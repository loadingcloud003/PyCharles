# -*- coding: utf-8 -*-
from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, ElementTransformUtils, CopyPasteOptions, Transaction, RevitLinkInstance, ElementId
from Autodesk.Revit.UI import TaskDialog, Selection
from System.Collections.Generic import List
import csv
import os
from Autodesk.Revit.DB import FamilyInstance

__doc__ = "Copy a selected element from a linked model and paste it into the current model using shared coordinates."
__title__ = "Copy Link Elements"
__author__ = "Your Name"

uidoc = revit.uidoc
output = script.get_output()

# Custom ISelectionFilter to allow only RevitLinkInstance selection
class LinkInstanceSelectionFilter(Selection.ISelectionFilter):
    def AllowElement(self, element):
        # Use built-in category check and class check for RevitLinkInstance
        return element.Category and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_RvtLinks)
    def AllowReference(self, ref, point):
        return True

# Custom ISelectionFilter for linked elements (accept all)
class LinkedElementSelectionFilter(Selection.ISelectionFilter):
    def AllowElement(self, element):
        return True  # Accept all elements in the linked model
    def AllowReference(self, ref, point):
        return True

try:
    # 1. Pick a link instance in the current view
    TaskDialog.Show("Step 1", "Please select a Revit Link instance in the current view.")
    link_ref = uidoc.Selection.PickObject(Selection.ObjectType.Element, LinkInstanceSelectionFilter(), "Select a Revit Link instance.")
    link_instance = revit.doc.GetElement(link_ref.ElementId)
    link_doc = link_instance.GetLinkDocument()
    if not link_doc:
        TaskDialog.Show("Error", "Failed to get linked document.")
        script.exit()

    # 2. Pick multiple elements in the linked model
    TaskDialog.Show("Step 2", "Now select one or more elements in the linked model.")
    linked_elem_refs = uidoc.Selection.PickObjects(Selection.ObjectType.LinkedElement, LinkedElementSelectionFilter(), "Select elements in the linked model.")
    linked_elem_ids = [ref.LinkedElementId for ref in linked_elem_refs]
    linked_elems = [link_doc.GetElement(eid) for eid in linked_elem_ids]

    # --- New: Ask user to select categories to copy ---
    # Gather all categories from selected elements
    selected_categories = set()
    for elem in linked_elems:
        try:
            if elem.Category:
                selected_categories.add(elem.Category.Name)
        except Exception:
            pass
    if not selected_categories:
        TaskDialog.Show("Error", "No categories found in selected elements.")
        script.exit()
    # Sort categories for display
    sorted_categories = sorted(selected_categories)
    # Build a string for user selection (comma separated)
    category_options = "\n".join(["[{}] {}".format(i+1, cat) for i, cat in enumerate(sorted_categories)])
    prompt = "Select categories to copy by entering their numbers separated by comma (e.g. 1,3,5):\n" + category_options
    # Use Windows Forms for input box
    import clr
    clr.AddReference('System.Windows.Forms')
    from System.Windows.Forms import Form, Label, Button, DialogResult, CheckedListBox
    class InputForm(Form):
        def __init__(self, prompt, categories):
            self.Text = "Choose Categories"
            self.Width = 700
            base_height = 200
            extra_height = min(max(len(categories), 1) * 22, 700)
            self.Height = base_height + extra_height
            self.label = Label()
            self.label.Text = prompt
            self.label.Width = 670
            self.label.Height = 60
            self.label.Top = 10
            self.label.Left = 10
            self.label.AutoSize = False
            self.Controls.Add(self.label)
            self.clb = CheckedListBox()
            self.clb.Width = 650
            self.clb.Height = min(22 * max(len(categories), 1), 700)
            self.clb.Top = self.label.Top + self.label.Height + 10
            self.clb.Left = 10
            self.clb.CheckOnClick = True  # Enable ticking on single click
            for cat in categories:
                self.clb.Items.Add(cat)
            self.Controls.Add(self.clb)
            self.ok_button = Button()
            self.ok_button.Text = "OK"
            self.ok_button.Top = self.clb.Top + self.clb.Height + 10
            self.ok_button.Left = 410
            self.ok_button.Width = 100
            self.ok_button.DialogResult = DialogResult.OK
            self.Controls.Add(self.ok_button)
            self.cancel_button = Button()
            self.cancel_button.Text = "Cancel"
            self.cancel_button.Top = self.clb.Top + self.clb.Height + 10
            self.cancel_button.Left = 530
            self.cancel_button.Width = 100
            self.cancel_button.DialogResult = DialogResult.Cancel
            self.Controls.Add(self.cancel_button)
            self.AcceptButton = self.ok_button
            self.CancelButton = self.cancel_button
    form = InputForm("Select categories to copy (check all that apply):", sorted_categories)
    result = form.ShowDialog()
    if result == DialogResult.OK:
        chosen_categories = set([str(form.clb.Items[i]) for i in range(form.clb.Items.Count) if form.clb.GetItemChecked(i)])
    else:
        chosen_categories = set()
    if not chosen_categories:
        TaskDialog.Show("Cancelled", "No categories selected.")
        script.exit()
    # Filter linked_elems by chosen categories
    linked_elems = [elem for elem in linked_elems if elem.Category and elem.Category.Name in chosen_categories]
    if not linked_elems:
        TaskDialog.Show("Error", "No elements match the selected categories.")
        script.exit()

    # 3. Copy the link elements and paste by shared coordinate
    t = Transaction(revit.doc, "Copy Link Elements by Shared Coordinate")
    t.Start()
    ids = List[ElementId]([elem.Id for elem in linked_elems])
    from Autodesk.Revit.DB import Transform
    total_transform = link_instance.GetTotalTransform()
    mapping = ElementTransformUtils.CopyElements(link_doc, ids, revit.doc, total_transform, CopyPasteOptions())
    for new_id in mapping:
        new_elem = revit.doc.GetElement(new_id)
        # No operation needed here, just ensure elements are copied and placed by transform
    t.Commit()

    # 4. Export to CSV: original info, copied info, match status (no mapping, just sort and align)
    export_rows = []
    link_name = link_instance.Name if hasattr(link_instance, 'Name') else ''
    export_rows.append([
        'Link Name', 'Original ElementId', 'Original Category', 'Original Family and Type', 'Original X', 'Original Y', 'Original Z',
        'Copied ElementId', 'Copied Category', 'Copied Family and Type', 'Copied X', 'Copied Y', 'Copied Z', 'XYZ Match', 'Family Type Match', 'Remark'
    ])
    # Build and sort original info (unique only)
    unique_elem_ids = set(elem.Id for elem in linked_elems)
    original_info = []
    # Get the inverse of the link's total transform for world coordinate conversion
    world_transform = link_instance.GetTotalTransform()
    inverse_transform = world_transform.Inverse
    for elem_id in unique_elem_ids:
        elem = link_doc.GetElement(elem_id)
        orig_loc = elem.Location
        if orig_loc is None:
            orig_xyz = ("no location api", "no location api", "no location api")
        elif hasattr(orig_loc, 'Point') and orig_loc.Point:
            orig_pt = orig_loc.Point
            # Convert to world coordinate
            world_pt = world_transform.OfPoint(orig_pt)
            orig_xyz = (round(world_pt.X, 6), round(world_pt.Y, 6), round(world_pt.Z, 6))
        elif hasattr(orig_loc, 'Curve') and orig_loc.Curve:
            curve = orig_loc.Curve
            mid_pt = curve.Evaluate(0.5, True)
            world_mid_pt = world_transform.OfPoint(mid_pt)
            orig_xyz = (round(world_mid_pt.X, 6), round(world_mid_pt.Y, 6), round(world_mid_pt.Z, 6))
        else:
            orig_xyz = ("no location api", "no location api", "no location api")
        orig_family_type = ''
        orig_category = ''
        try:
            param = elem.LookupParameter('Family and Type')
            if param:   
                orig_family_type = param.AsValueString()
        except Exception:
            pass
        try:
            if elem.Category:
                orig_category = elem.Category.Name
        except Exception:
            pass
        original_info.append((elem_id, orig_category, orig_family_type, orig_xyz))
    original_info.sort(key=lambda x: x[0].IntegerValue)
    # Build and sort copied info
    copied_info = []
    for new_elem in [revit.doc.GetElement(eid) for eid in mapping]:
        if new_elem is None:
            continue
        copied_id = new_elem.Id
        copied_loc = new_elem.Location
        if copied_loc is None:
            copied_xyz = ("no location api", "no location api", "no location api")
        elif hasattr(copied_loc, 'Point') and copied_loc.Point:
            pt = copied_loc.Point
            # Already in world coordinate
            copied_xyz = (round(pt.X, 6), round(pt.Y, 6), round(pt.Z, 6))
        elif hasattr(copied_loc, 'Curve') and copied_loc.Curve:
            curve = copied_loc.Curve
            mid_pt = curve.Evaluate(0.5, True)
            copied_xyz = (round(mid_pt.X, 6), round(mid_pt.Y, 6), round(mid_pt.Z, 6))
        else:
            copied_xyz = ("no location api", "no location api", "no location api")
        copied_family_type = ''
        copied_category = ''
        try:
            param = new_elem.LookupParameter('Family and Type')
            if param:
                copied_family_type = param.AsValueString()
        except Exception:
            pass
        try:
            if new_elem.Category:
                copied_category = new_elem.Category.Name
        except Exception:
            pass
        copied_info.append((copied_id, copied_category, copied_family_type, copied_xyz))
    copied_info.sort(key=lambda x: x[0].IntegerValue)
    # Combine by row order
    max_len = max(len(original_info), len(copied_info))
    xyz_error_count = 0
    family_type_error_count = 0
    for i in range(max_len):
        # Original
        if i < len(original_info):
            orig_id, orig_category, orig_family_type, orig_xyz = original_info[i]
        else:
            orig_id, orig_category, orig_family_type, orig_xyz = ('', '', '', ("", "", ""))
        # Copied
        if i < len(copied_info):
            copied_id, copied_category, copied_family_type, copied_xyz = copied_info[i]
        else:
            copied_id, copied_category, copied_family_type, copied_xyz = ('', '', '', ("", "", ""))
        # Match checks
        xyz_match = False
        try:
            xyz_match = all(abs(float(a) - float(b)) < 0.001 for a, b in zip(orig_xyz, copied_xyz))
        except Exception:
            pass
        match_status = 'Match' if xyz_match else 'No Match'
        if match_status == 'No Match':
            xyz_error_count += 1
        family_type_match = 'Match' if orig_family_type == copied_family_type and orig_family_type != '' else 'No Match'
        if family_type_match == 'No Match':
            family_type_error_count += 1
        # Remark logic
        remark = ''
        if match_status == 'No Match':
            if orig_xyz[2] != "no location api" and copied_xyz[2] != "no location api":
                try:
                    if abs(float(orig_xyz[2]) - float(copied_xyz[2])) > 0.001:
                        remark = "please check the current model level equal to link model or not"
                except Exception:
                    pass
            elif (orig_xyz[0] == "no location api" or copied_xyz[0] == "no location api"):
                remark = "Please check manually"
        export_rows.append([
            link_name,
            str(orig_id.IntegerValue) if orig_id else '', orig_category, orig_family_type, orig_xyz[0], orig_xyz[1], orig_xyz[2],
            str(copied_id.IntegerValue) if copied_id else '', copied_category, copied_family_type, copied_xyz[0], copied_xyz[1], copied_xyz[2],
            match_status, family_type_match, remark
        ])
    # Export to CSV in user's Documents folder
    docs = os.path.expanduser('~\\Documents')
    csv_path = os.path.join(docs, 'pyrevit_copy_link_elements_report.csv')
    with open(csv_path, 'w') as f:
        writer = csv.writer(f, lineterminator='\n')
        writer.writerows(export_rows)
    # Count copied element categories
    from collections import Counter
    copied_category_list = [cat for (_, cat, _, _) in copied_info if cat]
    copied_category_counter = Counter(copied_category_list)
    # Count original element categories
    original_category_list = [cat for (_, cat, _, _) in original_info if cat]
    original_category_counter = Counter(original_category_list)
    output.print_md("**Results exported to:** {}".format(csv_path))
    output.print_md("**Total elements copied:** {}".format(len(copied_info)))
    output.print_md("**Total XYZ errors:** {}".format(xyz_error_count))
    output.print_md("**Total Family and Type Name errors:** {}".format(family_type_error_count))
    output.print_md("**Original element category counts:**")
    for cat, count in original_category_counter.items():
        output.print_md("- {}: {}".format(cat, count))
    output.print_md("**Copied element category counts:**")
    for cat, count in copied_category_counter.items():
        output.print_md("- {}: {}".format(cat, count))

except Exception as e:
    TaskDialog.Show("Error", str(e))
    script.exit()
