# -*- coding: utf-8 -*-
# pyRevit addin script for reading model_comparison_summary_by_category.csv, creating filters, and printing PDF
import os
import csv
import json
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import FolderBrowserDialog, OpenFileDialog, Form, Label, ListBox, Button, DialogResult, CheckedListBox, ColorDialog, GroupBox, MessageBox, MessageBoxButtons, MessageBoxIcon
from System.Windows.Forms import SelectionMode
import System.Drawing
from pyrevit import revit, script
from datetime import datetime
from Autodesk.Revit.DB import ParameterFilterElement, ElementParameterFilter, FilteredElementCollector, BuiltInCategory, ElementId

CSV_FILENAME = "model_comparison_summary_by_category.csv"
SELECTION_RECORD = "last_selection.json"

# 1. Select folder and model
def select_folder():
    """Show a dialog to select a folder containing Revit models."""
    dialog = FolderBrowserDialog()
    dialog.Description = "Select the folder containing Revit models."
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.SelectedPath
    return None

def select_model(title, initial_dir=None):
    """Show a dialog to select a Revit model file."""
    dialog = OpenFileDialog()
    if initial_dir:
        dialog.InitialDirectory = initial_dir
    dialog.Title = title
    dialog.Filter = "Revit Files (*.rvt)|*.rvt"
    dialog.Multiselect = False
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    return None

def select_csv_file(initial_dir=None):
    """Show a dialog to select the CSV file for model comparison summary."""
    dialog = OpenFileDialog()
    if initial_dir:
        dialog.InitialDirectory = initial_dir
    dialog.Title = "Select the CSV file for model comparison summary"
    dialog.Filter = "CSV Files (*.csv)|*.csv"
    dialog.Multiselect = False
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    return None

# 2. Read and process CSV

def read_summary_csv(csv_path):
    """Read the summary CSV and return a list of rows as dictionaries."""
    rows = []
    with open(csv_path, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            rows.append(row)
    return rows

def group_results_by_category_and_type(summary_rows):
    """Group CSV results by category and result type."""
    grouped = {}
    for row in summary_rows:
        cat = row.get('category', 'Unknown')
        if cat not in grouped:
            grouped[cat] = {}
        # For each result type, collect parameter lists (including type parameter variants)
        for rtype in [
            'XY coordination move', 'Z coordination move', 'new parameter add',
            'parameter delete', 'parameter value change', 'new element added', 'element deleted',
            'type parameter add', 'type parameter delete', 'type parameter value change'
        ]:
            if rtype not in grouped[cat]:
                grouped[cat][rtype] = set()  # Use set for deduplication
        # Add counts and lists
        for rtype, key, listkey in [
            ('XY coordination move', 'xy_move_count', None),
            ('Z coordination move', 'z_move_count', None),
            ('new parameter add', 'new_param_count', 'new_param_list'),
            ('parameter delete', 'del_param_count', 'del_param_list'),
            ('parameter value change', 'param_value_change_count', 'param_value_change_list'),
            ('new element added', 'new_elem_count', None),
            ('element deleted', 'del_elem_count', None),
            ('type parameter add', 'type_param_add_count', 'type_param_add_list'),
            ('type parameter delete', 'type_param_del_count', 'type_param_del_list'),
            ('type parameter value change', 'type_param_value_change_count', 'type_param_value_change_list')
        ]:
            count = int(row.get(key, 0))
            if listkey:
                items = [i.strip() for i in row.get(listkey, '').split(',') if i.strip()]
                grouped[cat][rtype].update(items)
            elif count > 0:
                grouped[cat][rtype].add('')  # Use empty string as placeholder for count-based items
    # Convert sets back to lists for UI compatibility
    for cat in grouped:
        for rtype in grouped[cat]:
            grouped[cat][rtype] = list(grouped[cat][rtype])
    return grouped

# 3. Dialog for filter creation
class FilterDialog(Form):
    def __init__(self, grouped):
        self.Text = "Select Result Types to Filter"
        self.Width = 1200
        self.Height = 900
        self.selected_items = []  # (cat, result_type, pname)
        self.grouped = grouped
        self.listboxes = []  # Store listboxes for each category
        self.selection_box = ListBox()
        self.selection_box.Top = 10
        self.selection_box.Left = 950
        self.selection_box.Width = 220
        self.selection_box.Height = 760
        self.selection_box.Font = System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Regular)
        self.selection_box.SelectionMode = SelectionMode.MultiExtended
        self.Controls.Add(self.selection_box)
        self.add_btns = []  # Store add buttons for each category
        y = 10
        for cat, rtypes in grouped.items():
            gb = GroupBox()
            gb.Text = cat
            gb.Top = y
            gb.Left = 10
            gb.Width = 920
            gb.Height = 320
            gb.Font = System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
            gy = 20
            lb = ListBox()
            lb.Top = gy
            lb.Left = 10
            lb.Width = 800
            lb.Height = 250
            lb.SelectionMode = SelectionMode.MultiExtended
            lb.ScrollAlwaysVisible = True
            result_items = []
            for rtype, items in rtypes.items():
                if rtype in ["XY coordination move", "Z coordination move", "new element added", "element deleted"]:
                    if items:
                        for _ in items:
                            result_items.append(rtype)
                    else:
                        result_items.append(rtype)
                elif rtype in ["new parameter add", "type parameter add"]:
                    for pname in items:
                        result_items.append("{}: {}".format(rtype, pname))
                elif rtype in ["parameter delete", "type parameter delete"]:
                    for pname in items:
                        result_items.append("{}: {}".format(rtype, pname))
                elif rtype in ["parameter value change", "type parameter value_change"]:
                    for pname in items:
                        result_items.append("{}: {}".format(rtype, pname))
            for item in result_items:
                lb.Items.Add(item)
            lb.DoubleClick += self.add_selected_item_doubleclick
            gb.Controls.Add(lb)
            self.listboxes.append((cat, lb))
            # Add button for this category
            add_btn = Button()
            add_btn.Text = "Add >>"
            add_btn.Top = lb.Top + lb.Height // 2
            add_btn.Left = lb.Left + lb.Width + 10
            add_btn.Width = 80
            add_btn.Tag = (cat, lb)
            add_btn.Click += self.add_selected_item_button
            gb.Controls.Add(add_btn)
            self.add_btns.append(add_btn)
            self.Controls.Add(gb)
            y += gb.Height + 10
        # Remove button for selection box
        remove_btn = Button()
        remove_btn.Text = "<< Remove"
        remove_btn.Top = self.selection_box.Top + self.selection_box.Height + 10
        remove_btn.Left = self.selection_box.Left
        remove_btn.Width = 120
        remove_btn.Click += self.remove_selected_item
        self.Controls.Add(remove_btn)
        ok_btn = Button()
        ok_btn.Text = "Next"
        ok_btn.Top = y + 20
        ok_btn.Left = 220
        ok_btn.Width = 120
        ok_btn.DialogResult = DialogResult.OK
        self.Controls.Add(ok_btn)
        self.AcceptButton = ok_btn
    def add_selected_item_doubleclick(self, sender, args):
        for i in range(sender.Items.Count):
            if sender.GetSelected(i):
                item_text = sender.Items[i]
                # Find category for this listbox
                for cat, lb in self.listboxes:
                    if lb == sender:
                        entry = (cat, item_text)
                        if entry not in self.selected_items:
                            self.selected_items.append(entry)
                            self.selection_box.Items.Add("{}: {}".format(cat, item_text))
    def add_selected_item_button(self, sender, args):
        cat, lb = sender.Tag
        for i in range(lb.Items.Count):
            if lb.GetSelected(i):
                item_text = lb.Items[i]
                entry = (cat, item_text)
                if entry not in self.selected_items:
                    self.selected_items.append(entry)
                    self.selection_box.Items.Add("{}: {}".format(cat, item_text))
    def remove_selected_item(self, sender, args):
        # Remove selected items from selection_box and selected_items
        to_remove = []
        for i in range(self.selection_box.Items.Count):
            if self.selection_box.GetSelected(i):
                to_remove.append(i)
        # Remove from end to avoid index shift
        for i in reversed(to_remove):
            item_text = self.selection_box.Items[i]
            self.selection_box.Items.RemoveAt(i)
            # Parse category and item
            if ': ' in item_text:
                cat, item = item_text.split(': ', 1)
                entry = (cat, item)
                if entry in self.selected_items:
                    self.selected_items.remove(entry)
    def get_selected_items(self):
        return self.selected_items

class ColorAssignDialog(Form):
    def __init__(self, selected_items):
        self.Text = "Assign Colors to Selected Filters"
        self.Width = 800
        self.Height = 600
        self.color_map = {}  # (cat, item) -> color
        y = 10
        self.buttons = []
        for idx, (cat, item) in enumerate(selected_items):
            lbl = Label()
            lbl.Text = "{}: {}".format(cat, item)
            lbl.Top = y + idx * 40
            lbl.Left = 10
            lbl.Width = 500
            self.Controls.Add(lbl)
            btn = Button()
            btn.Text = "Select Color"
            btn.Top = y + idx * 40
            btn.Left = 520
            btn.Width = 120
            btn.Tag = (cat, item)
            btn.BackColor = System.Drawing.Color.White
            btn.Click += self.select_color
            self.Controls.Add(btn)
            self.color_map[(cat, item)] = System.Drawing.Color.White
            self.buttons.append(btn)
        ok_btn = Button()
        ok_btn.Text = "OK"
        ok_btn.Top = y + len(selected_items) * 40 + 20
        ok_btn.Left = 320
        ok_btn.Width = 120
        ok_btn.DialogResult = DialogResult.OK
        self.Controls.Add(ok_btn)
        self.AcceptButton = ok_btn
    def select_color(self, sender, args):
        cat, item = sender.Tag
        cd = ColorDialog()
        if cd.ShowDialog() == DialogResult.OK:
            c = cd.Color
            self.color_map[(cat, item)] = c
            sender.BackColor = c

# 4. Select views to apply filters
class ViewSelectForm(Form):
    def __init__(self, views):
        self.Text = "Select Views to Apply Filters"
        self.Width = 600
        self.Height = 600
        self.views = views
        self.clb = CheckedListBox()
        self.clb.Top = 10
        self.clb.Left = 10
        self.clb.Width = 560
        self.clb.Height = 500
        for v in views:
            self.clb.Items.Add(v)
        self.Controls.Add(self.clb)
        ok_btn = Button()
        ok_btn.Text = "OK"
        ok_btn.Top = 520
        ok_btn.Left = 250
        ok_btn.Width = 80
        ok_btn.DialogResult = DialogResult.OK
        self.Controls.Add(ok_btn)
        self.AcceptButton = ok_btn

# 5. Select views/sheets to print PDF
class PrintSelectForm(Form):
    def __init__(self, views, sheets):
        self.Text = "Select Views/Sheets to Print to PDF"
        self.Width = 800
        self.Height = 700
        self.views = views
        self.sheets = sheets
        self.clb_views = CheckedListBox()
        self.clb_views.Top = 10
        self.clb_views.Left = 10
        self.clb_views.Width = 360
        self.clb_views.Height = 600
        for v in views:
            self.clb_views.Items.Add(v)
        self.Controls.Add(self.clb_views)
        self.clb_sheets = CheckedListBox()
        self.clb_sheets.Top = 10
        self.clb_sheets.Left = 400
        self.clb_sheets.Width = 360
        self.clb_sheets.Height = 600
        for s in sheets:
            self.clb_sheets.Items.Add(s)
        self.Controls.Add(self.clb_sheets)
        ok_btn = Button()
        ok_btn.Text = "OK"
        ok_btn.Top = 620
        ok_btn.Left = 350
        ok_btn.Width = 80
        ok_btn.DialogResult = DialogResult.OK
        self.Controls.Add(ok_btn)
        self.AcceptButton = ok_btn

def get_views_from_model(doc):
    from Autodesk.Revit.DB import FilteredElementCollector, View, ViewType
    if doc is None:
        print("Error: No active Revit document. Please ensure the model is opened and set as current.")
        return []
    views = [v for v in FilteredElementCollector(doc).OfClass(View)
             if not v.IsTemplate and v.ViewType != ViewType.Internal]
    return [v.Name for v in views]


def get_sheets_from_model(doc):
    from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet

    if doc is None:
        print("Error: No active Revit document. Please ensure the model is opened and set as current.")
        return []

    sheets = [s for s in FilteredElementCollector(doc).OfClass(ViewSheet)]
    return [s.Name for s in sheets]

def get_selection_record_path(model_path):
    import os
    model_dir = os.path.dirname(model_path)
    return os.path.join(model_dir, "last_selection.json")

def save_selection_record(record, model_path):
    path = get_selection_record_path(model_path)
    try:
        with open(path, 'w') as f:
            import json
            json.dump(record, f)
    except Exception as e:
        print("Error saving selection record:", e)

def load_selection_record(model_path):
    path = get_selection_record_path(model_path)
    try:
        import os
        if os.path.exists(path):
            with open(path, 'r') as f:
                import json
                return json.load(f)
    except Exception as e:
        print("Error loading selection record:", e)
    return {}

def add_compare_result_filter(doc, entry):
    """
    For each item in entry (comma-separated string), split by ',' to get category and result type,
    then create a ParameterFilterElement for each pair.
    entry: string like "Category1,ResultType1;Category2,ResultType2;..."
    Returns a list of created filter elements.
    """
    from Autodesk.Revit.DB import ParameterFilterRuleFactory, ElementId, Transaction, FilteredElementCollector, BuiltInCategory, ParameterFilterElement, ElementParameterFilter
    from pyrevit import script
    from datetime import datetime
    today_str = datetime.now().strftime('%Y%m%d')
    print(entry)
    # entry is a tuple: (category, result_type)
    if isinstance(entry, tuple) and len(entry) == 2:
        category, result_type = entry
    else:
        script.get_output().print_md("**Input Error:** Entry '{}' is not a tuple of (category, result_type).".format(entry))
        return None
    # Get BuiltInCategory from doc.Settings.Categories
    def get_built_in_category_by_name(doc, category_name):
        categories = doc.Settings.Categories
        for cat in categories:
            if cat.Name == category_name:
                bic = cat.BuiltInCategory
                if bic != BuiltInCategory.INVALID:
                    return bic
        return None
    cat_enum = get_built_in_category_by_name(doc, category)
    if cat_enum is None:
        script.dialogs.alert("Category '{}' not found or not a built-in category.".format(category), title="Category Error", warn_icon=True)
    param_id = None
    collector = FilteredElementCollector(doc).OfCategory(cat_enum).WhereElementIsNotElementType()
    for elem in collector:
        p = elem.LookupParameter('compare_results')
        if p:
            param_id = p.Id
            break
        if param_id is None:
            script.dialogs.alert("Could not find 'compare_results' parameter in category '{}'.".format(category), title="Parameter Error", warn_icon=True)
            continue
    # Create rule
    rule = ParameterFilterRuleFactory.CreateContainsRule(param_id, result_type, False)
    param_rules = [rule]
    element_filter = ElementParameterFilter(param_rules)
    filter_name = today_str + '_' + category.replace(' ', '_') + '_' + result_type.replace(':', '_')
    from System.Collections.Generic import List
    cats = List[ElementId]()
    cats.Add(ElementId(int(cat_enum)))
    t = Transaction(doc, "Create Filter Element")
    t.Start()
    filter_elem = ParameterFilterElement.Create(doc, filter_name, cats, element_filter)
    t.Commit()
    return filter_elem

def apply_filter_to_views(doc, filter_elem, color, view_names):
    """
    Apply the filter to the given views and set the view's override color.
    """
    from Autodesk.Revit.DB import OverrideGraphicSettings, Transaction, FilteredElementCollector, View
    # Find views by name
    views = [v for v in FilteredElementCollector(doc).OfClass(View) if v.Name in view_names]
    # Set up color override for both projection and cut, lines and patterns
    ogs = OverrideGraphicSettings()
    # Projection
    ogs.SetProjectionLineColor(color)
    ogs.SetProjectionLinePatternId(ElementId.InvalidElementId)
    try:
        ogs.SetProjectionPatternColor(color)
    except Exception:
        pass
    # Cut
    ogs.SetCutLineColor(color)
    ogs.SetCutLinePatternId(ElementId.InvalidElementId)
    try:
        ogs.SetCutPatternColor(color)
    except Exception:
        pass
    # Apply filter and color in a transaction
    t = Transaction(doc, "Apply Filter and Color")
    t.Start()
    for v in views:
        v.AddFilter(filter_elem.Id)
        v.SetFilterOverrides(filter_elem.Id, ogs)
    t.Commit()
    print("Applied filter {} to views {} with color {}".format(filter_elem.Name, view_names, color))

# --- Main Workflow ---
def main():
    folder = select_folder()
    if not folder:
        print("No folder selected.")
        return
    model_path = select_model("Select the Revit model for reference", folder)
    if not model_path:
        print("No model selected.")
        return
    from Autodesk.Revit.DB import ModelPathUtils, OpenOptions, Color, BuiltInCategory
    app = revit.doc.Application
    model_path_obj = ModelPathUtils.ConvertUserVisiblePathToModelPath(model_path)
    opts = OpenOptions()
    opts.DetachFromCentralOption = 0
    doc = app.OpenDocumentFile(model_path_obj, opts)
    if doc is None:
        print("Failed to open or set the Revit model. Please ensure you are running inside Revit and the model path is valid.")
        return
    last = load_selection_record(model_path)
    csv_path = select_csv_file(folder) if not last.get('csv_path') else last['csv_path']
    if not csv_path or not os.path.exists(csv_path):
        print("CSV file not found: {}".format(csv_path))
        return
    summary_rows = read_summary_csv(csv_path)
    grouped = group_results_by_category_and_type(summary_rows)
    filter_form = FilterDialog(grouped)
    if filter_form.ShowDialog() == DialogResult.OK:
        selected_items = filter_form.get_selected_items() if not last.get('selected_items') else last['selected_items']
        if not selected_items:
            from pyrevit import script
            script.dialogs.alert("No result types selected.", title="Info", warn_icon=True)
            return
        # selected_items: list of (cat, result_type) or (cat, result_type, category)
        color_form = ColorAssignDialog([(cat, item) for (cat, item) in selected_items])
        if color_form.ShowDialog() == DialogResult.OK:
            color_map = color_form.color_map if not last.get('color_map') else last['color_map']
            filter_elems = []
            for entry in selected_items:
                if len(entry) == 3:
                    cat, item, category = entry
                else:
                    cat, item = entry
                    category = cat
                # Always ensure cat = category for downstream usage
                cat = category
                result_type = item.split(': ', 1)[-1] if ': ' in item else item
                print(entry)
                print(category)
                print(result_type)
                try:
                    # Robust BuiltInCategory conversion using doc.Settings.Categories
                    def get_built_in_category_by_name(doc, category_name):
                        categories = doc.Settings.Categories
                        for cat in categories:
                            if cat.Name == category_name:
                                bic = cat.BuiltInCategory
                                if bic != BuiltInCategory.INVALID:
                                    print(bic)
                                    return bic
                                else:
                                    print("Category found but is not a built-in category.")
                                    return None
                        print("Category named '{}' not found.".format(category_name))
                        return None

                    cat_enum = None
                    cat_enum = get_built_in_category_by_name(doc, category)
                    if cat_enum is None:
                        raise ValueError("Invalid BuiltInCategory: {}".format(category))
                    filter_elem = add_compare_result_filter(doc, entry)
                except Exception as e:
                    filter_elem = add_compare_result_filter(doc, entry)
                if filter_elem:
                    filter_elems.append((filter_elem, color_map[(cat, item)]))
            views = get_views_from_model(doc)
            view_form = ViewSelectForm(views)
            if view_form.ShowDialog() == DialogResult.OK:
                selected_views = [views[i] for i in range(view_form.clb.Items.Count) if view_form.clb.GetItemChecked(i)] if not last.get('selected_views') else last['selected_views']
                for filter_elem, color in filter_elems:
                    revit_color = Color(color.R, color.G, color.B)
                    apply_filter_to_views(doc, filter_elem, revit_color, selected_views)
                sheets = get_sheets_from_model(doc)
                print_form = PrintSelectForm(views, sheets)
                if print_form.ShowDialog() == DialogResult.OK:
                    selected_views_to_print = [views[i] for i in range(print_form.clb_views.Items.Count) if print_form.clb_views.GetItemChecked(i)] if not last.get('selected_views_to_print') else last['selected_views_to_print']
                    selected_sheets_to_print = [sheets[i] for i in range(print_form.clb_sheets.Items.Count) if print_form.clb_sheets.GetItemChecked(i)] if not last.get('selected_sheets_to_print') else last['selected_sheets_to_print']
                    save_selection_record({
                        'folder': folder,
                        'model_path': model_path,
                        'csv_path': csv_path,
                        'selected_items': selected_items,
                        'color_map': color_map,
                        'selected_views': selected_views,
                        'selected_views_to_print': selected_views_to_print,
                        'selected_sheets_to_print': selected_sheets_to_print
                    }, model_path)

    # Print selected views/sheets to PDF using selected print setting
    try:
        from Autodesk.Revit.DB import PrintManager, Transaction, ViewSheetSetting, ViewSheet, View, PrintRange, PaperPlacement, PrintSetup, PrintParameters
        # Load last selection record
        last = load_selection_record(model_path)
        selected_views_to_print = last.get('selected_views_to_print', [])
        selected_sheets_to_print = last.get('selected_sheets_to_print', [])
        # Get view and sheet objects
        views_to_print = [v for v in FilteredElementCollector(doc).OfClass(View) if v.Name in selected_views_to_print]
        sheets_to_print = [s for s in FilteredElementCollector(doc).OfClass(ViewSheet) if s.Name in selected_sheets_to_print]
        # Combine all for printing
        to_print = views_to_print + sheets_to_print
        if not to_print:
            print("No views or sheets selected for PDF printing.")
            return
        # Set up print manager
        pm = doc.PrintManager
        pm.PrintRange = PrintRange.Select
        pm.ViewSheetSetting.CurrentViewSheetSet.Views.Clear()
        from System.Collections.Generic import List
        view_ids = List[ElementId]([v.Id for v in to_print])
        pm.ViewSheetSetting.CurrentViewSheetSet.Views = view_ids
        # Use the currently selected print setting
        ps = pm.PrintSetup
        # Fit to page
        pparams = ps.CurrentPrintSetting.PrintParameters
        pparams.ZoomType = 1  # 1 = Fit to page
        pparams.PaperPlacement = PaperPlacement.Center
        ps.CurrentPrintSetting.PrintParameters = pparams
        # Print to PDF
        pm.PrintToFile = True
        pm.PrintToFileName = os.path.join(os.path.dirname(model_path), "CompareResults.pdf")
        t = Transaction(doc, "Print to PDF")
        t.Start()
        pm.SubmitPrint()
        t.Commit()
        print("PDF print job submitted for selected views/sheets using current print setting.")
    except Exception as e:
        print("Error during PDF printing:", e)

    # After all modifications, save the model
    doc.Save()
    print("Model saved after modifications.")

if __name__ == "__main__":
    main()