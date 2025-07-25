# -*- coding: utf-8 -*-
from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, ElementTransformUtils, CopyPasteOptions, Transaction, RevitLinkInstance, ElementId
from Autodesk.Revit.UI import TaskDialog, Selection
from System.Collections.Generic import List
import csv
import os
from Autodesk.Revit.DB import FamilyInstance
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import FolderBrowserDialog, OpenFileDialog, Form, Label, Button, DialogResult, CheckedListBox, ColorDialog, ComboBox, Panel
from Autodesk.Revit.DB import FilteredElementCollector, ParameterFilterElement, FilterRule, FilterStringRule, FilterStringEquals, OverrideGraphicSettings, Color, View
import datetime
import time
from Autodesk.Revit.DB import BuiltInParameterGroup, ViewType
from System.Windows.Forms import SelectionMode

# --- Helper Functions ---
def select_folder():
    dialog = FolderBrowserDialog()
    dialog.Description = "Select the folder containing Revit models."
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.SelectedPath
    return None

def select_model(title, initial_dir=None):
    dialog = OpenFileDialog()
    if initial_dir:
        dialog.InitialDirectory = initial_dir
    dialog.Title = title
    dialog.Filter = "Revit Files (*.rvt)|*.rvt"
    dialog.Multiselect = False
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    return None

def show_analysis_item_selection():
    items = [
        "XYZ deviation",
        "Parameter value change",
        "Newly/deleted elements"
    ]
    form = Form()
    form.Text = "Select Analysis Items"
    form.Width = 400
    form.Height = 300
    label = Label()
    label.Text = "Check analysis items to include:"
    label.Top = 10
    label.Left = 10
    label.Width = 350
    form.Controls.Add(label)
    clb = CheckedListBox()
    clb.Width = 350
    clb.Height = 120
    clb.Top = 40
    clb.Left = 10
    clb.CheckOnClick = True
    for item in items:
        clb.Items.Add(item)
    form.Controls.Add(clb)
    ok_button = Button()
    ok_button.Text = "OK"
    ok_button.Top = 180
    ok_button.Left = 200
    ok_button.Width = 80
    ok_button.DialogResult = DialogResult.OK
    form.Controls.Add(ok_button)
    form.AcceptButton = ok_button
    if form.ShowDialog() == DialogResult.OK:
        return [str(clb.Items[i]) for i in range(clb.Items.Count) if clb.GetItemChecked(i)]
    return []

def show_category_selection(categories):
    form = Form()
    form.Text = "Select Categories"
    form.Width = 400
    form.Height = 400
    label = Label()
    label.Text = "Check categories to include:"
    label.Top = 10
    label.Left = 10
    label.Width = 350
    form.Controls.Add(label)
    clb = CheckedListBox()
    clb.Width = 350
    clb.Height = 220
    clb.Top = 40
    clb.Left = 10
    clb.CheckOnClick = True
    for cat in categories:
        clb.Items.Add(cat)
    form.Controls.Add(clb)
    ok_button = Button()
    ok_button.Text = "OK"
    ok_button.Top = 280
    ok_button.Left = 200
    ok_button.Width = 80
    ok_button.DialogResult = DialogResult.OK
    form.Controls.Add(ok_button)
    form.AcceptButton = ok_button
    if form.ShowDialog() == DialogResult.OK:
        return [str(clb.Items[i]) for i in range(clb.Items.Count) if clb.GetItemChecked(i)]
    return []

def get_all_model_categories():
    return [
        "Walls",
        "Floors",
        "Roofs",
        "Doors",
        "Windows",
        "Columns",  # Architectural Columns
        "Structural Framing",  # Includes beams
        "Curtain Walls",
        "Curtain Panels",
        "Curtain Wall Mullions",
        "Stairs",
        "Railings",
        "Ceilings",
        "Rooms",
        "Spaces",
        "Furniture",
        "Casework",
        "Specialty Equipment",
        "Mass",
        "Topography",
        "Site",
        "Structural Foundations",
        "Structural Beam Systems",
        "Structural Columns",
        "Structural Trusses",
        "Structural Stiffeners",
        "Generic Models",
        "Detail Items"
    ]

def extract_xyz_by_category(doc, categories):
    """
    Extracts the XYZ location (in world coordinates), family and type, and category of elements in the given categories from the current opened model.
    Returns a dict: {element_id: (family_and_type, category, (x, y, z))}
    """
    from Autodesk.Revit.DB import FilteredElementCollector, Transform
    xyz_data = {}
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    # Get the model transform (identity for main model, or use GetTotalTransform for links)
    transform = Transform.Identity
    if hasattr(doc, 'ActiveProjectLocation') and doc.ActiveProjectLocation:
        try:
            transform = doc.ActiveProjectLocation.GetTotalTransform()
        except Exception:
            pass
    for elem in collector:
        try:
            if elem.Category and elem.Category.Name in categories:
                loc = elem.Location
                if loc is None:
                    continue
                if hasattr(loc, 'Point') and loc.Point:
                    pt = loc.Point
                    world_pt = transform.OfPoint(pt)
                    xyz = (round(world_pt.X, 6), round(world_pt.Y, 6), round(world_pt.Z, 6))
                elif hasattr(loc, 'Curve') and loc.Curve:
                    curve = loc.Curve
                    mid_pt = curve.Evaluate(0.5, True)
                    world_pt = transform.OfPoint(mid_pt)
                    xyz = (round(world_pt.X, 6), round(world_pt.Y, 6), round(world_pt.Z, 6))
                else:
                    continue
                fam_type = ''
                try:
                    param = elem.LookupParameter('Family and Type')
                    if param:
                        fam_type = param.AsValueString()
                except Exception:
                    pass
                xyz_data[elem.Id.IntegerValue] = (fam_type, elem.Category.Name, xyz)
        except Exception:
            pass
    return xyz_data

def extract_parameters_by_category(doc, categories):
    """
    Extracts all instance and type parameters and their values for elements in the given categories from the current opened model.
    Returns a dict: {element_id: {family_and_type: str, category: str, parameters: {param_name: param_value, ...}, type_parameters: {param_name: param_value, ...}}}
    """
    from Autodesk.Revit.DB import FilteredElementCollector
    param_data = {}
    type_param_cache = {}  # Cache type parameters by type id
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for elem in collector:
        try:
            if elem.Category and elem.Category.Name in categories:
                param_dict = {}
                for param in elem.Parameters:
                    try:
                        param_name = param.Definition.Name
                        if param.StorageType == 0:  # None
                            param_value = None
                        elif param.StorageType == 1:  # Integer
                            param_value = param.AsInteger()
                        elif param.StorageType == 2:  # Double
                            param_value = param.AsDouble()
                        elif param.StorageType == 3:  # String
                            param_value = param.AsString()
                        elif param.StorageType == 4:  # ElementId
                            param_value = param.AsElementId().IntegerValue
                        else:
                            param_value = param.AsValueString()
                        param_dict[param_name] = param_value
                    except Exception:
                        pass
                # Extract type parameters with caching
                type_param_dict = {}
                try:
                    type_id = elem.GetTypeId()
                    if type_id and type_id.IntegerValue != -1:
                        if type_id not in type_param_cache:
                            type_elem = doc.GetElement(type_id)
                            tdict = {}
                            if type_elem:
                                for tparam in type_elem.Parameters:
                                    try:
                                        tparam_name = tparam.Definition.Name
                                        if tparam.StorageType == 0:
                                            tparam_value = None
                                        elif tparam.StorageType == 1:
                                            tparam_value = tparam.AsInteger()
                                        elif tparam.StorageType == 2:
                                            tparam_value = tparam.AsDouble()
                                        elif tparam.StorageType == 3:
                                            tparam_value = tparam.AsString()
                                        elif tparam.StorageType == 4:
                                            tparam_value = tparam.AsElementId().IntegerValue
                                        else:
                                            tparam_value = tparam.AsValueString()
                                        tdict[tparam_name] = tparam_value
                                    except Exception:
                                        pass
                            type_param_cache[type_id] = tdict
                        type_param_dict = type_param_cache[type_id]
                except Exception:
                    pass
                fam_type = ''
                try:
                    param = elem.LookupParameter('Family and Type')
                    if param:
                        fam_type = param.AsValueString()
                except Exception:
                    pass
                param_data[elem.Id.IntegerValue] = {
                    'family_and_type': fam_type,
                    'category': elem.Category.Name,
                    'parameters': param_dict,
                    'type_parameters': type_param_dict
                }
        except Exception:
            pass
    return param_data

def get_elements_by_category(doc, categories):
    """
    Returns a list of tuples for all elements in the selected categories in the given model document.
    Each tuple: (element_id, family_and_type, category)
    """
    from Autodesk.Revit.DB import FilteredElementCollector
    result = []
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for elem in collector:
        try:
            if elem.Category and elem.Category.Name in categories:
                eid = elem.Id.IntegerValue
                fam_type = ''
                try:
                    param = elem.LookupParameter('Family and Type')
                    if param:
                        fam_type = param.AsValueString()
                except Exception:
                    pass
                category = elem.Category.Name if elem.Category else ''
                result.append((eid, fam_type, category))
        except Exception:
            pass
    return result

def compare_xyz_data(prev_xyz_data, latest_xyz_data):
    """
    Compares XYZ data between previous and latest models by element_id.
    Returns a list of dicts with keys:
    'previous_element_id', 'current_element_id', 'previous_family_and_type', 'current_family_and_type', 'previous_category', 'current_category', 'compare_result', 'compare_date'
    """
    import datetime
    results = []
    for prev_id, (fam_type, cat, prev_xyz) in prev_xyz_data.items():
        if prev_id in latest_xyz_data:
            lfam_type, lcat, latest_xyz = latest_xyz_data[prev_id]
            dx = latest_xyz[0] - prev_xyz[0]
            dy = latest_xyz[1] - prev_xyz[1]
            dz = latest_xyz[2] - prev_xyz[2]
            xy_moved = abs(dx) > 0.001 or abs(dy) > 0.001
            z_moved = abs(dz) > 0.001
            if not xy_moved and not z_moved:
                compare_result = "No movement"
            elif xy_moved and not z_moved:
                xy_dist = ((dx ** 2 + dy ** 2) ** 0.5) * 304.8  # Revit units to mm
                compare_result = "XY coordination move + '{0}mm'".format(int(round(xy_dist)))
            elif z_moved and not xy_moved:
                z_dist = abs(dz) * 304.8  # Revit units to mm
                if dz > 0:
                    compare_result = "Z coordination move upward + '{0}mm'".format(int(round(z_dist)))
                else:
                    compare_result = "Z coordination move downward + '{0}mm'".format(int(round(z_dist)))
            elif xy_moved and z_moved:
                xy_dist = ((dx ** 2 + dy ** 2) ** 0.5) * 304.8
                z_dist = abs(dz) * 304.8
                if dz > 0:
                    compare_result = "XY coordination move + '{0}mm', Z coordination move upward + '{1}mm'".format(int(round(xy_dist)), int(round(z_dist)))
                else:
                    compare_result = "XY coordination move + '{0}mm', Z coordination move downward + '{1}mm'".format(int(round(xy_dist)), int(round(z_dist)))
            best_match_id = prev_id
        else:
            compare_result = "No match found"
            best_match_id = ''
        if compare_result not in ["No movement", "No match found"]:
            compare_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            results.append({
                'previous_element_id': prev_id,
                'current_element_id': best_match_id,
                'previous_family_and_type': fam_type,
                'current_family_and_type': fam_type,  # XYZ only compares same element, so use fam_type for both
                'previous_category': cat,
                'current_category': cat,
                'compare_result': compare_result,
                'compare_date': compare_date
            })
    return results

def compare_param_data(prev_param_data, latest_param_data):
    """
    Compares parameter data (instance and type) between previous and latest models by element_id and parameter name.
    Groups results by element id, concatenating compare results with a comma.
    Returns a list of dicts with keys:
    'previous_element_id', 'current_element_id', 'previous_family_and_type', 'current_family_and_type', 'previous_category', 'current_category', 'compare_result', 'compare_date'
    """
    import datetime
    from collections import defaultdict
    grouped = defaultdict(lambda: {
        'previous_element_id': '',
        'current_element_id': '',
        'previous_family_and_type': '',
        'current_family_and_type': '',
        'previous_category': '',
        'current_category': '',
        'compare_result': [],
        'compare_date': ''
    })
    all_element_ids = set(prev_param_data.keys()) | set(latest_param_data.keys())
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    for eid in all_element_ids:
        prev_info = prev_param_data.get(eid)
        latest_info = latest_param_data.get(eid)
        prev_fam_type = prev_info['family_and_type'] if prev_info else ''
        curr_fam_type = latest_info['family_and_type'] if latest_info else ''
        prev_cat = prev_info['category'] if prev_info else ''
        curr_cat = latest_info['category'] if latest_info else ''
        key = (eid if eid in prev_param_data else '', eid if eid in latest_param_data else '', prev_fam_type, curr_fam_type, prev_cat, curr_cat)
        group = grouped[key]
        group['previous_element_id'] = eid if eid in prev_param_data else ''
        group['current_element_id'] = eid if eid in latest_param_data else ''
        group['previous_family_and_type'] = prev_fam_type
        group['current_family_and_type'] = curr_fam_type
        group['previous_category'] = prev_cat
        group['current_category'] = curr_cat
        group['compare_date'] = now
        # Instance parameters
        prev_params = prev_info['parameters'] if prev_info else {}
        latest_params = latest_info['parameters'] if latest_info else {}
        prev_param_names = set(prev_params.keys())
        latest_param_names = set(latest_params.keys())
        # New instance parameters
        for pname in latest_param_names - prev_param_names:
            group['compare_result'].append("new parameter add: {}".format(pname))
        # Deleted instance parameters
        for pname in prev_param_names - latest_param_names:
            group['compare_result'].append("parameter delete: {}".format(pname))
        # Changed instance parameters
        for pname in prev_param_names & latest_param_names:
            prev_val = prev_params[pname]
            latest_val = latest_params[pname]
            if prev_val != latest_val:
                group['compare_result'].append("parameter value change: {} ({} -> {})".format(pname, prev_val, latest_val))
        # Type parameters
        prev_type_params = prev_info['type_parameters'] if prev_info and 'type_parameters' in prev_info else {}
        latest_type_params = latest_info['type_parameters'] if latest_info and 'type_parameters' in latest_info else {}
        prev_type_param_names = set(prev_type_params.keys())
        latest_type_param_names = set(latest_type_params.keys())
        # New type parameters
        for pname in latest_type_param_names - prev_type_param_names:
            group['compare_result'].append("new type parameter add: {}".format(pname))
        # Deleted type parameters
        for pname in prev_type_param_names - latest_type_param_names:
            group['compare_result'].append("type parameter delete: {}".format(pname))
        # Changed type parameters
        for pname in prev_type_param_names & latest_type_param_names:
            prev_val = prev_type_params[pname]
            latest_val = latest_type_params[pname]
            if prev_val != latest_val:
                group['compare_result'].append("type parameter value change: {} ({} -> {})".format(pname, prev_val, latest_val))
    # Build final results, only include those with changes
    results = []
    for group in grouped.values():
        if group['compare_result']:
            group['compare_result'] = ', '.join(group['compare_result'])
            results.append(group)
    return results

def compare_element_data(prev_elements_data, latest_elements_data):
    """
    Compares element lists between previous and latest models.
    Returns a list of dicts with keys:
    'previous_element_id', 'current_element_id', 'previous_family_and_type', 'current_family_and_type', 'previous_category', 'current_category', 'compare_result', 'compare_date'
    """
    import datetime
    prev_ids = set(e[0] for e in prev_elements_data)
    latest_ids = set(e[0] for e in latest_elements_data)
    prev_dict = {e[0]: e for e in prev_elements_data}
    latest_dict = {e[0]: e for e in latest_elements_data}
    results = []
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # Deleted elements
    for eid in prev_ids - latest_ids:
        prev_e = prev_dict[eid]
        results.append({
            'previous_element_id': eid,
            'current_element_id': '',
            'previous_family_and_type': prev_e[1],
            'current_family_and_type': '',
            'previous_category': prev_e[2],
            'current_category': '',
            'compare_result': 'element deleted',
            'compare_date': now
        })
    # New elements
    for eid in latest_ids - prev_ids:
        latest_e = latest_dict[eid]
        results.append({
            'previous_element_id': '',
            'current_element_id': eid,
            'previous_family_and_type': '',
            'current_family_and_type': latest_e[1],
            'previous_category': '',
            'current_category': latest_e[2],
            'compare_result': 'new element added',
            'compare_date': now
        })
    return results

def combine_comparison_results(xyz_results, param_results, element_results):
    """
    Combines all comparison results by element id.
    Appends all results, then groups by element id, concatenating compare results by ','.
    If any compare result for an element is 'element deleted', just show 'element deleted'.
    If any compare result for an element is 'new element added', just show 'new element added'.
    Returns a list of dicts with keys:
    'previous_element_id', 'current_element_id', 'previous_family_and_type', 'current_family_and_type', 'previous_category', 'current_category', 'compare_result', 'compare_date'
    """
    from collections import defaultdict
    all_results = []
    # Append all results
    for row in xyz_results:
        all_results.append(row)
    for row in param_results:
        all_results.append(row)
    for row in element_results:
        all_results.append(row)
    # Group by element id (use previous_element_id or current_element_id)
    grouped = defaultdict(list)
    for row in all_results:
        key = row.get('previous_element_id') or row.get('current_element_id')
        grouped[key].append(row)
    # Build final results
    results = []
    for group_rows in grouped.values():
        # Merge fields, prefer non-empty, prefer previous_*
        merged = {
            'previous_element_id': '',
            'current_element_id': '',
            'previous_family_and_type': '',
            'current_family_and_type': '',
            'previous_category': '',
            'current_category': '',
            'compare_result': '',
            'compare_date': ''
        }
        compare_results = []
        for row in group_rows:
            for k in merged:
                if not merged[k] and row.get(k):
                    merged[k] = row.get(k)
            compare_results.append(row.get('compare_result', ''))
        # If any compare_result is 'element deleted', just show that
        if 'element deleted' in compare_results:
            merged['compare_result'] = 'element deleted'
        # If any compare_result is 'new element added', just show that (unless deleted)
        elif 'new element added' in compare_results:
            merged['compare_result'] = 'new element added'
        else:
            merged['compare_result'] = ', '.join([r for r in compare_results if r])
        results.append(merged)
    return results


def ensure_shared_parameters(doc, param_names, categories):
    """
    Ensure shared parameters with the given names exist and are bound to the given categories as instance parameters.
    If a parameter does not exist, create it in the shared parameter file and bind it.
    """
    from Autodesk.Revit.DB import CategorySet, InstanceBinding, Transaction, BuiltInParameterGroup, ExternalDefinitionCreationOptions, SpecTypeId
    app = doc.Application
    shared_param_file = app.OpenSharedParameterFile()
    if not shared_param_file:
        raise Exception("No shared parameter file is set. Please set a shared parameter file in Revit.")
    # Get or create the group for new parameters
    group = None
    if shared_param_file.Groups.Size > 0:
        group = list(shared_param_file.Groups)[0]
    else:
        group = shared_param_file.Groups.Create("00_ModelComparison")
    # Prepare category set
    categories_str = set([str(c) for c in categories])
    all_categories = [cat for cat in doc.Settings.Categories if cat.Name in categories_str]
    cat_set = app.Create.NewCategorySet()
    for cat in all_categories:
        cat_set.Insert(cat)
    # Bind parameters
    binding_map = doc.ParameterBindings
    with Transaction(doc, "Ensure shared parameters") as t:
        t.Start()
        for pname in param_names:
            # Check if parameter already bound
            found = False
            it = binding_map.ForwardIterator()
            it.Reset()
            while it.MoveNext():
                if it.Key.Name == pname:
                    found = True
                    break
            if not found:
                # Check if parameter exists in shared param file, else create
                param_def = None
                for g in shared_param_file.Groups:
                    for d in g.Definitions:
                        if d.Name == pname:
                            param_def = d
                            break
                    if param_def:
                        break
                if not param_def:
                    ext_opt = ExternalDefinitionCreationOptions(pname, SpecTypeId.String.Text)
                    param_def = group.Definitions.Create(ext_opt)
                binding = InstanceBinding(cat_set)
                binding_map.Insert(param_def, binding, BuiltInParameterGroup.PG_DATA)
        t.Commit()
    print("Shared parameter(s) ensured and bound to selected categories.")


# --- Main Workflow ---
if __name__ == "__main__":
    start_time = time.time()
    folder = select_folder()
    if not folder:
        print("No folder selected.")
        script.exit()
    latest_model = select_model("Select the LATEST Revit model", folder)
    if not latest_model:
        print("No latest model selected.")
        script.exit()
    previous_model = select_model("Select the PREVIOUS Revit model", folder)
    if not previous_model:
        print("No previous model selected.")
        script.exit()
    # List all model categories for selection
    categories = get_all_model_categories()
    selected_categories = show_category_selection(categories)
    if not selected_categories:
        print("No categories selected.")
        script.exit()
    analysis_items = show_analysis_item_selection()
    if not analysis_items:
        print("No analysis items selected.")
        script.exit()

    print('--- Timing: Start model extraction ---')
    extract_start = time.time()
    # --- Open previous and latest model in sequence and extract data ---
    from Autodesk.Revit.DB import ModelPathUtils, OpenOptions
    app = revit.doc.Application
    model_path_obj_prev = ModelPathUtils.ConvertUserVisiblePathToModelPath(previous_model)
    model_path_obj_latest = ModelPathUtils.ConvertUserVisiblePathToModelPath(latest_model)
    opts_prev = OpenOptions()
    opts_prev.DetachFromCentralOption = 0
    opts_latest = OpenOptions()
    opts_latest.DetachFromCentralOption = 0

    # Extract from previous model
    doc_prev = app.OpenDocumentFile(model_path_obj_prev, opts_prev)
    try:
        if "XYZ deviation" in analysis_items:
            t0 = time.time()
            prev_xyz_data = extract_xyz_by_category(doc_prev, selected_categories)
            print('Extract prev XYZ: {:.2f}s'.format(time.time() - t0))
        if "Parameter value change" in analysis_items:
            t0 = time.time()
            prev_param_data = extract_parameters_by_category(doc_prev, selected_categories)
            print('Extract prev params: {:.2f}s'.format(time.time() - t0))
        if "Newly/deleted elements" in analysis_items:
            t0 = time.time()
            prev_elements_data = get_elements_by_category(doc_prev, selected_categories)
            print('Extract prev elements: {:.2f}s'.format(time.time() - t0))
    finally:
        doc_prev.Close(False)

    # Initialize comparison result variables
    xyz_comparison_results = []
    param_comparison_results = []
    element_comparison_results = []
    # Extract from latest model
    doc_latest = app.OpenDocumentFile(model_path_obj_latest, opts_latest)
    try:
        if "XYZ deviation" in analysis_items:
            t0 = time.time()
            latest_xyz_data = extract_xyz_by_category(doc_latest, selected_categories)
            print('Extract latest XYZ: {:.2f}s'.format(time.time() - t0))
        if "Parameter value change" in analysis_items:
            t0 = time.time()
            latest_param_data = extract_parameters_by_category(doc_latest, selected_categories)
            print('Extract latest params: {:.2f}s'.format(time.time() - t0))
        if "Newly/deleted elements" in analysis_items:
            t0 = time.time()
            latest_elements_data = get_elements_by_category(doc_latest, selected_categories)
            print('Extract latest elements: {:.2f}s'.format(time.time() - t0))
        print('Model data extracted for selected analysis items.')

        # --- After all individual comparisons, combine results and export ---
        if any(item in analysis_items for item in ["XYZ deviation", "Parameter value change", "Newly/deleted elements"]):
            t0 = time.time()
            # Only run comparisons if data was extracted
            if "XYZ deviation" in analysis_items:
                xyz_comparison_results = compare_xyz_data(prev_xyz_data, latest_xyz_data)
            if "Parameter value change" in analysis_items:
                param_comparison_results = compare_param_data(prev_param_data, latest_param_data)
            if "Newly/deleted elements" in analysis_items:
                element_comparison_results = compare_element_data(prev_elements_data, latest_elements_data)
            combined_results = combine_comparison_results(xyz_comparison_results, param_comparison_results, element_comparison_results)
            print('Combine results: {:.2f}s'.format(time.time() - t0))
            csv_path_combined = os.path.join(folder, "model_comparison_combined_results.csv")
            if combined_results:
                fieldnames = [
                    'previous_element_id',
                    'current_element_id',
                    'previous_family_and_type',
                    'current_family_and_type',
                    'previous_category',
                    'current_category',
                    'compare_result',
                    'compare_date'
                ]
                with open(csv_path_combined, 'w') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')
                    writer.writeheader()
                    for row in combined_results:
                        writer.writerow(row)
                print("Combined model comparison results exported to: {}".format(csv_path_combined))
            else:
                print("No combined model comparison results to export.")

            # --- Ensure project parameters exist before writing ---
            ensure_shared_parameters(doc_latest, ["compare_results", "compare_date"], selected_categories)
            # --- Add compare_results and compare_date parameters to elements in latest model ---
            from Autodesk.Revit.DB import Transaction, BuiltInParameter, StorageType, SaveAsOptions
            t = Transaction(doc_latest, "Add comparison results")
            try:
                t.Start()
                added_count = 0
                for row in combined_results:
                    eid = row.get('current_element_id')
                    if not eid:
                        continue
                    elem = doc_latest.GetElement(ElementId(int(eid)))
                    if not elem:
                        continue
                    # Add or set 'compare_results' parameter
                    param = elem.LookupParameter('compare_results')
                    if param:
                        try:
                            param.Set(str(row.get('compare_result', '')))
                        except Exception:
                            pass
                    # Add or set 'compare_date' parameter
                    date_param = elem.LookupParameter('compare_date')
                    if date_param:
                        try:
                            date_val = str(row.get('compare_date', ''))
                            date_param.Set(date_val)
                        except Exception:
                            pass
                    added_count += 1
                t.Commit()
                print('Added/updated compare_results and compare_date parameters for {} elements.'.format(added_count))
            except Exception as e:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                raise e
            # --- Save the model with new name including current date ---
            import datetime
            save_name = os.path.splitext(os.path.basename(latest_model))[0] + "_compared_" + datetime.datetime.now().strftime("%Y%m%d") + ".rvt"
            save_path = os.path.join(folder, save_name)
            save_options = SaveAsOptions()
            save_options.OverwriteExistingFile = True
            doc_latest.SaveAs(save_path, save_options)
            print('Model saved as: {}'.format(save_path))

    finally:
        doc_latest.Close(False)
    print('--- Extraction total: {:.2f}s ---'.format(time.time() - extract_start))

    # --- Compare XYZ data if applicable ---
    if "XYZ deviation" in analysis_items:
        t0 = time.time()
        xyz_comparison_results = compare_xyz_data(prev_xyz_data, latest_xyz_data)
        elapsed = time.time() - t0
        print('Compare XYZ: {:.2f}s'.format(elapsed))
        # Export results to CSV
        import csv
        csv_path = os.path.join(folder, "xyz_comparison_results.csv")
        if xyz_comparison_results:
            fieldnames = [
                'previous_element_id',
                'current_element_id',
                'previous_family_and_type',
                'current_family_and_type',
                'previous_category',
                'current_category',
                'compare_result',
                'compare_date'
            ]
            with open(csv_path, 'w') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')
                writer.writeheader()
                for row in xyz_comparison_results:
                    writer.writerow(row)
            print("XYZ comparison results exported to: {}".format(csv_path))
        else:
            print("No XYZ comparison results to export.")
        print('--- XYZ comparison and export time: {:.2f}s ---'.format(elapsed))
    # --- Compare parameter data if applicable ---
    if "Parameter value change" in analysis_items:
        t0 = time.time()
        param_comparison_results = compare_param_data(prev_param_data, latest_param_data)
        elapsed = time.time() - t0
        print('Compare params: {:.2f}s'.format(elapsed))
        csv_path_param = os.path.join(folder, "param_comparison_results.csv")
        if param_comparison_results:
            fieldnames = [
                'previous_element_id',
                'current_element_id',
                'previous_family_and_type',
                'current_family_and_type',
                'previous_category',
                'current_category',
                'compare_result',
                'compare_date'
            ]
            with open(csv_path_param, 'w') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')
                writer.writeheader()
                for row in param_comparison_results:
                    writer.writerow(row)
            print("Parameter comparison results exported to: {}".format(csv_path_param))
        else:
            print("No parameter comparison results to export.")
        print('--- Parameter comparison and export time: {:.2f}s ---'.format(elapsed))
    # --- Compare element data if applicable ---
    if "Newly/deleted elements" in analysis_items:
        t0 = time.time()
        element_comparison_results = compare_element_data(prev_elements_data, latest_elements_data)
        elapsed = time.time() - t0
        print('Compare elements: {:.2f}s'.format(elapsed))
        csv_path_elem = os.path.join(folder, "element_comparison_results.csv")
        if element_comparison_results:
            fieldnames = [
                'previous_element_id',
                'current_element_id',
                'previous_family_and_type',
                'current_family_and_type',
                'previous_category',
                'current_category',
                'compare_result',
                'compare_date'
            ]
            with open(csv_path_elem, 'w') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')
                writer.writeheader()
                for row in element_comparison_results:
                    writer.writerow(row)
            print("Element comparison results exported to: {}".format(csv_path_elem))
        else:
            print("No element comparison results to export.")
        print('--- Element comparison and export time: {:.2f}s ---'.format(elapsed))
    print("--- Total script time: {:.2f}s ---".format(time.time() - start_time))
    print("Comparison complete.")

    # --- Summary printout ---
    def extract_summary_stats(combined_results):
        xy_move_count = 0
        z_move_count = 0
        new_param_count = 0
        new_param_list = set()
        del_param_count = 0
        del_param_list = set()
        param_value_change_count = 0
        param_value_change_list = set()
        new_elem_count = 0
        del_elem_count = 0
        for row in combined_results:
            result = row.get('compare_result', '')
            # 1. XY coordination move
            if "XY coordination move" in result:
                xy_move_count += 1
            # 2. Z coordination move
            if "Z coordination move" in result:
                z_move_count += 1
            # 3. new parameter add
            if "new parameter add:" in result:
                for part in result.split(','):
                    if "new parameter add:" in part:
                        pname = part.split(":",1)[1].strip()
                        new_param_list.add(pname)
                        new_param_count += 1
            # 4. parameter delete
            if "parameter delete:" in result:
                for part in result.split(','):
                    if "parameter delete:" in part:
                        pname = part.split(":",1)[1].strip()
                        del_param_list.add(pname)
                        del_param_count += 1
            # 5. parameter value change
            if "parameter value change:" in result:
                for part in result.split(','):
                    if "parameter value change:" in part:
                        pname = part.split(":",1)[1].split("(")[0].strip()
                        param_value_change_list.add(pname)
                        param_value_change_count += 1
            # 6. new element added
            if result == "new element added":
                new_elem_count += 1
            # 7. element deleted
            if result == "element deleted":
                del_elem_count += 1
        return {
            'xy_move_count': xy_move_count,
            'z_move_count': z_move_count,
            'new_param_count': len(new_param_list),
            'new_param_list': sorted(new_param_list),
            'del_param_count': len(del_param_list),
            'del_param_list': sorted(del_param_list),
            'param_value_change_count': len(param_value_change_list),
            'param_value_change_list': sorted(param_value_change_list),
            'new_elem_count': new_elem_count,
            'del_elem_count': del_elem_count
        }

    summary = extract_summary_stats(combined_results)
    print("\n--- Model Comparison Summary ---")
    print("1. Number of XY coordination move: {}".format(summary['xy_move_count']))
    print("2. Number of Z coordination move: {}".format(summary['z_move_count']))
    print("3. Number of new parameter added: {}".format(summary['new_param_count']))
    if summary['new_param_list']:
        print("   List of new parameters added:")
        for pname in summary['new_param_list']:
            print("     - {}".format(pname))
    print("4. Number of parameter deleted: {}".format(summary['del_param_count']))
    if summary['del_param_list']:
        print("   List of deleted parameters:")
        for pname in summary['del_param_list']:
            print("     - {}".format(pname))
    print("5. Number of parameter value change: {}".format(summary['param_value_change_count']))
    if summary['param_value_change_list']:
        print("   List of parameter value changes:")
        for pname in summary['param_value_change_list']:
            print("     - {}".format(pname))
    print("6. Number of new element added: {}".format(summary['new_elem_count']))
    print("7. Number of element deleted: {}".format(summary['del_elem_count']))

    # --- Summary printout ---
    def extract_summary_stats_by_category(combined_results, folder):
        summary_by_cat = {}
        for row in combined_results:
            cat = row.get('current_category') or row.get('previous_category') or 'Unknown'
            if cat not in summary_by_cat:
                summary_by_cat[cat] = {
                    'xy_move_count': 0,
                    'z_move_count': 0,
                    'new_param_list': set(),
                    'del_param_list': set(),
                    'param_value_change_list': set(),
                    'new_type_param_list': set(),
                    'del_type_param_list': set(),
                    'type_param_value_change_list': set(),
                    'new_elem_count': 0,
                    'del_elem_count': 0
                }
            result = row.get('compare_result', '')
            # 1. XY coordination move
            if "XY coordination move" in result:
                summary_by_cat[cat]['xy_move_count'] += 1
            # 2. Z coordination move
            if "Z coordination move" in result:
                summary_by_cat[cat]['z_move_count'] += 1
            # 3. new parameter add
            if "new parameter add:" in result:
                for part in result.split(','):
                    if "new parameter add:" in part:
                        pname = part.split(":",1)[1].strip()
                        summary_by_cat[cat]['new_param_list'].add(pname)
            # 4. parameter delete
            if "parameter delete:" in result:
                for part in result.split(','):
                    if "parameter delete:" in part:
                        pname = part.split(":",1)[1].strip()
                        summary_by_cat[cat]['del_param_list'].add(pname)
            # 5. parameter value change
            if "parameter value change:" in result:
                for part in result.split(','):
                    if "parameter value change:" in part:
                        pname = part.split(":",1)[1].split("(")[0].strip()
                        summary_by_cat[cat]['param_value_change_list'].add(pname)
            # 6. new type parameter add
            if "new type parameter add:" in result:
                for part in result.split(','):
                    if "new type parameter add:" in part:
                        pname = part.split(":",1)[1].strip()
                        summary_by_cat[cat]['new_type_param_list'].add(pname)
            # 7. type parameter delete
            if "type parameter delete:" in result:
                for part in result.split(','):
                    if "type parameter delete:" in part:
                        pname = part.split(":",1)[1].strip()
                        summary_by_cat[cat]['del_type_param_list'].add(pname)
            # 8. type parameter value change
            if "type parameter value change:" in result:
                for part in result.split(','):
                    if "type parameter value change:" in part:
                        pname = part.split(":",1)[1].split("(")[0].strip()
                        summary_by_cat[cat]['type_param_value_change_list'].add(pname)
            # 9. new element added
            if result == "new element added":
                summary_by_cat[cat]['new_elem_count'] += 1
            # 10. element deleted
            if result == "element deleted":
                summary_by_cat[cat]['del_elem_count'] += 1
        # Convert sets to sorted lists and add counts
        for cat, stats in summary_by_cat.items():
            stats['new_param_count'] = len(stats['new_param_list'])
            stats['new_param_list'] = sorted(stats['new_param_list'])
            stats['del_param_count'] = len(stats['del_param_list'])
            stats['del_param_list'] = sorted(stats['del_param_list'])
            stats['param_value_change_count'] = len(stats['param_value_change_list'])
            stats['param_value_change_list'] = sorted(stats['param_value_change_list'])
            stats['new_type_param_count'] = len(stats['new_type_param_list'])
            stats['new_type_param_list'] = sorted(stats['new_type_param_list'])
            stats['del_type_param_count'] = len(stats['del_type_param_list'])
            stats['del_type_param_list'] = sorted(stats['del_type_param_list'])
            stats['type_param_value_change_count'] = len(stats['type_param_value_change_list'])
            stats['type_param_value_change_list'] = sorted(stats['type_param_value_change_list'])
        # Export to CSV
        csv_path_summary_cat = os.path.join(folder, "model_comparison_summary_by_category.csv")
        with open(csv_path_summary_cat, 'w') as csvfile:
            fieldnames = [
                'category',
                'xy_move_count',
                'z_move_count',
                'new_param_count',
                'new_param_list',
                'del_param_count',
                'del_param_list',
                'param_value_change_count',
                'param_value_change_list',
                'new_type_param_count',
                'new_type_param_list',
                'del_type_param_count',
                'del_type_param_list',
                'type_param_value_change_count',
                'type_param_value_change_list',
                'new_elem_count',
                'del_elem_count'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')
            writer.writeheader()
            for cat, stats in summary_by_cat.items():
                row = stats.copy()
                row['category'] = cat
                row['new_param_list'] = ', '.join(row['new_param_list'])
                row['del_param_list'] = ', '.join(row['del_param_list'])
                row['param_value_change_list'] = ', '.join(row['param_value_change_list'])
                row['new_type_param_list'] = ', '.join(row['new_type_param_list'])
                row['del_type_param_list'] = ', '.join(row['del_type_param_list'])
                row['type_param_value_change_list'] = ', '.join(row['type_param_value_change_list'])
                writer.writerow(row)
        print("Summary by category exported to: {}".format(csv_path_summary_cat))
        return summary_by_cat

    summary_by_cat = extract_summary_stats_by_category(combined_results, folder)
    print("\n--- Model Comparison Summary by Category ---")
    for cat, summary in summary_by_cat.items():
        print("\nCategory: {}".format(cat))
        print("  1. Number of XY coordination move: {}".format(summary['xy_move_count']))
        print("  2. Number of Z coordination move: {}".format(summary['z_move_count']))
        print("  3. Number of new parameter added: {}".format(summary['new_param_count']))
        if summary['new_param_list']:
            print("     List of new parameters added:")
            for pname in summary['new_param_list']:
                print("       - {}".format(pname))
        print("  4. Number of parameter deleted: {}".format(summary['del_param_count']))
        if summary['del_param_list']:
            print("     List of deleted parameters:")
            for pname in summary['del_param_list']:
                print("       - {}".format(pname))
        print("  5. Number of parameter value change: {}".format(summary['param_value_change_count']))
        if summary['param_value_change_list']:
            print("     List of parameter value changes:")
            for pname in summary['param_value_change_list']:
                print("       - {}".format(pname))
        print("  6. Number of new element added: {}".format(summary['new_elem_count']))
        print("  7. Number of element deleted: {}".format(summary['del_elem_count']))

    # --- Export summary_by_cat to CSV ---
    csv_path_summary_cat = os.path.join(folder, "model_comparison_summary_by_category.csv")
    with open(csv_path_summary_cat, 'w') as csvfile:
        fieldnames = [
            'category',
            'xy_move_count',
            'z_move_count',
            'new_param_count',
            'new_param_list',
            'del_param_count',
            'del_param_list',
            'param_value_change_count',
            'param_value_change_list',
            'new_type_param_count',
            'new_type_param_list',
            'del_type_param_count',
            'del_type_param_list',
            'type_param_value_change_count',
            'type_param_value_change_list',
            'new_elem_count',
            'del_elem_count'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')
        writer.writeheader()
        for cat, stats in summary_by_cat.items():
            row = stats.copy()
            row['category'] = cat
            row['new_param_list'] = ', '.join(row['new_param_list'])
            row['del_param_list'] = ', '.join(row['del_param_list'])
            row['param_value_change_list'] = ', '.join(row['param_value_change_list'])
            row['new_type_param_list'] = ', '.join(row['new_type_param_list'])
            row['del_type_param_list'] = ', '.join(row['del_type_param_list'])
            row['type_param_value_change_list'] = ', '.join(row['type_param_value_change_list'])
            writer.writerow(row)
    print("Summary by category exported to: {}".format(csv_path_summary_cat))

