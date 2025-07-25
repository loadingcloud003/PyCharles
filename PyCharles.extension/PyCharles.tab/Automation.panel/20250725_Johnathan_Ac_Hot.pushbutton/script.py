from pyrevit import forms
# -*- coding: utf-8 -*-
from pyrevit import revit, script
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, Part, View, Element
import csv
import os

def get_all_parts_in_current_view(doc, view):
    # Get all Part elements visible in the current view
    return [e for e in FilteredElementCollector(doc, view.Id).OfClass(Part)]

def get_reference_element(doc, part):
    # Recursively get the parent element of the part until a non-Part element is found
    current_elem = part
    while True:
        if hasattr(current_elem, 'GetSourceElementIds'):
            source_ids = current_elem.GetSourceElementIds()
            if source_ids and len(source_ids) > 0:
                eid = source_ids[0].HostElementId
                parent_elem = doc.GetElement(eid)
                if parent_elem is None:
                    return None
                if isinstance(parent_elem, Part):
                    current_elem = parent_elem
                    continue
                else:
                    return parent_elem
            else:
                return None
        else:
            return None

def get_element_info(elem):
    # Get family name, type name, category, and all parameters as dict
    fam_name = ''
    type_name = ''
    cat_name = ''
    param_dict = {}
    if elem is None:
        return fam_name, type_name, cat_name, param_dict
    try:
        if hasattr(elem, 'Symbol') and elem.Symbol:
            fam_name = elem.Symbol.Family.Name
            type_name = elem.Symbol.Name
        elif hasattr(elem, 'FamilyName'):
            fam_name = elem.FamilyName
        if hasattr(elem, 'Name'):
            type_name = elem.Name
        if elem.Category:
            cat_name = elem.Category.Name
        for param in elem.Parameters:
            try:
                pname = param.Definition.Name
                pval = param.AsValueString() if param.StorageType != 4 else str(param.AsElementId().IntegerValue)
                param_dict[pname] = pval
            except Exception:
                pass
    except Exception:
        pass
    return fam_name, type_name, cat_name, param_dict

def export_parts_and_references_to_csv(doc, view, out_csv_path):
    parts = get_all_parts_in_current_view(doc, view)
    rows = []
    all_part_param_names = set()
    all_ref_param_names = set()
    part_data = []
    for part in parts:
        part_id = part.Id.IntegerValue
        part_fam, part_type, part_cat, part_params = get_element_info(part)
        ref_elem = get_reference_element(doc, part)
        ref_elem_id = ref_elem.Id.IntegerValue if ref_elem else ''
        ref_fam, ref_type, ref_cat, ref_params = get_element_info(ref_elem)
        all_part_param_names.update(part_params.keys())
        all_ref_param_names.update(ref_params.keys())
        part_data.append({
            'part_element_id': part_id,
            'part_family_name': part_fam,
            'part_type_name': part_type,
            'reference_element_id': ref_elem_id,
            'reference_family_name': ref_fam,
            'reference_type_name': ref_type,
            'reference_category': ref_cat,
            'part_params': part_params,
            'reference_params': ref_params
        })

    # Show parameter selection dialog
    part_param_list = sorted(all_part_param_names)
    ref_param_list = sorted(all_ref_param_names)
    selected_part_params = forms.SelectFromList.show(part_param_list, multiselect=True, title='Select Part Parameters to Export', button_name='Export')
    if not selected_part_params:
        script.exit('No Part parameters selected.')
    selected_ref_params = forms.SelectFromList.show(ref_param_list, multiselect=True, title='Select Reference Element Parameters to Export', button_name='Export')
    if not selected_ref_params:
        script.exit('No Reference Element parameters selected.')
    # Build fieldnames
    part_param_cols = ['part_param_' + k for k in selected_part_params]
    ref_param_cols = ['reference_param_' + k for k in selected_ref_params]
    fieldnames = [
        'part_element_id', 'part_family_name', 'part_type_name',
        'reference_element_id', 'reference_family_name', 'reference_type_name', 'reference_category'
    ] + part_param_cols + ref_param_cols
    # Build rows for CSV
    for data in part_data:
        row = {
            'part_element_id': data['part_element_id'],
            'part_family_name': data['part_family_name'],
            'part_type_name': data['part_type_name'],
            'reference_element_id': data['reference_element_id'],
            'reference_family_name': data['reference_family_name'],
            'reference_type_name': data['reference_type_name'],
            'reference_category': data['reference_category']
        }
        for k in selected_part_params:
            row['part_param_' + k] = data['part_params'].get(k, '')
        for k in selected_ref_params:
            row['reference_param_' + k] = data['reference_params'].get(k, '')
        rows.append(row)
    # Write to CSV with Excel-compatible quoting
    with open(out_csv_path, 'w') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n', quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
    print('Exported {} part(s) info to: {}'.format(len(rows), out_csv_path))

# --- Main pyRevit command ---
def main():
    doc = revit.doc
    view = doc.ActiveView
    # Use system temp directory for output
    import tempfile
    folder = tempfile.gettempdir()
    out_csv = os.path.join(folder, 'parts_and_references.csv')
    export_parts_and_references_to_csv(doc, view, out_csv)

if __name__ == '__main__':
    main()
