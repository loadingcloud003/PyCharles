# -*- coding: utf-8 -*-
# pyRevit addin script for reading model_comparison_summary_by_category.csv, creating filters, and printing PDF
import csv
import os
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import FolderBrowserDialog, OpenFileDialog, Form, Label, ListBox, Button, DialogResult, CheckedListBox, ColorDialog, GroupBox
from System.Windows.Forms import SelectionMode
clr.AddReference('System.Drawing')
import System.Drawing

CSV_FILENAME = "model_comparison_summary_by_category.csv"

# 1. Select folder and model
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

# 2. Read and process CSV

def read_summary_csv(csv_path):
    rows = []
    with open(csv_path, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            rows.append(row)
    return rows

def group_results_by_category_and_type(summary_rows):
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
        self.Text = "Create Filters by Category and Result Type"
        self.Width = 1200
        self.Height = 900
        self.selected_filters = []  # (cat, rtype, pname, color)
        self.grouped = grouped
        self.color_map = {}  # (cat, item) -> color
        y = 10
        for cat, rtypes in grouped.items():
            gb = GroupBox()
            gb.Text = cat
            gb.Top = y
            gb.Left = 10
            gb.Width = 1100
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
            self.item_buttons = []  # Store buttons for color selection
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
                elif rtype in ["parameter value change", "type parameter value change"]:
                    for pname in items:
                        result_items.append("{}: {}".format(rtype, pname))
            # Add items to listbox
            for idx, item in enumerate(result_items):
                lb.Items.Add(item)
                btn = Button()
                btn.Text = "Color"
                btn.Top = gy + idx * 28
                btn.Left = lb.Left + lb.Width + 10
                btn.Width = 60
                btn.Tag = (cat, item)
                btn.BackColor = System.Drawing.Color.White
                btn.Click += self.select_item_color
                gb.Controls.Add(btn)
                self.color_map[(cat, item)] = System.Drawing.Color.White
                self.item_buttons.append(btn)
            gb.Controls.Add(lb)
            self.Controls.Add(gb)
            y += gb.Height + 10
        ok_btn = Button()
        ok_btn.Text = "Create Filters"
        ok_btn.Top = y + 20
        ok_btn.Left = 220
        ok_btn.Width = 120
        ok_btn.DialogResult = DialogResult.OK
        self.Controls.Add(ok_btn)
        self.AcceptButton = ok_btn
    def select_item_color(self, sender, args):
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
    csv_path = os.path.join(folder, CSV_FILENAME)
    if not os.path.exists(csv_path):
        print("CSV file not found: {}".format(csv_path))
        return
    summary_rows = read_summary_csv(csv_path)
    grouped = group_results_by_category_and_type(summary_rows)
    # 3. Dialog for filter creation
    filter_form = FilterDialog(grouped)
    if filter_form.ShowDialog() == DialogResult.OK:
        # 4. Select views to apply filters
        # Replace with actual Revit view objects in real usage
        views = ["View 1", "View 2", "View 3"]
        view_form = ViewSelectForm(views)
        if view_form.ShowDialog() == DialogResult.OK:
            selected_views = [views[i] for i in range(view_form.clb.Items.Count) if view_form.clb.GetItemChecked(i)]
            print("Apply filters to views:", selected_views)
            # TODO: Apply filters to selected views using Revit API
        # 5. Select views/sheets to print PDF
        sheets = ["Sheet A", "Sheet B"]
        print_form = PrintSelectForm(views, sheets)
        if print_form.ShowDialog() == DialogResult.OK:
            selected_views_to_print = [views[i] for i in range(print_form.clb_views.Items.Count) if print_form.clb_views.GetItemChecked(i)]
            selected_sheets_to_print = [sheets[i] for i in range(print_form.clb_sheets.Items.Count) if print_form.clb_sheets.GetItemChecked(i)]
            print("Print PDF for views:", selected_views_to_print)
            print("Print PDF for sheets:", selected_sheets_to_print)
            # TODO: Implement actual PDF printing logic

if __name__ == "__main__":
    main()