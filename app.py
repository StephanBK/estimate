from flask import Flask, request, redirect, url_for, send_file, Response, session
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import os
import io
import csv
import datetime
import numpy as np
import json

app = Flask(__name__)
app.secret_key = "Ti5om4gm!"  # Replace with a secure random key

# Database Configuration
app.config[
    'SQLALCHEMY_DATABASE_URI'] = 'postgresql://u7vukdvn20pe3c:p918802c410825b956ccf24c5af8d168b4d9d69e1940182bae9bd8647eb606845@cb5ajfjosdpmil.cluster-czrs8kj4isg7.us-east-1.rds.amazonaws.com:5432/dcobttk99a5sie'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Initialize the database
db = SQLAlchemy(app)


# Define the Material model
class Material(db.Model):
    __tablename__ = 'materials'
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.Integer, nullable=False)
    nickname = db.Column(db.String(100), nullable=False)
    status = db.Column(db.String(50))
    description = db.Column(db.Text)
    cost = db.Column(db.Float, nullable=False)
    cost_unit = db.Column(db.String(50))
    in_inventory = db.Column(db.Boolean)
    moq = db.Column(db.Integer)
    moq_cost = db.Column(db.Float)
    yield_value = db.Column('yield', db.Float)
    yield_unit = db.Column(db.String(50))
    yield_unit_2 = db.Column(db.String(50))
    yield_cost = db.Column(db.Float)
    yield_unit_3 = db.Column(db.String(50))
    last_updated = db.Column(db.DateTime)
    manufacturer = db.Column(db.String(100))
    supplier = db.Column(db.String(100))


# Helper: Recursively convert NumPy int64 to native types
def make_serializable(obj):
    if isinstance(obj, dict):
        return {k: make_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [make_serializable(x) for x in obj]
    elif isinstance(obj, np.int64):
        return int(obj)
    else:
        return obj


# Helper: get or initialize current_project in session
def get_current_project():
    cp = session.get("current_project")
    if cp is None:
        cp = {}
        session["current_project"] = cp
    return cp


# Helper: save current_project after converting to serializable types
def save_current_project(cp):
    session["current_project"] = make_serializable(cp)


# Path to the CSV template
TEMPLATE_PATH = 'estimation_template_template.csv'

# Common CSS for consistent styling
common_css = """
    @import url('https://fonts.googleapis.com/css2?family=Exo+2:wght@400;600&display=swap');
    body { background-color: #121212; color: #ffffff; font-family: 'Exo 2', sans-serif; padding: 20px; }
    .container { background-color: #1e1e1e; padding: 20px; border-radius: 10px; max-width: 900px; margin: auto; box-shadow: 0 0 15px rgba(0,128,128,0.5); }
    input, select, button { width: 100%; padding: 10px; margin: 5px 0; border: none; border-radius: 5px; background-color: #333; color: #fff; }
    button { background-color: #008080; cursor: pointer; transition: background-color 0.3s ease; }
    button:hover { background-color: #00a0a0; }
    a { color: #00a0a0; text-decoration: none; font-weight: 600; }
    a:hover { text-decoration: underline; }
    h2, p { text-align: center; }
    .summary-table, .data-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    .summary-table th, .summary-table td, .data-table th, .data-table td { border: 1px solid #444; padding: 8px; text-align: center; }
    .summary-table th { background-color: #008080; color: white; }
    .data-table th { background-color: #00a0a0; color: white; }
    tr:nth-child(even) { background-color: #1e1e1f; }
    tr:nth-child(odd) { background-color: #2b2b2b; }
"""


# =========================
# INDEX PAGE
# =========================
@app.route('/', methods=['GET', 'POST'])
def index():
    cp = get_current_project()
    if request.method == 'POST':
        cp['project_name'] = request.form['project_name']
        cp['project_number'] = request.form['project_number']
        file = request.files['file']
        if file:
            file_path = os.path.join('uploads', file.filename)
            os.makedirs('uploads', exist_ok=True)
            file.save(file_path)
            cp['file_path'] = file_path
            save_current_project(cp)
            return redirect(url_for('summary'))
    return f"""
    <html>
      <head>
         <title>Estimation Tool</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Project Input Form</h2>
            <form method="POST" enctype="multipart/form-data">
               <label for="project_name">Project Name:</label>
               <input type="text" id="project_name" name="project_name" required>
               <label for="project_number">Project Number:</label>
               <input type="text" id="project_number" name="project_number" required>
               <label for="file">Upload Filled Template (CSV):</label>
               <input type="file" id="file" name="file" accept=".csv" required>
               <button type="submit">Submit</button>
            </form>
            <p><a href="/download-template">Download Template CSV</a></p>
         </div>
      </body>
    </html>
    """


@app.route('/download-template')
def download_template():
    return send_file(TEMPLATE_PATH, as_attachment=True)


# =========================
# SUMMARY PAGE
# =========================
@app.route('/summary')
def summary():
    cp = get_current_project()
    file_path = cp.get('file_path')
    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        return f"<h2 style='color: red;'>Error reading the file: {e}</h2>"
    if "Type" in df.columns:
        df_swr = df[df["Type"].str.upper() == "SWR"]
        df_igr = df[df["Type"].str.upper() == "IGR"]
    else:
        df_swr = df
        df_igr = pd.DataFrame()

    def compute_totals(df_subset):
        df_subset['Area (sq in)'] = df_subset['VGA Width in'] * df_subset['VGA Height in']
        df_subset['Total Area (sq ft)'] = (df_subset['Area (sq in)'] * df_subset['Qty']) / 144
        df_subset['Perimeter (in)'] = 2 * (df_subset['VGA Width in'] + df_subset['VGA Height in'])
        df_subset['Total Perimeter (ft)'] = (df_subset['Perimeter (in)'] * df_subset['Qty']) / 12
        df_subset['Total Vertical (ft)'] = (df_subset['VGA Height in'] * df_subset['Qty'] * 2) / 12
        df_subset['Total Horizontal (ft)'] = (df_subset['VGA Width in'] * df_subset['Qty'] * 2) / 12
        total_area = df_subset['Total Area (sq ft)'].sum()
        total_perimeter = df_subset['Total Perimeter (ft)'].sum()
        total_vertical = df_subset['Total Vertical (ft)'].sum()
        total_horizontal = df_subset['Total Horizontal (ft)'].sum()
        total_quantity = df_subset['Qty'].sum()
        return total_area, total_perimeter, total_vertical, total_horizontal, total_quantity

    swr_area, swr_perimeter, swr_vertical, swr_horizontal, swr_quantity = compute_totals(df_swr)
    igr_area, igr_perimeter, igr_vertical, igr_horizontal, igr_quantity = compute_totals(
        df_igr) if not df_igr.empty else (0, 0, 0, 0, 0)

    cp['swr_total_area'] = swr_area
    cp['swr_total_perimeter'] = swr_perimeter
    cp['swr_total_vertical_ft'] = swr_vertical
    cp['swr_total_horizontal_ft'] = swr_horizontal
    cp['swr_total_quantity'] = swr_quantity
    cp['igr_total_area'] = igr_area
    cp['igr_total_perimeter'] = igr_perimeter
    cp['igr_total_vertical_ft'] = igr_vertical
    cp['igr_total_horizontal_ft'] = igr_horizontal
    cp['igr_total_quantity'] = igr_quantity
    save_current_project(cp)

    summary_html = f"""
    <html>
      <head>
         <title>Project Summary</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Project Summary</h2>
            <p><strong>Project Name:</strong> {cp.get('project_name')}</p>
            <p><strong>Project Number:</strong> {cp.get('project_number')}</p>
            <h3>SWR Totals</h3>
            <table class="summary-table">
              <tr><th>Metric</th><th>Value</th></tr>
              <tr><td>Total Area (sq ft)</td><td>{swr_area:.2f}</td></tr>
              <tr><td>Total Perimeter (ft)</td><td>{swr_perimeter:.2f}</td></tr>
              <tr><td>Total Quantity</td><td>{swr_quantity:.2f}</td></tr>
              <tr><td>Total Vertical (ft)</td><td>{swr_vertical:.2f}</td></tr>
              <tr><td>Total Horizontal (ft)</td><td>{swr_horizontal:.2f}</td></tr>
            </table>
            <h3>IGR Totals</h3>
            <table class="summary-table">
              <tr><th>Metric</th><th>Value</th></tr>
              <tr><td>Total Area (sq ft)</td><td>{igr_area:.2f}</td></tr>
              <tr><td>Total Perimeter (ft)</td><td>{igr_perimeter:.2f}</td></tr>
              <tr><td>Total Quantity</td><td>{igr_quantity:.2f}</td></tr>
              <tr><td>Total Vertical (ft)</td><td>{igr_vertical:.2f}</td></tr>
              <tr><td>Total Horizontal (ft)</td><td>{igr_horizontal:.2f}</td></tr>
            </table>
            <button onclick="window.location.href='/'">Start New Project</button>
            <button onclick="window.location.href='/materials'">Next: SWR Materials</button>
         </div>
      </body>
    </html>
    """
    return summary_html


# =========================
# MATERIALS PAGE (with Updated Aluminum Split, Foam Baffle Split, and Tape Option renamed)
# =========================
@app.route('/materials', methods=['GET', 'POST'])
def materials():
    cp = get_current_project()
    try:
        # Glass from Category 15
        materials_glass = Material.query.filter_by(category=15).all()
        # Extrusions and Retainer from Category 1 (same pool)
        materials_extrusions = Material.query.filter_by(category=1).all()
        # Other categories remain unchanged
        materials_glazing = Material.query.filter_by(category=2).all()
        materials_gaskets = Material.query.filter_by(category=3).all()
        materials_corner_keys = Material.query.filter_by(category=4).all()
        materials_dual_lock = Material.query.filter_by(category=5).all()
        # Foam Baffle: Two separate items from Category 6 – one for Top/Head and one for Bottom/Sill
        materials_foam_baffle = Material.query.filter_by(category=6).all()
        materials_glass_protection = Material.query.filter_by(category=7).all()
        # For Tape, we keep the original name "Tape" (Category 10)
        materials_tape = Material.query.filter_by(category=10).all()
        materials_head_retainers = Material.query.filter_by(category=17).all()
    except Exception as e:
        return f"<h2 style='color: red;'>Error fetching materials: {e}</h2>"

    if request.method == 'POST':
        # Retrieve yield factors
        try:
            yield_cat15 = float(request.form.get('yield_cat15', 0.97))
        except:
            yield_cat15 = 0.97
        try:
            yield_extrusions = float(request.form.get('yield_extrusions', 0.75))
        except:
            yield_extrusions = 0.75
        try:
            yield_retainer = float(request.form.get('yield_retainer', 0.75))
        except:
            yield_retainer = 0.75
        try:
            yield_cat2 = float(request.form.get('yield_cat2', 0.91))
        except:
            yield_cat2 = 0.91
        try:
            yield_cat3 = float(request.form.get('yield_cat3', 0.91))
        except:
            yield_cat3 = 0.91
        try:
            yield_cat4 = float(request.form.get('yield_cat4', 0.91))
        except:
            yield_cat4 = 0.91
        try:
            yield_cat5 = float(request.form.get('yield_cat5', 0.91))
        except:
            yield_cat5 = 0.91
        try:
            yield_cat6 = float(request.form.get('yield_cat6', 0.91))
        except:
            yield_cat6 = 0.91
        try:
            yield_cat7 = float(request.form.get('yield_cat7', 0.91))
        except:
            yield_cat7 = 0.91
        try:
            yield_cat10 = float(request.form.get('yield_cat10', 0.91))
        except:
            yield_cat10 = 0.91
        try:
            yield_cat17 = float(request.form.get('yield_cat17', 0.91))
        except:
            yield_cat17 = 0.91

        # For Extrusions and Retainer (Category 1)
        selected_extrusions = request.form.get('material_extrusions')
        retainer_option = request.form.get('retainer_option')
        selected_retainer = request.form.get('material_retainer')
        # For Glass Protection
        glass_protection_side = request.form.get('glass_protection_side')

        # Other selections
        selected_glass = request.form.get('material_glass')
        selected_glazing = request.form.get('material_glazing')
        selected_gaskets = request.form.get('material_gaskets')
        selected_corner_keys = request.form.get('material_corner_keys')
        selected_dual_lock = request.form.get('material_dual_lock')
        # Foam Baffle selections:
        selected_foam_baffle_top = request.form.get('material_foam_baffle')
        selected_foam_baffle_bottom = request.form.get('material_foam_baffle_bottom')
        # For Tape, keep the label "Tape" but the dropdown will be titled "Retainer Attachment Option" with three choices.
        retainer_attachment_option = request.form.get('retainer_attachment_option')
        selected_tape = request.form.get('material_tape')
        selected_head_retainers = request.form.get('material_head_retainers')

        mat_glass = Material.query.get(selected_glass) if selected_glass else None
        mat_extrusions = Material.query.get(selected_extrusions) if selected_extrusions else None
        mat_retainer = Material.query.get(selected_retainer) if selected_retainer else None
        mat_glazing = Material.query.get(selected_glazing) if selected_glazing else None
        mat_gaskets = Material.query.get(selected_gaskets) if selected_gaskets else None
        mat_corner_keys = Material.query.get(selected_corner_keys) if selected_corner_keys else None
        # For Dual Lock, cost = Total quantity * cost / yield
        mat_dual_lock = Material.query.get(selected_dual_lock) if selected_dual_lock else None
        mat_foam_baffle_top = Material.query.get(selected_foam_baffle_top) if selected_foam_baffle_top else None
        mat_foam_baffle_bottom = Material.query.get(
            selected_foam_baffle_bottom) if selected_foam_baffle_bottom else None
        mat_glass_protection = Material.query.get(request.form.get('material_glass_protection')) if request.form.get(
            'material_glass_protection') else None
        mat_tape = Material.query.get(selected_tape) if selected_tape else None
        mat_head_retainers = Material.query.get(selected_head_retainers) if selected_head_retainers else None

        total_area = cp.get('swr_total_area', 0)
        total_perimeter = cp.get('swr_total_perimeter', 0)
        total_vertical = cp.get('swr_total_vertical_ft', 0)
        total_horizontal = cp.get('swr_total_horizontal_ft', 0)
        total_quantity = cp.get('swr_total_quantity', 0)

        cost_glass = (total_area * mat_glass.cost) / yield_cat15 if mat_glass else 0
        cost_extrusions = (total_perimeter * mat_extrusions.cost) / yield_extrusions if mat_extrusions else 0
        if retainer_option == "head_retainer":
            cost_retainer = (0.5 * total_horizontal * mat_retainer.cost) / yield_retainer if mat_retainer else 0
        elif retainer_option == "head_and_sill":
            cost_retainer = (total_horizontal * mat_retainer.cost * yield_retainer) if mat_retainer else 0
        else:
            cost_retainer = 0
        cost_glazing = (total_perimeter * mat_glazing.cost) / yield_cat2 if mat_glazing else 0
        cost_gaskets = (total_vertical * mat_gaskets.cost) / yield_cat3 if mat_gaskets else 0
        cost_corner_keys = (total_quantity * 4 * mat_corner_keys.cost) / yield_cat4 if mat_corner_keys else 0
        cost_dual_lock = (total_quantity * mat_dual_lock.cost) / yield_cat5 if mat_dual_lock else 0
        if mat_foam_baffle_top:
            cost_foam_baffle_top = (0.5 * total_horizontal * mat_foam_baffle_top.cost) / yield_cat6
        else:
            cost_foam_baffle_top = 0
        if mat_foam_baffle_bottom:
            cost_foam_baffle_bottom = (0.5 * total_horizontal * mat_foam_baffle_bottom.cost) / yield_cat6
        else:
            cost_foam_baffle_bottom = 0
        if glass_protection_side == "double":
            cost_glass_protection = (
                                                total_area * mat_glass_protection.cost * 2) / yield_cat7 if mat_glass_protection else 0
        elif glass_protection_side == "none":
            cost_glass_protection = 0
        else:
            cost_glass_protection = (total_area * mat_glass_protection.cost) / yield_cat7 if mat_glass_protection else 0
        if mat_tape:
            if retainer_attachment_option == 'head_retainer':
                cost_tape = ((total_horizontal / 2) * mat_tape.cost) / yield_cat10
            elif retainer_attachment_option == 'head_sill':
                cost_tape = (total_horizontal * mat_tape.cost) / yield_cat10
            elif retainer_attachment_option == 'no_tape':
                cost_tape = 0
            else:
                cost_tape = 0
        else:
            cost_tape = 0
        cost_head_retainers = ((
                                           total_horizontal / 2) * mat_head_retainers.cost) / yield_cat17 if mat_head_retainers else 0

        total_material_cost = (cost_glass + cost_extrusions + cost_retainer + cost_glazing + cost_gaskets +
                               cost_corner_keys + cost_dual_lock + cost_foam_baffle_top + cost_foam_baffle_bottom +
                               cost_glass_protection + cost_tape + cost_head_retainers)
        cp['material_total_cost'] = total_material_cost

        # Build the itemized list with extra columns: $ per SF and % Total Cost
        materials_list = [
            {
                "Category": "Glass (Cat 15)",
                "Selected Material": mat_glass.nickname if mat_glass else "N/A",
                "Unit Cost": mat_glass.cost if mat_glass else 0,
                "Calculation": f"Total Area {total_area:.2f} × Cost / {yield_cat15}",
                "Cost ($)": cost_glass
            },
            {
                "Category": "Extrusions (Cat 1)",
                "Selected Material": mat_extrusions.nickname if mat_extrusions else "N/A",
                "Unit Cost": mat_extrusions.cost if mat_extrusions else 0,
                "Calculation": f"Total Perimeter {total_perimeter:.2f} × Cost / {yield_extrusions}",
                "Cost ($)": cost_extrusions
            },
            {
                "Category": "Retainer (Cat 1)",
                "Selected Material": (
                    mat_retainer.nickname if mat_retainer else "N/A") if retainer_option != "no_retainer" else "N/A",
                "Unit Cost": (mat_retainer.cost if mat_retainer else 0) if retainer_option != "no_retainer" else 0,
                "Calculation": (
                    f"Head Retainer: 0.5 × Total Horizontal {total_horizontal:.2f} × Cost / {yield_retainer}"
                    if retainer_option == "head_retainer"
                    else (f"Head + Sill Retainer: Total Horizontal {total_horizontal:.2f} × Cost × {yield_retainer}"
                          if retainer_option == "head_and_sill"
                          else "No Retainer")),
                "Cost ($)": cost_retainer
            },
            {
                "Category": "Glazing Spline (Cat 2)",
                "Selected Material": mat_glazing.nickname if mat_glazing else "N/A",
                "Unit Cost": mat_glazing.cost if mat_glazing else 0,
                "Calculation": f"Total Perimeter {total_perimeter:.2f} × Cost / {yield_cat2}",
                "Cost ($)": cost_glazing
            },
            {
                "Category": "Gaskets (Cat 3)",
                "Selected Material": mat_gaskets.nickname if mat_gaskets else "N/A",
                "Unit Cost": mat_gaskets.cost if mat_gaskets else 0,
                "Calculation": f"Total Vertical {total_vertical:.2f} × Cost / {yield_cat3}",
                "Cost ($)": cost_gaskets
            },
            {
                "Category": "Corner Keys (Cat 4)",
                "Selected Material": mat_corner_keys.nickname if mat_corner_keys else "N/A",
                "Unit Cost": mat_corner_keys.cost if mat_corner_keys else 0,
                "Calculation": f"Total Quantity {total_quantity:.2f} × 4 × Cost / {yield_cat4}",
                "Cost ($)": cost_corner_keys
            },
            {
                "Category": "Dual Lock (Cat 5)",
                "Selected Material": mat_dual_lock.nickname if mat_dual_lock else "N/A",
                "Unit Cost": mat_dual_lock.cost if mat_dual_lock else 0,
                "Calculation": f"Total Quantity {total_quantity:.2f} × Cost / {yield_cat5}",
                "Cost ($)": cost_dual_lock
            },
            {
                "Category": "Foam Baffle Top/Head (Cat 6)",
                "Selected Material": mat_foam_baffle_top.nickname if mat_foam_baffle_top else "N/A",
                "Unit Cost": mat_foam_baffle_top.cost if mat_foam_baffle_top else 0,
                "Calculation": f"0.5 × Total Horizontal {total_horizontal:.2f} × Cost / {yield_cat6}",
                "Cost ($)": cost_foam_baffle_top
            },
            {
                "Category": "Foam Baffle Bottom/Sill (Cat 6)",
                "Selected Material": mat_foam_baffle_bottom.nickname if mat_foam_baffle_bottom else "N/A",
                "Unit Cost": mat_foam_baffle_bottom.cost if mat_foam_baffle_bottom else 0,
                "Calculation": f"0.5 × Total Horizontal {total_horizontal:.2f} × Cost / {yield_cat6}",
                "Cost ($)": cost_foam_baffle_bottom
            },
            {
                "Category": "Glass Protection (Cat 7)",
                "Selected Material": mat_glass_protection.nickname if mat_glass_protection else "N/A",
                "Unit Cost": mat_glass_protection.cost if mat_glass_protection else 0,
                "Calculation": (f"Total Area {total_area:.2f} × Cost / {yield_cat7}"
                                if glass_protection_side == "one"
                                else (f"Total Area {total_area:.2f} × Cost × 2 / {yield_cat7}"
                                      if glass_protection_side == "double"
                                      else "No Film")),
                "Cost ($)": cost_glass_protection
            },
            {
                "Category": "Tape (Cat 10)",
                "Selected Material": mat_tape.nickname if mat_tape else "N/A",
                "Unit Cost": mat_tape.cost if mat_tape else 0,
                "Calculation": f"Retainer Attachment Option: " + (
                    "(Head Retainer - Half Horizontal)" if retainer_attachment_option == "head_retainer"
                    else ("(Head+Sill - Full Horizontal)" if retainer_attachment_option == "head_sill"
                          else "No Tape")),
                "Cost ($)": cost_tape
            },
            {
                "Category": "Head Retainers (Cat 17)",
                "Selected Material": mat_head_retainers.nickname if mat_head_retainers else "N/A",
                "Unit Cost": mat_head_retainers.cost if mat_head_retainers else 0,
                "Calculation": f"0.5 × Total Horizontal × Cost / {yield_cat17}",
                "Cost ($)": cost_head_retainers
            }
        ]
        # Add extra columns: $ per SF and % Total Cost for each item
        for item in materials_list:
            cost = item["Cost ($)"]
            item["$ per SF"] = cost / total_area if total_area > 0 else 0
            item["% Total Cost"] = (cost / cp.get("material_total_cost", 1) * 100) if cp.get("material_total_cost",
                                                                                             0) > 0 else 0

        cp["itemized_costs"] = materials_list
        save_current_project(cp)

        table_html = "<table class='summary-table'><tr><th>Category</th><th>Selected Material</th><th>Unit Cost</th><th>Calculation</th><th>Cost ($)</th><th>$ per SF</th><th>% Total Cost</th></tr>"
        for item in materials_list:
            table_html += f"<tr><td>{item['Category']}</td><td>{item['Selected Material']}</td><td>{item['Unit Cost']:.2f}</td><td>{item['Calculation']}</td><td>{item['Cost ($)']:.2f}</td><td>{item['$ per SF']:.2f}</td><td>{item['% Total Cost']:.2f}</td></tr>"
        table_html += "</table>"

        cp["materials_breakdown"] = table_html
        save_current_project(cp)

        return f"""
         <html>
           <head>
             <title>SWR Material Cost Summary</title>
             <style>{common_css}</style>
           </head>
           <body>
             <div class="container">
               <h2>SWR Material Cost Summary</h2>
               {table_html}
               <button onclick="window.location.href='/other_costs'">Next: Enter Other Costs</button>
               <button onclick="window.location.href='/'">Start New Project</button>
             </div>
           </body>
         </html>
        """

    def generate_options(materials_list):
        options = ""
        for m in materials_list:
            options += f'<option value="{m.id}">{m.nickname} - ${m.cost:.2f}</option>'
        return options

    options_glass = generate_options(materials_glass)
    options_extrusions = generate_options(materials_extrusions)
    options_retainer = generate_options(materials_extrusions)  # Using same list for retainer choices
    options_glazing = generate_options(materials_glazing)
    options_gaskets = generate_options(materials_gaskets)
    options_corner_keys = generate_options(materials_corner_keys)
    options_dual_lock = generate_options(materials_dual_lock)
    options_foam_baffle = generate_options(materials_foam_baffle)
    options_glass_protection = generate_options(materials_glass_protection)
    options_tape = generate_options(materials_tape)
    options_head_retainers = generate_options(materials_head_retainers)

    return f"""
    <html>
      <head>
         <title>SWR Materials</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Select SWR Materials</h2>
            <form method="POST">
               <h3>Glass (Category 15)</h3>
               <label for="yield_cat15">Yield for Glass:</label>
               <input type="number" step="0.01" id="yield_cat15" name="yield_cat15" value="0.97" required>
               <label for="material_glass">Glass - Total Area (sq ft) × Cost / Yield:</label>
               <select name="material_glass" id="material_glass" required>
                  {options_glass}
               </select>

               <h3>Extrusions (Category 1)</h3>
               <label for="yield_extrusions">Yield for Extrusions:</label>
               <input type="number" step="0.01" id="yield_extrusions" name="yield_extrusions" value="0.75" required>
               <label for="material_extrusions">Extrusions - Total Perimeter (ft) × Cost / Yield:</label>
               <select name="material_extrusions" id="material_extrusions" required>
                  {options_extrusions}
               </select>

               <h3>Retainer (Category 1)</h3>
               <label for="yield_retainer">Yield for Retainer:</label>
               <input type="number" step="0.01" id="yield_retainer" name="yield_retainer" value="0.75" required>
               <label for="retainer_option">Retainer Option:</label>
               <select name="retainer_option" id="retainer_option" required>
                  <option value="head_retainer">Head Retainer</option>
                  <option value="head_and_sill">Head + Sill Retainer</option>
                  <option value="no_retainer">No Retainer</option>
               </select>
               <label for="material_retainer">Retainer - (Cost calculation based on option):</label>
               <select name="material_retainer" id="material_retainer" required>
                  {options_retainer}
               </select>

               <h3>Glazing Spline (Category 2)</h3>
               <label for="yield_cat2">Yield for Glazing Spline:</label>
               <input type="number" step="0.01" id="yield_cat2" name="yield_cat2" value="0.91" required>
               <label for="material_glazing">Glazing Spline - Total Perimeter (ft) × Cost / Yield:</label>
               <select name="material_glazing" id="material_glazing" required>
                  {options_glazing}
               </select>

               <h3>Gaskets (Category 3)</h3>
               <label for="yield_cat3">Yield for Gaskets:</label>
               <input type="number" step="0.01" id="yield_cat3" name="yield_cat3" value="0.91" required>
               <label for="material_gaskets">Gaskets - Total Vertical (ft) × Cost / Yield:</label>
               <select name="material_gaskets" id="material_gaskets" required>
                  {options_gaskets}
               </select>

               <h3>Corner Keys (Category 4)</h3>
               <label for="yield_cat4">Yield for Corner Keys:</label>
               <input type="number" step="0.01" id="yield_cat4" name="yield_cat4" value="0.91" required>
               <label for="material_corner_keys">Corner Keys - Total Quantity × 4 × Cost / Yield:</label>
               <select name="material_corner_keys" id="material_corner_keys" required>
                  {options_corner_keys}
               </select>

               <h3>Dual Lock (Category 5)</h3>
               <label for="yield_cat5">Yield for Dual Lock:</label>
               <input type="number" step="0.01" id="yield_cat5" name="yield_cat5" value="0.91" required>
               <label for="material_dual_lock">Dual Lock - Total Quantity × Cost / Yield:</label>
               <select name="material_dual_lock" id="material_dual_lock" required>
                  {options_dual_lock}
               </select>

               <h3>Foam Baffle Top/Head (Category 6)</h3>
               <label for="yield_cat6">Yield for Foam Baffle:</label>
               <input type="number" step="0.01" id="yield_cat6" name="yield_cat6" value="0.91" required>
               <label for="material_foam_baffle">Foam Baffle Top/Head - Select Material:</label>
               <select name="material_foam_baffle" id="material_foam_baffle" required>
                  {options_foam_baffle}
               </select>

               <h3>Foam Baffle Bottom/Sill (Category 6)</h3>
               <label for="material_foam_baffle_bottom">Foam Baffle Bottom/Sill - Select Material:</label>
               <select name="material_foam_baffle_bottom" id="material_foam_baffle_bottom" required>
                  {options_foam_baffle}
               </select>

               <h3>Glass Protection (Category 7)</h3>
               <label for="yield_cat7">Yield for Glass Protection:</label>
               <input type="number" step="0.01" id="yield_cat7" name="yield_cat7" value="0.91" required>
               <label for="glass_protection_side">Glass Protection Side:</label>
               <select name="glass_protection_side" id="glass_protection_side" required>
                  <option value="one">One Sided</option>
                  <option value="double">Double Sided</option>
                  <option value="none">No Film</option>
               </select>
               <label for="material_glass_protection">Glass Protection - Total Area (sq ft) × Cost / Yield (or ×2 if double sided):</label>
               <select name="material_glass_protection" id="material_glass_protection" required>
                  {options_glass_protection}
               </select>

               <h3>Tape (Category 10)</h3>
               <label for="yield_cat10">Yield for Tape:</label>
               <input type="number" step="0.01" id="yield_cat10" name="yield_cat10" value="0.91" required>
               <label for="retainer_attachment_option">Retainer Attachment Option:</label>
               <select name="retainer_attachment_option" id="retainer_attachment_option" required>
                  <option value="head_retainer">Head Retainer (Half Horizontal)</option>
                  <option value="head_sill">Head+Sill (Full Horizontal)</option>
                  <option value="no_tape">No Tape</option>
               </select>
               <label for="material_tape">Tape Material - Cost calculated / Yield:</label>
               <select name="material_tape" id="material_tape" required>
                  {options_tape}
               </select>

               <h3>Head Retainers (Category 17)</h3>
               <label for="yield_cat17">Yield for Head Retainers:</label>
               <input type="number" step="0.01" id="yield_cat17" name="yield_cat17" value="0.91" required>
               <label for="material_head_retainers">Head Retainers - 0.5 × Total Horizontal (ft) × Cost / Yield:</label>
               <select name="material_head_retainers" id="material_head_retainers" required>
                  {options_head_retainers}
               </select>

               <button type="submit">Calculate SWR Material Costs</button>
            </form>
            <button onclick="window.location.href='/summary'">Back to Summary</button>
         </div>
      </body>
    </html>
    """
    return form_html


# =========================
# OTHER COSTS PAGE (Unchanged)
# =========================
@app.route('/other_costs', methods=['GET', 'POST'])
def other_costs():
    cp = get_current_project()
    material_cost = cp.get('material_total_cost', 0)
    total_quantity = cp.get('swr_total_quantity', 0)
    if request.method == 'POST':
        num_trucks = float(request.form.get('num_trucks', 0))
        cost_per_truck = float(request.form.get('truck_cost', 0))
        total_truck_cost = num_trucks * cost_per_truck

        hourly_rate = float(request.form.get('hourly_rate', 0))
        hours_per_panel = float(request.form.get('hours_per_panel', 0))
        installation_cost = hourly_rate * hours_per_panel * total_quantity

        cost_scissor = float(request.form.get('cost_scissor', 0))
        cost_lull = float(request.form.get('cost_lull', 0))
        cost_baker = float(request.form.get('cost_baker', 0))
        cost_crane = float(request.form.get('cost_crane', 0))
        cost_blankets = float(request.form.get('cost_blankets', 0))
        equipment_cost = cost_scissor + cost_lull + cost_baker + cost_crane + cost_blankets

        airfare = float(request.form.get('airfare', 0))
        lodging = float(request.form.get('lodging', 0))
        meals = float(request.form.get('meals', 0))
        car_rental = float(request.form.get('car_rental', 0))
        travel_cost = airfare + lodging + meals + car_rental

        cost_audit = float(request.form.get('cost_audit', 0))
        cost_design = float(request.form.get('cost_design', 0))
        cost_thermal_stress = float(request.form.get('cost_thermal_stress', 0))
        cost_thermal_performance = float(request.form.get('cost_thermal_performance', 0))
        cost_mockup = float(request.form.get('cost_mockup', 0))
        cost_window = float(request.form.get('cost_window', 0))
        cost_energy = float(request.form.get('cost_energy', 0))
        cost_cost_benefit = float(request.form.get('cost_cost_benefit', 0))
        cost_utility = float(request.form.get('cost_utility', 0))
        sales_cost = (cost_audit + cost_design + cost_thermal_stress + cost_thermal_performance +
                      cost_mockup + cost_window + cost_energy + cost_cost_benefit + cost_utility)

        total_truck_cost = num_trucks * cost_per_truck
        additional_total = total_truck_cost + installation_cost + equipment_cost + travel_cost + sales_cost
        grand_total = material_cost + additional_total

        cp["total_truck_cost"] = total_truck_cost
        cp["installation_cost"] = installation_cost
        cp["equipment_cost"] = equipment_cost
        cp["travel_cost"] = travel_cost
        cp["sales_cost"] = sales_cost
        cp["additional_total"] = additional_total
        cp["grand_total"] = grand_total
        save_current_project(cp)

        result_html = f"""
         <html>
           <head>
             <title>Other Costs Summary</title>
             <style>{common_css}</style>
           </head>
           <body>
             <div class="container">
               <h2>Other Costs Summary</h2>
               <table class="summary-table">
                 <tr><th>Cost Item</th><th>Amount ($)</th></tr>
                 <tr><td>Material Cost</td><td>{material_cost:.2f}</td></tr>
                 <tr><td>Truck Cost</td><td>{total_truck_cost:.2f}</td></tr>
                 <tr><td>Installation</td><td>{installation_cost:.2f}</td></tr>
                 <tr><td>Equipment</td><td>{equipment_cost:.2f}</td></tr>
                 <tr><td>Travel</td><td>{travel_cost:.2f}</td></tr>
                 <tr><td>Sales</td><td>{sales_cost:.2f}</td></tr>
                 <tr><td><strong>Additional Total</strong></td><td><strong>{additional_total:.2f}</strong></td></tr>
                 <tr><th>Grand Total</th><th>{grand_total:.2f}</th></tr>
               </table>
               <button onclick="window.location.href='/margins'">Next: Set Margins</button>
               <button onclick="window.location.href='/materials'">Back to SWR Materials</button>
             </div>
           </body>
         </html>
        """
        return result_html

    return f"""
    <html>
      <head>
         <title>Other Costs</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Enter Additional Costs</h2>
            <form method="POST">
              <h3>Logistics (Truck Costs)</h3>
              <label for="num_trucks">Number of Trucks:</label>
              <input type="number" step="1" id="num_trucks" name="num_trucks" value="0" required>
              <label for="truck_cost">Cost per Truck ($):</label>
              <input type="number" step="100" id="truck_cost" name="truck_cost" value="0" required>

              <h3>Installation</h3>
              <label for="hourly_rate">Hourly Rate ($/hr):</label>
              <input type="number" step="1" id="hourly_rate" name="hourly_rate" value="0" required>
              <label for="hours_per_panel">Hours per Panel:</label>
              <input type="number" step="1" id="hours_per_panel" name="hours_per_panel" value="0" required>

              <h3>Equipment</h3>
              <p>Enter cost for each equipment (0 if not needed):</p>
              <label for="cost_scissor">Scissor Lift ($):</label>
              <input type="number" step="0.01" id="cost_scissor" name="cost_scissor" value="0">
              <label for="cost_lull">Lull Rental ($):</label>
              <input type="number" step="0.01" id="cost_lull" name="cost_lull" value="0">
              <label for="cost_baker">Baker Rolling Staging ($):</label>
              <input type="number" step="0.01" id="cost_baker" name="cost_baker" value="0">
              <label for="cost_crane">Crane ($):</label>
              <input type="number" step="0.01" id="cost_crane" name="cost_crane" value="0">
              <label for="cost_blankets">Finished Protected Board Blankets ($):</label>
              <input type="number" step="0.01" id="cost_blankets" name="cost_blankets" value="0">

              <h3>Travel</h3>
              <label for="airfare">Airfare ($):</label>
              <input type="number" step="0.01" id="airfare" name="airfare" value="0" required>
              <label for="lodging">Lodging ($):</label>
              <input type="number" step="0.01" id="lodging" name="lodging" value="0" required>
              <label for="meals">Meals &amp; Incidentals ($):</label>
              <input type="number" step="0.01" id="meals" name="meals" value="0" required>
              <label for="car_rental">Car Rental + Gas ($):</label>
              <input type="number" step="0.01" id="car_rental" name="car_rental" value="0" required>

              <h3>Sales</h3>
              <p>Enter cost for each sales item (0 if not needed):</p>
              <label for="cost_audit">Building Audit/Survey ($):</label>
              <input type="number" step="100" id="cost_audit" name="cost_audit" value="0">
              <label for="cost_design">System Design Customization ($):</label>
              <input type="number" step="100" id="cost_design" name="cost_design" value="0">
              <label for="cost_thermal_stress">Thermal Stress Analysis ($):</label>
              <input type="number" step="100" id="cost_thermal_stress" name="cost_thermal_stress" value="0">
              <label for="cost_thermal_performance">Thermal Performance Simulation ($):</label>
              <input type="number" step="100" id="cost_thermal_performance" name="cost_thermal_performance" value="0">
              <label for="cost_mockup">Visual &amp; Performance Mockup ($):</label>
              <input type="number" step="100" id="cost_mockup" name="cost_mockup" value="0">
              <label for="cost_window">Window Performance M&amp;V ($):</label>
              <input type="number" step="100" id="cost_window" name="cost_window" value="0">
              <label for="cost_energy">Building Energy Model ($):</label>
              <input type="number" step="100" id="cost_energy" name="cost_energy" value="0">
              <label for="cost_cost_benefit">Cost-Benefit Analysis ($):</label>
              <input type="number" step="100" id="cost_cost_benefit" name="cost_cost_benefit" value="0">
              <label for="cost_utility">Utility Incentive Application ($):</label>
              <input type="number" step="100" id="cost_utility" name="cost_utility" value="0">

              <button type="submit">Calculate Additional Costs</button>
            </form>
            <button onclick="window.location.href='/materials'">Back to SWR Materials</button>
         </div>
      </body>
    </html>
    """
    return form_html


# =========================
# MARGINS PAGE (with Dynamic $/SF Calculation and Navigation Buttons)
# =========================
@app.route('/margins', methods=['GET', 'POST'])
def margins():
    cp = get_current_project()
    base_costs = {
        "Panel Total": cp.get("material_total_cost", 0),
        "Logistics": cp.get("total_truck_cost", 0),
        "Installation": cp.get("installation_cost", 0),
        "Equipment": cp.get("equipment_cost", 0),
        "Travel": cp.get("travel_cost", 0),
        "Sales": cp.get("sales_cost", 0)
    }
    total_area = cp.get("swr_total_area", 0) + cp.get("igr_total_area", 0)
    if request.method == 'POST':
        margins = {}
        for category in base_costs:
            try:
                margins[category] = float(request.form.get(f"{category}_margin", 0))
            except:
                margins[category] = 0
        adjusted_costs = {cat: base_costs[cat] * (1 + margins[cat] / 100) for cat in base_costs}
        final_total = sum(adjusted_costs.values())

        summary_data = {
            "Category": list(base_costs.keys()) + ["Grand Total"],
            "Original Cost ($)": [base_costs[cat] for cat in base_costs] + [sum(base_costs.values())],
            "Margin (%)": [margins[cat] for cat in base_costs] + [""],
            "Cost with Margin ($)": [adjusted_costs[cat] for cat in base_costs] + [final_total]
        }
        df_summary = pd.DataFrame(summary_data)
        csv_buffer = io.StringIO()
        df_summary.to_csv(csv_buffer, index=False)
        csv_output = csv_buffer.getvalue()

        cp["final_summary"] = []
        for cat in base_costs:
            cp["final_summary"].append({
                "Category": cat,
                "Original Cost ($)": base_costs[cat],
                "Margin (%)": margins.get(cat, 0),
                "Cost with Margin ($)": adjusted_costs[cat]
            })
        cp["grand_total"] = final_total
        save_current_project(cp)

        result_html = f"""
         <html>
           <head>
             <title>Final Cost with Margins</title>
             <style>{common_css}</style>
           </head>
           <body>
             <div class="container">
               <h2>Final Cost with Margins</h2>
               <table class="summary-table">
                 <tr>
                   <th>Category</th>
                   <th>Original Cost ($)</th>
                   <th>Margin (%)</th>
                   <th>Cost with Margin ($)</th>
                 </tr>
        """
        for cat in base_costs:
            result_html += f"""
                 <tr>
                   <td>{cat}</td>
                   <td>{base_costs[cat]:.2f}</td>
                   <td>{margins[cat]:.2f}</td>
                   <td>{adjusted_costs[cat]:.2f}</td>
                 </tr>
            """
        result_html += f"""
                 <tr>
                   <th colspan="3">Grand Total</th>
                   <th>{final_total:.2f}</th>
                 </tr>
               </table>
               <form method="POST" action="/download_final_summary">
                 <input type="hidden" name="csv_data" value='{csv_output}'>
                 <button type="submit">Download Final Summary CSV</button>
               </form>
               <button style="font-size:0.8em; margin-top:5px;" onclick="window.location.href='/other_costs'">Back to Other Costs</button>
               <button style="font-size:0.8em; margin-top:5px;" onclick="window.location.href='/materials'">Back to SWR Materials</button>
               <button style="font-size:0.8em; margin-top:5px;" onclick="window.location.href='/'">Start New Project</button>
             </div>
             <div class="container">
               <h3>Dynamic Cost per SF</h3>
               <p>Total Cost with Margins divided by Total SF:</p>
               <p>Total SF: {total_area:.2f}</p>
               <p>$ per SF: $<span id="cost_per_sf">0.00</span></p>
             </div>
             <script>
                function updateOutput(sliderId, outputId) {{
                    var slider = document.getElementById(sliderId);
                    var output = document.getElementById(outputId);
                    output.value = slider.value;
                    recalcCostPerSF();
                }}
                function recalcCostPerSF() {{
                    var baseCosts = {json.dumps(base_costs)};
                    var totalArea = {total_area};
                    var adjustedTotal = 0;
                    for (var cat in baseCosts) {{
                        var slider = document.getElementById(cat.replace(/ /g, "_") + "_margin");
                        var margin = parseFloat(slider.value) || 0;
                        adjustedTotal += baseCosts[cat] * (1 + margin / 100);
                    }}
                    var costPerSF = adjustedTotal / totalArea;
                    document.getElementById("cost_per_sf").innerText = costPerSF.toFixed(2);
                }}
                window.addEventListener("load", function() {{
                    var sliders = document.querySelectorAll("input[type='range']");
                    sliders.forEach(function(slider) {{
                        slider.addEventListener("input", recalcCostPerSF);
                    }});
                    recalcCostPerSF();
                }});
             </script>
           </body>
         </html>
        """
        return result_html

    form_html = f"""
    <html>
      <head>
         <title>Set Margins</title>
         <style>{common_css}</style>
         <script>
            function updateOutput(sliderId, outputId) {{
                var slider = document.getElementById(sliderId);
                var output = document.getElementById(outputId);
                output.value = slider.value;
                recalcCostPerSF();
            }}
            function recalcCostPerSF() {{
                var baseCosts = {json.dumps(base_costs)};
                var totalArea = {total_area};
                var adjustedTotal = 0;
                for (var cat in baseCosts) {{
                    var slider = document.getElementById(cat.replace(/ /g, "_") + "_margin");
                    var margin = parseFloat(slider.value) || 0;
                    adjustedTotal += baseCosts[cat] * (1 + margin / 100);
                }}
                var costPerSF = adjustedTotal / totalArea;
                document.getElementById("cost_per_sf").innerText = costPerSF.toFixed(2);
            }}
            window.addEventListener("load", function() {{
                var sliders = document.querySelectorAll("input[type='range']");
                sliders.forEach(function(slider) {{
                    slider.addEventListener("input", recalcCostPerSF);
                }});
                recalcCostPerSF();
            }});
         </script>
      </head>
      <body>
         <div class="container">
            <h2>Set Margins for Each Cost Category</h2>
            <form method="POST">
    """
    for category in base_costs:
        var_id = category.replace(" ", "_")
        form_html += f"""
                <label for="{var_id}_margin">{category} Margin (%):</label>
                <input type="range" id="{var_id}_margin" name="{category}_margin" min="0" max="100" step="1" value="0" oninput="updateOutput('{var_id}_margin', '{var_id}_output')">
                <output id="{var_id}_output">0</output><br>
        """
    form_html += """
                <button type="submit">Calculate Final Cost with Margins</button>
            </form>
            <p>Total Cost per SF: $<span id="cost_per_sf">0.00</span></p>
            <button onclick="window.location.href='/other_costs'">Back to Other Costs</button>
         </div>
      </body>
    </html>
    """
    return form_html


# =========================
# DOWNLOAD FINAL SUMMARY CSV
# =========================
@app.route('/download_final_summary', methods=['POST'])
def download_final_summary():
    csv_data = create_final_summary_csv()
    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-disposition": "attachment; filename=final_cost_summary.csv"}
    )


def create_final_summary_csv():
    cp = get_current_project()
    output = io.StringIO()
    writer = csv.writer(output)

    project_name = cp.get("project_name", "Unnamed Project")
    project_number = cp.get("project_number", "")
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    writer.writerow(["Project Name:", project_name])
    writer.writerow(["Project Number:", project_number])
    writer.writerow(["Date:", current_date])
    writer.writerow([])

    swr_panels = cp.get("swr_total_quantity", 0)
    swr_area = cp.get("swr_total_area", 0)
    swr_perimeter = cp.get("swr_total_perimeter", 0)
    swr_horizontal = cp.get("swr_total_horizontal_ft", 0)
    swr_vertical = cp.get("swr_total_vertical_ft", 0)
    igr_panels = cp.get("igr_total_quantity", 0)
    igr_area = cp.get("igr_total_area", 0)
    igr_perimeter = cp.get("igr_total_perimeter", 0)
    igr_horizontal = cp.get("igr_total_horizontal_ft", 0)
    igr_vertical = cp.get("igr_total_vertical_ft", 0)
    combined_panels = swr_panels + igr_panels
    combined_area = swr_area + igr_area
    combined_perimeter = swr_perimeter + igr_perimeter
    combined_horizontal = swr_horizontal + igr_horizontal
    combined_vertical = swr_vertical + igr_vertical

    writer.writerow(["Combined Totals"])
    writer.writerow(["Panels", "Area (sq ft)", "Perimeter (ft)", "Horizontal (ft)", "Vertical (ft)"])
    writer.writerow([combined_panels, combined_area, combined_perimeter, combined_horizontal, combined_vertical])
    writer.writerow([])

    writer.writerow(["Project Summary"])
    writer.writerow(["Category", "Original Cost ($)", "Margin (%)", "Cost with Margin ($)"])
    final_summary = cp.get("final_summary", [])
    for row in final_summary:
        writer.writerow([
            row.get("Category", ""),
            row.get("Original Cost ($)", 0),
            row.get("Margin (%)", 0),
            row.get("Cost with Margin ($)", 0)
        ])
    writer.writerow([])

    writer.writerow(["Detailed Itemized Costs"])
    writer.writerow(
        ["Category", "Selected Material", "Unit Cost", "Calculation", "Cost ($)", "$ per SF", "% Total Cost"])
    line_items = cp.get("itemized_costs", [])
    for item in line_items:
        category = item.get("Category", "")
        material = item.get("Selected Material", "")
        unit_cost = item.get("Unit Cost", 0)
        calc_text = item.get("Calculation", "")
        cost = item.get("Cost ($)", 0)
        cost_per_sf = cost / cp.get("swr_total_area", 1) if cp.get("swr_total_area", 0) > 0 else 0
        percent_total = (cost / cp.get("grand_total", 1) * 100) if cp.get("grand_total", 0) > 0 else 0
        writer.writerow([category, material, unit_cost, calc_text, cost, cost_per_sf, percent_total])

    return output.getvalue()


# =========================
# EXCEL EXPORT (Detailed Final Export)
# =========================
def create_final_export_excel(margins_dict=None):
    cp = get_current_project()
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book
    ws = workbook.add_worksheet("Project Export")

    ws.write("A1", "Project Name:")
    ws.write("B1", cp.get("project_name", "Unnamed Project"))
    ws.write("A2", "Project Number:")
    ws.write("B2", cp.get("project_number", ""))
    ws.write("A3", "Date:")
    ws.write("B3", datetime.datetime.now().strftime("%Y-%m-%d"))

    ws.write("D1", "Combined Totals")
    ws.write("D2", "Panels")
    ws.write("E2", cp.get("swr_total_quantity", 0) + cp.get("igr_total_quantity", 0))
    ws.write("D3", "Area (sq ft)")
    ws.write("E3", cp.get("swr_total_area", 0) + cp.get("igr_total_area", 0))
    ws.write("D4", "Perimeter (ft)")
    ws.write("E4", cp.get("swr_total_perimeter", 0) + cp.get("igr_total_perimeter", 0))
    ws.write("D5", "Horizontal (ft)")
    ws.write("E5", cp.get("swr_total_horizontal_ft", 0) + cp.get("igr_total_horizontal_ft", 0))
    ws.write("D6", "Vertical (ft)")
    ws.write("E6", cp.get("swr_total_vertical_ft", 0) + cp.get("igr_total_vertical_ft", 0))

    ws.write("G1", "Project Summary")
    ws.write("G2", "Category")
    ws.write("H2", "Original Cost ($)")
    ws.write("I2", "Margin (%)")
    ws.write("J2", "Cost with Margin ($)")
    row = 3
    for item in cp.get("final_summary", []):
        ws.write(row, 6, item.get("Category", ""))
        ws.write(row, 7, item.get("Original Cost ($)", 0))
        ws.write(row, 8, item.get("Margin (%)", 0))
        ws.write(row, 9, item.get("Cost with Margin ($)", 0))
        row += 1

    start_row = 9
    headers_detail = ["Category", "Selected Material", "Unit Cost", "Calculation", "Cost ($)", "$ per SF",
                      "% Total Cost"]
    for col, header in enumerate(headers_detail):
        ws.write(start_row, col, header)
    line_items = cp.get("itemized_costs", [])
    for i, item in enumerate(line_items, start=start_row + 1):
        category = item.get("Category", "")
        material = item.get("Selected Material", "")
        unit_cost = item.get("Unit Cost", 0)
        calc_text = item.get("Calculation", "")
        cost = item.get("Cost ($)", 0)
        cost_per_sf = cost / cp.get("swr_total_area", 1) if cp.get("swr_total_area", 0) > 0 else 0
        percent_total = (cost / cp.get("grand_total", 1) * 100) if cp.get("grand_total", 0) > 0 else 0
        ws.write(i, 0, category)
        ws.write(i, 1, material)
        ws.write(i, 2, unit_cost)
        ws.write(i, 3, calc_text)
        ws.write(i, 4, cost)
        ws.write(i, 5, cost_per_sf)
        ws.write(i, 6, percent_total)

    writer.close()
    output.seek(0)
    return output


# =========================
# DOWNLOAD FINAL EXPORT (Excel) ROUTE
# =========================
@app.route('/download_final_export')
def download_final_export():
    excel_file = create_final_export_excel()
    return send_file(excel_file, attachment_filename="Project_Cost_Summary.xlsx", as_attachment=True)


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)