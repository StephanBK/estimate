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

# Database Configuration (update with your actual connection string)
app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://u7vukdvn20pe3c:p918802c410825b956ccf24c5af8d168b4d9d69e1940182bae9bd8647eb606845@cb5ajfjosdpmil.cluster-czrs8kj4isg7.us-east-1.rds.amazonaws.com:5432/dcobttk99a5sie'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Initialize the database
db = SQLAlchemy(app)


# Define the Material model (using 'nickname' for display)
class Material(db.Model):
    __tablename__ = 'materials'
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.Integer, nullable=False)
    nickname = db.Column(db.String(100), nullable=False)
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
    # New columns:
    quantity = db.Column(db.Float)
    min_lead = db.Column(db.Integer)
    max_lead = db.Column(db.Integer)


# Helper function: recursively convert NumPy types to native Python types
def make_serializable(obj):
    if isinstance(obj, dict):
        return {k: make_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [make_serializable(x) for x in obj]
    elif isinstance(obj, np.int64):
        return int(obj)
    else:
        return obj


# Session management helpers
def get_current_project():
    cp = session.get("current_project")
    if cp is None:
        cp = {}
        session["current_project"] = cp
    return cp


def save_current_project(cp):
    session["current_project"] = make_serializable(cp)


# Helper to generate select options (with safe yield_cost display)
def generate_options(materials_list, selected_value=None):
    options = ""
    for m in materials_list:
        yield_cost = m.yield_cost if m.yield_cost is not None else 0.0
        if str(m.id) == str(selected_value):
            options += f'<option value="{m.id}" selected>{m.nickname} - ${yield_cost:.2f}</option>'
        else:
            options += f'<option value="{m.id}">{m.nickname} - ${yield_cost:.2f}</option>'
    return options


# Path to the CSV template
TEMPLATE_PATH = 'estimation_template_template.csv'

# Common CSS for consistent styling
common_css = """
    @import url('https://fonts.googleapis.com/css2?family=Exo+2:wght@400;600&display=swap');
    body { background-color: #121212; color: #ffffff; font-family: 'Exo 2', sans-serif; padding: 20px; }
    .container { background-color: #1e1e1e; padding: 20px; border-radius: 10px; max-width: 900px; margin: auto; box-shadow: 0 0 15px rgba(0,128,128,0.5); }
    input, select, button, textarea { padding: 10px; margin: 5px 0; border: none; border-radius: 5px; background-color: #333; color: #fff; width: 100%; }
    button { background-color: #008080; cursor: pointer; transition: background-color 0.3s ease; }
    button:hover { background-color: #00a0a0; }
    a.btn {
        display: inline-block;
        background-color: #008080;
        color: #fff;
        text-align: center;
        padding: 10px 20px;
        border-radius: 5px;
        text-decoration: none;
        min-width: 140px;
    }
    a.btn:hover { background-color: #00a0a0; }
    h2, p { text-align: center; }
    .summary-table, .data-table { width: 100%; border-collapse: collapse; margin-top: 20px; table-layout: fixed; }
    .summary-table th, .summary-table td, .data-table th, .data-table td { border: 1px solid #444; padding: 8px; text-align: center; word-wrap: break-word; }
    .summary-table th { background-color: #008080; color: white; }
    .data-table th { background-color: #00a0a0; color: white; }
    tr:nth-child(even) { background-color: #1e1e1f; }
    tr:nth-child(odd) { background-color: #2b2b2b; }
    .margin-row { display: flex; align-items: center; margin-bottom: 10px; }
    .margin-row label { width: 180px; }
    .btn-group { display: flex; justify-content: space-between; margin-top: 10px; }
    .btn { min-width: 140px; }
    .btn-download { font-size: 1.5em; padding: 20px 40px; display: block; margin: 20px auto; }
    /* For custom installation fields */
    #custom_installation { display: none; }
"""


# ==================================================
# INDEX PAGE
# ==================================================
@app.route('/', methods=['GET', 'POST'])
def index():
    cp = get_current_project()
    if request.method == 'POST':
        cp['customer_name'] = request.form['customer_name']
        cp['project_name'] = request.form['project_name']
        cp['estimated_by'] = request.form['estimated_by']
        cp['swr_system'] = request.form['swr_system']
        cp['swr_mount'] = request.form['swr_mount']
        cp['igr_type'] = request.form['igr_type']
        cp['igr_location'] = request.form['igr_location']
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
               <div>
                  <label for="customer_name">Customer Name:</label>
                  <input type="text" id="customer_name" name="customer_name" value="{cp.get('customer_name', '')}" required>
               </div>
               <div>
                  <label for="project_name">Project Name:</label>
                  <input type="text" id="project_name" name="project_name" value="{cp.get('project_name', '')}" required>
               </div>
               <div>
                  <label for="estimated_by">Estimated by:</label>
                  <input type="text" id="estimated_by" name="estimated_by" value="{cp.get('estimated_by', '')}" required>
               </div>
               <div>
                  <label for="swr_system">SWR System:</label>
                  <select name="swr_system" id="swr_system" style="width:200px;" required>
                      <option value="SWR" {"selected" if cp.get('swr_system', '') == "SWR" else ""}>SWR</option>
                      <option value="SWR-IG" {"selected" if cp.get('swr_system', '') == "SWR-IG" else ""}>SWR-IG</option>
                      <option value="SWR-VIG" {"selected" if cp.get('swr_system', '') == "SWR-VIG" else ""}>SWR-VIG</option>
                  </select>
               </div>
               <div>
                  <label for="swr_mount">SWR Mount Type:</label>
                  <select name="swr_mount" id="swr_mount" style="width:200px;" required>
                      <option value="Inset-mount" {"selected" if cp.get('swr_mount', '') == "Inset-mount" else ""}>Inset-mount</option>
                      <option value="Overlap-mount" {"selected" if cp.get('swr_mount', '') == "Overlap-mount" else ""}>Overlap-mount</option>
                  </select>
               </div>
               <div>
                  <label for="igr_type">IGR Type:</label>
                  <select name="igr_type" id="igr_type" style="width:200px;" required>
                     <option value="Wet Seal IGR" {"selected" if cp.get('igr_type', '') == "Wet Seal IGR" else ""}>Wet Seal IGR</option>
                     <option value="Dry Seal IGR" {"selected" if cp.get('igr_type', '') == "Dry Seal IGR" else ""}>Dry Seal IGR</option>
                  </select>
               </div>
               <div>
                  <label for="igr_location">IGR Location:</label>
                  <select name="igr_location" id="igr_location" style="width:200px;" required>
                     <option value="Interior" {"selected" if cp.get('igr_location', '') == "Interior" else ""}>Interior</option>
                     <option value="Exterior" {"selected" if cp.get('igr_location', '') == "Exterior" else ""}>Exterior</option>
                  </select>
               </div>
               <div>
                  <label for="file">Upload Filled Template (CSV):</label>
                  <input type="file" id="file" name="file" accept=".csv" required>
               </div>
               <div class="btn-group">
                 <button type="button" class="btn" onclick="window.location.href='/download-template'">Download Template CSV</button>
                 <button type="submit" class="btn">Submit</button>
               </div>
            </form>
         </div>
      </body>
    </html>
    """


# ==================================================
# DOWNLOAD TEMPLATE
# ==================================================
@app.route('/download-template')
def download_template():
    return send_file(TEMPLATE_PATH, as_attachment=True)


# ==================================================
# SUMMARY PAGE (After CSV is Read)
# ==================================================
@app.route('/summary')
def summary():
    cp = get_current_project()
    file_path = cp.get('file_path')
    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        return f"<h2 style='color: red;'>Error reading the file: {e}</h2>"
    if "Type" in df.columns:
        type_col = "Type"
    elif "Type(IGR/SWR)" in df.columns:
        type_col = "Type(IGR/SWR)"
    else:
        type_col = None

    if type_col:
        df_swr = df[df[type_col].str.upper() == "SWR"]
        df_igr = df[df[type_col].str.upper() == "IGR"]
    else:
        df_swr = df
        df_igr = pd.DataFrame()

    def compute_totals(df_subset):
        df_subset['Area (sq in)'] = df_subset['VGA Width in'] * df_subset['VGA Height in']
        df_subset['Total Area (sq ft)'] = (df_subset['Area (sq in)'] * df_subset['Qty']) / 144
        df_subset['Perimeter (in)'] = 2 * (df_subset['VGA Width in'] + df_subset['VGA Height in'])
        df_subset['Total Perimeter (ft)'] = (df_subset['Perimeter (in)'] * df_subset['Qty']) / 12
        df_subset['Total Vertical (ft)'] = (df_subset['VGA Height in'] * df_subset['Qty'] * 2) / 12
        df_subset['Total Horizontal (ft)'] = (df_subset['VGA Width in'] * df_subset['Qty'] * 2) / 144
        total_area = df_subset['Total Area (sq ft)'].sum()
        total_perimeter = df_subset['Total Perimeter (ft)'].sum()
        total_vertical = df_subset['Total Vertical (ft)'].sum()
        total_horizontal = df_subset['Total Horizontal (ft)'].sum()
        total_quantity = df_subset['Qty'].sum()
        return total_area, total_perimeter, total_vertical, total_horizontal, total_quantity

    swr_area, swr_perimeter, swr_vertical, swr_horizontal, swr_quantity = compute_totals(df_swr)
    igr_area, igr_perimeter, igr_vertical, igr_horizontal, igr_quantity = (0, 0, 0, 0, 0)
    if not df_igr.empty:
        igr_area, igr_perimeter, igr_vertical, igr_horizontal, igr_quantity = compute_totals(df_igr)
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
    next_button = '<button type="button" class="btn" onclick="window.location.href=\'/materials\'">Next: SWR Materials</button>'
    btn_html = '<button type="button" class="btn" onclick="window.location.href=\'/new_project\'">Start New Project</button>' + next_button
    summary_html = f"""
    <html>
      <head>
         <title>Project Summary</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Project Summary</h2>
            <p><strong>Customer Name:</strong> {cp.get('customer_name', 'N/A')}</p>
            <p><strong>Project Name:</strong> {cp.get('project_name')}</p>
            <p><strong>Estimated by:</strong> {cp.get('estimated_by', 'N/A')}</p>
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
            <div class="btn-group">
              {btn_html}
            </div>
         </div>
      </body>
    </html>
    """
    return summary_html


# ==================================================
# SWR MATERIALS PAGE (Material Selection & Cost Summary)
# ==================================================
@app.route('/materials', methods=['GET', 'POST'])
def materials_page():
    cp = get_current_project()
    if cp.get('swr_total_area', 0) <= 0:
        return redirect(url_for('other_costs'))
    try:
        materials_glass = Material.query.filter_by(category=15).all()
        materials_aluminum = Material.query.filter_by(category=1).all()
        materials_glazing = Material.query.filter_by(category=2).all()
        materials_gaskets = Material.query.filter_by(category=3).all()
        materials_corner_keys = Material.query.filter_by(category=4).all()
        materials_dual_lock = Material.query.filter_by(category=5).all()
        materials_foam_baffle = Material.query.filter_by(category=6).all()
        materials_glass_protection = Material.query.filter_by(category=7).all()
        materials_tape = Material.query.filter_by(category=10).all()
        materials_head_retainers = Material.query.filter_by(category=17).all()
        materials_screws = Material.query.filter_by(category=18).all()
        # Added for Setting Block (Category 16)
        materials_setting_block = Material.query.filter_by(category=16).all()
    except Exception as e:
        return f"<h2 style='color: red;'>Error fetching materials: {e}</h2>"

    if request.method == 'POST':
        # Process yield values
        try:
            cp['yield_cat15'] = float(request.form.get('yield_cat15', 0.97))
        except:
            cp['yield_cat15'] = 0.97
        try:
            cp['yield_aluminum'] = float(request.form.get('yield_aluminum', 0.75))
        except:
            cp['yield_aluminum'] = 0.75
        try:
            cp['yield_cat2'] = float(request.form.get('yield_cat2', 0.91))
        except:
            cp['yield_cat2'] = 0.91
        try:
            cp['yield_cat3'] = float(request.form.get('yield_cat3', 0.91))
        except:
            cp['yield_cat3'] = 0.91
        try:
            cp['yield_cat4'] = float(request.form.get('yield_cat4', 0.91))
        except:
            cp['yield_cat4'] = 0.91
        try:
            cp['yield_cat5'] = float(request.form.get('yield_cat5', 0.91))
        except:
            cp['yield_cat5'] = 0.91
        try:
            cp['yield_cat6'] = float(request.form.get('yield_cat6', 0.91))
        except:
            cp['yield_cat6'] = 0.91
        try:
            cp['yield_cat7'] = float(request.form.get('yield_cat7', 0.91))
        except:
            cp['yield_cat7'] = 0.91
        try:
            cp['yield_cat10'] = float(request.form.get('yield_cat10', 0.91))
        except:
            cp['yield_cat10'] = 0.91
        # --- Added for Setting Block ---
        try:
            cp['yield_cat16'] = float(request.form.get('yield_cat16', 1.0))
        except:
            cp['yield_cat16'] = 1.0

        # Process material selections
        cp["material_glass"] = request.form.get('material_glass')
        cp["material_aluminum"] = request.form.get('material_aluminum')
        cp["retainer_option"] = request.form.get('retainer_option')
        cp["material_retainer"] = request.form.get('material_retainer')
        cp["material_glazing"] = request.form.get('material_glazing')
        cp["material_gaskets"] = request.form.get('material_gaskets')
        cp["jamb_plate"] = request.form.get('jamb_plate')
        cp["jamb_plate_screws"] = request.form.get('jamb_plate_screws')
        cp["material_corner_keys"] = request.form.get('material_corner_keys')
        cp["material_dual_lock"] = request.form.get('material_dual_lock')
        cp["material_foam_baffle"] = request.form.get('material_foam_baffle')
        cp["material_foam_baffle_bottom"] = request.form.get('material_foam_baffle_bottom')
        cp["glass_protection_side"] = request.form.get('glass_protection_side')
        cp["material_glass_protection"] = request.form.get('material_glass_protection')
        cp["retainer_attachment_option"] = request.form.get('retainer_attachment_option')
        cp["material_tape"] = request.form.get('material_tape')
        cp["swr_note"] = request.form.get("swr_note", "")
        # Capture the new glass thickness value
        cp["glass_thickness"] = request.form.get("glass_thickness")
        # --- Added for Setting Block ---
        cp["material_setting_block"] = request.form.get("material_setting_block")
        save_current_project(cp)

        # Retrieve the selected material objects
        mat_glass = Material.query.get(cp.get("material_glass")) if cp.get("material_glass") else None
        mat_aluminum = Material.query.get(cp.get("material_aluminum")) if cp.get("material_aluminum") else None
        mat_retainer = Material.query.get(cp.get("material_retainer")) if cp.get("material_retainer") else None
        mat_glazing = Material.query.get(cp.get("material_glazing")) if cp.get("material_glazing") else None
        mat_gaskets = Material.query.get(cp.get("material_gaskets")) if cp.get("material_gaskets") else None
        mat_corner_keys = Material.query.get(cp.get("material_corner_keys")) if cp.get("material_corner_keys") else None
        mat_dual_lock = Material.query.get(cp.get("material_dual_lock")) if cp.get("material_dual_lock") else None
        mat_foam_baffle_top = Material.query.get(cp.get("material_foam_baffle")) if cp.get(
            "material_foam_baffle") else None
        mat_foam_baffle_bottom = Material.query.get(cp.get("material_foam_baffle_bottom")) if cp.get(
            "material_foam_baffle_bottom") else None
        mat_glass_protection = Material.query.get(cp.get("material_glass_protection")) if cp.get(
            "material_glass_protection") else None
        mat_tape = Material.query.get(cp.get("material_tape")) if cp.get("material_tape") else None
        mat_screws = Material.query.get(cp.get("material_screws")) if cp.get("material_screws") else None
        # --- Retrieve Setting Block Material ---
        mat_setting_block = Material.query.get(cp.get("material_setting_block")) if cp.get(
            "material_setting_block") else None

        # --- Compute Lead Times for SWR Materials (unchanged) ---
        swr_materials = [mat_glass, mat_aluminum, mat_retainer, mat_glazing, mat_gaskets,
                         mat_corner_keys, mat_dual_lock, mat_foam_baffle_top,
                         mat_foam_baffle_bottom, mat_glass_protection, mat_tape, mat_screws]
        swr_min_lead = 0
        swr_max_lead = 0
        for m in swr_materials:
            if m:
                swr_min_lead = max(swr_min_lead, m.min_lead if m.min_lead is not None else 0)
                swr_max_lead = max(swr_max_lead, m.max_lead if m.max_lead is not None else 0)
        cp["min_lead_material"] = swr_min_lead
        cp["max_lead_material"] = swr_max_lead
        cp["min_lead_fabrication"] = swr_min_lead
        cp["max_lead_fabrication"] = swr_max_lead
        cp["min_total_lead"] = swr_min_lead * 2
        cp["max_total_lead"] = swr_max_lead * 2

        save_current_project(cp)

        # Use the previously computed totals (from CSV summary)
        total_area = cp.get('swr_total_area', 0)
        total_perimeter = cp.get('swr_total_perimeter', 0)
        total_vertical = cp.get('swr_total_vertical_ft', 0)
        total_horizontal = cp.get('swr_total_horizontal_ft', 0)
        total_quantity = cp.get('swr_total_quantity', 0)

        # --- Updated Cost Calculations with MOQ Logic ---
        # Glass (Cat 15) - using total_area
        if mat_glass:
            required_glass = total_area
            if required_glass > mat_glass.quantity:
                delta = required_glass - mat_glass.quantity
                moq = mat_glass.moq if mat_glass.moq is not None else 0
                if delta <= moq:
                    cost_glass = mat_glass.moq_cost
                    flag_glass = "MOQ applied"
                else:
                    cost_glass = (required_glass * mat_glass.yield_cost) / cp['yield_cat15']
                    flag_glass = ""
            else:
                cost_glass = (required_glass * mat_glass.yield_cost) / cp['yield_cat15']
                flag_glass = ""
        else:
            cost_glass = 0
            flag_glass = ""

        # Extrusions (Cat 1) - using total_perimeter
        if mat_aluminum:
            required_aluminum = total_perimeter
            if required_aluminum > mat_aluminum.quantity:
                delta = required_aluminum - mat_aluminum.quantity
                moq = mat_aluminum.moq if mat_aluminum.moq is not None else 0
                if delta <= moq:
                    cost_aluminum = mat_aluminum.moq_cost
                    flag_aluminum = "MOQ applied"
                else:
                    cost_aluminum = (required_aluminum * mat_aluminum.yield_cost) / cp['yield_aluminum']
                    flag_aluminum = ""
            else:
                cost_aluminum = (required_aluminum * mat_aluminum.yield_cost) / cp['yield_aluminum']
                flag_aluminum = ""
        else:
            cost_aluminum = 0
            flag_aluminum = ""

        # Retainer (Cat 17) - handling based on option
        if cp.get("retainer_option") == "head_retainer":
            if mat_retainer:
                required_retainer = 0.5 * total_horizontal
                if required_retainer > mat_retainer.quantity:
                    delta = required_retainer - mat_retainer.quantity
                    moq = mat_retainer.moq if mat_retainer.moq is not None else 0
                    if delta <= moq:
                        cost_retainer = mat_retainer.moq_cost
                        flag_retainer = "MOQ applied"
                    else:
                        cost_retainer = (required_retainer * mat_retainer.yield_cost) / cp['yield_aluminum']
                        flag_retainer = ""
                else:
                    cost_retainer = (required_retainer * mat_retainer.yield_cost) / cp['yield_aluminum']
                    flag_retainer = ""
            else:
                cost_retainer = 0
                flag_retainer = ""
        elif cp.get("retainer_option") == "head_and_sill":
            if mat_retainer:
                required_retainer = total_horizontal
                if required_retainer > mat_retainer.quantity:
                    delta = required_retainer - mat_retainer.quantity
                    moq = mat_retainer.moq if mat_retainer.moq is not None else 0
                    if delta <= moq:
                        cost_retainer = mat_retainer.moq_cost
                        flag_retainer = "MOQ applied"
                    else:
                        cost_retainer = (required_retainer * mat_retainer.yield_cost * cp['yield_aluminum'])
                        flag_retainer = ""
                else:
                    cost_retainer = (required_retainer * mat_retainer.yield_cost * cp['yield_aluminum'])
                    flag_retainer = ""
            else:
                cost_retainer = 0
                flag_retainer = ""
        else:
            cost_retainer = 0
            flag_retainer = ""

        # Glazing Spline (Cat 2) - using total_perimeter
        if mat_glazing:
            required_glazing = total_perimeter
            if required_glazing > mat_glazing.quantity:
                delta = required_glazing - mat_glazing.quantity
                moq = mat_glazing.moq if mat_glazing.moq is not None else 0
                if delta <= moq:
                    cost_glazing = mat_glazing.moq_cost
                    flag_glazing = "MOQ applied"
                else:
                    cost_glazing = (required_glazing * mat_glazing.yield_cost) / cp['yield_cat2']
                    flag_glazing = ""
            else:
                cost_glazing = (required_glazing * mat_glazing.yield_cost) / cp['yield_cat2']
                flag_glazing = ""
        else:
            cost_glazing = 0
            flag_glazing = ""

        # Gaskets (Cat 3) - using total_vertical
        if mat_gaskets:
            required_gaskets = total_vertical
            if required_gaskets > mat_gaskets.quantity:
                delta = required_gaskets - mat_gaskets.quantity
                moq = mat_gaskets.moq if mat_gaskets.moq is not None else 0
                if delta <= moq:
                    cost_gaskets = mat_gaskets.moq_cost
                    flag_gaskets = "MOQ applied"
                else:
                    cost_gaskets = (required_gaskets * mat_gaskets.yield_cost) / cp['yield_cat3']
                    flag_gaskets = ""
            else:
                cost_gaskets = (required_gaskets * mat_gaskets.yield_cost) / cp['yield_cat3']
                flag_gaskets = ""
        else:
            cost_gaskets = 0
            flag_gaskets = ""

        # Corner Keys (Cat 4) - using total_quantity * 4
        if mat_corner_keys:
            required_corner_keys = total_quantity * 4
            if required_corner_keys > mat_corner_keys.quantity:
                delta = required_corner_keys - mat_corner_keys.quantity
                moq = mat_corner_keys.moq if mat_corner_keys.moq is not None else 0
                if delta <= moq:
                    cost_corner_keys = mat_corner_keys.moq_cost
                    flag_corner_keys = "MOQ applied"
                else:
                    cost_corner_keys = (required_corner_keys * mat_corner_keys.yield_cost) / cp['yield_cat4']
                    flag_corner_keys = ""
            else:
                cost_corner_keys = (required_corner_keys * mat_corner_keys.yield_cost) / cp['yield_cat4']
                flag_corner_keys = ""
        else:
            cost_corner_keys = 0
            flag_corner_keys = ""

        # Dual Lock (Cat 5) - using total_quantity
        if mat_dual_lock:
            required_dual_lock = total_quantity
            if required_dual_lock > mat_dual_lock.quantity:
                delta = required_dual_lock - mat_dual_lock.quantity
                moq = mat_dual_lock.moq if mat_dual_lock.moq is not None else 0
                if delta <= moq:
                    cost_dual_lock = mat_dual_lock.moq_cost
                    flag_dual_lock = "MOQ applied"
                else:
                    cost_dual_lock = (required_dual_lock * mat_dual_lock.yield_cost) / cp['yield_cat5']
                    flag_dual_lock = ""
            else:
                cost_dual_lock = (required_dual_lock * mat_dual_lock.yield_cost) / cp['yield_cat5']
                flag_dual_lock = ""
        else:
            cost_dual_lock = 0
            flag_dual_lock = ""

        # Foam Baffle Top (Cat 6) - using 0.5 * total_horizontal
        if mat_foam_baffle_top:
            required_foam_baffle_top = 0.5 * total_horizontal
            if required_foam_baffle_top > mat_foam_baffle_top.quantity:
                delta = required_foam_baffle_top - mat_foam_baffle_top.quantity
                moq = mat_foam_baffle_top.moq if mat_foam_baffle_top.moq is not None else 0
                if delta <= moq:
                    cost_foam_baffle_top = mat_foam_baffle_top.moq_cost
                    flag_foam_baffle_top = "MOQ applied"
                else:
                    cost_foam_baffle_top = (required_foam_baffle_top * mat_foam_baffle_top.yield_cost) / cp[
                        'yield_cat6']
                    flag_foam_baffle_top = ""
            else:
                cost_foam_baffle_top = (required_foam_baffle_top * mat_foam_baffle_top.yield_cost) / cp['yield_cat6']
                flag_foam_baffle_top = ""
        else:
            cost_foam_baffle_top = 0
            flag_foam_baffle_top = ""

        # Foam Baffle Bottom (Cat 6) - using 0.5 * total_horizontal
        if mat_foam_baffle_bottom:
            required_foam_baffle_bottom = 0.5 * total_horizontal
            if required_foam_baffle_bottom > mat_foam_baffle_bottom.quantity:
                delta = required_foam_baffle_bottom - mat_foam_baffle_bottom.quantity
                moq = mat_foam_baffle_bottom.moq if mat_foam_baffle_bottom.moq is not None else 0
                if delta <= moq:
                    cost_foam_baffle_bottom = mat_foam_baffle_bottom.moq_cost
                    flag_foam_baffle_bottom = "MOQ applied"
                else:
                    cost_foam_baffle_bottom = (required_foam_baffle_bottom * mat_foam_baffle_bottom.yield_cost) / cp[
                        'yield_cat6']
                    flag_foam_baffle_bottom = ""
            else:
                cost_foam_baffle_bottom = (required_foam_baffle_bottom * mat_foam_baffle_bottom.yield_cost) / cp[
                    'yield_cat6']
                flag_foam_baffle_bottom = ""
        else:
            cost_foam_baffle_bottom = 0
            flag_foam_baffle_bottom = ""

        # Glass Protection (Cat 7) - using total_area
        if mat_glass_protection:
            required_glass_protection = total_area
            if required_glass_protection > mat_glass_protection.quantity:
                delta = required_glass_protection - mat_glass_protection.quantity
                moq = mat_glass_protection.moq if mat_glass_protection.moq is not None else 0
                if delta <= moq:
                    cost_glass_protection = mat_glass_protection.moq_cost
                    flag_glass_protection = "MOQ applied"
                else:
                    cost_glass_protection = (required_glass_protection * mat_glass_protection.yield_cost) / cp[
                        'yield_cat7']
                    flag_glass_protection = ""
            else:
                cost_glass_protection = (required_glass_protection * mat_glass_protection.yield_cost) / cp['yield_cat7']
                flag_glass_protection = ""
        else:
            cost_glass_protection = 0
            flag_glass_protection = ""

        # Tape (Cat 10)
        if mat_tape:
            if cp.get("retainer_attachment_option") == 'head_retainer':
                required_tape = total_horizontal / 2
            elif cp.get("retainer_attachment_option") == 'head_sill':
                required_tape = total_horizontal
            else:
                required_tape = 0
            if required_tape > mat_tape.quantity:
                delta = required_tape - mat_tape.quantity
                moq = mat_tape.moq if mat_tape.moq is not None else 0
                if delta <= moq:
                    cost_tape = mat_tape.moq_cost
                    flag_tape = "MOQ applied"
                else:
                    cost_tape = (required_tape * mat_tape.yield_cost) / cp['yield_cat10']
                    flag_tape = ""
            else:
                cost_tape = (required_tape * mat_tape.yield_cost) / cp['yield_cat10']
                flag_tape = ""
        else:
            cost_tape = 0
            flag_tape = ""

        # Screws (Cat 18)
        if mat_screws:
            if cp.get("screws_option") == "head_retainer":
                required_screws = 0.5 * total_horizontal * 4
            elif cp.get("screws_option") == "head_and_sill":
                required_screws = total_horizontal * 4
            else:
                required_screws = 0
            if required_screws > mat_screws.quantity:
                delta = required_screws - mat_screws.quantity
                moq = mat_screws.moq if mat_screws.moq is not None else 0
                if delta <= moq:
                    cost_screws = mat_screws.moq_cost
                    flag_screws = "MOQ applied"
                else:
                    cost_screws = (required_screws * mat_screws.yield_cost)
                    flag_screws = ""
            else:
                cost_screws = (required_screws * mat_screws.yield_cost)
                flag_screws = ""
        else:
            cost_screws = 0
            flag_screws = ""

        # Setting Block (Cat 16) - using 2 per panel
        if mat_setting_block:
            required_setting_block = total_quantity * 2
            if required_setting_block > mat_setting_block.quantity:
                delta = required_setting_block - mat_setting_block.quantity
                moq = mat_setting_block.moq if mat_setting_block.moq is not None else 0
                if delta <= moq:
                    cost_setting_block = mat_setting_block.moq_cost
                    flag_setting_block = "MOQ applied"
                else:
                    cost_setting_block = (required_setting_block * mat_setting_block.yield_cost) / cp['yield_cat16']
                    flag_setting_block = ""
            else:
                cost_setting_block = (required_setting_block * mat_setting_block.yield_cost) / cp['yield_cat16']
                flag_setting_block = ""
        else:
            cost_setting_block = 0
            flag_setting_block = ""

        total_material_cost = (cost_glass + cost_aluminum + cost_retainer + cost_glazing + cost_gaskets +
                               cost_corner_keys + cost_dual_lock + cost_foam_baffle_top + cost_foam_baffle_bottom +
                               cost_glass_protection + cost_tape + cost_screws + cost_setting_block)
        cp['material_total_cost'] = total_material_cost

        # Build the full materials list with Calculation fields updated to show "MOQ Cost" if the MOQ cost was applied.
        materials_list = [
            {
                "Category": "Glass (Cat 15)",
                "Selected Material": mat_glass.nickname if mat_glass else "N/A",
                "Unit Cost": mat_glass.yield_cost if mat_glass else 0,
                "Calculation": "MOQ Cost" if flag_glass else f"Total Area {total_area:.2f} × Yield Cost / {cp['yield_cat15']}",
                "Cost ($)": cost_glass
            },
            {
                "Category": "Extrusions (Cat 1)",
                "Selected Material": mat_aluminum.nickname if mat_aluminum else "N/A",
                "Unit Cost": mat_aluminum.yield_cost if mat_aluminum else 0,
                "Calculation": "MOQ Cost" if flag_aluminum else f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {cp['yield_aluminum']}",
                "Cost ($)": cost_aluminum
            },
            {
                "Category": "Retainer (Cat 17)",
                "Selected Material": (mat_retainer.nickname if mat_retainer else "N/A") if cp.get(
                    "retainer_option") != "no_retainer" else "N/A",
                "Unit Cost": (mat_retainer.yield_cost if mat_retainer else 0) if cp.get(
                    "retainer_option") != "no_retainer" else 0,
                "Calculation": (
                    (
                        "MOQ Cost" if flag_retainer else f"Head Retainer: 0.5 × Total Horizontal {total_horizontal:.2f} × Yield Cost / {cp['yield_aluminum']}")
                    if cp.get("retainer_option") == "head_retainer"
                    else ((
                              "MOQ Cost" if flag_retainer else f"Head + Sill Retainer: Total Horizontal {total_horizontal:.2f} × Yield Cost × {cp['yield_aluminum']}") if cp.get(
                        "retainer_option") == "head_and_sill" else "No Retainer")
                ),
                "Cost ($)": cost_retainer
            },
            {
                "Category": "Glazing Spline (Cat 2)",
                "Selected Material": mat_glazing.nickname if mat_glazing else "N/A",
                "Unit Cost": mat_glazing.yield_cost if mat_glazing else 0,
                "Calculation": "MOQ Cost" if flag_glazing else f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {cp['yield_cat2']}",
                "Cost ($)": cost_glazing
            },
            {
                "Category": "Gaskets (Cat 3)",
                "Selected Material": mat_gaskets.nickname if mat_gaskets else "N/A",
                "Unit Cost": mat_gaskets.yield_cost if mat_gaskets else 0,
                "Calculation": "MOQ Cost" if flag_gaskets else f"Total Vertical {total_vertical:.2f} × Yield Cost / {cp['yield_cat3']}",
                "Cost ($)": cost_gaskets
            },
            {
                "Category": "Corner Keys (Cat 4)",
                "Selected Material": mat_corner_keys.nickname if mat_corner_keys else "N/A",
                "Unit Cost": mat_corner_keys.yield_cost if mat_corner_keys else 0,
                "Calculation": "MOQ Cost" if flag_corner_keys else f"Total Quantity {total_quantity:.2f} × 4 × Yield Cost / {cp['yield_cat4']}",
                "Cost ($)": cost_corner_keys
            },
            {
                "Category": "Dual Lock (Cat 5)",
                "Selected Material": mat_dual_lock.nickname if mat_dual_lock else "N/A",
                "Unit Cost": mat_dual_lock.yield_cost if mat_dual_lock else 0,
                "Calculation": "MOQ Cost" if flag_dual_lock else f"Total Quantity {total_quantity:.2f} × Yield Cost / {cp['yield_cat5']}",
                "Cost ($)": cost_dual_lock
            },
            {
                "Category": "Foam Baffle Top/Head (Cat 6)",
                "Selected Material": mat_foam_baffle_top.nickname if mat_foam_baffle_top else "N/A",
                "Unit Cost": mat_foam_baffle_top.yield_cost if mat_foam_baffle_top else 0,
                "Calculation": "MOQ Cost" if flag_foam_baffle_top else f"0.5 × Total Horizontal {total_horizontal:.2f} × Yield Cost / {cp['yield_cat6']}",
                "Cost ($)": cost_foam_baffle_top
            },
            {
                "Category": "Foam Baffle Bottom/Sill (Cat 6)",
                "Selected Material": mat_foam_baffle_bottom.nickname if mat_foam_baffle_bottom else "N/A",
                "Unit Cost": mat_foam_baffle_bottom.yield_cost if mat_foam_baffle_bottom else 0,
                "Calculation": "MOQ Cost" if flag_foam_baffle_bottom else f"0.5 × Total Horizontal {total_horizontal:.2f} × Yield Cost / {cp['yield_cat6']}",
                "Cost ($)": cost_foam_baffle_bottom
            },
            {
                "Category": "Glass Protection (Cat 7)",
                "Selected Material": mat_glass_protection.nickname if mat_glass_protection else "N/A",
                "Unit Cost": mat_glass_protection.yield_cost if mat_glass_protection else 0,
                "Calculation": (
                    "MOQ Cost" if flag_glass_protection else f"Total Area {total_area:.2f} × Yield Cost / {cp['yield_cat7']}") if cp.get(
                    "glass_protection_side") == "one"
                else ((
                          "MOQ Cost" if flag_glass_protection else f"Total Area {total_area:.2f} × Yield Cost × 2 / {cp['yield_cat7']}") if cp.get(
                    "glass_protection_side") == "double" else "No Film"),
                "Cost ($)": cost_glass_protection
            },
            {
                "Category": "Tape (Cat 10)",
                "Selected Material": mat_tape.nickname if mat_tape else "N/A",
                "Unit Cost": mat_tape.yield_cost if mat_tape else 0,
                "Calculation": "MOQ Cost" if flag_tape else "Retainer Attachment Option: " + (
                    "(Head Retainer - Half Horizontal)" if cp.get("retainer_attachment_option") == "head_retainer"
                    else ("(Head+Sill - Full Horizontal)" if cp.get(
                        "retainer_attachment_option") == "head_sill" else "No Tape")),
                "Cost ($)": cost_tape
            },
            {
                "Category": "Screws (Cat 18)",
                "Selected Material": mat_screws.nickname if mat_screws else "N/A",
                "Unit Cost": mat_screws.yield_cost if mat_screws else 0,
                "Calculation": ("MOQ Cost" if flag_screws else (
                    f"Head Retainer: 0.5 × Total Horizontal {total_horizontal:.2f} × 4 × Yield Cost" if cp.get(
                        "screws_option") == "head_retainer"
                    else (f"Head + Sill: Total Horizontal {total_horizontal:.2f} × 4 × Yield Cost" if cp.get(
                        "screws_option") == "head_and_sill" else "No Screws"))),
                "Cost ($)": cost_screws
            },
            # --- Added Setting Block Item ---
            {
                "Category": "Setting Block (Cat 16)",
                "Selected Material": mat_setting_block.nickname if mat_setting_block else "N/A",
                "Unit Cost": mat_setting_block.yield_cost if mat_setting_block else 0,
                "Calculation": "MOQ Cost" if flag_setting_block else f"Total Panels {total_quantity:.2f} × 2 × Yield Cost / {cp['yield_cat16']}",
                "Cost ($)": cost_setting_block
            },
        ]
        cp["itemized_costs"] = materials_list
        save_current_project(cp)
        next_button = '<button type="button" class="btn" onclick="window.location.href=\'/other_costs\'">Next: Additional Costs</button>'
        result_html = f"""
         <html>
           <head>
             <title>SWR Material Cost Summary</title>
             <style>{common_css}</style>
           </head>
           <body>
             <div class="container">
               <h2>SWR Material Cost Summary</h2>
               <form method="POST">
                 <table class="summary-table">
                   <tr>{"".join(f"<th>{h}</th>" for h in ["Category", "Selected Material", "Unit Cost", "Calculation", "Cost ($)", "$ per SF", "% Total Cost", "Stock Level", "Min Lead", "Max Lead", "Discount/Surcharge", "Final Cost"])}</tr>
                   {''.join(f"""
                   <tr>
                     <td>{item['Category']}</td>
                     <td>{item['Selected Material']}</td>
                     <td>{item['Unit Cost']:.2f}</td>
                     <td>{item['Calculation']}</td>
                     <td class="cost">{item['Cost ($)']:.2f}</td>
                     <td>{(item['Cost ($)'] / total_area if total_area > 0 else 0):.2f}</td>
                     <td>{(item['Cost ($)'] / total_material_cost * 100 if total_material_cost > 0 else 0):.2f}</td>
                     <td>N/A</td>
                     <td>N/A</td>
                     <td>N/A</td>
                     <td><input type="number" step="1" name="discount_{item["Category"].replace(" ", "_")}" value="0" oninput="updateFinalCost(this)" /></td>
                     <td class="final-cost">{item['Cost ($)']:.2f}</td>
                   </tr>
                   """ for item in materials_list)}
                 </table>
                 <div style="margin-top:10px;">
                   <label for="swr_note">Materials Note:</label>
                   <textarea id="swr_note" name="swr_note" rows="3" style="width:100%;">{cp.get("swr_note", "")}</textarea>
                 </div>
                 <script>
                   function updateFinalCost(input) {{
                       var row = input.closest("tr");
                       var costCell = row.querySelector(".cost");
                       var finalCostCell = row.querySelector(".final-cost");
                       var discountValue = parseFloat(input.value) || 0;
                       var costValue = parseFloat(costCell.innerText) || 0;
                       var finalCost = costValue + discountValue;
                       finalCostCell.innerText = finalCost.toFixed(2);
                   }}
                 </script>
               </form>
               <div class="btn-group">
                 <button type="button" class="btn" onclick="window.location.href='/materials'">Back: Edit Materials</button>
                 {next_button}
               </div>
             </div>
           </body>
         </html>
        """
        return result_html
    else:
        # GET branch: Render the form without the duplicate Head Retainer fields.
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
               <div>
                  <label for="yield_cat15">Glass (Cat 15) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat15" name="yield_cat15" value="{cp.get('yield_cat15', '0.97')}" required>
               </div>
               <div>
                  <label for="material_glass">Select Glass:</label>
                  <select name="material_glass" id="material_glass" required>
                     {generate_options(materials_glass, cp.get("material_glass"))}
                  </select>
               </div>
               <!-- New Glass Thickness Dropdown -->
               <div>
                  <label for="glass_thickness">Glass Thickness (mm):</label>
                  <select name="glass_thickness" id="glass_thickness" required>
                     <option value="4">4mm</option>
                     <option value="5">5mm</option>
                     <option value="6">6mm</option>
                     <option value="8">8mm</option>
                     <option value="10">10mm</option>
                  </select>
               </div>
               <div>
                  <label for="yield_aluminum">Extrusions (Cat 1) Yield:</label>
                  <input type="number" step="0.01" id="yield_aluminum" name="yield_aluminum" value="{cp.get('yield_aluminum', '0.75')}" required>
               </div>
               <div>
                  <label for="material_aluminum">Select Extrusions:</label>
                  <select name="material_aluminum" id="material_aluminum" required>
                     {generate_options(materials_aluminum, cp.get("material_aluminum"))}
                  </select>
               </div>
               <div>
                  <label for="retainer_option">Retainer Option:</label>
                  <select name="retainer_option" id="retainer_option" required>
                     <option value="head_retainer" {"selected" if cp.get('retainer_option', '') == "head_retainer" else ""}>Head Retainer</option>
                     <option value="head_and_sill" {"selected" if cp.get('retainer_option', '') == "head_and_sill" else ""}>Head + Sill Retainer</option>
                     <option value="no_retainer" {"selected" if cp.get('retainer_option', '') == "no_retainer" else ""}>No Retainer</option>
                  </select>
               </div>
               <div>
                  <label for="material_retainer">Select Retainer Material:</label>
                  <select name="material_retainer" id="material_retainer" required>
                     {generate_options(materials_head_retainers, cp.get("material_retainer"))}
                  </select>
               </div>
               <div>
                  <label for="yield_cat2">Glazing Spline (Cat 2) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat2" name="yield_cat2" value="{cp.get('yield_cat2', '0.91')}" required>
               </div>
               <div>
                  <label for="material_glazing">Select Glazing Spline:</label>
                  <select name="material_glazing" id="material_glazing" required>
                     {generate_options(materials_glazing, cp.get("material_glazing"))}
                  </select>
               </div>
               <div>
                  <label for="yield_cat3">Gaskets (Cat 3) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat3" name="yield_cat3" value="{cp.get('yield_cat3', '0.91')}" required>
               </div>
               <div>
                  <label for="material_gaskets">Select Gaskets:</label>
                  <select name="material_gaskets" id="material_gaskets" required>
                     {generate_options(materials_gaskets, cp.get("material_gaskets"))}
                  </select>
               </div>
               <div>
                  <label for="jamb_plate">Jamb Plate:</label>
                  <select name="jamb_plate" id="jamb_plate" required>
                     <option value="Yes" {"selected" if cp.get('jamb_plate', '') == "Yes" else ""}>Yes</option>
                     <option value="No" {"selected" if cp.get('jamb_plate', '') == "No" else ""}>No</option>
                  </select>
               </div>
               <div>
                  <label for="jamb_plate_screws">Jamb Plate Screws:</label>
                  <select name="jamb_plate_screws" id="jamb_plate_screws" required>
                     <option value="Yes" {"selected" if cp.get('jamb_plate_screws', '') == "Yes" else ""}>Yes</option>
                     <option value="No" {"selected" if cp.get('jamb_plate_screws', '') == "No" else ""}>No</option>
                  </select>
               </div>
               <div>
                  <label for="yield_cat4">Corner Keys (Cat 4) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat4" name="yield_cat4" value="{cp.get('yield_cat4', '0.91')}" required>
               </div>
               <div>
                  <label for="material_corner_keys">Select Corner Keys:</label>
                  <select name="material_corner_keys" id="material_corner_keys" required>
                     {generate_options(materials_corner_keys, cp.get("material_corner_keys"))}
                  </select>
               </div>
               <div>
                  <label for="yield_cat5">Dual Lock (Cat 5) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat5" name="yield_cat5" value="{cp.get('yield_cat5', '0.91')}" required>
               </div>
               <div>
                  <label for="material_dual_lock">Select Dual Lock:</label>
                  <select name="material_dual_lock" id="material_dual_lock" required>
                     {generate_options(materials_dual_lock, cp.get("material_dual_lock"))}
                  </select>
               </div>
               <div>
                  <label for="yield_cat6">Foam Baffle Yield (Cat 6):</label>
                  <input type="number" step="0.01" id="yield_cat6" name="yield_cat6" value="{cp.get('yield_cat6', '0.91')}" required>
               </div>
               <div>
                  <label for="material_foam_baffle">Select Foam Baffle Top/Head:</label>
                  <select name="material_foam_baffle" id="material_foam_baffle" required>
                     {generate_options(materials_foam_baffle, cp.get("material_foam_baffle"))}
                  </select>
               </div>
               <div>
                  <label for="material_foam_baffle_bottom">Select Foam Baffle Bottom/Sill:</label>
                  <select name="material_foam_baffle_bottom" id="material_foam_baffle_bottom" required>
                     {generate_options(materials_foam_baffle, cp.get("material_foam_baffle_bottom"))}
                  </select>
               </div>
               <div>
                  <label for="yield_cat7">Glass Protection (Cat 7) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat7" name="yield_cat7" value="{cp.get('yield_cat7', '0.91')}" required>
               </div>
               <div>
                  <label for="glass_protection_side">Glass Protection Side:</label>
                  <select name="glass_protection_side" id="glass_protection_side" required>
                     <option value="one" {"selected" if cp.get('glass_protection_side', '') == "one" else ""}>One Sided</option>
                     <option value="double" {"selected" if cp.get('glass_protection_side', '') == "double" else ""}>Double Sided</option>
                     <option value="none" {"selected" if cp.get('glass_protection_side', '') == "none" else ""}>No Film</option>
                  </select>
               </div>
               <div>
                  <label for="material_glass_protection">Select Glass Protection:</label>
                  <select name="material_glass_protection" id="material_glass_protection" required>
                     {generate_options(materials_glass_protection, cp.get("material_glass_protection"))}
                  </select>
               </div>
               <div>
                  <label for="yield_cat10">Tape (Cat 10) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat10" name="yield_cat10" value="{cp.get('yield_cat10', '0.91')}" required>
               </div>
               <div>
                  <label for="retainer_attachment_option">Retainer Attachment Option:</label>
                  <select name="retainer_attachment_option" id="retainer_attachment_option" required>
                     <option value="head_retainer" {"selected" if cp.get('retainer_attachment_option', '') == "head_retainer" else ""}>Head Retainer (Half Horizontal)</option>
                     <option value="head_sill" {"selected" if cp.get('retainer_attachment_option', '') == "head_sill" else ""}>Head+Sill (Full Horizontal)</option>
                     <option value="no_tape" {"selected" if cp.get('retainer_attachment_option', '') == "no_tape" else ""}>No Tape</option>
                  </select>
               </div>
               <div>
                  <label for="material_tape">Select Tape Material:</label>
                  <select name="material_tape" id="material_tape" required>
                     {generate_options(materials_tape, cp.get("material_tape"))}
                  </select>
               </div>
               <!-- Added Setting Block Fields -->
               <div>
                  <label for="yield_cat16">Setting Block (Cat 16) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat16" name="yield_cat16" value="{cp.get('yield_cat16', '1.0')}" required>
               </div>
               <div>
                  <label for="material_setting_block">Select Setting Block:</label>
                  <select name="material_setting_block" id="material_setting_block" required>
                     {generate_options(materials_setting_block, cp.get("material_setting_block"))}
                  </select>
               </div>
               <!-- Duplicate Head Retainer fields have been removed here -->
               <div class="btn-group">
                  <button type="button" class="btn" onclick="window.location.href='/summary'">Back: Edit Summary</button>
                  <button type="submit" class="btn">Calculate SWR Material Costs</button>
               </div>
            </form>
         </div>
      </body>
    </html>
    """


@app.route('/igr_materials', methods=['GET', 'POST'])
def igr_materials():
    cp = get_current_project()
    try:
        igr_glass = Material.query.filter_by(category=15).all()
        igr_extrusions = Material.query.filter_by(category=1).all()
        igr_gaskets = Material.query.filter_by(category=3).all()
        igr_glass_protection = Material.query.filter_by(category=7).all()
        igr_tape = Material.query.filter_by(category=10).all()
    except Exception as e:
        return f"<h2 style='color: red;'>Error fetching IGR materials: {e}</h2>"
    if request.method == 'POST':
        try:
            cp['yield_igr_glass'] = float(request.form.get('yield_igr_glass', 0.97))
        except:
            cp['yield_igr_glass'] = 0.97
        try:
            cp['yield_igr_extrusions'] = float(request.form.get('yield_igr_extrusions', 0.75))
        except:
            cp['yield_igr_extrusions'] = 0.75
        try:
            cp['yield_igr_gaskets'] = float(request.form.get('yield_igr_gaskets', 0.91))
        except:
            cp['yield_igr_gaskets'] = 0.91
        try:
            cp['yield_igr_glass_protection'] = float(request.form.get('yield_igr_glass_protection', 0.91))
        except:
            cp['yield_igr_glass_protection'] = 0.91
        try:
            cp['yield_igr_perimeter_tape'] = float(request.form.get('yield_igr_perimeter_tape', 0.91))
        except:
            cp['yield_igr_perimeter_tape'] = 0.91
        try:
            cp['yield_igr_structural_tape'] = float(request.form.get('yield_igr_structural_tape', 0.91))
        except:
            cp['yield_igr_structural_tape'] = 0.91

        cp["igr_material_glass"] = request.form.get('igr_material_glass')
        # Capture the new IGR glass thickness value
        cp["igr_glass_thickness"] = request.form.get("igr_glass_thickness")
        cp["igr_material_extrusions"] = request.form.get('igr_material_extrusions')
        cp["igr_material_gaskets"] = request.form.get('igr_material_gaskets')
        cp["igr_material_glass_protection"] = request.form.get('igr_material_glass_protection')
        cp["igr_material_perimeter_tape"] = request.form.get('igr_material_perimeter_tape')
        cp["igr_material_structural_tape"] = request.form.get('igr_material_structural_tape')
        cp["igr_note"] = request.form.get("igr_note", "")
        save_current_project(cp)

        mat_igr_glass = Material.query.get(cp.get("igr_material_glass")) if cp.get("igr_material_glass") else None
        mat_igr_extrusions = Material.query.get(cp.get("igr_material_extrusions")) if cp.get(
            "igr_material_extrusions") else None
        mat_igr_gaskets = Material.query.get(cp.get("igr_material_gaskets")) if cp.get("igr_material_gaskets") else None
        mat_igr_glass_protection = Material.query.get(cp.get("igr_material_glass_protection")) if cp.get(
            "igr_material_glass_protection") else None
        mat_igr_perimeter_tape = Material.query.get(cp.get("igr_material_perimeter_tape")) if cp.get(
            "igr_material_perimeter_tape") else None
        mat_igr_structural_tape = Material.query.get(cp.get("igr_material_structural_tape")) if cp.get(
            "igr_material_structural_tape") else None

        total_area = cp.get('igr_total_area', 0)
        total_perimeter = cp.get('igr_total_perimeter', 0)
        total_vertical = cp.get('igr_total_vertical_ft', 0)

        # --- IGR Glass (Cat 15) ---
        if mat_igr_glass:
            required_igr_glass = total_area
            if required_igr_glass > mat_igr_glass.quantity:
                delta = required_igr_glass - mat_igr_glass.quantity
                moq = mat_igr_glass.moq if mat_igr_glass.moq is not None else 0
                if delta <= moq:
                    cost_igr_glass = mat_igr_glass.moq_cost
                    flag_igr_glass = "MOQ applied"
                else:
                    cost_igr_glass = (required_igr_glass * mat_igr_glass.yield_cost) / cp['yield_igr_glass']
                    flag_igr_glass = ""
            else:
                cost_igr_glass = (required_igr_glass * mat_igr_glass.yield_cost) / cp['yield_igr_glass']
                flag_igr_glass = ""
        else:
            cost_igr_glass = 0
            flag_igr_glass = ""

        # --- IGR Extrusions (Cat 1) ---
        if mat_igr_extrusions:
            required_igr_extrusions = total_perimeter
            if required_igr_extrusions > mat_igr_extrusions.quantity:
                delta = required_igr_extrusions - mat_igr_extrusions.quantity
                moq = mat_igr_extrusions.moq if mat_igr_extrusions.moq is not None else 0
                if delta <= moq:
                    cost_igr_extrusions = mat_igr_extrusions.moq_cost
                    flag_igr_extrusions = "MOQ applied"
                else:
                    cost_igr_extrusions = (required_igr_extrusions * mat_igr_extrusions.yield_cost) / cp[
                        'yield_igr_extrusions']
                    flag_igr_extrusions = ""
            else:
                cost_igr_extrusions = (required_igr_extrusions * mat_igr_extrusions.yield_cost) / cp[
                    'yield_igr_extrusions']
                flag_igr_extrusions = ""
        else:
            cost_igr_extrusions = 0
            flag_igr_extrusions = ""

        # --- IGR Gaskets (Cat 3) ---
        if mat_igr_gaskets:
            required_igr_gaskets = total_vertical
            if required_igr_gaskets > mat_igr_gaskets.quantity:
                delta = required_igr_gaskets - mat_igr_gaskets.quantity
                moq = mat_igr_gaskets.moq if mat_igr_gaskets.moq is not None else 0
                if delta <= moq:
                    cost_igr_gaskets = mat_igr_gaskets.moq_cost
                    flag_igr_gaskets = "MOQ applied"
                else:
                    cost_igr_gaskets = (required_igr_gaskets * mat_igr_gaskets.yield_cost) / cp['yield_igr_gaskets']
                    flag_igr_gaskets = ""
            else:
                cost_igr_gaskets = (required_igr_gaskets * mat_igr_gaskets.yield_cost) / cp['yield_igr_gaskets']
                flag_igr_gaskets = ""
        else:
            cost_igr_gaskets = 0
            flag_igr_gaskets = ""

        # --- IGR Glass Protection (Cat 7) ---
        if mat_igr_glass_protection:
            required_igr_glass_protection = total_area
            if required_igr_glass_protection > mat_igr_glass_protection.quantity:
                delta = required_igr_glass_protection - mat_igr_glass_protection.quantity
                moq = mat_igr_glass_protection.moq if mat_igr_glass_protection.moq is not None else 0
                if delta <= moq:
                    cost_igr_glass_protection = mat_igr_glass_protection.moq_cost
                    flag_igr_glass_protection = "MOQ applied"
                else:
                    cost_igr_glass_protection = (required_igr_glass_protection * mat_igr_glass_protection.yield_cost) / \
                                                cp['yield_igr_glass_protection']
                    flag_igr_glass_protection = ""
            else:
                cost_igr_glass_protection = (required_igr_glass_protection * mat_igr_glass_protection.yield_cost) / cp[
                    'yield_igr_glass_protection']
                flag_igr_glass_protection = ""
        else:
            cost_igr_glass_protection = 0
            flag_igr_glass_protection = ""

        # --- IGR Perimeter Butyl Tape (Cat 10) ---
        if cp.get('igr_type') == "Dry Seal IGR":
            cost_igr_perimeter_tape = 0
            flag_igr_perimeter_tape = ""
        else:
            if mat_igr_perimeter_tape:
                required_igr_perimeter_tape = total_perimeter
                if required_igr_perimeter_tape > mat_igr_perimeter_tape.quantity:
                    delta = required_igr_perimeter_tape - mat_igr_perimeter_tape.quantity
                    moq = mat_igr_perimeter_tape.moq if mat_igr_perimeter_tape.moq is not None else 0
                    if delta <= moq:
                        cost_igr_perimeter_tape = mat_igr_perimeter_tape.moq_cost
                        flag_igr_perimeter_tape = "MOQ applied"
                    else:
                        cost_igr_perimeter_tape = (required_igr_perimeter_tape * mat_igr_perimeter_tape.yield_cost) / \
                                                  cp['yield_igr_perimeter_tape']
                        flag_igr_perimeter_tape = ""
                else:
                    cost_igr_perimeter_tape = (required_igr_perimeter_tape * mat_igr_perimeter_tape.yield_cost) / cp[
                        'yield_igr_perimeter_tape']
                    flag_igr_perimeter_tape = ""
            else:
                cost_igr_perimeter_tape = 0
                flag_igr_perimeter_tape = ""

        # --- IGR Structural Glazing Tape (Cat 10) ---
        if mat_igr_structural_tape:
            required_igr_structural_tape = total_perimeter
            if required_igr_structural_tape > mat_igr_structural_tape.quantity:
                delta = required_igr_structural_tape - mat_igr_structural_tape.quantity
                moq = mat_igr_structural_tape.moq if mat_igr_structural_tape.moq is not None else 0
                if delta <= moq:
                    cost_igr_structural_tape = mat_igr_structural_tape.moq_cost
                    flag_igr_structural_tape = "MOQ applied"
                else:
                    cost_igr_structural_tape = (required_igr_structural_tape * mat_igr_structural_tape.yield_cost) / cp[
                        'yield_igr_structural_tape']
                    flag_igr_structural_tape = ""
            else:
                cost_igr_structural_tape = (required_igr_structural_tape * mat_igr_structural_tape.yield_cost) / cp[
                    'yield_igr_structural_tape']
                flag_igr_structural_tape = ""
        else:
            cost_igr_structural_tape = 0
            flag_igr_structural_tape = ""

        total_igr_material_cost = (cost_igr_glass + cost_igr_extrusions + cost_igr_gaskets +
                                   cost_igr_glass_protection + cost_igr_perimeter_tape + cost_igr_structural_tape)
        cp['igr_material_total_cost'] = total_igr_material_cost

        # Build IGR items list with MOQ flags in the Calculation fields
        igr_items = [
            {
                "Category": "IGR Glass (Cat 15)",
                "Selected Material": mat_igr_glass.nickname if mat_igr_glass else "N/A",
                "Unit Cost": mat_igr_glass.yield_cost if mat_igr_glass else 0,
                "Calculation": "MOQ Cost" if flag_igr_glass else f"Total Area {total_area:.2f} × Yield Cost / {cp['yield_igr_glass']}",
                "Cost ($)": cost_igr_glass
            },
            {
                "Category": "IGR Extrusions (Cat 1)",
                "Selected Material": mat_igr_extrusions.nickname if mat_igr_extrusions else "N/A",
                "Unit Cost": mat_igr_extrusions.yield_cost if mat_igr_extrusions else 0,
                "Calculation": "MOQ Cost" if flag_igr_extrusions else f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {cp['yield_igr_extrusions']}",
                "Cost ($)": cost_igr_extrusions
            },
            {
                "Category": "IGR Gaskets (Cat 3)",
                "Selected Material": mat_igr_gaskets.nickname if mat_igr_gaskets else "N/A",
                "Unit Cost": mat_igr_gaskets.yield_cost if mat_igr_gaskets else 0,
                "Calculation": "MOQ Cost" if flag_igr_gaskets else f"Total Vertical {total_vertical:.2f} × Yield Cost / {cp['yield_igr_gaskets']}",
                "Cost ($)": cost_igr_gaskets
            },
            {
                "Category": "IGR Glass Protection (Cat 7)",
                "Selected Material": mat_igr_glass_protection.nickname if mat_igr_glass_protection else "N/A",
                "Unit Cost": mat_igr_glass_protection.yield_cost if mat_igr_glass_protection else 0,
                "Calculation": (
                    "MOQ Cost" if flag_igr_glass_protection else f"Total Area {total_area:.2f} × Yield Cost / {cp['yield_igr_glass_protection']}"),
                "Cost ($)": cost_igr_glass_protection
            },
            {
                "Category": "IGR Perimeter Butyl Tape (Cat 10)",
                "Selected Material": mat_igr_perimeter_tape.nickname if mat_igr_perimeter_tape else "N/A",
                "Unit Cost": mat_igr_perimeter_tape.yield_cost if mat_igr_perimeter_tape else 0,
                "Calculation": (
                    "MOQ Cost" if flag_igr_perimeter_tape else f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {cp['yield_igr_perimeter_tape']}"),
                "Cost ($)": cost_igr_perimeter_tape
            },
            {
                "Category": "IGR Structural Glazing Tape (Cat 10)",
                "Selected Material": mat_igr_structural_tape.nickname if mat_igr_structural_tape else "N/A",
                "Unit Cost": mat_igr_structural_tape.yield_cost if mat_igr_structural_tape else 0,
                "Calculation": (
                    "MOQ Cost" if flag_igr_structural_tape else f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {cp['yield_igr_structural_tape']}"),
                "Cost ($)": cost_igr_structural_tape
            }
        ]
        cp["igr_itemized_costs"] = igr_items
        save_current_project(cp)
        next_button = '<button type="button" class="btn" onclick="window.location.href=\'/other_costs\'">Next: Additional Costs</button>'
        result_html = f"""
         <html>
           <head>
             <title>IGR Material Cost Summary</title>
             <style>{common_css}</style>
           </head>
           <body>
             <div class="container">
               <h2>IGR Material Cost Summary</h2>
               <form method='POST'>
               <table class='summary-table'>
               <tr>{"".join(f"<th>{h}</th>" for h in ["Category", "Selected Material", "Unit Cost", "Calculation", "Cost ($)", "$ per SF", "% Total Cost", "Stock Level", "Min Lead", "Max Lead", "Discount/Surcharge", "Final Cost"])}</tr>
        """
        for item in igr_items:
            discount_key = "discount_" + "".join(c if c.isalnum() else "_" for c in item["Category"])
            result_html += f"""
               <tr>
                 <td>{item['Category']}</td>
                 <td>{item['Selected Material']}</td>
                 <td>{item['Unit Cost']:.2f}</td>
                 <td>{item['Calculation']}</td>
                 <td class='cost'>{item['Cost ($)']:.2f}</td>
                 <td>{(item['Cost ($)'] / total_area if total_area > 0 else 0):.2f}</td>
                 <td>{(item['Cost ($)'] / cp.get("igr_material_total_cost", 1) * 100 if cp.get("igr_material_total_cost", 0) > 0 else 0):.2f}</td>
                 <td>N/A</td>
                 <td>N/A</td>
                 <td>N/A</td>
                 <td><input type='number' step='1' name='{discount_key}' value='0' oninput='updateFinalCost(this)' /></td>
                 <td class='final-cost'>{item['Cost ($)']:.2f}</td>
               </tr>
            """
        result_html += f"""
               </table>
               <div style='margin-top:10px;'><label for='igr_note'>IGR Materials Note:</label><textarea id='igr_note' name='igr_note' rows='3' style='width:100%;'>{cp.get("igr_note", "")}</textarea></div>
               <script>
               function updateFinalCost(input) {{
                   var row = input.closest("tr");
                   var costCell = row.querySelector(".cost");
                   var finalCostCell = row.querySelector(".final-cost");
                   var discountValue = parseFloat(input.value) || 0;
                   var costValue = parseFloat(costCell.innerText) || 0;
                   var finalCost = costValue + discountValue;
                   finalCostCell.innerText = finalCost.toFixed(2);
               }}
               </script>
               </form>
               <div class="btn-group">
                  <button type="button" class="btn" onclick="window.location.href='/igr_materials'">Back: Edit IGR Materials</button>
                  {next_button}
               </div>
             </div>
           </body>
         </html>
        """
        return result_html
    else:
        return f"""
    <html>
      <head>
         <title>IGR Materials</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Select IGR Materials</h2>
            <form method="POST">
               <div>
                  <label for="yield_igr_glass">IGR Glass (Cat 15) Yield:</label>
                  <input type="number" step="0.01" id="yield_igr_glass" name="yield_igr_glass" value="{cp.get('yield_igr_glass', '0.97')}" required>
               </div>
               <div>
                  <label for="igr_material_glass">Select IGR Glass:</label>
                  <select name="igr_material_glass" id="igr_material_glass" required>
                     {generate_options(igr_glass, cp.get("igr_material_glass"))}
                  </select>
               </div>
               <!-- New IGR Glass Thickness Dropdown -->
               <div>
                  <label for="igr_glass_thickness">Glass Thickness (mm):</label>
                  <select name="igr_glass_thickness" id="igr_glass_thickness" required>
                     <option value="4">4mm</option>
                     <option value="5">5mm</option>
                     <option value="6">6mm</option>
                     <option value="8">8mm</option>
                     <option value="10">10mm</option>
                  </select>
               </div>
               <div>
                  <label for="yield_igr_extrusions">IGR Extrusions (Cat 1) Yield:</label>
                  <input type="number" step="0.01" id="yield_igr_extrusions" name="yield_igr_extrusions" value="{cp.get('yield_igr_extrusions', '0.75')}" required>
               </div>
               <div>
                  <label for="igr_material_extrusions">Select IGR Extrusions:</label>
                  <select name="igr_material_extrusions" id="igr_material_extrusions" required>
                     {generate_options(igr_extrusions, cp.get("igr_material_extrusions"))}
                  </select>
               </div>
               <div>
                  <label for="yield_igr_gaskets">IGR Gaskets (Cat 3) Yield:</label>
                  <input type="number" step="0.01" id="yield_igr_gaskets" name="yield_igr_gaskets" value="{cp.get('yield_igr_gaskets', '0.91')}" required>
               </div>
               <div>
                  <label for="igr_material_gaskets">Select IGR Gaskets:</label>
                  <select name="igr_material_gaskets" id="igr_material_gaskets" required>
                     {generate_options(igr_gaskets, cp.get("igr_material_gaskets"))}
                  </select>
               </div>
               <div>
                  <label for="yield_igr_glass_protection">IGR Glass Protection (Cat 7) Yield:</label>
                  <input type="number" step="0.01" id="yield_igr_glass_protection" name="yield_igr_glass_protection" value="{cp.get('yield_igr_glass_protection', '0.91')}" required>
               </div>
               <div>
                  <label for="igr_material_glass_protection">Select IGR Glass Protection:</label>
                  <select name="igr_material_glass_protection" id="igr_material_glass_protection" required>
                     {generate_options(igr_glass_protection, cp.get("igr_material_glass_protection"))}
                  </select>
               </div>
               <div>
                  <label for="yield_igr_perimeter_tape">Perimeter Butyl Tape (Cat 10) Yield:</label>
                  <input type="number" step="0.01" id="yield_igr_perimeter_tape" name="yield_igr_perimeter_tape" value="{cp.get('yield_igr_perimeter_tape', '0.91')}" required>
               </div>
               <div>
                  <label for="igr_material_perimeter_tape">Select IGR Perimeter Butyl Tape:</label>
                  <select name="igr_material_perimeter_tape" id="igr_material_perimeter_tape" required>
                     {generate_options(igr_tape, cp.get("igr_material_perimeter_tape"))}
                  </select>
               </div>
               <div>
                  <label for="yield_igr_structural_tape">Structural Glazing Tape (Cat 10) Yield:</label>
                  <input type="number" step="0.01" id="yield_igr_structural_tape" name="yield_igr_structural_tape" value="{cp.get('yield_igr_structural_tape', '0.91')}" required>
               </div>
               <div>
                  <label for="igr_material_structural_tape">Select Structural Glazing Tape:</label>
                  <select name="igr_material_structural_tape" id="igr_material_structural_tape" required>
                     {generate_options(igr_tape, cp.get("igr_material_structural_tape"))}
                  </select>
               </div>
               <div class="btn-group">
                  <button type="button" class="btn" onclick="window.location.href='/materials'">Back to SWR Materials</button>
                  <button type="submit" class="btn">Calculate IGR Material Costs</button>
               </div>
               <div style="margin-top:10px;">
                  <label for="igr_note">IGR Materials Note:</label>
                  <textarea id="igr_note" name="igr_note" rows="3" style="width:100%;">{cp.get("igr_note", "")}</textarea>
               </div>
            </form>
         </div>
      </body>
    </html>
    """
@app.route('/other_costs', methods=['GET', 'POST'])
def other_costs():
    cp = get_current_project()
    swr_area = cp.get("swr_total_area", 0)
    igr_area = cp.get("igr_total_area", 0)
    total_area_all = swr_area + igr_area
    total_panels = cp.get("swr_total_quantity", 0) + cp.get("igr_total_quantity", 0)
    if request.method == 'POST':
        # --- Fabrication Section ---
        fabrication_rate = float(request.form.get('fabrication_rate', 5))
        cp["fabrication_rate"] = fabrication_rate
        fabrication_cost = fabrication_rate * total_area_all

        # --- Packaging & Shipping Section ---
        num_trucks = float(request.form.get('num_trucks', 0))
        truck_cost = float(request.form.get('truck_cost', 0))
        num_crates = float(request.form.get('num_crates', 0))
        crate_cost = float(request.form.get('crate_cost', 0))
        packaging_cost = (num_trucks * truck_cost) + (num_crates * crate_cost)

        # --- Installation Section ---
        installation_option = request.form.get('installation_option')
        if installation_option == "inovues":
            labor_rate = 76.21
        elif installation_option == "nonunion":
            labor_rate = 85.0
        elif installation_option == "union":
            labor_rate = 112.0
        elif installation_option == "custom":
            labor_rate = float(request.form.get('custom_hourly_rate', 0))
            cp["custom_installation_label"] = request.form.get('custom_installation_label', "Custom")
        else:
            labor_rate = 0
        cp["installation_option"] = installation_option
        cp["custom_hourly_rate"] = request.form.get('custom_hourly_rate', '0')
        cp["custom_installation_label"] = request.form.get('custom_installation_label', "")
        hours_per_panel = float(request.form.get('hours_per_panel', 0))
        base_installation_cost = labor_rate * hours_per_panel * cp.get("swr_total_quantity", 0)
        additional_install_cost = 0
        if request.form.get('window_takeoff_checkbox'):
            additional_install_cost += float(request.form.get('window_takeoff_cost', 0))
        if request.form.get('pe_review_checkbox'):
            additional_install_cost += float(request.form.get('pe_review_cost', 0))
        if request.form.get('pm_checkbox'):
            additional_install_cost += float(request.form.get('pm_cost', 0))
        installation_cost = base_installation_cost + additional_install_cost

        # --- Equipment Section ---
        cost_scissor = float(request.form.get('cost_scissor', 0))
        cost_lull = float(request.form.get('cost_lull', 0))
        cost_baker = float(request.form.get('cost_baker', 0))
        cost_crane = float(request.form.get('cost_crane', 0))
        cost_blankets = float(request.form.get('cost_blankets', 0))
        equipment_cost = cost_scissor + cost_lull + cost_baker + cost_crane + cost_blankets

        # --- Travel Section ---
        units_per_day = float(request.form.get('units_per_day', 1))
        cp["units_per_day"] = units_per_day
        total_panels = cp.get("swr_total_quantity", 0) + cp.get("igr_total_quantity", 0)
        days_needed = total_panels / units_per_day if units_per_day > 0 else 0
        cp["days_needed"] = days_needed
        daily_rate = float(request.form.get('daily_rate', 0))
        airfare = float(request.form.get('airfare', 0))
        lodging = float(request.form.get('lodging', 0))
        meals = float(request.form.get('meals', 0))
        car_rental = float(request.form.get('car_rental', 0))
        travel_cost = airfare + (daily_rate + lodging + meals + car_rental) * days_needed
        observation = request.form.get('installation_observation', 'No')
        if observation == "Yes":
            obs_daily_rate = float(request.form.get('obs_daily_rate', 207))
            observation_cost = obs_daily_rate * days_needed
        else:
            observation_cost = 0
        travel_cost += observation_cost

        # --- Sales Section ---
        sales_items = [
            "Building Audit/Survey", "Detailed audit to inventory existing windows",
            "System Design Customization", "Thermal Stress Analysis", "Structural Analysis",
            "Thermal Performance Simulation/Analysis", "Visual & Performance Mockup",
            "CEO Time (management & development)", "Additional Design Development for nontypical conditions",
            "CFD analysis", "Window Performance M&V", "Building Energy Model",
            "Cost-Benefit Analysis", "Utility Incentive Application"
        ]
        sales_cost = 0
        for item in sales_items:
            safe_item = item.replace(" ", "_")
            if request.form.get(safe_item):
                sales_cost += float(request.form.get(safe_item + '_cost', 0))
        cp["fabrication_note"] = request.form.get("fabrication_note", "")
        cp["packaging_note"] = request.form.get("packaging_note", "")
        cp["installation_note"] = request.form.get("installation_note", "")
        cp["equipment_note"] = request.form.get("equipment_note", "")
        cp["travel_note"] = request.form.get("travel_note", "")
        cp["sales_note"] = request.form.get("sales_note", "")
        save_current_project(cp)

        additional_total = fabrication_cost + packaging_cost + installation_cost + equipment_cost + travel_cost + sales_cost
        material_cost = cp.get("material_total_cost", 0) + cp.get("igr_material_total_cost", 0)
        grand_total = material_cost + additional_total

        cp["fabrication_cost"] = fabrication_cost
        cp["packaging_cost"] = packaging_cost
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
                 <tr><th>Cost Category</th><th>Amount ($)</th></tr>
                 <tr><td>Fabrication</td><td>{fabrication_cost:.2f}</td></tr>
                 <tr><td>Packaging & Shipping</td><td>{packaging_cost:.2f}</td></tr>
                 <tr><td>Installation</td><td>{installation_cost:.2f}</td></tr>
                 <tr><td>Equipment</td><td>{equipment_cost:.2f}</td></tr>
                 <tr><td>Travel</td><td>{travel_cost:.2f}</td></tr>
                 <tr><td>Sales</td><td>{sales_cost:.2f}</td></tr>
                 <tr><td><strong>Additional Total</strong></td><td><strong>{additional_total:.2f}</strong></td></tr>
                 <tr><th>Grand Total</th><th>{grand_total:.2f}</th></tr>
               </table>
               <div style="margin-top:10px;">
                  <label for="fabrication_note">Fabrication Note:</label>
                  <textarea id="fabrication_note" name="fabrication_note" rows="2" style="width:100%;">{cp.get("fabrication_note", "")}</textarea>
               </div>
               <div style="margin-top:10px;">
                  <label for="packaging_note">Packaging & Shipping Note:</label>
                  <textarea id="packaging_note" name="packaging_note" rows="2" style="width:100%;">{cp.get("packaging_note", "")}</textarea>
               </div>
               <div style="margin-top:10px;">
                  <label for="installation_note">Installation Note:</label>
                  <textarea id="installation_note" name="installation_note" rows="2" style="width:100%;">{cp.get("installation_note", "")}</textarea>
               </div>
               <div style="margin-top:10px;">
                  <label for="equipment_note">Equipment Note:</label>
                  <textarea id="equipment_note" name="equipment_note" rows="2" style="width:100%;">{cp.get("equipment_note", "")}</textarea>
               </div>
               <div style="margin-top:10px;">
                  <label for="travel_note">Travel Note:</label>
                  <textarea id="travel_note" name="travel_note" rows="2" style="width:100%;">{cp.get("travel_note", "")}</textarea>
               </div>
               <div style="margin-top:10px;">
                  <label for="sales_note">Sales Note:</label>
                  <textarea id="sales_note" name="sales_note" rows="2" style="width:100%;">{cp.get("sales_note", "")}</textarea>
               </div>
               <div class="btn-group" style="margin-top:15px;">
                  <div class="btn-left">
                     <button type="button" class="btn" onclick="window.location.href='/igr_materials'">Back to IGR Materials</button>
                  </div>
                  <div class="btn-right">
                     <button type="button" class="btn" onclick="window.location.href='/margins'">Next: Set Margins</button>
                  </div>
               </div>
             </div>
           </body>
         </html>
        """
        return result_html
    else:
        # --- Calculate default number of crates based on glass thickness and panel type ---
        import math
        thickness_factor_lookup = {"4": 2.05, "5": 2.56, "6": 3.25, "8": 4.1, "10": 5.12, "12": 6.5}
        swr_extra = {"SWR": 0.25, "SWR-IG": 0.5, "SWR-VIG": 0.35}
        swr_thickness = cp.get("glass_thickness", "4")
        igr_thickness = cp.get("igr_glass_thickness", "4")
        base_factor_swr = thickness_factor_lookup.get(str(swr_thickness), 2.05)
        base_factor_igr = thickness_factor_lookup.get(str(igr_thickness), 2.05)
        swr_panel_type = cp.get("swr_system", "SWR")
        extra_weight_swr = swr_extra.get(swr_panel_type, 0.25)
        extra_weight_igr = 0.5  # Fixed for IGR
        factor_swr = base_factor_swr + extra_weight_swr
        factor_igr = base_factor_igr + extra_weight_igr
        crates_swr = math.ceil((swr_area * factor_swr) / 2000) if swr_area > 0 else 0
        crates_igr = math.ceil((igr_area * factor_igr) / 2000) if igr_area > 0 else 0
        default_crates = crates_swr + crates_igr

        form_html = f"""
        <html>
          <head>
             <title>Enter Additional Costs</title>
             <style>{common_css}</style>
             <script>
               function checkInstallationOption() {{
                    var sel = document.getElementById("installation_option").value;
                    var customDiv = document.getElementById("custom_installation");
                    if(sel === "custom") {{
                        customDiv.style.display = "block";
                    }} else {{
                        customDiv.style.display = "none";
                    }}
               }}
               window.onload = checkInstallationOption;
             </script>
          </head>
          <body>
             <div class="container">
                <h2>Enter Additional Costs</h2>
                <form method="POST">
                  <fieldset style="margin-bottom:20px; border:1px solid #444; padding:10px;">
                    <legend>Fabrication</legend>
                    <label for="fabrication_rate">Lump Sum Rate ($ per sq ft):</label>
                    <input type="number" step="0.01" id="fabrication_rate" name="fabrication_rate" value="{cp.get('fabrication_rate', '5')}" required>
                    <div style="margin-top:10px;">
                      <label for="fabrication_note">Fabrication Note:</label>
                      <textarea id="fabrication_note" name="fabrication_note" rows="2" style="width:100%;">{cp.get("fabrication_note", "")}</textarea>
                    </div>
                  </fieldset>
                  <fieldset style="margin-bottom:20px; border:1px solid #444; padding:10px;">
                    <legend>Packaging & Shipping</legend>
                    <label for="num_trucks">Number of Trucks:</label>
                    <input type="number" step="1" id="num_trucks" name="num_trucks" value="{cp.get('num_trucks', '0')}" required>
                    <label for="truck_cost">Cost per Truck ($):</label>
                    <input type="number" step="0.01" id="truck_cost" name="truck_cost" value="{cp.get('truck_cost', '0')}" required>
                    <label for="num_crates">Number of Crates/Racks:</label>
                    <input type="number" step="1" id="num_crates" name="num_crates" value="{cp.get('num_crates', default_crates)}" required>
                    <label for="crate_cost">Cost per Crate/Rack ($):</label>
                    <input type="number" step="0.01" id="crate_cost" name="crate_cost" value="{cp.get('crate_cost', '0')}" required>
                    <div style="margin-top:10px;">
                      <label for="packaging_note">Packaging & Shipping Note:</label>
                      <textarea id="packaging_note" name="packaging_note" rows="2" style="width:100%;">{cp.get("packaging_note", "")}</textarea>
                    </div>
                  </fieldset>
                  <fieldset style="margin-bottom:20px; border:1px solid #444; padding:10px;">
                    <legend>Installation</legend>
                    <label for="installation_option">Select Labor Rate Option:</label>
                    <select id="installation_option" name="installation_option" onchange="checkInstallationOption()" required>
                      <option value="inovues" {"selected" if cp.get('installation_option', '') == "inovues" else ""}>$76.21 INOVUES installation labor rate</option>
                      <option value="nonunion" {"selected" if cp.get('installation_option', '') == "nonunion" else ""}>$85 general non-union labor rate</option>
                      <option value="union" {"selected" if cp.get('installation_option', '') == "union" else ""}>$112 general union labor rate</option>
                      <option value="custom" {"selected" if cp.get('installation_option', '') == "custom" else ""}>Custom</option>
                    </select>
                    <div id="custom_installation">
                      <label for="custom_installation_label">Custom Labor Label:</label>
                      <input type="text" id="custom_installation_label" name="custom_installation_label" value="{cp.get('custom_installation_label', '')}">
                      <label for="custom_hourly_rate">Custom Hourly Rate ($):</label>
                      <input type="number" step="0.01" id="custom_hourly_rate" name="custom_hourly_rate" value="{cp.get('custom_hourly_rate', '0')}">
                    </div>
                    <label for="hours_per_panel">Man Hours per Panel:</label>
                    <input type="number" step="0.01" id="hours_per_panel" name="hours_per_panel" value="{cp.get('hours_per_panel', '0')}" required>
                    <p><strong>Additional Installation Tasks:</strong></p>
                    <div>
                      <input type="checkbox" id="window_takeoff_checkbox" name="window_takeoff_checkbox" {"checked" if cp.get("window_takeoff_checkbox") else ""}>
                      <label for="window_takeoff_checkbox">Window Takeoff</label>
                      <input type="number" step="0.01" id="window_takeoff_cost" name="window_takeoff_cost" placeholder="Cost ($)" value="{cp.get('window_takeoff_cost', '0')}">
                    </div>
                    <div>
                      <input type="checkbox" id="pe_review_checkbox" name="pe_review_checkbox" {"checked" if cp.get("pe_review_checkbox") else ""}>
                      <label for="pe_review_checkbox">PE Review</label>
                      <input type="number" step="0.01" id="pe_review_cost" name="pe_review_cost" placeholder="Cost ($)" value="{cp.get('pe_review_cost', '0')}">
                    </div>
                    <div>
                      <input type="checkbox" id="pm_checkbox" name="pm_checkbox" {"checked" if cp.get("pm_checkbox") else ""}>
                      <label for="pm_checkbox">PM</label>
                      <input type="number" step="0.01" id="pm_cost" name="pm_cost" placeholder="Cost ($)" value="{cp.get('pm_cost', '0')}">
                    </div>
                    <div style="margin-top:10px;">
                      <label for="installation_note">Installation Note:</label>
                      <textarea id="installation_note" name="installation_note" rows="2" style="width:100%;">{cp.get("installation_note", "")}</textarea>
                    </div>
                  </fieldset>
                  <!-- Additional fieldsets for Travel, Equipment, Sales, etc. remain unchanged -->
                  <div class="btn-group">
                     <div class="btn-left">
                         <button type="button" class="btn" onclick="window.location.href='/igr_materials'">Back to IGR Materials</button>
                     </div>
                     <div class="btn-right">
                         <button type="submit" class="btn">Calculate Additional Costs</button>
                     </div>
                  </div>
                </form>
             </div>
             <script>
               window.onload = checkInstallationOption;
             </script>
          </body>
        </html>
        """
        return form_html
@app.route('/margins', methods=['GET', 'POST'])
def margins():
    cp = get_current_project()
    # "Panel Total" = material cost + fabrication cost
    material_cost = cp.get("material_total_cost", 0) + cp.get("igr_material_total_cost", 0)
    fabrication_cost = cp.get("fabrication_cost", 0)
    base_costs = {
        "Panel Total": material_cost + fabrication_cost,
        "Packaging & Shipping": cp.get("packaging_cost", 0),
        "Installation": cp.get("installation_cost", 0),
        "Equipment": cp.get("equipment_cost", 0),
        "Travel": cp.get("travel_cost", 0),
        "Sales": cp.get("sales_cost", 0)
    }
    total_area = cp.get("swr_total_area", 0)

    if request.method == 'POST':
        # Process category margins
        margins_values = {}
        for cat in base_costs:
            try:
                margins_values[cat] = float(request.form.get(f"{cat}_margin", 0))
            except:
                margins_values[cat] = 0
        adjusted_costs = {cat: base_costs[cat] * (1 + margins_values[cat] / 100) for cat in base_costs}
        final_total = sum(adjusted_costs.values())
        cp["final_summary"] = []
        for cat in base_costs:
            cp["final_summary"].append({
                "Category": cat,
                "Original Cost ($)": base_costs[cat],
                "Margin (%)": margins_values.get(cat, 0),
                "Cost with Margin ($)": adjusted_costs[cat]
            })
        cp["grand_total"] = final_total
        save_current_project(cp)

        # Process additional pricing adjustments (all percentages)
        try:
            product_discount_pct = float(request.form.get("product_discount_pct", 0))
        except:
            product_discount_pct = 0
        try:
            finders_fee_pct = float(request.form.get("finders_fee_pct", 0))
        except:
            finders_fee_pct = 0
        try:
            sales_commission_pct = float(request.form.get("sales_commission_pct", 0))
        except:
            sales_commission_pct = 0
        try:
            sales_tax_pct = float(request.form.get("sales_tax_pct", 0))
        except:
            sales_tax_pct = 0

        cp["product_discount_pct"] = product_discount_pct
        cp["finders_fee_pct"] = finders_fee_pct
        cp["sales_commission_pct"] = sales_commission_pct
        cp["sales_tax_pct"] = sales_tax_pct

        # Note: Discount subtracts; finders fee and sales commission are added.
        panelTotal = adjusted_costs["Panel Total"]
        discount_amount = panelTotal * (product_discount_pct / 100)
        finders_fee_amount = panelTotal * (finders_fee_pct / 100)
        sales_commission_amount = panelTotal * (sales_commission_pct / 100)
        product_total = panelTotal - discount_amount + finders_fee_amount + sales_commission_amount
        sales_tax_amount = product_total * (sales_tax_pct / 100)
        product_total_after_tax = product_total + sales_tax_amount
        other_costs = (base_costs["Packaging & Shipping"] + base_costs["Installation"] +
                       base_costs["Equipment"] + base_costs["Travel"] + base_costs["Sales"])
        total_sell_price = product_total_after_tax + other_costs
        product_plus_installation = product_total + base_costs["Installation"]
        total_cost = sum(base_costs.values())
        actual_profit = total_sell_price - total_cost
        profit_margin = (actual_profit / total_sell_price * 100) if total_sell_price > 0 else 0

        cp["product_discount_amount"] = discount_amount
        cp["finders_fee_amount"] = finders_fee_amount
        cp["sales_commission_amount"] = sales_commission_amount
        cp["sales_tax_amount"] = sales_tax_amount
        cp["product_total"] = product_total
        cp["product_total_after_tax"] = product_total_after_tax
        cp["product_plus_installation"] = product_plus_installation
        cp["total_sell_price"] = total_sell_price
        cp["total_cost"] = total_cost
        cp["actual_profit"] = actual_profit
        cp["profit_margin"] = profit_margin

        # --- Compute Ultimate Lead Times (for SWR and IGR) ---
        ultimate_min_lead_material = max(cp.get("min_lead_material", 0), cp.get("igr_min_lead_material", 0))
        ultimate_max_lead_material = max(cp.get("max_lead_material", 0), cp.get("igr_max_lead_material", 0))
        ultimate_min_lead_fabrication = max(cp.get("min_lead_fabrication", 0), cp.get("igr_min_lead_fabrication", 0))
        ultimate_max_lead_fabrication = max(cp.get("max_lead_fabrication", 0), cp.get("igr_max_lead_fabrication", 0))
        ultimate_min_total_lead = max(cp.get("min_total_lead", 0), cp.get("igr_min_total_lead", 0))
        ultimate_max_total_lead = max(cp.get("max_total_lead", 0), cp.get("igr_max_total_lead", 0))

        cp["ultimate_min_lead_material"] = ultimate_min_lead_material
        cp["ultimate_max_lead_material"] = ultimate_max_lead_material
        cp["ultimate_min_lead_fabrication"] = ultimate_min_lead_fabrication
        cp["ultimate_max_lead_fabrication"] = ultimate_max_lead_fabrication
        cp["ultimate_min_total_lead"] = ultimate_min_total_lead
        cp["ultimate_max_total_lead"] = ultimate_max_total_lead

        # --- Update Final Cost for Each Material Item ---
        # Final Cost is calculated as: (Cost ($) + Discount/Surcharge) * (1 + Margin (%) / 100)
        for item in cp.get("itemized_costs", []):
            cat = item.get("Category")
            margin_pct = margins_values.get(cat, 0)
            base_cost = item.get("Cost ($)", 0)
            discount = item.get("Discount/Surcharge", 0)
            item["Final Cost"] = (base_cost + discount) * (1 + margin_pct / 100)
        save_current_project(cp)

        # Render final static result page (including updated lead time summary)
        result_html = f"""
        <html>
          <head>
            <title>Final Cost & Pricing Adjustments</title>
            <style>{common_css}</style>
          </head>
          <body>
            <div class="container">
              <h2>Final Cost & Pricing Adjustments</h2>
              <h3>Cost Summary (with Margins)</h3>
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
                  <td>{margins_values[cat]:.2f}</td>
                  <td>{adjusted_costs[cat]:.2f}</td>
                </tr>
            """
        result_html += f"""
                <tr>
                  <th colspan="3">Grand Total</th>
                  <th>{final_total:.2f}</th>
                </tr>
              </table>
              <h3>Additional Pricing Adjustments</h3>
              <table class="summary-table">
                <tr>
                  <th>Adjustment</th>
                  <th>Percentage (%)</th>
                  <th>Amount ($)</th>
                </tr>
                <tr>
                  <td>Product Discount</td>
                  <td>{product_discount_pct:.2f}</td>
                  <td>{discount_amount:.2f}</td>
                </tr>
                <tr>
                  <td>Finders Fee</td>
                  <td>{finders_fee_pct:.2f}</td>
                  <td>{finders_fee_amount:.2f}</td>
                </tr>
                <tr>
                  <td>Sales Commission</td>
                  <td>{sales_commission_pct:.2f}</td>
                  <td>{sales_commission_amount:.2f}</td>
                </tr>
                <tr>
                  <td>Sales Tax</td>
                  <td>{sales_tax_pct:.2f}</td>
                  <td>{sales_tax_amount:.2f}</td>
                </tr>
              </table>
              <h3>Product Pricing</h3>
              <table class="summary-table">
                <tr>
                  <th>Product Total ($)</th>
                  <th>Product + Installation ($)</th>
                  <th>Total Sell Price ($)</th>
                </tr>
                <tr>
                  <td>{product_total:.2f}</td>
                  <td>{product_plus_installation:.2f}</td>
                  <td>{total_sell_price:.2f}</td>
                </tr>
              </table>
              <h3>Final Analysis</h3>
              <table class="summary-table">
                <tr>
                  <th>Total Sell Price per SF</th>
                  <th>Total Sell Price per Panel</th>
                  <th>Total Cost ($)</th>
                  <th>Actual Profit ($)</th>
                  <th>Profit Margin (%)</th>
                </tr>
                <tr>
                  <td>{(total_sell_price / total_area) if total_area > 0 else 0:.2f}</td>
                  <td>{(total_sell_price / cp.get("swr_total_quantity", 1)) if cp.get("swr_total_quantity", 0) > 0 else 0:.2f}</td>
                  <td>{total_cost:.2f}</td>
                  <td>{actual_profit:.2f}</td>
                  <td>{profit_margin:.2f}</td>
                </tr>
              </table>
              <h3>Lead Time Summary (weeks)</h3>
              <table class="summary-table">
                <tr>
                  <th>Metric</th>
                  <th>Min (weeks)</th>
                  <th>Max (weeks)</th>
                </tr>
                <tr>
                  <td>Lead Material</td>
                  <td>{cp.get("ultimate_min_lead_material", 0):.2f}</td>
                  <td>{cp.get("ultimate_max_lead_material", 0):.2f}</td>
                </tr>
                <tr>
                  <td>Lead Fabrication</td>
                  <td>{cp.get("ultimate_min_lead_fabrication", 0):.2f}</td>
                  <td>{cp.get("ultimate_max_lead_fabrication", 0):.2f}</td>
                </tr>
                <tr>
                  <td>Total Lead</td>
                  <td>{cp.get("ultimate_min_total_lead", 0):.2f}</td>
                  <td>{cp.get("ultimate_max_total_lead", 0):.2f}</td>
                </tr>
              </table>
              <div class="btn-group" style="margin-top:15px;">
                <div class="btn-left">
                  <button type="button" class="btn" onclick="window.location.href='/other_costs'">Back to Additional Costs</button>
                </div>
                <div class="btn-right">
                  <button type="button" class="btn" onclick="window.location.href='/new_project'">Start New Project</button>
                </div>
              </div>
              <div style="margin-top:10px;">
                <form method="POST" action="/download_final_summary">
                  <input type="hidden" name="csv_data" value=''>
                  <button type="submit" class="btn btn-download">Download Final Summary</button>
                </form>
              </div>
            </div>
          </body>
        </html>
        """
        return result_html

    else:
        # GET branch: Render interactive form with sliders for margins and additional pricing adjustments
        return f"""
        <html>
          <head>
            <title>Set Margins & Pricing Adjustments</title>
            <style>{common_css}</style>
            <script>
              var baseCosts = {json.dumps(base_costs)};
              function updateOutput(sliderId, outputId) {{
                  var slider = document.getElementById(sliderId);
                  var output = document.getElementById(outputId);
                  output.innerText = slider.value;
                  updateCalculations();
              }}

              function updateCalculations() {{
                  // Calculate adjusted costs for each category
                  var categories = ["Panel Total", "Packaging & Shipping", "Installation", "Equipment", "Travel", "Sales"];
                  var margins = {{}};
                  var adjusted = {{}};
                  var grandTotal = 0;
                  for (var i = 0; i < categories.length; i++) {{
                      var cat = categories[i];
                      var var_id = cat.replace(/\\s+/g, "_");
                      var margin = parseFloat(document.getElementById(var_id + "_margin").value);
                      margins[cat] = margin;
                      var original = baseCosts[cat];
                      var adj = original * (1 + margin/100);
                      adjusted[cat] = adj;
                      grandTotal += adj;
                  }}
                  document.getElementById("dynamic_panel_total").innerText = "$ " + adjusted["Panel Total"].toFixed(2);

                  // Process additional pricing adjustments
                  var product_discount_pct = parseFloat(document.getElementById("product_discount_pct").value);
                  var finders_fee_pct = parseFloat(document.getElementById("finders_fee_pct").value);
                  var sales_commission_pct = parseFloat(document.getElementById("sales_commission_pct").value);
                  var sales_tax_pct = parseFloat(document.getElementById("sales_tax_pct").value);

                  var discount_amount = adjusted["Panel Total"] * (product_discount_pct / 100);
                  var finders_fee_amount = adjusted["Panel Total"] * (finders_fee_pct / 100);
                  var sales_commission_amount = adjusted["Panel Total"] * (sales_commission_pct / 100);
                  var product_total = adjusted["Panel Total"] - discount_amount + finders_fee_amount + sales_commission_amount;
                  var sales_tax_amount = product_total * (sales_tax_pct / 100);
                  var product_total_after_tax = product_total + sales_tax_amount;
                  var other_costs = baseCosts["Packaging & Shipping"] + baseCosts["Installation"] + baseCosts["Equipment"] + baseCosts["Travel"] + baseCosts["Sales"];
                  var total_sell_price = product_total_after_tax + other_costs;
                  var product_plus_installation = product_total + baseCosts["Installation"];
                  var total_cost = 0;
                  for (var key in baseCosts) {{
                      total_cost += baseCosts[key];
                  }}
                  var actual_profit = total_sell_price - total_cost;
                  var profit_margin = (actual_profit / total_sell_price * 100) || 0;

                  // Update dynamic displays for additional adjustments (with spacing)
                  document.getElementById("discount_amount_display").innerText = "$ " + discount_amount.toFixed(2);
                  document.getElementById("finders_fee_amount_display").innerText = "$ " + finders_fee_amount.toFixed(2);
                  document.getElementById("sales_commission_amount_display").innerText = "$ " + sales_commission_amount.toFixed(2);
                  document.getElementById("sales_tax_amount_display").innerText = "$ " + sales_tax_amount.toFixed(2);

                  // Update dynamic displays for final calculations
                  document.getElementById("dynamic_product_total").innerText = "$ " + product_total.toFixed(2);
                  document.getElementById("dynamic_product_total_after_tax").innerText = "$ " + product_total_after_tax.toFixed(2);
                  document.getElementById("dynamic_total_sell_price").innerText = "$ " + total_sell_price.toFixed(2);
                  document.getElementById("dynamic_product_plus_installation").innerText = "$ " + product_plus_installation.toFixed(2);
                  document.getElementById("dynamic_total_cost").innerText = "$ " + total_cost.toFixed(2);
                  document.getElementById("dynamic_actual_profit").innerText = "$ " + actual_profit.toFixed(2);
                  document.getElementById("dynamic_profit_margin").innerText = profit_margin.toFixed(2) + " %";
              }}

              window.onload = updateCalculations;
            </script>
          </head>
          <body>
            <div class="container">
              <h2>Set Margins & Pricing Adjustments</h2>
              <form method="POST">
                <h3>Category Margins</h3>
                {"".join([
                  f"""
                <div class="margin-row">
                  <label for="{cat.replace(' ', '_')}_margin">{cat} Margin (%):</label>
                  <input type="range" style="width:200px;" id="{cat.replace(' ', '_')}_margin" name="{cat}_margin" min="0" max="100" step="1" value="{cp.get(cat + '_margin', 0)}" oninput="updateOutput('{cat.replace(' ', '_')}_margin', '{cat.replace(' ', '_')}_output')">
                  <output id="{cat.replace(' ', '_')}_output" style="margin-right:10px;"></output>
                </div>
                  """ for cat in base_costs
                ])}
                <h3>Additional Pricing Adjustments</h3>
                <div class="margin-row">
                  <label for="product_discount_pct">Product Discount (%):</label>
                  <input type="range" style="width:200px;" id="product_discount_pct" name="product_discount_pct" min="0" max="100" step="1" value="{cp.get('product_discount_pct', 0)}" oninput="updateOutput('product_discount_pct', 'product_discount_output')">
                  <output id="product_discount_output" style="margin-right:10px;"></output>
                  <br><span>Discount Amount: <span id="discount_amount_display">$ 0.00</span></span>
                </div>
                <div class="margin-row">
                  <label for="finders_fee_pct">Finders Fee (%):</label>
                  <input type="range" style="width:200px;" id="finders_fee_pct" name="finders_fee_pct" min="0" max="100" step="1" value="{cp.get('finders_fee_pct', 0)}" oninput="updateOutput('finders_fee_pct', 'finders_fee_output')">
                  <output id="finders_fee_output" style="margin-right:10px;"></output>
                  <br><span>Finders Fee Amount: <span id="finders_fee_amount_display">$ 0.00</span></span>
                </div>
                <div class="margin-row">
                  <label for="sales_commission_pct">Sales Commission (%):</label>
                  <input type="range" style="width:200px;" id="sales_commission_pct" name="sales_commission_pct" min="0" max="100" step="1" value="{cp.get('sales_commission_pct', 0)}" oninput="updateOutput('sales_commission_pct', 'sales_commission_output')">
                  <output id="sales_commission_output" style="margin-right:10px;"></output>
                  <br><span>Sales Commission Amount: <span id="sales_commission_amount_display">$ 0.00</span></span>
                </div>
                <div class="margin-row">
                  <label for="sales_tax_pct">Sales Tax (%):</label>
                  <input type="range" style="width:200px;" id="sales_tax_pct" name="sales_tax_pct" min="0" max="100" step="1" value="{cp.get('sales_tax_pct', 0)}" oninput="updateOutput('sales_tax_pct', 'sales_tax_output')">
                  <output id="sales_tax_output" style="margin-right:10px;"></output>
                  <br><span>Sales Tax Amount: <span id="sales_tax_amount_display">$ 0.00</span></span>
                </div>
                <h3>Final Calculations</h3>
                <table class="summary-table">
                  <tr>
                    <th>Dynamic Panel Total</th>
                    <th>Dynamic Product Total</th>
                    <th>Dynamic Product Total After Tax</th>
                    <th>Dynamic Total Sell Price</th>
                    <th>Dynamic Product + Installation</th>
                    <th>Dynamic Total Cost</th>
                    <th>Dynamic Actual Profit</th>
                    <th>Dynamic Profit Margin (%)</th>
                  </tr>
                  <tr>
                    <td id="dynamic_panel_total">$ 0.00</td>
                    <td id="dynamic_product_total">$ 0.00</td>
                    <td id="dynamic_product_total_after_tax">$ 0.00</td>
                    <td id="dynamic_total_sell_price">$ 0.00</td>
                    <td id="dynamic_product_plus_installation">$ 0.00</td>
                    <td id="dynamic_total_cost">$ 0.00</td>
                    <td id="dynamic_actual_profit">$ 0.00</td>
                    <td id="dynamic_profit_margin">0.00 %</td>
                  </tr>
                </table>
                <div class="btn-group">
                  <div class="btn-left">
                    <button type="button" class="btn" onclick="window.location.href='/other_costs'">Back to Additional Costs</button>
                  </div>
                  <div class="btn-right">
                    <button type="submit" class="btn">Save Margins & Pricing Adjustments</button>
                  </div>
                </div>
              </form>
            </div>
          </body>
        </html>
        """
        return form_html
@app.route('/download_final_summary', methods=['POST'])
def download_final_summary():
    excel_file = create_final_export_excel()  # Use the Excel export function
    cp = get_current_project()  # Retrieve project details from the session
    customer_name = cp.get("customer_name", "customer").replace(" ", "_")
    project_name = cp.get("project_name", "project").replace(" ", "_")
    filename = f"{project_name}_{customer_name}_estimate.xlsx"
    return send_file(
        excel_file,
        download_name=filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def create_final_summary_csv():
    cp = get_current_project()
    output = io.StringIO()
    writer = csv.writer(output)

    # Basic Project Details
    writer.writerow(["Customer Name:", cp.get("customer_name", "N/A")])
    writer.writerow(["Project Name:", cp.get("project_name", "Unnamed Project")])
    writer.writerow(["Estimated by:", cp.get("estimated_by", "N/A")])
    writer.writerow(["Date:", datetime.datetime.now().strftime("%Y-%m-%d")])
    writer.writerow([])

    # Panels Summary
    writer.writerow(["Panels Summary"])
    writer.writerow(["Total Panels", "Total Area (sq ft)", "Total Perimeter (ft)", "Total Vertical (ft)", "Total Horizontal (ft)"])
    swr_panels = cp.get("swr_total_quantity", 0)
    swr_area = cp.get("swr_total_area", 0)
    swr_perimeter = cp.get("swr_total_perimeter", 0)
    swr_vertical = cp.get("swr_total_vertical_ft", 0)
    swr_horizontal = cp.get("swr_total_horizontal_ft", 0)
    igr_panels = cp.get("igr_total_quantity", 0)
    igr_area = cp.get("igr_total_area", 0)
    igr_perimeter = cp.get("igr_total_perimeter", 0)
    igr_vertical = cp.get("igr_total_vertical_ft", 0)
    igr_horizontal = cp.get("igr_total_horizontal_ft", 0)
    total_panels = swr_panels + igr_panels
    total_area = swr_area + igr_area
    total_perimeter = swr_perimeter + igr_perimeter
    total_vertical = swr_vertical + igr_vertical
    total_horizontal = swr_horizontal + igr_horizontal
    writer.writerow([total_panels, total_area, total_perimeter, total_vertical, total_horizontal])
    writer.writerow([])

    # Detailed SWR Materials Itemized Costs
    writer.writerow(["Detailed SWR Materials Itemized Costs"])
    writer.writerow(["Category", "Selected Material", "Unit Cost", "Cost ($)", "Discount/Surcharge", "Final Cost ($)"])
    for item in cp.get("itemized_costs", []):
        writer.writerow([
            item.get("Category", ""),
            item.get("Selected Material", ""),
            item.get("Unit Cost", 0),
            item.get("Cost ($)", 0),
            item.get("Discount/Surcharge", 0),
            item.get("Final Cost", 0)
        ])
    writer.writerow(["Materials Note:", cp.get("swr_note", "")])
    writer.writerow([])

    # Detailed IGR Materials Itemized Costs (with lead times)
    writer.writerow(["Detailed IGR Materials Itemized Costs"])
    writer.writerow(["Category", "Selected Material", "Unit Cost", "Cost ($)", "Discount/Surcharge", "Final Cost ($)", "Min Lead", "Max Lead"])
    for item in cp.get("igr_itemized_costs", []):
        writer.writerow([
            item.get("Category", ""),
            item.get("Selected Material", ""),
            item.get("Unit Cost", 0),
            item.get("Cost ($)", 0),
            item.get("Discount/Surcharge", 0),
            item.get("Final Cost", 0),
            item.get("Min Lead", "N/A"),
            item.get("Max Lead", "N/A")
        ])
    writer.writerow(["IGR Materials Note:", cp.get("igr_note", "")])
    writer.writerow([])

    # Margins and Pricing Adjustments Summary
    writer.writerow(["Margins and Pricing Adjustments"])
    writer.writerow(["Category", "Original Cost ($)", "Margin (%)", "Cost with Margin ($)"])
    for item in cp.get("final_summary", []):
        writer.writerow([
            item.get("Category", ""),
            item.get("Original Cost ($)", 0),
            item.get("Margin (%)", 0),
            item.get("Cost with Margin ($)", 0)
        ])
    writer.writerow(["Grand Total with Margins:", cp.get("grand_total", 0)])
    writer.writerow([])

    # Additional Pricing Adjustments
    writer.writerow(["Additional Pricing Adjustments"])
    writer.writerow(["Adjustment", "Percentage (%)", "Amount ($)"])
    writer.writerow(["Product Discount", cp.get("product_discount_pct", 0), cp.get("product_discount_amount", 0)])
    writer.writerow(["Finders Fee", cp.get("finders_fee_pct", 0), cp.get("finders_fee_amount", 0)])
    writer.writerow(["Sales Commission", cp.get("sales_commission_pct", 0), cp.get("sales_commission_amount", 0)])
    writer.writerow(["Sales Tax", cp.get("sales_tax_pct", 0), cp.get("sales_tax_amount", 0)])
    writer.writerow([])

    # Product Pricing Details
    writer.writerow(["Product Pricing"])
    writer.writerow(["Product Total ($)", "Product Total After Tax ($)", "Product + Installation ($)", "Total Sell Price ($)"])
    writer.writerow([
        cp.get("product_total", 0),
        cp.get("product_total_after_tax", 0),
        cp.get("product_plus_installation", 0),
        cp.get("total_sell_price", 0)
    ])
    writer.writerow([])

    # Final Analysis
    writer.writerow(["Final Analysis"])
    writer.writerow(["Total Sell Price per SF", "Total Sell Price per Panel", "Total Cost ($)", "Actual Profit ($)", "Profit Margin (%)"])
    swr_total_area = cp.get("swr_total_area", 0)
    swr_total_quantity = cp.get("swr_total_quantity", 0)
    total_sell_price = cp.get("total_sell_price", 0)
    total_cost = cp.get("total_cost", 0)
    actual_profit = cp.get("actual_profit", 0)
    profit_margin = cp.get("profit_margin", 0)
    writer.writerow([
        (total_sell_price / swr_total_area) if swr_total_area > 0 else 0,
        (total_sell_price / swr_total_quantity) if swr_total_quantity > 0 else 0,
        total_cost,
        actual_profit,
        profit_margin
    ])
    writer.writerow([])

    # Lead Time Summary (side-by-side display)
    writer.writerow(["Lead Time Summary (weeks)"])
    writer.writerow(["Metric", "Min (weeks)", "Max (weeks)"])
    writer.writerow(["Lead Material", cp.get("min_lead_material", 0), cp.get("max_lead_material", 0)])
    writer.writerow(["Lead Fabrication", cp.get("min_lead_fabrication", 0), cp.get("max_lead_fabrication", 0)])
    writer.writerow(["Total Lead", cp.get("min_total_lead", 0), cp.get("max_total_lead", 0)])
    writer.writerow([])

    # Additional Notes
    writer.writerow(["Notes"])
    writer.writerow(["Fabrication Note:", cp.get("fabrication_note", "")])
    writer.writerow(["Packaging & Shipping Note:", cp.get("packaging_note", "")])
    writer.writerow(["Installation Note:", cp.get("installation_note", "")])
    writer.writerow(["Equipment Note:", cp.get("equipment_note", "")])
    writer.writerow(["Travel Note:", cp.get("travel_note", "")])
    writer.writerow(["Sales Note:", cp.get("sales_note", "")])
    writer.writerow([])

    return output.getvalue()
# ==================================================
# EXCEL EXPORT (Detailed Final Export)
# ==================================================
def create_final_export_excel(margins_dict=None):
    cp = get_current_project()
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book

    # ----------------- Project Export Worksheet -----------------
    ws = workbook.add_worksheet("Project Export")

    # Basic Project Information
    ws.write("A1", "Customer Name:")
    ws.write("B1", cp.get("customer_name", "N/A"))
    ws.write("A2", "Project Name:")
    ws.write("B2", cp.get("project_name", "Unnamed Project"))
    ws.write("A3", "Estimated by:")
    ws.write("B3", cp.get("estimated_by", "N/A"))
    ws.write("A4", "Date:")
    ws.write("B4", datetime.datetime.now().strftime("%Y-%m-%d"))

    # Combined Totals
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

    # ----------------- Detailed SWR Itemized Costs -----------------
    ws.write("G1", "Detailed SWR Itemized Costs")
    headers_swr = ["Category", "Selected Material", "Unit Cost", "Calculation", "MOQ Applied", "Cost ($)",
                   "Final Cost ($)"]
    for col, header in enumerate(headers_swr):
        ws.write(1, col, header)
    row = 2
    for item in cp.get("itemized_costs", []):
        moq_flag = "Yes" if "MOQ Cost" in item.get("Calculation", "") else "No"
        ws.write(row, 0, item.get("Category", ""))
        ws.write(row, 1, item.get("Selected Material", ""))
        ws.write(row, 2, item.get("Unit Cost", 0))
        ws.write(row, 3, item.get("Calculation", ""))
        ws.write(row, 4, moq_flag)
        ws.write(row, 5, item.get("Cost ($)", 0))
        ws.write(row, 6, item.get("Final Cost", 0))
        row += 1
    ws.write(row, 0, "Materials Note:")
    ws.write(row, 1, cp.get("swr_note", ""))
    row += 2

    # ----------------- Detailed IGR Itemized Costs -----------------
    ws.write(row, 0, "Detailed IGR Itemized Costs")
    headers_igr = ["Category", "Selected Material", "Unit Cost", "Calculation", "MOQ Applied", "Cost ($)",
                   "Final Cost ($)"]
    for col, header in enumerate(headers_igr):
        ws.write(row + 1, col, header)
    row = row + 2
    for item in cp.get("igr_itemized_costs", []):
        moq_flag = "Yes" if "MOQ Cost" in item.get("Calculation", "") else "No"
        ws.write(row, 0, item.get("Category", ""))
        ws.write(row, 1, item.get("Selected Material", ""))
        ws.write(row, 2, item.get("Unit Cost", 0))
        ws.write(row, 3, item.get("Calculation", ""))
        ws.write(row, 4, moq_flag)
        ws.write(row, 5, item.get("Cost ($)", 0))
        ws.write(row, 6, item.get("Final Cost", 0))
        row += 1
    ws.write(row, 0, "IGR Materials Note:")
    ws.write(row, 1, cp.get("igr_note", ""))
    row += 2

    # ----------------- Notes from Other Sections -----------------
    ws.write(row, 0, "Fabrication Note:")
    ws.write(row, 1, cp.get("fabrication_note", ""))
    row += 1
    ws.write(row, 0, "Packaging & Shipping Note:")
    ws.write(row, 1, cp.get("packaging_note", ""))
    row += 1
    ws.write(row, 0, "Installation Note:")
    ws.write(row, 1, cp.get("installation_note", ""))
    row += 1
    ws.write(row, 0, "Equipment Note:")
    ws.write(row, 1, cp.get("equipment_note", ""))
    row += 1
    ws.write(row, 0, "Travel Note:")
    ws.write(row, 1, cp.get("travel_note", ""))
    row += 1
    ws.write(row, 0, "Sales Note:")
    ws.write(row, 1, cp.get("sales_note", ""))
    row += 2

    # ----------------- Lead Time Summary (weeks) -----------------
    ws.write(row, 0, "Lead Time Summary (weeks)")
    row += 1
    ws.write(row, 0, "Lead Material")
    ws.write(row, 1, "Min:")
    ws.write(row, 2, cp.get("min_lead_material", 0))
    ws.write(row, 3, "Max:")
    ws.write(row, 4, cp.get("max_lead_material", 0))
    row += 1
    ws.write(row, 0, "Lead Fabrication")
    ws.write(row, 1, "Min:")
    ws.write(row, 2, cp.get("min_lead_fabrication", 0))
    ws.write(row, 3, "Max:")
    ws.write(row, 4, cp.get("max_lead_fabrication", 0))
    row += 1
    ws.write(row, 0, "Total Lead")
    ws.write(row, 1, "Min:")
    ws.write(row, 2, cp.get("min_total_lead", 0))
    ws.write(row, 3, "Max:")
    ws.write(row, 4, cp.get("max_total_lead", 0))
    row += 2

    # ----------------- Margins Summary Section -----------------
    ws.write(row, 0, "Margins Summary")
    row += 1
    margins_headers = ["Category", "Original Cost ($)", "Margin (%)", "Cost with Margin ($)"]
    for col, header in enumerate(margins_headers):
        ws.write(row, col, header)
    row += 1
    if cp.get("final_summary"):
        for summary in cp["final_summary"]:
            ws.write(row, 0, summary.get("Category", ""))
            ws.write(row, 1, summary.get("Original Cost ($)", 0))
            ws.write(row, 2, summary.get("Margin (%)", 0))
            ws.write(row, 3, summary.get("Cost with Margin ($)", 0))
            row += 1
    row += 2

    # ----------------- Product Pricing Details Section -----------------
    ws.write(row, 0, "Product Pricing Details")
    row += 1
    ws.write(row, 0, "Product Price per Unit:")
    ws.write(row, 1, cp.get("product_price_unit", 0))
    row += 1
    ws.write(row, 0, "Product Price per SF:")
    ws.write(row, 1, cp.get("product_price_sf", 0))
    row += 1
    ws.write(row, 0, "Pricing Area (SF):")
    ws.write(row, 1, cp.get("pricing_area", 0))
    row += 2

    # ----------------- Additional Pricing Adjustments Section -----------------
    ws.write(row, 0, "Additional Pricing Adjustments")
    row += 1
    ws.write(row, 0, "Product Discount (%)")
    ws.write(row, 1, cp.get("product_discount_pct", 0))
    ws.write(row, 2, "Product Discount ($)")
    ws.write(row, 3, cp.get("product_discount_dollar", 0))
    row += 1
    ws.write(row, 0, "Finders Fee (%)")
    ws.write(row, 1, cp.get("finders_fee_pct", 0))
    ws.write(row, 2, "Finders Fee ($)")
    ws.write(row, 3, cp.get("finders_fee_dollar", 0))
    row += 1
    ws.write(row, 0, "Sales Commission (%)")
    ws.write(row, 1, cp.get("sales_commission_pct", 0))
    ws.write(row, 2, "Sales Commission ($)")
    ws.write(row, 3, cp.get("sales_commission_dollar", 0))
    row += 1
    ws.write(row, 0, "Sales Tax (%)")
    ws.write(row, 1, cp.get("sales_tax_pct", 0))
    ws.write(row, 2, "Sales Tax ($)")
    ws.write(row, 3, cp.get("sales_tax_dollar", 0))
    row += 2

    # ----------------- Final Pricing Summary Section -----------------
    ws.write(row, 0, "Final Pricing Summary")
    row += 1
    ws.write(row, 0, "Product Total (after adjustments)")
    ws.write(row, 1, cp.get("product_total", 0))
    row += 1
    ws.write(row, 0, "Product + Installation")
    ws.write(row, 1, cp.get("product_plus_installation", 0))
    row += 1
    ws.write(row, 0, "Total Sell Price")
    ws.write(row, 1, cp.get("total_sell_price", 0))
    row += 1
    ws.write(row, 0, "Total Sell Price per SF")
    ws.write(row, 1, cp.get("total_sell_price_per_sf", 0))
    row += 1
    ws.write(row, 0, "Total Sell Price per Panel")
    ws.write(row, 1, cp.get("total_sell_price_per_panel", 0))
    row += 1
    ws.write(row, 0, "Total Cost (pre-margins)")
    ws.write(row, 1, cp.get("total_cost", 0))
    row += 1
    ws.write(row, 0, "Actual Profit ($)")
    ws.write(row, 1, cp.get("actual_profit", 0))
    row += 1
    ws.write(row, 0, "Profit Margin (%)")
    ws.write(row, 1, cp.get("profit_margin", 0))

    # ------------------- Logistics Worksheet -------------------
    import math
    thickness_factor_lookup = {"4": 2.05, "5": 2.56, "6": 3.25, "8": 4.1, "10": 5.12, "12": 6.5}
    swr_extra = {"SWR": 0.25, "SWR-IG": 0.5, "SWR-VIG": 0.35}
    swr_thickness = str(cp.get("glass_thickness", "4"))
    igr_thickness = str(cp.get("igr_glass_thickness", "4"))
    base_factor_swr = thickness_factor_lookup.get(swr_thickness, 2.05)
    base_factor_igr = thickness_factor_lookup.get(igr_thickness, 2.05)
    swr_panel_type = cp.get("swr_system", "SWR")
    extra_weight_swr = swr_extra.get(swr_panel_type, 0.25)
    extra_weight_igr = 0.5  # Fixed for IGR
    factor_swr = base_factor_swr + extra_weight_swr
    factor_igr = base_factor_igr + extra_weight_igr
    crates_swr = math.ceil((cp.get("swr_total_area", 0) * factor_swr) / 2000) if cp.get("swr_total_area", 0) > 0 else 0
    crates_igr = math.ceil((cp.get("igr_total_area", 0) * factor_igr) / 2000) if cp.get("igr_total_area", 0) > 0 else 0
    default_crates = crates_swr + crates_igr

    ws2 = workbook.add_worksheet("Logistics")
    ws2.write("A1", "Logistics Calculations")
    ws2.write("A3", "SWR Total Area (sq ft)")
    ws2.write("B3", cp.get("swr_total_area", 0))
    ws2.write("A4", "IGR Total Area (sq ft)")
    ws2.write("B4", cp.get("igr_total_area", 0))
    ws2.write("A5", "SWR Glass Thickness (mm)")
    ws2.write("B5", cp.get("glass_thickness", "4"))
    ws2.write("A6", "IGR Glass Thickness (mm)")
    ws2.write("B6", cp.get("igr_glass_thickness", "4"))
    ws2.write("A7", "SWR Base Factor")
    ws2.write("B7", base_factor_swr)
    ws2.write("A8", "SWR Extra Weight")
    ws2.write("B8", extra_weight_swr)
    ws2.write("A9", "SWR Total Factor")
    ws2.write("B9", factor_swr)
    ws2.write("A10", "Crates for SWR")
    ws2.write("B10", crates_swr)
    ws2.write("A11", "IGR Base Factor")
    ws2.write("B11", base_factor_igr)
    ws2.write("A12", "IGR Extra Weight")
    ws2.write("B12", extra_weight_igr)
    ws2.write("A13", "IGR Total Factor")
    ws2.write("B13", factor_igr)
    ws2.write("A14", "Crates for IGR")
    ws2.write("B14", crates_igr)
    ws2.write("A15", "Total Crates (Logistics)")
    ws2.write("B15", default_crates)

    # ------------------- Order Worksheet -------------------
    # This sheet calculates the order quantities and prices based on inventory shortfall.
    # For each SWR category, we determine the required amount, get the available stock,
    # compute delta, then order quantity = ceil(delta / yield), and order price = order_qty * yield_cost.
    ws_order = workbook.add_worksheet("Order")
    ws_order.write("A1", "Category")
    ws_order.write("B1", "Required Amount")
    ws_order.write("C1", "Available")
    ws_order.write("D1", "Delta")
    ws_order.write("E1", "Yield Factor")
    ws_order.write("F1", "Order Quantity")
    ws_order.write("G1", "Yield Cost")
    ws_order.write("H1", "Order Price")

    # Helper function: Calculate the required amount for each category
    def calc_required(category):
        if category == "Glass (Cat 15)":
            return cp.get("swr_total_area", 0)
        elif category == "Extrusions (Cat 1)":
            return cp.get("swr_total_perimeter", 0)
        elif category == "Retainer (Cat 17)":
            option = cp.get("retainer_option", "")
            if option == "head_retainer":
                return 0.5 * cp.get("swr_total_horizontal_ft", 0)
            elif option == "head_and_sill":
                return cp.get("swr_total_horizontal_ft", 0)
            else:
                return 0
        elif category == "Glazing Spline (Cat 2)":
            return cp.get("swr_total_perimeter", 0)
        elif category == "Gaskets (Cat 3)":
            return cp.get("swr_total_vertical_ft", 0)
        elif category == "Corner Keys (Cat 4)":
            return cp.get("swr_total_quantity", 0) * 4
        elif category == "Dual Lock (Cat 5)":
            return cp.get("swr_total_quantity", 0)
        elif category == "Foam Baffle Top/Head (Cat 6)":
            return 0.5 * cp.get("swr_total_horizontal_ft", 0)
        elif category == "Foam Baffle Bottom/Sill (Cat 6)":
            return 0.5 * cp.get("swr_total_horizontal_ft", 0)
        elif category == "Glass Protection (Cat 7)":
            return cp.get("swr_total_area", 0)
        elif category == "Tape (Cat 10)":
            option = cp.get("retainer_attachment_option", "")
            if option == "head_retainer":
                return cp.get("swr_total_horizontal_ft", 0) / 2
            elif option == "head_sill":
                return cp.get("swr_total_horizontal_ft", 0)
            else:
                return 0
        elif category == "Screws (Cat 18)":
            option = cp.get("screws_option", "")
            if option == "head_retainer":
                return 0.5 * cp.get("swr_total_horizontal_ft", 0) * 4
            elif option == "head_and_sill":
                return cp.get("swr_total_horizontal_ft", 0) * 4
            else:
                return 0
        elif category == "Setting Block (Cat 16)":
            return cp.get("swr_total_quantity", 0) * 2
        else:
            return 0

    # Mapping of category to yield factors stored in cp.
    yield_mapping = {
        "Glass (Cat 15)": cp.get("yield_cat15", 1),
        "Extrusions (Cat 1)": cp.get("yield_aluminum", 1),
        "Retainer (Cat 17)": cp.get("yield_aluminum", 1),
        "Glazing Spline (Cat 2)": cp.get("yield_cat2", 1),
        "Gaskets (Cat 3)": cp.get("yield_cat3", 1),
        "Corner Keys (Cat 4)": cp.get("yield_cat4", 1),
        "Dual Lock (Cat 5)": cp.get("yield_cat5", 1),
        "Foam Baffle Top/Head (Cat 6)": cp.get("yield_cat6", 1),
        "Foam Baffle Bottom/Sill (Cat 6)": cp.get("yield_cat6", 1),
        "Glass Protection (Cat 7)": cp.get("yield_cat7", 1),
        "Tape (Cat 10)": cp.get("yield_cat10", 1),
        "Screws (Cat 18)": 1,
        "Setting Block (Cat 16)": cp.get("yield_cat16", 1)
    }

    # Reconstruct a mapping from category to material object for SWR materials
    material_order_map = {}
    if cp.get("material_glass"):
        material_order_map["Glass (Cat 15)"] = Material.query.get(cp.get("material_glass"))
    if cp.get("material_aluminum"):
        material_order_map["Extrusions (Cat 1)"] = Material.query.get(cp.get("material_aluminum"))
    if cp.get("material_retainer"):
        material_order_map["Retainer (Cat 17)"] = Material.query.get(cp.get("material_retainer"))
    if cp.get("material_glazing"):
        material_order_map["Glazing Spline (Cat 2)"] = Material.query.get(cp.get("material_glazing"))
    if cp.get("material_gaskets"):
        material_order_map["Gaskets (Cat 3)"] = Material.query.get(cp.get("material_gaskets"))
    if cp.get("material_corner_keys"):
        material_order_map["Corner Keys (Cat 4)"] = Material.query.get(cp.get("material_corner_keys"))
    if cp.get("material_dual_lock"):
        material_order_map["Dual Lock (Cat 5)"] = Material.query.get(cp.get("material_dual_lock"))
    if cp.get("material_foam_baffle"):
        material_order_map["Foam Baffle Top/Head (Cat 6)"] = Material.query.get(cp.get("material_foam_baffle"))
    if cp.get("material_foam_baffle_bottom"):
        material_order_map["Foam Baffle Bottom/Sill (Cat 6)"] = Material.query.get(
            cp.get("material_foam_baffle_bottom"))
    if cp.get("material_glass_protection"):
        material_order_map["Glass Protection (Cat 7)"] = Material.query.get(cp.get("material_glass_protection"))
    if cp.get("material_tape"):
        material_order_map["Tape (Cat 10)"] = Material.query.get(cp.get("material_tape"))
    if cp.get("material_screws"):
        material_order_map["Screws (Cat 18)"] = Material.query.get(cp.get("material_screws"))
    if cp.get("material_setting_block"):
        material_order_map["Setting Block (Cat 16)"] = Material.query.get(cp.get("material_setting_block"))

    # Now prepare the order items
    order_items = []
    import math
    categories = ["Glass (Cat 15)", "Extrusions (Cat 1)", "Retainer (Cat 17)", "Glazing Spline (Cat 2)",
                  "Gaskets (Cat 3)", "Corner Keys (Cat 4)", "Dual Lock (Cat 5)", "Foam Baffle Top/Head (Cat 6)",
                  "Foam Baffle Bottom/Sill (Cat 6)", "Glass Protection (Cat 7)", "Tape (Cat 10)", "Screws (Cat 18)",
                  "Setting Block (Cat 16)"]
    for cat in categories:
        required_amt = calc_required(cat)
        mat_obj = material_order_map.get(cat)
        available = mat_obj.quantity if (mat_obj and mat_obj.quantity is not None) else 0
        delta = required_amt - available if required_amt > available else 0
        yld = yield_mapping.get(cat, 1)
        order_qty = math.ceil(delta / yld) if delta > 0 and yld > 0 else 0
        yield_cost = mat_obj.yield_cost if mat_obj and mat_obj.yield_cost is not None else 0
        order_price = order_qty * yield_cost
        order_items.append({
            "Category": cat,
            "Required": required_amt,
            "Available": available,
            "Delta": delta,
            "Yield": yld,
            "Order Quantity": order_qty,
            "Yield Cost": yield_cost,
            "Order Price": order_price
        })
    total_order_price = sum(item["Order Price"] for item in order_items)

    # Write the "Order" worksheet
    ws_order.write("A2", "Category")
    ws_order.write("B2", "Required")
    ws_order.write("C2", "Available")
    ws_order.write("D2", "Delta")
    ws_order.write("E2", "Yield")
    ws_order.write("F2", "Order Quantity")
    ws_order.write("G2", "Yield Cost")
    ws_order.write("H2", "Order Price")
    row_order = 3
    for item in order_items:
        ws_order.write(row_order, 0, item["Category"])
        ws_order.write(row_order, 1, item["Required"])
        ws_order.write(row_order, 2, item["Available"])
        ws_order.write(row_order, 3, item["Delta"])
        ws_order.write(row_order, 4, item["Yield"])
        ws_order.write(row_order, 5, item["Order Quantity"])
        ws_order.write(row_order, 6, item["Yield Cost"])
        ws_order.write(row_order, 7, item["Order Price"])
        row_order += 1
    ws_order.write(row_order, 0, "Total Order Price:")
    ws_order.write(row_order, 1, total_order_price)

    writer.close()
    output.seek(0)
    return output
@app.route('/new_project')
def new_project():
    session.clear()  # This clears all project data from the session
    return redirect(url_for('index'))

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)