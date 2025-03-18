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
    # Removed the 'status' column since it no longer exists.
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

# Helper function: Recursively convert NumPy int64 (and similar types) to native Python int
def make_serializable(obj):
    if isinstance(obj, dict):
        return {k: make_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [make_serializable(x) for x in obj]
    elif isinstance(obj, np.int64):
        return int(obj)
    else:
        return obj

# Helper functions for session management
def get_current_project():
    cp = session.get("current_project")
    if cp is None:
        cp = {}
        session["current_project"] = cp
    return cp

def save_current_project(cp):
    session["current_project"] = make_serializable(cp)

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
    a { color: #00a0a0; text-decoration: none; font-weight: 600; }
    a:hover { text-decoration: underline; }
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
"""

# =========================
# INDEX PAGE
# =========================
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
                      <option value="SWR" {"selected" if cp.get('swr_system','')=="SWR" else ""}>SWR</option>
                      <option value="SWR-IG" {"selected" if cp.get('swr_system','')=="SWR-IG" else ""}>SWR-IG</option>
                      <option value="SWR-VIG" {"selected" if cp.get('swr_system','')=="SWR-VIG" else ""}>SWR-VIG</option>
                  </select>
               </div>
               <div>
                  <label for="swr_mount">SWR Mount Type:</label>
                  <select name="swr_mount" id="swr_mount" style="width:200px;" required>
                      <option value="Inset-mount" {"selected" if cp.get('swr_mount','')=="Inset-mount" else ""}>Inset-mount</option>
                      <option value="Overlap-mount" {"selected" if cp.get('swr_mount','')=="Overlap-mount" else ""}>Overlap-mount</option>
                  </select>
               </div>
               <div>
                  <label for="igr_type">IGR Type:</label>
                  <select name="igr_type" id="igr_type" style="width:200px;" required>
                     <option value="Wet Seal IGR" {"selected" if cp.get('igr_type','')=="Wet Seal IGR" else ""}>Wet Seal IGR</option>
                     <option value="Dry Seal IGR" {"selected" if cp.get('igr_type','')=="Dry Seal IGR" else ""}>Dry Seal IGR</option>
                  </select>
               </div>
               <div>
                  <label for="igr_location">IGR Location:</label>
                  <select name="igr_location" id="igr_location" style="width:200px;" required>
                     <option value="Interior" {"selected" if cp.get('igr_location','')=="Interior" else ""}>Interior</option>
                     <option value="Exterior" {"selected" if cp.get('igr_location','')=="Exterior" else ""}>Exterior</option>
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
    igr_area, igr_perimeter, igr_vertical, igr_horizontal, igr_quantity = compute_totals(df_igr) if not df_igr.empty else (0, 0, 0, 0, 0)

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

    next_page = '/materials'
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
              <button type="button" class="btn" onclick="window.location.href='/'">Start New Project</button>
              <button type="button" class="btn" onclick="window.location.href='{next_page}'">Next: SWR Materials</button>
            </div>
         </div>
      </body>
    </html>
    """
    return summary_html

# --- Helper to generate options with previous selection ---
def generate_options(materials_list, selected_value=None):
    options = ""
    for m in materials_list:
        if str(m.id) == str(selected_value):
            options += f'<option value="{m.id}" selected>{m.nickname} - ${m.yield_cost:.2f}</option>'
        else:
            options += f'<option value="{m.id}">{m.nickname} - ${m.yield_cost:.2f}</option>'
    return options

# =========================
# SWR MATERIALS PAGE (Always shown)
# =========================
@app.route('/materials', methods=['GET', 'POST'])
def materials():
    cp = get_current_project()
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
    except Exception as e:
        return f"<h2 style='color: red;'>Error fetching materials: {e}</h2>"

    if request.method == 'POST':
        try:
            yield_cat15 = float(request.form.get('yield_cat15', 0.97))
        except:
            yield_cat15 = 0.97
        try:
            yield_aluminum = float(request.form.get('yield_aluminum', 0.75))
        except:
            yield_aluminum = 0.75
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

        cp['yield_cat15'] = yield_cat15
        cp['yield_aluminum'] = yield_aluminum
        cp['yield_cat2'] = yield_cat2
        cp['yield_cat3'] = yield_cat3
        cp['yield_cat4'] = yield_cat4
        cp['yield_cat5'] = yield_cat5
        cp['yield_cat6'] = yield_cat6
        cp['yield_cat7'] = yield_cat7
        cp['yield_cat10'] = yield_cat10
        cp['yield_cat17'] = yield_cat17

        selected_glass = request.form.get('material_glass')
        selected_aluminum = request.form.get('material_aluminum')
        retainer_option = request.form.get('retainer_option')
        selected_retainer = request.form.get('material_retainer')
        selected_glazing = request.form.get('material_glazing')
        selected_gaskets = request.form.get('material_gaskets')
        jamb_plate = request.form.get('jamb_plate')
        jamb_plate_screws = request.form.get('jamb_plate_screws')
        cp['jamb_plate'] = jamb_plate
        cp['jamb_plate_screws'] = jamb_plate_screws
        selected_corner_keys = request.form.get('material_corner_keys')
        selected_dual_lock = request.form.get('material_dual_lock')
        selected_foam_baffle_top = request.form.get('material_foam_baffle')
        selected_foam_baffle_bottom = request.form.get('material_foam_baffle_bottom')
        glass_protection_side = request.form.get('glass_protection_side')
        selected_glass_protection = request.form.get('material_glass_protection')
        retainer_attachment_option = request.form.get('retainer_attachment_option')
        selected_tape = request.form.get('material_tape')
        selected_head_retainers = request.form.get('material_head_retainers')
        screws_option = request.form.get('screws_option')
        selected_screws = request.form.get('material_screws')
        # Retrieve the Materials Note from the form
        cp["swr_note"] = request.form.get("swr_note", "")
        save_current_project(cp)

        mat_glass = Material.query.get(selected_glass) if selected_glass else None
        mat_aluminum = Material.query.get(selected_aluminum) if selected_aluminum else None
        mat_retainer = Material.query.get(selected_retainer) if selected_retainer else None
        mat_glazing = Material.query.get(selected_glazing) if selected_glazing else None
        mat_gaskets = Material.query.get(selected_gaskets) if selected_gaskets else None
        mat_corner_keys = Material.query.get(selected_corner_keys) if selected_corner_keys else None
        mat_dual_lock = Material.query.get(selected_dual_lock) if selected_dual_lock else None
        mat_foam_baffle_top = Material.query.get(selected_foam_baffle_top) if selected_foam_baffle_top else None
        mat_foam_baffle_bottom = Material.query.get(selected_foam_baffle_bottom) if selected_foam_baffle_bottom else None
        mat_glass_protection = Material.query.get(selected_glass_protection) if selected_glass_protection else None
        mat_tape = Material.query.get(selected_tape) if selected_tape else None
        mat_head_retainers = Material.query.get(selected_head_retainers) if selected_head_retainers else None
        mat_screws = Material.query.get(selected_screws) if selected_screws else None

        total_area = cp.get('swr_total_area', 0)
        total_perimeter = cp.get('swr_total_perimeter', 0)
        total_vertical = cp.get('swr_total_vertical_ft', 0)
        total_horizontal = cp.get('swr_total_horizontal_ft', 0)
        total_quantity = cp.get('swr_total_quantity', 0)

        cost_glass = (total_area * mat_glass.yield_cost) / yield_cat15 if mat_glass else 0
        cost_aluminum = (total_perimeter * mat_aluminum.yield_cost) / yield_aluminum if mat_aluminum else 0
        if retainer_option == "head_retainer":
            cost_retainer = (0.5 * total_horizontal * mat_retainer.yield_cost) / yield_aluminum if mat_retainer else 0
        elif retainer_option == "head_and_sill":
            cost_retainer = (total_horizontal * mat_retainer.yield_cost * yield_aluminum) if mat_retainer else 0
        else:
            cost_retainer = 0
        cost_glazing = (total_perimeter * mat_glazing.yield_cost) / yield_cat2 if mat_glazing else 0
        cost_gaskets = (total_vertical * mat_gaskets.yield_cost) / yield_cat3 if mat_gaskets else 0
        cost_corner_keys = (total_quantity * 4 * mat_corner_keys.yield_cost) / yield_cat4 if mat_corner_keys else 0
        cost_dual_lock = (total_quantity * mat_dual_lock.yield_cost) / yield_cat5 if mat_dual_lock else 0
        cost_foam_baffle_top = (0.5 * total_horizontal * mat_foam_baffle_top.yield_cost) / yield_cat6 if mat_foam_baffle_top else 0
        cost_foam_baffle_bottom = (0.5 * total_horizontal * mat_foam_baffle_bottom.yield_cost) / yield_cat6 if mat_foam_baffle_bottom else 0
        if glass_protection_side == "one":
            cost_glass_protection = (total_area * mat_glass_protection.yield_cost) / yield_cat7 if mat_glass_protection else 0
        elif glass_protection_side == "double":
            cost_glass_protection = (total_area * mat_glass_protection.yield_cost * 2) / yield_cat7 if mat_glass_protection else 0
        elif glass_protection_side == "none":
            cost_glass_protection = 0
        else:
            cost_glass_protection = 0
        if mat_tape:
            if retainer_attachment_option == 'head_retainer':
                cost_tape = ((total_horizontal / 2) * mat_tape.yield_cost) / yield_cat10
            elif retainer_attachment_option == 'head_sill':
                cost_tape = (total_horizontal * mat_tape.yield_cost) / yield_cat10
            elif retainer_attachment_option == 'no_tape':
                cost_tape = 0
            else:
                cost_tape = 0
        else:
            cost_tape = 0
        cost_head_retainers = ((total_horizontal / 2) * mat_head_retainers.yield_cost) / yield_cat17 if mat_head_retainers else 0

        if screws_option == "head_retainer":
            cost_screws = (0.5 * total_horizontal * 4 * (mat_screws.yield_cost if mat_screws else 0))
        elif screws_option == "head_and_sill":
            cost_screws = (total_horizontal * 4 * (mat_screws.yield_cost if mat_screws else 0))
        else:
            cost_screws = 0

        total_material_cost = (cost_glass + cost_aluminum + cost_retainer + cost_glazing + cost_gaskets +
                               cost_corner_keys + cost_dual_lock + cost_foam_baffle_top + cost_foam_baffle_bottom +
                               cost_glass_protection + cost_tape + cost_head_retainers + cost_screws)
        cp['material_total_cost'] = total_material_cost

        material_map = {
            "Glass (Cat 15)": mat_glass,
            "Extrusions (Cat 1)": mat_aluminum,
            "Retainer (Cat 17)": mat_retainer,
            "Glazing Spline (Cat 2)": mat_glazing,
            "Gaskets (Cat 3)": mat_gaskets,
            "Corner Keys (Cat 4)": mat_corner_keys,
            "Dual Lock (Cat 5)": mat_dual_lock,
            "Foam Baffle Top/Head (Cat 6)": mat_foam_baffle_top,
            "Foam Baffle Bottom/Sill (Cat 6)": mat_foam_baffle_bottom,
            "Glass Protection (Cat 7)": mat_glass_protection,
            "Tape (Cat 10)": mat_tape,
            "Screws (Cat 18)": mat_screws
        }

        def get_required_quantity(category):
            if category == "Glass (Cat 15)":
                return total_area
            elif category == "Extrusions (Cat 1)":
                return total_perimeter
            elif category == "Retainer (Cat 17)":
                if retainer_option == "head_retainer":
                    return 0.5 * total_horizontal
                elif retainer_option == "head_and_sill":
                    return total_horizontal
                else:
                    return 0
            elif category == "Glazing Spline (Cat 2)":
                return total_perimeter
            elif category == "Gaskets (Cat 3)":
                return total_vertical
            elif category == "Corner Keys (Cat 4)":
                return total_quantity * 4
            elif category == "Dual Lock (Cat 5)":
                return total_quantity
            elif category == "Foam Baffle Top/Head (Cat 6)":
                return 0.5 * total_horizontal
            elif category == "Foam Baffle Bottom/Sill (Cat 6)":
                return 0.5 * total_horizontal
            elif category == "Glass Protection (Cat 7)":
                if glass_protection_side == "double":
                    return total_area * 2
                else:
                    return total_area
            elif category == "Tape (Cat 10)":
                if retainer_attachment_option == 'head_retainer':
                    return total_horizontal / 2
                elif retainer_attachment_option == 'head_sill':
                    return total_horizontal
                else:
                    return 0
            elif category == "Screws (Cat 18)":
                if screws_option == "head_retainer":
                    return 0.5 * total_horizontal * 4
                elif screws_option == "head_and_sill":
                    return total_horizontal * 4
                else:
                    return 0
            else:
                return 0

        materials_list = [
            {
                "Category": "Glass (Cat 15)",
                "Selected Material": mat_glass.nickname if mat_glass else "N/A",
                "Unit Cost": mat_glass.yield_cost if mat_glass else 0,
                "Calculation": f"Total Area {total_area:.2f} × Yield Cost / {yield_cat15}",
                "Cost ($)": cost_glass
            },
            {
                "Category": "Extrusions (Cat 1)",
                "Selected Material": mat_aluminum.nickname if mat_aluminum else "N/A",
                "Unit Cost": mat_aluminum.yield_cost if mat_aluminum else 0,
                "Calculation": f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {yield_aluminum}",
                "Cost ($)": cost_aluminum
            },
            {
                "Category": "Retainer (Cat 17)",
                "Selected Material": (mat_retainer.nickname if mat_retainer else "N/A") if retainer_option != "no_retainer" else "N/A",
                "Unit Cost": (mat_retainer.yield_cost if mat_retainer else 0) if retainer_option != "no_retainer" else 0,
                "Calculation": (f"Head Retainer: 0.5 × Total Horizontal {total_horizontal:.2f} × Yield Cost / {yield_aluminum}"
                                if retainer_option == "head_retainer"
                                else (f"Head + Sill Retainer: Total Horizontal {total_horizontal:.2f} × Yield Cost × {yield_aluminum}"
                                      if retainer_option == "head_and_sill"
                                      else "No Retainer")),
                "Cost ($)": cost_retainer
            },
            {
                "Category": "Glazing Spline (Cat 2)",
                "Selected Material": mat_glazing.nickname if mat_glazing else "N/A",
                "Unit Cost": mat_glazing.yield_cost if mat_glazing else 0,
                "Calculation": f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {yield_cat2}",
                "Cost ($)": cost_glazing
            },
            {
                "Category": "Gaskets (Cat 3)",
                "Selected Material": mat_gaskets.nickname if mat_gaskets else "N/A",
                "Unit Cost": mat_gaskets.yield_cost if mat_gaskets else 0,
                "Calculation": f"Total Vertical {total_vertical:.2f} × Yield Cost / {yield_cat3}",
                "Cost ($)": cost_gaskets
            },
            {
                "Category": "Corner Keys (Cat 4)",
                "Selected Material": mat_corner_keys.nickname if mat_corner_keys else "N/A",
                "Unit Cost": mat_corner_keys.yield_cost if mat_corner_keys else 0,
                "Calculation": f"Total Quantity {total_quantity:.2f} × 4 × Yield Cost / {yield_cat4}",
                "Cost ($)": cost_corner_keys
            },
            {
                "Category": "Dual Lock (Cat 5)",
                "Selected Material": mat_dual_lock.nickname if mat_dual_lock else "N/A",
                "Unit Cost": mat_dual_lock.yield_cost if mat_dual_lock else 0,
                "Calculation": f"Total Quantity {total_quantity:.2f} × Yield Cost / {yield_cat5}",
                "Cost ($)": cost_dual_lock
            },
            {
                "Category": "Foam Baffle Top/Head (Cat 6)",
                "Selected Material": mat_foam_baffle_top.nickname if mat_foam_baffle_top else "N/A",
                "Unit Cost": mat_foam_baffle_top.yield_cost if mat_foam_baffle_top else 0,
                "Calculation": f"0.5 × Total Horizontal {total_horizontal:.2f} × Yield Cost / {yield_cat6}",
                "Cost ($)": cost_foam_baffle_top
            },
            {
                "Category": "Foam Baffle Bottom/Sill (Cat 6)",
                "Selected Material": mat_foam_baffle_bottom.nickname if mat_foam_baffle_bottom else "N/A",
                "Unit Cost": mat_foam_baffle_bottom.yield_cost if mat_foam_baffle_bottom else 0,
                "Calculation": f"0.5 × Total Horizontal {total_horizontal:.2f} × Yield Cost / {yield_cat6}",
                "Cost ($)": cost_foam_baffle_bottom
            },
            {
                "Category": "Glass Protection (Cat 7)",
                "Selected Material": mat_glass_protection.nickname if mat_glass_protection else "N/A",
                "Unit Cost": mat_glass_protection.yield_cost if mat_glass_protection else 0,
                "Calculation": (f"Total Area {total_area:.2f} × Yield Cost / {yield_cat7}"
                                if glass_protection_side == "one"
                                else (f"Total Area {total_area:.2f} × Yield Cost × 2 / {yield_cat7}"
                                      if glass_protection_side == "double"
                                      else "No Film")),
                "Cost ($)": cost_glass_protection
            },
            {
                "Category": "Tape (Cat 10)",
                "Selected Material": mat_tape.nickname if mat_tape else "N/A",
                "Unit Cost": mat_tape.yield_cost if mat_tape else 0,
                "Calculation": "Retainer Attachment Option: " + (
                    "(Head Retainer - Half Horizontal)" if retainer_attachment_option == "head_retainer"
                    else ("(Head+Sill - Full Horizontal)" if retainer_attachment_option == "head_sill" else "No Tape")),
                "Cost ($)": cost_tape
            },
            {
                "Category": "Screws (Cat 18)",
                "Selected Material": mat_screws.nickname if mat_screws else "N/A",
                "Unit Cost": mat_screws.yield_cost if mat_screws else 0,
                "Calculation": (f"Head Retainer: 0.5 × Total Horizontal {total_horizontal:.2f} × 4 × Yield Cost"
                                if screws_option == "head_retainer"
                                else (f"Head + Sill: Total Horizontal {total_horizontal:.2f} × 4 × Yield Cost"
                                      if screws_option == "head_and_sill"
                                      else "No Screws")),
                "Cost ($)": cost_screws
            },
        ]

        # Process discount/surcharge and compute Final Cost for each SWR item
        for item in materials_list:
            discount_key = "discount_" + "".join(c if c.isalnum() else "_" for c in item["Category"])
            discount_str = (request.form.get(discount_key) or "").strip()
            if discount_str != "":
                try:
                    discount_value = float(discount_str)
                except:
                    discount_value = 0
            else:
                discount_value = cp.get(discount_key, 0)
            cp[discount_key] = discount_value
            item["Discount/Surcharge"] = discount_value
            item["Final Cost"] = item["Cost ($)"] + discount_value

        total_final_cost = sum(item["Final Cost"] for item in materials_list)

        # Build interactive table HTML for SWR materials
        table_html = "<form method='POST'>"
        headers = ["Category", "Selected Material", "Unit Cost", "Calculation", "Cost ($)", "$ per SF", "% Total Cost", "Stock Level", "Min Lead", "Max Lead", "Discount/Surcharge", "Final Cost"]
        table_html += "<table class='summary-table'>"
        table_html += "<tr>"
        for header in headers:
            table_html += f"<th>{header}</th>"
        table_html += "</tr>"
        for item in materials_list:
            discount_key = "discount_" + "".join(c if c.isalnum() else "_" for c in item["Category"])
            table_html += "<tr>"
            table_html += f"<td>{item['Category']}</td>"
            table_html += f"<td>{item['Selected Material']}</td>"
            table_html += f"<td>{item['Unit Cost']:.2f}</td>"
            table_html += f"<td>{item['Calculation']}</td>"
            table_html += f"<td class='cost'>{item['Cost ($)']:.2f}</td>"
            table_html += f"<td>{(item['Final Cost'] / total_area if total_area > 0 else 0):.2f}</td>"
            table_html += f"<td>{(item['Final Cost'] / total_final_cost * 100 if total_final_cost > 0 else 0):.2f}</td>"
            table_html += f"<td>N/A</td>"  # Stock Level
            table_html += f"<td>N/A</td>"  # Min Lead
            table_html += f"<td>N/A</td>"  # Max Lead
            table_html += f"<td><input type='number' step='1' name='{discount_key}' value='{item['Discount/Surcharge']}' oninput='updateFinalCost(this)' /></td>"
            table_html += f"<td class='final-cost'>{item['Final Cost']:.2f}</td>"
            table_html += "</tr>"
        table_html += "</table>"
        # Place the note field outside the table:
        table_html += "<div style='margin-top:10px;'><label for='swr_note'>Materials Note:</label><textarea id='swr_note' name='swr_note' rows='3' style='width:100%;'>" + cp.get("swr_note", "") + "</textarea></div>"
        table_html += """
        <script>
        function updateFinalCost(input) {
            var row = input.closest("tr");
            var costCell = row.querySelector(".cost");
            var finalCostCell = row.querySelector(".final-cost");
            var discountValue = parseFloat(input.value) || 0;
            var costValue = parseFloat(costCell.innerText) || 0;
            var finalCost = costValue + discountValue;
            finalCostCell.innerText = finalCost.toFixed(2);
        }
        </script>
        """
        table_html += "</form>"

        cp["itemized_costs"] = materials_list
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
               <div class="btn-group">
                 <button type="button" class="btn" onclick="window.location.href='/materials'">Back: Edit Materials</button>
                 <button type="button" class="btn" onclick="window.location.href='/igr_materials'">Next: IGR Materials</button>
               </div>
             </div>
           </body>
         </html>
        """

    def generate_options(materials_list, selected_value=None):
        options = ""
        for m in materials_list:
            if str(m.id) == str(selected_value):
                options += f'<option value="{m.id}" selected>{m.nickname} - ${m.yield_cost:.2f}</option>'
            else:
                options += f'<option value="{m.id}">{m.nickname} - ${m.yield_cost:.2f}</option>'
        return options

    form_html = f"""
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
                     <option value="head_retainer" {"selected" if cp.get('retainer_option','')=="head_retainer" else ""}>Head Retainer</option>
                     <option value="head_and_sill" {"selected" if cp.get('retainer_option','')=="head_and_sill" else ""}>Head + Sill Retainer</option>
                     <option value="no_retainer" {"selected" if cp.get('retainer_option','')=="no_retainer" else ""}>No Retainer</option>
                  </select>
               </div>
               <div>
                  <label for="material_retainer">Select Retainer Material:</label>
                  <select name="material_retainer" id="material_retainer" required>
                     {generate_options(materials_head_retainers, cp.get("material_retainer"))}
                  </select>
               </div>
               <div>
                  <label for="screws_option">Retainer Screws Option:</label>
                  <select name="screws_option" id="screws_option" required>
                     <option value="none" {"selected" if cp.get('screws_option','')=="none" else ""}>None</option>
                     <option value="head_retainer" {"selected" if cp.get('screws_option','')=="head_retainer" else ""}>Head Retainer</option>
                     <option value="head_and_sill" {"selected" if cp.get('screws_option','')=="head_and_sill" else ""}>Head + Sill</option>
                  </select>
               </div>
               <div>
                  <label for="material_screws">Select Retainer Screws:</label>
                  <select name="material_screws" id="material_screws" required>
                     {generate_options(materials_screws, cp.get("material_screws"))}
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
                     <option value="Yes" {"selected" if cp.get('jamb_plate','')=="Yes" else ""}>Yes</option>
                     <option value="No" {"selected" if cp.get('jamb_plate','')=="No" else ""}>No</option>
                  </select>
               </div>
               <div>
                  <label for="jamb_plate_screws">Jamb Plate Screws:</label>
                  <select name="jamb_plate_screws" id="jamb_plate_screws" required>
                     <option value="Yes" {"selected" if cp.get('jamb_plate_screws','')=="Yes" else ""}>Yes</option>
                     <option value="No" {"selected" if cp.get('jamb_plate_screws','')=="No" else ""}>No</option>
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
                     <option value="one" {"selected" if cp.get('glass_protection_side','')=="one" else ""}>One Sided</option>
                     <option value="double" {"selected" if cp.get('glass_protection_side','')=="double" else ""}>Double Sided</option>
                     <option value="none" {"selected" if cp.get('glass_protection_side','')=="none" else ""}>No Film</option>
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
                     <option value="head_retainer" {"selected" if cp.get('retainer_attachment_option','')=="head_retainer" else ""}>Head Retainer (Half Horizontal)</option>
                     <option value="head_sill" {"selected" if cp.get('retainer_attachment_option','')=="head_sill" else ""}>Head+Sill (Full Horizontal)</option>
                     <option value="no_tape" {"selected" if cp.get('retainer_attachment_option','')=="no_tape" else ""}>No Tape</option>
                  </select>
               </div>
               <div>
                  <label for="material_tape">Select Tape Material:</label>
                  <select name="material_tape" id="material_tape" required>
                     {generate_options(materials_tape, cp.get("material_tape"))}
                  </select>
               </div>
               <div>
                  <label for="yield_cat17">Head Retainers (Cat 17) Yield:</label>
                  <input type="number" step="0.01" id="yield_cat17" name="yield_cat17" value="{cp.get('yield_cat17', '0.91')}" required>
               </div>
               <div>
                  <label for="material_head_retainers">Select Head Retainers:</label>
                  <select name="material_head_retainers" id="material_head_retainers" required>
                     {generate_options(materials_head_retainers, cp.get("material_head_retainers"))}
                  </select>
               </div>
               <div class="btn-group">
                  <button type="button" class="btn" onclick="window.location.href='/summary'">Back: Edit Summary</button>
                  <button type="submit" class="btn">Calculate SWR Material Costs</button>
               </div>
            </form>
         </div>
      </body>
    </html>
    """
    return form_html

# =========================
# IGR MATERIALS PAGE (Always shown)
# =========================
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
            yield_igr_glass = float(request.form.get('yield_igr_glass', 0.97))
        except:
            yield_igr_glass = 0.97
        try:
            yield_igr_extrusions = float(request.form.get('yield_igr_extrusions', 0.75))
        except:
            yield_igr_extrusions = 0.75
        try:
            yield_igr_gaskets = float(request.form.get('yield_igr_gaskets', 0.91))
        except:
            yield_igr_gaskets = 0.91
        try:
            yield_igr_glass_protection = float(request.form.get('yield_igr_glass_protection', 0.91))
        except:
            yield_igr_glass_protection = 0.91
        try:
            yield_igr_perimeter_tape = float(request.form.get('yield_igr_perimeter_tape', 0.91))
        except:
            yield_igr_perimeter_tape = 0.91
        try:
            yield_igr_structural_tape = float(request.form.get('yield_igr_structural_tape', 0.91))
        except:
            yield_igr_structural_tape = 0.91

        cp['yield_igr_glass'] = yield_igr_glass
        cp['yield_igr_extrusions'] = yield_igr_extrusions
        cp['yield_igr_gaskets'] = yield_igr_gaskets
        cp['yield_igr_glass_protection'] = yield_igr_glass_protection
        cp['yield_igr_perimeter_tape'] = yield_igr_perimeter_tape
        cp['yield_igr_structural_tape'] = yield_igr_structural_tape

        selected_igr_glass = request.form.get('igr_material_glass')
        selected_igr_extrusions = request.form.get('igr_material_extrusions')
        selected_igr_gaskets = request.form.get('igr_material_gaskets')
        selected_igr_glass_protection = request.form.get('igr_material_glass_protection')
        selected_igr_perimeter_tape = request.form.get('igr_material_perimeter_tape')
        selected_igr_structural_tape = request.form.get('igr_material_structural_tape')

        cp['igr_material_glass'] = selected_igr_glass
        cp['igr_material_extrusions'] = selected_igr_extrusions
        cp['igr_material_gaskets'] = selected_igr_gaskets
        cp['igr_material_glass_protection'] = selected_igr_glass_protection
        cp['igr_material_perimeter_tape'] = selected_igr_perimeter_tape
        cp['igr_material_structural_tape'] = selected_igr_structural_tape
        # Retrieve IGR Materials Note from the form
        cp["igr_note"] = request.form.get("igr_note", "")
        save_current_project(cp)

        mat_igr_glass = Material.query.get(selected_igr_glass) if selected_igr_glass else None
        mat_igr_extrusions = Material.query.get(selected_igr_extrusions) if selected_igr_extrusions else None
        mat_igr_gaskets = Material.query.get(selected_igr_gaskets) if selected_igr_gaskets else None
        mat_igr_glass_protection = Material.query.get(selected_igr_glass_protection) if selected_igr_glass_protection else None
        mat_igr_perimeter_tape = Material.query.get(selected_igr_perimeter_tape) if selected_igr_perimeter_tape else None
        mat_igr_structural_tape = Material.query.get(selected_igr_structural_tape) if selected_igr_structural_tape else None

        total_area = cp.get('igr_total_area', 0)
        total_perimeter = cp.get('igr_total_perimeter', 0)
        total_vertical = cp.get('igr_total_vertical_ft', 0)

        cost_igr_glass = (total_area * mat_igr_glass.yield_cost) / yield_igr_glass if mat_igr_glass else 0
        cost_igr_extrusions = (total_perimeter * mat_igr_extrusions.yield_cost) / yield_igr_extrusions if mat_igr_extrusions else 0
        cost_igr_gaskets = (total_vertical * mat_igr_gaskets.yield_cost) / yield_igr_gaskets if mat_igr_gaskets else 0
        cost_igr_glass_protection = (total_area * mat_igr_glass_protection.yield_cost) / yield_igr_glass_protection if mat_igr_glass_protection else 0

        if cp.get('igr_type') == "Dry Seal IGR":
            cost_igr_perimeter_tape = 0
        else:
            cost_igr_perimeter_tape = (total_perimeter * mat_igr_perimeter_tape.yield_cost) / yield_igr_perimeter_tape if mat_igr_perimeter_tape else 0

        cost_igr_structural_tape = (total_perimeter * mat_igr_structural_tape.yield_cost) / yield_igr_structural_tape if mat_igr_structural_tape else 0

        total_igr_material_cost = (cost_igr_glass + cost_igr_extrusions + cost_igr_gaskets +
                                   cost_igr_glass_protection + cost_igr_perimeter_tape + cost_igr_structural_tape)
        cp['igr_material_total_cost'] = total_igr_material_cost

        igr_material_map = {
            "IGR Glass (Cat 15)": mat_igr_glass,
            "IGR Extrusions (Cat 1)": mat_igr_extrusions,
            "IGR Gaskets (Cat 3)": mat_igr_gaskets,
            "IGR Glass Protection (Cat 7)": mat_igr_glass_protection,
            "IGR Perimeter Butyl Tape (Cat 10)": mat_igr_perimeter_tape,
            "IGR Structural Glazing Tape (Cat 10)": mat_igr_structural_tape
        }

        def get_igr_required_quantity(category):
            if category == "IGR Glass (Cat 15)":
                return total_area
            elif category == "IGR Extrusions (Cat 1)":
                return total_perimeter
            elif category == "IGR Gaskets (Cat 3)":
                return total_vertical
            elif category == "IGR Glass Protection (Cat 7)":
                return total_area
            elif category == "IGR Perimeter Butyl Tape (Cat 10)":
                return total_perimeter
            elif category == "IGR Structural Glazing Tape (Cat 10)":
                return total_perimeter
            else:
                return 0

        igr_items = [
            {
                "Category": "IGR Glass (Cat 15)",
                "Selected Material": mat_igr_glass.nickname if mat_igr_glass else "N/A",
                "Unit Cost": mat_igr_glass.yield_cost if mat_igr_glass else 0,
                "Calculation": f"Total Area {total_area:.2f} × Yield Cost / {yield_igr_glass}",
                "Cost ($)": cost_igr_glass
            },
            {
                "Category": "IGR Extrusions (Cat 1)",
                "Selected Material": mat_igr_extrusions.nickname if mat_igr_extrusions else "N/A",
                "Unit Cost": mat_igr_extrusions.yield_cost if mat_igr_extrusions else 0,
                "Calculation": f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {yield_igr_extrusions}",
                "Cost ($)": cost_igr_extrusions
            },
            {
                "Category": "IGR Gaskets (Cat 3)",
                "Selected Material": mat_igr_gaskets.nickname if mat_igr_gaskets else "N/A",
                "Unit Cost": mat_igr_gaskets.yield_cost if mat_igr_gaskets else 0,
                "Calculation": f"Total Vertical {total_vertical:.2f} × Yield Cost / {yield_igr_gaskets}",
                "Cost ($)": cost_igr_gaskets
            },
            {
                "Category": "IGR Glass Protection (Cat 7)",
                "Selected Material": mat_igr_glass_protection.nickname if mat_igr_glass_protection else "N/A",
                "Unit Cost": mat_igr_glass_protection.yield_cost if mat_igr_glass_protection else 0,
                "Calculation": f"Total Area {total_area:.2f} × Yield Cost / {yield_igr_glass_protection}",
                "Cost ($)": cost_igr_glass_protection
            },
            {
                "Category": "IGR Perimeter Butyl Tape (Cat 10)",
                "Selected Material": mat_igr_perimeter_tape.nickname if mat_igr_perimeter_tape else "N/A",
                "Unit Cost": mat_igr_perimeter_tape.yield_cost if mat_igr_perimeter_tape else 0,
                "Calculation": f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {yield_igr_perimeter_tape}",
                "Cost ($)": cost_igr_perimeter_tape
            },
            {
                "Category": "IGR Structural Glazing Tape (Cat 10)",
                "Selected Material": mat_igr_structural_tape.nickname if mat_igr_structural_tape else "N/A",
                "Unit Cost": mat_igr_structural_tape.yield_cost if mat_igr_structural_tape else 0,
                "Calculation": f"Total Perimeter {total_perimeter:.2f} × Yield Cost / {yield_igr_structural_tape}",
                "Cost ($)": cost_igr_structural_tape
            }
        ]

        for item in igr_items:
            req_qty = get_igr_required_quantity(item["Category"])
            mat_obj = igr_material_map.get(item["Category"])
            if mat_obj and req_qty > 0:
                stock_pct = (mat_obj.quantity / req_qty) * 100
            else:
                stock_pct = 0
            item["$ per SF"] = (item["Cost ($)"] / total_area if total_area > 0 else 0)
            item["% Total Cost"] = (item["Cost ($)"] / cp.get("igr_material_total_cost", 1) * 100 if cp.get("igr_material_total_cost", 0) > 0 else 0)
            item["Stock Level"] = f"{stock_pct:.2f}%"
            if mat_obj:
                item["Min Lead"] = mat_obj.min_lead if mat_obj.min_lead is not None else "N/A"
                item["Max Lead"] = mat_obj.max_lead if mat_obj.max_lead is not None else "N/A"
            else:
                item["Min Lead"] = "N/A"
                item["Max Lead"] = "N/A"
            discount_key = "discount_" + "".join(c if c.isalnum() else "_" for c in item["Category"])
            discount_str = (request.form.get(discount_key) or "").strip()
            if discount_str != "":
                try:
                    discount_value = float(discount_str)
                except:
                    discount_value = 0
            else:
                discount_value = cp.get(discount_key, 0)
            cp[discount_key] = discount_value
            item["Discount/Surcharge"] = discount_value
            item["Final Cost"] = item["Cost ($)"] + discount_value

        total_final_cost = sum(item["Final Cost"] for item in igr_items)

        table_html = "<form method='POST'>"
        igr_headers = ["Category", "Selected Material", "Unit Cost", "Calculation", "Cost ($)", "$ per SF", "% Total Cost", "Stock Level", "Min Lead", "Max Lead", "Discount/Surcharge", "Final Cost"]
        table_html += "<table class='summary-table'><tr>"
        for header in igr_headers:
            table_html += f"<th>{header}</th>"
        table_html += "</tr>"
        for item in igr_items:
            discount_key = "discount_" + "".join(c if c.isalnum() else "_" for c in item["Category"])
            table_html += "<tr>"
            table_html += f"<td>{item['Category']}</td>"
            table_html += f"<td>{item['Selected Material']}</td>"
            table_html += f"<td>{item['Unit Cost']:.2f}</td>"
            table_html += f"<td>{item['Calculation']}</td>"
            table_html += f"<td class='cost'>{item['Cost ($)']:.2f}</td>"
            table_html += f"<td>{(item['Final Cost']/total_area if total_area>0 else 0):.2f}</td>"
            table_html += f"<td>{(item['Final Cost']/total_final_cost*100 if total_final_cost>0 else 0):.2f}</td>"
            table_html += f"<td>{item['Stock Level']}</td>"
            table_html += f"<td>{item['Min Lead']}</td>"
            table_html += f"<td>{item['Max Lead']}</td>"
            table_html += f"<td><input type='number' step='1' name='{discount_key}' value='{item['Discount/Surcharge']}' oninput='updateFinalCost(this)' /></td>"
            table_html += f"<td class='final-cost'>{item['Final Cost']:.2f}</td>"
            table_html += "</tr>"
        table_html += "</table>"
        table_html += "<div style='margin-top:10px;'><label for='igr_note'>IGR Materials Note:</label><textarea id='igr_note' name='igr_note' rows='3' style='width:100%;'>" + cp.get("igr_note", "") + "</textarea></div>"
        table_html += """
        <script>
        function updateFinalCost(input) {
            var row = input.closest("tr");
            var costCell = row.querySelector(".cost");
            var finalCostCell = row.querySelector(".final-cost");
            var discountValue = parseFloat(input.value) || 0;
            var costValue = parseFloat(costCell.innerText) || 0;
            var finalCost = costValue + discountValue;
            finalCostCell.innerText = finalCost.toFixed(2);
        }
        </script>
        """
        table_html += "</form>"

        cp["igr_itemized_costs"] = igr_items
        save_current_project(cp)

        return f"""
         <html>
           <head>
             <title>IGR Material Cost Summary</title>
             <style>{common_css}</style>
           </head>
           <body>
             <div class="container">
               <h2>IGR Material Cost Summary</h2>
               {table_html}
               <div class="btn-group">
                 <button type="button" class="btn" onclick="window.location.href='/igr_materials'">Back: Edit IGR Materials</button>
                 <button type="button" class="btn" onclick="window.location.href='/other_costs'">Next: Other Costs</button>
               </div>
             </div>
           </body>
         </html>
        """

    def generate_options(materials_list, selected_value=None):
        options = ""
        for m in materials_list:
            if str(m.id) == str(selected_value):
                options += f'<option value="{m.id}" selected>{m.nickname} - ${m.yield_cost:.2f}</option>'
            else:
                options += f'<option value="{m.id}">{m.nickname} - ${m.yield_cost:.2f}</option>'
        return options

    form_html = f"""
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
                  <label for="igr_material_perimeter_tape">Select Perimeter Butyl Tape:</label>
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
                  <button type="button" class="btn" onclick="window.location.href='/materials'">Back: SWR Materials</button>
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
    return form_html

# =========================
# OTHER COSTS PAGE (Combined SWR + IGR)
# =========================
@app.route('/other_costs', methods=['GET', 'POST'])
def other_costs():
    cp = get_current_project()
    swr_material = cp.get('material_total_cost', 0)
    igr_material = cp.get('igr_material_total_cost', 0)
    combined_material_cost = swr_material + igr_material
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
        grand_total = combined_material_cost + additional_total

        cp["total_truck_cost"] = total_truck_cost
        cp["installation_cost"] = installation_cost
        cp["equipment_cost"] = equipment_cost
        cp["travel_cost"] = travel_cost
        cp["sales_cost"] = sales_cost
        cp["additional_total"] = additional_total
        cp["grand_total"] = grand_total
        # Retrieve Travel Note from the form
        cp["travel_note"] = request.form.get("travel_note", "")
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
                 <tr><td>Combined Material Cost</td><td>{combined_material_cost:.2f}</td></tr>
                 <tr><td>Truck Cost</td><td>{total_truck_cost:.2f}</td></tr>
                 <tr><td>Installation</td><td>{installation_cost:.2f}</td></tr>
                 <tr><td>Equipment</td><td>{equipment_cost:.2f}</td></tr>
                 <tr><td>Travel</td><td>{travel_cost:.2f}</td></tr>
                 <tr><td>Sales</td><td>{sales_cost:.2f}</td></tr>
                 <tr><td><strong>Additional Total</strong></td><td><strong>{additional_total:.2f}</strong></td></tr>
                 <tr><th>Grand Total</th><th>{grand_total:.2f}</th></tr>
               </table>
               <div style="margin-top:10px;">
                  <label for="travel_note">Travel Note:</label>
                  <textarea id="travel_note" name="travel_note" rows="3" style="width:100%;">{cp.get("travel_note", "")}</textarea>
               </div>
               <div class="btn-group">
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

    form_html = f"""
    <html>
      <head>
         <title>Other Costs</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Enter Additional Costs</h2>
            <form method="POST">
              <div>
                <label for="num_trucks">Number of Trucks:</label>
                <input type="number" step="1" id="num_trucks" name="num_trucks" value="{cp.get('num_trucks', '0')}" required>
              </div>
              <div>
                <label for="truck_cost">Cost per Truck ($):</label>
                <input type="number" step="100" id="truck_cost" name="truck_cost" value="{cp.get('truck_cost', '0')}" required>
              </div>
              <div>
                <label for="hourly_rate">Hourly Rate ($/hr):</label>
                <input type="number" step="1" id="hourly_rate" name="hourly_rate" value="{cp.get('hourly_rate', '0')}" required>
              </div>
              <div>
                <label for="hours_per_panel">Hours per Panel:</label>
                <input type="number" step="0.01" id="hours_per_panel" name="hours_per_panel" value="{cp.get('hours_per_panel', '0')}" required>
              </div>
              <!-- Additional cost fields remain unchanged -->
              <div style="margin-top:10px;">
                  <label for="travel_note">Travel Note:</label>
                  <textarea id="travel_note" name="travel_note" rows="3" style="width:100%;">{cp.get("travel_note", "")}</textarea>
              </div>
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
      </body>
    </html>
    """
    return form_html

# =========================
# MARGINS PAGE (Combined)
# =========================
@app.route('/margins', methods=['GET', 'POST'])
def margins():
    cp = get_current_project()
    combined_material = cp.get("material_total_cost", 0) + cp.get("igr_material_total_cost", 0)
    base_costs = {
        "Panel Total": combined_material,
        "Logistics": cp.get("total_truck_cost", 0),
        "Installation": cp.get("installation_cost", 0),
        "Equipment": cp.get("equipment_cost", 0),
        "Travel": cp.get("travel_cost", 0),
        "Sales": cp.get("sales_cost", 0)
    }
    total_area = cp.get("swr_total_area", 0)

    if request.method == 'POST':
        margins_values = {}
        for category in base_costs:
            try:
                margins_values[category] = float(request.form.get(f"{category}_margin", 0))
            except:
                margins_values[category] = 0
        adjusted_costs = {cat: base_costs[cat] * (1 + margins_values[cat] / 100) for cat in base_costs}
        final_total = sum(adjusted_costs.values())

        summary_data = {
            "Category": list(base_costs.keys()) + ["Grand Total"],
            "Original Cost ($)": [base_costs[cat] for cat in base_costs] + [sum(base_costs.values())],
            "Margin (%)": [margins_values[cat] for cat in base_costs] + [""],
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
                "Margin (%)": margins_values.get(cat, 0),
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
               <div style="margin-top:10px;">
                 <label for="travel_note">Travel Note:</label>
                 <textarea id="travel_note" name="travel_note" rows="3" style="width:100%;">{cp.get("travel_note", "")}</textarea>
               </div>
               <div class="btn-group">
                 <div class="btn-left">
                    <button type="button" class="btn" onclick="window.location.href='/other_costs'">Back to Other Costs</button>
                 </div>
                 <div class="btn-right">
                    <button type="button" class="btn" onclick="window.location.href='/materials'">Back to SWR Materials</button>
                 </div>
                 <div class="btn-right">
                    <button type="button" class="btn" onclick="window.location.href='/'">Start New Project</button>
                 </div>
               </div>
             </div>
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
            var baseCosts = {json.dumps(base_costs)};
            var totalArea = {total_area};
            function updateOutput(sliderId, outputId) {{
                var slider = document.getElementById(sliderId);
                var output = document.getElementById(outputId);
                output.value = slider.value;
                updateGrandPSF();
            }}
            function updateGrandPSF() {{
                var categories = ["Panel Total", "Logistics", "Installation", "Equipment", "Travel", "Sales"];
                var sum = 0;
                for (var i = 0; i < categories.length; i++) {{
                    var cat = categories[i];
                    var var_id = cat.replace(" ", "_");
                    var slider = document.getElementById(var_id + "_margin");
                    var margin = parseFloat(slider.value);
                    var cost = baseCosts[cat];
                    var adjusted = cost * (1 + margin / 100);
                    sum += adjusted;
                }}
                var grandDisplay = document.getElementById("grand_psf");
                if(totalArea > 0) {{
                    var psf = sum / totalArea;
                    grandDisplay.innerText = "$" + psf.toFixed(2) + " per SF";
                }} else {{
                    grandDisplay.innerText = "N/A";
                }}
            }}
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
                <div class="margin-row">
                   <label for="{var_id}_margin">{category} Margin (%):</label>
                   <input type="range" id="{var_id}_margin" name="{category}_margin" min="0" max="100" step="1" value="{cp.get(category + '_margin', '0')}" oninput="updateOutput('{var_id}_margin', '{var_id}_output')">
                   <output id="{var_id}_output">{cp.get(category + '_margin', '0')}</output>
                </div>
        """
    form_html += """
                <div class="margin-row" style="justify-content:center;">
                    <label style="margin-right:10px;">Grand Total $ per SF:</label>
                    <span id="grand_psf">$0.00 per SF</span>
                </div>
                <div class="btn-group">
                    <div class="btn-left">
                        <button type="button" class="btn" onclick="window.location.href='/other_costs'">Back to Other Costs</button>
                    </div>
                    <div class="btn-right">
                        <button type="submit" class="btn">Calculate Final Cost with Margins</button>
                    </div>
                </div>
            </form>
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
    # Header information
    writer.writerow(["Customer Name:", cp.get("customer_name", "N/A")])
    writer.writerow(["Project Name:", cp.get("project_name", "Unnamed Project")])
    writer.writerow(["Estimated by:", cp.get("estimated_by", "N/A")])
    writer.writerow(["Date:", datetime.datetime.now().strftime("%Y-%m-%d")])
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
    writer.writerow(["Category", "Final Cost"])
    writer.writerow([])
    writer.writerow(["Detailed SWR Itemized Costs"])
    writer.writerow(["Category", "Selected Material", "Final Cost"])
    for item in cp.get("itemized_costs", []):
        writer.writerow([
            item.get("Category", ""),
            item.get("Selected Material", ""),
            item.get("Final Cost", 0)
        ])
    writer.writerow(["Materials Note:", cp.get("swr_note", "")])
    writer.writerow([])
    writer.writerow(["Detailed IGR Itemized Costs"])
    writer.writerow(["Category", "Selected Material", "Final Cost"])
    for item in cp.get("igr_itemized_costs", []):
        writer.writerow([
            item.get("Category", ""),
            item.get("Selected Material", ""),
            item.get("Final Cost", 0)
        ])
    writer.writerow(["IGR Materials Note:", cp.get("igr_note", "")])
    writer.writerow([])
    writer.writerow(["Other Costs Note (Travel):", cp.get("travel_note", "")])
    writer.writerow([])
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
    ws.write("A1", "Customer Name:")
    ws.write("B1", cp.get("customer_name", "N/A"))
    ws.write("A2", "Project Name:")
    ws.write("B2", cp.get("project_name", "Unnamed Project"))
    ws.write("A3", "Estimated by:")
    ws.write("B3", cp.get("estimated_by", "N/A"))
    ws.write("A4", "Date:")
    ws.write("B4", datetime.datetime.now().strftime("%Y-%m-%d"))
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
    ws.write("G1", "Detailed SWR Itemized Costs")
    headers_detail = ["Category", "Selected Material", "Final Cost"]
    for col, header in enumerate(headers_detail):
        ws.write(1, col, header)
    row = 2
    for item in cp.get("itemized_costs", []):
        ws.write(row, 0, item.get("Category", ""))
        ws.write(row, 1, item.get("Selected Material", ""))
        ws.write(row, 2, item.get("Final Cost", 0))
        row += 1
    ws.write(row, 0, "Materials Note:")
    ws.write(row, 1, cp.get("swr_note", ""))
    row += 2
    ws.write(row, 0, "Detailed IGR Itemized Costs")
    for col, header in enumerate(headers_detail):
        ws.write(row+1, col, header)
    row = row + 2
    for item in cp.get("igr_itemized_costs", []):
        ws.write(row, 0, item.get("Category", ""))
        ws.write(row, 1, item.get("Selected Material", ""))
        ws.write(row, 2, item.get("Final Cost", 0))
        row += 1
    ws.write(row, 0, "IGR Materials Note:")
    ws.write(row, 1, cp.get("igr_note", ""))
    row += 2
    ws.write(row, 0, "Other Costs Note (Travel):")
    ws.write(row, 1, cp.get("travel_note", ""))
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