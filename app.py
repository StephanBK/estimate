from flask import Flask, request, redirect, url_for, send_file, Response
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import os
import io

app = Flask(__name__)

# Database Configuration
app.config[
    'SQLALCHEMY_DATABASE_URI'] = 'postgresql://u7vukdvn20pe3c:p918802c410825b956ccf24c5af8d168b4d9d69e1940182bae9bd8647eb606845@cb5ajfjosdpmil.cluster-czrs8kj4isg7.us-east-1.rds.amazonaws.com:5432/dcobttk99a5sie'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Initialize the database
db = SQLAlchemy(app)


# Define the Material model (using 'nickname' for display)
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


# Global storage for current project details
current_project = {}

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


# -------------------------
# INDEX & SUMMARY PAGES
# -------------------------
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        project_name = request.form['project_name']
        project_number = request.form['project_number']
        file = request.files['file']
        if file:
            file_path = os.path.join('uploads', file.filename)
            os.makedirs('uploads', exist_ok=True)
            file.save(file_path)
            current_project['project_name'] = project_name
            current_project['project_number'] = project_number
            current_project['file_path'] = file_path
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
               <label for="file">Upload Filled Template:</label>
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


@app.route('/summary')
def summary():
    # Read CSV and compute basic metrics
    file_path = current_project.get('file_path')
    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        return f"<h2 style='color: red;'>Error reading the file: {e}</h2>"

    # Separate by Type column if present; otherwise assume all SWR
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
        df_subset['Total Vertical (ft)'] = (df_subset['VGA Height in'] * df_subset['Qty']) / 12
        df_subset['Total Horizontal (ft)'] = (df_subset['VGA Width in'] * df_subset['Qty']) / 12
        total_area = df_subset['Total Area (sq ft)'].sum()
        total_perimeter = df_subset['Total Perimeter (ft)'].sum()
        total_vertical = df_subset['Total Vertical (ft)'].sum()
        total_horizontal = df_subset['Total Horizontal (ft)'].sum()
        total_quantity = df_subset['Qty'].sum()
        return total_area, total_perimeter, total_vertical, total_horizontal, total_quantity

    swr_area, swr_perimeter, swr_vertical, swr_horizontal, swr_quantity = compute_totals(df_swr)
    igr_area, igr_perimeter, igr_vertical, igr_horizontal, igr_quantity = compute_totals(
        df_igr) if not df_igr.empty else (0, 0, 0, 0, 0)

    current_project['swr_total_area'] = swr_area
    current_project['swr_total_perimeter'] = swr_perimeter
    current_project['swr_total_vertical_ft'] = swr_vertical
    current_project['swr_total_horizontal_ft'] = swr_horizontal
    current_project['swr_total_quantity'] = swr_quantity

    current_project['igr_total_area'] = igr_area
    current_project['igr_total_perimeter'] = igr_perimeter
    current_project['igr_total_vertical_ft'] = igr_vertical
    current_project['igr_total_horizontal_ft'] = igr_horizontal
    current_project['igr_total_quantity'] = igr_quantity

    def add_calculation_columns(df_subset):
        df_subset = df_subset.copy()
        df_subset["Square Footage (sq ft)"] = (df_subset["VGA Width in"] * df_subset["VGA Height in"] * df_subset[
            "Qty"]) / 144
        df_subset["Total Perimeter (ft)"] = (2 * (df_subset["VGA Width in"] + df_subset["VGA Height in"]) * df_subset[
            "Qty"]) / 12
        df_subset["Total Vertical (ft)"] = (df_subset["VGA Height in"] * df_subset["Qty"]) / 12
        df_subset["Total Horizontal (ft)"] = (df_subset["VGA Width in"] * df_subset["Qty"]) / 12
        return df_subset

    # Add new columns and sort by Qty descending
    df_swr_calc = add_calculation_columns(df_swr).sort_values(by="Qty", ascending=False)
    if not df_igr.empty:
        df_igr_calc = add_calculation_columns(df_igr).sort_values(by="Qty", ascending=False)
    else:
        df_igr_calc = pd.DataFrame()

    summary_html = f"""
    <html>
      <head>
         <title>Project Summary</title>
         <style>{common_css}</style>
      </head>
      <body>
         <div class="container">
            <h2>Project Summary</h2>
            <p><strong>Project Name:</strong> {current_project.get('project_name')}</p>
            <p><strong>Project Number:</strong> {current_project.get('project_number')}</p>
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
            <h3>Detailed SWR Items</h3>
            {df_swr_calc.to_html(classes='data-table', index=False, float_format="%.2f")}
    """
    if not df_igr_calc.empty:
        summary_html += f"""
            <h3>Detailed IGR Items</h3>
            {df_igr_calc.to_html(classes='data-table', index=False, float_format="%.2f")}
        """
    summary_html += f"""
            <button onclick="window.location.href='/'">Start New Project</button>
            <button onclick="window.location.href='/materials'">Next: SWR Materials</button>
         </div>
      </body>
    </html>
    """
    return summary_html


# -------------------------
# MATERIALS PAGE (SWR Materials Only)
# -------------------------
@app.route('/materials', methods=['GET', 'POST'])
def materials():
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
    except Exception as e:
        return f"<h2 style='color: red;'>Error fetching materials: {e}</h2>"

    if request.method == 'POST':
        selected_glass = request.form.get('material_glass')
        aluminum_option = request.form.get('aluminum_option')
        selected_aluminum = request.form.get('material_aluminum')
        selected_glazing = request.form.get('material_glazing')
        selected_gaskets = request.form.get('material_gaskets')
        selected_corner_keys = request.form.get('material_corner_keys')
        selected_dual_lock = request.form.get('material_dual_lock')
        selected_foam_baffle = request.form.get('material_foam_baffle')
        selected_glass_protection = request.form.get('material_glass_protection')
        selected_tape = request.form.get('material_tape')
        tape_option = request.form.get('tape_option')

        mat_glass = Material.query.get(selected_glass) if selected_glass else None
        mat_aluminum = Material.query.get(selected_aluminum) if selected_aluminum else None
        mat_glazing = Material.query.get(selected_glazing) if selected_glazing else None
        mat_gaskets = Material.query.get(selected_gaskets) if selected_gaskets else None
        mat_corner_keys = Material.query.get(selected_corner_keys) if selected_corner_keys else None
        mat_dual_lock = Material.query.get(selected_dual_lock) if selected_dual_lock else None
        mat_foam_baffle = Material.query.get(selected_foam_baffle) if selected_foam_baffle else None
        mat_glass_protection = Material.query.get(selected_glass_protection) if selected_glass_protection else None
        mat_tape = Material.query.get(selected_tape) if selected_tape else None

        total_area = current_project.get('swr_total_area', 0)
        total_perimeter = current_project.get('swr_total_perimeter', 0)
        total_vertical = current_project.get('swr_total_vertical_ft', 0)
        total_horizontal = current_project.get('swr_total_horizontal_ft', 0)
        total_quantity = current_project.get('swr_total_quantity', 0)

        cost_glass = (total_area * mat_glass.cost) if mat_glass else 0
        if mat_aluminum:
            if aluminum_option == 'head_retainer':
                cost_aluminum = ((total_perimeter + (total_horizontal / 2)) * mat_aluminum.cost)
            elif aluminum_option == 'head_and_sill':
                cost_aluminum = ((total_perimeter + total_horizontal) * mat_aluminum.cost)
            else:
                cost_aluminum = 0
        else:
            cost_aluminum = 0
        cost_glazing = (total_perimeter * mat_glazing.cost) if mat_glazing else 0
        cost_gaskets = (total_vertical * mat_gaskets.cost) if mat_gaskets else 0
        cost_corner_keys = (total_quantity * 4 * mat_corner_keys.cost) if mat_corner_keys else 0
        cost_dual_lock = (total_vertical * mat_dual_lock.cost) if mat_dual_lock else 0
        cost_foam_baffle = (total_horizontal * mat_foam_baffle.cost) if mat_foam_baffle else 0
        cost_glass_protection = (total_horizontal * mat_glass_protection.cost) if mat_glass_protection else 0
        if mat_tape:
            if tape_option == 'head_retainer':
                cost_tape = ((total_horizontal / 2) * mat_tape.cost)
            else:
                cost_tape = (total_horizontal * mat_tape.cost)
        else:
            cost_tape = 0

        total_material_cost = (cost_glass + cost_aluminum + cost_glazing + cost_gaskets +
                               cost_corner_keys + cost_dual_lock + cost_foam_baffle +
                               cost_glass_protection + cost_tape)
        current_project['material_total_cost'] = total_material_cost

        result_html = f"""
         <html>
           <head>
             <title>SWR Material Cost Summary</title>
             <style>{common_css}</style>
           </head>
           <body>
             <div class="container">
               <h2>SWR Material Cost Summary</h2>
               <table class="summary-table">
                 <tr>
                   <th>Category</th>
                   <th>Selected Material</th>
                   <th>Unit Cost</th>
                   <th>Calculation</th>
                   <th>Cost ($)</th>
                 </tr>
                 <tr>
                   <td>Glass (Cat 15)</td>
                   <td>{mat_glass.nickname if mat_glass else "N/A"}</td>
                   <td>{mat_glass.cost if mat_glass else 0:.2f}</td>
                   <td>Total Area (sq ft): {total_area:.2f} × Cost</td>
                   <td>{cost_glass:.2f}</td>
                 </tr>
                 <tr>
                   <td>Aluminum (Cat 1)</td>
                   <td>{mat_aluminum.nickname if mat_aluminum else "N/A"}</td>
                   <td>{mat_aluminum.cost if mat_aluminum else 0:.2f}</td>
                   <td>
                     Option: {aluminum_option}<br>
                     {"(Perimeter + 0.5×Horizontal)" if aluminum_option == "head_retainer" else "(Perimeter + Horizontal)"} × Cost
                   </td>
                   <td>{cost_aluminum:.2f}</td>
                 </tr>
                 <tr>
                   <td>Glazing Spline (Cat 2)</td>
                   <td>{mat_glazing.nickname if mat_glazing else "N/A"}</td>
                   <td>{mat_glazing.cost if mat_glazing else 0:.2f}</td>
                   <td>Total Perimeter (ft): {total_perimeter:.2f} × Cost</td>
                   <td>{cost_glazing:.2f}</td>
                 </tr>
                 <tr>
                   <td>Gaskets (Cat 3)</td>
                   <td>{mat_gaskets.nickname if mat_gaskets else "N/A"}</td>
                   <td>{mat_gaskets.cost if mat_gaskets else 0:.2f}</td>
                   <td>Total Vertical (ft): {total_vertical:.2f} × Cost</td>
                   <td>{cost_gaskets:.2f}</td>
                 </tr>
                 <tr>
                   <td>Corner Keys (Cat 4)</td>
                   <td>{mat_corner_keys.nickname if mat_corner_keys else "N/A"}</td>
                   <td>{mat_corner_keys.cost if mat_corner_keys else 0:.2f}</td>
                   <td>Total Quantity: {total_quantity:.2f} × 4 × Cost</td>
                   <td>{cost_corner_keys:.2f}</td>
                 </tr>
                 <tr>
                   <td>Dual Lock (Cat 5)</td>
                   <td>{mat_dual_lock.nickname if mat_dual_lock else "N/A"}</td>
                   <td>{mat_dual_lock.cost if mat_dual_lock else 0:.2f}</td>
                   <td>Total Vertical (ft): {total_vertical:.2f} × Cost</td>
                   <td>{cost_dual_lock:.2f}</td>
                 </tr>
                 <tr>
                   <td>Foam Baffle (Cat 6)</td>
                   <td>{mat_foam_baffle.nickname if mat_foam_baffle else "N/A"}</td>
                   <td>{mat_foam_baffle.cost if mat_foam_baffle else 0:.2f}</td>
                   <td>Total Horizontal (ft): {total_horizontal:.2f} × Cost</td>
                   <td>{cost_foam_baffle:.2f}</td>
                 </tr>
                 <tr>
                   <td>Glass Protection (Cat 7)</td>
                   <td>{mat_glass_protection.nickname if mat_glass_protection else "N/A"}</td>
                   <td>{mat_glass_protection.cost if mat_glass_protection else 0:.2f}</td>
                   <td>Total Horizontal (ft): {total_horizontal:.2f} × Cost</td>
                   <td>{cost_glass_protection:.2f}</td>
                 </tr>
                 <tr>
                   <td>Tape (Cat 10)</td>
                   <td>{mat_tape.nickname if mat_tape else "N/A"}</td>
                   <td>{mat_tape.cost if mat_tape else 0:.2f}</td>
                   <td>
                     Option: {tape_option}<br>
                     {"(Half Horizontal)" if tape_option == "head_retainer" else "(Full Horizontal)"}<br>
                     × Total Horizontal (ft): {total_horizontal:.2f} × Cost
                   </td>
                   <td>{cost_tape:.2f}</td>
                 </tr>
                 <tr>
                   <th colspan="4">SWR Material Total Cost</th>
                   <th>{total_material_cost:.2f}</th>
                 </tr>
               </table>
               <button onclick="window.location.href='/other_costs'">Other Costs</button>
               <button onclick="window.location.href='/'">Start New Project</button>
             </div>
           </body>
         </html>
        """
        return result_html

    def generate_options(materials_list):
        options = ""
        for m in materials_list:
            options += f'<option value="{m.id}">{m.nickname} - ${m.cost:.2f}</option>'
        return options

    options_glass = generate_options(materials_glass)
    options_aluminum = generate_options(materials_aluminum)
    options_glazing = generate_options(materials_glazing)
    options_gaskets = generate_options(materials_gaskets)
    options_corner_keys = generate_options(materials_corner_keys)
    options_dual_lock = generate_options(materials_dual_lock)
    options_foam_baffle = generate_options(materials_foam_baffle)
    options_glass_protection = generate_options(materials_glass_protection)
    options_tape = generate_options(materials_tape)

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
               <label for="material_glass">Glass (Cat 15 - Total Area (sq ft) × Cost):</label>
               <select name="material_glass" id="material_glass" required>
                  {options_glass}
               </select>

               <!-- Aluminum: option appears above its picker -->
               <label for="aluminum_option">Aluminum Option:</label>
               <select name="aluminum_option" id="aluminum_option" required>
                  <option value="head_retainer">Head Retainer (Total Perimeter + 0.5 × Total Horizontal)</option>
                  <option value="head_and_sill">Head and Sill (Total Perimeter + Total Horizontal)</option>
               </select>
               <label for="material_aluminum">Aluminum (Cat 1):</label>
               <select name="material_aluminum" id="material_aluminum" required>
                  {options_aluminum}
               </select>

               <label for="material_glazing">Glazing Spline (Cat 2 - Total Perimeter (ft) × Cost):</label>
               <select name="material_glazing" id="material_glazing" required>
                  {options_glazing}
               </select>

               <label for="material_gaskets">Gaskets (Cat 3 - Total Vertical (ft) × Cost):</label>
               <select name="material_gaskets" id="material_gaskets" required>
                  {options_gaskets}
               </select>

               <label for="material_corner_keys">Corner Keys (Cat 4 - Total Quantity × 4 × Cost):</label>
               <select name="material_corner_keys" id="material_corner_keys" required>
                  {options_corner_keys}
               </select>

               <label for="material_dual_lock">Dual Lock (Cat 5 - Total Vertical (ft) × Cost):</label>
               <select name="material_dual_lock" id="material_dual_lock" required>
                  {options_dual_lock}
               </select>

               <label for="material_foam_baffle">Foam Baffle (Cat 6 - Total Horizontal (ft) × Cost):</label>
               <select name="material_foam_baffle" id="material_foam_baffle" required>
                  {options_foam_baffle}
               </select>

               <label for="material_glass_protection">Glass Protection (Cat 7 - Total Horizontal (ft) × Cost):</label>
               <select name="material_glass_protection" id="material_glass_protection" required>
                  {options_glass_protection}
               </select>

               <label for="tape_option">Tape Option:</label>
               <select name="tape_option" id="tape_option" required>
                  <option value="head_retainer">Head Retainer (Half Horizontal)</option>
                  <option value="head_sill">Head+Sill (Full Horizontal)</option>
               </select>
               <label for="material_tape">Tape (Cat 10):</label>
               <select name="material_tape" id="material_tape" required>
                  {options_tape}
               </select>

               <button type="submit">Calculate SWR Material Costs</button>
            </form>
            <button onclick="window.location.href='/summary'">Back to Summary</button>
         </div>
      </body>
    </html>
    """


# -------------------------
# OTHER COSTS PAGE
# -------------------------
@app.route('/other_costs', methods=['GET', 'POST'])
def other_costs():
    # Capture additional costs: Logistics, Installation, Equipment, Travel, Sales.
    material_cost = current_project.get('material_total_cost', 0)
    total_quantity = current_project.get('total_quantity', 0)
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

        additional_total = total_truck_cost + installation_cost + equipment_cost + travel_cost + sales_cost
        grand_total = material_cost + additional_total

        current_project["total_truck_cost"] = total_truck_cost
        current_project["installation_cost"] = installation_cost
        current_project["equipment_cost"] = equipment_cost
        current_project["travel_cost"] = travel_cost
        current_project["sales_cost"] = sales_cost
        current_project["additional_total"] = additional_total
        current_project["grand_total"] = grand_total

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


# -------------------------
# MARGINS PAGE
# -------------------------
@app.route('/margins', methods=['GET', 'POST'])
def margins():
    base_costs = {
        "Panel Total": current_project.get("material_total_cost", 0),
        "Logistics": current_project.get("total_truck_cost", 0),
        "Installation": current_project.get("installation_cost", 0),
        "Equipment": current_project.get("equipment_cost", 0),
        "Travel": current_project.get("travel_cost", 0),
        "Sales": current_project.get("sales_cost", 0)
    }
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
               <button onclick="window.location.href='/'">Start New Project</button>
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
            function updateOutput(sliderId, outputId) {{
                var slider = document.getElementById(sliderId);
                var output = document.getElementById(outputId);
                output.value = slider.value;
            }}
         </script>
      </head>
      <body>
         <div class="container">
            <h2>Set Margins for Each Cost Category</h2>
            <form method="POST">
    """
    for category in base_costs:
        form_html += f"""
                <label for="{category}_margin">{category} Margin (%):</label>
                <input type="range" id="{category}_margin" name="{category}_margin" min="0" max="100" step="1" value="0" oninput="updateOutput('{category}_margin', '{category}_output')">
                <output id="{category}_output">0</output><br>
        """
    form_html += """
                <button type="submit">Calculate Final Cost with Margins</button>
            </form>
            <button onclick="window.location.href='/other_costs'">Back to Other Costs</button>
         </div>
      </body>
    </html>
    """
    return form_html


# -------------------------
# DOWNLOAD FINAL SUMMARY
# -------------------------
@app.route('/download_final_summary', methods=['POST'])
def download_final_summary():
    csv_data = request.form.get('csv_data', '')
    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-disposition": "attachment; filename=final_cost_summary.csv"}
    )


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)