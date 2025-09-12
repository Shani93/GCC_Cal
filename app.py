from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import numpy as np
from datetime import datetime
import logging
import io
import re

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Helper function to clean tier names
def clean_tier_name(tier):
    """Clean tier names by removing extra spaces and standardizing format"""
    if pd.isna(tier):
        return tier
    # Remove extra spaces and standardize format
    cleaned = re.sub(r'\s+', ' ', str(tier).strip())
    # Ensure format is "Tier X"
    if re.match(r'^Tier\s+\d+$', cleaned):
        return cleaned
    return tier

# Load data from Excel
def load_data():
    try:
        # Read the Excel file
        excel_file = "GCC Calculator 3.xlsx"
        
        # Load the necessary sheets
        real_estate_df = pd.read_excel(excel_file, sheet_name="Real_Estate")
        it_infra_df = pd.read_excel(excel_file, sheet_name="IT_Infra")
        plans_df = pd.read_excel(excel_file, sheet_name="Plans")
        lookup_helper_df = pd.read_excel(excel_file, sheet_name="Lookup_Helper")
        
        # Clean tier names in the dataframes
        real_estate_df["Tier"] = real_estate_df["Tier"].apply(clean_tier_name)
        it_infra_df["Tier"] = it_infra_df["Tier"].apply(clean_tier_name)
        
        # Validate data
        dfs = {
            'Real_Estate': real_estate_df,
            'IT_Infra': it_infra_df,
            'Plans': plans_df
        }
        
        required_columns = {
            'Real_Estate': ['Tier', 'City', 'Cost_INR_PM'],
            'IT_Infra': ['Tier', 'City', 'Cost_INR_PM'],
            'Plans': ['MinHC', 'MaxHC', 'Enab_Basic', 'Enab_Premium', 'Enab_Advance', 'Tech_Basic', 'Tech_Premium', 'Tech_Advance']
        }
        
        for sheet, columns in required_columns.items():
            if sheet in dfs and not all(col in dfs[sheet].columns for col in columns):
                missing = [col for col in columns if col not in dfs[sheet].columns]
                raise ValueError(f"Missing required columns in {sheet} sheet: {missing}")
        
        logger.info("Data loaded successfully")
        return real_estate_df, it_infra_df, plans_df, lookup_helper_df
        
    except Exception as e:
        logger.error(f"Error loading data: {str(e)}")
        return None, None, None, None

# Calculate costs
def calculate_costs(headcount, selected_city, selected_plan, real_estate_toggle, 
                   it_infra_toggle, enabling_toggle, technology_toggle,
                   real_estate_df, it_infra_df, plans_df):
    try:
        # Initialize all costs to zero
        total_real_estate_cost = 0.0
        total_it_infra_cost = 0.0
        enab_cost = 0.0
        tech_cost = 0.0
        
        # Get real estate cost only if enabled
        if real_estate_toggle:
            real_estate_cost_per_seat = real_estate_df[real_estate_df["City"] == selected_city]["Cost_INR_PM"].values[0]
            total_real_estate_cost = float(real_estate_cost_per_seat * headcount)
        
        # Get IT infrastructure cost only if enabled
        if it_infra_toggle:
            it_infra_cost_per_seat = it_infra_df[it_infra_df["City"] == selected_city]["Cost_INR_PM"].values[0]
            total_it_infra_cost = float(it_infra_cost_per_seat * headcount)
        
        # Get enabling functions and technology cost based on headcount and plan
        plan_row = None
        for index, row in plans_df.iterrows():
            if row["MinHC"] <= headcount <= row["MaxHC"]:
                plan_row = row
                break
        
        if plan_row is None:
            return None, None, None, None, None, None
        
        # Get enabling functions cost only if enabled
        if enabling_toggle:
            if selected_plan == "Basic":
                enab_cost = float(plan_row["Enab_Basic"])
            elif selected_plan == "Premium":
                enab_cost = float(plan_row["Enab_Premium"])
            else:  # Advance
                enab_cost = float(plan_row["Enab_Advance"])
        
        # Get technology cost only if enabled
        if technology_toggle:
            if selected_plan == "Basic":
                tech_cost = float(plan_row["Tech_Basic"])
            elif selected_plan == "Premium":
                tech_cost = float(plan_row["Tech_Premium"])
            else:  # Advance
                tech_cost = float(plan_row["Tech_Advance"])
        
        total_plan_cost = enab_cost + tech_cost
        
        # Calculate total cost
        total_cost = total_real_estate_cost + total_it_infra_cost + total_plan_cost
        
        # Calculate hourly cost per head in USD
        usd_inr_rate = 85  # From Lookup_Helper sheet
        hours_in_month = 120  # From Lookup_Helper sheet
        hourly_cost_per_head_usd = float((total_cost / headcount / hours_in_month) / usd_inr_rate)
        
        return total_real_estate_cost, total_it_infra_cost, enab_cost, tech_cost, total_cost, hourly_cost_per_head_usd
        
    except Exception as e:
        logger.error(f"Error calculating costs: {str(e)}")
        return None, None, None, None, None, None

# Generate report
def generate_report(headcount, selected_tier, selected_city, selected_plan, 
                   real_estate_toggle, it_infra_toggle, enabling_toggle, technology_toggle,
                   total_real_estate_cost, total_it_infra_cost, enab_cost, tech_cost, 
                   total_cost, hourly_cost_per_head_usd):
    report = f"""
    GCC SETUP COST REPORT
    Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    
    CONFIGURATION:
    - Headcount: {headcount}
    - Tier: {selected_tier}
    - City: {selected_city}
    - Plan: {selected_plan}
    
    COMPONENTS INCLUDED:
    - Real Estate: {'Yes' if real_estate_toggle else 'No'}
    - IT Infrastructure: {'Yes' if it_infra_toggle else 'No'}
    - Enabling Functions: {'Yes' if enabling_toggle else 'No'}
    - Technology: {'Yes' if technology_toggle else 'No'}
    
    COST BREAKDOWN (Monthly):
    - Real Estate: ₹{total_real_estate_cost:,.0f}
    - IT Infrastructure: ₹{total_it_infra_cost:,.0f}
    - Enabling Functions: ₹{enab_cost:,.0f}
    - Technology: ₹{tech_cost:,.0f}
    - TOTAL COST: ₹{total_cost:,.0f}
    
    HOURLY COST PER HEAD: ${hourly_cost_per_head_usd:.6f}
    
    NOTE: All costs include a 30% markup for company services.
    """
    return report

# Plan details function
def get_plan_details(plan):
    if plan == "Basic":
        return {
            "name": "Basic",
            "description": "Essential GCC setup with core functionality",
            "real_estate": "Managed Workspace",
            "it_infra": "Hardware, Networking, Security solutions, Cloud, IT Support, Collaboration tools, Data-centre, Disaster-recovery backups, Compliance & audits, End-user peripherals, cybersecurity (SIEM/DLP), Biometric attendance/access",
            "enabling_functions": "Based on team size: HR/Admin, Finance, IT, Admin, HRBP+Ops",
            "technology": "Platforms like Talent Trail, Zoho People, Freshteam, Keka Foundation, BambooHR, SAP SuccessFactors, Recruiter Fusion, Zoho Books"
        }
    elif plan == "Premium":
        return {
            "name": "Premium",
            "description": "Enhanced GCC setup with additional features",
            "real_estate": "Managed Workspace",
            "it_infra": "Hardware, Networking, Security solutions, Cloud/SaaS, IT Support & AMC, Collaboration tools, Data-centre/Colocation, Disaster-recovery backups, Compliance & audits, End-user peripherals, Advanced cybersecurity (SIEM/DLP), Biometric attendance/access",
            "enabling_functions": "HR/Admin, Finance, IT, Admin, HRBP+Ops, Marketing, Legal, Finance, Other",
            "technology": "Talent Trail, Zoho Campaigns, Contracts, Books, Premium, Keka Foundation, BambooHR, Darwinbox, SAP SuccessFactors, Recruiter Fusion"
        }
    else:  # Advance
        return {
            "name": "Advance",
            "description": "Comprehensive GCC setup with full customization",
            "real_estate": "Managed Workspace",
            "it_infra": "Hardware, Networking, Security solutions, Cloud/SaaS, IT Support & AMC, Collaboration tools, Data-centre/Colocation, Disaster-recovery backups, Compliance & audits, End-user peripherals, Advanced cybersecurity (SIEM/DLP), Biometric attendance/access",
            "enabling_functions": "HR/Admin, Finance, IT, Admin, HRBP+Ops, Marketing, Legal, Finance, Vendor Mgmt, Other",
            "technology": "Talent Trial, HubSpot Pro, DocuSign CLM, QuickBooks Adv, Salesforce Marketing Cloud, Xero, Notion, Marketo Pro, ContractWorks, NetSuite, Oracle Eloqua, Agifo, SAP B1, Slack Grid, Keka Foundation, BambooHR, Darwinbox, SAP SuccessFactors, Recruiter Fusion"
        }

# Helper function to convert numpy types to native Python types
def convert_numpy_types(obj):
    if isinstance(obj, (np.integer, np.floating)):
        return obj.item()
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, dict):
        return {k: convert_numpy_types(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_numpy_types(item) for item in obj]
    else:
        return obj

@app.route('/')
def index():
    # Load data
    real_estate_df, it_infra_df, plans_df, lookup_helper_df = load_data()
    
    # Check if data loaded successfully
    if real_estate_df is None:
        return render_template('error.html', message="Failed to load data. Please check the Excel file and try again.")
    
    # Get unique tiers and cities - ensure no duplicates
    # Use drop_duplicates() to remove duplicate rows before extracting unique tiers
    unique_tiers_df = real_estate_df[['Tier']].drop_duplicates()
    tiers = unique_tiers_df['Tier'].tolist()
    tiers.sort()  # Sort the tiers for consistent ordering
    
    cities_by_tier = {}
    for tier in tiers:
        # Get unique cities for each tier
        cities_df = real_estate_df[real_estate_df["Tier"] == tier][['City']].drop_duplicates()
        cities = cities_df['City'].tolist()
        cities.sort()  # Sort cities alphabetically
        cities_by_tier[tier] = cities
    
    # Get average costs for each tier
    avg_costs = {}
    for tier in tiers:
        avg_real_estate = real_estate_df[real_estate_df["Tier"] == tier]["Cost_INR_PM"].mean()
        avg_it_infra = it_infra_df[it_infra_df["Tier"] == tier]["Cost_INR_PM"].mean()
        avg_costs[tier] = {
            "real_estate": convert_numpy_types(avg_real_estate),
            "it_infra": convert_numpy_types(avg_it_infra)
        }
    
    # Get plan cost ranges
    plan_ranges = {
        "Basic": {
            "min": convert_numpy_types(plans_df['Enab_Basic'].min()),
            "max": convert_numpy_types(plans_df['Enab_Basic'].max())
        },
        "Premium": {
            "min": convert_numpy_types(plans_df['Enab_Premium'].min()),
            "max": convert_numpy_types(plans_df['Enab_Premium'].max())
        },
        "Advance": {
            "min": convert_numpy_types(plans_df['Enab_Advance'].min()),
            "max": convert_numpy_types(plans_df['Enab_Advance'].max())
        }
    }
    
    return render_template('index.html', 
                         tiers=tiers, 
                         cities_by_tier=cities_by_tier,
                         avg_costs=avg_costs,
                         plan_ranges=plan_ranges)

@app.route('/calculate', methods=['POST'])
def calculate():
    # Get form data
    headcount = int(request.form.get('headcount', 100))
    selected_tier = request.form.get('tier', 'Tier 1')
    selected_city = request.form.get('city')
    selected_plan = request.form.get('plan', 'Basic')
    real_estate_toggle = request.form.get('real_estate') == 'on'
    it_infra_toggle = request.form.get('it_infra') == 'on'
    enabling_toggle = request.form.get('enabling') == 'on'
    technology_toggle = request.form.get('technology') == 'on'
    
    # Load data
    real_estate_df, it_infra_df, plans_df, lookup_helper_df = load_data()
    
    # Calculate costs
    total_real_estate_cost, total_it_infra_cost, enab_cost, tech_cost, total_cost, hourly_cost_per_head_usd = calculate_costs(
        headcount, selected_city, selected_plan, real_estate_toggle, 
        it_infra_toggle, enabling_toggle, technology_toggle,
        real_estate_df, it_infra_df, plans_df
    )
    
    if total_cost is None:
        return jsonify({"error": "Failed to calculate costs. Please check your inputs."})
    
    # Get cities by tier for the JavaScript functionality
    # Use drop_duplicates() to remove duplicate rows before extracting unique tiers
    unique_tiers_df = real_estate_df[['Tier']].drop_duplicates()
    tiers = unique_tiers_df['Tier'].tolist()
    tiers.sort()  # Sort the tiers for consistent ordering
    
    cities_by_tier = {}
    for tier in tiers:
        # Get unique cities for each tier
        cities_df = real_estate_df[real_estate_df["Tier"] == tier][['City']].drop_duplicates()
        cities = cities_df['City'].tolist()
        cities.sort()  # Sort cities alphabetically
        cities_by_tier[tier] = cities
    
    # Get plan details
    plan_details = get_plan_details(selected_plan)
    
    # Prepare results
    results = {
        "headcount": headcount,
        "tier": selected_tier,
        "city": selected_city,
        "plan": selected_plan,
        "real_estate_toggle": real_estate_toggle,
        "it_infra_toggle": it_infra_toggle,
        "enabling_toggle": enabling_toggle,
        "technology_toggle": technology_toggle,
        "total_real_estate_cost": total_real_estate_cost,
        "total_it_infra_cost": total_it_infra_cost,
        "enab_cost": enab_cost,
        "tech_cost": tech_cost,
        "total_cost": total_cost,
        "hourly_cost_per_head_usd": hourly_cost_per_head_usd,
        "plan_details": plan_details
    }
    
    # Pass datetime and cities_by_tier to the template context
    return render_template('results.html', results=results, datetime=datetime, cities_by_tier=cities_by_tier)

@app.route('/download_report', methods=['POST'])
def download_report():
    # Get data from form
    data = request.json
    
    # Convert all values in data
    converted_data = {k: convert_numpy_types(v) for k, v in data.items()}
    
    # Generate report
    report = generate_report(
        converted_data['headcount'], converted_data['tier'], converted_data['city'], converted_data['plan'],
        converted_data['real_estate_toggle'], converted_data['it_infra_toggle'], 
        converted_data['enabling_toggle'], converted_data['technology_toggle'],
        converted_data['total_real_estate_cost'], converted_data['total_it_infra_cost'],
        converted_data['enab_cost'], converted_data['tech_cost'], converted_data['total_cost'],
        converted_data['hourly_cost_per_head_usd']
    )
    
    # Create file in memory
    mem = io.BytesIO()
    mem.write(report.encode('utf-8'))
    mem.seek(0)
    
    # Return file
    filename = f"gcc_cost_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    return send_file(mem, as_attachment=True, download_name=filename, mimetype='text/plain')

if __name__ == '__main__':
    app.run(debug=True)