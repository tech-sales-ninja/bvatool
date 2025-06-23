# v1.6 - June 2025 - Added Export/Import Configuration functionality

import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
from datetime import datetime
import csv
from io import StringIO
import json

# Executive Report Dependencies
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    import matplotlib.pyplot as plt
    from io import BytesIO
    REPORT_DEPENDENCIES_AVAILABLE = True
except ImportError:
    REPORT_DEPENDENCIES_AVAILABLE = False

# Set page configuration
st.set_page_config(page_title="Business Value Assessment Tool", layout="wide")

# --- EXPORT/IMPORT FUNCTIONS ---

def get_all_input_values():
    """Collect all input values from the current session state"""
    input_values = {}
    
    # Get all the input keys from your existing inputs
    input_keys = [
        # Basic Configuration
        'solution_name', 'industry_template', 'currency', 
        
        # Implementation Timeline
        'implementation_delay', 'benefits_ramp_up',
        
        # Working Hours Configuration
        'hours_per_day', 'days_per_week', 'weeks_per_year', 'holiday_sick_days',
        
        # Alert Management
        'alert_volume', 'alert_ftes', 'avg_alert_triage_time', 'avg_alert_fte_salary',
        'alert_reduction_pct', 'alert_triage_time_saved_pct',
        
        # Incident Management
        'incident_volume', 'incident_ftes', 'avg_incident_triage_time', 'avg_incident_fte_salary',
        'incident_reduction_pct', 'incident_triage_time_savings_pct',
        
        # Major Incidents
        'major_incident_volume', 'avg_major_incident_cost', 'avg_mttr_hours', 'mttr_improvement_pct',
        
        # Additional Benefits
        'tool_savings', 'people_efficiency', 'fte_avoidance', 'sla_penalty', 
        'revenue_growth', 'capex_savings', 'opex_savings',
        
        # Costs
        'platform_cost', 'services_cost',
        
        # Financial Settings
        'evaluation_years', 'discount_rate'
    ]
    
    # Collect values from session state
    for key in input_keys:
        if key in st.session_state:
            input_values[key] = st.session_state[key]
        else:
            # Fallback to default values if not in session state
            input_values[key] = get_default_value(key)
    
    return input_values

def get_default_value(key):
    """Get default values for inputs"""
    defaults = {
        'solution_name': 'AIOPs',
        'industry_template': 'Custom',
        'currency': '$',
        'implementation_delay': 6,
        'benefits_ramp_up': 3,
        'hours_per_day': 8.0,
        'days_per_week': 5,
        'weeks_per_year': 52,
        'holiday_sick_days': 25,
        'alert_volume': 0,
        'alert_ftes': 0,
        'avg_alert_triage_time': 0,
        'avg_alert_fte_salary': 50000,
        'alert_reduction_pct': 0,
        'alert_triage_time_saved_pct': 0,
        'incident_volume': 0,
        'incident_ftes': 0,
        'avg_incident_triage_time': 0,
        'avg_incident_fte_salary': 50000,
        'incident_reduction_pct': 0,
        'incident_triage_time_savings_pct': 0,
        'major_incident_volume': 0,
        'avg_major_incident_cost': 0,
        'avg_mttr_hours': 0.0,
        'mttr_improvement_pct': 0,
        'tool_savings': 0,
        'people_efficiency': 0,
        'fte_avoidance': 0,
        'sla_penalty': 0,
        'revenue_growth': 0,
        'capex_savings': 0,
        'opex_savings': 0,
        'platform_cost': 0,
        'services_cost': 0,
        'evaluation_years': 3,
        'discount_rate': 10
    }
    return defaults.get(key, 0)

def export_to_csv(input_values):
    """Export input values to CSV format"""
    output = StringIO()
    writer = csv.writer(output)
    
    # Write header
    writer.writerow(['Parameter', 'Value', 'Description'])
    
    # Define parameter descriptions for better readability
    descriptions = {
        'solution_name': 'Solution Name',
        'industry_template': 'Industry Template',
        'currency': 'Currency Symbol',
        'implementation_delay': 'Implementation Delay (months)',
        'benefits_ramp_up': 'Benefits Ramp-up Period (months)',
        'hours_per_day': 'Working Hours per Day',
        'days_per_week': 'Working Days per Week',
        'weeks_per_year': 'Working Weeks per Year',
        'holiday_sick_days': 'Holiday + Sick Days per Year',
        'alert_volume': 'Total Infrastructure Related Alerts per Year',
        'alert_ftes': 'Total FTEs Managing Infrastructure Alerts',
        'avg_alert_triage_time': 'Average Alert Triage Time (minutes)',
        'avg_alert_fte_salary': 'Average Annual Salary per Alert Management FTE',
        'alert_reduction_pct': '% Alert Reduction',
        'alert_triage_time_saved_pct': '% Alert Triage Time Reduction',
        'incident_volume': 'Total Infrastructure Related Incident Volumes per Year',
        'incident_ftes': 'Total FTEs Managing Infrastructure Incidents',
        'avg_incident_triage_time': 'Average Incident Triage Time (minutes)',
        'avg_incident_fte_salary': 'Average Annual Salary per Incident Management FTE',
        'incident_reduction_pct': '% Incident Reduction',
        'incident_triage_time_savings_pct': '% Incident Triage Time Reduction',
        'major_incident_volume': 'Total Infrastructure Related Major Incidents per Year (Sev1)',
        'avg_major_incident_cost': 'Average Major Incident Cost per Hour',
        'avg_mttr_hours': 'Average MTTR (hours)',
        'mttr_improvement_pct': 'MTTR Improvement Percentage',
        'tool_savings': 'Tool Consolidation Savings',
        'people_efficiency': 'People Efficiency Gains',
        'fte_avoidance': 'FTE Avoidance (annualized value)',
        'sla_penalty': 'SLA Penalty Avoidance',
        'revenue_growth': 'Revenue Growth',
        'capex_savings': 'Capital Expenditure Savings',
        'opex_savings': 'Operational Expenditure Savings',
        'platform_cost': 'Annual Subscription Cost',
        'services_cost': 'Implementation & Services (One-Time)',
        'evaluation_years': 'Evaluation Period (Years)',
        'discount_rate': 'NPV Discount Rate (%)'
    }
    
    # Write data rows
    for key, value in input_values.items():
        description = descriptions.get(key, key.replace('_', ' ').title())
        writer.writerow([key, value, description])
    
    return output.getvalue()

def import_from_csv(csv_content):
    """Import input values from CSV content and update session state"""
    try:
        # Parse CSV content
        reader = csv.DictReader(StringIO(csv_content))
        imported_values = {}
        
        for row in reader:
            key = row['Parameter']
            value = row['Value']
            
            # Convert value to appropriate type
            try:
                # Try to convert to number first
                if '.' in str(value):
                    value = float(value)
                else:
                    value = int(value)
            except (ValueError, TypeError):
                # Keep as string if not a number
                value = str(value)
            
            imported_values[key] = value
        
        # Update session state with imported values
        for key, value in imported_values.items():
            st.session_state[key] = value
        
        return True, f"Successfully imported {len(imported_values)} parameters"
    
    except Exception as e:
        return False, f"Error importing CSV: {str(e)}"

def export_to_json(input_values):
    """Export input values to JSON format"""
    # Add metadata
    export_data = {
        'metadata': {
            'export_date': datetime.now().isoformat(),
            'version': '1.6',
            'tool': 'BVA Business Value Assessment'
        },
        'configuration': input_values
    }
    return json.dumps(export_data, indent=2)

def import_from_json(json_content):
    """Import input values from JSON content and update session state"""
    try:
        data = json.loads(json_content)
        
        # Extract configuration data
        if 'configuration' in data:
            imported_values = data['configuration']
        else:
            # Assume the entire JSON is the configuration
            imported_values = data
        
        # Update session state with imported values
        for key, value in imported_values.items():
            st.session_state[key] = value
        
        return True, f"Successfully imported {len(imported_values)} parameters"
    
    except Exception as e:
        return False, f"Error importing JSON: {str(e)}"

# --- Sidebar Export/Import Section ---
st.sidebar.header("🔄 Configuration Export/Import")

# Export Section
with st.sidebar.expander("📤 Export Configuration"):
    st.write("Export your current configuration to save or share with others.")
    
    export_format = st.selectbox("Export Format", ["CSV", "JSON"], key="export_format")
    
    if st.button("Generate Export File"):
        current_values = get_all_input_values()
        
        if export_format == "CSV":
            export_data = export_to_csv(current_values)
            file_extension = "csv"
            mime_type = "text/csv"
        else:  # JSON
            export_data = export_to_json(current_values)
            file_extension = "json"
            mime_type = "application/json"
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"BVA_Config_{timestamp}.{file_extension}"
        
        st.download_button(
            label=f"Download {export_format} Configuration",
            data=export_data,
            file_name=filename,
            mime=mime_type
        )

# Import Section
with st.sidebar.expander("📥 Import Configuration"):
    st.write("Import a previously saved configuration.")
    
    uploaded_file = st.file_uploader(
        "Choose configuration file",
        type=['csv', 'json'],
        help="Upload a CSV or JSON configuration file"
    )
    
    if uploaded_file is not None:
        try:
            # Read file content
            file_content = uploaded_file.read().decode('utf-8')
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            if st.button("Import Configuration"):
                if file_extension == 'csv':
                    success, message = import_from_csv(file_content)
                elif file_extension == 'json':
                    success, message = import_from_json(file_content)
                else:
                    success, message = False, "Unsupported file format"
                
                if success:
                    st.success(message)
                    st.info("Please scroll down to see the imported values. You may need to refresh the page to see all changes.")
                else:
                    st.error(message)
        
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")

st.sidebar.markdown("---")

# --- Sidebar Inputs ---
st.sidebar.header("Customize Your Financial Impact Model Inputs")

# Solution Name Input
solution_name = st.sidebar.text_input("Solution Name", value="AIOPs", key="solution_name")

# --- Implementation Timeline ---
st.sidebar.subheader("📅 Implementation Timeline")
implementation_delay_months = st.sidebar.slider(
    "Implementation Delay (months)", 
    0, 24, 6, 
    help="Time from project start until benefits begin to be realized",
    key="implementation_delay"
)
benefits_ramp_up_months = st.sidebar.slider(
    "Benefits Ramp-up Period (months)", 
    0, 12, 3,
    help="Time to reach full benefits after go-live (gradual adoption)",
    key="benefits_ramp_up"
)

# --- Industry Benchmark Templates ---
industry_templates = {
    "Custom": {},
    "Financial Services": {
        "alert_volume": 1_200_000,
        "major_incident_volume": 140,
        "avg_alert_triage_time": 25,
        "alert_reduction_pct": 40,
        "incident_volume": 400_000,
        "avg_incident_triage_time": 30,
        "incident_reduction_pct": 40,
        "mttr_improvement_pct": 40
    },
    "Retail": {
        "alert_volume": 600_000,
        "major_incident_volume": 80,
        "avg_alert_triage_time": 20,
        "alert_reduction_pct": 30,
        "incident_volume": 200_000,
        "avg_incident_triage_time": 25,
        "incident_reduction_pct": 30,
        "mttr_improvement_pct": 30
    },
    "MSP": {
        "alert_volume": 2_500_000,
        "major_incident_volume": 200,
        "avg_alert_triage_time": 35,
        "alert_reduction_pct": 50,
        "incident_volume": 800_000,
        "avg_incident_triage_time": 35,
        "incident_reduction_pct": 50,
        "mttr_improvement_pct": 50
    },
    "Healthcare": {
        "alert_volume": 800_000,
        "major_incident_volume": 100,
        "avg_alert_triage_time": 30,
        "alert_reduction_pct": 35,
        "incident_volume": 300_000,
        "avg_incident_triage_time": 30,
        "incident_reduction_pct": 35,
        "mttr_improvement_pct": 35
    },
    "Telecom": {
        "alert_volume": 1_800_000,
        "major_incident_volume": 160,
        "avg_alert_triage_time": 35,
        "alert_reduction_pct": 45,
        "incident_volume": 600_000,
        "avg_incident_triage_time": 35,
        "incident_reduction_pct": 40,
        "mttr_improvement_pct": 45
    }
}

selected_template = st.sidebar.selectbox("Select Industry Template", list(industry_templates.keys()), key="industry_template")
template = industry_templates[selected_template]
st.sidebar.caption("📌 Industry templates provide baseline values for estimation only. Adjust any field as needed.")

# --- Currency Selection ---
currency_symbol = st.sidebar.selectbox("Currency", ["$", "€", "£", "Kč"], key="currency")

# --- Working Hours Configuration ---
st.sidebar.subheader("⏰ Working Hours Configuration")
hours_per_day = st.sidebar.number_input(
    "Working Hours per Day", 
    value=8.0, 
    min_value=1.0, 
    max_value=24.0,
    step=0.5,
    key="hours_per_day",
    help="Standard working hours per day for your FTEs"
)
days_per_week = st.sidebar.number_input(
    "Working Days per Week", 
    value=5, 
    min_value=1, 
    max_value=7,
    key="days_per_week",
    help="Standard working days per week"
)
weeks_per_year = st.sidebar.number_input(
    "Working Weeks per Year", 
    value=52, 
    min_value=1, 
    max_value=52,
    key="weeks_per_year",
    help="Total weeks worked per year"
)
holiday_sick_days = st.sidebar.number_input(
    "Holiday + Sick Days per Year", 
    value=25, 
    min_value=0, 
    max_value=100,
    key="holiday_sick_days",
    help="Total days off per year (holidays, vacation, sick leave)"
)

# Calculate and display total working hours
total_working_days = (weeks_per_year * days_per_week) - holiday_sick_days
working_hours_per_fte_per_year = total_working_days * hours_per_day
st.sidebar.info(f"**Calculated: {working_hours_per_fte_per_year:,.0f} working hours per FTE per year**")

# --- ALERT INPUTS ---
st.sidebar.subheader("🚨 Alert Management")
alert_volume = st.sidebar.number_input(
    "Total Infrastructure Related Alerts Managed per Year", 
    value=template.get("alert_volume", 0),
    key="alert_volume"
)
alert_ftes = st.sidebar.number_input(
    "Total FTEs Managing Infrastructure Alerts", 
    value=0,
    key="alert_ftes"
)
avg_alert_triage_time = st.sidebar.number_input(
    "Average Alert Triage Time (minutes)", 
    value=template.get("avg_alert_triage_time", 0),
    key="avg_alert_triage_time"
)
avg_alert_fte_salary = st.sidebar.number_input(
    "Average Annual Salary per Alert Management FTE", 
    value=50000,
    key="avg_alert_fte_salary"
)
alert_reduction_pct = st.sidebar.slider(
    "% Alert Reduction", 
    0, 100, 
    value=template.get("alert_reduction_pct", 0),
    key="alert_reduction_pct"
)
alert_triage_time_saved_pct = st.sidebar.slider(
    "% Alert Triage Time Reduction", 
    0, 100, 0,
    key="alert_triage_time_saved_pct"
)

# --- INCIDENT INPUTS ---
st.sidebar.subheader("🔧 Incident Management")
incident_volume = st.sidebar.number_input(
    "Total Infrastructure Related Incident Volumes Managed per Year", 
    value=template.get("incident_volume", 0),
    key="incident_volume"
)
incident_ftes = st.sidebar.number_input(
    "Total FTEs Managing Infrastructure Incidents", 
    value=0,
    key="incident_ftes"
)
avg_incident_triage_time = st.sidebar.number_input(
    "Average Incident Triage Time (minutes)", 
    value=template.get("avg_incident_triage_time", 0),
    key="avg_incident_triage_time"
)
avg_incident_fte_salary = st.sidebar.number_input(
    "Average Annual Salary per Incident Management FTE", 
    value=50000,
    key="avg_incident_fte_salary"
)
incident_reduction_pct = st.sidebar.slider(
    "% Incident Reduction", 
    0, 100, 
    value=template.get("incident_reduction_pct", 0),
    key="incident_reduction_pct"
)
incident_triage_time_savings_pct = st.sidebar.slider(
    "% Incident Triage Time Reduction", 
    0, 100, 0,
    key="incident_triage_time_savings_pct"
)

# --- MAJOR INCIDENT INPUTS ---
st.sidebar.subheader("🚨 Major Incidents (Sev1)")
major_incident_volume = st.sidebar.number_input(
    "Total Infrastructure Related Major Incidents per Year (Sev1)", 
    value=template.get("major_incident_volume", 0),
    key="major_incident_volume"
)
avg_major_incident_cost = st.sidebar.number_input(
    "Average Major Incident Cost per Hour", 
    value=0,
    key="avg_major_incident_cost"
)
avg_mttr_hours = st.sidebar.number_input(
    "Average MTTR (hours)", 
    value=0.0,
    key="avg_mttr_hours"
)
mttr_improvement_pct = st.sidebar.slider(
    "MTTR Improvement Percentage", 
    0, 100, 
    value=template.get("mttr_improvement_pct", 0),
    key="mttr_improvement_pct"
)

# --- OTHER BENEFITS ---
st.sidebar.subheader("💰 Additional Benefits")
tool_savings = st.sidebar.number_input(
    "Tool Consolidation Savings", 
    value=0,
    key="tool_savings"
)
people_cost_per_year = st.sidebar.number_input(
    "People Efficiency Gains", 
    value=0,
    key="people_efficiency"
)
fte_avoidance = st.sidebar.number_input(
    "FTE Avoidance (annualized value in local currency)", 
    value=0,
    key="fte_avoidance"
)
sla_penalty_avoidance = st.sidebar.number_input(
    "SLA Penalty Avoidance (Service Providers)", 
    value=0,
    key="sla_penalty"
)
revenue_growth = st.sidebar.number_input(
    "Revenue Growth (Service Providers)", 
    value=0,
    key="revenue_growth"
)
capex_savings = st.sidebar.number_input(
    "Capital Expenditure Savings (Hardware)", 
    value=0,
    key="capex_savings"
)
opex_savings = st.sidebar.number_input(
    "Operational Expenditure Savings (e.g. Storage Costs)", 
    value=0,
    key="opex_savings"
)

# --- COSTS ---
st.sidebar.subheader("💳 Solution Costs")
platform_cost = st.sidebar.number_input(
    "Annual Subscription Cost (After discounts)", 
    value=0,
    key="platform_cost"
)
services_cost = st.sidebar.number_input(
    "Implementation & Services (One-Time)", 
    value=0,
    key="services_cost"
)

# --- FINANCIAL SETTINGS ---
st.sidebar.subheader("📊 Financial Analysis Settings")
evaluation_years = st.sidebar.slider(
    "Evaluation Period (Years)", 
    1, 5, 3,
    key="evaluation_years"
)
discount_rate = st.sidebar.slider(
    "NPV Discount Rate (%)", 
    0, 20, 10,
    key="discount_rate"
) / 100

# --- CORRECTED CALCULATIONS WITH CONFIGURABLE WORKING HOURS ---

# Function to calculate alert costs based on FTE time allocation
def calculate_alert_costs(alert_volume, alert_ftes, avg_alert_triage_time, avg_salary_per_year, 
                         hours_per_day, days_per_week, weeks_per_year, holiday_sick_days):
    """Calculate the true cost per alert based on FTE time allocation"""
    if alert_volume == 0 or alert_ftes == 0:
        return 0, 0, 0, 0
    
    total_alert_time_minutes_per_year = alert_volume * avg_alert_triage_time
    total_alert_time_hours_per_year = total_alert_time_minutes_per_year / 60
    
    total_working_days = (weeks_per_year * days_per_week) - holiday_sick_days
    working_hours_per_fte_per_year = total_working_days * hours_per_day
    total_available_fte_hours = alert_ftes * working_hours_per_fte_per_year
    
    fte_time_percentage_on_alerts = total_alert_time_hours_per_year / total_available_fte_hours if total_available_fte_hours > 0 else 0
    
    total_fte_cost = alert_ftes * avg_salary_per_year
    total_alert_handling_cost = total_fte_cost * fte_time_percentage_on_alerts
    cost_per_alert = total_alert_handling_cost / alert_volume if alert_volume > 0 else 0
    
    return cost_per_alert, total_alert_handling_cost, fte_time_percentage_on_alerts, working_hours_per_fte_per_year

# Function to calculate incident costs based on FTE time allocation
def calculate_incident_costs(incident_volume, incident_ftes, avg_incident_triage_time, avg_salary_per_year,
                           hours_per_day, days_per_week, weeks_per_year, holiday_sick_days):
    """Calculate the true cost per incident based on FTE time allocation"""
    if incident_volume == 0 or incident_ftes == 0:
        return 0, 0, 0, 0
    
    total_incident_time_minutes_per_year = incident_volume * avg_incident_triage_time
    total_incident_time_hours_per_year = total_incident_time_minutes_per_year / 60
    
    total_working_days = (weeks_per_year * days_per_week) - holiday_sick_days
    working_hours_per_fte_per_year = total_working_days * hours_per_day
    total_available_fte_hours = incident_ftes * working_hours_per_fte_per_year
    
    fte_time_percentage_on_incidents = total_incident_time_hours_per_year / total_available_fte_hours if total_available_fte_hours > 0 else 0
    
    total_fte_cost = incident_ftes * avg_salary_per_year
    total_incident_handling_cost = total_fte_cost * fte_time_percentage_on_incidents
    cost_per_incident = total_incident_handling_cost / incident_volume if incident_volume > 0 else 0
    
    return cost_per_incident, total_incident_handling_cost, fte_time_percentage_on_incidents, working_hours_per_fte_per_year

# Calculate alert and incident costs
cost_per_alert, total_alert_handling_cost, alert_fte_percentage, alert_working_hours = calculate_alert_costs(
    alert_volume, alert_ftes, avg_alert_triage_time, avg_alert_fte_salary,
    hours_per_day, days_per_week, weeks_per_year, holiday_sick_days
)

cost_per_incident, total_incident_handling_cost, incident_fte_percentage, incident_working_hours = calculate_incident_costs(
    incident_volume, incident_ftes, avg_incident_triage_time, avg_incident_fte_salary,
    hours_per_day, days_per_week, weeks_per_year, holiday_sick_days
)

# Calculate baseline savings
avoided_alerts = alert_volume * (alert_reduction_pct / 100)
remaining_alerts = alert_volume - avoided_alerts
alert_reduction_savings = avoided_alerts * cost_per_alert
remaining_alert_handling_cost = remaining_alerts * cost_per_alert
alert_triage_savings = remaining_alert_handling_cost * (alert_triage_time_saved_pct / 100)

avoided_incidents = incident_volume * (incident_reduction_pct / 100)
remaining_incidents = incident_volume - avoided_incidents
incident_reduction_savings = avoided_incidents * cost_per_incident
remaining_incident_handling_cost = remaining_incidents * cost_per_incident
incident_triage_savings = remaining_incident_handling_cost * (incident_triage_time_savings_pct / 100)

mttr_hours_saved_per_incident = (mttr_improvement_pct / 100) * avg_mttr_hours
total_mttr_hours_saved = major_incident_volume * mttr_hours_saved_per_incident
major_incident_savings = total_mttr_hours_saved * avg_major_incident_cost

# Total Annual Benefits (baseline)
total_annual_benefits = (
    alert_reduction_savings + alert_triage_savings + incident_reduction_savings +
    incident_triage_savings + major_incident_savings + tool_savings + people_cost_per_year +
    fte_avoidance + sla_penalty_avoidance + revenue_growth + capex_savings + opex_savings
)

# --- Implementation Delay Functions ---
def calculate_benefit_realization_factor(month, implementation_delay_months, ramp_up_months):
    """Calculate what percentage of benefits are realized in a given month"""
    if month <= implementation_delay_months:
        return 0.0  # No benefits during implementation
    elif month <= implementation_delay_months + ramp_up_months:
        # Linear ramp-up during ramp-up period
        months_since_golive = month - implementation_delay_months
        return months_since_golive / ramp_up_months
    else:
        return 1.0  # Full benefits realized

def calculate_scenario_results(benefits_multiplier, implementation_delay_multiplier, scenario_name):
    """Calculate NPV, ROI, and payback for a given scenario"""
    # Adjust benefits and timeline
    scenario_benefits = total_annual_benefits * benefits_multiplier
    scenario_impl_delay = max(0, int(implementation_delay_months * implementation_delay_multiplier)) # Ensure not negative
    scenario_ramp_up = benefits_ramp_up_months
    
    # Calculate cash flows
    scenario_cash_flows = []
    for year in range(1, evaluation_years + 1):
        year_start_month = (year - 1) * 12 + 1
        year_end_month = year * 12
        
        monthly_factors = []
        for month in range(year_start_month, year_end_month + 1):
            factor = calculate_benefit_realization_factor(month, scenario_impl_delay, scenario_ramp_up)
            monthly_factors.append(factor)
        
        avg_realization_factor = np.mean(monthly_factors)
        year_benefits = scenario_benefits * avg_realization_factor
        year_platform_cost = platform_cost
        year_services_cost = services_cost if year == 1 else 0
        year_net_cash_flow = year_benefits - year_platform_cost - year_services_cost
        
        scenario_cash_flows.append({
            'year': year,
            'benefits': year_benefits,
            'platform_cost': year_platform_cost,
            'services_cost': year_services_cost,
            'net_cash_flow': year_net_cash_flow,
            'realization_factor': avg_realization_factor
        })
    
    # Calculate metrics
    scenario_npv = sum([cf['net_cash_flow'] / ((1 + discount_rate) ** cf['year']) for cf in scenario_cash_flows])
    scenario_tco = sum([cf['platform_cost'] + cf['services_cost'] for cf in scenario_cash_flows])
    scenario_roi = scenario_npv / scenario_tco if scenario_tco != 0 else 0
    
    # Calculate payback
    scenario_payback = "N/A"
    cumulative_net_cash_flow = 0
    for cf in scenario_cash_flows:
        cumulative_net_cash_flow += cf['net_cash_flow']
        if cumulative_net_cash_flow >= 0:
            scenario_payback = f"{cf['year']} years"
            break
    
    return {
        'npv': scenario_npv,
        'roi': scenario_roi,
        'payback': scenario_payback,
        'impl_delay': scenario_impl_delay,
        'benefits_mult': benefits_multiplier,
        'cash_flows': scenario_cash_flows,
        'annual_benefits': scenario_benefits
    }

# Calculate scenarios
scenarios = {
    "Conservative": {
        "benefits_multiplier": 0.7,  # 30% lower benefits
        "implementation_delay_multiplier": 1.3,  # 30% longer implementation
        "description": "Benefits 30% lower, implementation 30% longer",
        "color": "#ff6b6b",
        "icon": "🔴"
    },
    "Expected": {
        "benefits_multiplier": 1.0,  # Baseline
        "implementation_delay_multiplier": 1.0,  # Baseline
        "description": "Baseline assumptions as entered",
        "color": "#4ecdc4",
        "icon": "🟢"
    },
    "Optimistic": {
        "benefits_multiplier": 1.2,  # 20% higher benefits
        "implementation_delay_multiplier": 0.8,  # 20% faster implementation
        "description": "Benefits 20% higher, implementation 20% faster",
        "color": "#45b7d1",
        "icon": "🔵"
    }
}

scenario_results = {}
for scenario_name, params in scenarios.items():
    scenario_results[scenario_name] = calculate_scenario_results(
        params["benefits_multiplier"], 
        params["implementation_delay_multiplier"],
        scenario_name
    )
    scenario_results[scenario_name].update({
        "color": params["color"],
        "description": params["description"],
        "icon": params["icon"]
    })

# --- NEW FUNCTIONALITY ADDITIONS ---

# 1. Calculate the total cost savings from alert and incident management
total_operational_savings_from_time_saved = (
    alert_reduction_savings + alert_triage_savings +
    incident_reduction_savings + incident_triage_savings +
    major_incident_savings
)

# 2. Determine the equivalent number of full-time employees (FTEs) from savings
effective_avg_fte_salary = 0
if avg_alert_fte_salary > 0 and avg_incident_fte_salary > 0:
    effective_avg_fte_salary = (avg_alert_fte_salary + avg_incident_fte_salary) / 2
elif avg_alert_fte_salary > 0:
    effective_avg_fte_salary = avg_alert_fte_salary
elif avg_incident_fte_salary > 0:
    effective_avg_fte_salary = avg_incident_fte_salary

equivalent_ftes_from_savings = 0
if effective_avg_fte_salary > 0:
    equivalent_ftes_from_savings = total_operational_savings_from_time_saved / effective_avg_fte_salary

# 3. Payback Periods in Months (More granular calculation)

def calculate_payback_months(annual_benefits, annual_platform_cost, one_time_services_cost, 
                             implementation_delay_months, benefits_ramp_up_months, max_months_eval=60):
    """Calculates the payback period in months."""
    
    cumulative_cash_flow = 0
    payback_month = "N/A"
    
    # Initial investment (services cost) incurred at the beginning
    cumulative_cash_flow -= one_time_services_cost

    for month in range(1, max_months_eval + 1):
        factor = calculate_benefit_realization_factor(month, implementation_delay_months, benefits_ramp_up_months)
        
        monthly_benefit = (annual_benefits / 12) * factor
        monthly_platform_cost = annual_platform_cost / 12
        
        monthly_net_cash_flow = monthly_benefit - monthly_platform_cost
        
        cumulative_cash_flow += monthly_net_cash_flow
        
        if cumulative_cash_flow >= 0:
            payback_month = f"{month} months"
            break
            
    return payback_month

# Update scenario results with monthly payback for each scenario
for scenario_name, params in scenarios.items():
    s_result = scenario_results[scenario_name]
    
    scenario_annual_benefits_for_payback = total_annual_benefits * params["benefits_multiplier"]
    scenario_impl_delay_for_payback = max(0, int(implementation_delay_months * params["implementation_delay_multiplier"]))
    
    s_result['payback_months'] = calculate_payback_months(
        annual_benefits=scenario_annual_benefits_for_payback,
        annual_platform_cost=platform_cost,
        one_time_services_cost=services_cost,
        implementation_delay_months=scenario_impl_delay_for_payback,
        benefits_ramp_up_months=benefits_ramp_up_months,
        max_months_eval=evaluation_years * 12
    )

# --- END OF NEW FUNCTIONALITY ADDITIONS ---


def create_implementation_timeline_chart(implementation_delay_months, ramp_up_months, evaluation_years, currency_symbol, total_annual_benefits):
    """Create a visual timeline showing benefit realization over time"""
    
    total_months = evaluation_years * 12
    months = list(range(1, total_months + 1))
    realization_factors = []
    monthly_benefits = []
    
    for month in months:
        factor = calculate_benefit_realization_factor(month, implementation_delay_months, ramp_up_months)
        realization_factors.append(factor * 100)
        monthly_benefits.append(total_annual_benefits * factor / 12)
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=months, y=realization_factors, mode='lines+markers', name='Benefit Realization %',
        line=dict(color='#2E86AB', width=3), marker=dict(size=4),
        hovertemplate='<b>Month %{x}</b><br>Benefit Realization: %{y:.1f}%<br><extra></extra>',
        yaxis='y'
    ))
    
    fig.add_trace(go.Scatter(
        x=months, y=[b/1000 for b in monthly_benefits], mode='lines', name=f'Monthly Benefits ({currency_symbol}K)',
        line=dict(color='#A23B72', width=2), fill='tonexty', fillcolor='rgba(162, 59, 114, 0.2)',
        hovertemplate='<b>Month %{x}</b><br>' + f'Monthly Benefit: {currency_symbol}' + '%{customdata:,.0f}<br><extra></extra>',
        customdata=monthly_benefits, yaxis='y2'
    ))
    
    if implementation_delay_months > 0:
        fig.add_vline(x=implementation_delay_months, line_dash="dash", line_color="red", line_width=2,
                      annotation_text="Go-Live", annotation_position="top",
                      annotation=dict(bgcolor="white", bordercolor="red"))
    
    if ramp_up_months > 0:
        fig.add_vline(x=implementation_delay_months + ramp_up_months, line_dash="dash", line_color="green", line_width=2,
                      annotation_text="Full Benefits", annotation_position="top",
                      annotation=dict(bgcolor="white", bordercolor="green"))
    
    if implementation_delay_months > 0:
        fig.add_vrect(x0=0, x1=implementation_delay_months, fillcolor="red", opacity=0.1, layer="below", line_width=0,
                      annotation_text="Implementation Phase", annotation_position="top left",
                      annotation=dict(textangle=0, font=dict(size=10, color="red")))
    
    if ramp_up_months > 0:
        fig.add_vrect(x0=implementation_delay_months, x1=implementation_delay_months + ramp_up_months,
                      fillcolor="orange", opacity=0.1, layer="below", line_width=0,
                      annotation_text="Ramp-up Phase", annotation_position="top left",
                      annotation=dict(textangle=0, font=dict(size=10, color="orange")))
    
    fig.add_vrect(x0=implementation_delay_months + ramp_up_months, x1=total_months,
                  fillcolor="green", opacity=0.1, layer="below", line_width=0,
                  annotation_text="Full Benefits Phase", annotation_position="top left",
                  annotation=dict(textangle=0, font=dict(size=10, color="green")))
    
    fig.update_layout(
        title={'text': 'Implementation Timeline & Benefit Realization', 'x': 0.5, 'xanchor': 'center', 'font': {'size': 18}},
        xaxis_title="Months from Project Start",
        yaxis=dict(title="Benefit Realization (%)", side="left", range=[0, 105], color='#2E86AB'),
        yaxis2=dict(title=f"Monthly Benefits ({currency_symbol}K)", side="right", overlaying="y", color='#A23B72'),
        height=500, hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=80), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
    )
    
    fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='lightgray')
    
    return fig

# --- EXECUTIVE REPORT GENERATOR FUNCTIONS ---

def create_executive_summary_data(scenario_results, currency_symbol):
    """Create data structure for executive summary"""
    return {
        'investment_summary': {
            'conservative_npv': scenario_results['Conservative']['npv'],
            'expected_npv': scenario_results['Expected']['npv'],
            'optimistic_npv': scenario_results['Optimistic']['npv'],
            'expected_roi': scenario_results['Expected']['roi'],
            'payback_period': scenario_results['Expected']['payback'],
            'currency': currency_symbol,
            'expected_payback_months': scenario_results['Expected']['payback_months']
        },
        'key_benefits': {
            'alert_reduction_savings': alert_reduction_savings,
            'alert_triage_savings': alert_triage_savings,
            'incident_reduction_savings': incident_reduction_savings,
            'incident_triage_savings': incident_triage_savings,
            'major_incident_savings': major_incident_savings,
            'total_operational_savings': alert_reduction_savings + alert_triage_savings + incident_reduction_savings + incident_triage_savings + major_incident_savings,
            'additional_benefits': tool_savings + people_cost_per_year + fte_avoidance + revenue_growth
        },
        'implementation': {
            'delay_months': implementation_delay_months,
            'ramp_up_months': benefits_ramp_up_months,
            'full_benefits_month': implementation_delay_months + benefits_ramp_up_months,
            'evaluation_years': evaluation_years
        },
        'reallocation_and_fte': {
            'total_cost_savings_for_reallocation': total_operational_savings_from_time_saved,
            'equivalent_ftes_from_savings': equivalent_ftes_from_savings
        }
    }

def create_timeline_chart_for_pdf():
    """Create implementation timeline chart for PDF"""
    if not REPORT_DEPENDENCIES_AVAILABLE:
        return None
        
    fig, ax = plt.subplots(figsize=(10, 4))
    
    # Timeline data
    phases = ['Implementation', 'Ramp-up', 'Full Benefits']
    starts = [0, implementation_delay_months, implementation_delay_months + benefits_ramp_up_months]
    durations = [implementation_delay_months, benefits_ramp_up_months, max(0, (evaluation_years*12) - (implementation_delay_months + benefits_ramp_up_months))]
    colors_list = ['#ff6b6b', '#ffa500', '#4ecdc4']
    
    # Create Gantt chart
    for i, (phase, start, duration, color) in enumerate(zip(phases, starts, durations, colors_list)):
        if duration > 0:
            ax.barh(i, duration, left=start, height=0.6, color=color, alpha=0.7, label=phase)
            ax.text(start + duration/2, i, phase, ha='center', va='center', fontweight='bold', fontsize=10)
    
    ax.set_ylim(-0.5, len(phases) - 0.5)
    ax.set_xlim(0, evaluation_years * 12)
    ax.set_xlabel('Months from Project Start', fontsize=12)
    ax.set_title('Implementation Timeline & Benefit Realization', fontsize=14, fontweight='bold')
    ax.grid(axis='x', alpha=0.3)
    
    # Remove y-axis labels
    ax.set_yticks([])
    
    plt.tight_layout()
    
    # Save to BytesIO
    img_buffer = BytesIO()
    plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
    img_buffer.seek(0)
    plt.close()
    
    return img_buffer

def create_scenario_chart_for_pdf(scenario_results, currency_symbol):
    """Create scenario comparison chart for PDF"""
    if not REPORT_DEPENDENCIES_AVAILABLE:
        return None
        
    fig, ax = plt.subplots(figsize=(10, 6))
    
    scenarios_list = list(scenario_results.keys())
    npvs = [scenario_results[scenario]['npv'] for scenario in scenarios_list]
    colors_list = ['#ff6b6b', '#4ecdc4', '#45b7d1']
    
    bars = ax.bar(scenarios_list, npvs, color=colors_list, alpha=0.8)
    
    # Add value labels on bars
    for bar, npv in zip(bars, npvs):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + max(npvs)*0.01,
                f'{currency_symbol}{npv:,.0f}', ha='center', va='bottom', fontweight='bold')
    
    ax.set_ylabel(f'Net Present Value ({currency_symbol})', fontsize=12)
    ax.set_title('Scenario Analysis - NPV Comparison', fontsize=14, fontweight='bold')
    ax.grid(axis='y', alpha=0.3)
    
    # Format y-axis
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{currency_symbol}{x/1000:.0f}K'))
    
    plt.tight_layout()
    
    # Save to BytesIO
    img_buffer = BytesIO()
    plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
    img_buffer.seek(0)
    plt.close()
    
    return img_buffer

def generate_executive_report_pdf(summary_data, scenario_results, solution_name, organization_name="Your Organization"):
    """Generate comprehensive executive report PDF"""
    
    if not REPORT_DEPENDENCIES_AVAILABLE:
        return None
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.darkblue,
        borderWidth=1,
        borderColor=colors.darkblue,
        borderPadding=5
    )
    
    subheading_style = ParagraphStyle(
        'CustomSubheading',
        parent=styles['Heading3'],
        fontSize=14,
        spaceAfter=8,
        textColor=colors.darkblue
    )
    
    # Title Page
    story.append(Paragraph(f"Business Value Assessment", title_style))
    story.append(Paragraph(f"{solution_name} Implementation", styles['Title']))
    story.append(Spacer(1, 0.5*inch))
    story.append(Paragraph(f"Prepared for: {organization_name}", styles['Heading2']))
    story.append(Paragraph(f"Date: {datetime.now().strftime('%B %d, %Y')}", styles['Normal']))
    story.append(Spacer(1, 0.5*inch))
    
    # Custom style for white header text
    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        textColor=colors.white,
        fontName='Helvetica-Bold'
    )
    
    # Executive Summary Box with proper column widths and white headers
    exec_summary_data = [
        [Paragraph('<b>Metric</b>', header_style), 
         Paragraph('<b>Conservative</b>', header_style), 
         Paragraph('<b>Expected</b>', header_style), 
         Paragraph('<b>Optimistic</b>', header_style)],
        [Paragraph('Net Present Value', styles['Normal']), 
         Paragraph(f"{summary_data['investment_summary']['currency']}{scenario_results['Conservative']['npv']:,.0f}", styles['Normal']),
         Paragraph(f"{summary_data['investment_summary']['currency']}{scenario_results['Expected']['npv']:,.0f}", styles['Normal']),
         Paragraph(f"{summary_data['investment_summary']['currency']}{scenario_results['Optimistic']['npv']:,.0f}", styles['Normal'])],
        [Paragraph('Return on Investment', styles['Normal']),
         Paragraph(f"{scenario_results['Conservative']['roi']*100:.1f}%", styles['Normal']),
         Paragraph(f"{scenario_results['Expected']['roi']*100:.1f}%", styles['Normal']),
         Paragraph(f"{scenario_results['Optimistic']['roi']*100:.1f}%", styles['Normal'])],
        [Paragraph('Payback Period (Years)', styles['Normal']),
         Paragraph(scenario_results['Conservative']['payback'], styles['Normal']),
         Paragraph(scenario_results['Expected']['payback'], styles['Normal']),
         Paragraph(scenario_results['Optimistic']['payback'], styles['Normal'])],
        [Paragraph('Payback Period (Months)', styles['Normal']),
         Paragraph(scenario_results['Conservative']['payback_months'], styles['Normal']),
         Paragraph(scenario_results['Expected']['payback_months'], styles['Normal']),
         Paragraph(scenario_results['Optimistic']['payback_months'], styles['Normal'])]
    ]
    
    exec_table = Table(exec_summary_data, colWidths=[2.2*inch, 1.5*inch, 1.5*inch, 1.5*inch])
    exec_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    story.append(exec_table)
    story.append(Spacer(1, 0.2*inch))
    
    # Add overall cost reallocation and FTE equivalency to Executive Summary (optional, can be its own section)
    story.append(Paragraph("Operational Savings for Reallocation:", subheading_style))
    story.append(Paragraph(
        f"Annually, **{summary_data['investment_summary']['currency']}{summary_data['reallocation_and_fte']['total_cost_savings_for_reallocation']:,.0f}** can be reallocated to higher-margin projects. "
        f"This represents **{summary_data['reallocation_and_fte']['equivalent_ftes_from_savings']:,.1f}** equivalent full-time employees (FTEs) in savings.",
        styles['Normal']
    ))
    story.append(PageBreak())
    
    # 1. Executive Summary
    story.append(Paragraph("Executive Summary", heading_style))
    
    exec_text = f"""
    This Business Value Assessment demonstrates the financial and operational benefits of implementing {solution_name} at {organization_name}. Our analysis shows strong positive returns across all scenarios: 
    <b>Key Financial Highlights:</b><br/> 
    • Expected NPV: {summary_data['investment_summary']['currency']}{summary_data['investment_summary']['expected_npv']:,.0f}<br/> 
    • Expected ROI: {summary_data['investment_summary']['expected_roi']*100:.1f}%<br/> 
    • Payback Period: {summary_data['investment_summary']['payback_period']} ({summary_data['investment_summary']['expected_payback_months']})<br/> 
    • NPV Range: {summary_data['investment_summary']['currency']}{scenario_results['Conservative']['npv']:,.0f} to {summary_data['investment_summary']['currency']}{scenario_results['Optimistic']['npv']:,.0f}<br/><br/> 
    <b>Primary Value Drivers:</b><br/> 
    • Alert Management Optimization: {summary_data['investment_summary']['currency']}{summary_data['key_benefits']['alert_reduction_savings'] + summary_data['key_benefits']['incident_reduction_savings']:,.0f} annually<br/> 
    • Incident Management Efficiency: {summary_data['investment_summary']['currency']}{summary_data['key_benefits']['incident_reduction_savings'] + summary_data['key_benefits']['incident_triage_savings']:,.0f} annually<br/> 
    • Major Incident Impact Reduction: {summary_data['investment_summary']['currency']}{summary_data['key_benefits']['major_incident_savings']:,.0f} annually<br/><br/> 
    <b>Operational Savings for Reallocation:</b><br/>
    • Total Annual Cost Savings from A&I Management: {summary_data['investment_summary']['currency']}{summary_data['reallocation_and_fte']['total_cost_savings_for_reallocation']:,.0f}<br/>
    • Equivalent FTEs from Savings: {summary_data['reallocation_and_fte']['equivalent_ftes_from_savings']:,.1f} FTEs<br/><br/>
    <b>Implementation Timeline:</b><br/> 
    • Implementation Phase: {summary_data['implementation']['delay_months']} months<br/> 
    • Ramp-up to Full Benefits: {summary_data['implementation']['ramp_up_months']} months<br/> 
    • Full ROI Realization: Month {summary_data['implementation']['full_benefits_month']}<br/><br/> 
    Even under conservative assumptions (30% lower benefits, 30% longer implementation), the investment delivers **{scenario_results['Conservative']['roi']*100:.1f}% ROI** with a **{scenario_results['Conservative']['payback_months']}** payback period. 
    """ 
    story.append(Paragraph(exec_text, styles['Normal'])) 
    story.append(Spacer(1, 0.3*inch)) 
    # Add scenario chart 
    scenario_chart = create_scenario_chart_for_pdf(scenario_results, summary_data['investment_summary']['currency']) 
    if scenario_chart: 
        story.append(Image(scenario_chart, width=6*inch, height=3.6*inch)) 
    story.append(PageBreak()) 
    # 2. Implementation Roadmap with wrapped text and white headers 
    story.append(Paragraph("Implementation Roadmap & Milestones", heading_style)) 
    roadmap_data = [ 
        [Paragraph('<b>Phase</b>', header_style), Paragraph('<b>Duration</b>', header_style), Paragraph('<b>Key Activities</b>', header_style), Paragraph('<b>Benefits Realization</b>', header_style)], 
        [Paragraph('Planning & Setup', styles['Normal']), Paragraph(f"Months 1-2", styles['Normal']), Paragraph('Environment setup, integration planning, team training', styles['Normal']), Paragraph('0%', styles['Normal'])], 
        [Paragraph('Core Implementation', styles['Normal']), Paragraph(f"Months 3-{implementation_delay_months}", styles['Normal']), Paragraph('Data integration, alert configuration, dashboard creation', styles['Normal']), Paragraph('0%', styles['Normal'])], 
        [Paragraph('Go-Live & Ramp-up', styles['Normal']), Paragraph(f"Months {implementation_delay_months+1}-{implementation_delay_months + benefits_ramp_up_months}", styles['Normal']), Paragraph('Deployment, user adoption, process optimization', styles['Normal']), Paragraph('0% → 100%', styles['Normal'])], 
        [Paragraph('Full Operation', styles['Normal']), Paragraph(f"Month {implementation_delay_months + benefits_ramp_up_months}+", styles['Normal']), Paragraph('Business as usual, continuous improvement', styles['Normal']), Paragraph('100%', styles['Normal'])], 
    ] 
    roadmap_table = Table(roadmap_data, colWidths=[1.3*inch, 1.1*inch, 3*inch, 1.3*inch]) 
    roadmap_table.setStyle(TableStyle([ 
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue), 
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), 
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'), 
        ('VALIGN', (0, 0), (-1, -1), 'TOP'), 
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), 
        ('FONTSIZE', (0, 0), (-1, -1), 9), 
        ('TOPPADDING', (0, 0), (-1, -1), 8), 
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8), 
        ('LEFTPADDING', (0, 0), (-1, -1), 6), 
        ('RIGHTPADDING', (0, 0), (-1, -1), 6), 
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]), 
        ('GRID', (0, 0), (-1, -1), 1, colors.black) 
    ])) 
    story.append(roadmap_table) 
    story.append(Spacer(1, 0.3*inch)) 
    # Add timeline chart 
    timeline_chart = create_timeline_chart_for_pdf() 
    if timeline_chart: 
        story.append(Image(timeline_chart, width=6*inch, height=2.4*inch)) 
    story.append(Spacer(1, 0.3*inch)) 
    # Key Milestones 
    story.append(Paragraph("Key Success Milestones", subheading_style)) 
    milestones_text = f""" 
    <b>Month {implementation_delay_months}: Go-Live Milestone</b><br/> 
    • Solution deployed and operational<br/> 
    • Initial benefits begin to materialize<br/> 
    • User training completed<br/><br/> 
    <b>Month {implementation_delay_months + benefits_ramp_up_months}: Full Benefits Milestone</b><br/> 
    • 100% benefit realization achieved<br/> 
    • All processes optimized<br/> 
    • ROI tracking established<br/><br/> 
    <b>Month {evaluation_years * 12}: Final Review & Optimization</b><br/>
    • Comprehensive review of benefits realization<br/>
    • Identify areas for further optimization<br/>
    • Plan for future initiatives and expansion
    """
    story.append(Paragraph(milestones_text, styles['Normal']))
    story.append(Spacer(1, 0.3*inch))

    # Add a concluding statement
    story.append(Paragraph(
        "By implementing the proposed solution, " + organization_name + " can expect to achieve significant financial benefits and operational efficiencies, driving enhanced business value.",
        styles['Normal']
    ))

    # Build the PDF
    doc.build(story)
    buffer.seek(0)
    return buffer

# --- Main App Layout ---
st.title(f"Business Value Assessment for {solution_name} Implementation")
st.markdown("This tool helps estimate the financial impact of implementing the solution, providing a comprehensive business case with scenario analysis.")

st.header("Financial Impact Summary")

# --- Key Metrics Cards ---
col1, col2, col3 = st.columns(3)

with col1:
    expected_npv = scenario_results['Expected']['npv']
    st.metric(label=f"Expected Net Present Value (NPV) over {evaluation_years} years",
              value=f"{currency_symbol}{expected_npv:,.0f}")
    
with col2:
    expected_roi = scenario_results['Expected']['roi'] * 100
    st.metric(label=f"Expected Return on Investment (ROI) over {evaluation_years} years",
              value=f"{expected_roi:.1f}%")

with col3:
    expected_payback = scenario_results['Expected']['payback']
    expected_payback_months = scenario_results['Expected']['payback_months']
    st.metric(label="Expected Payback Period",
              value=f"{expected_payback_months}")

st.markdown("---")

# --- Value Reallocation & FTE Equivalency (Overall Project) ---
st.subheader("🚀 Value Reallocation & FTE Equivalency (Overall Project)")
st.write(f"**Cost Available for Higher Margin Projects (Annually):** {currency_symbol}{total_operational_savings_from_time_saved:,.0f}")
if effective_avg_fte_salary > 0:
    st.write(f"**Equivalent FTEs from Savings (Annually):** {equivalent_ftes_from_savings:,.1f} FTEs")
else:
    st.write("Average FTE salary not provided, unable to calculate equivalent FTEs.")

st.markdown("---")

# --- Scenario Analysis ---
st.header("Scenario Analysis")
st.info("Explore the potential financial outcomes under different assumptions.")

tabs = st.tabs(list(scenarios.keys()))

for i, (scenario_name, params) in enumerate(scenarios.items()):
    with tabs[i]:
        result = scenario_results[scenario_name]
        st.subheader(f"{params['icon']} {scenario_name} Scenario")
        st.markdown(f"*{params['description']}*")

        st.markdown("---")

        st.write(f"**Annual Benefits (Year {evaluation_years}):** {currency_symbol}{result['annual_benefits']:,.0f}")
        st.write(f"**Total Annual Platform Cost:** {currency_symbol}{platform_cost:,.0f}")
        st.write(f"**One-Time Services Cost:** {currency_symbol}{services_cost:,.0f}")

        st.markdown("---")

        st.write(f"**Net Present Value (NPV):** {currency_symbol}{result['npv']:,.0f}")
        st.write(f"**Return on Investment (ROI):** {result['roi']*100:.1f}%")
        st.write(f"**Payback Period (Years):** {result['payback']}")
        st.write(f"**Payback Period (Months):** {result['payback_months']}") # New monthly payback display

        # Display cash flows in a table
        st.markdown("#### Detailed Cash Flows")
        cash_flow_df = pd.DataFrame(result['cash_flows'])
        cash_flow_df['net_cash_flow_cumulative'] = cash_flow_df['net_cash_flow'].cumsum()

        # Format for display
        cash_flow_display_df = cash_flow_df.copy()
        for col in ['benefits', 'platform_cost', 'services_cost', 'net_cash_flow', 'net_cash_flow_cumulative']:
            cash_flow_display_df[col] = cash_flow_display_df[col].apply(lambda x: f"{currency_symbol}{x:,.0f}")
        
        cash_flow_display_df['realization_factor'] = cash_flow_display_df['realization_factor'].apply(lambda x: f"{x*100:.1f}%")

        st.dataframe(cash_flow_display_df[[
            'year', 'benefits', 'platform_cost', 'services_cost', 
            'net_cash_flow', 'net_cash_flow_cumulative', 'realization_factor'
        ]].rename(columns={
            'year': 'Year',
            'benefits': 'Benefits',
            'platform_cost': 'Platform Cost',
            'services_cost': 'Services Cost',
            'net_cash_flow': 'Net Cash Flow',
            'net_cash_flow_cumulative': 'Cumulative Net Cash Flow',
            'realization_factor': 'Benefit Realization Factor'
        }), hide_index=True)


st.markdown("---")

# --- Monthly Cumulative Cash Flow Chart (Expected Scenario - showing initial months) ---
st.subheader("Cumulative Net Cash Flow Over Time (Expected Scenario)")

def get_monthly_cumulative_cash_flow(annual_benefits, annual_platform_cost, one_time_services_cost, 
                                     implementation_delay_months, benefits_ramp_up_months, evaluation_years):
    total_months = evaluation_years * 12
    monthly_data = []
    cumulative_cash_flow = 0
    
    # Start with initial services cost as a negative cash flow at month 0
    monthly_data.append({'month': 0, 'net_cash_flow': -one_time_services_cost, 'cumulative_net_cash_flow': -one_time_services_cost})
    cumulative_cash_flow = -one_time_services_cost

    for month in range(1, total_months + 1):
        factor = calculate_benefit_realization_factor(month, implementation_delay_months, benefits_ramp_up_months)
        
        monthly_benefit = (annual_benefits / 12) * factor
        monthly_platform_cost = annual_platform_cost / 12
        
        monthly_net_cash_flow = monthly_benefit - monthly_platform_cost
        
        cumulative_cash_flow += monthly_net_cash_flow
        
        monthly_data.append({
            'month': month,
            'net_cash_flow': monthly_net_cash_flow,
            'cumulative_net_cash_flow': cumulative_cash_flow
        })
    return pd.DataFrame(monthly_data)

expected_monthly_cf_df = get_monthly_cumulative_cash_flow(
    total_annual_benefits * scenarios['Expected']['benefits_multiplier'],
    platform_cost,
    services_cost,
    implementation_delay_months,
    benefits_ramp_up_months,
    evaluation_years
)

fig_monthly_cf = px.line(expected_monthly_cf_df, x='month', y='cumulative_net_cash_flow',
                 labels={'cumulative_net_cash_flow': f'Cumulative Net Cash Flow ({currency_symbol})', 'month': 'Month'},
                 title='Cumulative Net Cash Flow (Expected Scenario - Monthly View)')
fig_monthly_cf.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Payback Point", 
                  annotation_position="bottom right")
fig_monthly_cf.update_layout(hovermode="x unified")
st.plotly_chart(fig_monthly_cf, use_container_width=True)


st.markdown("---")

# Implementation Timeline Chart
st.subheader("Implementation Timeline & Benefit Realization")
timeline_fig = create_implementation_timeline_chart(
    implementation_delay_months, 
    benefits_ramp_up_months, 
    evaluation_years, 
    currency_symbol, 
    total_annual_benefits
)
st.plotly_chart(timeline_fig, use_container_width=True)


st.markdown("---")

# --- Show BVA Calculations Section ---
st.header("Show BVA Calculations")
with st.expander("Click to view detailed calculations"):
    st.markdown("### Cost per Alert/Incident")
    st.write(f"**Cost per Alert:** {currency_symbol}{cost_per_alert:,.2f}")
    st.write(f"**Total Annual Alert Handling Cost:** {currency_symbol}{total_alert_handling_cost:,.0f}")
    st.write(f"**FTE Time % on Alerts:** {alert_fte_percentage*100:.1f}%")
    st.write(f"**Cost per Incident:** {currency_symbol}{cost_per_incident:,.2f}")
    st.write(f"**Total Annual Incident Handling Cost:** {currency_symbol}{total_incident_handling_cost:,.0f}")
    st.write(f"**Incident Triage Time Savings (Annual):** {currency_symbol}{incident_triage_savings:,.0f}")
    st.write(f"**Major Incident MTTR Savings (Annual):** {currency_symbol}{major_incident_savings:,.0f}")
    st.write(f"---")
    st.write(f"**Total Operational Savings from Alert/Incident Management (Annual):** {currency_symbol}{total_operational_savings_from_time_saved:,.0f}")
    st.write(f"**Total Additional Benefits (Tool Consolidaton, FTE Avoidance, etc.) (Annual):** {currency_symbol}{tool_savings + people_cost_per_year + fte_avoidance + sla_penalty_avoidance + revenue_growth + capex_savings + opex_savings:,.0f}")
    st.write(f"**TOTAL ANNUAL BASELINE BENEFITS:** {currency_symbol}{total_annual_benefits:,.0f}")
    st.write(f"---")
    st.write(f"**Effective Average FTE Salary:** {currency_symbol}{effective_avg_fte_salary:,.0f}")
    st.write(f"**Equivalent FTEs from Operational Savings:** {equivalent_ftes_from_savings:,.1f} FTEs")

st.markdown("---")

# --- BVA Story by Stakeholder Section ---
st.header("Stakeholder Value Propositions")
st.info("Tailored value messages for key stakeholders, highlighting the benefits most relevant to their roles.")

stakeholder_tabs = st.tabs(["CIO", "CTO", "CFO", "Operations Manager", "Service Desk Manager"])

with stakeholder_tabs[0]: # CIO
    st.subheader("For the CIO (Chief Information Officer)")
    st.markdown(f"""
    **Strategic Alignment & Digital Transformation:**
    Implementing {solution_name} is a strategic move towards a more proactive and agile IT environment. 
    By automating repetitive tasks and providing unified visibility, we can free up IT resources to focus on innovation and digital transformation initiatives that directly impact business growth. 
    The projected **{currency_symbol}{scenario_results['Expected']['npv']:,.0f} NPV** and **{scenario_results['Expected']['roi']*100:.1f}% ROI** over {evaluation_years} years demonstrate a strong financial case for this investment.
    
    Even under the conservative scenario, the solution still delivers a positive **{scenario_results['Conservative']['roi']*100:.1f}% ROI** with a payback period of **{scenario_results['Conservative']['payback_months']}**, affirming its robust value.

    **Key Benefits for the CIO:**
    * **Enhanced Service Delivery:** Proactive identification and resolution of issues lead to higher application availability and improved customer satisfaction.
    * **Operational Excellence:** Standardizes and automates IT operations, reducing manual effort and human error across the organization.
    * **Resource Optimization:** Reallocates **{equivalent_ftes_from_savings:,.1f} FTEs equivalent in savings** from reactive tasks to strategic projects, optimizing IT spending.
    * **Improved Decision Making:** Provides comprehensive insights into IT performance, enabling data-driven strategic planning.
    """)

with stakeholder_tabs[1]: # CTO
    st.subheader("For the CTO (Chief Technology Officer)")
    st.markdown(f"""
    **Technology Modernization & Resiliency:**
    {solution_name} directly addresses the complexities of our hybrid IT landscape, improving overall system resiliency and performance. 
    Its advanced AI/ML capabilities will enable us to move from reactive troubleshooting to predictive problem resolution, ensuring our technology stack supports business demands effectively.
    
    Even with conservative assumptions, the technology proves its worth, offering a **{scenario_results['Conservative']['roi']*100:.1f}% ROI** and reaching payback in **{scenario_results['Conservative']['payback_months']}**.

    **Key Benefits for the CTO:**
    * **Reduced MTTR:** A **{mttr_improvement_pct:.0f}% reduction in MTTR for major incidents** translates to significant cost savings of **{currency_symbol}{major_incident_savings:,.0f} annually** and minimized business disruption.
    * **Proactive Problem Solving:** AI-driven insights help identify root causes faster and even predict potential issues before they impact services.
    * **Scalability & Efficiency:** Automates routine operational tasks, allowing technical teams to scale operations without proportional headcount increases.
    * **Unified Observability:** Provides a single pane of glass for all infrastructure and application performance, breaking down data silos.
    """)

with stakeholder_tabs[2]: # CFO
    st.subheader("For the CFO (Chief Financial Officer)")
    st.markdown(f"""
    **Strong Financial Returns & Cost Optimization:**
    This investment in {solution_name} is projected to deliver substantial financial returns, with an **Expected Net Present Value of {currency_symbol}{scenario_results['Expected']['npv']:,.0f}** and an **ROI of {scenario_results['Expected']['roi']*100:.1f}%** over {evaluation_years} years. 
    The rapid payback period of **{scenario_results['Expected']['payback_months']}** ensures a quick return on our investment.
    
    Critically, even in the most conservative scenario, the solution demonstrates a positive **{scenario_results['Conservative']['roi']*100:.1f}% ROI** and achieves payback within **{scenario_results['Conservative']['payback_months']}**, confirming its financial viability under various conditions.

    **Key Benefits for the CFO:**
    * **Significant Cost Savings:** Achieves **{currency_symbol}{total_operational_savings_from_time_saved:,.0f} in annual operational savings** from reduced alert/incident volumes and improved efficiency.
    * **Predictable Budgeting:** Streamlined operations lead to more predictable and manageable IT operational expenditures.
    """)

with stakeholder_tabs[3]: # Operations Manager
    st.subheader("For the Operations Manager")
    st.markdown(f"""
    **Streamlined Operations & Reduced Toil:**
    {solution_name} will significantly enhance our operational efficiency by reducing noise and automating routine tasks. 
    This means fewer false alarms, faster triage, and more time for your teams to focus on impactful work rather than constant firefighting.
    
    **Key Benefits for the Operations Manager:**
    * **Alert & Incident Reduction:** Expect a **{alert_reduction_pct:.0f}% reduction in alerts** and **{incident_reduction_pct:.0f}% reduction in incidents**, leading to less operational burden.
    * **Faster Triage & Resolution:** Improve average alert triage time by **{alert_triage_time_saved_pct:.0f}%** and incident triage by **{incident_triage_time_savings_pct:.0f}%**, saving significant time and effort.
    * **Automated Workflows:** Automate repetitive responses to common issues, improving consistency and speed.
    * **Improved Team Morale:** Reduce alert fatigue and empower your team with better tools and a clearer focus.
    """)

with stakeholder_tabs[4]: # Service Desk Manager
    st.subheader("For the Service Desk Manager")
    st.markdown(f"""
    **Enhanced Service Quality & Customer Satisfaction:**
    {solution_name} will empower your service desk with more accurate and actionable information, enabling faster resolution of user-reported issues and even preventing issues before users notice them.
    
    **Key Benefits for the Service Desk Manager:**
    * **Reduced Ticket Volume:** Fewer incidents mean fewer tickets, easing the burden on the service desk team.
    * **Improved First-Call Resolution:** Better diagnostics and automated runbooks provide service desk agents with the information needed to resolve issues quickly.
    * **Proactive Issue Resolution:** By integrating with IT operations, many issues can be resolved before they escalate to user-impacting problems.
    * **Clearer Communication:** Provides real-time status and impact assessments, improving communication with end-users during outages.
    """)


st.markdown("---")

# --- Executive Report Generation ---
st.header("Generate Executive Report")

if REPORT_DEPENDENCIES_AVAILABLE:
    st.write("Generate a professional PDF executive summary of this Business Value Assessment.")
    
    org_name_for_report = st.text_input("Your Organization Name (for report)", value="My Company", key="org_name_report")

    if st.button("Generate PDF Report"):
        with st.spinner("Generating PDF report..."):
            summary_data_for_report = create_executive_summary_data(scenario_results, currency_symbol)
            pdf_buffer = generate_executive_report_pdf(summary_data_for_report, scenario_results, solution_name, org_name_for_report)
            if pdf_buffer:
                st.download_button(
                    label="Download PDF Report",
                    data=pdf_buffer,
                    file_name=f"{org_name_for_report}_{solution_name}_BVA_Report.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("Failed to generate PDF report. Please check if reportlab dependencies are installed correctly.")
else:
    st.warning("To generate PDF reports, please install `reportlab` and `matplotlib` (`pip install reportlab matplotlib`).")

