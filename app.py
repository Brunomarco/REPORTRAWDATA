import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Configure Streamlit page
st.set_page_config(
page_title="LFS Amsterdam - TMS Performance Dashboard",
page_icon="📊",
layout="wide",
initial_sidebar_state="expanded"
)

# Custom CSS - minimal styling
st.markdown("""
<style>
.main-header {
font-size: 2.5rem;
font-weight: bold;
color: #1f77b4;
text-align: center;
margin-bottom: 2rem;
}
.section-header {
font-size: 1.8rem;
font-weight: bold;
color: #2c3e50;
margin: 2rem 0 1.5rem 0;
padding: 0.8rem 0;
border-bottom: 2px solid #3498db;
}
.insight-box {
background: #f0f8ff;
padding: 1.5rem;
border-radius: 8px;
margin: 1.5rem 0;
border-left: 4px solid #3498db;
}
.report-section {
margin: 2rem 0;
padding: 1.5rem;
background: #fafafa;
border-radius: 8px;
}
.chart-title {
font-size: 1.2rem;
font-weight: bold;
color: #2c3e50;
margin-bottom: 0.5rem;
text-align: center;
}
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<h1 class="main-header">LFS Amsterdam TMS Performance Dashboard</h1>', unsafe_allow_html=True)

# Sidebar
st.sidebar.title("📊 Dashboard Controls")
st.sidebar.markdown("---")

uploaded_file = st.sidebar.file_uploader(
"Upload TMS Excel File",
type=['xlsx', 'xls'],
help="Upload your 'report raw data.xls' file"
)

# Define service types and countries correctly
SERVICE_TYPES = ['CTX', 'CX', 'EF', 'EGD', 'FF', 'RGD', 'ROU', 'SF']
COUNTRIES = ['AT', 'AU', 'BE', 'DE', 'DK', 'ES', 'FR', 'GB', 'IT', 'N1', 'NL', 'NZ', 'SE', 'US']

# Complete QC Name mapping
QC_CATEGORIES = {
'MNX-Incorrect QDT': 'System Error',
'Customer-Changed delivery parameters': 'Customer Related',
'Consignee-Driver waiting at delivery': 'Delivery Issue',
'Customer-Requested delay': 'Customer Related',
'Customer-Shipment not ready': 'Customer Related',
'Del Agt-Late del': 'Delivery Issue',
'Consignee-Changed delivery parameters': 'Delivery Issue'
}

def safe_date_conversion(date_series):
 """Safely convert Excel dates"""
 try:
  if date_series.dtype in ['int64', 'float64']:
   return pd.to_datetime(date_series, origin='1899-12-30', unit='D', errors='coerce')
  else:
   return pd.to_datetime(date_series, errors='coerce')
 except:
  return date_series

@st.cache_data
def load_tms_data(uploaded_file):
 """Load and process TMS Excel file"""
 if uploaded_file is not None:
  try:
   excel_sheets = pd.read_excel(uploaded_file, sheet_name=None)
   data = {}
   
   # 1. Raw Data
   if "AMS RAW DATA" in excel_sheets:
    data['raw_data'] = excel_sheets["AMS RAW DATA"].copy()
   
   # 2. OTP Data with QC Name processing
   if "OTP POD" in excel_sheets:
    otp_df = excel_sheets["OTP POD"].copy()
    # Get first 6 columns to include QC Name
    if len(otp_df.columns) >= 6:
     otp_df = otp_df.iloc[:, :6]
     otp_df.columns = ['TMS_Order', 'QDT', 'POD_DateTime', 'Time_Diff', 'Status', 'QC_Name']
    else:
     # Handle case with fewer columns
     cols = ['TMS_Order', 'QDT', 'POD_DateTime', 'Time_Diff', 'Status'][:len(otp_df.columns)]
     otp_df.columns = cols
    otp_df = otp_df.dropna(subset=['TMS_Order'])
    data['otp'] = otp_df
   
   # 3. Volume Data - Process dynamically from Excel
   if "Volume per SVC" in excel_sheets:
    volume_df = excel_sheets["Volume per SVC"].copy()
    
    # Try to read the actual matrix from Excel
    if not volume_df.empty:
     # Assuming first column is countries and rest are services
     volume_df = volume_df.set_index(volume_df.columns[0])
     
     # Initialize dictionaries
     service_country_matrix = {}
     service_volumes = {}
     country_volumes = {}
     
     # Process the matrix
     for country in volume_df.index:
      if pd.notna(country) and country in COUNTRIES:
       service_country_matrix[country] = {}
       country_total = 0
       
       for service in volume_df.columns:
        if service in SERVICE_TYPES:
         value = volume_df.loc[country, service]
         if pd.notna(value) and value > 0:
          service_country_matrix[country][service] = int(value)
          country_total += int(value)
          
          # Update service totals
          if service not in service_volumes:
           service_volumes[service] = 0
          service_volumes[service] += int(value)
       
       if country_total > 0:
        country_volumes[country] = country_total
     
     # Calculate total volume
     total_vol = sum(service_volumes.values())
     
     data['service_volumes'] = service_volumes
     data['country_volumes'] = country_volumes
     data['service_country_matrix'] = service_country_matrix
     data['total_volume'] = total_vol
   
   # 4. Lane Usage - Process dynamically from Excel
   if "Lane usage " in excel_sheets:
    lane_df = excel_sheets["Lane usage "].copy()
    
    if not lane_df.empty:
     # Process lane matrix dynamically
     lane_df = lane_df.set_index(lane_df.columns[0])
     
     lanes_list = []
     origin_totals = {}
     dest_totals = {}
     
     for origin in lane_df.index:
      if pd.notna(origin):
       origin_total = 0
       for dest in lane_df.columns:
        if pd.notna(dest):
         value = lane_df.loc[origin, dest]
         if pd.notna(value) and value > 0:
          lanes_list.append({
           'Origin': str(origin),
           'Destination': str(dest),
           'Volume': int(value)
          })
          origin_total += int(value)
          
          # Update destination totals
          if dest not in dest_totals:
           dest_totals[dest] = 0
          dest_totals[dest] += int(value)
       
       if origin_total > 0:
        origin_totals[str(origin)] = origin_total
     
     data['lanes_list'] = lanes_list
     data['origin_totals'] = origin_totals
     data['dest_totals'] = dest_totals
    
    data['lanes'] = lane_df
   
   # 5. Cost Sales - Fixed to properly process financial data
   if "cost sales" in excel_sheets:
    cost_df = excel_sheets["cost sales"].copy()
    
    # Dynamically map columns based on content
    if not cost_df.empty:
     # Try to identify columns by their content patterns
     for i, col in enumerate(cost_df.columns):
      sample_values = cost_df[col].dropna().head(10)
      
      # Date column
      if any(isinstance(v, (pd.Timestamp, datetime)) for v in sample_values):
       cost_df.rename(columns={col: 'Order_Date'}, inplace=True)
      # Revenue column (look for positive monetary values)
      elif 'revenue' in str(col).lower() or 'sales' in str(col).lower():
       cost_df.rename(columns={col: 'Net_Revenue'}, inplace=True)
      # Cost columns
      elif 'pu' in str(col).lower() and 'cost' in str(col).lower():
       cost_df.rename(columns={col: 'PU_Cost'}, inplace=True)
      elif 'ship' in str(col).lower() and 'cost' in str(col).lower():
       cost_df.rename(columns={col: 'Ship_Cost'}, inplace=True)
      elif 'man' in str(col).lower() and 'cost' in str(col).lower():
       cost_df.rename(columns={col: 'Man_Cost'}, inplace=True)
      elif 'del' in str(col).lower() and 'cost' in str(col).lower():
       cost_df.rename(columns={col: 'Del_Cost'}, inplace=True)
      elif 'total' in str(col).lower() and 'cost' in str(col).lower():
       cost_df.rename(columns={col: 'Total_Cost'}, inplace=True)
      # Gross percent
      elif 'gross' in str(col).lower() and '%' in str(col).lower():
       cost_df.rename(columns={col: 'Gross_Percent'}, inplace=True)
      # Country column
      elif 'country' in str(col).lower() or (len(sample_values) > 0 and all(len(str(v)) == 2 for v in sample_values if pd.notna(v))):
       cost_df.rename(columns={col: 'PU_Country'}, inplace=True)
     
     # Convert date columns
     if 'Order_Date' in cost_df.columns:
      cost_df['Order_Date'] = safe_date_conversion(cost_df['Order_Date'])
     
     # Clean financial data - remove rows with missing financial values
     financial_cols = ['Net_Revenue', 'Total_Cost']
     existing_fin_cols = [col for col in financial_cols if col in cost_df.columns]
     
     if existing_fin_cols:
      cost_df = cost_df.dropna(subset=existing_fin_cols)
      # Only keep rows with actual financial activity
      if 'Net_Revenue' in cost_df.columns and 'Total_Cost' in cost_df.columns:
       cost_df = cost_df[(cost_df['Net_Revenue'] != 0) | (cost_df['Total_Cost'] != 0)]
     
     data['cost_sales'] = cost_df
   
   return data
   
  except Exception as e:
   st.error(f"Error processing Excel file: {str(e)}")
   return None
 return None

# Load data
tms_data = None
if uploaded_file is not None:
 tms_data = load_tms_data(uploaded_file)
 if tms_data:
  st.sidebar.success("✅ Data loaded successfully")
 else:
  st.sidebar.error("❌ Error loading data")
else:
 st.sidebar.info("📁 Upload Excel file to begin")

# Calculate global metrics for use across tabs
avg_otp = 0
total_orders = 0
total_revenue = 0
total_cost = 0
profit_margin = 0
total_services = 0

if tms_data is not None:
 # Calculate key metrics
 total_services = tms_data.get('total_volume', sum(tms_data.get('service_volumes', {}).values()))
 
 # OTP metrics
 if 'otp' in tms_data and not tms_data['otp'].empty:
  otp_df = tms_data['otp']
  if 'Status' in otp_df.columns:
   status_series = otp_df['Status'].dropna()
   total_orders = len(status_series)
   on_time_orders = len(status_series[status_series == 'ON TIME'])
   avg_otp = (on_time_orders / total_orders * 100) if total_orders > 0 else 0
 
 # Financial metrics - Fixed to only use rows with actual financial data
 if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
  cost_df = tms_data['cost_sales']
  if 'Net_Revenue' in cost_df.columns:
   total_revenue = cost_df['Net_Revenue'].sum()
  if 'Total_Cost' in cost_df.columns:
   total_cost = cost_df['Total_Cost'].sum()
  profit_margin = ((total_revenue - total_cost) / total_revenue * 100) if total_revenue > 0 else 0

# Create tabs for each sheet
if tms_data is not None:
 tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
  "📊 Overview", 
  "📦 Volume Analysis", 
  "⏱️ OTP Performance", 
  "💰 Financial Analysis", 
  "🛣️ Lane Network",
  "📄 Executive Report"
 ])
 
 # TAB 1: Overview
 with tab1:
  st.markdown('<h2 class="section-header">Executive Dashboard Overview</h2>', unsafe_allow_html=True)
  
  # KPI Dashboard
  col1, col2, col3, col4 = st.columns(4)
  
  with col1:
   st.metric("📦 Total Volume", f"{int(total_services):,}", "shipments")
  
  with col2:
   st.metric("⏱️ OTP Rate", f"{avg_otp:.1f}%", f"{avg_otp-95:.1f}% vs target")
  
  with col3:
   st.metric("💰 Revenue", f"€{total_revenue:,.0f}", "total")
  
  with col4:
   st.metric("📈 Margin", f"{profit_margin:.1f}%", f"{profit_margin-20:.1f}% vs target")
  
  # Performance Summary
  col1, col2 = st.columns(2)
  
  with col1:
   st.markdown('<div class="insight-box">', unsafe_allow_html=True)
   st.markdown("### 📊 What These Numbers Mean")
   st.markdown(f"""
   **Volume Analysis:**
   - The **{total_services} shipments** represent all packages handled by LFS Amsterdam
   - With **{len(tms_data.get('country_volumes', {}))} active countries**, we process diverse international operations
   - Netherlands handles the largest volume, confirming Amsterdam as the main hub
   
   **Service Distribution:**
   - **{len([s for s in tms_data.get('service_volumes', {}).values() if s > 0])} active service types** provide flexibility for different customer needs
   - Service mix shows balanced operations between speed and cost-efficiency
   """)
   st.markdown('</div>', unsafe_allow_html=True)
  
  with col2:
   st.markdown('<div class="insight-box">', unsafe_allow_html=True)
   st.markdown("### 🎯 Performance Interpretation")
   
   if avg_otp >= 95:
    st.markdown(f"""
    ✅ **OTP at {avg_otp:.1f}%** means we deliver on-time {int(avg_otp/100 * total_orders)} out of {total_orders} orders
    - This exceeds industry standard (95%), showing reliable service
    - Customers can trust our delivery promises
    """)
   else:
    st.markdown(f"""
    ⚠️ **OTP at {avg_otp:.1f}%** means we're late on {total_orders - int(avg_otp/100 * total_orders)} out of {total_orders} orders
    - We need {int((95-avg_otp)/100 * total_orders)} more on-time deliveries to hit target
    - Each 1% improvement = {total_orders/100:.0f} more satisfied customers
    """)
   
   if profit_margin >= 20:
    st.markdown(f"""
    ✅ **{profit_margin:.1f}% margin** means €{profit_margin:.0f} profit per €100 revenue
    - Healthy profitability above 20% target
    - Strong financial position for growth investments
    """)
   else:
    st.markdown(f"""
    ⚠️ **{profit_margin:.1f}% margin** needs improvement
    - Currently €{profit_margin:.0f} profit per €100 revenue
    - Need to increase by €{20-profit_margin:.0f} per €100 to hit target
    """)
   st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 2: Volume Analysis
 with tab2:
  st.markdown('<h2 class="section-header">Volume Analysis by Service & Country</h2>', unsafe_allow_html=True)
  
  if 'service_volumes' in tms_data and tms_data['service_volumes']:
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown('<p class="chart-title">Service Type Distribution - What We Ship</p>', unsafe_allow_html=True)
    
    service_data = pd.DataFrame(list(tms_data['service_volumes'].items()), 
                              columns=['Service', 'Volume'])
    service_data = service_data[service_data['Volume'] > 0]
    
    # Use darker colors
    fig = px.bar(service_data, x='Service', y='Volume', 
                color='Volume', 
                color_continuous_scale=[[0, '#08519c'], [0.5, '#3182bd'], [1, '#6baed6']],
                title='')
    fig.update_layout(showlegend=False, height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Service breakdown with interpretation
    service_table = service_data.copy()
    service_table['Share %'] = (service_table['Volume'] / service_table['Volume'].sum() * 100).round(1)
    service_table['Interpretation'] = service_table.apply(
     lambda x: f"{'Leading' if x['Share %'] > 20 else 'Secondary' if x['Share %'] > 10 else 'Niche'} service",
     axis=1
    )
    service_table = service_table.sort_values('Volume', ascending=False)
    st.dataframe(service_table, hide_index=True, use_container_width=True)
   
   with col2:
    st.markdown('<p class="chart-title">Country Distribution - Where We Operate</p>', unsafe_allow_html=True)
    
    if 'country_volumes' in tms_data and tms_data['country_volumes']:
     country_data = pd.DataFrame(list(tms_data['country_volumes'].items()), 
                               columns=['Country', 'Volume'])
     
     # Use darker green colors
     fig = px.bar(country_data, x='Country', y='Volume',
                 color='Volume', 
                 color_continuous_scale=[[0, '#006d2c'], [0.5, '#31a354'], [1, '#74c476']],
                 title='')
     fig.update_layout(showlegend=False, height=400)
     st.plotly_chart(fig, use_container_width=True)
     
     # Country breakdown with regions
     country_table = country_data.copy()
     country_table['Share %'] = (country_table['Volume'] / country_table['Volume'].sum() * 100).round(1)
     country_table['Region'] = country_table['Country'].apply(
      lambda x: 'Europe' if x in ['AT', 'BE', 'DE', 'DK', 'ES', 'FR', 'GB', 'IT', 'NL', 'SE'] 
      else 'Americas' if x in ['US'] 
      else 'Asia-Pacific' if x in ['AU', 'NZ'] 
      else 'Other'
     )
     country_table = country_table.sort_values('Volume', ascending=False)
     st.dataframe(country_table, hide_index=True, use_container_width=True)
  
  # Service-Country Matrix Heatmap
  if 'service_country_matrix' in tms_data:
   st.markdown('<p class="chart-title">Service-Country Matrix - What Services Go Where</p>', unsafe_allow_html=True)
   
   # Create matrix dataframe
   matrix_data = []
   all_countries = sorted(tms_data['service_country_matrix'].keys())
   all_services = sorted(set(service for country_data in tms_data['service_country_matrix'].values() for service in country_data.keys()))
   
   for country in all_countries:
    row = {'Country': country}
    for service in all_services:
     if country in tms_data['service_country_matrix'] and service in tms_data['service_country_matrix'][country]:
      row[service] = tms_data['service_country_matrix'][country][service]
     else:
      row[service] = 0
    matrix_data.append(row)
   
   matrix_df = pd.DataFrame(matrix_data)
   matrix_df = matrix_df.set_index('Country')
   
   # Create heatmap
   fig = px.imshow(matrix_df.T, 
                  labels=dict(x="Country", y="Service Type", color="Volume"),
                  title="",
                  color_continuous_scale='YlOrRd',
                  aspect='auto')
   fig.update_layout(height=500)
   st.plotly_chart(fig, use_container_width=True)
  
  # Detailed Analysis with meaning
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### 📦 Understanding the Volume Patterns")
  
  # Display top services dynamically
  if 'service_volumes' in tms_data:
   top_services = sorted([(k, v) for k, v in tms_data['service_volumes'].items() if v > 0], 
                        key=lambda x: x[1], reverse=True)[:3]
   
   st.markdown(f"""
   **What the Service Distribution Tells Us:**
   - **{top_services[0][0] if top_services else 'Top'} Service ({top_services[0][1] if top_services else 0} shipments)**: Leading service type
   - **{top_services[1][0] if len(top_services) > 1 else 'Second'} Service ({top_services[1][1] if len(top_services) > 1 else 0} shipments)**: Secondary volume driver
   - Service mix shows operational focus areas
   
   **Geographic Insights:**
   - Netherlands processes the highest volume as the central hub
   - European markets dominate the operation
   - Clear opportunities in underserved markets
   
   **Business Implications:**
   - Hub-and-spoke model is working with Amsterdam central
   - Service concentration suggests operational efficiency focus
   - Geographic spread provides risk diversification
   """)
  st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 3: OTP Performance
 with tab3:
  st.markdown('<h2 class="section-header">On-Time Performance Analysis</h2>', unsafe_allow_html=True)
  
  if 'otp' in tms_data and not tms_data['otp'].empty:
   otp_df = tms_data['otp']
   
   # OTP Status Analysis
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown('<p class="chart-title">Delivery Performance Breakdown</p>', unsafe_allow_html=True)
    
    if 'Status' in otp_df.columns:
     status_counts = otp_df['Status'].value_counts()
     
     fig = px.pie(values=status_counts.values, names=status_counts.index,
                 title='',
                 color_discrete_map={'ON TIME': '#2ca02c', 'LATE': '#d62728'})
     fig.update_traces(textposition='inside', textinfo='percent+label')
     st.plotly_chart(fig, use_container_width=True)
    
    # Performance Metrics with explanations
    on_time_count = int(avg_otp/100 * total_orders)
    late_count = total_orders - on_time_count
    
    metrics_data = pd.DataFrame({
     'Metric': ['Total Orders', 'On-Time', 'Late', 'OTP Rate'],
     'Value': [
      f"{total_orders:,}",
      f"{on_time_count:,}",
      f"{late_count:,}",
      f"{avg_otp:.1f}%"
     ],
     'Target': [
      '-',
      '-',
      '-',
      '95%'
     ]
    })
    st.dataframe(metrics_data, hide_index=True, use_container_width=True)
   
   with col2:
    st.markdown('<p class="chart-title">Root Causes of Delays</p>', unsafe_allow_html=True)
    
    if 'QC_Name' in otp_df.columns:
     # Process all QC reasons
     qc_data = []
     for idx, value in otp_df['QC_Name'].dropna().items():
      reasons = str(value).strip()
      if reasons and reasons != 'nan':
       qc_data.append(reasons)
     
     # Count occurrences
     qc_counts = {}
     for reasons in qc_data:
      # Common delay reasons from the data
      delay_reasons = [
       'MNX-Incorrect QDT',
       'Customer-Changed delivery parameters',
       'Consignee-Driver waiting at delivery',
       'Customer-Requested delay',
       'Customer-Shipment not ready',
       'Del Agt-Late del',
       'Consignee-Changed delivery parameters'
      ]
      
      for reason in delay_reasons:
       if reason in reasons:
        if reason not in qc_counts:
         qc_counts[reason] = 0
        qc_counts[reason] += 1
     
     if qc_counts:
      # Categorize for visualization
      category_summary = {
       'Customer Issues': 0,
       'System Errors': 0,
       'Delivery Problems': 0
      }
      
      for reason, count in qc_counts.items():
       if 'Customer' in reason:
        category_summary['Customer Issues'] += count
       elif 'MNX' in reason:
        category_summary['System Errors'] += count
       else:
        category_summary['Delivery Problems'] += count
      
      fig = px.bar(x=list(category_summary.keys()), y=list(category_summary.values()),
                  title='',
                  color=list(category_summary.values()),
                  color_continuous_scale='Reds')
      fig.update_layout(showlegend=False, xaxis_title='Category', yaxis_title='Count')
      st.plotly_chart(fig, use_container_width=True)
      
      # Show detailed reasons
      st.markdown("**Detailed Delay Reasons:**")
      qc_detail_df = pd.DataFrame(list(qc_counts.items()), columns=['Reason', 'Count'])
      qc_detail_df = qc_detail_df.sort_values('Count', ascending=False)
      st.dataframe(qc_detail_df, hide_index=True, use_container_width=True)
  
  # OTP Summary Insights
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### ⏱️ OTP Performance Summary")
  st.markdown(f"""
  **Key Findings:**
  - Current OTP: {avg_otp:.1f}% {'(Above target)' if avg_otp >= 95 else '(Below target)'}
  - Total orders tracked: {total_orders:,}
  - On-time deliveries: {int(avg_otp/100 * total_orders):,}
  - Late deliveries: {total_orders - int(avg_otp/100 * total_orders):,}
  
  **Main Delay Causes:**
  - Customer-related issues (parameter changes, not ready)
  - System errors (incorrect delivery time calculations)
  - Delivery execution challenges
  
  **Improvement Opportunities:**
  - Better customer communication to reduce last-minute changes
  - System fixes for accurate delivery time calculations
  - Operational improvements in last-mile delivery
  """)
  st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 4: Financial Analysis
 with tab4:
  st.markdown('<h2 class="section-header">Financial Performance & Profitability</h2>', unsafe_allow_html=True)
  
  if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
   cost_df = tms_data['cost_sales']
   
   # Financial Overview
   st.markdown('<p class="chart-title">Overall Financial Health</p>', unsafe_allow_html=True)
   
   # First row - Revenue/Cost/Profit and Cost Breakdown
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown("**Revenue vs Cost Analysis**")
    st.markdown("<small>Shows total income, expenses, and resulting profit</small>", unsafe_allow_html=True)
    
    profit = total_revenue - total_cost
    financial_data = pd.DataFrame({
     'Category': ['Revenue', 'Cost', 'Profit'],
     'Amount': [total_revenue, total_cost, profit]
    })
    
    fig = px.bar(financial_data, x='Category', y='Amount',
                color='Category',
                color_discrete_map={'Revenue': '#2ca02c', 
                                  'Cost': '#ff7f0e',
                                  'Profit': '#2ca02c' if profit >= 0 else '#d62728'},
                title='')
    fig.update_layout(showlegend=False, height=350)
    st.plotly_chart(fig, use_container_width=True)
    
    # Financial summary
    st.write(f"**Total Revenue**: €{total_revenue:,.0f}")
    st.write(f"**Total Cost**: €{total_cost:,.0f}")
    st.write(f"**Net Profit**: €{profit:,.0f}")
    st.write(f"**Overall Margin**: {profit_margin:.1f}%")
   
   with col2:
    st.markdown("**Cost Structure Breakdown**")
    st.markdown("<small>Understanding where money is spent</small>", unsafe_allow_html=True)
    
    cost_components = {}
    cost_cols = ['PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost']
    for col in cost_cols:
     if col in cost_df.columns:
      cost_sum = cost_df[col].sum()
      if cost_sum > 0:
       cost_components[col.replace('_Cost', '')] = cost_sum
    
    if cost_components:
     # Add percentages to labels
     total_costs = sum(cost_components.values())
     labels = [f"{k}<br>{v/total_costs*100:.1f}%" for k, v in cost_components.items()]
     
     fig = px.pie(values=list(cost_components.values()), 
                names=labels,
                title='')
     fig.update_traces(textposition='inside', textinfo='value+label')
     fig.update_layout(height=350, showlegend=False)
     st.plotly_chart(fig, use_container_width=True)
    
    # Cost insights
    if cost_components:
     largest_cost = max(cost_components, key=cost_components.get)
     st.write(f"**Biggest expense**: {largest_cost} ({cost_components[largest_cost]/total_costs*100:.1f}%)")
     st.write(f"**Total costs**: €{total_costs:,.0f}")
   
   # Second row - Profit Margin Distribution (separate plot as requested)
   st.markdown('<p class="chart-title">Profit Margin Analysis</p>', unsafe_allow_html=True)
   
   if 'Gross_Percent' in cost_df.columns:
    # Create a dedicated section for margin distribution
    col1, col2 = st.columns(2)
    
    with col1:
     st.markdown("**Margin Distribution by Order**")
     margin_data = cost_df['Gross_Percent'].dropna()
     if not margin_data.empty:
      # Convert to percentage if needed
      if margin_data.max() <= 1:
       margin_data = margin_data * 100
      
      fig = px.histogram(margin_data, nbins=30,
                       title='',
                       labels={'value': 'Margin %', 'count': 'Number of Orders'})
      fig.add_vline(x=20, line_dash="dash", line_color="green", 
                  annotation_text="Target 20%")
      fig.add_vline(x=0, line_dash="dash", line_color="red", 
                  annotation_text="Break-even")
      fig.update_traces(marker_color='lightcoral')
      fig.update_layout(height=400)
      st.plotly_chart(fig, use_container_width=True)
    
    with col2:
     st.markdown("**Margin Statistics**")
     
     # Calculate detailed margin statistics
     profitable_orders = len(margin_data[margin_data > 0])
     loss_making_orders = len(margin_data[margin_data < 0])
     high_margin_orders = len(margin_data[margin_data >= 20])
     
     margin_stats = pd.DataFrame({
      'Category': [
       'Total Orders Analyzed',
       'Profitable Orders',
       'Loss-making Orders',
       'High Margin (≥20%)',
       'Average Margin',
       'Median Margin'
      ],
      'Value': [
       f"{len(margin_data):,}",
       f"{profitable_orders:,} ({profitable_orders/len(margin_data)*100:.1f}%)",
       f"{loss_making_orders:,} ({loss_making_orders/len(margin_data)*100:.1f}%)",
       f"{high_margin_orders:,} ({high_margin_orders/len(margin_data)*100:.1f}%)",
       f"{margin_data.mean():.1f}%",
       f"{margin_data.median():.1f}%"
      ]
     })
     st.dataframe(margin_stats, hide_index=True, use_container_width=True)
   
   # Country Financial Performance - FIXED to only show countries with financial data
   if 'PU_Country' in cost_df.columns:
    st.markdown('<p class="chart-title">Country-by-Country Financial Performance</p>', unsafe_allow_html=True)
    
    # Only aggregate countries that have financial data
    country_financials = cost_df.groupby('PU_Country').agg({
     'Net_Revenue': 'sum',
     'Total_Cost': 'sum',
     'Gross_Percent': 'mean'
    }).round(2)
    
    country_financials['Profit'] = country_financials['Net_Revenue'] - country_financials['Total_Cost']
    country_financials['Margin_Percent'] = (country_financials['Gross_Percent'] * 100).round(1)
    
    # Sort by revenue
    country_financials = country_financials.sort_values('Net_Revenue', ascending=False)
    
    # Create subplots with better spacing
    col1, col2 = st.columns([1, 1])
    
    with col1:
     st.markdown("**Revenue by Country**")
     st.markdown("<small>Which markets generate most income?</small>", unsafe_allow_html=True)
     
     revenue_data = country_financials.reset_index()
     revenue_data = revenue_data[revenue_data['Net_Revenue'] > 0]
     
     fig = px.bar(revenue_data, x='PU_Country', y='Net_Revenue',
                title='',
                color='Net_Revenue',
                color_continuous_scale=[[0, '#006d2c'], [0.5, '#31a354'], [1, '#74c476']])
     fig.update_layout(showlegend=False, height=400)
     st.plotly_chart(fig, use_container_width=True)
    
    with col2:
     st.markdown("**Profit/Loss by Country**")
     st.markdown("<small>Which routes are actually profitable?</small>", unsafe_allow_html=True)
     
     profit_data = country_financials[['Profit']].reset_index()
     profit_data['Color'] = profit_data['Profit'].apply(lambda x: 'Profit' if x >= 0 else 'Loss')
     
     fig = px.bar(profit_data, x='PU_Country', y='Profit',
                title='',
                color='Color',
                color_discrete_map={'Profit': '#2ca02c', 'Loss': '#d62728'})
     fig.update_layout(showlegend=False, height=400)
     st.plotly_chart(fig, use_container_width=True)
    
    # Detailed financial table with insights - only show countries with data
    st.markdown("**Detailed Country Financial Performance**")
    
    display_financials = country_financials.copy()
    display_financials['Revenue'] = display_financials['Net_Revenue'].round(0).astype(int)
    display_financials['Cost'] = display_financials['Total_Cost'].round(0).astype(int)
    display_financials['Profit'] = display_financials['Profit'].round(0).astype(int)
    display_financials['Status'] = display_financials['Profit'].apply(
     lambda x: '🟢 Profitable' if x > 0 else '🔴 Loss-making'
    )
    display_financials = display_financials[['Revenue', 'Cost', 'Profit', 'Margin_Percent', 'Status']]
    display_financials.columns = ['Revenue (€)', 'Cost (€)', 'Profit (€)', 'Margin (%)', 'Status']
    
    st.dataframe(display_financials, use_container_width=True)
    
    # Add summary statistics
    st.markdown("**Financial Summary by Country:**")
    profitable_countries = len(display_financials[display_financials['Profit (€)'] > 0])
    total_countries = len(display_financials)
    st.write(f"- **Profitable countries**: {profitable_countries} out of {total_countries}")
    st.write(f"- **Average country margin**: {display_financials['Margin (%)'].mean():.1f}%")
    st.write(f"- **Best performing**: {display_financials.index[0]} with {display_financials.iloc[0]['Margin (%)']}% margin")
  
  # Financial Insights with business meaning
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### 💰 Financial Performance Summary")
  st.markdown(f"""
  **Overall Financial Health:**
  - **Total Revenue**: €{total_revenue:,.0f}
  - **Total Costs**: €{total_cost:,.0f}
  - **Net Profit**: €{total_revenue - total_cost:,.0f}
  - **Overall Margin**: {profit_margin:.1f}% {'(Above target)' if profit_margin >= 20 else '(Below 20% target)'}
  - **Per shipment**: €{total_revenue/total_services:.2f} revenue, €{total_cost/total_services:.2f} cost
  
  **Cost Structure Insights:**
  - Identify largest cost drivers for optimization
  - Balance between pickup, shipping, handling, and delivery costs
  - Focus efficiency improvements on highest cost areas
  
  **Margin Distribution Analysis:**
  - Shows profitability variation across orders
  - Identifies pricing opportunities and problem areas
  - Target: Increase orders above 20% margin threshold
  
  **Country Performance:**
  - Not all high-revenue countries are profitable
  - Some smaller markets show better margins
  - Consider volume vs. profitability trade-offs
  """)
  st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 5: Lane Network
 with tab5:
  st.markdown('<h2 class="section-header">Lane Network & Route Analysis</h2>', unsafe_allow_html=True)
  
  if 'lanes_list' in tms_data and tms_data['lanes_list']:
   st.markdown('<p class="chart-title">Trade Lane Network Visualization</p>', unsafe_allow_html=True)
   
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown("**Top Origin Countries**")
    st.markdown("<small>Countries sending most shipments</small>", unsafe_allow_html=True)
    
    if 'origin_totals' in tms_data:
     origin_data = pd.DataFrame(list(tms_data['origin_totals'].items()), 
                              columns=['Origin', 'Volume'])
     origin_data = origin_data.sort_values('Volume', ascending=False).head(10)
     
     fig = px.bar(origin_data, x='Origin', y='Volume',
                title='',
                color='Volume',
                color_continuous_scale='Blues')
     fig.update_layout(showlegend=False, height=350)
     st.plotly_chart(fig, use_container_width=True)
   
   with col2:
    st.markdown("**Top Destination Countries**")
    st.markdown("<small>Countries receiving most shipments</small>", unsafe_allow_html=True)
    
    if 'dest_totals' in tms_data:
     dest_data = pd.DataFrame(list(tms_data['dest_totals'].items()), 
                            columns=['Destination', 'Volume'])
     dest_data = dest_data.sort_values('Volume', ascending=False).head(10)
     
     fig = px.bar(dest_data, x='Destination', y='Volume',
                title='',
                color='Volume',
                color_continuous_scale='Greens')
     fig.update_layout(showlegend=False, height=350)
     st.plotly_chart(fig, use_container_width=True)
   
   # Complete Lane Matrix - USING ACTUAL DATA FROM EXCEL
   st.markdown('<p class="chart-title">Complete Lane Network Matrix</p>', unsafe_allow_html=True)
   
   # Create matrix from actual lane data
   all_lanes = tms_data['lanes_list']
   
   # Get unique origins and destinations
   origins = sorted(set(lane['Origin'] for lane in all_lanes))
   destinations = sorted(set(lane['Destination'] for lane in all_lanes))
   
   # Create empty matrix
   matrix = pd.DataFrame(0, index=origins, columns=destinations)
   
   # Fill matrix with actual volumes
   for lane in all_lanes:
    matrix.loc[lane['Origin'], lane['Destination']] = lane['Volume']
   
   # Create heatmap
   fig = px.imshow(matrix, 
                  labels=dict(x="Destination", y="Origin", color="Volume"),
                  title="",
                  color_continuous_scale='YlOrRd',
                  aspect='auto')
   fig.update_layout(height=600)
   st.plotly_chart(fig, use_container_width=True)
   
   # All trade lanes sorted by volume
   st.markdown('<p class="chart-title">All Trade Lanes by Volume</p>', unsafe_allow_html=True)
   
   lanes_df = pd.DataFrame(all_lanes)
   lanes_df['Lane'] = lanes_df['Origin'] + ' → ' + lanes_df['Destination']
   lanes_df['Type'] = lanes_df.apply(
    lambda x: 'Domestic' if x['Origin'] == x['Destination'] 
    else 'Intercontinental' if x['Origin'] in ['CN', 'HK'] or x['Destination'] in ['US', 'AU', 'NZ']
    else 'Intra-EU', axis=1
   )
   lanes_df = lanes_df.sort_values('Volume', ascending=False)
   
   # Show top 20 lanes
   top_lanes = lanes_df.head(20)
   
   fig = px.bar(top_lanes, x='Lane', y='Volume',
              color='Type',
              title='Top 20 Trade Lanes',
              color_discrete_map={'Intra-EU': '#3182bd', 
                                'Domestic': '#31a354',
                                'Intercontinental': '#de2d26'})
   fig.update_layout(xaxis_tickangle=-45, height=400)
   st.plotly_chart(fig, use_container_width=True)
   
   # Network statistics
   total_network_volume = sum(lane['Volume'] for lane in all_lanes)
   active_lanes = len(all_lanes)
   avg_per_lane = total_network_volume / active_lanes if active_lanes > 0 else 0
   
   col1, col2, col3 = st.columns(3)
   
   with col1:
    st.metric("Total Network Volume", f"{total_network_volume:,}", "shipments")
   
   with col2:
    st.metric("Active Trade Lanes", f"{active_lanes:,}", "routes")
   
   with col3:
    st.metric("Average per Lane", f"{avg_per_lane:.1f}", "shipments")
   
   # Lane details table
   st.markdown("**Detailed Lane Information (Top 30)**")
   lane_table = lanes_df[['Origin', 'Destination', 'Volume', 'Type']].head(30)
   st.dataframe(lane_table, hide_index=True, use_container_width=True)
  
  # Network Insights
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### 🛣️ Network Structure Analysis")
  st.markdown(f"""
  **Network Overview:**
  - **Total volume**: {total_network_volume:,} shipments
  - **Active lanes**: {active_lanes} routes
  - **Average volume per lane**: {avg_per_lane:.1f} shipments
  
  **Hub Analysis:**
  - Primary hub locations based on high origin volumes
  - Key destination markets identified
  - Hub-and-spoke model efficiency confirmed
  
  **Trade Patterns:**
  - Intra-EU dominance shows regional focus
  - Domestic volumes indicate local distribution
  - Intercontinental routes show global reach
  
  **Optimization Opportunities:**
  - Low-volume lanes for consolidation
  - High-volume corridors for dedicated service
  - Imbalanced flows for backhaul optimization
  """)
  st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 6: Executive Report
 with tab6:
  st.markdown('<h2 class="section-header">Executive Summary Report</h2>', unsafe_allow_html=True)
  
  # Report Header
  st.markdown(f"**Report Date**: {datetime.now().strftime('%B %d, %Y')}")
  st.markdown(f"**Reporting Period**: Based on uploaded TMS data")
  st.markdown("**Prepared for**: LFS Amsterdam Management Team")
  
  # Executive Summary
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 1. Executive Summary")
  
  performance_status = "Meeting Targets" if avg_otp >= 95 and profit_margin >= 20 else "Below Targets"
  
  st.markdown(f"""
  LFS Amsterdam operates a **{performance_status}** logistics network processing **{total_services} shipments** 
  across multiple countries. The operation shows {'strong' if performance_status == "Meeting Targets" else 'improving'} 
  operational and financial performance.
  
  **Key Performance Indicators:**
  - **On-Time Performance**: {avg_otp:.1f}% (Target: 95%)
  - **Profit Margin**: {profit_margin:.1f}% (Target: 20%)
  - **Revenue per Shipment**: €{total_revenue/total_services:.2f}
  - **Active Trade Lanes**: {tms_data.get('lanes_list', []).__len__()} routes
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Performance Analysis
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 2. Performance Analysis")
  
  st.markdown(f"""
  **Operational Performance:**
  - OTP at {avg_otp:.1f}% {'exceeds' if avg_otp >= 95 else 'below'} industry standard
  - {total_orders:,} orders tracked with {int(avg_otp/100 * total_orders):,} delivered on time
  - Main delay causes: Customer issues, system errors, delivery challenges
  
  **Financial Performance:**
  - Total revenue: €{total_revenue:,.0f}
  - Total costs: €{total_cost:,.0f}
  - Net profit: €{total_revenue - total_cost:,.0f}
  - Margin of {profit_margin:.1f}% {'above' if profit_margin >= 20 else 'below'} 20% target
  
  **Network Efficiency:**
  - {total_services} shipments across {len(tms_data.get('country_volumes', {}))} countries
  - {len(tms_data.get('service_volumes', {}))} service types in operation
  - Hub-and-spoke model with primary concentration in key markets
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Strategic Recommendations
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 3. Strategic Recommendations")
  
  st.markdown("""
  **Immediate Actions:**
  1. Address OTP improvement through customer communication and system fixes
  2. Review pricing in loss-making countries
  3. Optimize high-cost operational areas
  
  **Short-term Initiatives:**
  1. Expand in high-margin markets
  2. Consolidate low-volume lanes
  3. Implement operational efficiency programs
  
  **Long-term Strategy:**
  1. Develop secondary hubs for regional coverage
  2. Invest in technology for better tracking and prediction
  3. Build strategic partnerships in growth markets
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Conclusion
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 4. Conclusion")
  
  st.markdown(f"""
  LFS Amsterdam shows {'strong' if performance_status == "Meeting Targets" else 'solid'} operational foundation 
  with clear opportunities for growth and optimization. Focus areas include:
  
  - {'Maintaining' if avg_otp >= 95 else 'Improving'} on-time performance
  - {'Protecting' if profit_margin >= 20 else 'Enhancing'} profit margins
  - Optimizing network efficiency
  - Expanding in profitable markets
  
  Regular monitoring and targeted improvements will drive continued success.
  """)
  st.markdown('</div>', unsafe_allow_html=True)
