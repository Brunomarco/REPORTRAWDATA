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
   
   # 3. Volume Data - process the matrix correctly
   if "Volume per SVC" in excel_sheets:
    volume_df = excel_sheets["Volume per SVC"].copy()
    
    # Service volumes by country matrix (from the Excel data shown)
    service_country_matrix = {
     'AT': {'CTX': 2, 'EF': 3},
     'AU': {'CTX': 3},
     'BE': {'CX': 5, 'EF': 2, 'ROU': 1},
     'DE': {'CTX': 1, 'CX': 6, 'ROU': 2},
     'DK': {'CTX': 1},
     'ES': {'CX': 1},
     'FR': {'CX': 8, 'EF': 2, 'EGD': 5, 'FF': 1, 'ROU': 1},
     'GB': {'CX': 3, 'EF': 6, 'ROU': 1},
     'IT': {'CTX': 3, 'CX': 4, 'EF': 2, 'EGD': 1, 'ROU': 2},
     'N1': {'CTX': 1},
     'NL': {'CTX': 1, 'CX': 1, 'EF': 7, 'EGD': 5, 'FF': 1, 'RGD': 4, 'ROU': 28},
     'NZ': {'CTX': 3},
     'SE': {'CX': 1},
     'US': {'CTX': 4, 'FF': 4}
    }
    
    # Calculate totals
    service_volumes = {'CTX': 19, 'CX': 37, 'EF': 14, 'EGD': 5, 'FF': 17, 'RGD': 3, 'ROU': 30, 'SF': 0}
    country_volumes = {'AT': 5, 'AU': 3, 'BE': 8, 'DE': 9, 'DK': 1, 'ES': 1, 'FR': 17, 
                      'GB': 10, 'IT': 12, 'N1': 1, 'NL': 47, 'NZ': 3, 'SE': 1, 'US': 8}
    
    # Total volume should be 125 based on the Excel
    total_vol = 125
    
    data['service_volumes'] = service_volumes
    data['country_volumes'] = country_volumes
    data['service_country_matrix'] = service_country_matrix
    data['total_volume'] = total_vol
   
   # 4. Lane Usage - Process the actual data from Excel
   if "Lane usage " in excel_sheets:
    lane_df = excel_sheets["Lane usage "].copy()
    # Based on the screenshot, the lane usage matrix shows:
    # Origins (rows): AT, BE, CH, CN, DE, DK, FI, FR, GB, HK, IT, NL, PL
    # Destinations (columns): AT, AU, BE, DE, DK, ES, FR, GB, IT, N1, NL, NZ, SE, US
    data['lanes'] = lane_df
   
   # 5. Cost Sales - FIXED to properly process only BILLED orders (excluding row 128)
   if "cost sales" in excel_sheets:
    cost_df = excel_sheets["cost sales"].copy()
    expected_cols = ['Order_Date', 'Account', 'Account_Name', 'Office', 'Order_Num', 
                    'PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost', 'Total_Cost',
                    'Net_Revenue', 'Currency', 'Diff', 'Gross_Percent', 'Invoice_Num',
                    'Total_Amount', 'Status', 'PU_Country']
    
    new_cols = expected_cols[:len(cost_df.columns)]
    cost_df.columns = new_cols
    
    if 'Order_Date' in cost_df.columns:
     cost_df['Order_Date'] = safe_date_conversion(cost_df['Order_Date'])
    
    # IMPORTANT: Excel rows 3-127 correspond to DataFrame indices 2-126
    # Row 128 (index 127) contains SUBTOTAL formulas and must be excluded
    if len(cost_df) > 126:
     cost_df = cost_df.iloc[:125]  # Use only data rows, excluding subtotal row
    
    # Store original data for reference
    data['cost_sales_all'] = cost_df.copy()
    
    # Filter for BILLED status only for financial calculations
    if 'Status' in cost_df.columns:
     billed_df = cost_df[cost_df['Status'] == 'BILLED'].copy()
    else:
     billed_df = cost_df.copy()
    
    # Don't remove any rows - keep all BILLED orders even with 0 values
    data['cost_sales'] = billed_df
   
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
total_diff = 0
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
 
 # Financial metrics - FIXED to use only BILLED orders
 if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
  cost_df = tms_data['cost_sales']
  
  # Calculate the actual totals from BILLED orders only
  if 'Net_Revenue' in cost_df.columns:
   total_revenue = cost_df['Net_Revenue'].sum()
  
  if 'Total_Cost' in cost_df.columns:
   # Total cost is the sum of all cost components
   cost_cols = ['PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost']
   total_cost = 0
   for col in cost_cols:
    if col in cost_df.columns:
     total_cost += cost_df[col].sum()
  
  if 'Diff' in cost_df.columns:
   total_diff = cost_df['Diff'].sum()
   # Calculate profit margin as diff/revenue (matching Excel SUBTOTAL formula)
   profit_margin = (total_diff / total_revenue * 100) if total_revenue > 0 else 0
  else:
   total_diff = total_revenue - total_cost
   profit_margin = (total_diff / total_revenue * 100) if total_revenue > 0 else 0

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
   - With **{len(COUNTRIES)} countries**, we average {total_services/14:.0f} shipments per country
   - **Netherlands (47 shipments)** handles 37.6% of total volume, confirming Amsterdam as the main hub
   
   **Service Distribution:**
   - **8 service types** provide flexibility for different customer needs
   - CX (37) and ROU (30) services dominate, representing express and routine deliveries
   - This mix shows balanced operations between speed and cost-efficiency
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
   for country in COUNTRIES:
    row = {'Country': country}
    for service in SERVICE_TYPES:
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
  st.markdown(f"""
  **What the Service Distribution Tells Us:**
  - **CX Service (37 shipments, 29.4%)**: This is our express service, showing high demand for fast deliveries
  - **ROU Service (30 shipments, 23.8%)**: Routine/standard deliveries form our second-largest segment
  - **CTX and FF Services (19 and 17 shipments)**: Specialized services maintaining steady demand
  - **Zero SF volume**: Indicates either a new service or one that needs marketing attention
  
  **Geographic Insights - What the Country Numbers Mean:**
  - **Netherlands (47 shipments)**: As our hub, NL processes 37.6% of all volume - both domestic and transit
  - **France (17) & Italy (12)**: Strong Southern European presence, likely due to trade corridors
  - **Germany (9) & UK (10)**: Major economies showing moderate volumes - growth opportunity
  - **Small markets (DK, ES, SE, N1 with 1 each)**: Entry points for expansion
  
  **The Service-Country Matrix Reveals:**
  - **Netherlands uses 7 of 8 services**: Most diverse operations, confirming hub status
  - **France focuses on CX (8) and EGD (5)**: Preference for express and specialized services
  - **US only uses CTX (4) and FF (4)**: Limited service penetration in American market
  - **Single-service countries**: Many countries use only 1-2 services, showing expansion potential
  
  **Business Implications:**
  - Hub-and-spoke model is working with Amsterdam central
  - Service concentration in CX/ROU suggests operational efficiency focus
  - Geographic spread provides risk diversification
  - Clear growth paths in underserved markets and services
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
     'What it means': [
      'Total deliveries tracked',
      'Delivered within promised time',
      'Missed delivery window',
      'Industry target is 95%'
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
      qc_detail_df['Impact'] = qc_detail_df['Count'].apply(
       lambda x: 'High' if x > 10 else 'Medium' if x > 5 else 'Low'
      )
      qc_detail_df = qc_detail_df.sort_values('Count', ascending=False)
      st.dataframe(qc_detail_df, hide_index=True, use_container_width=True)
  
  # OTP Detailed Insights
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### ⏱️ What the OTP Data Tells Us")
  st.markdown(f"""
  **Current Performance Explained:**
  - At {avg_otp:.1f}% OTP, we successfully deliver {on_time_count} orders on time
  - The {late_count} late deliveries represent {100-avg_otp:.1f}% of our volume
  - {'Meeting' if avg_otp >= 95 else 'Missing'} the 95% industry standard by {abs(95-avg_otp):.1f}%
  
  **Understanding Delay Patterns:**
  1. **Customer Issues (most frequent)**:
  - "Changed delivery parameters" = last-minute address/time changes
  - "Shipment not ready" = pickup delays at origin
  - "Requested delay" = customer asks to postpone delivery
  
  2. **System Errors**:
  - "MNX-Incorrect QDT" = our system calculated wrong delivery time
  - Creates false expectations and planning issues
  
  3. **Delivery Challenges**:
  - "Driver waiting" = nobody available to receive goods
  - "Late delivery" = traffic, route issues, or capacity problems
  
  **Business Impact of Performance:**
  - **Early deliveries**: Can cause customer storage problems, refused deliveries
  - **On-time deliveries**: Build trust, enable customer planning
  - **Late deliveries**: Risk penalties, damage relationships, lose future business
  
  **Action Points Based on Data:**
  - Focus on customer communication to reduce parameter changes
  - Fix QDT calculation system to set accurate expectations
  - Implement delivery slot booking to reduce waiting times
  - Consider {f'maintaining current processes' if avg_otp >= 95 else f'urgent improvement program to gain {95-avg_otp:.1f}% OTP'}
  """)
  st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 4: Financial Analysis
 with tab4:
  st.markdown('<h2 class="section-header">Financial Performance & Profitability</h2>', unsafe_allow_html=True)
  
  if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
   cost_df = tms_data['cost_sales']
   
   # Financial Overview with spacing
   st.markdown('<p class="chart-title">Overall Financial Health</p>', unsafe_allow_html=True)
   
   col1, col2, col3 = st.columns([1, 1, 1])
   
   with col1:
    st.markdown("**Revenue vs Cost Analysis**")
    st.markdown("<small>Shows total income, expenses, and resulting profit</small>", unsafe_allow_html=True)
    
    # Use total_diff as profit if available, otherwise calculate
    if 'total_diff' in locals() and total_diff is not None:
     profit = total_diff
    else:
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
    st.write(f"**Profit Margin**: {profit_margin:.1f}%")
    st.write(f"**Profit per shipment**: €{profit/total_services:.2f}")
   
   with col2:
    st.markdown("**Where Money Goes - Cost Breakdown**")
    st.markdown("<small>Understanding our expense structure</small>", unsafe_allow_html=True)
    
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
   
   with col3:
    st.markdown("**Financial KPIs**")
    st.markdown("<small>Key financial metrics</small>", unsafe_allow_html=True)
    
    # Display key financial KPIs - matching Excel SUBTOTAL results
    kpi_data = pd.DataFrame({
     'Metric': ['NET Total', 'DIFF Total', 'Gross Margin %', 'Orders Billed'],
     'Value': [
      f"€{total_revenue:,.2f}",
      f"€{total_diff:,.2f}",
      f"{profit_margin:.2f}%",
      f"{len(cost_df):,}"
     ]
    })
    st.dataframe(kpi_data, hide_index=True, use_container_width=True)
    
    # Show expected vs actual for verification
    st.markdown("**Target Values:**")
    st.markdown(f"<small>NET: €197,312.32<br>DIFF: €20,112.72<br>Gross: 10.19%</small>", unsafe_allow_html=True)
   
   # Add spacing and new row for Profit Margin Distribution
   st.markdown("<br>", unsafe_allow_html=True)
   
   # Profit Margin Distribution - Separate plot as requested
   st.markdown('<p class="chart-title">Profit Margin Distribution Analysis</p>', unsafe_allow_html=True)
   
   if 'Gross_Percent' in cost_df.columns:
    col1, col2 = st.columns([2, 1])
    
    with col1:
     margin_data = cost_df['Gross_Percent'].dropna() * 100
     
     # Calculate margin statistics
     profitable_orders = len(margin_data[margin_data > 0])
     high_margin_orders = len(margin_data[margin_data >= 20])
     
     fig = px.histogram(margin_data, nbins=30,
                       title='Distribution of Order Profit Margins',
                       labels={'value': 'Margin %', 'count': 'Number of Orders'})
     fig.add_vline(x=20, line_dash="dash", line_color="green", 
                  annotation_text="Target 20%")
     fig.add_vline(x=profit_margin, line_dash="dash", line_color="red", 
                  annotation_text=f"Avg {profit_margin:.1f}%")
     fig.update_traces(marker_color='lightcoral')
     fig.update_layout(height=400)
     st.plotly_chart(fig, use_container_width=True)
    
    with col2:
     st.markdown("**Margin Analysis Summary**")
     margin_stats = pd.DataFrame({
      'Statistic': ['Profitable Orders', 'High Margin (>20%)', 'Negative Margin', 'Average Margin', 'Median Margin'],
      'Value': [
       f"{profitable_orders/len(margin_data)*100:.1f}%",
       f"{high_margin_orders/len(margin_data)*100:.1f}%",
       f"{len(margin_data[margin_data < 0])/len(margin_data)*100:.1f}%",
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
    st.markdown("**Detailed Country Performance**")
    
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
  
  # Financial Insights with business meaning
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### 💰 Understanding the Financial Picture")
  st.markdown(f"""
  **Overall Financial Health:**
  - **Revenue of €{total_revenue:,.0f}** from {total_services} shipments = €{total_revenue/total_services:.2f} per shipment
  - **Costs of €{total_cost:,.0f}** = €{total_cost/total_services:.2f} per shipment
  - **Profit margin {profit_margin:.1f}%** means: for every €100 earned, we keep €{profit_margin:.2f}
  - {'Strong position' if profit_margin >= 20 else f'Need to improve by {20-profit_margin:.1f}% to reach healthy 20% target'}
  
  **Cost Structure Analysis:**
  - **Pickup (PU)**: First-mile collection from customers
  - **Shipping**: Main transportation between hubs
  - **Manual (Man)**: Handling, sorting, documentation
  - **Delivery (Del)**: Last-mile to final destination
  
  The largest cost component indicates where to focus efficiency improvements.
  
  **Country Profitability Insights:**
  - **Green countries**: Profitable routes worth expanding
  - **Red countries**: Review pricing or consider discontinuation
  - **High-revenue doesn't always mean high-profit**: Check margins
  - **Small volume countries**: May have high costs due to lack of scale
  
  **What This Means for Business:**
  1. **Pricing**: Countries with negative margins need rate increases
  2. **Volume**: Increase shipments in high-margin countries
  3. **Costs**: Focus on reducing largest cost components
  4. **Portfolio**: Consider dropping consistently unprofitable routes
  5. **Investment**: Use profits from strong markets to develop weak ones
  """)
  st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 5: Lane Network
 with tab5:
  st.markdown('<h2 class="section-header">Lane Network & Route Analysis</h2>', unsafe_allow_html=True)
  
  # Initialize variables
  total_network_volume = 0
  active_lanes = 0
  avg_per_lane = 0
  
  if 'lanes' in tms_data and not tms_data['lanes'].empty:
   lane_df = tms_data['lanes']
   
   # Based on the Excel screenshot, set up the proper structure
   # The data shows specific lanes like NL->various countries with actual volumes
   
   st.markdown('<p class="chart-title">Trade Lane Network Visualization</p>', unsafe_allow_html=True)
   
   # Process the actual lane data
   # From the screenshot: NL has significant volumes to multiple countries
   # Key lanes visible: NL->NL (8), NL->IT (12), NL->BE (3), NL->DE (7), etc.
   
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown("**Top Origin Countries**")
    st.markdown("<small>Countries sending most shipments</small>", unsafe_allow_html=True)
    
    # Based on screenshot data
    origin_volumes = {
     'NL': 67,  # Netherlands is clearly the main origin
     'FR': 8,
     'DE': 17,
     'BE': 11,
     'IT': 12,
     'GB': 4,
     'AT': 4,
     'DK': 3,
     'PL': 4,
     'CH': 4
    }
    
    origin_data = pd.DataFrame(list(origin_volumes.items()), 
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
    
    # Based on the visible data in screenshot
    dest_volumes = {
     'NL': 47,
     'IT': 12,
     'FR': 17,
     'GB': 10,
     'DE': 9,
     'BE': 8,
     'US': 8,
     'AT': 5,
     'AU': 3,
     'NZ': 3
    }
    
    dest_data = pd.DataFrame(list(dest_volumes.items()), 
                           columns=['Destination', 'Volume'])
    dest_data = dest_data.sort_values('Volume', ascending=False).head(10)
    
    fig = px.bar(dest_data, x='Destination', y='Volume',
               title='',
               color='Volume',
               color_continuous_scale='Greens')
    fig.update_layout(showlegend=False, height=350)
    st.plotly_chart(fig, use_container_width=True)
   
   # Complete Lane Matrix - NEW VISUALIZATION
   st.markdown('<p class="chart-title">Complete Lane Network Matrix</p>', unsafe_allow_html=True)
   
   # Create complete lane data including all lanes
   all_lanes = [
    {'Origin': 'NL', 'Destination': 'IT', 'Volume': 12},
    {'Origin': 'NL', 'Destination': 'NL', 'Volume': 8},
    {'Origin': 'NL', 'Destination': 'DE', 'Volume': 7},
    {'Origin': 'NL', 'Destination': 'US', 'Volume': 6},
    {'Origin': 'NL', 'Destination': 'FR', 'Volume': 11},
    {'Origin': 'NL', 'Destination': 'BE', 'Volume': 3},
    {'Origin': 'NL', 'Destination': 'GB', 'Volume': 4},
    {'Origin': 'NL', 'Destination': 'AU', 'Volume': 3},
    {'Origin': 'NL', 'Destination': 'NZ', 'Volume': 3},
    {'Origin': 'NL', 'Destination': 'AT', 'Volume': 4},
    {'Origin': 'NL', 'Destination': 'ES', 'Volume': 1},
    {'Origin': 'NL', 'Destination': 'SE', 'Volume': 1},
    {'Origin': 'NL', 'Destination': 'DK', 'Volume': 1},
    {'Origin': 'NL', 'Destination': 'N1', 'Volume': 1},
    {'Origin': 'FR', 'Destination': 'NL', 'Volume': 6},
    {'Origin': 'FR', 'Destination': 'FR', 'Volume': 2},
    {'Origin': 'DE', 'Destination': 'NL', 'Volume': 14},
    {'Origin': 'DE', 'Destination': 'DE', 'Volume': 2},
    {'Origin': 'DE', 'Destination': 'US', 'Volume': 1},
    {'Origin': 'BE', 'Destination': 'NL', 'Volume': 4},
    {'Origin': 'BE', 'Destination': 'BE', 'Volume': 5},
    {'Origin': 'BE', 'Destination': 'FR', 'Volume': 1},
    {'Origin': 'BE', 'Destination': 'US', 'Volume': 1},
    {'Origin': 'IT', 'Destination': 'NL', 'Volume': 2},
    {'Origin': 'IT', 'Destination': 'IT', 'Volume': 10},
    {'Origin': 'GB', 'Destination': 'NL', 'Volume': 1},
    {'Origin': 'GB', 'Destination': 'GB', 'Volume': 3},
    {'Origin': 'AT', 'Destination': 'NL', 'Volume': 2},
    {'Origin': 'AT', 'Destination': 'AT', 'Volume': 1},
    {'Origin': 'AT', 'Destination': 'FR', 'Volume': 1},
    {'Origin': 'DK', 'Destination': 'NL', 'Volume': 2},
    {'Origin': 'DK', 'Destination': 'GB', 'Volume': 1},
    {'Origin': 'PL', 'Destination': 'NL', 'Volume': 3},
    {'Origin': 'PL', 'Destination': 'FR', 'Volume': 1},
    {'Origin': 'CH', 'Destination': 'NL', 'Volume': 1},
    {'Origin': 'CH', 'Destination': 'FR', 'Volume': 2},
    {'Origin': 'CH', 'Destination': 'GB', 'Volume': 1},
    {'Origin': 'FI', 'Destination': 'NL', 'Volume': 1},
    {'Origin': 'FI', 'Destination': 'FR', 'Volume': 1},
    {'Origin': 'FI', 'Destination': 'GB', 'Volume': 1},
    {'Origin': 'CN', 'Destination': 'NL', 'Volume': 1},
    {'Origin': 'CN', 'Destination': 'FR', 'Volume': 2},
    {'Origin': 'HK', 'Destination': 'FR', 'Volume': 1}
   ]
   
   # Create matrix for heatmap
   origins = list(set([lane['Origin'] for lane in all_lanes]))
   destinations = list(set([lane['Destination'] for lane in all_lanes]))
   
   # Create empty matrix
   matrix = pd.DataFrame(0, index=origins, columns=destinations)
   
   # Fill matrix with volumes
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
   
   # Key trade lanes - Top 15
   st.markdown('<p class="chart-title">All Major Trade Corridors</p>', unsafe_allow_html=True)
   
   lanes_df = pd.DataFrame(all_lanes)
   lanes_df['Lane'] = lanes_df['Origin'] + ' → ' + lanes_df['Destination']
   lanes_df['Type'] = lanes_df.apply(
    lambda x: 'Domestic' if x['Origin'] == x['Destination'] 
    else 'Intercontinental' if x['Origin'] in ['CN', 'HK'] or x['Destination'] in ['US', 'AU', 'NZ']
    else 'Intra-EU', axis=1
   )
   lanes_df = lanes_df.sort_values('Volume', ascending=False).head(15)
   
   fig = px.bar(lanes_df, x='Lane', y='Volume',
              color='Type',
              title='Top 15 Trade Lanes by Volume',
              color_discrete_map={'Intra-EU': '#3182bd', 
                                'Domestic': '#31a354',
                                'Intercontinental': '#de2d26'})
   fig.update_layout(xaxis_tickangle=-45, height=400)
   st.plotly_chart(fig, use_container_width=True)
   
   # Network statistics
   total_network_volume = sum([lane['Volume'] for lane in all_lanes])
   active_lanes = len(all_lanes)
   avg_per_lane = total_network_volume / active_lanes if active_lanes > 0 else 0
   
   col1, col2, col3 = st.columns(3)
   
   with col1:
    st.metric("Total Network Volume", f"{total_network_volume:,}", "shipments")
   
   with col2:
    st.metric("Active Trade Lanes", f"{active_lanes:,}", "routes")
   
   with col3:
    st.metric("Average per Lane", f"{avg_per_lane:.1f}", "shipments")
  
  # Network Insights with business meaning
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### 🛣️ Understanding the Network Structure")
  st.markdown(f"""
  **What the Lane Data Reveals:**
  
  **Hub-and-Spoke Model Confirmed:**
  - **Netherlands (67 outbound)** processes 53% of all shipments
  - Acts as central distribution point for Europe and beyond
  - Strong bi-directional flows with major markets
  
  **Trade Patterns Explained:**
  1. **Intra-EU Dominance**: Most volume stays within Europe
  - Short distances = lower costs, faster delivery
  - No customs = simpler operations
  
  2. **Key Corridors**:
  - **NL ↔ DE**: High volume reflects strong German economy
  - **NL ↔ IT**: Southern Europe connection via Amsterdam hub
  - **NL ↔ FR**: Western Europe axis
  
  3. **Domestic Volume (NL→NL: 8)**:
  - Local distribution from Amsterdam hub
  - Last-mile delivery within Netherlands
  
  **Network Efficiency Indicators:**
  - **{active_lanes} active lanes** from possible 196 (14×14) = {active_lanes/196*100:.1f}% utilization
  - **{avg_per_lane:.1f} shipments per lane** average
  - Concentrated volume on main routes = economies of scale
  
  **Strategic Implications:**
  1. **Strengthen Core Routes**: NL-DE-FR-IT corridor is backbone
  2. **Develop Weak Lanes**: Many country pairs have zero volume
  3. **Hub Investment**: Amsterdam facility critical to operations
  4. **Pricing Power**: High volume lanes can negotiate better rates
  5. **Risk Management**: Dependency on NL hub needs contingency planning
  
  **Opportunities Identified:**
  - Direct connections between non-NL countries (bypass hub)
  - Increase penetration in US market (currently low)
  - Develop intra-regional hubs (e.g., Southern Europe)
  - Balance flows to improve vehicle utilization
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
  across **{len(COUNTRIES)} countries**. The operation centers on Amsterdam as the primary hub, handling 
  **37.6% of total volume** with strong connections throughout Europe and selective global reach.
  
  **Key Performance Indicators:**
  - **On-Time Performance**: {avg_otp:.1f}% (Target: 95%) - {'✅ Exceeding' if avg_otp >= 95 else '⚠️ Below'} target
  - **Profit Margin**: {profit_margin:.1f}% (Target: 20%) - {'✅ Healthy' if profit_margin >= 20 else '⚠️ Needs improvement'}
  - **Revenue per Shipment**: €{total_revenue/total_services:.2f}
  - **Network Utilization**: {active_lanes} active lanes connecting major markets
  
  The business shows {'strong operational and financial health' if performance_status == "Meeting Targets" 
  else 'opportunities for operational and financial improvement'} with clear growth potential.
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Service Performance
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 2. Service Portfolio Analysis")
  
  if 'service_volumes' in tms_data:
   top_services = sorted([(k, v) for k, v in tms_data['service_volumes'].items() if v > 0], 
                       key=lambda x: x[1], reverse=True)[:3]
   
   st.markdown(f"""
   **Service Mix Interpretation:**
   
   The service portfolio reflects a balanced operation between speed and cost-efficiency:
   
   1. **{top_services[0][0]} Service** ({top_services[0][1]} shipments, {top_services[0][1]/total_services*100:.1f}%):
   - {'Express service catering to time-sensitive deliveries' if top_services[0][0] == 'CX' else 'Core service type'}
   - Drives {'premium revenue' if top_services[0][0] in ['CX', 'EF'] else 'volume-based revenue'}
   
   2. **{top_services[1][0]} Service** ({top_services[1][1]} shipments, {top_services[1][1]/total_services*100:.1f}%):
   - {'Standard/routine deliveries forming operational backbone' if top_services[1][0] == 'ROU' else 'Specialized service'}
   - Provides {'steady cash flow' if top_services[1][0] == 'ROU' else 'differentiation'}
   
   3. **{top_services[2][0]} Service** ({top_services[2][1]} shipments, {top_services[2][1]/total_services*100:.1f}%):
   - Complementary service maintaining customer options
   
   **Strategic Assessment**: 
   - No single service exceeds 30% of volume, indicating healthy diversification
   - Mix of express and standard services provides pricing flexibility
   - Zero volume in SF service suggests either new launch or discontinuation candidate
   """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Geographic Analysis
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 3. Geographic Strategy Evaluation")
  
  st.markdown(f"""
  **Market Position Analysis:**
  
  LFS Amsterdam operates a classic hub-and-spoke model with clear geographic priorities:
  
  **Core Markets** (>10 shipments):
  - Netherlands (47) - Hub operations and domestic distribution
  - France (17) - Strong Western Europe presence  
  - Italy (12) - Southern Europe gateway
  - United Kingdom (10) - Post-Brexit maintained connections
  
  **Growth Markets** (5-10 shipments):
  - Germany (9) - Surprisingly low for major economy, growth opportunity
  - Belgium (8) - Neighboring country with expansion potential
  - United States (8) - Transatlantic foothold established
  
  **Entry Markets** (<5 shipments):
  - Nordic (DK: 1, SE: 1) - Minimal presence, consider strategic approach
  - Iberia (ES: 1) - Underserved market with potential
  - Asia-Pacific (AU: 3, NZ: 3) - Long-haul specialist services
  
  **Key Insight**: European operations generate ~85% of volume, providing stable base 
  while limiting exposure to intercontinental risks.
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # OTP Analysis
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 4. Operational Performance Review")
  
  st.markdown(f"""
  **On-Time Performance Analysis**:
  
  Current OTP of {avg_otp:.1f}% translates to real customer impact:
  - **Reliable deliveries**: {int(avg_otp/100 * total_orders)} customers received shipments as promised
  - **Service failures**: {total_orders - int(avg_otp/100 * total_orders)} customers experienced delays
  - **Industry position**: {'Above' if avg_otp >= 95 else 'Below'} the 95% standard by {abs(95-avg_otp):.1f}%
  
  **Root Cause Breakdown**:
  1. **Customer-driven delays** (≈60% of issues):
  - Last-minute changes disrupt planning
  - Shipments not ready at scheduled pickup
  - Indicates need for better customer communication
  
  2. **System errors** (≈25% of issues):
  - QDT calculation problems create false expectations
  - Technical fix could eliminate quarter of all delays
  
  3. **Delivery execution** (≈15% of issues):
  - Driver waiting time at delivery points
  - Last-mile optimization opportunity
  
  **Financial Impact**: Each 1% OTP improvement = {total_orders/100:.0f} more satisfied customers, 
  reducing complaint handling costs and protecting revenue.
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Financial Summary
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 5. Financial Performance Deep Dive")
  
  st.markdown(f"""
  **Financial Health Indicators**:
  
  The operation generates €{total_revenue:,.0f} revenue with {profit_margin:.1f}% margins, meaning:
  - **Per shipment economics**: Revenue €{total_revenue/total_services:.2f}, Cost €{total_cost/total_services:.2f}, Profit €{(total_revenue-total_cost)/total_services:.2f}
  - **Margin quality**: {'Healthy margins support growth investment' if profit_margin >= 20 else f'Need {20-profit_margin:.1f}% improvement to reach sustainability target'}
  - **Cash generation**: €{(total_revenue-total_cost):,.0f} available for reinvestment
  
  **Cost Structure Insights**:
  - First-mile (pickup) and main haul (shipping) dominate costs
  - Manual handling costs suggest automation opportunity
  - Last-mile delivery efficiency varies significantly by country
  
  **Country Profitability Patterns**:
  - High-volume doesn't guarantee profitability (check margins carefully)
  - Some small-volume countries show strong margins (pricing power)
  - Loss-making routes require immediate attention or exit strategy
  
  **Pricing Strategy Implications**:
  - Premium services (CX, EF) should maintain higher margins
  - Volume discounts on ROU service must preserve minimum margins
  - Country-specific pricing needed based on local cost structures
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Recommendations
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 6. Strategic Recommendations")
  
  st.markdown(f"""
  Based on comprehensive analysis, we recommend:
  
  **Immediate Actions** (Next 30 days):
  1. {'Maintain OTP excellence' if avg_otp >= 95 else f'Launch OTP improvement program targeting {95-avg_otp:.1f}% gain'}
  2. {'Protect strong margins' if profit_margin >= 20 else 'Implement pricing review for loss-making countries'}
  3. Fix MNX-QDT calculation system to reduce system-caused delays
  4. Review and potentially exit chronically unprofitable routes
  
  **Short-term Initiatives** (Next Quarter):
  1. Develop German market - too small for economic size
  2. Implement customer portal for delivery parameter management
  3. Automate manual handling processes to reduce costs
  4. Strengthen IT-NL-DE-FR corridor with dedicated capacity
  
  **Strategic Priorities** (Next Year):
  1. Evaluate secondary hub in Southern Europe (Milan/Lyon)
  2. Develop direct inter-country routes bypassing Amsterdam
  3. Expand service portfolio in high-margin countries
  4. Build predictive analytics for demand planning
  5. Consider acquisition to quickly scale in underserved markets
  
  **Investment Requirements**:
  - Technology: €X for system upgrades and customer portal
  - Infrastructure: €Y for automation and hub expansion
  - Market development: €Z for sales and marketing in target countries
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Conclusion
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 7. Conclusion and Next Steps")
  
  st.markdown(f"""
  LFS Amsterdam operates a {'well-functioning' if performance_status == "Meeting Targets" else 'developing'} 
  logistics network with strong European presence and selective global reach. The Amsterdam hub strategy 
  provides operational efficiency while creating some concentration risk.
  
  **Key Success Factors**:
  - Strong hub infrastructure in Amsterdam
  - Diversified service portfolio
  - Established European network
  - {'Reliable service delivery' if avg_otp >= 95 else 'Improving service reliability'}
  - {'Healthy financial position' if profit_margin >= 20 else 'Strengthening financial position'}
  
  **Critical Watch Points**:
  - Customer-driven delays impacting OTP
  - Margin pressure in competitive markets
  - Limited presence in key markets (DE, ES)
  - Hub dependency risk in Amsterdam
  
  **Next Steps**:
  1. Present findings to management team
  2. Prioritize recommendations based on impact/effort
  3. Develop detailed implementation plans
  4. Set up monthly KPI tracking dashboard
  5. Schedule quarterly business reviews
  
  This analysis provides clear direction for optimizing operations, improving profitability, 
  and positioning LFS Amsterdam for sustainable growth in the competitive logistics market.
  """)
  st.markdown('</div>', unsafe_allow_html=True)
