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

def safe_date_conversion(date_series):
 """Safely convert Excel dates"""
 try:
  if date_series.dtype in ['int64', 'float64']:
   return pd.to_datetime(date_series, origin='1899-12-30', unit='D', errors='coerce')
  else:
   return pd.to_datetime(date_series, errors='coerce')
 except:
  return date_series

def safe_float_conversion(value):
 """Safely convert value to float"""
 try:
  if pd.isna(value):
   return 0.0
  if isinstance(value, (int, float)):
   return float(value)
  if isinstance(value, str):
   # Remove currency symbols and spaces
   value = value.replace('€', '').replace('$', '').replace(',', '').strip()
   return float(value) if value else 0.0
  return 0.0
 except:
  return 0.0

@st.cache_data
def load_tms_data(uploaded_file):
 """Load and process TMS Excel file"""
 if uploaded_file is not None:
  try:
   # Read all sheets
   excel_sheets = pd.read_excel(uploaded_file, sheet_name=None)
   data = {}
   
   # 1. Raw Data
   if "AMS RAW DATA" in excel_sheets:
    data['raw_data'] = excel_sheets["AMS RAW DATA"].copy()
   
   # 2. OTP Data
   if "OTP POD" in excel_sheets:
    otp_df = excel_sheets["OTP POD"].copy()
    if len(otp_df.columns) >= 6:
     otp_df = otp_df.iloc[:, :6]
     otp_df.columns = ['TMS_Order', 'QDT', 'POD_DateTime', 'Time_Diff', 'Status', 'QC_Name']
    else:
     cols = ['TMS_Order', 'QDT', 'POD_DateTime', 'Time_Diff', 'Status'][:len(otp_df.columns)]
     otp_df.columns = cols
    otp_df = otp_df.dropna(subset=['TMS_Order'])
    data['otp'] = otp_df
   
   # 3. Volume Data - Based on the screenshot structure
   if "Volume per SVC" in excel_sheets:
    volume_df = excel_sheets["Volume per SVC"].copy()
    
    # Initialize dictionaries
    service_country_matrix = {}
    service_volumes = {svc: 0 for svc in SERVICE_TYPES}
    country_volumes = {country: 0 for country in COUNTRIES}
    
    # Process the volume matrix
    # Assuming the first row contains service types and first column contains countries
    for idx, row in volume_df.iterrows():
     if idx == 0:  # Skip header row if needed
      continue
     
     country = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else None
     if country and country in COUNTRIES:
      service_country_matrix[country] = {}
      
      for col_idx, service in enumerate(volume_df.columns[1:], 1):
       if service in SERVICE_TYPES:
        value = safe_float_conversion(row.iloc[col_idx])
        if value > 0:
         service_country_matrix[country][service] = int(value)
         service_volumes[service] += int(value)
         country_volumes[country] += int(value)
    
    # Remove services/countries with zero volume
    service_volumes = {k: v for k, v in service_volumes.items() if v > 0}
    country_volumes = {k: v for k, v in country_volumes.items() if v > 0}
    
    # Calculate total
    total_vol = sum(service_volumes.values())
    
    data['service_volumes'] = service_volumes
    data['country_volumes'] = country_volumes
    data['service_country_matrix'] = service_country_matrix
    data['total_volume'] = total_vol
   
   # 4. Lane Usage
   if "Lane usage " in excel_sheets:
    lane_df = excel_sheets["Lane usage "].copy()
    
    lanes_list = []
    origin_totals = {}
    dest_totals = {}
    
    # Process lane matrix - first column is origins, first row is destinations
    for idx, row in lane_df.iterrows():
     if idx == 0:  # Skip if header row
      continue
      
     origin = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else None
     if origin:
      for col_idx, col_name in enumerate(lane_df.columns[1:], 1):
       dest = str(col_name) if not pd.isna(col_name) else None
       if dest:
        try:
         value = safe_float_conversion(row.iloc[col_idx])
         if value > 0:
          lanes_list.append({
           'Origin': origin,
           'Destination': dest,
           'Volume': int(value)
          })
          
          # Update totals
          origin_totals[origin] = origin_totals.get(origin, 0) + int(value)
          dest_totals[dest] = dest_totals.get(dest, 0) + int(value)
        except:
         continue
    
    data['lanes_list'] = lanes_list
    data['origin_totals'] = origin_totals
    data['dest_totals'] = dest_totals
    data['lanes'] = lane_df
   
   # 5. Cost Sales
   if "cost sales" in excel_sheets:
    cost_df = excel_sheets["cost sales"].copy()
    
    # Clean column names
    cost_df.columns = [str(col).strip() for col in cost_df.columns]
    
    # Try to identify columns by patterns
    for col in cost_df.columns:
     col_lower = col.lower()
     
     # Revenue columns
     if any(term in col_lower for term in ['revenue', 'net revenue', 'sales']):
      cost_df.rename(columns={col: 'Net_Revenue'}, inplace=True)
     # Total cost
     elif 'total' in col_lower and 'cost' in col_lower:
      cost_df.rename(columns={col: 'Total_Cost'}, inplace=True)
     # Individual costs
     elif 'pu' in col_lower and 'cost' in col_lower:
      cost_df.rename(columns={col: 'PU_Cost'}, inplace=True)
     elif 'ship' in col_lower and 'cost' in col_lower:
      cost_df.rename(columns={col: 'Ship_Cost'}, inplace=True)
     elif 'man' in col_lower and 'cost' in col_lower:
      cost_df.rename(columns={col: 'Man_Cost'}, inplace=True)
     elif 'del' in col_lower and 'cost' in col_lower:
      cost_df.rename(columns={col: 'Del_Cost'}, inplace=True)
     # Gross percent
     elif 'gross' in col_lower and ('%' in col_lower or 'percent' in col_lower):
      cost_df.rename(columns={col: 'Gross_Percent'}, inplace=True)
     # Country
     elif 'country' in col_lower:
      cost_df.rename(columns={col: 'PU_Country'}, inplace=True)
     # Date
     elif 'date' in col_lower:
      cost_df.rename(columns={col: 'Order_Date'}, inplace=True)
    
    # Convert numeric columns
    numeric_cols = ['Net_Revenue', 'Total_Cost', 'PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost', 'Gross_Percent']
    for col in numeric_cols:
     if col in cost_df.columns:
      cost_df[col] = cost_df[col].apply(safe_float_conversion)
    
    # Convert date
    if 'Order_Date' in cost_df.columns:
     cost_df['Order_Date'] = safe_date_conversion(cost_df['Order_Date'])
    
    # Clean data - only keep rows with financial activity
    if 'Net_Revenue' in cost_df.columns and 'Total_Cost' in cost_df.columns:
     cost_df = cost_df[(cost_df['Net_Revenue'] != 0) | (cost_df['Total_Cost'] != 0)]
     cost_df = cost_df.dropna(subset=['Net_Revenue', 'Total_Cost'])
    
    data['cost_sales'] = cost_df
   
   return data
   
  except Exception as e:
   st.error(f"Error processing Excel file: {str(e)}")
   import traceback
   st.error(traceback.format_exc())
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

# Calculate global metrics
avg_otp = 0
total_orders = 0
total_revenue = 0
total_cost = 0
profit_margin = 0
total_services = 0

if tms_data is not None:
 # Calculate key metrics
 total_services = tms_data.get('total_volume', 0)
 
 # OTP metrics
 if 'otp' in tms_data and not tms_data['otp'].empty:
  otp_df = tms_data['otp']
  if 'Status' in otp_df.columns:
   status_series = otp_df['Status'].dropna()
   total_orders = len(status_series)
   on_time_orders = len(status_series[status_series == 'ON TIME'])
   avg_otp = (on_time_orders / total_orders * 100) if total_orders > 0 else 0
 
 # Financial metrics
 if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
  cost_df = tms_data['cost_sales']
  if 'Net_Revenue' in cost_df.columns:
   total_revenue = cost_df['Net_Revenue'].sum()
  if 'Total_Cost' in cost_df.columns:
   total_cost = cost_df['Total_Cost'].sum()
  profit_margin = ((total_revenue - total_cost) / total_revenue * 100) if total_revenue > 0 else 0

# Create tabs
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
   - Active in **{len(tms_data.get('country_volumes', {}))} countries**
   - Operating **{len(tms_data.get('service_volumes', {}))} service types**
   
   **Service Distribution:**
   - Multiple service types provide flexibility for customer needs
   - Mix shows balance between express and standard services
   """)
   st.markdown('</div>', unsafe_allow_html=True)
  
  with col2:
   st.markdown('<div class="insight-box">', unsafe_allow_html=True)
   st.markdown("### 🎯 Performance Interpretation")
   
   if avg_otp >= 95:
    st.markdown(f"""
    ✅ **OTP at {avg_otp:.1f}%** - Exceeding target
    - Delivering {int(avg_otp/100 * total_orders)} out of {total_orders} orders on time
    - Industry-leading performance
    """)
   else:
    st.markdown(f"""
    ⚠️ **OTP at {avg_otp:.1f}%** - Below 95% target
    - Need {int((95-avg_otp)/100 * total_orders)} more on-time deliveries
    - Focus on improvement initiatives
    """)
   
   if profit_margin >= 20:
    st.markdown(f"""
    ✅ **{profit_margin:.1f}% margin** - Healthy profitability
    - Above 20% target threshold
    - Strong financial position
    """)
   else:
    st.markdown(f"""
    ⚠️ **{profit_margin:.1f}% margin** - Below target
    - Need {20-profit_margin:.1f}% improvement
    - Review pricing and costs
    """)
   st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 2: Volume Analysis
 with tab2:
  st.markdown('<h2 class="section-header">Volume Analysis by Service & Country</h2>', unsafe_allow_html=True)
  
  if 'service_volumes' in tms_data and tms_data['service_volumes']:
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown('<p class="chart-title">Service Type Distribution</p>', unsafe_allow_html=True)
    
    service_data = pd.DataFrame(list(tms_data['service_volumes'].items()), 
                              columns=['Service', 'Volume'])
    service_data = service_data[service_data['Volume'] > 0].sort_values('Volume', ascending=False)
    
    fig = px.bar(service_data, x='Service', y='Volume', 
                color='Volume', 
                color_continuous_scale='Blues',
                title='')
    fig.update_layout(showlegend=False, height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Service table
    service_table = service_data.copy()
    service_table['Share %'] = (service_table['Volume'] / service_table['Volume'].sum() * 100).round(1)
    st.dataframe(service_table, hide_index=True, use_container_width=True)
   
   with col2:
    st.markdown('<p class="chart-title">Country Distribution</p>', unsafe_allow_html=True)
    
    if 'country_volumes' in tms_data and tms_data['country_volumes']:
     country_data = pd.DataFrame(list(tms_data['country_volumes'].items()), 
                               columns=['Country', 'Volume'])
     country_data = country_data.sort_values('Volume', ascending=False)
     
     fig = px.bar(country_data, x='Country', y='Volume',
                 color='Volume', 
                 color_continuous_scale='Greens',
                 title='')
     fig.update_layout(showlegend=False, height=400)
     st.plotly_chart(fig, use_container_width=True)
     
     # Country table
     country_table = country_data.copy()
     country_table['Share %'] = (country_table['Volume'] / country_table['Volume'].sum() * 100).round(1)
     st.dataframe(country_table, hide_index=True, use_container_width=True)
  
  # Service-Country Matrix
  if 'service_country_matrix' in tms_data and tms_data['service_country_matrix']:
   st.markdown('<p class="chart-title">Service-Country Matrix</p>', unsafe_allow_html=True)
   
   # Create matrix
   countries = sorted(tms_data['service_country_matrix'].keys())
   services = sorted(set(s for c in tms_data['service_country_matrix'].values() for s in c.keys()))
   
   matrix_data = []
   for country in countries:
    row = {'Country': country}
    for service in services:
     row[service] = tms_data['service_country_matrix'].get(country, {}).get(service, 0)
    matrix_data.append(row)
   
   matrix_df = pd.DataFrame(matrix_data).set_index('Country')
   
   fig = px.imshow(matrix_df.T, 
                  labels=dict(x="Country", y="Service Type", color="Volume"),
                  title="",
                  color_continuous_scale='YlOrRd',
                  aspect='auto')
   fig.update_layout(height=500)
   st.plotly_chart(fig, use_container_width=True)
 
 # TAB 3: OTP Performance
 with tab3:
  st.markdown('<h2 class="section-header">On-Time Performance Analysis</h2>', unsafe_allow_html=True)
  
  if 'otp' in tms_data and not tms_data['otp'].empty:
   otp_df = tms_data['otp']
   
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown('<p class="chart-title">Delivery Performance</p>', unsafe_allow_html=True)
    
    if 'Status' in otp_df.columns:
     status_counts = otp_df['Status'].value_counts()
     
     fig = px.pie(values=status_counts.values, names=status_counts.index,
                 title='',
                 color_discrete_map={'ON TIME': '#2ca02c', 'LATE': '#d62728'})
     fig.update_traces(textposition='inside', textinfo='percent+label')
     st.plotly_chart(fig, use_container_width=True)
    
    # Metrics table
    on_time_count = int(avg_otp/100 * total_orders)
    late_count = total_orders - on_time_count
    
    metrics_data = pd.DataFrame({
     'Metric': ['Total Orders', 'On-Time', 'Late', 'OTP Rate'],
     'Value': [f"{total_orders:,}", f"{on_time_count:,}", f"{late_count:,}", f"{avg_otp:.1f}%"],
     'Target': ['-', '-', '-', '95%']
    })
    st.dataframe(metrics_data, hide_index=True, use_container_width=True)
   
   with col2:
    st.markdown('<p class="chart-title">Delay Causes</p>', unsafe_allow_html=True)
    
    if 'QC_Name' in otp_df.columns:
     # Process QC reasons
     delay_categories = {'Customer Issues': 0, 'System Errors': 0, 'Delivery Problems': 0}
     
     for value in otp_df['QC_Name'].dropna():
      reason = str(value).strip()
      if 'Customer' in reason:
       delay_categories['Customer Issues'] += 1
      elif 'MNX' in reason:
       delay_categories['System Errors'] += 1
      elif reason and reason != 'nan':
       delay_categories['Delivery Problems'] += 1
     
     if sum(delay_categories.values()) > 0:
      fig = px.bar(x=list(delay_categories.keys()), y=list(delay_categories.values()),
                  title='',
                  color=list(delay_categories.values()),
                  color_continuous_scale='Reds')
      fig.update_layout(showlegend=False, xaxis_title='Category', yaxis_title='Count')
      st.plotly_chart(fig, use_container_width=True)
  
  # OTP Summary
  st.markdown('<div class="insight-box">', unsafe_allow_html=True)
  st.markdown("### ⏱️ OTP Performance Summary")
  st.markdown(f"""
  **Performance Metrics:**
  - OTP Rate: {avg_otp:.1f}% {'✅ Above' if avg_otp >= 95 else '⚠️ Below'} 95% target
  - On-time deliveries: {int(avg_otp/100 * total_orders):,} orders
  - Late deliveries: {total_orders - int(avg_otp/100 * total_orders):,} orders
  
  **Improvement Focus:**
  - Address customer-related delays
  - Fix system calculation errors
  - Optimize delivery operations
  """)
  st.markdown('</div>', unsafe_allow_html=True)
 
 # TAB 4: Financial Analysis
 with tab4:
  st.markdown('<h2 class="section-header">Financial Performance & Profitability</h2>', unsafe_allow_html=True)
  
  if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
   cost_df = tms_data['cost_sales']
   
   # First row - Revenue/Cost and Cost Breakdown
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown("**Revenue vs Cost Analysis**")
    
    profit = total_revenue - total_cost
    financial_data = pd.DataFrame({
     'Category': ['Revenue', 'Cost', 'Profit'],
     'Amount': [total_revenue, total_cost, profit]
    })
    
    fig = px.bar(financial_data, x='Category', y='Amount',
                color='Category',
                color_discrete_map={'Revenue': '#2ca02c', 'Cost': '#ff7f0e',
                                  'Profit': '#2ca02c' if profit >= 0 else '#d62728'},
                title='')
    fig.update_layout(showlegend=False, height=350)
    st.plotly_chart(fig, use_container_width=True)
    
    st.write(f"**Total Revenue**: €{total_revenue:,.0f}")
    st.write(f"**Total Cost**: €{total_cost:,.0f}")
    st.write(f"**Net Profit**: €{profit:,.0f}")
   
   with col2:
    st.markdown("**Cost Structure Breakdown**")
    
    cost_components = {}
    for col in ['PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost']:
     if col in cost_df.columns:
      cost_sum = cost_df[col].sum()
      if cost_sum > 0:
       cost_components[col.replace('_Cost', '')] = cost_sum
    
    if cost_components:
     total_costs = sum(cost_components.values())
     
     fig = px.pie(values=list(cost_components.values()), 
                names=list(cost_components.keys()),
                title='')
     fig.update_traces(textposition='inside', textinfo='percent+label')
     fig.update_layout(height=350)
     st.plotly_chart(fig, use_container_width=True)
   
   # Profit Margin Distribution
   st.markdown('<p class="chart-title">Profit Margin Analysis</p>', unsafe_allow_html=True)
   
   if 'Gross_Percent' in cost_df.columns:
    col1, col2 = st.columns(2)
    
    with col1:
     st.markdown("**Margin Distribution**")
     margin_data = cost_df['Gross_Percent'].dropna()
     
     if not margin_data.empty:
      # Convert to percentage if needed
      if margin_data.max() <= 1:
       margin_data = margin_data * 100
      
      fig = px.histogram(margin_data, nbins=30,
                       title='',
                       labels={'value': 'Margin %', 'count': 'Orders'})
      fig.add_vline(x=20, line_dash="dash", line_color="green", 
                  annotation_text="Target 20%")
      fig.add_vline(x=0, line_dash="dash", line_color="red", 
                  annotation_text="Break-even")
      fig.update_traces(marker_color='lightcoral')
      st.plotly_chart(fig, use_container_width=True)
    
    with col2:
     st.markdown("**Margin Statistics**")
     
     if not margin_data.empty:
      profitable = len(margin_data[margin_data > 0])
      high_margin = len(margin_data[margin_data >= 20])
      
      stats = pd.DataFrame({
       'Metric': ['Total Orders', 'Profitable', 'Loss-making', 'High Margin (≥20%)', 
                  'Avg Margin', 'Median Margin'],
       'Value': [
        f"{len(margin_data):,}",
        f"{profitable:,} ({profitable/len(margin_data)*100:.1f}%)",
        f"{len(margin_data) - profitable:,}",
        f"{high_margin:,} ({high_margin/len(margin_data)*100:.1f}%)",
        f"{margin_data.mean():.1f}%",
        f"{margin_data.median():.1f}%"
       ]
      })
      st.dataframe(stats, hide_index=True, use_container_width=True)
   
   # Country Performance
   if 'PU_Country' in cost_df.columns:
    st.markdown('<p class="chart-title">Country Financial Performance</p>', unsafe_allow_html=True)
    
    country_fin = cost_df.groupby('PU_Country').agg({
     'Net_Revenue': 'sum',
     'Total_Cost': 'sum'
    }).round(2)
    
    country_fin['Profit'] = country_fin['Net_Revenue'] - country_fin['Total_Cost']
    country_fin['Margin %'] = ((country_fin['Profit'] / country_fin['Net_Revenue']) * 100).round(1)
    country_fin = country_fin.sort_values('Net_Revenue', ascending=False)
    
    col1, col2 = st.columns(2)
    
    with col1:
     st.markdown("**Revenue by Country**")
     revenue_data = country_fin[country_fin['Net_Revenue'] > 0]
     
     fig = px.bar(revenue_data.reset_index(), x='PU_Country', y='Net_Revenue',
                title='', color='Net_Revenue',
                color_continuous_scale='Greens')
     fig.update_layout(showlegend=False, height=400)
     st.plotly_chart(fig, use_container_width=True)
    
    with col2:
     st.markdown("**Profit/Loss by Country**")
     profit_data = country_fin.reset_index()
     profit_data['Color'] = profit_data['Profit'].apply(lambda x: 'Profit' if x >= 0 else 'Loss')
     
     fig = px.bar(profit_data, x='PU_Country', y='Profit',
                title='', color='Color',
                color_discrete_map={'Profit': '#2ca02c', 'Loss': '#d62728'})
     fig.update_layout(showlegend=False, height=400)
     st.plotly_chart(fig, use_container_width=True)
    
    # Country table
    st.markdown("**Detailed Country Performance**")
    display_fin = country_fin.copy()
    display_fin['Status'] = display_fin['Profit'].apply(
     lambda x: '🟢 Profitable' if x > 0 else '🔴 Loss-making'
    )
    display_fin = display_fin[['Net_Revenue', 'Total_Cost', 'Profit', 'Margin %', 'Status']]
    display_fin.columns = ['Revenue (€)', 'Cost (€)', 'Profit (€)', 'Margin (%)', 'Status']
    display_fin = display_fin.round(0)
    st.dataframe(display_fin, use_container_width=True)
 
 # TAB 5: Lane Network
 with tab5:
  st.markdown('<h2 class="section-header">Lane Network & Route Analysis</h2>', unsafe_allow_html=True)
  
  if 'lanes_list' in tms_data and tms_data['lanes_list']:
   # Origin and Destination summaries
   col1, col2 = st.columns(2)
   
   with col1:
    st.markdown("**Top Origin Countries**")
    if 'origin_totals' in tms_data and tms_data['origin_totals']:
     origin_data = pd.DataFrame(list(tms_data['origin_totals'].items()), 
                              columns=['Origin', 'Volume'])
     origin_data = origin_data.sort_values('Volume', ascending=False).head(10)
     
     fig = px.bar(origin_data, x='Origin', y='Volume',
                title='', color='Volume',
                color_continuous_scale='Blues')
     fig.update_layout(showlegend=False, height=350)
     st.plotly_chart(fig, use_container_width=True)
   
   with col2:
    st.markdown("**Top Destination Countries**")
    if 'dest_totals' in tms_data and tms_data['dest_totals']:
     dest_data = pd.DataFrame(list(tms_data['dest_totals'].items()), 
                            columns=['Destination', 'Volume'])
     dest_data = dest_data.sort_values('Volume', ascending=False).head(10)
     
     fig = px.bar(dest_data, x='Destination', y='Volume',
                title='', color='Volume',
                color_continuous_scale='Greens')
     fig.update_layout(showlegend=False, height=350)
     st.plotly_chart(fig, use_container_width=True)
   
   # Lane Matrix
   st.markdown('<p class="chart-title">Complete Lane Network Matrix</p>', unsafe_allow_html=True)
   
   # Create matrix from lane data
   lanes_df = pd.DataFrame(tms_data['lanes_list'])
   
   # Get unique origins and destinations
   origins = sorted(lanes_df['Origin'].unique())
   destinations = sorted(lanes_df['Destination'].unique())
   
   # Create pivot table for matrix
   matrix = lanes_df.pivot_table(index='Origin', columns='Destination', 
                                values='Volume', fill_value=0, aggfunc='sum')
   
   # Create heatmap
   fig = px.imshow(matrix, 
                  labels=dict(x="Destination", y="Origin", color="Volume"),
                  title="Origin-Destination Volume Matrix",
                  color_continuous_scale='YlOrRd',
                  aspect='auto')
   fig.update_layout(height=600)
   st.plotly_chart(fig, use_container_width=True)
   
   # Top lanes bar chart
   st.markdown('<p class="chart-title">Top Trade Lanes</p>', unsafe_allow_html=True)
   
   # Add lane description
   lanes_df['Lane'] = lanes_df['Origin'] + ' → ' + lanes_df['Destination']
   lanes_df['Type'] = lanes_df.apply(
    lambda x: 'Domestic' if x['Origin'] == x['Destination'] 
    else 'Intercontinental' if x['Destination'] in ['US', 'AU', 'NZ'] or x['Origin'] in ['CN', 'HK']
    else 'Regional', axis=1
   )
   
   # Sort by volume and show top 20
   top_lanes = lanes_df.sort_values('Volume', ascending=False).head(20)
   
   fig = px.bar(top_lanes, x='Lane', y='Volume',
              color='Type',
              title='Top 20 Trade Lanes by Volume',
              color_discrete_map={'Regional': '#3182bd', 
                                'Domestic': '#31a354',
                                'Intercontinental': '#de2d26'})
   fig.update_layout(xaxis_tickangle=-45, height=400)
   st.plotly_chart(fig, use_container_width=True)
   
   # Network statistics
   total_volume = lanes_df['Volume'].sum()
   num_lanes = len(lanes_df)
   avg_volume = total_volume / num_lanes if num_lanes > 0 else 0
   
   col1, col2, col3 = st.columns(3)
   with col1:
    st.metric("Total Network Volume", f"{total_volume:,}")
   with col2:
    st.metric("Active Lanes", f"{num_lanes:,}")
   with col3:
    st.metric("Avg Volume/Lane", f"{avg_volume:.1f}")
   
   # Detailed lane table
   st.markdown("**Lane Details (Top 30)**")
   lane_details = lanes_df[['Origin', 'Destination', 'Volume', 'Type']].sort_values(
    'Volume', ascending=False).head(30)
   st.dataframe(lane_details, hide_index=True, use_container_width=True)
  
  else:
   st.warning("No lane data available in the uploaded file.")
 
 # TAB 6: Executive Report
 with tab6:
  st.markdown('<h2 class="section-header">Executive Summary Report</h2>', unsafe_allow_html=True)
  
  # Report Header
  st.markdown(f"**Report Date**: {datetime.now().strftime('%B %d, %Y')}")
  st.markdown("**Prepared for**: LFS Amsterdam Management Team")
  
  # Executive Summary
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 1. Executive Summary")
  
  performance_status = "Meeting Targets" if avg_otp >= 95 and profit_margin >= 20 else "Below Targets"
  
  st.markdown(f"""
  LFS Amsterdam operates a **{performance_status}** logistics network with:
  - **{total_services:,} shipments** processed
  - **{len(tms_data.get('country_volumes', {}))} countries** served
  - **{len(tms_data.get('service_volumes', {}))} service types** offered
  
  **Key Metrics:**
  - OTP: {avg_otp:.1f}% (Target: 95%)
  - Profit Margin: {profit_margin:.1f}% (Target: 20%)
  - Revenue/Shipment: €{total_revenue/total_services:.2f if total_services > 0 else 0}
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Performance Analysis
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 2. Performance Analysis")
  
  st.markdown(f"""
  **Operational Performance:**
  - On-time delivery rate: {avg_otp:.1f}%
  - Total orders tracked: {total_orders:,}
  - Service reliability: {'Above' if avg_otp >= 95 else 'Below'} industry standard
  
  **Financial Performance:**
  - Total revenue: €{total_revenue:,.0f}
  - Operating costs: €{total_cost:,.0f}
  - Net profit: €{total_revenue - total_cost:,.0f}
  
  **Network Efficiency:**
  - Active trade lanes: {len(tms_data.get('lanes_list', []))}
  - Geographic coverage: {len(tms_data.get('country_volumes', {}))} countries
  - Service diversity: {len(tms_data.get('service_volumes', {}))} types
  """)
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Recommendations
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 3. Strategic Recommendations")
  
  recommendations = []
  
  if avg_otp < 95:
   recommendations.append("• Implement OTP improvement program targeting customer communication and system accuracy")
  if profit_margin < 20:
   recommendations.append("• Review pricing strategy and cost optimization opportunities")
  if len(tms_data.get('lanes_list', [])) > 0:
   recommendations.append("• Optimize low-volume lanes and strengthen high-volume corridors")
  
  if recommendations:
   st.markdown("**Priority Actions:**")
   for rec in recommendations:
    st.markdown(rec)
  else:
   st.markdown("**Continue current strategies while exploring growth opportunities**")
  
  st.markdown('</div>', unsafe_allow_html=True)
  
  # Conclusion
  st.markdown('<div class="report-section">', unsafe_allow_html=True)
  st.markdown("## 4. Conclusion")
  
  st.markdown(f"""
  LFS Amsterdam demonstrates {'strong' if performance_status == "Meeting Targets" else 'developing'} 
  operational capabilities with clear paths for optimization:
  
  - {'Maintain' if avg_otp >= 95 else 'Improve'} service reliability
  - {'Protect' if profit_margin >= 20 else 'Enhance'} profit margins
  - Optimize network efficiency
  - Expand strategic markets
  
  Regular monitoring and targeted improvements will ensure continued success.
  """)
  st.markdown('</div>', unsafe_allow_html=True)
