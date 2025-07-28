import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')
import base64
from io import BytesIO

# Configure Streamlit page
st.set_page_config(
    page_title="LFS Amsterdam - TMS Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Professional CSS styling
st.markdown("""
<style>
.main-header {
    font-size: 2.5rem;
    font-weight: 700;
    color: #1a1a1a;
    text-align: center;
    margin-bottom: 2rem;
    font-family: 'Arial', sans-serif;
}
.section-header {
    font-size: 1.8rem;
    font-weight: 600;
    color: #2c3e50;
    margin: 2rem 0 1.5rem 0;
    padding: 0.8rem 0;
    border-bottom: 3px solid #3498db;
}
.metric-card {
    background: #ffffff;
    padding: 1.5rem;
    border-radius: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    border-left: 4px solid #3498db;
    margin-bottom: 1rem;
}
.insight-box {
    background: #f8f9fa;
    padding: 1.5rem;
    border-radius: 8px;
    margin: 1.5rem 0;
    border-left: 4px solid #3498db;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
.report-section {
    margin: 2rem 0;
    padding: 2rem;
    background: #ffffff;
    border-radius: 10px;
    box-shadow: 0 2px 5px rgba(0,0,0,0.1);
}
.chart-title {
    font-size: 1.3rem;
    font-weight: 600;
    color: #2c3e50;
    margin-bottom: 1rem;
    text-align: center;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 24px;
}
.stTabs [data-baseweb="tab"] {
    height: 50px;
    padding-left: 20px;
    padding-right: 20px;
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

# Define service types and countries
SERVICE_TYPES = ['CTX', 'CX', 'EF', 'EGD', 'FF', 'RGD', 'ROU', 'SF']
COUNTRIES = ['AT', 'AU', 'BE', 'DE', 'DK', 'ES', 'FR', 'GB', 'IT', 'N1', 'NL', 'NZ', 'SE', 'US']

# QC Name mapping
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
            
            # 2. OTP Data
            if "OTP POD" in excel_sheets:
                otp_df = excel_sheets["OTP POD"].copy()
                # Process columns properly
                if len(otp_df.columns) >= 6:
                    otp_df = otp_df.iloc[:, :6]
                    otp_df.columns = ['TMS_Order', 'QDT', 'POD_DateTime', 'Time_Diff', 'Status', 'QC_Name']
                else:
                    cols = ['TMS_Order', 'QDT', 'POD_DateTime', 'Time_Diff', 'Status'][:len(otp_df.columns)]
                    if len(otp_df.columns) > len(cols):
                        cols.append('QC_Name')
                    otp_df.columns = cols[:len(otp_df.columns)]
                
                otp_df = otp_df.dropna(subset=['TMS_Order'])
                # Convert date columns
                if 'QDT' in otp_df.columns:
                    otp_df['QDT'] = safe_date_conversion(otp_df['QDT'])
                if 'POD_DateTime' in otp_df.columns:
                    otp_df['POD_DateTime'] = safe_date_conversion(otp_df['POD_DateTime'])
                
                data['otp'] = otp_df
            
            # 3. Volume Data - Process the actual Excel structure
            if "Volume per SVC" in excel_sheets:
                volume_df = excel_sheets["Volume per SVC"].copy()
                
                # Process the volume matrix from Excel
                service_volumes = {}
                country_volumes = {}
                service_country_matrix = {}
                
                # Read the actual data structure
                if not volume_df.empty:
                    # Assuming first column is service/country and rest are volumes
                    if len(volume_df.columns) > 1:
                        # Process rows
                        for idx, row in volume_df.iterrows():
                            if pd.notna(row.iloc[0]):
                                country = str(row.iloc[0]).strip()
                                if country in COUNTRIES:
                                    country_total = 0
                                    service_country_matrix[country] = {}
                                    # Process each service column
                                    for i, service in enumerate(SERVICE_TYPES):
                                        if i + 1 < len(row):
                                            val = row.iloc[i + 1]
                                            if pd.notna(val) and val != 0:
                                                service_country_matrix[country][service] = int(val)
                                                country_total += int(val)
                                                if service not in service_volumes:
                                                    service_volumes[service] = 0
                                                service_volumes[service] += int(val)
                                    if country_total > 0:
                                        country_volumes[country] = country_total
                
                # Calculate total volume
                total_vol = sum(country_volumes.values()) if country_volumes else 0
                
                data['service_volumes'] = service_volumes
                data['country_volumes'] = country_volumes
                data['service_country_matrix'] = service_country_matrix
                data['total_volume'] = total_vol
            
            # 4. Lane Usage
            if "Lane usage " in excel_sheets or "Lane usage" in excel_sheets:
                sheet_name = "Lane usage " if "Lane usage " in excel_sheets else "Lane usage"
                lane_df = excel_sheets[sheet_name].copy()
                
                # Process lane matrix
                lanes_list = []
                if not lane_df.empty and len(lane_df.columns) > 1:
                    # First column should be origins
                    origins = lane_df.iloc[:, 0].dropna()
                    # First row should be destinations
                    destinations = lane_df.columns[1:]
                    
                    for idx, origin in enumerate(origins):
                        if pd.notna(origin) and str(origin).strip():
                            origin_str = str(origin).strip()
                            for j, dest in enumerate(destinations):
                                if j + 1 < len(lane_df.columns):
                                    volume = lane_df.iloc[idx, j + 1]
                                    if pd.notna(volume) and volume != 0:
                                        lanes_list.append({
                                            'Origin': origin_str,
                                            'Destination': str(dest).strip(),
                                            'Volume': int(volume)
                                        })
                
                data['lanes'] = lanes_list
                data['lanes_df'] = lane_df
            
            # 5. Cost Sales - Fixed for accurate financial data
            if "cost sales" in excel_sheets:
                cost_df = excel_sheets["cost sales"].copy()
                
                # Map columns based on actual Excel structure
                if len(cost_df.columns) >= 18:
                    expected_cols = ['Order_Date', 'Account', 'Account_Name', 'Office', 'Order_Num', 
                                    'PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost', 'Total_Cost',
                                    'Net_Revenue', 'Currency', 'Diff', 'Gross_Percent', 'Invoice_Num',
                                    'Total_Amount', 'Status', 'PU_Country']
                    
                    cost_df.columns = expected_cols[:len(cost_df.columns)]
                
                # Convert dates
                if 'Order_Date' in cost_df.columns:
                    cost_df['Order_Date'] = safe_date_conversion(cost_df['Order_Date'])
                
                # Clean numeric columns
                numeric_cols = ['PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost', 'Total_Cost', 
                               'Net_Revenue', 'Diff', 'Gross_Percent', 'Total_Amount']
                for col in numeric_cols:
                    if col in cost_df.columns:
                        cost_df[col] = pd.to_numeric(cost_df[col], errors='coerce')
                
                # Remove rows with invalid financial data
                if 'Net_Revenue' in cost_df.columns and 'Total_Cost' in cost_df.columns:
                    cost_df = cost_df.dropna(subset=['Net_Revenue', 'Total_Cost'])
                    # Keep only rows with actual financial activity
                    cost_df = cost_df[(cost_df['Net_Revenue'] != 0) | (cost_df['Total_Cost'] != 0)]
                
                data['cost_sales'] = cost_df
            
            return data
            
        except Exception as e:
            st.error(f"Error processing Excel file: {str(e)}")
            return None
    return None

def generate_pdf_report(tms_data, metrics):
    """Generate a simple text-based report that can be saved as PDF"""
    try:
        report_content = f"""
LFS AMSTERDAM TMS PERFORMANCE REPORT
{'='*50}

Report Date: {datetime.now().strftime('%B %d, %Y')}
Prepared for: LFS Amsterdam Management Team

EXECUTIVE SUMMARY
{'='*50}

LFS Amsterdam processes {metrics['total_services']:,} shipments with the following performance metrics:

KEY PERFORMANCE INDICATORS:
- Total Volume: {metrics['total_services']:,} shipments
- On-Time Performance: {metrics['otp']:.1f}% (Target: 95%)
- Total Revenue: €{metrics['revenue']:,.2f}
- Total Cost: €{metrics['cost']:,.2f}
- Gross Profit: €{metrics['profit']:,.2f}
- Profit Margin: {metrics['margin']:.2f}% (Target: 20%)

PERFORMANCE STATUS:
- OTP: {'✓ Meeting Target' if metrics['otp'] >= 95 else '✗ Below Target - Improvement Needed'}
- Margin: {'✓ Meeting Target' if metrics['margin'] >= 20 else '✗ Below Target - Improvement Needed'}

VOLUME ANALYSIS
{'='*50}
"""
        
        if 'service_volumes' in tms_data:
            report_content += "\nService Distribution:\n"
            for service, volume in sorted(tms_data['service_volumes'].items(), key=lambda x: x[1], reverse=True):
                if volume > 0:
                    percentage = (volume/metrics['total_services']*100) if metrics['total_services'] > 0 else 0
                    report_content += f"- {service}: {volume} shipments ({percentage:.1f}%)\n"
        
        if 'country_volumes' in tms_data:
            report_content += "\nTop Countries by Volume:\n"
            top_countries = sorted([(k, v) for k, v in tms_data['country_volumes'].items()], 
                                 key=lambda x: x[1], reverse=True)[:5]
            for country, volume in top_countries:
                percentage = (volume/metrics['total_services']*100) if metrics['total_services'] > 0 else 0
                report_content += f"- {country}: {volume} shipments ({percentage:.1f}%)\n"
        
        report_content += f"""

FINANCIAL PERFORMANCE
{'='*50}

Revenue Breakdown:
- Total Revenue: €{metrics['revenue']:,.2f}
- Total Costs: €{metrics['cost']:,.2f}
- Gross Profit: €{metrics['profit']:,.2f}
- Profit Margin: {metrics['margin']:.2f}%

Financial Health: {'Strong - Exceeding targets' if metrics['margin'] >= 20 else 'Needs Improvement - Below 20% target'}

ON-TIME PERFORMANCE
{'='*50}

Delivery Performance:
- Total Orders: {metrics['on_time_orders'] + metrics['late_orders']:,}
- On-Time Deliveries: {metrics['on_time_orders']:,} ({metrics['otp']:.1f}%)
- Late Deliveries: {metrics['late_orders']:,} ({100-metrics['otp']:.1f}%)

Performance Status: {'Excellent - Meeting industry standards' if metrics['otp'] >= 95 else f"Below Standard - {95-metrics['otp']:.1f}% improvement needed"}

STRATEGIC RECOMMENDATIONS
{'='*50}

"""
        
        # Add recommendations based on performance
        rec_num = 1
        
        if metrics['otp'] < 95:
            report_content += f"{rec_num}. IMPROVE ON-TIME PERFORMANCE\n"
            report_content += "   - Address system errors in delivery time calculations\n"
            report_content += "   - Enhance customer communication to reduce parameter changes\n"
            report_content += "   - Implement predictive analytics for better planning\n\n"
            rec_num += 1
        
        if metrics['margin'] < 20:
            report_content += f"{rec_num}. ENHANCE PROFITABILITY\n"
            report_content += "   - Review pricing for low-margin routes\n"
            report_content += "   - Optimize cost structure, particularly high-cost components\n"
            report_content += "   - Consider automation to reduce manual costs\n\n"
            rec_num += 1
        
        report_content += f"{rec_num}. OPTIMIZE NETWORK STRUCTURE\n"
        report_content += "   - Strengthen high-volume trade corridors\n"
        report_content += "   - Evaluate secondary hub opportunities\n"
        report_content += "   - Develop underserved markets with growth potential\n"
        
        report_content += f"""

CONCLUSION
{'='*50}

LFS Amsterdam demonstrates {'strong' if metrics['otp'] >= 95 and metrics['margin'] >= 20 else 'developing'} operational 
capabilities with clear opportunities for growth. Continued focus on operational excellence 
and strategic expansion will drive sustainable success.

Next Steps:
1. Review and prioritize recommendations
2. Develop implementation roadmaps
3. Establish KPI monitoring frameworks
4. Schedule quarterly business reviews

{'='*50}
End of Report
"""
        
        return report_content
        
    except Exception as e:
        st.error(f"Error generating report: {str(e)}")
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
total_diff = 0
profit_margin = 0
total_services = 0
on_time_orders = 0
late_orders = 0

if tms_data is not None:
    # Calculate key metrics
    total_services = tms_data.get('total_volume', 0)
    if total_services == 0 and 'service_volumes' in tms_data:
        total_services = sum(tms_data['service_volumes'].values())
    
    # OTP metrics
    if 'otp' in tms_data and not tms_data['otp'].empty:
        otp_df = tms_data['otp']
        if 'Status' in otp_df.columns:
            status_counts = otp_df['Status'].value_counts()
            total_orders = status_counts.sum()
            if 'ON TIME' in status_counts:
                on_time_orders = status_counts['ON TIME']
            if 'LATE' in status_counts:
                late_orders = status_counts['LATE']
            avg_otp = (on_time_orders / total_orders * 100) if total_orders > 0 else 0
    
    # Financial metrics - Calculate for BILLED orders only
    if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
        cost_df = tms_data['cost_sales'].copy()
        
        # Remove the total row (row 128) and filter for BILLED only
        if len(cost_df) > 127:
            cost_df = cost_df.iloc[:127]
        
        if 'Status' in cost_df.columns:
            cost_df = cost_df[cost_df['Status'] == 'BILLED']
        
        if 'Net_Revenue' in cost_df.columns:
            total_revenue = cost_df['Net_Revenue'].sum()
        if 'Total_Cost' in cost_df.columns:
            total_cost = cost_df['Total_Cost'].sum()
        if 'Diff' in cost_df.columns:
            total_diff = cost_df['Diff'].sum()
        else:
            total_diff = total_revenue - total_cost
            
        if total_revenue > 0:
            profit_margin = (total_diff / total_revenue * 100)

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
            otp_delta = f"{avg_otp-95:.1f}%" if avg_otp != 0 else "N/A"
            st.metric("⏱️ OTP Rate", f"{avg_otp:.1f}%", otp_delta)
        
        with col3:
            st.metric("💰 Revenue", f"€{total_revenue:,.0f}", "total")
        
        with col4:
            margin_delta = f"{profit_margin-20:.1f}%" if profit_margin != 0 else "N/A"
            st.metric("📈 Margin", f"{profit_margin:.1f}%", margin_delta)
        
        # Performance Summary
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown("### 📊 Performance Metrics")
            
            # Volume breakdown
            if 'service_volumes' in tms_data:
                st.markdown("**Service Distribution:**")
                for service, volume in sorted(tms_data['service_volumes'].items(), key=lambda x: x[1], reverse=True):
                    if volume > 0:
                        percentage = (volume/total_services*100) if total_services > 0 else 0
                        st.write(f"- {service}: {volume} ({percentage:.1f}%)")
            
            st.markdown('</div>', unsafe_allow_html=True)
    
    # TAB 5: Lane Network - Corrected with actual data structure
    with tab5:
        st.markdown('<h2 class="section-header">Lane Network & Route Analysis</h2>', unsafe_allow_html=True)
        
        if 'lanes' in tms_data and tms_data['lanes']:
            lanes_data = tms_data['lanes']
            
            # Process lanes data for visualization
            all_origins = set()
            all_destinations = set()
            lane_volumes = {}
            
            for lane in lanes_data:
                origin = lane['Origin']
                dest = lane['Destination']
                volume = lane['Volume']
                
                all_origins.add(origin)
                all_destinations.add(dest)
                lane_key = f"{origin}->{dest}"
                lane_volumes[lane_key] = volume
            
            # Calculate statistics
            total_lane_volume = sum(lane['Volume'] for lane in lanes_data)
            active_lanes = len(lanes_data)
            avg_volume_per_lane = total_lane_volume / active_lanes if active_lanes > 0 else 0
            
            # Key metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Lane Volume", f"{total_lane_volume:,}")
            with col2:
                st.metric("Active Lanes", f"{active_lanes:,}")
            with col3:
                st.metric("Avg per Lane", f"{avg_volume_per_lane:.1f}")
            with col4:
                st.metric("Countries Connected", f"{len(all_origins | all_destinations)}")
            
            # Origin and Destination Analysis
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<p class="chart-title">Top Origin Countries</p>', unsafe_allow_html=True)
                
                # Calculate origin volumes
                origin_volumes = {}
                for lane in lanes_data:
                    origin = lane['Origin']
                    if origin not in origin_volumes:
                        origin_volumes[origin] = 0
                    origin_volumes[origin] += lane['Volume']
                
                origin_df = pd.DataFrame(list(origin_volumes.items()), 
                                       columns=['Origin', 'Volume'])
                origin_df = origin_df.sort_values('Volume', ascending=False).head(10)
                
                fig = px.bar(origin_df, x='Origin', y='Volume',
                           text='Volume',
                           color='Volume',
                           color_continuous_scale='Blues',
                           title='')
                fig.update_traces(texttemplate='%{text}', textposition='outside')
                fig.update_layout(showlegend=False, height=350)
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.markdown('<p class="chart-title">Top Destination Countries</p>', unsafe_allow_html=True)
                
                # Calculate destination volumes
                dest_volumes = {}
                for lane in lanes_data:
                    dest = lane['Destination']
                    if dest not in dest_volumes:
                        dest_volumes[dest] = 0
                    dest_volumes[dest] += lane['Volume']
                
                dest_df = pd.DataFrame(list(dest_volumes.items()), 
                                     columns=['Destination', 'Volume'])
                dest_df = dest_df.sort_values('Volume', ascending=False).head(10)
                
                fig = px.bar(dest_df, x='Destination', y='Volume',
                           text='Volume',
                           color='Volume',
                           color_continuous_scale='Greens',
                           title='')
                fig.update_traces(texttemplate='%{text}', textposition='outside')
                fig.update_layout(showlegend=False, height=350)
                st.plotly_chart(fig, use_container_width=True)
            
            # Lane Matrix Heatmap
            st.markdown('<p class="chart-title">Complete Lane Network Matrix</p>', unsafe_allow_html=True)
            
            # Create matrix
            origins = sorted(list(all_origins))
            destinations = sorted(list(all_destinations))
            
            matrix = pd.DataFrame(0, index=origins, columns=destinations)
            
            for lane in lanes_data:
                matrix.loc[lane['Origin'], lane['Destination']] = lane['Volume']
            
            # Remove empty rows and columns
            matrix = matrix.loc[(matrix.sum(axis=1) != 0), (matrix.sum(axis=0) != 0)]
            
            fig = px.imshow(matrix, 
                          labels=dict(x="Destination", y="Origin", color="Volume"),
                          title="",
                          color_continuous_scale='YlOrRd',
                          aspect='auto',
                          text_auto=True)
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)
            
            # Top Trade Lanes
            st.markdown('<p class="chart-title">Top 20 Trade Corridors</p>', unsafe_allow_html=True)
            
            lanes_df = pd.DataFrame(lanes_data)
            lanes_df['Lane'] = lanes_df['Origin'] + ' → ' + lanes_df['Destination']
            lanes_df['Type'] = lanes_df.apply(
                lambda x: 'Domestic' if x['Origin'] == x['Destination'] 
                else 'Intercontinental' if x['Origin'] in ['CN', 'HK', 'US', 'AU', 'NZ'] or x['Destination'] in ['US', 'AU', 'NZ']
                else 'Intra-EU', axis=1
            )
            lanes_df = lanes_df.sort_values('Volume', ascending=False).head(20)
            
            fig = px.bar(lanes_df, x='Lane', y='Volume',
                      text='Volume',
                      color='Type',
                      title='',
                      color_discrete_map={'Intra-EU': '#3498db', 
                                        'Domestic': '#27ae60',
                                        'Intercontinental': '#e74c3c'})
            fig.update_traces(texttemplate='%{text}', textposition='outside')
            fig.update_layout(xaxis_tickangle=-45, height=400)
            st.plotly_chart(fig, use_container_width=True)
            
            # Network Analysis
            st.markdown('<div class="insight-box">', unsafe_allow_html=True)
            st.markdown("### 🛣️ Lane Network Analysis")
            
            # Hub analysis
            if origin_volumes:
                top_origin = max(origin_volumes, key=origin_volumes.get)
                st.markdown(f"• **Primary Hub**: {top_origin} with {origin_volumes[top_origin]} outbound shipments ({origin_volumes[top_origin]/total_lane_volume*100:.1f}% of network)")
            
            # Lane concentration
            top_10_volume = sum(lane['Volume'] for lane in sorted(lanes_data, key=lambda x: x['Volume'], reverse=True)[:10])
            st.markdown(f"• **Lane Concentration**: Top 10 lanes handle {top_10_volume/total_lane_volume*100:.1f}% of volume")
            
            # Network balance
            domestic_volume = sum(lane['Volume'] for lane in lanes_data if lane['Origin'] == lane['Destination'])
            international_volume = total_lane_volume - domestic_volume
            st.markdown(f"• **Network Split**: {international_volume/total_lane_volume*100:.1f}% international, {domestic_volume/total_lane_volume*100:.1f}% domestic")
            
            st.markdown('</div>', unsafe_allow_html=True)
    
    # TAB 6: Executive Report with PDF Export
    with tab6:
        st.markdown('<h2 class="section-header">Executive Summary Report</h2>', unsafe_allow_html=True)
        
        # Add PDF export button at the top
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("📥 Generate PDF Report", type="primary", use_container_width=True):
                with st.spinner("Generating PDF report..."):
                    # Prepare metrics for PDF
                    pdf_metrics = {
                        'total_services': total_services,
                        'otp': avg_otp,
                        'revenue': total_revenue,
                        'cost': total_cost,
                        'profit': total_diff,
                        'margin': profit_margin,
                        'on_time_orders': on_time_orders,
                        'late_orders': late_orders
                    }
                    
                    # Generate PDF
                    pdf_data = generate_pdf_report(tms_data, pdf_metrics)
                    
                    if pdf_data:
                        # Create download button for text file
                        b64 = base64.b64encode(pdf_data.encode()).decode()
                        file_name = f"LFS_Amsterdam_TMS_Report_{datetime.now().strftime('%Y%m%d')}.txt"
                        href = f'<a href="data:text/plain;base64,{b64}" download="{file_name}" style="display: inline-block; padding: 0.5rem 1rem; background-color: #3498db; color: white; text-decoration: none; border-radius: 4px;">📄 Download Report</a>'
                        st.markdown(href, unsafe_allow_html=True)
                        st.success("✅ Report generated successfully!")
                    else:
                        st.error("Failed to generate PDF report")
        
        # Report metadata
        report_date = datetime.now().strftime('%B %d, %Y')
        st.markdown(f"**Report Date**: {report_date}")
        st.markdown(f"**Reporting Period**: Based on uploaded TMS data")
        st.markdown("**Prepared for**: LFS Amsterdam Management Team")
        
        # Executive Summary
        st.markdown('<div class="report-section">', unsafe_allow_html=True)
        st.markdown("## 1. Executive Summary")
        
        performance_status = "Meeting Targets" if avg_otp >= 95 and profit_margin >= 20 else "Below Targets"
        
        st.markdown(f"""
        LFS Amsterdam operates a **{performance_status}** logistics network processing **{total_services:,} shipments** 
        across **{len(COUNTRIES)} countries**. The operation demonstrates a hub-and-spoke model with Amsterdam as the 
        primary distribution center.
        
        **Key Performance Indicators:**
        - **On-Time Performance**: {avg_otp:.1f}% (Target: 95%) - {'✅ Exceeding' if avg_otp >= 95 else '⚠️ Below'} target
        - **Profit Margin**: {profit_margin:.2f}% (Target: 20%) - {'✅ Healthy' if profit_margin >= 20 else '⚠️ Needs improvement'}
        - **Total Revenue**: €{total_revenue:,.2f} (Billed orders only)
        - **Network Coverage**: {active_lanes} active trade lanes
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Volume Analysis Summary
        st.markdown('<div class="report-section">', unsafe_allow_html=True)
        st.markdown("## 2. Volume & Service Analysis")
        
        if 'service_volumes' in tms_data and 'country_volumes' in tms_data:
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Service Portfolio Performance:**")
                # Get top 3 services
                top_services = sorted([(k, v) for k, v in tms_data['service_volumes'].items() if v > 0], 
                                    key=lambda x: x[1], reverse=True)[:3]
                for service, volume in top_services:
                    percentage = (volume/total_services*100) if total_services > 0 else 0
                    st.write(f"- {service}: {volume} shipments ({percentage:.1f}%)")
            
            with col2:
                st.markdown("**Geographic Distribution:**")
                # Get top 3 countries
                top_countries = sorted([(k, v) for k, v in tms_data['country_volumes'].items()], 
                                     key=lambda x: x[1], reverse=True)[:3]
                for country, volume in top_countries:
                    percentage = (volume/total_services*100) if total_services > 0 else 0
                    st.write(f"- {country}: {volume} shipments ({percentage:.1f}%)")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Financial Performance Summary
        st.markdown('<div class="report-section">', unsafe_allow_html=True)
        st.markdown("## 3. Financial Performance")
        
        st.markdown(f"""
        **Financial Overview (Billed Orders):**
        - **Total Revenue**: €{total_revenue:,.2f}
        - **Total Costs**: €{total_cost:,.2f}
        - **Gross Profit**: €{total_diff:,.2f}
        - **Gross Margin**: {profit_margin:.2f}%
        
        The financial analysis shows {'strong profitability' if profit_margin >= 20 else 'margin pressure'} with 
        {'sustainable' if profit_margin >= 15 else 'concerning'} profit levels. Cost structure analysis reveals 
        opportunities for optimization in operational expenses.
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Operational Performance Summary
        st.markdown('<div class="report-section">', unsafe_allow_html=True)
        st.markdown("## 4. Operational Excellence")
        
        st.markdown(f"""
        **On-Time Performance Analysis:**
        - **Total Orders Tracked**: {total_orders:,}
        - **On-Time Deliveries**: {on_time_orders:,} ({avg_otp:.1f}%)
        - **Late Deliveries**: {late_orders:,} ({100-avg_otp:.1f}%)
        
        Primary delay causes include customer-related issues, system errors, and delivery execution challenges. 
        {'Current performance exceeds industry standards.' if avg_otp >= 95 else f'Improvement of {95-avg_otp:.1f}% needed to meet industry standards.'}
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Strategic Recommendations
        st.markdown('<div class="report-section">', unsafe_allow_html=True)
        st.markdown("## 5. Strategic Recommendations")
        
        recommendations = []
        
        # OTP recommendations
        if avg_otp < 95:
            recommendations.append("**1. Improve On-Time Performance**")
            recommendations.append("   - Implement predictive analytics for delivery planning")
            recommendations.append("   - Address system errors in QDT calculations")
            recommendations.append("   - Enhance customer communication protocols")
        
        # Financial recommendations
        if profit_margin < 20:
            recommendations.append("**2. Enhance Profitability**")
            recommendations.append("   - Review pricing strategy for low-margin routes")
            recommendations.append("   - Optimize cost structure, particularly in high-cost categories")
            recommendations.append("   - Consider automation for manual processes")
        
        # Network recommendations
        recommendations.append("**3. Optimize Network Structure**")
        recommendations.append("   - Strengthen high-volume corridors")
        recommendations.append("   - Evaluate secondary hub opportunities")
        recommendations.append("   - Develop underserved markets with growth potential")
        
        for rec in recommendations:
            st.markdown(rec)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Conclusion
        st.markdown('<div class="report-section">', unsafe_allow_html=True)
        st.markdown("## 6. Conclusion")
        
        st.markdown(f"""
        LFS Amsterdam demonstrates {'strong' if performance_status == "Meeting Targets" else 'developing'} operational 
        capabilities with clear opportunities for growth. The combination of established European presence, 
        diversified service portfolio, and {'robust' if profit_margin >= 20 else 'improving'} financial performance 
        positions the company for sustainable expansion.
        
        **Next Steps:**
        1. Review and prioritize strategic recommendations
        2. Develop detailed implementation roadmaps
        3. Establish KPI monitoring frameworks
        4. Schedule quarterly business reviews
        
        This dashboard provides real-time visibility into operational and financial performance, enabling 
        data-driven decision making for continuous improvement.
        """)
        st.markdown('</div>', unsafe_allow_html=True)

# Add footer
st.markdown("---")
st.markdown("*Dashboard developed for LFS Amsterdam TMS Performance Analysis*")
        
        with col2:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown("### 🎯 Target Achievement")
            
            # OTP status
            if avg_otp >= 95:
                st.success(f"✅ OTP Target Met: {avg_otp:.1f}% (Target: 95%)")
            else:
                st.warning(f"⚠️ OTP Below Target: {avg_otp:.1f}% (Target: 95%)")
            
            # Margin status
            if profit_margin >= 20:
                st.success(f"✅ Margin Target Met: {profit_margin:.1f}% (Target: 20%)")
            else:
                st.warning(f"⚠️ Margin Below Target: {profit_margin:.1f}% (Target: 20%)")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Key insights
        st.markdown('<div class="insight-box">', unsafe_allow_html=True)
        st.markdown("### 💡 Key Insights")
        
        insights = []
        if total_services > 0:
            insights.append(f"• Processing an average of {total_services/30:.0f} shipments per day")
        if 'country_volumes' in tms_data and tms_data['country_volumes']:
            top_country = max(tms_data['country_volumes'], key=tms_data['country_volumes'].get)
            insights.append(f"• {top_country} is the largest market with {tms_data['country_volumes'][top_country]} shipments")
        if avg_otp < 95:
            insights.append(f"• Need to improve OTP by {95-avg_otp:.1f}% to meet target")
        if profit_margin < 20:
            insights.append(f"• Margin improvement of {20-profit_margin:.1f}% required")
        
        for insight in insights:
            st.markdown(insight)
        
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
                            text='Volume',
                            color='Volume', 
                            color_continuous_scale='Blues',
                            title='')
                fig.update_traces(texttemplate='%{text}', textposition='outside')
                fig.update_layout(showlegend=False, height=400, yaxis_title="Shipment Volume")
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.markdown('<p class="chart-title">Country Distribution</p>', unsafe_allow_html=True)
                
                if 'country_volumes' in tms_data and tms_data['country_volumes']:
                    country_data = pd.DataFrame(list(tms_data['country_volumes'].items()), 
                                              columns=['Country', 'Volume'])
                    country_data = country_data.sort_values('Volume', ascending=False)
                    
                    fig = px.bar(country_data, x='Country', y='Volume',
                                text='Volume',
                                color='Volume', 
                                color_continuous_scale='Greens',
                                title='')
                    fig.update_traces(texttemplate='%{text}', textposition='outside')
                    fig.update_layout(showlegend=False, height=400, yaxis_title="Shipment Volume")
                    st.plotly_chart(fig, use_container_width=True)
        
        # Service-Country Matrix
        if 'service_country_matrix' in tms_data and tms_data['service_country_matrix']:
            st.markdown('<p class="chart-title">Service-Country Distribution Matrix</p>', unsafe_allow_html=True)
            
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
            # Remove empty rows and columns
            matrix_df = matrix_df.loc[(matrix_df.sum(axis=1) != 0), (matrix_df.sum(axis=0) != 0)]
            
            if not matrix_df.empty:
                fig = px.imshow(matrix_df.T, 
                              labels=dict(x="Country", y="Service Type", color="Volume"),
                              title="",
                              color_continuous_scale='YlOrRd',
                              aspect='auto',
                              text_auto=True)
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
        
        # Volume Summary Statistics
        st.markdown('<div class="insight-box">', unsafe_allow_html=True)
        st.markdown("### 📊 Volume Analysis Summary")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**Service Concentration**")
            if 'service_volumes' in tms_data and total_services > 0:
                top_3_services = sorted([(k, v) for k, v in tms_data['service_volumes'].items() if v > 0], 
                                      key=lambda x: x[1], reverse=True)[:3]
                top_3_volume = sum([v for _, v in top_3_services])
                st.write(f"Top 3 services: {top_3_volume/total_services*100:.1f}% of volume")
        
        with col2:
            st.markdown("**Geographic Focus**")
            if 'country_volumes' in tms_data and total_services > 0:
                eu_countries = ['AT', 'BE', 'DE', 'DK', 'ES', 'FR', 'GB', 'IT', 'NL', 'SE']
                eu_volume = sum([v for k, v in tms_data['country_volumes'].items() if k in eu_countries])
                st.write(f"European volume: {eu_volume/total_services*100:.1f}%")
        
        with col3:
            st.markdown("**Network Density**")
            if 'service_country_matrix' in tms_data:
                active_combinations = sum([len(services) for services in tms_data['service_country_matrix'].values()])
                possible_combinations = len(COUNTRIES) * len(SERVICE_TYPES)
                st.write(f"Active routes: {active_combinations}/{possible_combinations}")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # TAB 3: OTP Performance - Improved without statistical overview
    with tab3:
        st.markdown('<h2 class="section-header">On-Time Performance Analysis</h2>', unsafe_allow_html=True)
        
        if 'otp' in tms_data and not tms_data['otp'].empty:
            otp_df = tms_data['otp']
            
            # Main OTP metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Orders", f"{total_orders:,}")
            with col2:
                st.metric("On-Time", f"{on_time_orders:,}")
            with col3:
                st.metric("Late", f"{late_orders:,}")
            with col4:
                st.metric("OTP Rate", f"{avg_otp:.1f}%")
            
            # Performance visualization
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<p class="chart-title">Delivery Performance Status</p>', unsafe_allow_html=True)
                
                if 'Status' in otp_df.columns:
                    status_counts = otp_df['Status'].value_counts()
                    
                    fig = px.pie(values=status_counts.values, names=status_counts.index,
                                title='',
                                color_discrete_map={'ON TIME': '#27ae60', 'LATE': '#e74c3c'},
                                hole=0.3)
                    fig.update_traces(textposition='inside', textinfo='percent+label+value')
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.markdown('<p class="chart-title">Delay Categories Analysis</p>', unsafe_allow_html=True)
                
                if 'QC_Name' in otp_df.columns:
                    # Analyze QC reasons
                    qc_analysis = otp_df[otp_df['Status'] == 'LATE']['QC_Name'].dropna()
                    
                    if len(qc_analysis) > 0:
                        # Categorize delays
                        delay_categories = {
                            'Customer Related': 0,
                            'System Errors': 0,
                            'Delivery Issues': 0,
                            'Other': 0
                        }
                        
                        for reason in qc_analysis:
                            reason_str = str(reason)
                            if 'Customer' in reason_str or 'Consignee' in reason_str:
                                delay_categories['Customer Related'] += 1
                            elif 'MNX' in reason_str or 'QDT' in reason_str:
                                delay_categories['System Errors'] += 1
                            elif 'Del' in reason_str or 'delivery' in reason_str.lower():
                                delay_categories['Delivery Issues'] += 1
                            else:
                                delay_categories['Other'] += 1
                        
                        # Create bar chart
                        delay_df = pd.DataFrame(list(delay_categories.items()), 
                                              columns=['Category', 'Count'])
                        delay_df = delay_df[delay_df['Count'] > 0]
                        
                        fig = px.bar(delay_df, x='Category', y='Count',
                                   text='Count',
                                   color='Count',
                                   color_continuous_scale='Reds',
                                   title='')
                        fig.update_traces(texttemplate='%{text}', textposition='outside')
                        fig.update_layout(showlegend=False, height=400)
                        st.plotly_chart(fig, use_container_width=True)
            
            # Time difference analysis
            if 'Time_Diff' in otp_df.columns:
                st.markdown('<p class="chart-title">Delivery Time Performance Distribution</p>', unsafe_allow_html=True)
                
                time_diff_clean = pd.to_numeric(otp_df['Time_Diff'], errors='coerce').dropna()
                
                if len(time_diff_clean) > 0:
                    # Create histogram
                    fig = px.histogram(time_diff_clean, nbins=50,
                                     title='',
                                     labels={'value': 'Days (Early < 0 < Late)', 'count': 'Number of Orders'})
                    fig.add_vline(x=0, line_dash="dash", line_color="green", 
                                annotation_text="On Time")
                    fig.update_traces(marker_color='lightblue')
                    fig.update_layout(height=350)
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Performance zones
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        early_count = len(time_diff_clean[time_diff_clean < -0.5])
                        st.metric("Early Deliveries", f"{early_count:,}", 
                                f"{early_count/len(time_diff_clean)*100:.1f}%")
                    
                    with col2:
                        window_count = len(time_diff_clean[(time_diff_clean >= -0.5) & (time_diff_clean <= 0.5)])
                        st.metric("Within Window", f"{window_count:,}", 
                                f"{window_count/len(time_diff_clean)*100:.1f}%")
                    
                    with col3:
                        late_count = len(time_diff_clean[time_diff_clean > 0.5])
                        st.metric("Late Deliveries", f"{late_count:,}", 
                                f"{late_count/len(time_diff_clean)*100:.1f}%")
            
            # Root cause details
            if 'QC_Name' in otp_df.columns:
                st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                st.markdown("### 🔍 Detailed Delay Analysis")
                
                # Get all delay reasons
                delay_reasons = otp_df[otp_df['Status'] == 'LATE']['QC_Name'].dropna()
                if len(delay_reasons) > 0:
                    reason_counts = delay_reasons.value_counts().head(10)
                    
                    reason_df = pd.DataFrame({
                        'Delay Reason': reason_counts.index,
                        'Occurrences': reason_counts.values,
                        'Percentage': (reason_counts.values / len(delay_reasons) * 100).round(1)
                    })
                    
                    st.dataframe(reason_df, hide_index=True, use_container_width=True)
                
                st.markdown('</div>', unsafe_allow_html=True)
    
    # TAB 4: Financial Analysis - Corrected calculations for BILLED orders only
    with tab4:
        st.markdown('<h2 class="section-header">Financial Performance & Profitability</h2>', unsafe_allow_html=True)
        
        if 'cost_sales' in tms_data and not tms_data['cost_sales'].empty:
            cost_df = tms_data['cost_sales'].copy()
            
            # Filter for BILLED orders only (exclude row 128 which is the total)
            if 'Status' in cost_df.columns:
                # Remove the last row if it's a total row
                if len(cost_df) > 127:
                    cost_df = cost_df.iloc[:127]  # Keep only first 127 rows
                
                # Filter for billed orders
                billed_df = cost_df[cost_df['Status'] == 'BILLED'].copy()
            else:
                billed_df = cost_df.copy()
            
            # Add filters in sidebar for financial analysis
            st.sidebar.markdown("### 💰 Financial Filters")
            
            # Office filter
            if 'Office' in billed_df.columns:
                offices = ['All'] + sorted(billed_df['Office'].dropna().unique().tolist())
                selected_office = st.sidebar.selectbox("Select Office", offices)
                if selected_office != 'All':
                    billed_df = billed_df[billed_df['Office'] == selected_office]
            
            # Country filter
            if 'PU_Country' in billed_df.columns:
                countries = ['All'] + sorted(billed_df['PU_Country'].dropna().unique().tolist())
                selected_country = st.sidebar.selectbox("Select Country", countries)
                if selected_country != 'All':
                    billed_df = billed_df[billed_df['PU_Country'] == selected_country]
            
            # Account filter
            if 'Account_Name' in billed_df.columns:
                accounts = ['All'] + sorted(billed_df['Account_Name'].dropna().unique().tolist())
                selected_account = st.sidebar.selectbox("Select Account", accounts, key='account_filter')
                if selected_account != 'All':
                    billed_df = billed_df[billed_df['Account_Name'] == selected_account]
            
            # Calculate financial metrics for BILLED orders only
            if 'Net_Revenue' in billed_df.columns:
                total_revenue = billed_df['Net_Revenue'].sum()
            else:
                total_revenue = 0
                
            if 'Total_Cost' in billed_df.columns:
                total_cost = billed_df['Total_Cost'].sum()
            else:
                total_cost = 0
            
            if 'Diff' in billed_df.columns:
                total_diff = billed_df['Diff'].sum()
            else:
                total_diff = total_revenue - total_cost
            
            # Calculate profit margin correctly
            if total_revenue > 0:
                profit_margin = (total_diff / total_revenue * 100)
            else:
                profit_margin = 0
            
            # Display current filter status
            filter_status = []
            if selected_office != 'All':
                filter_status.append(f"Office: {selected_office}")
            if selected_country != 'All':
                filter_status.append(f"Country: {selected_country}")
            if selected_account != 'All':
                filter_status.append(f"Account: {selected_account}")
            
            if filter_status:
                st.info(f"🔍 Filtered by: {' | '.join(filter_status)}")
            
            # Financial Overview with correct values
            st.markdown('<p class="chart-title">Financial Summary (Billed Orders Only)</p>', unsafe_allow_html=True)
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("💰 NET Revenue", f"€{total_revenue:,.2f}")
            
            with col2:
                st.metric("📉 Total Cost", f"€{total_cost:,.2f}")
            
            with col3:
                st.metric("💵 Gross Profit (Diff)", f"€{total_diff:,.2f}")
            
            with col4:
                st.metric("📊 Gross Margin %", f"{profit_margin:.2f}%")
            
            # Detailed financial analysis
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<p class="chart-title">Revenue vs Cost Analysis</p>', unsafe_allow_html=True)
                
                financial_summary = pd.DataFrame({
                    'Category': ['Revenue', 'Cost', 'Profit'],
                    'Amount': [total_revenue, total_cost, total_diff]
                })
                
                fig = px.bar(financial_summary, x='Category', y='Amount',
                            text=[f'€{x:,.0f}' for x in financial_summary['Amount']],
                            color='Category',
                            color_discrete_map={'Revenue': '#27ae60', 
                                              'Cost': '#e74c3c',
                                              'Profit': '#3498db'},
                            title='')
                fig.update_traces(texttemplate='%{text}', textposition='outside')
                fig.update_layout(showlegend=False, height=400, yaxis_title="Amount (€)")
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                st.markdown('<p class="chart-title">Cost Breakdown Analysis</p>', unsafe_allow_html=True)
                
                # Cost components
                cost_components = {}
                cost_cols = ['PU_Cost', 'Ship_Cost', 'Man_Cost', 'Del_Cost']
                for col in cost_cols:
                    if col in billed_df.columns:
                        cost_sum = billed_df[col].sum()
                        if cost_sum > 0:
                            cost_components[col.replace('_Cost', '')] = cost_sum
                
                if cost_components:
                    cost_df_pie = pd.DataFrame(list(cost_components.items()), 
                                             columns=['Cost Type', 'Amount'])
                    
                    fig = px.pie(cost_df_pie, values='Amount', names='Cost Type',
                               title='',
                               color_discrete_sequence=px.colors.qualitative.Set3)
                    fig.update_traces(textposition='inside', 
                                    textinfo='percent+label',
                                    hovertemplate='%{label}: €%{value:,.2f}<extra></extra>')
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)
            
            # Profit Margin Distribution - Separate plot as requested
            st.markdown('<p class="chart-title">Profit Margin Distribution Analysis</p>', unsafe_allow_html=True)
            
            if 'Gross_Percent' in billed_df.columns:
                # Convert to percentage
                margin_data = billed_df['Gross_Percent'].dropna() * 100
                
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    # Create histogram
                    fig = px.histogram(margin_data, nbins=30,
                                     title='Distribution of Order Profit Margins',
                                     labels={'value': 'Margin (%)', 'count': 'Number of Orders'})
                    fig.add_vline(x=20, line_dash="dash", line_color="green", 
                                annotation_text="Target 20%")
                    fig.add_vline(x=profit_margin, line_dash="solid", line_color="red", 
                                annotation_text=f"Current Avg: {profit_margin:.1f}%")
                    fig.update_traces(marker_color='lightcoral')
                    fig.update_layout(height=400, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    # Margin statistics
                    st.markdown("**Margin Performance Metrics**")
                    st.write(f"Average Margin: {margin_data.mean():.2f}%")
                    st.write(f"Median Margin: {margin_data.median():.2f}%")
                    st.write(f"Std Deviation: {margin_data.std():.2f}%")
                    
                    profitable_orders = len(margin_data[margin_data > 0])
                    high_margin_orders = len(margin_data[margin_data >= 20])
                    
                    st.write(f"Profitable Orders: {profitable_orders} ({profitable_orders/len(margin_data)*100:.1f}%)")
                    st.write(f"High Margin (≥20%): {high_margin_orders} ({high_margin_orders/len(margin_data)*100:.1f}%)")
            
            # Country/Office Performance Analysis
            if len(billed_df) > 0:
                st.markdown('<p class="chart-title">Performance by Country</p>', unsafe_allow_html=True)
                
                if 'PU_Country' in billed_df.columns:
                    country_performance = billed_df.groupby('PU_Country').agg({
                        'Net_Revenue': 'sum',
                        'Total_Cost': 'sum',
                        'Diff': 'sum'
                    }).round(2)
                    
                    country_performance['Margin %'] = (country_performance['Diff'] / country_performance['Net_Revenue'] * 100).round(2)
                    country_performance = country_performance.sort_values('Net_Revenue', ascending=False)
                    
                    # Display top countries
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig = px.bar(country_performance.reset_index(), 
                                   x='PU_Country', y='Net_Revenue',
                                   title='Revenue by Country',
                                   text=[f'€{x:,.0f}' for x in country_performance['Net_Revenue']],
                                   color='Net_Revenue',
                                   color_continuous_scale='Greens')
                        fig.update_traces(texttemplate='%{text}', textposition='outside')
                        fig.update_layout(showlegend=False, height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        fig = px.bar(country_performance.reset_index(), 
                                   x='PU_Country', y='Margin %',
                                   title='Profit Margin by Country',
                                   text=[f'{x:.1f}%' for x in country_performance['Margin %']],
                                   color='Margin %',
                                   color_continuous_scale='RdYlGn',
                                   range_color=[-20, 40])
                        fig.update_traces(texttemplate='%{text}', textposition='outside')
                        fig.add_hline(y=20, line_dash="dash", line_color="black", 
                                    annotation_text="Target 20%")
                        fig.update_layout(showlegend=False, height=400)
                        st.plotly_chart(fig, use_container_width=True)
                
                # Office Performance if selected
                if 'Office' in billed_df.columns and selected_office == 'All':
                    st.markdown('<p class="chart-title">Performance by Office</p>', unsafe_allow_html=True)
                    
                    office_performance = billed_df.groupby('Office').agg({
                        'Net_Revenue': 'sum',
                        'Total_Cost': 'sum',
                        'Diff': 'sum'
                    }).round(2)
                    
                    office_performance['Margin %'] = (office_performance['Diff'] / office_performance['Net_Revenue'] * 100).round(2)
                    office_performance = office_performance.sort_values('Net_Revenue', ascending=False)
                    
                    # Create table
                    display_office = office_performance.copy()
                    display_office['Revenue'] = display_office['Net_Revenue'].apply(lambda x: f'€{x:,.2f}')
                    display_office['Cost'] = display_office['Total_Cost'].apply(lambda x: f'€{x:,.2f}')
                    display_office['Profit'] = display_office['Diff'].apply(lambda x: f'€{x:,.2f}')
                    display_office = display_office[['Revenue', 'Cost', 'Profit', 'Margin %']]
                    
                    st.dataframe(display_office, use_container_width=True)
            
            # Financial insights
            st.markdown('<div class="insight-box">', unsafe_allow_html=True)
            st.markdown("### 💰 Financial Performance Insights")
            
            insights = []
            
            # Overall performance
            if profit_margin >= 20:
                insights.append(f"✅ **Healthy Margin**: {profit_margin:.2f}% exceeds the 20% target")
            else:
                insights.append(f"⚠️ **Margin Below Target**: {profit_margin:.2f}% needs {20-profit_margin:.2f}% improvement")
            
            # Cost structure
            if cost_components:
                largest_cost = max(cost_components, key=cost_components.get)
                largest_pct = cost_components[largest_cost] / sum(cost_components.values()) * 100
                insights.append(f"• **{largest_cost} costs** represent {largest_pct:.1f}% of total costs - focus area for optimization")
            
            # Revenue per order
            num_orders = len(billed_df)
            if num_orders > 0:
                rev_per_order = total_revenue / num_orders
                insights.append(f"• Average revenue per order: €{rev_per_order:.2f}")
            
            # Country insights
            if 'PU_Country' in billed_df.columns and len(country_performance) > 0:
                best_margin_country = country_performance['Margin %'].idxmax()
                worst_margin_country = country_performance['Margin %'].idxmin()
                insights.append(f"• Best margin country: {best_margin_country} ({country_performance.loc[best_margin_country, 'Margin %']:.1f}%)")
                insights.append(f"• Worst margin country: {worst_margin_country} ({country_performance.loc[worst_margin_country, 'Margin %']:.1f}%)")
            
            for insight in insights:
                st.markdown(insight)
            
            st.markdown('</div>', unsafe_allow_html=True)
