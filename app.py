# app.py - Executive SCP Savings Dashboard - Clean Version

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import numpy as np
import requests
from io import BytesIO
import base64

# Configure Streamlit page
st.set_page_config(
    page_title="SCP Savings Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced CSS with Flex logo in top-right corner
st.markdown("""
<style>
    /* Hide default Streamlit elements */
    #MainMenu {visibility: hidden;}
    .stDeployButton {display:none;}
    footer {visibility: hidden;}
    
    .main-header {
        font-size: 42px;
        color: #003366;
        text-align: center;
        font-weight: 700;
        margin-bottom: 40px;
        font-family: 'Helvetica Neue', sans-serif;
        letter-spacing: -0.5px;
    }
    
    /* Flex Logo - Fixed to top-right corner */
    .flex-logo-fixed {
        position: fixed;
        top: 10px;
        right: 20px;
        z-index: 999999;
        background: white;
        padding: 8px 12px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        border: 1px solid #e2e8f0;
    }
    
    .flex-logo-img {
        width: 120px;
        height: auto;
        display: block;
    }
    
    /* Fallback text logo */
    .flex-text-logo {
        width: 120px;
        height: 35px;
        background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 100%);
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-weight: 700;
        font-size: 16px;
        font-family: 'Arial', sans-serif;
    }
    
    .summary-container {
        margin: 30px 0;
    }
    
    .tile-row {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 20px;
        margin-bottom: 20px;
        width: 100%;
    }
    
    .summary-tile {
        background: linear-gradient(135deg, #003366 0%, #004080 100%);
        color: white;
        padding: 18px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,51,102,0.2);
        text-align: center;
        transition: all 0.3s ease;
        min-height: 110px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        width: 100%;
    }
    
    .summary-tile:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0,51,102,0.3);
    }
    
    .summary-tile.positive {
        background: linear-gradient(135deg, #1e7e34 0%, #28a745 100%);
    }
    
    .summary-tile.negative {
        background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
    }
    
    .tile-label {
        font-size: 9px;
        font-weight: 600;
        margin-bottom: 6px;
        opacity: 0.9;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        line-height: 1.2;
    }
    
    .tile-amount {
        font-size: 20px;
        font-weight: 700;
        margin: 4px 0;
        line-height: 1;
    }
    
    .tile-desc {
        font-size: 7px;
        opacity: 0.8;
        letter-spacing: 0.3px;
        line-height: 1.1;
    }
    
    .filter-section {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        padding: 25px;
        border-radius: 15px;
        border: 1px solid #cbd5e0;
        margin-bottom: 30px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.05);
    }
    
    .chart-container {
        background-color: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        margin-bottom: 25px;
        border: 1px solid #e2e8f0;
    }
    
    .section-header {
        font-size: 24px;
        color: #003366;
        font-weight: 600;
        margin-bottom: 20px;
        font-family: 'Helvetica Neue', sans-serif;
    }
    
    .data-summary {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        text-align: center;
        font-weight: 500;
    }
    
    .sidebar-clean {
        background: linear-gradient(135deg, #e6f3ff 0%, #ccebff 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 4px solid #003366;
        margin: 20px 0;
        text-align: center;
    }
    
    .status-active {
        color: #28a745;
        font-weight: 600;
        font-size: 16px;
        margin-top: 8px;
    }
</style>
""", unsafe_allow_html=True)

# OneDrive integration functions
def extract_direct_link(onedrive_url):
    """Extract direct download link from OneDrive sharing URL"""
    try:
        if "onedrive.live.com" in onedrive_url:
            if "resid=" in onedrive_url:
                parts = onedrive_url.split("resid=")[1]
                file_id = parts.split("&")[0] if "&" in parts else parts
                direct_url = f"https://onedrive.live.com/download?resid={file_id}"
                return direct_url
        
        if "?" in onedrive_url:
            return onedrive_url + "&download=1"
        else:
            return onedrive_url + "?download=1"
            
    except Exception as e:
        return onedrive_url

@st.cache_data(ttl=300)
def load_onedrive_data(onedrive_url):
    """Load data from OneDrive with robust error handling"""
    methods = [
        ("Direct Download", extract_direct_link(onedrive_url)),
        ("Original URL", onedrive_url),
        ("With Download Param", onedrive_url + ("&download=1" if "?" in onedrive_url else "?download=1"))
    ]
    
    for method_name, url in methods:
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*',
                'Accept-Language': 'en-US,en;q=0.9',
                'Referer': 'https://onedrive.live.com/',
            }
            
            response = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
            
            if (response.status_code == 200 and 
                len(response.content) > 1000 and 
                (response.headers.get('content-type', '').startswith('application/') or 
                 response.content[:4] == b'PK\x03\x04')):
                
                try:
                    df = pd.read_excel(BytesIO(response.content), sheet_name="Savings_WIP_Data")
                    if len(df) > 0:
                        return df, f"OneDrive file loaded successfully via {method_name}"
                except Exception:
                    continue
                    
        except Exception:
            continue
    
    # Fallback to local file
    try:
        df = pd.read_excel("SCP_Savings_FY26_dummy_v3.xlsx", sheet_name="Savings_WIP_Data")
        return df, "Using local file - OneDrive connection unsuccessful"
    except FileNotFoundError:
        return None, "No data source available"
    except Exception as e:
        return None, f"Unable to load data: {str(e)}"

def get_base64_image(image_path):
    """Convert image to base64"""
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except:
        return ""

# Add Flex Logo to center of page
flex_logo_base64 = get_base64_image("flex_logo.png")

if flex_logo_base64:
    st.markdown(f'''
    <div class="flex-logo-centered">
        <img src="data:image/png;base64,{flex_logo_base64}" class="flex-logo-img" alt="Flex">
    </div>
    ''', unsafe_allow_html=True)
else:
    # Fallback text logo
    st.markdown('''
    <div class="flex-logo-centered">
        <div class="flex-text-logo">flex</div>
    </div>
    ''', unsafe_allow_html=True)

# CLEAN SIDEBAR - ONLY Dashboard Status
with st.sidebar:
    st.markdown('''
    <div class="sidebar-clean">
        <div style="font-weight: 600; color: #003366;">📊 Dashboard Status:</div>
        <div class="status-active">Active</div>
    </div>
    ''', unsafe_allow_html=True)

# Dashboard Header
st.markdown('<h1 class="main-header">Executive SCP Savings Dashboard</h1>', unsafe_allow_html=True)

# Hidden OneDrive configuration (not shown in UI)
default_onedrive_url = "https://onedrive.live.com/:x:/g/personal/9E1C07238F947303/EbI62L-aBvdDgxmyFIMOdugB5BoH7r7ATZcU1ywNSR1Psw?resid=9E1C07238F947303!sbfd83ab2069a43f78319b214830e76e8&ithint=file%2Cxlsx&e=bSpS5T"

# Load data
with st.spinner("Loading data..."):
    df, status_message = load_onedrive_data(default_onedrive_url)

# Only show error messages, not success messages
if df is None:
    st.error(status_message)

# Main dashboard content
if df is not None:
    # Data preprocessing
    df = df.rename(columns={
        "Difference (PA)-Finance": "Savings_Finance",
        "Difference (PA) -SCP": "Savings_SCP",
    })
    
    df["Savings_Finance"] = pd.to_numeric(df["Savings_Finance"], errors="coerce").fillna(0)
    df["Savings_SCP"] = pd.to_numeric(df["Savings_SCP"], errors="coerce").fillna(0)
    
    # Date columns
    for col in ["Contract Start", "Contract End"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # FILTERS SECTION
    st.markdown('<div class="filter-section">', unsafe_allow_html=True)
    st.markdown('<h3 class="section-header">📊 Business Intelligence Filters</h3>', unsafe_allow_html=True)
    
    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
    
    with filter_col1:
        if "Contract Start" in df.columns and df["Contract Start"].notna().any():
            min_date = df["Contract Start"].min().date()
            max_date = df["Contract Start"].max().date()
            start_date_filter = st.date_input("Contract Start Date", value=min_date, min_value=min_date, max_value=max_date)
        else:
            start_date_filter = None

    with filter_col2:
        if "Contract End" in df.columns and df["Contract End"].notna().any():
            min_date = df["Contract End"].min().date()
            max_date = df["Contract End"].max().date()
            end_date_filter = st.date_input("Contract End Date", value=max_date, min_value=min_date, max_value=max_date)
        else:
            end_date_filter = None

    with filter_col3:
        if "FY of Savings-Finance" in df.columns:
            finance_options = ["All"] + sorted(df["FY of Savings-Finance"].dropna().unique().tolist())
            finance_filter = st.selectbox("Finance FY", options=finance_options, index=0)
        else:
            finance_filter = "All"

    with filter_col4:
        if "FY of Savings-SCP" in df.columns:
            scp_options = ["All"] + sorted(df["FY of Savings-SCP"].dropna().unique().tolist())
            scp_filter = st.selectbox("SCP FY", options=scp_options, index=0)
        else:
            scp_filter = "All"

    # Domain filter
    if "Domain" in df.columns:
        domain_options = ["All Domains"] + sorted(df["Domain"].dropna().unique().tolist())
        domain_filter = st.selectbox("🏢 Business Domain", options=domain_options, index=0)
    else:
        domain_filter = "All Domains"
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Apply filters
    filtered_df = df.copy()
    
    if start_date_filter and "Contract Start" in df.columns:
        filtered_df = filtered_df[filtered_df["Contract Start"] >= pd.Timestamp(start_date_filter)]
    
    if end_date_filter and "Contract End" in df.columns:
        filtered_df = filtered_df[filtered_df["Contract End"] <= pd.Timestamp(end_date_filter)]
    
    if finance_filter != "All":
        filtered_df = filtered_df[filtered_df["FY of Savings-Finance"] == finance_filter]
    
    if scp_filter != "All":
        filtered_df = filtered_df[filtered_df["FY of Savings-SCP"] == scp_filter]
    
    if domain_filter != "All Domains":
        filtered_df = filtered_df[filtered_df["Domain"] == domain_filter]

    # Calculate metrics
    total_finance = filtered_df["Savings_Finance"].sum()
    positive_finance = filtered_df.loc[filtered_df["Savings_Finance"] > 0, "Savings_Finance"].sum()
    negative_finance = filtered_df.loc[filtered_df["Savings_Finance"] < 0, "Savings_Finance"].sum()
    
    total_scp = filtered_df["Savings_SCP"].sum()
    positive_scp = filtered_df.loc[filtered_df["Savings_SCP"] > 0, "Savings_SCP"].sum()
    negative_scp = filtered_df.loc[filtered_df["Savings_SCP"] < 0, "Savings_SCP"].sum()

    # EXECUTIVE SUMMARY
    st.markdown('<div class="summary-container">', unsafe_allow_html=True)
    st.markdown('<h2 class="section-header">📈 Executive Summary</h2>', unsafe_allow_html=True)
    
    # Finance metrics row
    st.markdown('<div class="tile-row">', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="summary-tile">
            <div class="tile-label">Net Finance Impact</div>
            <div class="tile-amount">${total_finance:,.0f}</div>
            <div class="tile-desc">Total Portfolio</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="summary-tile positive">
            <div class="tile-label">Finance Upside</div>
            <div class="tile-amount">${positive_finance:,.0f}</div>
            <div class="tile-desc">Value Creation</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div class="summary-tile negative">
            <div class="tile-label">Finance Exposure</div>
            <div class="tile-amount">${abs(negative_finance):,.0f}</div>
            <div class="tile-desc">Risk Management</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # SCP metrics row
    st.markdown('<div class="tile-row">', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="summary-tile">
            <div class="tile-label">Net SCP Impact</div>
            <div class="tile-amount">${total_scp:,.0f}</div>
            <div class="tile-desc">Total Portfolio</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="summary-tile positive">
            <div class="tile-label">SCP Upside</div>
            <div class="tile-amount">${positive_scp:,.0f}</div>
            <div class="tile-desc">Value Creation</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div class="summary-tile negative">
            <div class="tile-label">SCP Exposure</div>
            <div class="tile-amount">${abs(negative_scp):,.0f}</div>
            <div class="tile-desc">Risk Management</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # STRATEGIC ANALYTICS
    st.markdown('<h2 class="section-header">📊 Strategic Analytics</h2>', unsafe_allow_html=True)

    mckinsey_colors = ['#001f3f', '#003366', '#004080', '#0066cc', '#3399ff', '#66b3ff', '#99ccff']
    
    chart_col1, chart_col2 = st.columns(2)
    
    # Finance chart
    with chart_col1:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        
        if "FY of Savings-Finance" in filtered_df.columns:
            finance_data = filtered_df.groupby("FY of Savings-Finance")["Savings_Finance"].sum().reset_index()
            finance_data = finance_data.sort_values("FY of Savings-Finance")
            
            colors = [mckinsey_colors[i % len(mckinsey_colors)] for i in range(len(finance_data))]
            
            fig_finance = go.Figure()
            
            for i, row in finance_data.iterrows():
                fig_finance.add_trace(go.Bar(
                    x=[row["FY of Savings-Finance"]],
                    y=[row["Savings_Finance"]],
                    marker_color=colors[i],
                    text=f"${row['Savings_Finance']:,.0f}",
                    textposition="outside",
                    showlegend=False,
                    hovertemplate=f"<b>FY:</b> {row['FY of Savings-Finance']}<br><b>Impact:</b> ${row['Savings_Finance']:,.0f}<extra></extra>"
                ))
            
            fig_finance.update_layout(
                title={'text': "Finance Impact by Fiscal Year", 'x': 0.5, 'font': {'size': 18, 'color': '#003366'}},
                xaxis_title="Fiscal Year",
                yaxis_title="Financial Impact ($)",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(size=11, color="#003366"),
                height=450
            )
            
            st.plotly_chart(fig_finance, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # SCP chart
    with chart_col2:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        
        if "FY of Savings-SCP" in filtered_df.columns:
            scp_data = filtered_df.groupby("FY of Savings-SCP")["Savings_SCP"].sum().reset_index()
            scp_data = scp_data.sort_values("FY of Savings-SCP")
            
            colors = [mckinsey_colors[i % len(mckinsey_colors)] for i in range(len(scp_data))]
            
            fig_scp = go.Figure()
            
            for i, row in scp_data.iterrows():
                fig_scp.add_trace(go.Bar(
                    x=[row["FY of Savings-SCP"]],
                    y=[row["Savings_SCP"]],
                    marker_color=colors[i],
                    text=f"${row['Savings_SCP']:,.0f}",
                    textposition="outside",
                    showlegend=False,
                    hovertemplate=f"<b>FY:</b> {row['FY of Savings-SCP']}<br><b>Impact:</b> ${row['Savings_SCP']:,.0f}<extra></extra>"
                ))
            
            fig_scp.update_layout(
                title={'text': "SCP Impact by Fiscal Year", 'x': 0.5, 'font': {'size': 18, 'color': '#003366'}},
                xaxis_title="Fiscal Year",
                yaxis_title="SCP Impact ($)",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(size=11, color="#003366"),
                height=450
            )
            
            st.plotly_chart(fig_scp, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    # DOMAIN ANALYSIS
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h3 class="section-header">🏢 Business Domain Analysis</h3>', unsafe_allow_html=True)
    
    if "Domain" in filtered_df.columns:
        domain_data = filtered_df.groupby("Domain")["Savings_Finance"].sum().reset_index()
        domain_data = domain_data.sort_values("Savings_Finance", ascending=True)
        
        domain_colors = [mckinsey_colors[i % len(mckinsey_colors)] for i in range(len(domain_data))]
        
        fig_domain = go.Figure(go.Bar(
            x=domain_data["Savings_Finance"],
            y=domain_data["Domain"],
            orientation='h',
            marker=dict(color=domain_colors, line=dict(color='white', width=1)),
            text=[f"${val:,.0f}" for val in domain_data["Savings_Finance"]],
            textposition="outside"
        ))
        
        fig_domain.update_layout(
            title={'text': "Financial Impact by Business Domain", 'x': 0.5, 'font': {'size': 18, 'color': '#003366'}},
            xaxis_title="Financial Impact ($)",
            yaxis_title="Business Domain",
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(size=11, color="#003366"),
            height=max(400, len(domain_data) * 45),
            showlegend=False
        )
        
        st.plotly_chart(fig_domain, use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

    # PORTFOLIO OVERVIEW
    st.markdown('<h2 class="section-header">📋 Portfolio Overview</h2>', unsafe_allow_html=True)
    
    overview_col1, overview_col2, overview_col3, overview_col4 = st.columns(4)
    
    with overview_col1:
        st.metric("Active Contracts", len(filtered_df))
    
    with overview_col2:
        avg_finance = filtered_df["Savings_Finance"].mean()
        st.metric("Avg Finance Impact", f"${avg_finance:,.0f}")
    
    with overview_col3:
        avg_scp = filtered_df["Savings_SCP"].mean()
        st.metric("Avg SCP Impact", f"${avg_scp:,.0f}")
    
    with overview_col4:
        if "Domain" in filtered_df.columns:
            unique_domains = filtered_df["Domain"].nunique()
            st.metric("Business Domains", unique_domains)

    # DATA EXPORT
    st.markdown("### 💾 Data Export")
    
    total_records = len(filtered_df)
    if total_records != len(df):
        st.markdown(f'<div class="data-summary">Analysis Results: {total_records:,} of {len(df):,} contracts</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="data-summary">Complete Portfolio: {total_records:,} contracts analyzed</div>', unsafe_allow_html=True)

    export_col1, export_col2 = st.columns(2)
    
    with export_col1:
        summary_data = pd.DataFrame({
            'Metric': ['Net Finance Impact', 'Finance Upside', 'Finance Exposure', 'Net SCP Impact', 'SCP Upside', 'SCP Exposure'],
            'Value': [total_finance, positive_finance, abs(negative_finance), total_scp, positive_scp, abs(negative_scp)]
        })
        
        st.download_button(
            label="📊 Download Executive Summary",
            data=summary_data.to_csv(index=False),
            file_name=f"executive_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )
    
    with export_col2:
        st.download_button(
            label="📁 Download Portfolio Data",
            data=filtered_df.to_csv(index=False),
            file_name=f"portfolio_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

else:
    # Error state
    st.error("Unable to Connect to OneDrive")
    
    with st.expander("🔧 Connection Troubleshooting"):
        st.markdown("""
        **Common OneDrive Issues:**
        
        1. **Sharing Permissions** - File must be shared with "Anyone with the link can view"
        2. **URL Format** - Must be the sharing link from OneDrive
        3. **File Requirements** - Excel file must contain "Savings_WIP_Data" sheet
        4. **Network Issues** - Check if corporate firewall blocks OneDrive
        """)

# Footer
st.markdown("---")
st.markdown("**Flex Confidential")

