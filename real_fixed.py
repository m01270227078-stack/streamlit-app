import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import openpyxl

# =================================================================================
# Page Configuration & Initial Setup
# =================================================================================
st.set_page_config(
    page_title="Visionary Analytics Dashboard 3.0",
    page_icon="👑", # Changed icon to better match the Gold theme
    layout="wide",
    initial_sidebar_state="expanded"
)

# =================================================================================
# ✨ Golden Age CSS & Styling ✨
# =================================================================================
st.markdown("""
<style>
/* --- Font Imports --- */
@import url('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css');
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=Roboto:wght@300;400;700&family=Cinzel:wght@400;700&display=swap'); /* Added Cinzel for a more luxurious feel */

/* --- CSS Variables for Theming (GOLDEN THEME) --- */
:root {
    --primary-color: #ffd700; /* Gold */
    --secondary-color: #daa520; /* Goldenrod */
    --background-color: #1a150a; /* Dark Brown/Black */
    --card-background-color: rgba(26, 21, 10, 0.8);
    --border-color: rgba(255, 215, 0, 0.3);
    --glow-color: rgba(255, 215, 0, 0.5);
    --text-color: #f0e68c; /* Khaki/Light Gold */
    --header-font: 'Cinzel', serif; /* Changed header font */
    --body-font: 'Roboto', sans-serif;
}

/* --- General Body Styling --- */
body {
    background-color: var(--background-color);
    color: var(--text-color);
    font-family: var(--body-font);
}
.main .block-container {
    padding: 2rem 3rem;
    /* Subtle geometric pattern in gold */
    background-image: linear-gradient(45deg, rgba(255, 215, 0, 0.02) 25%, transparent 25%),
                      linear-gradient(-45deg, rgba(255, 215, 0, 0.02) 25%, transparent 25%),
                      linear-gradient(45deg, transparent 75%, rgba(255, 215, 0, 0.02) 75%),
                      linear-gradient(-45deg, transparent 75%, rgba(255, 215, 0, 0.02) 75%);
    background-size: 20px 20px;
}

/* --- Sidebar (Glassmorphism & Glow) - STABLE SELECTOR --- */
[data-testid="stSidebar"] {
    background-color: rgba(26, 21, 10, 0.6);
    backdrop-filter: blur(15px);
    border-right: 1px solid var(--border-color);
    box-shadow: 0 0 20px rgba(255, 215, 0, 0.1);
}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
    color: #ffffff;
    font-family: var(--header-font);
    text-shadow: 0 0 5px var(--glow-color);
}

/* --- Metric Cards (Golden Style) --- */
@property --angle {
    syntax: '<angle>';
    initial-value: 0deg;
    inherits: false;
}
.metric-card {
    background: var(--card-background-color);
    border: 1px solid transparent;
    border-radius: 15px;
    padding: 25px;
    position: relative;
    overflow: hidden;
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}
.metric-card::before {
    content: '';
    position: absolute;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: conic-gradient(
        from var(--angle),
        transparent 0%,
        var(--primary-color) 20%,
        transparent 25%
    );
    animation: rotate 6s linear infinite;
    opacity: 0.3; /* Reduced opacity for subtlety */
}
.metric-card:hover {
    transform: translateY(-10px) scale(1.02);
    box-shadow: 0 0 35px var(--glow-color);
}
.metric-card-content {
    position: relative;
    z-index: 1;
    background: var(--card-background-color);
    border-radius: 12px;
    padding: 20px;
    margin: 2px;
}
.metric-card .icon { font-size: 2.5rem; color: var(--primary-color); margin-bottom: 10px; text-shadow: 0 0 10px var(--glow-color); }
.metric-card h3 { color: #f0f0f0; font-size: 1.1rem; font-family: var(--header-font); }
.metric-card .value { color: var(--primary-color); font-size: 2.8rem; font-weight: 900; font-family: var(--header-font); text-shadow: 0 0 8px var(--glow-color); }
.metric-card .delta { font-size: 1rem; }
/* Changed delta colors for better contrast on dark/gold theme */
.metric-card .delta .fa-arrow-down { color: #ff6347; } /* Tomato Red */
.metric-card .delta .fa-arrow-up { color: #3cb371; } /* Medium Sea Green */

/* --- Page Header --- */
.page-header {
    text-align: center;
    margin-bottom: 2rem;
    animation: fadeIn 1s ease-in-out;
}
.page-header h2 {
    font-family: var(--header-font);
    font-size: 3.5rem;
    color: #fff;
    text-shadow: 0 0 15px var(--glow-color), 0 0 25px var(--secondary-color);
}
.page-header p {
    color: var(--text-color);
    font-size: 1.2rem;
}

/* --- Chart Styling --- */
.stPlotlyChart {
    border-radius: 12px;
    padding: 10px;
    background: transparent;
    transition: box-shadow 0.3s ease, transform 0.3s ease;
}
.stPlotlyChart:hover {
    box-shadow: 0 0 25px rgba(255, 215, 0, 0.4);
    transform: scale(1.01);
}
/* Style for Tabs */
.st-emotion-cache-1qg26k p {
    font-size: 1.1rem;
    color: var(--text-color);
}
/* --- Animations --- */
@keyframes rotate { to { --angle: 360deg; } }
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
</style>
""", unsafe_allow_html=True)
# =================================================================================
# Data Loading & Generation
# =================================================================================
@st.cache_data
def generate_demo_data():
    """Generates a sample DataFrame for demonstration."""
    np.random.seed(42)
    start_date = datetime(2023, 1, 1)
    end_date = datetime(2025, 9, 30)
    date_range = pd.to_datetime(pd.date_range(start=start_date, end=end_date, freq='D'))
    num_records = 1500
    
    data = {
        'Order Date': np.random.choice(date_range, num_records),
        'Customer ID': [f'CUST-{i}' for i in np.random.randint(100, 500, num_records)],
        'Customer Name': [f'Customer {chr(65+np.random.randint(0,26))}{np.random.randint(10,99)}' for _ in range(num_records)],
        'Product Name': [f'Product {chr(88+np.random.randint(0,3))}{np.random.randint(1,10)}' for _ in range(num_records)],
        'Category': np.random.choice(['Electronics', 'Office Supplies', 'Furniture'], num_records, p=[0.3, 0.4, 0.3]),
        'Sales': np.random.uniform(50, 1500, num_records).round(2),
        'Quantity': np.random.randint(1, 10, num_records),
        'Duration_new': np.random.uniform(1, 10, num_records).round(1)
    }
    df = pd.DataFrame(data)
    
    sub_cat_map = {
        'Electronics': ['Phones', 'Accessories', 'Laptops'],
        'Office Supplies': ['Paper', 'Binders', 'Pens'],
        'Furniture': ['Chairs', 'Tables', 'Desks']
    }
    df['Sub-Category'] = df['Category'].apply(lambda x: np.random.choice(sub_cat_map[x]))
    
    us_states = ['CA', 'TX', 'NY', 'FL', 'IL', 'PA', 'OH', 'GA', 'NC', 'MI', 'WA', 'VA', 'AZ', 'MA', 'CO']
    df['State'] = np.random.choice(us_states, size=len(df))
    
    return df

@st.cache_data
def load_data(uploaded_file):
    """Loads and preprocesses data from an Excel file."""
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [col.strip() for col in df.columns]
        rename_map = {
            "Customer ID": "Customer ID", "Customer Name": "Customer Name", "Order Date": "Order Date",
            "Product Name": "Product Name", "Category": "Category", "Sub-Category": "Sub-Category",
            "Sales": "Sales", "Quantity": "Quantity", "duration": "Duration", "Shipping Duration": "Duration"
        }
        df = df.rename(columns=lambda c: rename_map.get(c, c))
        df["Order Date"] = pd.to_datetime(df["Order Date"], errors='coerce')
        if "Duration" in df.columns:
            df["Duration"] = pd.to_numeric(df["Duration"], errors='coerce')
        df.dropna(subset=['Order Date', 'Customer ID', 'Sales', 'Category'], inplace=True)
        if 'State' not in df.columns:
            us_states = ['CA', 'TX', 'NY', 'FL', 'IL', 'PA', 'OH', 'GA', 'NC', 'MI', 'WA', 'VA', 'AZ', 'MA', 'CO']
            np.random.seed(42)
            df['State'] = np.random.choice(us_states, size=len(df))
        return df
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        return None

# =================================================================================
# Helper & UI Functions
# =================================================================================
def display_metric(title, value, icon_class, delta=None):
    delta_html = ""
    if delta is not None and pd.notna(delta) and delta != float('inf'):
        arrow = "fa-arrow-down" if delta < 0 else "fa-arrow-up"
        delta_html = f'<p class="delta" style="font-size: 1rem;">{delta:.2f}% <i class="fas {arrow}"></i> vs previous period</p>'
    
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-card-content">
            <div class="icon"><i class="fas {icon_class}"></i></div>
            <h3>{title}</h3>
            <p class="value">{value}</p>
            {delta_html}
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_page_header(title, subtitle):
    st.markdown(f"""
    <div class="page-header">
        <h2>{title}</h2>
        <p>{subtitle}</p>
    </div>
    """, unsafe_allow_html=True)

@st.cache_data
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# =================================================================================
# Page Rendering Functions
# =================================================================================

def render_overview_page(df, df_prev):
    render_page_header("Dashboard Overview", "A high-level view of key performance indicators.")

    # --- Metrics ---
    total_sales = df['Sales'].sum()
    prev_sales = df_prev['Sales'].sum() if not df_prev.empty else 0
    sales_delta = ((total_sales - prev_sales) / prev_sales * 100) if prev_sales > 0 else float('inf')

    total_customers = df['Customer ID'].nunique()
    prev_customers = df_prev['Customer ID'].nunique() if not df_prev.empty else 0
    customers_delta = ((total_customers - prev_customers) / prev_customers * 100) if prev_customers > 0 else float('inf')

    avg_duration = df['duration_new'].mean() if 'duration_new' in df.columns and not df.empty else 0
    prev_duration = df_prev['Duration'].mean() if 'Duration' in df_prev.columns and not df_prev.empty else 0
    duration_delta = ((avg_duration - prev_duration) / prev_duration * 100) if prev_duration > 0 else float('inf')

    col1, col2, col3 = st.columns(3)
    with col1:
        display_metric("Total Sales", f"${total_sales:,.0f}", "fa-satellite-dish", delta=sales_delta)
    with col2:
        display_metric("Unique Customers", f"{total_customers:,}", "fa-users-cog", delta=customers_delta)
    with col3:
        display_metric("Avg. Shipping", f"{avg_duration:.1f} days" if avg_duration > 0 else "N/A", "fa-rocket", delta=duration_delta)
    
    st.markdown("<br><br>", unsafe_allow_html=True)

    # --- Charts ---
    tab1, tab2 = st.tabs(["📈 **Trend & Forecast**", "📦 **Category Drill-Down**"])

    with tab1:
        st.subheader("Sales Trend and Future Forecast")
        time_agg = st.radio("Aggregate by", ["Day", "Week", "Month"], horizontal=True, index=2, key="time_agg")
        agg_map = {'Day': 'D', 'Week': 'W', 'Month': 'M'}
        sales_over_time = df.set_index('Order Date').resample(agg_map[time_agg])['Sales'].sum().reset_index()

        # NEW: Simple forecasting using moving average
        window_size = 3
        sales_over_time['Forecast'] = sales_over_time['Sales'].rolling(window=window_size, min_periods=1).mean().shift(1)
        
        fig_line = go.Figure()
        # Actual Sales Trace
        fig_line.add_trace(go.Scatter(
            x=sales_over_time['Order Date'], y=sales_over_time['Sales'],
            mode='lines+markers', name='Actual Sales', fill='tozeroy',
            line=dict(color='#ffd700', width=3), # Gold
            marker=dict(symbol='diamond', size=8)
        ))
        # Forecast Trace
        fig_line.add_trace(go.Scatter(
            x=sales_over_time['Order Date'], y=sales_over_time['Forecast'],
            mode='lines', name='Forecast (3-Period MA)',
            line=dict(color='#daa520', width=2, dash='dot') # Goldenrod
        ))
        fig_line.update_layout(
            template="plotly_dark", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(gridcolor='rgba(255,255,255,0.1)'), yaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
            hovermode="x unified", legend=dict(y=0.99, x=0.01, bgcolor='rgba(0,0,0,0.5)')
        )
        st.plotly_chart(fig_line, use_container_width=True)

    with tab2:
        st.subheader("Interactive Sales Treemap")
        # NEW: Replaced Polar chart with a more effective and interactive Treemap
        category_subcat_sales = df.groupby(['Category', 'Sub-Category'])['Sales'].sum().reset_index()
        fig_treemap = px.treemap(category_subcat_sales, path=[px.Constant("All Categories"), 'Category', 'Sub-Category'], values='Sales',
                                 color='Sales', template='plotly_dark',
                                 color_continuous_scale='Sunset', # Changed color scale to a more golden/warm one
                                 hover_data={'Sales': ':.2f'})
        fig_treemap.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=0, l=0, r=0, b=0))
        st.plotly_chart(fig_treemap, use_container_width=True)

    st.markdown("<br><br>", unsafe_allow_html=True)
    
    # --- Map and Data Preview ---
    col1, col2 = st.columns([2, 1])
    state_abbrev = {
    "Alabama": "AL","Alaska": "AK","Arizona": "AZ","Arkansas": "AR","California": "CA",
    "Colorado": "CO","Connecticut": "CT","Delaware": "DE","District of Columbia": "DC",
    "Florida": "FL","Georgia": "GA","Hawaii": "HI","Idaho": "ID","Illinois": "IL",
    "Indiana": "IN","Iowa": "IA","Kansas": "KS","Kentucky": "KY","Louisiana": "LA",
    "Maine": "ME","Maryland": "MD","Massachusetts": "MA","Michigan": "MI","Minnesota": "MN",
    "Mississippi": "MS","Missouri": "MO","Montana": "MT","Nebraska": "NE","Nevada": "NV",
    "New Hampshire": "NH","New Jersey": "NJ","New Mexico": "NM","New York": "NY",
    "North Carolina": "NC","North Dakota": "ND","Ohio": "OH","Oklahoma": "OK","Oregon": "OR",
    "Pennsylvania": "PA","Rhode Island": "RI","South Carolina": "SC","South Dakota": "SD",
    "Tennessee": "TN","Texas": "TX","Utah": "UT","Vermont": "VT","Virginia": "VA",
    "Washington": "WA","West Virginia": "WV","Wisconsin": "WI","Wyoming": "WY"
    }
    with col1:
        st.subheader("🌍 Sales Distribution by State")
        df["State_abbrev"] = df["State"].map(state_abbrev)
        sales_by_state = df.groupby('State')['Sales'].sum().reset_index()

        fig_map = go.Figure(data=go.Choropleth(
            locations=df["State_abbrev"],
            z=df["Sales"],
            locationmode='USA-states',
            colorscale='Sunset', # Changed colorscale
            colorbar_title="Sales ($)",
            hovertemplate='<b>State:</b> %{location}<br><b>Total Sales:</b> $%{z:,.0f}<extra></extra>'
        ))

        fig_map.update_layout(
            geo=dict(
                scope='usa',
                projection=dict(type='albers usa'),
                bgcolor='rgba(0,0,0,0)',
                landcolor='rgba(20,20,40,0.8)',
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=0, r=0, t=0, b=0),
            template="plotly_dark"
        )

        st.plotly_chart(fig_map, use_container_width=True)
    with col2:
        st.subheader("🔑 Key Insights")

        # أعلى ولاية
        top_state = sales_by_state.loc[sales_by_state['Sales'].idxmax()]
        st.info(
            f"**Top State:** {top_state['State']} with ${top_state['Sales']:,.0f} in sales.",
            icon="🏆"
        )

        # أعلى كاتيجوري
        top_category = df.groupby('Category')['Sales'].sum().idxmax()
        st.info(
            f"**Top Category:** {top_category} is the best performing category.",
            icon="🥇"
        )

        # Preview للبيانات
        with st.expander("📋 Show Filtered Data Preview", expanded=False):
            st.dataframe(df.head(100), use_container_width=True)


# ... (The rest of the page rendering functions remain the same)
def render_products_page(df):
    render_page_header("Product & Shipping Deep Dive", "Analyze performance by category and shipping times.")
    
    category_option = st.selectbox("Filter by Category", ["All Categories"] + sorted(df["Category"].dropna().unique().tolist()))
    
    filtered_df = df.copy()
    if category_option != "All Categories":
        filtered_df = filtered_df[filtered_df["Category"] == category_option]
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("Category Metrics")
        display_metric("Total Products Sold", f"{filtered_df['Quantity'].sum():,}", "fa-boxes-stacked")
        st.markdown("<br>", unsafe_allow_html=True)
        display_metric("Avg. Sales/Product", f"${filtered_df['Sales'].mean():,.2f}", "fa-tag")
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.subheader("Top 5 Selling Products")
        top_products = filtered_df.groupby("Product Name")['Sales'].sum().nlargest(5).reset_index()
        st.dataframe(top_products, use_container_width=True, hide_index=True)

    with col2:
        st.subheader("Sales Distribution by Sub-Category")
        fig_sunburst = px.sunburst(filtered_df, path=['Category', 'Sub-Category'], values='Sales',
                                         color='Sales', color_continuous_scale='Sunset', # Changed color scale
                                         template='plotly_dark')
        fig_sunburst.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_sunburst, use_container_width=True)
    
    st.markdown("---")
    
    if 'duration_new' in filtered_df.columns:
        st.subheader("🕒 Shipping Duration Analysis")
        fig_box = px.box(filtered_df, x="duration_new", y="Sub-Category", template="plotly_dark", color="Category",
                          labels={"Duration": "Shipping Duration (days)", "Sub-Category": ""})
        fig_box.update_traces(marker=dict(color='#ffd700', outliercolor='#daa520', size=8)) # Gold colors
        fig_box.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_box, use_container_width=True)

def render_customer_page(df):
    render_page_header("Customer Insights", "Explore customer behavior and purchase history.")

    customer_list = ["All Customers"] + sorted(df["Customer Name"].dropna().unique().tolist())
    customer_option = st.selectbox("Select a Customer", customer_list)

    if customer_option != "All Customers":
        customer_df = df[df["Customer Name"] == customer_option]
        
        col1, col2, col3 = st.columns(3)
        with col1:
            display_metric("Total Orders", f"{customer_df.shape[0]}", "fa-shopping-cart")
        with col2:
            display_metric("Total Spend", f"${customer_df['Sales'].sum():,.2f}", "fa-credit-card")
        with col3:
            display_metric("Avg. Shipping", f"{customer_df['duration_new'].mean():.1f} days" if 'duration_new' in customer_df.columns else "N/A", "fa-hourglass-half")
            
        st.subheader(f"Purchase History for {customer_option}")
        display_cols = ["Order Date", "Product Name", "Category", "Sub-Category", "Quantity", "Sales"]
        if 'Duration' in customer_df.columns: display_cols.append("Duration")
        st.dataframe(customer_df[display_cols].sort_values(by="Order Date", ascending=False), use_container_width=True, hide_index=True)
    else:
        st.subheader("Top 10 Customers by Total Sales")
        customer_summary = df.groupby("Customer Name").agg(TotalSales=("Sales", "sum"), Orders=("Order Date", "count")).nlargest(10, 'TotalSales').reset_index()
        fig_top_cust = px.bar(customer_summary, x='TotalSales', y='Customer Name', orientation='h', 
                                 template='plotly_dark', color='TotalSales', color_continuous_scale='Sunset', # Changed color scale
                                 labels={'TotalSales': 'Total Sales', 'Customer Name': ''})
        fig_top_cust.update_layout(yaxis={'categoryorder':'total ascending'}, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_top_cust, use_container_width=True)

def render_advanced_analytics_page(df):
    render_page_header("Advanced Analytics", "3D Customer Segmentation with RFM Model.")
    
    st.info("""
    *RFM (Recency, Frequency, Monetary) analysis* is a marketing model for behavior-based customer segmentation.
    - **Recency:** How recently a customer has made a purchase.
    - **Frequency:** How often they make a purchase.
    - **Monetary:** How much money they spend.
    """, icon="💡")
    
    snapshot_date = df['Order Date'].max() + timedelta(days=1)
    rfm_df = df.groupby('Customer ID').agg({
        'Order Date': lambda date: (snapshot_date - date.max()).days,
        'Customer ID': 'count', 'Sales': 'sum'
    }).rename(columns={'Order Date': 'Recency', 'Customer ID': 'Frequency', 'Sales': 'Monetary'})
    
    try:
        rfm_df['R_Score'] = pd.qcut(rfm_df['Recency'].rank(method='first'), 4, labels=range(4, 0, -1)).astype(int)
        rfm_df['F_Score'] = pd.qcut(rfm_df['Frequency'].rank(method='first'), 4, labels=range(1, 5)).astype(int)
        rfm_df['M_Score'] = pd.qcut(rfm_df['Monetary'].rank(method='first'), 4, labels=range(1, 5)).astype(int)
        
        def assign_segment(row):
            if row['R_Score'] >= 4 and row['F_Score'] >= 4: return 'Champions'
            if row['R_Score'] >= 3 and row['F_Score'] >= 3: return 'Loyal Customers'
            if row['R_Score'] >= 3 and row['M_Score'] >= 3: return 'Potential Loyalists'
            if row['R_Score'] <= 2 and row['F_Score'] >= 3: return 'At Risk'
            if row['R_Score'] <= 2: return 'Hibernating'
            return 'Others'
        rfm_df['Segment'] = rfm_df.apply(assign_segment, axis=1)

        st.subheader("👑 3D Customer Segmentation with RFM Analysis")
        fig_3d = px.scatter_3d(
            rfm_df.reset_index(), x='Recency', y='Frequency', z='Monetary',
            color='Segment', symbol='Segment', hover_data=['Customer ID'], template='plotly_dark',
            color_discrete_map={
                'Champions': '#ffd700', 'Loyal Customers': '#f0e68c', # Gold and Khaki
                'Potential Loyalists': '#daa520', 'At Risk': '#ff8800', # Goldenrod and Orange
                'Hibernating': '#ff0000', 'Others': '#888888'
            }
        )
        fig_3d.update_layout(
            margin=dict(l=0, r=0, b=0, t=0),
            scene=dict(
                xaxis_title='Recency (Days)', yaxis_title='Frequency (Orders)', zaxis_title='Monetary Value ($)'
            ),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            legend=dict(y=0.9, x=0.1, title=dict(text='Customer Segment'))
        )
        fig_3d.update_traces(marker=dict(size=5, opacity=0.8))
        st.plotly_chart(fig_3d, use_container_width=True, height=700)

    except ValueError:
        st.warning("Could not generate RFM scores due to non-unique bin edges. Try a larger dataset or date range.")

# =================================================================================
# Main App Logic
# =================================================================================
st.sidebar.title("Visionary Analytics 👑")
st.sidebar.markdown("---")

data_source = st.sidebar.radio("Select Data Source", ["Use Demo Data", "Upload Excel File"], key="data_source")

df_master = None
if data_source == "Use Demo Data":
    with st.spinner('Generating demo data...'):
        df_master = generate_demo_data()
else:
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

    if uploaded_file:
        with st.spinner('Processing uploaded file...'):
            # Read Excel file into a dataframe
            df_master = pd.read_excel(uploaded_file)

            # Show a preview of the data
            st.success("File successfully uploaded and processed!")
            
    else:
        st.info("👋 Please upload an Excel file to begin.", icon="⬆")
        st.stop()

if df_master is None or df_master.empty:
    st.error("Data could not be loaded or is empty. Please check your file or try the demo data.")
    st.stop()

st.sidebar.header("GLOBAL FILTERS")
min_date = df_master["Order Date"].min().date()
max_date = df_master["Order Date"].max().date()
date_range = st.sidebar.date_input("Select Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])

all_categories = sorted(df_master["Category"].dropna().unique().tolist())
selected_categories = st.sidebar.multiselect("Select Categories", options=all_categories, default=all_categories)

df_filtered = df_master[
    (df_master["Order Date"] >= start_date) & (df_master["Order Date"] <= end_date) &
    (df_master["Category"].isin(selected_categories))
]

period_duration = (end_date - start_date).days + 1
prev_start_date = start_date - timedelta(days=period_duration)
prev_end_date = start_date - timedelta(days=1)
df_prev_period = df_master[
    (df_master["Order Date"] >= prev_start_date) & (df_master["Order Date"] <= prev_end_date) &
    (df_master["Category"].isin(selected_categories))
]

if df_filtered.empty:
    st.warning("No data available for the selected filters.", icon="⚠")
    st.stop()

# NEW: Add Export Button to Sidebar
st.sidebar.markdown("---")
st.sidebar.download_button(
    label="📥 Export Filtered Data",
    data=convert_df_to_csv(df_filtered),
    file_name=f"filtered_data_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.csv",
    mime="text/csv",
    use_container_width=True
)

st.sidebar.markdown("---")
st.sidebar.header("Navigation")
page_options = {
    "📊 Overview": render_overview_page,
    "📦 Products": render_products_page,
    "👤 Customers": render_customer_page,
    "🔬 Advanced Analytics": render_advanced_analytics_page
}
page_selection = st.sidebar.radio("Go to", list(page_options.keys()), key="navigation")

# Render the selected page
if page_selection == "📊 Overview":
    page_options[page_selection](df_filtered, df_prev_period)
else:
    page_options[page_selection](df_filtered)