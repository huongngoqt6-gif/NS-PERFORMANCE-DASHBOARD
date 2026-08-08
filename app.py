import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os

# Config page
st.set_page_config(page_title="NS Dashboard", layout="wide")

# CSS Styling to match requirements
st.markdown('''
<style>
    /* Main Background & Header */
    .stApp, .stApp > header {
        background-color: #607D8B !important;
    }
    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #1E3A8A !important;
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    /* Cards */
    .metric-card {
        background-color: white;
        border-radius: 10px;
        padding: 15px;
        border: 1px solid #e0e0e0;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        margin-bottom: 15px;
        text-align: center; /* Căn giữa toàn bộ text trong card */
    }
    .metric-card h4 {
        color: #7f8c8d;
        margin-top: 0;
        font-size: 14px;
        text-transform: uppercase;
        text-align: center; /* Căn giữa tiêu đề H4 */
    }
    .metric-card .value-container {
        display: flex;
        justify-content: center; /* Đưa các giá trị ra giữa */
        gap: 30px; /* Khoảng cách giữa 2 cột (ví dụ MNG và PIC) */
    }
    .metric-card .value-box {
        text-align: center; /* Căn giữa nội dung từng giá trị */
    }
    .metric-card .value-label {
        color: #7f8c8d;
        font-size: 12px;
        margin-bottom: 5px;
    }
    .metric-card .value-number {
        color: #e67e22;
        font-size: 24px;
        font-weight: bold;
        margin: 0;
    }
    /* Charts container */
    .stPlotlyChart {
        background-color: white !important;
        border-radius: 10px;
        border: 1px solid #dcdde1;
        padding: 10px;
    }
    /* Dataframes */
    .stDataFrame {
        background-color: white;
        border-radius: 10px;
        padding: 10px;
    }
</style>
''', unsafe_allow_html=True)

def create_card(title, v1_label, v1, v2_label=None, v2=None):
    # Viết liền HTML để tránh lỗi khoảng trắng của Markdown
    html = f'<div class="metric-card"><h4>{title}</h4><div class="value-container"><div class="value-box"><div class="value-label">{v1_label}</div><div class="value-number">{v1}</div></div>'
    
    if v2_label:
        html += f'<div class="value-box"><div class="value-label">{v2_label}</div><div class="value-number">{v2}</div></div>'
        
    html += '</div></div>'
    
    st.markdown(html, unsafe_allow_html=True)

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    
    # HC Sheet
    hc = pd.read_excel(xls, sheet_name='HC', skiprows=3, header=None)
    hc.columns = ['Office', 'Month', 'Approved_HC_MNG', 'Approved_HC_PIC', 'Approved_HC_Total', 
                  'Actual_HC_MNG', 'Actual_HC_PIC', 'Actual_HC_Total', 
                  'Required_HC_MNG', 'Required_HC_PIC', 'Required_HC_Total', 'Capacity_Pct', 'Capacity_Status']
                  
    # Shipment Volume Sheet
    sv = pd.read_excel(xls, sheet_name='Shipment volume', skiprows=3, header=None)
    sv.columns = ['Office', 'Month', 'Active_customer', 'AI', 'AE', 'OILCL', 'OIFCL', 'OELCL', 'OEFCL', 'DI', 'DE', 'DM', 'CE', 'CI', 'HE', 'HI', 'RE', 'RI', 'RD', 'Total']

    # BU Allocation Sheet
    bu = pd.read_excel(xls, sheet_name='BU allocation', skiprows=3, header=None)
    bu.columns = ['Office', 'Month', 'Segment', 'Core_Volume', 'Core_Time', 'Ancillary_Volume', 'Ancillary_Time', 'Supporting_Volume', 'Supporting_Time', 'Exception_Volume', 'Exception_Time', 'Total_workload', 'Pct_of_Network']

    # N-S Customer List
    nsc = pd.read_excel(xls, sheet_name='N-S Customer list', skiprows=3, header=None)
    nsc.columns = ['No', 'Office', 'Customer', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Total']

    # CS FTE
    csfte = pd.read_excel(xls, sheet_name='CS FTE', skiprows=2, header=None)
    csfte.columns = ['Office', 'CSPIC', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']

    # Sheets C, A, S, E
    sc = pd.read_excel(xls, sheet_name='C', skiprows=3, header=None)
    sc.columns = ['Office', 'Scope', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Total']

    sa = pd.read_excel(xls, sheet_name='A', skiprows=3, header=None)
    sa.columns = ['Office', 'Scope', 'Service_detail', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Total']

    ss = pd.read_excel(xls, sheet_name='S', skiprows=3, header=None)
    ss.columns = ['Office', 'Scope', 'Supporting_Details', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Total']

    se = pd.read_excel(xls, sheet_name='E', skiprows=3, header=None)
    se.columns = ['Office', 'Scope', 'Cretia', 'Exception_Details', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Total']

    return {'HC': hc, 'SV': sv, 'BU': bu, 'NSC': nsc, 'CSFTE': csfte, 'C': sc, 'A': sa, 'S': ss, 'E': se}

def clean_empty_months(df):
    months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
    cols_to_keep = [c for c in df.columns if c not in months]
    for m in months:
        if m in df.columns and not df[m].isnull().all() and (df[m] != 0).any():
            cols_to_keep.append(m)
    return df[[c for c in df.columns if c in cols_to_keep]]

st.sidebar.title("NS Dashboard")

# Chỉ định đường dẫn file trực tiếp thay vì upload
file_path = "NTW DATA FOR DASHBOARD.xlsx"

if os.path.exists(file_path):
    try:
        data = load_data(file_path)
        
        # FILTERS
        st.sidebar.subheader("Filters (Sort)")
        all_offices = sorted(data['HC']['Office'].dropna().unique().tolist())
        all_months = sorted(data['HC']['Month'].dropna().unique().tolist())
        all_customers = sorted(data['NSC']['Customer'].dropna().unique().tolist())
        
        sel_office = st.sidebar.multiselect("Office", all_offices)
        sel_month = st.sidebar.multiselect("Month", all_months)
        sel_customer = st.sidebar.multiselect("Customer", all_customers)
        
        def filter_df(df, has_month=True, has_customer=False):
            res = df.copy()
            if sel_office and 'Office' in res.columns:
                res = res[res['Office'].isin(sel_office)]
            if has_month and sel_month and 'Month' in res.columns:
                res = res[res['Month'].isin(sel_month)]
            if has_customer and sel_customer and 'Customer' in res.columns:
                res = res[res['Customer'].isin(sel_customer)]
            return res

        hc_df = filter_df(data['HC'])
        sv_df = filter_df(data['SV'])
        bu_df = filter_df(data['BU'])
        nsc_df = filter_df(data['NSC'], has_month=False, has_customer=True)
        
        page = st.sidebar.radio("Navigation", ["Overview", "Shipment volume", "FTE", "BU Allocation"])
        
        if page == "Overview":
            st.header("Overview")
            c1, c2, c3, c4, c5 = st.columns(5)
            
            with c1:
                create_card("Approved HC", "MNG", round(hc_df['Approved_HC_MNG'].mean(), 1), "PIC", round(hc_df['Approved_HC_PIC'].mean(), 1))
            with c2:
                create_card("Actual HC", "MNG", round(hc_df['Actual_HC_MNG'].mean(), 1), "PIC", round(hc_df['Actual_HC_PIC'].mean(), 1))
            with c3:
                create_card("Required HC", "MNG", round(hc_df['Required_HC_MNG'].mean(), 1), "PIC", round(hc_df['Required_HC_PIC'].mean(), 1))
            with c4:
                cap_status = hc_df['Capacity_Status'].mode()[0] if not hc_df['Capacity_Status'].empty else ""
                create_card("Capacity", "%", f"{round(hc_df['Capacity_Pct'].mean() * 100, 1)}%", "Status", cap_status)
            with c5:
                create_card("Shipment Volume", "Total Avg", round(sv_df['Total'].mean(), 1))
                
            c_chart1, c_chart2 = st.columns(2)
            with c_chart1:
                st.markdown("<div style='background: white; padding: 10px; border-radius: 10px;'>", unsafe_allow_html=True)
                hc_trend = hc_df.groupby('Month')[['Required_HC_Total', 'Actual_HC_Total']].mean().reset_index()
                fig1 = go.Figure()
                fig1.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Required_HC_Total'], name='Required HC', line=dict(color='gray')))
                fig1.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Actual_HC_Total'], name='Actual HC', fill='tonexty', line=dict(color='orange')))
                fig1.update_layout(title="Actual HC vs Required HC", margin=dict(l=0, r=0, t=30, b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig1, use_container_width=True)
                st.markdown("</div>", unsafe_allow_html=True)
                
            with c_chart2:
                st.markdown("<div style='background: white; padding: 10px; border-radius: 10px;'>", unsafe_allow_html=True)
                modes = ['AI', 'AE', 'OILCL', 'OIFCL', 'OELCL', 'OEFCL', 'DI', 'DE', 'DM', 'CE', 'CI', 'HE', 'HI', 'RE', 'RI', 'RD']
                sv_trend = sv_df.groupby('Month')[modes].sum().reset_index()
                fig2 = px.area(sv_trend, x='Month', y=modes, title="Trend Volume by Transportation Mode")
                fig2.update_layout(margin=dict(l=0, r=0, t=30, b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig2, use_container_width=True)
                st.markdown("</div>", unsafe_allow_html=True)
                
            st.markdown("<div style='color: white; margin-top: 20px; font-style: italic;'>* Ghi chú: 1 FTE tương ứng với 8 tiếng x 95% hiệu suất x 22 ngày trong tháng.</div>", unsafe_allow_html=True)
            
        elif page == "Shipment volume":
            st.header("Shipment volume")
            c1, c2 = st.columns(2)
            with c1:
                create_card("Active Customer", "Avg Active", round(sv_df['Active_customer'].mean(), 1))
            with c2:
                create_card("Shipment Volume", "Avg Total", round(sv_df['Total'].mean(), 1))
                
            c_chart1, c_chart2 = st.columns(2)
            with c_chart1:
                top_customers = nsc_df.sort_values(by='Total', ascending=False).head(10)
                fig1 = px.bar(top_customers, x='Total', y='Customer', orientation='h', text='Total', title="Top 10 Customers by Volume")
                fig1.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig1, use_container_width=True)
                
            with c_chart2:
                modes = ['AI', 'AE', 'OILCL', 'OIFCL', 'OELCL', 'OEFCL', 'DI', 'DE', 'DM', 'CE', 'CI', 'HE', 'HI', 'RE', 'RI', 'RD']
                mode_sums = sv_df[modes].sum().reset_index()
                mode_sums.columns = ['Mode', 'Volume']
                mode_sums = mode_sums.sort_values(by='Volume', ascending=False)
                fig2 = px.bar(mode_sums, x='Volume', y='Mode', orientation='h', text='Volume', title="Top Transportation Mode")
                fig2.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig2, use_container_width=True)
                
            st.subheader("Data Tables")
            c_t1, c_t2 = st.columns(2)
            with c_t1:
                st.markdown("**N-S Customer List**")
                st.dataframe(clean_empty_months(nsc_df))
            with c_t2:
                st.markdown("**Shipment Volume**")
                st.dataframe(sv_df)
                
        elif page == "FTE":
            st.header("FTE")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                create_card("Approved HC", "MNG", round(hc_df['Approved_HC_MNG'].mean(), 1), "PIC", round(hc_df['Approved_HC_PIC'].mean(), 1))
            with c2:
                create_card("Actual HC", "MNG", round(hc_df['Actual_HC_MNG'].mean(), 1), "PIC", round(hc_df['Actual_HC_PIC'].mean(), 1))
            with c3:
                create_card("Required HC", "MNG", round(hc_df['Required_HC_MNG'].mean(), 1), "PIC", round(hc_df['Required_HC_PIC'].mean(), 1))
            with c4:
                cap_status = hc_df['Capacity_Status'].mode()[0] if not hc_df['Capacity_Status'].empty else ""
                create_card("Capacity", "%", f"{round(hc_df['Capacity_Pct'].mean() * 100, 1)}%", "Status", cap_status)
                
            hc_trend = hc_df.groupby('Month')[['Required_HC_Total', 'Actual_HC_Total', 'Approved_HC_Total']].mean().reset_index()
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Approved_HC_Total'], name='Approved HC', line=dict(color='blue')))
            fig.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Required_HC_Total'], name='Required HC', line=dict(color='gray')))
            fig.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Actual_HC_Total'], name='Actual HC', fill='tonexty', line=dict(color='orange')))
            fig.update_layout(title="HC Trends", plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
            
            st.subheader("CS FTE Table")
            st.dataframe(clean_empty_months(filter_df(data['CSFTE'], has_month=False)))
            
        elif page == "BU Allocation":
            st.header("BU Allocation")
            segments = ['AE', 'AI', 'OE', 'OI', 'CC', 'TR', 'WH']
            cols = st.columns(7)
            for idx, seg in enumerate(segments):
                with cols[idx]:
                    val = bu_df[bu_df['Segment'] == seg]['Pct_of_Network'].mean()
                    if pd.isna(val): val = 0
                    create_card(seg, "% of Network", f"{round(val * 100, 1)}%")
                    
            c_left, c_right = st.columns(2)
            with c_left:
                seg_data = bu_df.groupby('Segment')['Pct_of_Network'].mean().reset_index()
                fig = px.pie(seg_data, values='Pct_of_Network', names='Segment', hole=0.5, title="% of Network by Segment")
                fig.update_traces(pull=[0.05]*len(seg_data))
                st.plotly_chart(fig, use_container_width=True)
                
            with c_right:
                st.markdown("**BU Allocation Data**")
                st.dataframe(bu_df)
                
            st.subheader("Details by Services")
            t_a, t_c, t_s, t_e = st.tabs(["Core Service (A)", "Ancillary Service (C)", "Supporting Activity (S)", "Exception Handling (E)"])
            with t_a:
                df_a = filter_df(data['A'], has_month=False)
                if 'Total' in df_a.columns: df_a = df_a.sort_values(by='Total', ascending=False)
                st.dataframe(clean_empty_months(df_a))
            with t_c:
                df_c = filter_df(data['C'], has_month=False)
                if 'Total' in df_c.columns: df_c = df_c.sort_values(by='Total', ascending=False)
                st.dataframe(clean_empty_months(df_c))
            with t_s:
                df_s = filter_df(data['S'], has_month=False)
                if 'Total' in df_s.columns: df_s = df_s.sort_values(by='Total', ascending=False)
                st.dataframe(clean_empty_months(df_s))
            with t_e:
                df_e = filter_df(data['E'], has_month=False)
                if 'Total' in df_e.columns: df_e = df_e.sort_values(by='Total', ascending=False)
                st.dataframe(clean_empty_months(df_e))
    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}. Vui lòng kiểm tra lại cấu trúc file Excel.")
else:
    st.error(f"Không tìm thấy file '{file_path}'. Vui lòng đảm bảo file Excel đã được đưa lên cùng thư mục trên GitHub.")
