import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os
import streamlit as st

# Config page
st.set_page_config(page_title="CSD Performance", layout="wide")

# CSS Styling để tùy biến st.metric thành các card nền trắng, bo viền, căn giữa, chữ xám, số cam
st.markdown('''
<style>
    .stApp {
        background-color: #607D8B !important;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
    }
    header[data-testid="stHeader"] {
        background-color: rgba(0,0,0,0) !important;
    }
    [data-testid="stSidebar"] {
        background-color: #1E3A8A !important;
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    .main-title {
        color: white;
        font-size: 32px;
        font-weight: bold;
        margin-top: -10px;
        margin-bottom: 10px;
        text-transform: uppercase;
    }
    
    /* Thiết kế thẻ Card chuẩn, nền trắng, bo viền, căn giữa tuyệt đối */
    .custom-metric-card {
        background-color: white;
        border-radius: 10px;
        border: 1px solid #dcdde1;
        padding: 15px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        text-align: center;
        margin-bottom: 10px;
    }
    .card-title {
        color: #7f8c8d;
        font-size: 12px;
        font-weight: bold;
        text-transform: uppercase;
        margin-bottom: 5px;
    }
    .card-value {
        color: #e67e22;
        font-size: 26px;
        font-weight: bold;
        margin-bottom: 5px;
    }
    .card-sub {
        color: #7f8c8d;
        font-size: 13px;
        display: flex;
        justify-content: space-around;
        padding: 0 10px;
    }
    .label-part {
        color: #7f8c8d;
    }
    .num-part {
        color: #e67e22;
        font-weight: bold;
    }
    .stPlotlyChart {
        background-color: white !important;
        border-radius: 10px;
        border: 1px solid #dcdde1;
        padding: 10px;
    }
    .stDataFrame {
        background-color: white;
        border-radius: 10px;
        padding: 10px;
    }
</style>
''', unsafe_allow_html=True)

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    
    hc = pd.read_excel(xls, sheet_name='HC', skiprows=3, header=None)
    hc.columns = ['Office', 'Month', 'Approved_HC_MNG', 'Approved_HC_PIC', 'Approved_HC_Total', 
                 'Actual_HC_MNG', 'Actual_HC_PIC', 'Actual_HC_Total', 
                 'Required_HC_MNG', 'Required_HC_PIC', 'Required_HC_Total', 'Capacity_Pct', 'Capacity_Status']
                 
    sv = pd.read_excel(xls, sheet_name='Shipment volume', skiprows=3, header=None)
    sv.columns = ['Office', 'Month', 'Active_customer', 'AI', 'AE', 'OILCL', 'OIFCL', 'OELCL', 'OEFCL', 'DI', 'DE', 'DM', 'CE', 'CI', 'HE', 'HI', 'RE', 'RI', 'RD', 'Total']

    bu = pd.read_excel(xls, sheet_name='BU allocation', skiprows=3, header=None)
    bu.columns = ['Office', 'Month', 'Segment', 'Core_Volume', 'Core_Time', 'Ancillary_Volume', 'Ancillary_Time', 'Supporting_Volume', 'Supporting_Time', 'Exception_Volume', 'Exception_Time', 'Total_workload', 'Pct_of_Network']
    
    # Lấy trực tiếp giá trị từ cột M ('% of Net Work') của nguồn thay vì tự tính
    raw_bu_col_m = pd.read_excel(xls, sheet_name='BU allocation', skiprows=3, header=None)
    if len(raw_bu_col_m.columns) > 12:
        bu['Pct_of_Network'] = raw_bu_col_m[12]

    nsc = pd.read_excel(xls, sheet_name='N-S Customer list', skiprows=3, header=None)
    nsc.columns = ['No', 'Office', 'Customer', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Total']

    csfte = pd.read_excel(xls, sheet_name='CS FTE', skiprows=2, header=None)
    csfte.columns = ['Office', 'CSPIC', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']

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

st.sidebar.title("CSD Performance")
file_path = "NTW DATA FOR DASHBOARD.xlsx"

if os.path.exists(file_path):
    try:
        data = load_data(file_path)
        
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
            st.markdown('<div class="main-title">CSD Operations performance dashboard</div>', unsafe_allow_html=True)
            st.header("Overview")
            
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Approved HC (Total)</div>
                    <div class="card-value">{round(hc_df['Approved_HC_Total'].mean(), 1)}</div>
                    <div class="card-sub">
                        <span><span class="label-part">MNG:</span> <span class="num-part">{round(hc_df['Approved_HC_MNG'].mean(), 1)}</span></span>
                        <span><span class="label-part">PIC:</span> <span class="num-part">{round(hc_df['Approved_HC_PIC'].mean(), 1)}</span></span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Actual HC (Total)</div>
                    <div class="card-value">{round(hc_df['Actual_HC_Total'].mean(), 1)}</div>
                    <div class="card-sub">
                        <span><span class="label-part">MNG:</span> <span class="num-part">{round(hc_df['Actual_HC_MNG'].mean(), 1)}</span></span>
                        <span><span class="label-part">PIC:</span> <span class="num-part">{round(hc_df['Actual_HC_PIC'].mean(), 1)}</span></span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c3:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Required HC (Total)</div>
                    <div class="card-value">{round(hc_df['Required_HC_Total'].mean(), 1)}</div>
                    <div class="card-sub">
                        <span><span class="label-part">MNG:</span> <span class="num-part">{round(hc_df['Required_HC_MNG'].mean(), 1)}</span></span>
                        <span><span class="label-part">PIC:</span> <span class="num-part">{round(hc_df['Required_HC_PIC'].mean(), 1)}</span></span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c4:
                cap_status = hc_df['Capacity_Status'].mode()[0] if not hc_df['Capacity_Status'].empty else ""
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Capacity %</div>
                    <div class="card-value">{round(hc_df['Capacity_Pct'].mean() * 100, 1)}%</div>
                    <div class="card-sub"><span class="label-part">Status: {cap_status}</span></div>
                </div>
                """, unsafe_allow_html=True)
            with c5:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Shipment Volume</div>
                    <div class="card-value">{round(sv_df['Total'].mean(), 1)}</div>
                    <div class="card-sub"><span class="label-part">Total Avg</span></div>
                </div>
                """, unsafe_allow_html=True)
                
            c_chart1, c_chart2 = st.columns(2)
            with c_chart1:
                hc_trend = hc_df.groupby('Month')[['Required_HC_Total', 'Actual_HC_Total']].mean().reset_index()
                fig1 = go.Figure()
                fig1.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Required_HC_Total'], name='Required HC', line=dict(color='gray')))
                fig1.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Actual_HC_Total'], name='Actual HC', fill='tonexty', line=dict(color='orange')))
                fig1.update_layout(title="Actual HC vs Required HC", autosize=True, height=350, margin=dict(l=20, r=80, t=40, b=20), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig1, use_container_width=True)
                
            with c_chart2:
                sv_trend = sv_df.groupby(['Month', 'Office'])['Total'].sum().reset_index()
                fig2 = px.bar(sv_trend, x='Month', y='Total', color='Office', barmode='group', title="Total Volume Trend by Month & Office")
                fig2.update_layout(autosize=True, height=350, margin=dict(l=20, r=100, t=40, b=20), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig2, use_container_width=True)
                
            st.markdown("<div style='color: white; margin-top: 20px; font-style: italic;'>* Ghi chú: 1 FTE tương ứng với 8 tiếng x 95% hiệu suất x 22 ngày trong tháng.</div>", unsafe_allow_html=True)
            
        elif page == "Shipment volume":
            st.markdown('<div class="main-title">CSD Operations performance dashboard</div>', unsafe_allow_html=True)
            st.header("Shipment volume")
            
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Active Customer</div>
                    <div class="card-value" style="margin-bottom: 0px;">{round(sv_df['Active_customer'].mean(), 1)}</div>
                </div>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Shipment Volume</div>
                    <div class="card-value" style="margin-bottom: 0px;">{round(sv_df['Total'].mean(), 1)}</div>
                </div>
                """, unsafe_allow_html=True)
                
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
                fig2 = px.bar(mode_sums, x='Volume', y='Mode', orientation='h', text='Volume', title="Top Volume by Transportation Mode")
                fig2.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig2, use_container_width=True)
                
            st.subheader("Data Tables")
            
            c_t1, c_space, c_t2 = st.columns([5, 0.4, 5])
            
            with c_t1:
                st.markdown("**CSD Customer List**")
                st.dataframe(clean_empty_months(nsc_df), use_container_width=True, hide_index=True)
                
            with c_t2:
                st.markdown("**Shipment Volume**")
                st.dataframe(sv_df, use_container_width=True, hide_index=True)
                
        elif page == "FTE":
            st.markdown('<div class="main-title">N-S Operations performance dashboard</div>', unsafe_allow_html=True)
            st.header("FTE")
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Approved HC</div>
                    <div class="card-value">{round(hc_df['Approved_HC_Total'].mean(), 1)}</div>
                    <div class="card-sub">
                        <span><span class="label-part">MNG:</span> <span class="num-part">{round(hc_df['Approved_HC_MNG'].mean(), 1)}</span></span>
                        <span><span class="label-part">PIC:</span> <span class="num-part">{round(hc_df['Approved_HC_PIC'].mean(), 1)}</span></span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Actual HC</div>
                    <div class="card-value">{round(hc_df['Actual_HC_Total'].mean(), 1)}</div>
                    <div class="card-sub">
                        <span><span class="label-part">MNG:</span> <span class="num-part">{round(hc_df['Actual_HC_MNG'].mean(), 1)}</span></span>
                        <span><span class="label-part">PIC:</span> <span class="num-part">{round(hc_df['Actual_HC_PIC'].mean(), 1)}</span></span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c3:
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Required HC</div>
                    <div class="card-value">{round(hc_df['Required_HC_Total'].mean(), 1)}</div>
                    <div class="card-sub">
                        <span><span class="label-part">MNG:</span> <span class="num-part">{round(hc_df['Required_HC_MNG'].mean(), 1)}</span></span>
                        <span><span class="label-part">PIC:</span> <span class="num-part">{round(hc_df['Required_HC_PIC'].mean(), 1)}</span></span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c4:
                cap_status = hc_df['Capacity_Status'].mode()[0] if not hc_df['Capacity_Status'].empty else ""
                st.markdown(f"""
                <div class="custom-metric-card">
                    <div class="card-title">Capacity %</div>
                    <div class="card-value">{round(hc_df['Capacity_Pct'].mean() * 100, 1)}%</div>
                    <div class="card-sub"><span class="label-part">Status: {cap_status}</span></div>
                </div>
                """, unsafe_allow_html=True)
                
            hc_trend = hc_df.groupby('Month')[['Required_HC_Total', 'Actual_HC_Total', 'Approved_HC_Total']].mean().reset_index()
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Approved_HC_Total'], name='Approved HC', line=dict(color='blue')))
            fig.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Required_HC_Total'], name='Required HC', line=dict(color='gray')))
            fig.add_trace(go.Scatter(x=hc_trend['Month'], y=hc_trend['Actual_HC_Total'], name='Actual HC', fill='tonexty', line=dict(color='orange')))
            fig.update_layout(title="HC Trends", plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
            
            # --- PHẦN BẢNG HC DATA MỚI ĐÍNH KÈM ---
            HC Performance Data
            # --------------------------------------
            
            st.subheader("CS FTE Table")
            df_csfte = clean_empty_months(filter_df(data['CSFTE'], has_month=False))
            
            def get_fte_status(val):
                if pd.isna(val) or val == "":
                    return None
                if val > 1.0:
                    return "Overload"
                elif val > 0.95:
                    return "High load"
                elif val >= 0.90:
                    return "Balanced"
                else:
                    return "Less load"
            
            all_months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
            month_cols = [c for c in df_csfte.columns if c in all_months]
            
            df_csfte['Average FTE'] = df_csfte[month_cols].mean(axis=1)
            df_csfte['Average Status'] = df_csfte['Average FTE'].apply(get_fte_status)
            
            ordered_cols = ['Office', 'CSPIC']
            status_cols = []
            
            for m in month_cols:
                status_col_name = f"{m} Status"
                df_csfte[status_col_name] = df_csfte[m].apply(get_fte_status)
                ordered_cols.extend([m, status_col_name])
                status_cols.append(status_col_name)
                
            ordered_cols.extend(['Average FTE', 'Average Status'])
            status_cols.append('Average Status')
            
            df_csfte = df_csfte[ordered_cols]
            df_csfte.index = range(1, len(df_csfte) + 1)
            
            def color_status(val):
                if val == "Overload":
                    return 'background-color: #ffcccc; color: #cc0000; font-weight: bold;'
                elif val == "High load":
                    return 'background-color: #ffe6cc; color: #cc6600; font-weight: bold;'
                elif val == "Balanced":
                    return 'background-color: #cce6ff; color: #0066cc; font-weight: bold;'
                elif val == "Less load":
                    return 'background-color: #d9ffcc; color: #2e8b57; font-weight: bold;'
                return ''

            styled_df = df_csfte.style.map(color_status, subset=status_cols)

            st.dataframe(
                styled_df, 
                use_container_width=True,
                column_config={
                    "Average FTE": st.column_config.NumberColumn(
                        "Average FTE",
                        format="%.2f"
                    )
                }
            )
            
        elif page == "BU Allocation":
            st.markdown('<div class="main-title">CSD Operations performance dashboard</div>', unsafe_allow_html=True)
            st.header("BU Allocation")
            
            segments = ['AE', 'AI', 'OE', 'OI', 'CC', 'TR', 'WH']
            cols = st.columns(7)
            for idx, seg in enumerate(segments):
                with cols[idx]:
                    val = bu_df[bu_df['Segment'] == seg]['Pct_of_Network'].mean()
                    if pd.isna(val): val = 0
                    st.markdown(f"""
                    <div class="custom-metric-card">
                        <div class="card-title">{seg}</div>
                        <div class="card-value" style="margin-bottom: 0px;">{round(val * 100, 1)}%</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
            c_left, c_right = st.columns(2)
            with c_left:
                seg_data = bu_df.groupby('Segment')['Pct_of_Network'].mean().reset_index()
                
                fig = px.pie(
                    seg_data, 
                    values='Pct_of_Network', 
                    names='Segment', 
                    hole=0.5, 
                    title="% of Network by Segment"
                )
                
                fig.update_traces(
                    pull=[0.05] * len(seg_data),
                    textinfo='label+percent',
                    textposition='inside'
                )
                
                fig.update_layout(
                    showlegend=False,
                    height=520
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            with c_right:
                st.markdown("**BU Allocation Data**")
                df_bu_display = bu_df.copy()
                df_bu_display.index = range(1, len(df_bu_display) + 1)
                
                st.dataframe(
                    df_bu_display, 
                    use_container_width=True,
                    column_config={
                        "Pct_of_Network": st.column_config.NumberColumn(
                            "Pct_of_Network",
                            format="%.1f%%"
                        )
                    }
                )
                
            st.subheader("Details by Services")
            t_a, t_c, t_s, t_e = st.tabs(["Core Service (A)", "Ancillary Service (C)", "Supporting Activity (S)", "Exception Handling (E)"])
            
            with t_a:
                df_a = filter_df(data['A'], has_month=False)
                if 'Total' in df_a.columns: 
                    df_a = df_a.sort_values(by='Total', ascending=False)
                df_a = clean_empty_months(df_a)
                df_a.index = range(1, len(df_a) + 1)
                st.dataframe(df_a, use_container_width=True)
                
            with t_c:
                df_c = filter_df(data['C'], has_month=False)
                if 'Total' in df_c.columns: 
                    df_c = df_c.sort_values(by='Total', ascending=False)
                df_c = clean_empty_months(df_c)
                df_c.index = range(1, len(df_c) + 1)
                st.dataframe(df_c, use_container_width=True)
                
            with t_s:
                df_s = filter_df(data['S'], has_month=False)
                if 'Total' in df_s.columns: 
                    df_s = df_s.sort_values(by='Total', ascending=False)
                df_s = clean_empty_months(df_s)
                df_s.index = range(1, len(df_s) + 1)
                st.dataframe(df_s, use_container_width=True)
                
            with t_e:
                df_e = filter_df(data['E'], has_month=False)
                if len(df_e) > 0:
                    df_e = df_e.iloc[1:]
                df_e = clean_empty_months(df_e)
                df_e.index = range(1, len(df_e) + 1)
                st.dataframe(df_e, use_container_width=True)
    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}. Vui lòng kiểm tra lại cấu trúc file Excel.")
else:
    st.error(f"Không tìm thấy file '{file_path}'. Vui lòng đảm bảo file Excel đã được đưa lên cùng thư mục trên GitHub.")
