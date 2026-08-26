from pathlib import Path
from plotly.subplots import make_subplots
import pandas as pd  # pip install openpyxl
import plotly.express as px  # pip install plotly-express
import numpy as np
import plotly.graph_objects as go
from datetime import datetime
from datetime import date
from datetime import datetime, timedelta
import streamlit as st  # pip install streamlit
from streamlit_autorefresh import st_autorefresh  # pip install streamlit_autorefresh
import streamlit_authenticator as stauth
from streamlit_option_menu import option_menu  # pip install streamlit_option_menu
import pickle
from streamlit_card import card
import base64
from PIL import Image
from plotly.graph_objects import Figure
import os
import streamlit.components.v1 as components

# ==========================================
# 1. SƏHİFƏ VƏ DATA KONFİQURASİYASI
# ==========================================
st.set_page_config(
    page_title="Branch Performance",
    page_icon="🏦",
    layout="wide"
)

today = datetime.today().strftime('%d-%b-%Y %H:%M:%S')
st.sidebar.info(f'🗓️ updated {today}')

# Faylların yerləşdiyi qovluq yolu
DATA_FOLDER = r'C:\Users\009688\Desktop\MXSPerformance'

@st.cache_data
def load_single_excel(file_name):
    """
    Qovluqdakı tək bir Excel faylını oxuyur.
    Fayl tapılmadıqda proqramın çökməməsi üçün boş DataFrame qaytarır.
    """
    full_path = os.path.join(DATA_FOLDER, f"{file_name}.xlsx")
    if os.path.exists(full_path):
        return pd.read_excel(full_path)
    else:
        st.warning(f"Xəbərdarlıq: '{file_name}.xlsx' faylı tapılmadı!")
        return pd.DataFrame()

try:
    
    # Hər bir faylı ayrı-ayrı oxuyuruq (Fayl adları tam uyğun olmalıdır)
    Branch_info = load_single_excel('Branch_info')
    Branch_manager = load_single_excel('Branch_manager')
    Supervisor = load_single_excel('Supervisor')
    Loan = load_single_excel('Loan')
    Card = load_single_excel('Card')
    Insurance = load_single_excel('Insurance')
    Deposit = load_single_excel('Deposit')
    ABB_biz = load_single_excel('ABB_biz')
    Employees = load_single_excel('Employees')
    Actual_employee = load_single_excel('Actual_employee')
    Workload = load_single_excel('Workload')
    ST_KPI = load_single_excel('ST_KPI')
    Paperless = load_single_excel('Paperless')
    ABB_biz_satis = load_single_excel('ABB_biz_satis')
    Noqsan = load_single_excel('Noqsan')
    Kredit_noqsan = load_single_excel('Kredit_noqsan')
    ST_kenarlasma = load_single_excel('ST_kenarlasma')
    PL = load_single_excel('PL')
    Detailed_PL = load_single_excel('Detailed_PL')
    İmtahan = load_single_excel('İmtahan')
    Qmeter = load_single_excel('Qmeter')
    Sikayet_indeksi = load_single_excel('Sikayet_indeksi')
    Ulduz = load_single_excel('Ulduz')
    Ferqliler = load_single_excel('Ferqliler')
    Direct_Sales = load_single_excel('Direct_Sales')
    Tamkart_portfel = load_single_excel('Tamkart_portfel')
    Istehlak_portfel = load_single_excel('Istehlak_portfel')
    Depozit_portfel = load_single_excel('Depozit_portfel')
    Korporat_portfel = load_single_excel('Korporat_portfel')
    Ipoteka_portfeli = load_single_excel('Ipoteka_portfeli')
    Baza = load_single_excel('Baza')
    Potential_customer = load_single_excel('Potential_customer')
    Portfolio_distribution = load_single_excel('Portfolio_distribution')
    Portfolio_percentage = load_single_excel('Portfolio_percentage')
    Kart_sayi = load_single_excel('Kart_sayi')
    Bazar_payi = load_single_excel('Bazar_payi')


    # Kodun digər hissələrində 'data_dict' istifadə olunursa, lüğəti yenidən formalaşdırırıq:
    data_dict = {
        'Branch_info': Branch_info,
        'Branch_manager': Branch_manager,
        'Supervisor': Supervisor,
        'Loan': Loan,
        'Card': Card,
        'Insurance': Insurance,
        'Deposit': Deposit,
        'Abb_biz': ABB_biz,
        'Employees': Employees,
        'Actual_employee': Actual_employee,
        'Workload': Workload,
        'ST_KPI': ST_KPI,
        'Paperless': Paperless,
        'ABB_biz_satis': ABB_biz_satis,
        'Noqsan': Noqsan,
        'Kredit_noqsan': Kredit_noqsan,
        'ST_kenarlasma': ST_kenarlasma,
        'PL': PL,
        'Detailed_PL': Detailed_PL,
        'İmtahan': İmtahan,
        'Qmeter': Qmeter,
        'Sikayet_indeksi' : Sikayet_indeksi,
        'Ulduz' : Ulduz,
        'Ferqliler' : Ferqliler,
        'Direct_Sales' : Direct_Sales,
        'Tamkart_portfel' : Tamkart_portfel,
        'Istehlak_portfel' : Istehlak_portfel,
        'Depozit_portfel' : Depozit_portfel,
        'Korporat_portfel' : Korporat_portfel,
        'Ipoteka_portfeli' : Ipoteka_portfeli,
        'Baza' : Baza,
        'Potential_customer' : Potential_customer,
        'Portfolio_distribution' : Portfolio_distribution,
        'Portfolio_percentage'  : Portfolio_percentage,
        'Kart_sayi' : Kart_sayi,
        'Bazar_payi' : Bazar_payi
    }

except Exception as e:
    st.error(f"Fayllar oxunarkən xəta baş verdi: {e}")

@st.cache_data
def all_list():
    branch = Branch_info['MXS_name']
    data_types = ['Amount', 'Count']
    return branch, data_types

branch, data_types = all_list()

def get_date_range():
    curr = datetime.today()
    start_date = datetime(curr.year, 1, 1)
    return start_date, curr

# ==========================================
# 2. VİZUAL ÜSLUB VƏ CARD FUNKSİYALARI
# ==========================================
st.markdown("""
<style>
    .custom-card { width: 150px; border-radius: 8px; box-shadow: 0 2px 4px 0 rgba(0,0,0,0.05); overflow: hidden; background: white; }
    .card-image { width: 100%; background-size: cover; background-position: center; }
    .card-content { padding: 0.1px; background: white; }
    .card-title { font-weight: bold; color: #333; margin: 0; text-align: center; }
</style>
""", unsafe_allow_html=True)

def get_image_base64(image_path):
    with open(image_path, "rb") as img_file:
        encoded = base64.b64encode(img_file.read()).decode()
    return f"data:image/jpeg;base64,{encoded}"

def create_custom_card(image_path, title, width=150, height=210, font_size=16):
    image_data = get_image_base64(image_path)
    st.markdown(f"""
    <div class="custom-card" style="width: {width}px;">
        <div class="card-image" style="background-image: url('{image_data}'); height: {height}px;"></div>
        <div class="card-content">
            <p class="card-title" style="font-size: {font_size}px;">{title}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# 3. FİLTER PANELİ
# ==========================================
col_1, spacer, col_2, col_3 = st.columns([1.5, 0.1, 1, 1.5])

with col_1:
    branch_name = st.selectbox('Branch selection:', options=branch)

with col_2:
    selected_data_type = st.radio('Data type selection:', options=data_types)

start_date_def, end_date_def = get_date_range()

with col_3:
    col_from, col_to = st.columns(2)
    with col_from:
        tarixden = st.date_input("Start Date:", value=start_date_def.date(), format="YYYY/MM/DD")
    with col_to:
        tarixe = st.date_input("End Date:", value=end_date_def.date(), format="YYYY/MM/DD")

# ==========================================
# 4. BRANCH REVIEW (KADR VƏ ŞƏKİLLƏR)
# ==========================================
def branch_review():
    col1, col2, col3, col4, col5, space, col6, col7 = st.columns([0.6, 0.35, 0.6, 0.35, 0.5, 0.2, 0.5, 1.8])
  
    with col1:
        manager_name = Branch_manager[Branch_manager['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)['Branch_manager'].iloc[0]
        st.text('Branch Manager')
        path = f"Photos/{manager_name}.jpg" if os.path.exists(f"Photos/{manager_name}.jpg") else "Photos/Default.jpg"
        create_custom_card(path, manager_name, 150, 210, 18)
    
    with col2:
        st.text('Team Lead')
        tl1 = Branch_manager[Branch_manager['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)['Team_lead1'].iloc[0]
        tl2 = Branch_manager[Branch_manager['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)['Team_lead2'].iloc[0]
        for tl in [tl1, tl2]:
            path = f"Photos/{tl}.jpg" if os.path.exists(f"Photos/{tl}.jpg") else "Photos/Default.jpg"
            create_custom_card(path, tl, 73, 102.2, 10)

    with col3:
        supervisor_name = Supervisor[Supervisor['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)['Supervisor'].iloc[0]
        st.text('Supervisor')
        path = f"Photos/{supervisor_name}.jpg" if os.path.exists(f"Photos/{supervisor_name}.jpg") else "Photos/Default.jpg"
        create_custom_card(path, supervisor_name, 150, 210, 18)
    
    with col4:
        st.text('Member')
        m1 = Supervisor[Supervisor['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)['Member1'].iloc[0]
        m2 = Supervisor[Supervisor['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)['Member2'].iloc[0]
        for mb in [m1, m2]:
            path = f"Photos/{mb}.jpg" if os.path.exists(f"Photos/{mb}.jpg") else "Photos/Default.jpg"
            create_custom_card(path, mb, 73, 102.2, 10)

    with col5:
        coordinator = Supervisor[Supervisor['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)['Koordinator'].iloc[0]
        st.text('Coordinator')
        path = f"Photos/{coordinator}.jpg" if os.path.exists(f"Photos/{coordinator}.jpg") else "Photos/Default.jpg"
        create_custom_card(path, coordinator, 150, 210, 18)




    with col6:
        # 1. Metriklərin ölçüsünü kiçiltmək üçün xüsusi CSS
        st.markdown("""
            <style>
                div[data-testid="stMetric"] {
                    padding: 2px 0px !important;
                }
                div[data-testid="stMetricLabel"] {
                    font-size: 13px !important;
                }
                div[data-testid="stMetricValue"] {
                    font-size: 18px !important;
                }
            </style>
        """, unsafe_allow_html=True)

        # 2. Tarixləri təhlükəsiz formata gətiririk
        dt_from = pd.to_datetime(tarixden)
        dt_to = pd.to_datetime(tarixe)
        search_branch = str(branch_name).strip().lower()

        # --- DATALARIN HESABLANMASI ---

        # 1. Actual & Current Staff
        act_dates = pd.to_datetime(Actual_employee['Date'], errors='coerce')
        act_filtered = Actual_employee[
            (Actual_employee['MXS_name'] == branch_name) & 
            (act_dates >= dt_from) & 
            (act_dates <= dt_to)
        ].sort_values(by='Date', ascending=False)

        curr_val = act_filtered.iloc[0]['Current_Staff'].round(1) if not act_filtered.empty else "-"
        act_val = act_filtered.iloc[0]['Actual_Staff'].round(1) if not act_filtered.empty else "-"

        # 2. Qmeter Score
        qmeter_names = Qmeter['MXS_name'].astype(str).str.strip().str.lower()
        qmeter_dates = pd.to_datetime(Qmeter['Date'], format='%d-%m-%y', errors='coerce')

        qmeter_filtered = Qmeter[
            (qmeter_names == search_branch) & 
            (qmeter_dates >= dt_from) & 
            (qmeter_dates <= dt_to)
        ]
        qmeter_val = round(qmeter_filtered['Score'].mean(), 1) if not qmeter_filtered.empty else "-"

        # 3. Şikayət İndeksi
        sikayet_names = Sikayet_indeksi['MXS_name'].astype(str).str.strip().str.lower()
        sikayet_dates = pd.to_datetime(Sikayet_indeksi['Date'], errors='coerce')

        sikayet_filtered = Sikayet_indeksi[
            (sikayet_names == search_branch) & 
            (sikayet_dates >= dt_from) & 
            (sikayet_dates <= dt_to)
        ]

        if not sikayet_filtered.empty:
            idx_series = sikayet_filtered['İndex'].astype(str).str.replace('%', '').str.strip()
            idx_numeric = pd.to_numeric(idx_series, errors='coerce')
            avg_idx = idx_numeric.mean()
            if avg_idx < 1:
                avg_idx = avg_idx * 100
            sikayet_val = f"{avg_idx:.2f}%"
        else:
            sikayet_val = "-"

        # --- ALT-ALTA YIĞCAM ÇIXARIŞ ---
        st.metric("Current Staff", curr_val)
        st.metric("Actual Staff", act_val)
        st.metric("Qmeter Score", qmeter_val)
        st.metric("Şikayət İndeksi", sikayet_val)
    
    with col7:
        current_stf = Employees[Employees['MXS_name'] == branch_name].sort_values(by='Date', ascending=False)
        st.dataframe(current_stf[['Staff','Position']], width=800, height=280, hide_index=True)

branch_review()



# ==========================================
# 5. ULDUZ NOMİNATİON BÖLMƏSİ
# ==========================================
st.subheader('"Ulduz" Nomination', divider="grey")

if 'Ulduz' in globals() and not Ulduz.empty and 'branch_name' in globals() and branch_name:
    # Ulduz cədvəlində sütun adlarını təmizləyirik
    u_df = Ulduz.copy()
    u_df.columns = u_df.columns.astype(str).str.strip()
    
    # Sütunları tapırıq
    mxs_col = next((c for c in u_df.columns if str(c).strip().lower() in ['mxs_name', 'mxs', 'filial', 'branch']), None)
    mfx_col = next((c for c in u_df.columns if any(k in str(c).strip().lower() for k in ['mukafat', 'mükafat', 'award', 'star'])), None)
    dovr_col = next((c for c in u_df.columns if 'dovr' in str(c).strip().lower() or 'dövr' in str(c).strip().lower()), None)
    
    star_list = []
    
    if mxs_col and mfx_col:
        # Seçilən filiala görə filtrləyirik
        branch_rows = u_df[u_df[mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()]
        
        for _, row in branch_rows.iterrows():
            val = row[mfx_col]
            dovr_text = str(row[dovr_col]).strip() if dovr_col and not pd.isna(row[dovr_col]) else ""
            
            try:
                num = int(float(str(val).strip()))
                if num in [1, 2, 3]:
                    # Ulduz və tam altında Dövr məlumatı
                    star_html = f"""<div style="display: flex; flex-direction: column; align-items: center; justify-content: center; margin-right: 18px; margin-bottom: 10px;"><div style="position:
                      relative; display: inline-flex; align-items: center; justify-content: center; width: 65px; height: 65px;"><svg viewBox="0 0 500 500" width="65" height="65" style="position: 
                      absolute; top: 0; left: 0;"><path fill="#FFC107" d="M250,30 L250,380 L120,450 L150,300 L30,200 L180,180 Z" /><path fill="#E6A100" d="M250,30 L320,180 L470,200 L350,300 L380,450 L250,380 Z"
                        /></svg><span style="position: relative; z-index: 2; font-size: 26px; font-weight: 900; color: #FFFFFF; font-family: 'Arial Rounded MT Bold', 'Nunito', 'Segoe UI', sans-serif; text-shadow:
                          0px 1px 2px rgba(0,0,0,0.2);">{num}</span></div><span style="font-size: 13px; font-weight: 700; color: #4A5568; margin-top: 4px; font-family: 'Segoe UI',
                            sans-serif;">{dovr_text}</span></div>"""
                    star_list.append(star_html)
            except:
                continue

    # Əgər ulduz varsa ekrana çıxarırıq, yoxdursa mədəni bildiriş göstəririk
    if star_list:
        all_stars_html = "".join(star_list)
        st.markdown(f"<div style='display: flex; flex-wrap: wrap; align-items: flex-start; margin-top: 10px;'>{all_stars_html}</div>", unsafe_allow_html=True)
    else:
        st.info("Seçilmiş filial Ulduz nominasiyası üzrə mükafata layiq görülməyib.")

# --- NƏZƏRƏ ÇARPAN GÜCLÜ MƏSAFƏ (BOŞLUQ) ---
st.write("")
st.write("")
st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)


# ==========================================
# FƏRQLƏNƏNLƏR (NƏTİCƏLƏR) BÖLMƏSİ
# ==========================================

st.subheader('"Hədəf Liderləri" Nomination', divider="grey")

if 'Ferqliler' in globals() and not Ferqliler.empty and 'branch_name' in globals() and branch_name:
    # Cədvəl sütun adlarını təmizləyirik
    f_df = Ferqliler.copy()
    f_df.columns = f_df.columns.astype(str).str.strip()
    
    # Sütun adlarını dinamik tapırıq
    mxs_col = next((c for c in f_df.columns if str(c).strip().lower() in ['mxs_name', 'mxs', 'filial', 'branch']), None)
    dovr_col = next((c for c in f_df.columns if 'dovr' in str(c).strip().lower() or 'dövr' in str(c).strip().lower()), None)
    kat_col = next((c for c in f_df.columns if 'kateqor' in str(c).strip().lower() or 'nominasiya' in str(c).strip().lower()), None)
    yer_col = next((c for c in f_df.columns if 'yer' in str(c).strip().lower()), None)

    if mxs_col:
        # Seçilən filiala görə filtrləyirik
        filtered_f = f_df[f_df[mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()]

        if not filtered_f.empty:
            cards_html = ""
            for _, row in filtered_f.iterrows():
                dovr_val = str(row[dovr_col]).strip() if dovr_col and not pd.isna(row[dovr_col]) else "-"
                kat_val = str(row[kat_col]).strip() if kat_col and not pd.isna(row[kat_col]) else "-"
                
                # Bütün dərəcələr üçün 3-cü yerlə eyni vahid rəng stili
                badge_bg = "#F8FAFC"
                badge_color = "#1E293B"    # Tünd, aydın oxunan rəng
                badge_border = "#CBD5E1"
                
                try:
                    yer_val = int(float(str(row[yer_col]).strip()))
                    if yer_val == 1:
                        place_icon = "🥇 1-ci Yer"
                    elif yer_val == 2:
                        place_icon = "🥈 2-ci Yer"
                    elif yer_val == 3:
                        place_icon = "🥉 3-cü Yer"
                    else:
                        place_icon = f"🎖️ {yer_val}-ci Yer"
                except:
                    place_icon = "🎗️ -"

                # ABB Stilində Xüsusi Kart
                card_item = f"""
                <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 5px solid #0052C2; border-radius: 10px; padding: 14px 18px; margin-bottom: 12px; box-shadow: 0 2px 6px rgba(0, 82, 194, 0.05); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                    <div style="flex: 2; min-width: 200px;">
                        <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px; display: block; margin-bottom: 2px;">Nominasiya</span>
                        <span style="font-size: 15px; font-weight: 700; color: #0F172A; font-family: 'Segoe UI', sans-serif;">{kat_val}</span>
                    </div>
                    <div style="flex: 1; min-width: 100px; text-align: center;">
                        <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px; display: block; margin-bottom: 2px;">Dövr</span>
                        <span style="font-size: 14px; font-weight: 600; color: #0052C2; background: #E6F0FA; padding: 4px 12px; border-radius: 6px;">{dovr_val}</span>
                    </div>
                    <div style="flex: 1; min-width: 120px; text-align: right;">
                        <span style="background-color: {badge_bg}; color: {badge_color}; border: 1px solid {badge_border}; font-size: 15px; font-weight: 800; padding: 8px 16px; border-radius: 20px; display: inline-block;">
                            {place_icon}
                        </span>
                    </div>
                </div>
                """
                cards_html += card_item

            st.markdown(cards_html, unsafe_allow_html=True)
        else:
            st.info("Seçilmiş filial üzrə fərqlənmə nəticəsi qeydə alınmayıb.")
else:
    st.info("Məlumat tapılmadı və ya filial seçilməyib.")


# --- NƏZƏRƏ ÇARPAN GÜCLÜ MƏSAFƏ (BOŞLUQ) ---
st.write("")
st.write("")
st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)


# ==========================================
# 6. BRANCH P&L PERFORMANCE BÖLMƏSİ
# ==========================================
st.subheader("Branch P&L", divider="grey")

def display_pl_table(data_dict, branch_name):
    pl_df = data_dict['PL'].copy()
    pl_df.columns = pl_df.columns.astype(str).str.strip()

    # Filiala görə filtrləmə
    mxs_col = next((c for c in ['MXS', 'MXS_name', 'Filial', 'Branch'] if c in pl_df.columns), None)
    if mxs_col:
        mask = pl_df[mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()
        filtered_df = pl_df.loc[mask].copy()
    else:
        filtered_df = pl_df.copy()

    if filtered_df.empty:
        st.info(f"'{branch_name}' filialı üzrə PL məlumatı tapılmadı.")
        return

    # Göstəriləcək sütunların seçilməsi
    available_cols = [c for c in filtered_df.columns if c not in ['MXS', 'MXS_CODE', 'MXS_name', 'Filial', 'Branch']]
    total_rows = len(filtered_df)

    # HTML Sütun Kartlarının Hazırlanması
    header_styles = ["header-dark", "header-mid", "header-light", "header-soft"]
    data_columns_html = ""

    for c_idx, col_name in enumerate(available_cols):
        style_class = header_styles[c_idx % len(header_styles)]
        
        col_card = f"""
        <div class="column-card">
            <div class="card-header {style_class}">
                <div class="col-title">{col_name}</div>
            </div>
        """
        
        for _, row in filtered_df.iterrows():
            val = row[col_name]
            is_text_col = col_name.strip().lower() in ['il', 'i̇l', 'dövr', 'dovr', 'year', 'period']

            if pd.isna(val) or val is None:
                display_val = "-"
            elif is_text_col:
                display_val = str(val).replace(".0", "").replace(" ", "")
            elif isinstance(val, (int, float)):
                display_val = f"{val:,.0f} ₼".replace(",", " ")
            else:
                display_val = str(val)

            col_card += f'<div class="card-row data-val">{display_val}</div>'
            
        col_card += "</div>"
        data_columns_html += col_card

    num_cols = len(available_cols)
    grid_template_columns = f"repeat({num_cols}, 1fr)"

    # Layout HTML & CSS
    custom_component = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            background-color: transparent;
            margin: 0;
            padding: 2px;
        }}

        .pnl-container {{
            width: 100%;
            box-sizing: border-box;
        }}

        .pnl-grid {{
            display: grid;
            grid-template-columns: {grid_template_columns};
            gap: 4px;
            width: 100%;
        }}

        .column-card {{
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 1px 4px rgba(0,0,0,0.05);
            background-color: #ffffff;
            display: flex;
            flex-direction: column;
            border: 1px solid #e2e8f0;
            min-width: 0;
        }}

        .card-header {{
            height: 48px;
            padding: 4px 2px;
            box-sizing: border-box;
            display: flex;
            align-items: center;
            justify-content: center;
            text-align: center;
            color: #ffffff;
        }}

        .header-dark {{ background: linear-gradient(135deg, #102a45 0%, #1a365d 100%); }}
        .header-mid {{ background: linear-gradient(135deg, #1e3a5f 0%, #2b4c7e 100%); }}
        .header-light {{ background: linear-gradient(135deg, #2b5c8f 0%, #3a75ab 100%); }}
        .header-soft {{ background: linear-gradient(135deg, #417bb8 0%, #5b9bd5 100%); }}

        .col-title {{
            font-size: 10px;
            font-weight: 700;
            line-height: 1.1;
            word-wrap: break-word;
            hyphens: auto;
        }}

        .card-row {{
            height: 34px;
            padding: 2px 2px;
            box-sizing: border-box;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 10px;
            border-bottom: 1px solid #edf2f7;
        }}

        .data-val {{
            color: #2d3748;
            font-weight: 700;
            text-align: center;
            white-space: nowrap;
        }}

        .data-val:nth-child(even) {{
            background-color: #f8fafc;
        }}
    </style>
    </head>
    <body>

    <div class="pnl-container">
        <div class="pnl-grid">
            {data_columns_html}
        </div>
    </div>

    </body>
    </html>
    """

    calc_height = (total_rows * 34) + 75
    components.html(custom_component, height=calc_height, scrolling=False)

# --- Çağırış ---
display_pl_table(data_dict, branch_name)

# ==========================================
# 1. DETAILED P&L REPORT FUNKSİYASI
# ==========================================
def build_detailed_pl_report(Detailed_PL, branch_name, start_date, end_date):
    if Detailed_PL is None or Detailed_PL.empty:
        st.warning("'Detailed_PL' faylında məlumat tapılmadı.")
        return

    df = Detailed_PL.copy()
    df.columns = df.columns.astype(str).str.strip()

    # Filiala (MXS_name) görə süzmə
    if 'MXS_name' in df.columns:
        df = df[df['MXS_name'].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()]

    # Tarixə (DATE) görə süzmə
    if 'DATE' in df.columns and not df.empty:
        df['DATE_parsed'] = pd.to_datetime(df['DATE'], dayfirst=True, errors='coerce').dt.date
        s_date = pd.to_datetime(start_date).date() if start_date else df['DATE_parsed'].min()
        e_date = pd.to_datetime(end_date).date() if end_date else df['DATE_parsed'].max()
        df = df[(df['DATE_parsed'] >= s_date) & (df['DATE_parsed'] <= e_date)]

    # AMOUNT sütununu rəqəmə çevirmək
    if 'AMOUNT' in df.columns:
        df['AMOUNT'] = pd.to_numeric(df['AMOUNT'], errors='coerce').fillna(0)
    else:
        df['AMOUNT'] = 0.0

    # Şablon strukturu
    template_structure = [
        ("Əmlak ilə bağlı xərclər", "İcarə haqqı"),
        ("Əmlak ilə bağlı xərclər", "Mühafizə xərcləri"),
        ("Əmlak ilə bağlı xərclər", "Vergilər"),
        ("Əmlak ilə bağlı xərclər", "Sığorta xərcləri"),
        
        ("Əsas vəsaitlər ilə bağlı xərclər", "Mebel və avadanlıqların təmir və saxlanması"),
        ("Əsas vəsaitlər ilə bağlı xərclər", "Bina və qurğuların təmir və saxlanması"),
        ("Əsas vəsaitlər ilə bağlı xərclər", "Köhnəlmə xərcləri"),
        ("Əsas vəsaitlər ilə bağlı xərclər", "Kompüter və rabitə avadanlıqlarının təmir və saxlanması"),
        
        ("İşçilər ilə bağlı xərclər", "Mükafatlar"),
        ("İşçilər ilə bağlı xərclər", "Ezamiyyə ödənişləri"),
        ("İşçilər ilə bağlı xərclər", "İdman-sağlamlıq xərcləri"),
        ("İşçilər ilə bağlı xərclər", "İşçilərin sığortası"),
        ("İşçilər ilə bağlı xərclər", "Əmək haqqı (Ştat Üzrə)"),
        ("İşçilər ilə bağlı xərclər", "İşçilərin tədris xərcləri"),
        ("İşçilər ilə bağlı xərclər", "Fondlara ayırmalar"),
        ("İşçilər ilə bağlı xərclər", "Maddi yardım və sair ödənişlər"),
        ("İşçilər ilə bağlı xərclər", "Digər xərclər"),
        
        ("Kommunal xidmətlər üzrə xərclər", "Tullantıların daşınması və dezinfeksiya"),
        ("Kommunal xidmətlər üzrə xərclər", "Elektrik enerji xərcləri"),
        
        ("Məsləhət, audit və digər peşəkar xidmətlər", "Hüquqi məsrəflər"),
        ("Məsləhət, audit və digər peşəkar xidmətlər", "Peşəkar xidmətlər"),
        
        ("Mətbəə, mal və materialların alınması", "Mətbəə, mal və materialların alınması"),
        
        ("Nəqliyyat xərcləri", "Nəqliyyat vasitələrinin təmir və saxlanması"),
        
        ("Rabitə və proqram xərcləri", "Poçt xərcləri"),
        ("Rabitə və proqram xərcləri", "Telefon danışıq və abunə haqqı"),
        ("Rabitə və proqram xərcləri", "Optik kabel xərcləri"),
        
        ("Üzvlük, nümayəndəlik və digər xərclər", "Digər xərclər"),
        ("Üzvlük, nümayəndəlik və digər xərclər", "Müxtəlif tədbirlər"),
        ("Üzvlük, nümayəndəlik və digər xərclər", "Ödənilmiş cərimə və dəbbə məbləği")
    ]

    template_df = pd.DataFrame(template_structure, columns=['TIP', 'BOLGU'])

    # Datanı cəmləmək
    if not df.empty and 'TIP' in df.columns and 'BOLGU' in df.columns:
        actual_summary = df.groupby(['TIP', 'BOLGU'], as_index=False)['AMOUNT'].sum()
        merged = pd.merge(template_df, actual_summary, on=['TIP', 'BOLGU'], how='left')
    else:
        merged = template_df.copy()
        merged['AMOUNT'] = 0.0

    merged['AMOUNT'] = merged['AMOUNT'].fillna(0.0)

    # HTML Cədvəl - Ultra Yığcam Və Qrupları Ayıran Dizayn
    html_table = """
    <style>
        .custom-pl-table-container {
            width: 100%;
            overflow-x: auto;
            margin-top: 5px;
            border-radius: 4px;
            border: 1px solid #cfd8dc;
        }
        .custom-pl-table {
            width: 100%;
            border-collapse: collapse;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 12px;
            color: #2c3e50;
            background-color: #ffffff;
            line-height: 1.2;
        }
        .custom-pl-table th, .custom-pl-table td {
            border: 1px solid #e0e0e0;
            padding: 4px 8px; /* Hündürlüyü azaldıldı */
        }
        .custom-pl-table tr:hover {
            background-color: #f5f9ff;
        }
        .tip-cell {
            background-color: #e3f2fd;
            color: #0d47a1;
            font-weight: 700;
            text-align: center;
            vertical-align: middle;
            font-size: 12.5px;
            width: 18%; /* Eni kiçildildi */
            border-right: 2px solid #90caf9 !important;
            border-bottom: 3px solid #90a4ae !important; /* Sol xananın alt tərəfini ayırır */
        }
        .total-row {
            font-weight: 700;
            background-color: #f1f5f9 !important;
            color: #000000;
            border-bottom: 3px solid #90a4ae !important; /* Qrupları bir-birindən ayıran qalın xətt */
        }
        .amount-cell {
            text-align: right;
            white-space: nowrap;
            font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
        }
    </style>
    <div class="custom-pl-table-container">
    <table class="custom-pl-table">
    """

    for tip_val, group in merged.groupby('TIP', sort=False):
        group_rows = group.reset_index(drop=True)
        row_count = len(group_rows) + 1  # Sub-sətirlər + TOTAL sətri
        
        tip_total = group['AMOUNT'].sum()

        for idx, row in group_rows.iterrows():
            html_table += "<tr>"
            if idx == 0:
                html_table += f'<td rowspan="{row_count}" class="tip-cell">{tip_val}</td>'
            
            amt_str = f"{row['AMOUNT']:,.2f} ₼".replace(",", " ")
            html_table += f'<td>{row["BOLGU"]}</td>'
            html_table += f'<td class="amount-cell">{amt_str}</td>'
            html_table += "</tr>"
        
        tot_str = f"{tip_total:,.2f} ₼".replace(",", " ")
        html_table += f'<tr class="total-row"><td style="text-align:center;">TOTAL</td><td class="amount-cell">{tot_str}</td></tr>'

    html_table += "</table></div>"

    # Tarix başlığı
    date_str = ""
    if start_date and end_date:
        s_str = start_date.strftime('%Y/%m/%d') if hasattr(start_date, 'strftime') else str(start_date)
        e_str = end_date.strftime('%Y/%m/%d') if hasattr(end_date, 'strftime') else str(end_date)
        date_str = f"({s_str} - {e_str})"

    with st.expander(f"📈 Admin Expenses {date_str}", expanded=False):
        st.markdown(html_table, unsafe_allow_html=True)

# Mövcud filtrlərinin altından bu tək sətr yazılır:
build_detailed_pl_report(data_dict['Detailed_PL'], branch_name, tarixden, tarixe)        


# --- NƏZƏRƏ ÇARPAN GÜCLÜ MƏSAFƏ (BOŞLUQ) ---
st.write("")
st.write("")
st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)



# ==========================================
# 5. SALES PERFORMANCE BÖLMƏSİ
# ==========================================
st.subheader("Sales Performance", divider="grey")
def create_simple_bar_chart(x_data, y_data, title, color):
    fig = go.Figure()
    
    # Y oxunun maksimumunu tapıb tavanı 20% yuxarı qaldırırıq (yazı kəsilməsin)
    max_y = max(y_data) if len(y_data) > 0 and max(y_data) > 0 else 1
    
    fig.add_trace(go.Bar(
        x=x_data,
        y=y_data,
        name=title,
        marker_color=color,
        text=y_data,
        texttemplate='<b>%{y}</b>',   # <-- Barların üstündəki yazıları BOLD edir
        textposition='outside',
        textfont=dict(size=11, color='#2C3E50') # Şrifti daha oxunaqlı edir
    ))
    
    fig.update_layout(
        title={'text': title, 'x': 0, 'xanchor': 'left'},
        margin=dict(l=10, r=10, t=30, b=20),
        height=280,
        showlegend=False,
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            showgrid=False,
            tickangle=0,
            nticks=5,
            tickfont=dict(size=10, color='#7F8C8D')
        ),
        yaxis=dict(
            showgrid=True, 
            gridcolor='#F0F0F0',
            range=[0, max_y * 1.20]  # <-- Y oxunun maksimumunu 20% artırır ki, yazı kəsilməsin
        )
    )
    return fig

def create_kpi_chart(dates, values, kpi_percents, legend_name, graph_name, sale_color):
    fig = go.Figure()
    
    # Bar - Satışlar
    amount_text = [f"<b>{val/1000000:,.2f}M</b>" if val >= 100000 else f"<b>{val}</b>" for val in values]
    fig.add_trace(go.Bar(
        x=dates, y=values, name=legend_name, text=amount_text, textposition='outside',
        textfont=dict(family="Bahnschrift Condensed", size=12, color="grey"),
        marker=dict(color=sale_color, line=dict(width=1, color="#E6E6E6"))
    ))

    # Line - KPI (Orijinal saxlanıldı)
    fig.add_trace(go.Scatter(
        x=dates, y=kpi_percents, name='KPI %', yaxis='y2', mode='lines+markers+text',
        text=[f"<b>{val:.0f}%</b>" for val in kpi_percents], textposition='top center',
        textfont=dict(family="Bahnschrift Condensed", size=12, color="#D2691E"),
        line=dict(color='rgba(222, 160, 115, 0.8)', width=1.5),
        marker=dict(color='#D2691E', symbol="diamond", size=6)
    ))

    # Layout (Orijinal saxlanıldı)
    fig.update_layout(
        title=dict(text=f"<b>{graph_name}</b>", font=dict(size=12), x=0.01),
        height=320, plot_bgcolor='white', paper_bgcolor='white',
        margin=dict(l=5, r=5, t=50, b=70), hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.45, xanchor="left", x=0, font=dict(size=10)),
        modebar_remove=['zoomin','zoomout','zoom','pan','select','lasso','autoscale']
    )

    # 📌 Yalnız aşağıdakı ayların yazılış tərzi dəyişdirildi:
    fig.update_xaxes(
        tickangle=0,            # Ayları tam düz (horizontal) edir
        nticks=6,               # Şəkildəki kimi ehtiyac olduqda seyrəkləşdirir
        tickfont=dict(size=10)  # Sığması üçün şrifti səliqəli ölçüyə gətirir
    )

    # Oxlar (Orijinal saxlanıldı)
    y_max = (values.max() if not values.empty else 0) * 1.5
    fig.update_yaxes(range=[0, y_max], showgrid=True, gridcolor='#E9E4E4')
    
    kpi_max = max(kpi_percents.max() if not kpi_percents.empty else 100, 100) * 1.6
    fig.update_layout(yaxis2=dict(range=[0, kpi_max], overlaying='y', side='right', showgrid=False, showticklabels=False))
    
    return fig

def branch_sales():
    
    # 📌 Birinci sıra (4 Sütun)
    c1, c2, c3, c4 = st.columns(4)
    
    row1_charts = [
        (Loan, 'Loan_count' if selected_data_type == 'Count' else 'Loan_amount', 'Kpi', 'Loan Sales', 'rgba(135, 187, 224, 0.8)', c1),
        (Card, 'Card_count', 'Kpi', 'Card Sales', 'rgba(159, 207, 172, 0.8)', c2),
        (Deposit, 'Deposit_count' if selected_data_type == 'Count' else 'Deposit_amount', 'Kpi', 'Deposit Sales', 'rgba(213, 207, 227, 0.8)', c3),
        (Insurance, 'Insurance_count', 'Kpi', 'Insurance Sales', 'rgba(235, 223, 185, 0.8)', c4)
    ]

    # 📌 İkinci sıra (4 Sütun: c5, c6, c7 və c8 burda təyin olunur)
    c5, c6, c7, c8 = st.columns(4)

    row2_charts = [
        (ABB_biz, 'abb_biz_count', 'Kpi', 'ABB Biz Register', 'rgba(235, 206, 204, 0.8)', c5),
        (ABB_biz_satis, 'Card', 'Card_Kpi' if 'Card_Kpi' in ABB_biz_satis.columns else 'Kpi', 'ABB Biz Card Sales', 'rgba(180, 220, 210, 0.8)', c6),
        (ABB_biz_satis, 'Loan_count' if selected_data_type == 'Count' else 'Loan_amount', 'Loan_Kpi' if 'Loan_Kpi' in ABB_biz_satis.columns else 'Kpi', 'ABB Biz Loan Sales', 'rgba(220, 190, 230, 0.8)', c7)
    ]

    # 1. Əvvəlki 7 chart-ı loop ilə çəkirik:
    for df, col, kpi_col, label, color, pos in row1_charts + row2_charts:
        if df is not None and col in df.columns:
            f_df = df[(df['MXS_name'] == branch_name) & (df['Date'] >= pd.to_datetime(tarixden)) & (df['Date'] <= pd.to_datetime(tarixe))].sort_values('Date')
            f_df['Date_str'] = f_df['Date'].dt.strftime('%b %Y')
            
            actual_kpi_col = kpi_col if kpi_col in f_df.columns else 'Kpi'
            
            with pos:
                fig = create_kpi_chart(f_df['Date_str'], f_df[col], f_df[actual_kpi_col], label, f"{label} / KPI", color)
                st.plotly_chart(fig, use_container_width=True, config={'displaylogo': False})

    # ==========================================
    # 8-ci SÜTUN: Direct Sales (c8 daxilində)
    # ==========================================
    with c8:
        target_col = 'DS_amount' if selected_data_type == 'Amount' else 'DS_count'
        
        f_ds = Direct_Sales[(Direct_Sales['MXS_name'] == branch_name) & 
                            (Direct_Sales['Date'] >= pd.to_datetime(tarixden)) & 
                            (Direct_Sales['Date'] <= pd.to_datetime(tarixe))].sort_values('Date')
        
        if not f_ds.empty:
            f_ds['Date_str'] = f_ds['Date'].dt.strftime('%b %Y')
            
            fig_ds = create_simple_bar_chart(
                f_ds['Date_str'], 
                f_ds[target_col], 
                'Direct Sales', 
                'rgba(255, 218, 185, 0.8)'
            )
            st.plotly_chart(fig_ds, use_container_width=True, config={'displaylogo': False})

# Funksiyanı çağırırıq:
branch_sales()






# ==========================================
# 5. PORTFEL BÖLMƏSİ
# ==========================================
st.subheader("Portfolio", divider="grey")

# 📌 Kəsilmənin qarşısını alan tam xətasız Line Chart funksiyası
def create_grid_line_chart(x_data, y_data, title, line_color, fill_color):
    fig = go.Figure()
    
    max_y = max(y_data) if len(y_data) > 0 and max(y_data) > 0 else 1
    
    formatted_text = []
    for val in y_data:
        if val >= 1_000_000:
            formatted_text.append(f"{val/1_000_000:.1f}M")
        elif val >= 1_000:
            formatted_text.append(f"{val/1_000:.0f}K")
        else:
            formatted_text.append(f"{val:.0f}")
    
    # Line Trace
    fig.add_trace(go.Scatter(
        x=x_data,
        y=y_data,
        mode='lines+markers+text',
        name=title,
        line=dict(color=line_color, width=2.5),
        marker=dict(size=6, color=line_color),
        fill='tozeroy',
        fillcolor=fill_color,
        text=formatted_text,
        textposition='top center',
        cliponaxis=False,                     # Rəqəmlərin kəsilməsini söndürür
        textfont=dict(size=9, color='#2C3E50', family="Arial")
    ))
    
    fig.update_layout(
        title={'text': f"<b>{title}</b>", 'x': 0, 'xanchor': 'left', 'font': dict(size=13, color="#2C3E50", family="Arial")},
        margin=dict(l=35, r=25, t=35, b=20),
        height=280,
        showlegend=False,
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            showgrid=False,
            tickangle=0,
            nticks=5,
            tickfont=dict(size=9, color='#7F8C8D')
        ),
        yaxis=dict(
            showgrid=True, 
            gridcolor='#F0F0F0',
            range=[0, max_y * 1.35] 
        )
    )
    return fig


# 📌 2. Yenilənmiş 3 Yuxarıda + 3 Aşağıda (Yoxlamasız) Düzülüş Funksiyası
def render_3_2_portfolio_charts():
    target_col = 'Amount' if selected_data_type == 'Amount' else 'Count'
    
    # Tarix filtrlərini təhlükəsiz datetime formatına salırıq
    dt_tarixden = pd.to_datetime(tarixden, dayfirst=True, errors='coerce')
    dt_tarixe = pd.to_datetime(tarixe, dayfirst=True, errors='coerce')
    
    # 📌 YUXARI SIRA: 3 Sütun (İstehlak, Tamkart, Depozit)
    r1_col1, r1_col2, r1_col3 = st.columns(3)
    
    top_row_portfolios = [
        (Istehlak_portfel, 'İstehlak Portfeli', '#8E44AD', 'rgba(142, 68, 173, 0.08)', r1_col1),
        (Tamkart_portfel, 'Tamkart Portfeli', '#E67E22', 'rgba(230, 126, 34, 0.08)', r1_col2),
        (Depozit_portfel, 'Depozit Portfeli', '#2980B9', 'rgba(41, 128, 185, 0.08)', r1_col3)
    ]
    
    for df, title, line_color, fill_color, pos in top_row_portfolios:
        f_df = df.copy()
        f_df['Parsed_Date'] = pd.to_datetime(f_df['Date'], dayfirst=True, errors='coerce')
        
        f_df = f_df[
            (f_df['MXS_name'].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()) & 
            (f_df['Parsed_Date'] >= dt_tarixden) & 
            (f_df['Parsed_Date'] <= dt_tarixe)
        ].sort_values('Parsed_Date')
        
        f_df['Date_str'] = f_df['Parsed_Date'].dt.strftime('%b %Y')
        fig = create_grid_line_chart(f_df['Date_str'], f_df[target_col], title, line_color, fill_color)
        with pos:
            st.plotly_chart(fig, use_container_width=True, config={'displaylogo': False})

    st.markdown("---") # Bölücü xətt

    # 📌 AŞAĞI SIRA: 3 Sütun (Korporativ, İpoteka və Pie Chart)
    r2_col1, r2_col2, r2_col3 = st.columns(3)
    
    bottom_row_portfolios = [
        (Korporat_portfel, 'Korporativ Portfel', '#27AE60', 'rgba(39, 174, 96, 0.08)', r2_col1),
        (Ipoteka_portfeli, 'İpoteka Portfeli', '#C0392B', 'rgba(192, 57, 43, 0.08)', r2_col2)
    ]
    
    # 1. İlk 2 Line Chart-ın çəkilməsi
    for df, title, line_color, fill_color, pos in bottom_row_portfolios:
        f_df = df.copy()
        f_df['Parsed_Date'] = pd.to_datetime(f_df['Date'], dayfirst=True, errors='coerce')
        
        f_df = f_df[
            (f_df['MXS_name'].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()) & 
            (f_df['Parsed_Date'] >= dt_tarixden) & 
            (f_df['Parsed_Date'] <= dt_tarixe)
        ].sort_values('Parsed_Date')
        
        f_df['Date_str'] = f_df['Parsed_Date'].dt.strftime('%b %Y')
        fig = create_grid_line_chart(f_df['Date_str'], f_df[target_col], title, line_color, fill_color)
        with pos:
            st.plotly_chart(fig, use_container_width=True, config={'displaylogo': False})

# 📌 r2_col3 daxilində PIE CHART
    with r2_col3:
        raw_pie = data_dict['Portfolio_percentage']
        f_pie = raw_pie[raw_pie["MXS_name"].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()].copy()
        
        # Faiz təmizlənməsi
        def parse_pct(val):
            if isinstance(val, str):
                return float(val.replace('%', '').strip())
            return val * 100 if val <= 1 else val
        
        f_pie["Clean_Pct"] = f_pie["Percentage"].apply(parse_pct)
        
        soft_pastel_colors = ["#AB91D4", '#4FA3D1', "#5BAFA8", "#87D8B6", '#8B5CF6']
        
        fig_pie = go.Figure(data=[go.Pie(
            labels=f_pie["Segment"],
            values=f_pie["Clean_Pct"],
            hole=0.5,
            # 📌 Pie chart-ın diametrini kiçildib yan qrafiklərlə eyni boya gətirir:
            domain=dict(x=[0.15, 0.85], y=[0.05, 0.85]),
            textinfo='label+percent',
            textposition='outside',
            textfont=dict(size=9, color='#000000', family='Arial', weight='bold'),
            marker=dict(
                colors=soft_pastel_colors,
                line=dict(color='#FFFFFF', width=1.5)
            ),
            hovertemplate="<b>%{label}</b>: %{value:.1f}%<extra></extra>"
        )])
        
        fig_pie.update_layout(
            title=dict(
                text="<b>Portfolio Segment Distribution</b>", 
                font=dict(size=13, color="#2C3E50", family="Arial"), 
                x=0, 
                xanchor='left'
            ),
            template='plotly_white',
            height=280,
            # 📌 Margin hissəsini Line Chart margin-ləri ilə eyniləşdirdik:
            margin=dict(l=35, r=25, t=35, b=20),
            showlegend=False,
            paper_bgcolor='white',
            plot_bgcolor='white'
        )
        
        st.plotly_chart(fig_pie, use_container_width=True, config={'displaylogo': False})



# Funksiyanı çağırırıq:
render_3_2_portfolio_charts()


st.write("")
st.write("")
st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)



# ==========================================
# 6. BRANCH STATİSTİK GÖSTƏRİCİLƏR
# ==========================================
st.subheader("Branch Statistics", divider="grey")

COLORS = {
    'Kredit nisbet': "#7eaee2",
    'Kredit nisbet (w/o Pension)': "#f85639",
    'Depozit nisbet': '#ffc107',
    'Kart nisbet': "#d283f7"
}

def create_workload_kpi_chart(dates, bar_data, line_data, graph_name, bar_name, line_name, bar_color='#4A90E2', line_color='#1B4F72', is_percent=False):
    fig = go.Figure()

    # Y oxlarının limitlərini rəqəmlərin kəsilməməsi üçün artırırıq
    y1_limit = (max(bar_data) if len(bar_data) > 0 and max(bar_data) > 0 else 100) * 1.25
    y2_limit = (max(line_data) if len(line_data) > 0 and max(line_data) > 0 else 1) * 1.35

    # 📌 Bar-ların dəyərini barın daxilində yuvarlaq göstəririk ki, faizlə toqquşmasın
    fig.add_trace(go.Bar(
        x=dates, y=bar_data, name=bar_name, 
        text=bar_data, texttemplate='<b>%{text:.0f}</b>',
        textposition='inside', # Rəqəmlər barın içində yerləşəcək
        insidetextanchor='middle',
        marker=dict(color=bar_color, opacity=0.85)
    ))

    # 📌 Xətt və faiz göstəricisi (Tünd Modern Accent)
    fig.add_trace(go.Scatter(
        x=dates, y=line_data, name=line_name, yaxis='y2', mode='lines+markers+text',
        text=[f"<b>{val*100:.0f}%</b>" if is_percent else f"<b>{val:.0f}</b>" for val in line_data],
        textposition='top center', cliponaxis=False,
        textfont=dict(family="Arial", size=12, color=line_color),
        line=dict(color=line_color, width=2),
        marker=dict(color=line_color, symbol="diamond", size=7)
    ))

    fig.update_layout(
        title_text=f"<b>{graph_name}</b>", height=360, plot_bgcolor='white', 
        margin=dict(l=10, r=10, t=50, b=70),
        legend=dict(orientation="h", yanchor="bottom", y=-0.45, xanchor="left", x=0),
        yaxis=dict(range=[0, y1_limit], showgrid=True, gridcolor='#EAEAEA'),
        yaxis2=dict(overlaying='y', side='right', showgrid=False, range=[0, y2_limit], showticklabels=False)
    )
    fig.update_xaxes(tickangle=0, nticks=6, tickfont=dict(size=10, color="#555"))
    return fig

def create_workload_kpi_chart2(dates, c1, c2, c3, graph_name):
    fig = go.Figure()
    colors, names = ['#5DADE2', '#2E86C1', '#1B4F72'], ['Credit Load', 'Card Load', 'Other Load']
    totals = np.where((np.array(c1)+np.array(c2)+np.array(c3)) == 0, 1, (np.array(c1)+np.array(c2)+np.array(c3)))
    for d, n, c in zip([c1, c2, c3], names, colors):
        percs = [(v / t) * 100 for v, t in zip(d, totals)]
        fig.add_trace(go.Bar(x=dates, y=d, name=n, marker_color=c, text=[f"<b>{p:.0f}%</b>" if p > 0 else "" for p in percs], textposition='auto'))
    fig.update_layout(title_text=graph_name, height=350, barmode='stack', plot_bgcolor='white', margin=dict(l=10, r=10, t=60, b=80),
                      legend=dict(orientation="h", yanchor="bottom", y=-0.5, xanchor="left", x=0))
    fig.update_xaxes(tickangle=0, nticks=6, tickfont=dict(size=10))
    return fig

def create_ST_chart(dates, v1, v2, v3, graph_name, names=['Served', 'Not Served', 'Other']):
    fig = go.Figure()
    colors = ["#2ecc71", "#D8452B", '#F4D03F']

    for i, d in enumerate([v1, v2, v3]):
        if i == 2 and sum(d) == 0: continue
        fig.add_trace(go.Bar(
            x=dates, y=d, name=names[i], marker_color=colors[i],
            text=d, texttemplate='<b>%{text:.0f}</b>', textposition='auto',
            hovertemplate=f"<b>{names[i]}</b>: %{{y:.0f}}<extra></extra>"
        ))

    totals = [sum(x) for x in zip(v1, v2, v3)]
    max_val = max(totals) if totals else 0
    y_max = max_val * 1.15 if max_val > 0 else 10

    fig.update_layout(
        title_text=graph_name, height=350, barmode='stack', plot_bgcolor='white',
        margin=dict(l=10, r=10, t=60, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=-0.5, xanchor="left", x=0),
        hovermode="x unified"
    )
    fig.update_yaxes(range=[0, y_max])
    fig.update_xaxes(tickangle=0, nticks=6, tickfont=dict(size=10))
    return fig

def create_paperless_chart(df, title):
    fig = go.Figure()

    for col, color in COLORS.items():
        if col in df.columns:
            labels = [f"<b>{v*100:.0f}%</b>" if pd.notnull(v) and v > 0 else "" for v in df[col]]
            fig.add_trace(go.Scatter(
                x=df['Date_str'], y=df[col], name=col,
                mode='lines+markers+text', text=labels, textposition='top center', cliponaxis=False,
                textfont=dict(family="Arial", size=10, color="#333"),
                line=dict(width=1.2, color=color), marker=dict(size=5, symbol="circle", color=color),
                hovertemplate=f"<b>{col}</b>: %{{y:.1%}}<extra></extra>"
            ))

    max_val = df[[c for c in COLORS if c in df.columns]].max().max() if not df.empty else 0.1
    y_max = min((max_val or 0.1) + 0.15, 1.15)

    fig.update_layout(
        title=dict(text=f"<b>{title}</b>", x=0.01, font=dict(size=13)),
        height=350, plot_bgcolor='white', paper_bgcolor='white',
        margin=dict(l=10, r=10, t=60, b=80),
        legend=dict(orientation="h", yanchor="bottom", y=-0.5, xanchor="left", x=0, font=dict(size=9)),
        hovermode="x unified",
        modebar_remove=['zoomin', 'zoomout', 'zoom', 'pan', 'select', 'lasso', 'autoscale']
    )
    fig.update_yaxes(range=[0, y_max], showgrid=True, gridcolor='#F0F0F0', tickformat=".0%")
    fig.update_xaxes(showgrid=False, type='category', tickangle=0, nticks=6, tickfont=dict(size=10))
    return fig

def get_paperless_data():
    df = Paperless.copy()
    df.columns = df.columns.str.strip()

    date_col = next(c for c in ['Date', 'DATE', 'Tarix', 'tarix'] if c in df.columns)
    mxs_col = next(c for c in ['MXS_name', 'Filial', 'Branch'] if c in df.columns)

    df['Date_clean'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    mask = (
        (df[mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()) &
        (df['Date_clean'] >= pd.to_datetime(tarixden)) &
        (df['Date_clean'] <= pd.to_datetime(tarixe))
    )
    filtered_df = df.loc[mask].copy()

    percent_cols = list(COLORS.keys())
    for col in percent_cols:
        if col in filtered_df.columns:
            s = filtered_df[col].astype(str).str.replace('%', '', regex=False).str.replace(',', '.', regex=False).str.strip()
            vals = pd.to_numeric(s if filtered_df[col].dtype == object else filtered_df[col], errors='coerce')
            filtered_df[col] = vals.apply(lambda x: x / 100 if pd.notnull(x) and x > 1 else x)

    aggregated_df = (
        filtered_df
        .groupby('Date_clean', as_index=False)[percent_cols]
        .mean()
        .sort_values('Date_clean')
    )
    aggregated_df['Date_str'] = aggregated_df['Date_clean'].dt.strftime('%b %Y')
    return aggregated_df

def workload_and_paperless():
    df = Workload[(Workload['MXS_name'] == branch_name) & (Workload['Date'] >= pd.to_datetime(tarixden)) & (Workload['Date'] <= pd.to_datetime(tarixe))].sort_values('Date').fillna(0).copy()
    df['Date_str'] = df['Date'].dt.strftime('%b %Y')

    # 📌 1-ci sıra
    w1, w2, w3 = st.columns(3)
    with w1:
        st.plotly_chart(create_workload_kpi_chart(
            df['Date_str'], df['Total_daily_customer'], df['Avg_waiting_time'],
            'Avg WT / Customer', 'Customer count', 'Avg Waiting Time',
            bar_color='rgba(173, 217, 237, 0.8)', line_color='#D2691E'
        ), use_container_width=True)

    with w2:
        st.plotly_chart(create_workload_kpi_chart(
            df['Date_str'], df['Avg_st'], df['St_kpi'],
            'Avg ST / ST KPI', 'Avg Service Time', 'ST KPI Rate',
            bar_color='rgba(19, 151, 214, 0.8)', line_color='#D2691E', is_percent=True
        ), use_container_width=True)

    with w3:
        st.plotly_chart(create_workload_kpi_chart2(
            df['Date_str'], df['Credit_load'], df['Card_load'], df['Other_load'],
            'Service Workload Distribution'
        ), use_container_width=True)

    # 📌 2-ci sıra
    w4, w5, w6 = st.columns(3)

    with w4:
        st.plotly_chart(create_ST_chart(
            df['Date_str'], df['Monthly_served_customer'], df['Monthly_not_served_customer'], [0]*len(df),
            'Customer Distribution'
        ), use_container_width=True)

    with w5:
        paperless_df = get_paperless_data()
        fig_paperless = create_paperless_chart(paperless_df, f"Paperless Performance")
        st.plotly_chart(fig_paperless, use_container_width=True, config={'displaylogo': False})

    # 📌 3-cü sütun (w6): Baza DataFrame-i üzərindən filtrlənmiş vizual
    # 📌 3-cü sütun (w6): Baza cədvəlinin düzgün cəmlənməsi ilə
    with w6:
        b_df = Baza.copy()
        b_df.columns = b_df.columns.str.strip()
        
        mxs_c = next(c for c in ['MXS_name', 'Branch', 'Filial', 'mxs_name'] if c in b_df.columns)
        date_c = next(c for c in ['Date', 'DATE', 'Tarix', 'tarix'] if c in b_df.columns)
        
        b_df['Date_clean'] = pd.to_datetime(b_df[date_c], dayfirst=True, errors='coerce')
        
        target = str(branch_name).strip().lower().replace('filialı', '').replace('filiali', '').replace('mxs', '').strip()
        alt_target = target.replace('qazax', 'gazakh').replace('gazakh', 'qazax')
        
        b_df['mxs_clean'] = b_df[mxs_c].astype(str).str.strip().str.lower()
        
        mask_baza = (
            (b_df['mxs_clean'].str.contains(target, regex=False, na=False) | 
             b_df['mxs_clean'].str.contains(alt_target, regex=False, na=False)) &
            (b_df['Date_clean'] >= pd.to_datetime(tarixden)) &
            (b_df['Date_clean'] <= pd.to_datetime(tarixe))
        )
        f_baza = b_df.loc[mask_baza].sort_values('Date_clean').fillna(0).copy()
        
        # 📌 HƏLL: Faiz sütununu rəqəmə çeviririk
        if 'Percentage' in f_baza.columns:
            s_p = f_baza['Percentage'].astype(str).str.replace('%', '', regex=False).str.replace(',', '.', regex=False).str.strip()
            f_baza['Perc_clean'] = pd.to_numeric(s_p, errors='coerce').fillna(0)
            f_baza['Perc_clean'] = f_baza['Perc_clean'].apply(lambda x: x / 100 if x > 1 else x)
        else:
            f_baza['Perc_clean'] = 0

        # 📌 HƏLL: Təkrarlanan ayları Qruplaşdırırıq (Groupby)
        f_baza_grouped = f_baza.groupby('Date_clean', as_index=False).agg({
            'Total_count': 'sum',      # Barları cəmləyirik
            'Perc_clean': 'mean'       # Faizlərin ortalamasını götürürük
        }).sort_values('Date_clean')

        f_baza_grouped['Date_str'] = f_baza_grouped['Date_clean'].dt.strftime('%b %Y')

        st.plotly_chart(create_workload_kpi_chart(
            f_baza_grouped['Date_str'], 
            f_baza_grouped['Total_count'], 
            f_baza_grouped['Perc_clean'],
            'Total Calls / Answer rate', 
            'Calls', 
            'Answer rate',
            bar_color='rgba(19, 151, 214, 0.8)', 
            line_color='#D2691E', 
            is_percent=True
        ), use_container_width=True)

workload_and_paperless()
# --- NƏZƏRƏ ÇARPAN GÜCLÜ MƏSAFƏ (BOŞLUQ) ---
st.write("")
st.write("")
st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)





# ==========================================
# 7. BRANCH OVERVIEW
# ==========================================

st.subheader("Branch Overview", divider="grey")

# 1. Müxtəlif sütun adları ilə işləyən dinamik modulyar funksiya (fig1, fig2, fig3 üçün)
def create_distribution_chart(data, title, primary_color, secondary_color, cat_col="Category", type_col="Salary /Street"):
    fig = go.Figure() # Hər ehtimala qarşı boş obyekt yaradırıq

    if data is not None and isinstance(data, pd.DataFrame) and not data.empty and "MXS_name" in data.columns:
        filtered_data = data[data["MXS_name"] == branch_name]
        categories = filtered_data[cat_col].unique().tolist() if not filtered_data.empty else []

        salary_vals, street_vals = [], []
        for cat in categories:
            s_val = filtered_data[(filtered_data[cat_col] == cat) & (filtered_data[type_col] == "Salary")]["Count"].sum() if "Count" in filtered_data.columns else 0
            st_val = filtered_data[(filtered_data[cat_col] == cat) & (filtered_data[type_col] == "Street")]["Count"].sum() if "Count" in filtered_data.columns else 0
            salary_vals.append(s_val)
            street_vals.append(st_val)

        max_val = max(max(salary_vals, default=0), max(street_vals, default=0))
        y_upper = max_val * 1.25 if max_val > 0 else 10

        # Salary Bar
        fig.add_trace(go.Bar(
            x=categories, y=salary_vals, name='Salary',
            marker=dict(color=primary_color, cornerradius=4),
            text=[f"{int(v):,}" if v > 0 else "" for v in salary_vals],
            textposition='outside',
            textfont=dict(size=10, color='#000000', family='Segoe UI', weight='bold')
        ))
        
        # Street Bar
        fig.add_trace(go.Bar(
            x=categories, y=street_vals, name='Street',
            marker=dict(color=secondary_color, cornerradius=4),
            text=[f"{int(v):,}" if v > 0 else "" for v in street_vals],
            textposition='outside',
            textfont=dict(size=10, color='#000000', family='Segoe UI', weight='bold')
        ))

        fig.update_layout(yaxis=dict(range=[0, y_upper]))

    # Layout
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>", 
            font=dict(size=11, color="#2C3E50", family="Segoe UI"), 
            x=0, xanchor='left', y=0.90, yanchor='bottom'
        ),
        barmode='group', 
        template='plotly_white', 
        height=320, 
        margin=dict(l=10, r=10, t=60, b=20),
        legend=dict(
            orientation="h", yanchor="bottom", y=0.98, xanchor="right", x=1, 
            itemwidth=30,
            font=dict(size=10, color='#2C3E50')
        ),
        yaxis=dict(showgrid=True, gridcolor='#F0F4F8', zeroline=False),
        xaxis=dict(showgrid=False, tickfont=dict(size=10, color='#000000', weight='bold')),
        paper_bgcolor='rgba(0,0,0,0)', 
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig


def create_market_share_donut(data, title):
    # Filial adını filterləyirik
    filtered_data = data[data["MXS_name"] == branch_name]
    
    # Dəyərləri götürürük
    abb_val = filtered_data["ABB"].values[0]
    others_val = filtered_data["Others"].values[0]

    fig = go.Figure(data=[go.Pie(
        labels=['ABB', 'Others'],
        values=[abb_val, others_val],
        hole=0.6,
        marker=dict(colors=['#002B66', '#00A3E0']),
        textinfo='percent',
        textfont=dict(size=11, family='Segoe UI', weight='bold'),
        hovertemplate="<b>%{label}</b>: %{value:,.0f} ₼<extra></extra>"
    )])

    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>", 
            font=dict(size=11, color="#2C3E50", family="Segoe UI"), 
            x=0, xanchor='left', y=0.90, yanchor='bottom'
        ),
        template='plotly_white', 
        height=320, 
        margin=dict(l=10, r=10, t=60, b=20),
        legend=dict(
            orientation="h", yanchor="bottom", y=0.98, xanchor="right", x=1, 
            itemwidth=30,
            font=dict(size=10, color='#2C3E50')
        ),
        paper_bgcolor='rgba(0,0,0,0)', 
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig


# 3. Chart-ların yaradılması (.get() ilə təhlükəsiz çağırma)
fig1 = create_distribution_chart(
    data=data_dict.get('Potential_customer'),
    title="POTENTIAL CUSTOMER DISTRIBUTION",
    primary_color="#002B66",
    secondary_color="#00A3E0"
)

fig2 = create_distribution_chart(
    data=data_dict.get('Portfolio_distribution'),
    title="PORTFOLIO DISTRIBUTION",
    primary_color="#0A2540",
    secondary_color="#008080"
)

fig3 = create_distribution_chart(
    data=data_dict.get('Kart_sayi'),
    title="KART SAYI DISTRIBUTION",
    primary_color="#1E40AF",
    secondary_color="#4FA3D1",
    cat_col="Segment",
    type_col="Percentage"
)

fig4 = create_market_share_donut(
    data=data_dict.get('Bazar_payi'),
    title=" CREDİT MARKET SHARE DISTRIBUTION"
)


# 4. Streamlit Sütunlar
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.plotly_chart(fig1, use_container_width=True, config={'displaylogo': False})

with col2:
    st.plotly_chart(fig2, use_container_width=True, config={'displaylogo': False})

with col3:
    st.plotly_chart(fig3, use_container_width=True, config={'displaylogo': False})

with col4:
    st.plotly_chart(fig4, use_container_width=True, config={'displaylogo': False})

st.write("")
st.write("")
st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)




# ==========================================
# 7. EMPLOYEE STATISTICS
# ==========================================
st.subheader("Employees Statistics", divider="grey")

def branch_review_employee_stats():
    current_stf = Employees[Employees['MXS_name'] == branch_name].sort_values(by='Date', ascending=False).copy()

    # -------------------------------------------------------------
    # A. ST_KENARLASMA DATASI
    # -------------------------------------------------------------
    st_key = next((k for k in data_dict.keys() if 'kenarlasma' in str(k).lower().strip()), None)
    if st_key and not data_dict[st_key].empty:
        st_df = data_dict[st_key].copy()
        st_df.columns = st_df.columns.astype(str).str.strip()
        date_col = next((c for c in ['Date', 'DATE', 'Tarix', 'tarix', 'TARIX'] if c in st_df.columns), None)
        mxs_col = next((c for c in ['MXS', 'MXS_name', 'Filial', 'Branch'] if c in st_df.columns), None)
        staff_col = next((c for c in ['Staff', 'Əməkdaş', 'Emekdas', 'Employee'] if c in st_df.columns), None)

        if date_col and mxs_col and staff_col:
            st_df['Date_clean'] = pd.to_datetime(st_df[date_col], dayfirst=True, errors='coerce')
            st_mask = (
                (st_df[mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()) &
                (st_df['Date_clean'] >= pd.to_datetime(tarixden)) &
                (st_df['Date_clean'] <= pd.to_datetime(tarixe))
            )
            st_filtered = st_df.loc[st_mask].copy()

            if 'KENARLASMA_FAIZI' in st_filtered.columns:
                st_filtered['KENARLASMA_FAIZI'] = pd.to_numeric(
                    st_filtered['KENARLASMA_FAIZI'].astype(str).str.replace('%', '', regex=False).str.replace(',', '.', regex=False),
                    errors='coerce'
                )
                st_grouped = st_filtered.groupby(staff_col, as_index=False)['KENARLASMA_FAIZI'].mean()
                current_stf = pd.merge(current_stf, st_grouped, left_on='Staff', right_on=staff_col, how='left')
                current_stf['Kənarlaşma %'] = current_stf['KENARLASMA_FAIZI'].apply(
                    lambda x: f"{x:.1f}%" if pd.notnull(x) and not pd.isna(x) else "-"
                )
            else:
                current_stf['Kənarlaşma %'] = "-"
        else:
            current_stf['Kənarlaşma %'] = "-"
    else:
        current_stf['Kənarlaşma %'] = "-"


    # -------------------------------------------------------------
    # F. İMTAHAN NƏTİCƏLƏRİ SHEET-İNİN İŞLƏNMƏSİ
    # -------------------------------------------------------------
    exam_key = next((k for k in data_dict.keys() if 'imtahan' in str(k).lower().replace('i̇', 'i')), None)

    if exam_key and not data_dict[exam_key].empty:
        ex_df = data_dict[exam_key].copy()

        if any('unnamed' in str(c).lower() for c in ex_df.columns):
            for idx, row in ex_df.head(10).iterrows():
                row_str = " ".join([str(val).lower() for val in row.values])
                if ('employee' in row_str or 'əməkdaş' in row_str or 'user' in row_str) and ('il' in row_str or 'rub' in row_str):
                    ex_df.columns = ex_df.iloc[idx].astype(str).str.strip()
                    ex_df = ex_df.iloc[idx + 1:].reset_index(drop=True)
                    break
        else:
            ex_df.columns = ex_df.columns.astype(str).str.strip()

        def clean_col_name(col):
            return str(col).strip().lower().replace('i̇', 'i').replace('ı', 'i').replace('ə', 'e')

        ex_emp_col = next((c for c in ex_df.columns if clean_col_name(c) in ['employee', 'user', 'staff', 'emekdas', 'əmekdas']), None)
        ex_year_col = next((c for c in ex_df.columns if clean_col_name(c) in ['il', 'year']), None)
        ex_rub_col = next((c for c in ex_df.columns if clean_col_name(c) in ['rub', 'rubu', 'quarter']), None)
        ex_score_col = next((c for c in ex_df.columns if 'netice' in clean_col_name(c) or 'imtahan' in clean_col_name(c) or 'score' in clean_col_name(c)), None)

        if ex_emp_col and ex_year_col and ex_rub_col and ex_score_col:
            # Qiyməti təmizləyib faiz formatına salan funksiya
            def clean_val(val):
                val_str = str(val).strip()
                try:
                    num_val = float(val_str.replace('%', '').replace(',', '.'))
                    if 0 < num_val <= 1.0:
                        return f"{int(round(num_val * 100))}%"
                    else:
                        return f"{int(round(num_val))}%"
                except:
                    return val_str

            ex_df['clean_yr'] = ex_df[ex_year_col].astype(str).str.strip().str.split('.').str[0]
            ex_df['clean_rub'] = ex_df[ex_rub_col].astype(str).str.strip()
            ex_df['clean_score'] = ex_df[ex_score_col].apply(clean_val)

            # İl üzrə rübləri Q1-dən Q4-ə qədər sıralayıb ayrıca HTML sətrinə çeviririk
            def group_quarters(group):
                quarter_order = {'Q1': 1, 'Q2': 2, 'Q3': 3, 'Q4': 4}
                group = group.sort_values(
                    by='clean_rub',
                    key=lambda values: values.map(
                        lambda value: quarter_order.get(str(value).strip().upper(), 99)
                    )
                )
                q_list = [f"{row['clean_rub']} {row['clean_score']}" for _, row in group.iterrows()]
                yr = group['clean_yr'].iloc[0]
                return f"<div><strong>{yr}:</strong> {' '.join(q_list)}</div>"

            year_grouped = ex_df.groupby([ex_emp_col, 'clean_yr']).apply(group_quarters).reset_index(name='Year_Line')

            # Hər əməkdaş üçün illəri alt-alta yığırıq
            ex_summary = year_grouped.groupby(ex_emp_col)['Year_Line'].apply(lambda x: "".join(x)).reset_index()

            def get_exam_result_for_staff(staff_name):
                st_clean = str(staff_name).strip().lower()
                
                # Tam adı üst-üstə düşənlər
                match = ex_summary[ex_summary[ex_emp_col].astype(str).str.strip().str.lower() == st_clean]
                if not match.empty:
                    return match['Year_Line'].values[0]
                
                # Ad və Soyad uyğunlaşdırılması
                st_words = [w for w in st_clean.split() if w not in ['oğlu', 'qızı', 'oglu', 'qizi'] and len(w) > 2]
                for _, row in ex_summary.iterrows():
                    emp_name = str(row[ex_emp_col]).strip().lower()
                    matched_words = sum(1 for word in st_words if word in emp_name)
                    if matched_words >= min(2, len(st_words)):
                        return row['Year_Line']
                        
                return "-"

            current_stf['Exam_Result'] = current_stf['Staff'].apply(get_exam_result_for_staff)
        else:
            current_stf['Exam_Result'] = "-"
    else:
        current_stf['Exam_Result'] = "-"

    current_stf['Exam_Result'] = current_stf['Exam_Result'].fillna("-")


    # -------------------------------------------------------------
    # B. KREDIT_NOQSAN SHEET-İ
    # -------------------------------------------------------------
    kredit_noqsan_key = next((k for k in data_dict.keys() if 'kredit_noqsan' in str(k).lower().strip()), None)
    if kredit_noqsan_key and not data_dict[kredit_noqsan_key].empty:
        kr_df = data_dict[kredit_noqsan_key].copy()
        kr_df.columns = kr_df.columns.astype(str).str.strip()
        kr_date_col = next((c for c in ['Date', 'DATE', 'Tarix', 'tarix'] if c in kr_df.columns), None)
        kr_mxs_col = next((c for c in ['MXS_name', 'MXS', 'Filial', 'Branch'] if c in kr_df.columns), None)
        kr_emp_col = next((c for c in ['Employee', 'Staff', 'Əməkdaş', 'Emekdas'] if c in kr_df.columns), None)
        kr_type_col = next((c for c in ['Type', 'Nöqsan növləri', 'Nöqsan Tipi'] if c in kr_df.columns), None)

        if kr_date_col and kr_mxs_col and kr_emp_col and kr_type_col:
            kr_df['Date_clean'] = pd.to_datetime(kr_df[kr_date_col], dayfirst=True, errors='coerce')
            kr_mask = (
                (kr_df[kr_mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()) &
                (kr_df['Date_clean'] >= pd.to_datetime(tarixden)) &
                (kr_df['Date_clean'] <= pd.to_datetime(tarixe))
            )
            kr_filtered = kr_df.loc[kr_mask].copy()
            kr_filtered['Is_High_Risk_Kr'] = kr_filtered[kr_type_col].astype(str).str.strip().str.upper() == 'YÜKSƏK'
            kr_filtered['Is_Other_Error_Kr'] = kr_filtered[kr_type_col].astype(str).str.strip().str.upper() == 'DİGƏR'

            kr_summary = kr_filtered.groupby(kr_emp_col).agg(
                High_Risk_Kr_Count=('Is_High_Risk_Kr', 'sum'),
                Other_Risk_Kr_Count=('Is_Other_Error_Kr', 'sum')
            ).reset_index()

            current_stf = pd.merge(current_stf, kr_summary, left_on='Staff', right_on=kr_emp_col, how='left')
            current_stf['Kr_Yuksek'] = current_stf['High_Risk_Kr_Count'].apply(lambda x: int(x) if pd.notnull(x) and x > 0 else "0")
            current_stf['Kr_Diger'] = current_stf['Other_Risk_Kr_Count'].apply(lambda x: int(x) if pd.notnull(x) and x > 0 else "0")
        else:
            current_stf['Kr_Yuksek'] = "-"
            current_stf['Kr_Diger'] = "-"
    else:
        current_stf['Kr_Yuksek'] = "-"
        current_stf['Kr_Diger'] = "-"

    # -------------------------------------------------------------
    # C. NOQSAN SHEET-İ
    # -------------------------------------------------------------
    noqsan_key = next((k for k in data_dict.keys() if str(k).strip().lower() == 'noqsan'), None)
    if noqsan_key and not data_dict[noqsan_key].empty:
        noqsan_df = data_dict[noqsan_key].copy()
        noqsan_df.columns = noqsan_df.columns.astype(str).str.strip()
        n_date_col = next((c for c in ['Date', 'DATE', 'Tarix', 'tarix'] if c in noqsan_df.columns), None)
        n_mxs_col = next((c for c in ['MXS_name', 'MXS', 'Filial', 'Branch'] if c in noqsan_df.columns), None)
        n_emp_col = next((c for c in ['Employee', 'Staff', 'Əməkdaş', 'Emekdas'] if c in noqsan_df.columns), None)
        n_type_col = next((c for c in ['Type', 'Nöqsan növləri', 'Nöqsan Tipi'] if c in noqsan_df.columns), None)

        if n_date_col and n_mxs_col and n_emp_col and n_type_col:
            noqsan_df['Date_clean'] = pd.to_datetime(noqsan_df[n_date_col], dayfirst=True, errors='coerce')
            n_mask = (
                (noqsan_df[n_mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()) &
                (noqsan_df['Date_clean'] >= pd.to_datetime(tarixden)) &
                (noqsan_df['Date_clean'] <= pd.to_datetime(tarixe))
            )
            n_filtered = noqsan_df.loc[n_mask].copy()
            n_filtered['Is_High_Risk'] = n_filtered[n_type_col].astype(str).str.strip().str.lower() == 'yüksək riskli'
            n_filtered['Is_Other_Error'] = n_filtered[n_type_col].astype(str).str.strip().str.lower() == 'digər səhv'

            noqsan_summary = n_filtered.groupby(n_emp_col).agg(
                High_Risk_Count=('Is_High_Risk', 'sum'),
                Other_Risk_Count=('Is_Other_Error', 'sum')
            ).reset_index()

            current_stf = pd.merge(current_stf, noqsan_summary, left_on='Staff', right_on=n_emp_col, how='left')
            current_stf['QKr_Yuksek'] = current_stf['High_Risk_Count'].apply(lambda x: int(x) if pd.notnull(x) and x > 0 else "0")
            current_stf['QKr_Diger'] = current_stf['Other_Risk_Count'].apply(lambda x: int(x) if pd.notnull(x) and x > 0 else "0")
        else:
            current_stf['QKr_Yuksek'] = "-"
            current_stf['QKr_Diger'] = "-"
    else:
        current_stf['QKr_Yuksek'] = "-"
        current_stf['QKr_Diger'] = "-"

    # -------------------------------------------------------------
    # D. SCORE FAYLI
    # -------------------------------------------------------------
    score_key = next((k for k in data_dict.keys() if 'score' in str(k).lower().strip()), None)
    if score_key and not data_dict[score_key].empty:
        score_df = data_dict[score_key].copy()
        score_df.columns = score_df.columns.astype(str).str.strip()
        sc_date_col = next((c for c in ['Date', 'DATE', 'Tarix', 'tarix'] if c in score_df.columns), None)
        sc_mxs_col = next((c for c in ['MXS_name', 'MXS', 'Filial', 'Branch'] if c in score_df.columns), None)
        sc_emp_col = next((c for c in ['User_name', 'Staff', 'Əməkdaş', 'Emekdas', 'Employee'] if c in score_df.columns), None)
        sc_val_col = next((c for c in ['Score', 'SCORE', 'Xal', 'Bal'] if c in score_df.columns), None)

        if sc_date_col and sc_mxs_col and sc_emp_col and sc_val_col:
            score_df['Date_clean'] = pd.to_datetime(score_df[sc_date_col], dayfirst=True, errors='coerce')
            score_df['mxs_clean'] = score_df[sc_mxs_col].astype(str).str.strip().str.lower()
            sc_mask = (
                (score_df['mxs_clean'] == str(branch_name).strip().lower()) &
                (score_df['Date_clean'] >= pd.to_datetime(tarixden)) &
                (score_df['Date_clean'] <= pd.to_datetime(tarixe))
            )
            sc_filtered = score_df.loc[sc_mask].copy()
            sc_filtered[sc_val_col] = pd.to_numeric(sc_filtered[sc_val_col].astype(str).str.replace(',', '.', regex=False), errors='coerce')
            sc_grouped = sc_filtered.groupby(sc_emp_col, as_index=False)[sc_val_col].mean()

            current_stf = pd.merge(current_stf, sc_grouped, left_on='Staff', right_on=sc_emp_col, how='left')
            current_stf['Score_Val'] = current_stf[sc_val_col].apply(
                lambda x: f"{x:.1f}" if pd.notnull(x) and not pd.isna(x) else "-"
            )
        else:
            current_stf['Score_Val'] = "-"
    else:
        current_stf['Score_Val'] = "-"


    # -------------------------------------------------------------
    # E. HTML GRID RENDER
    # -------------------------------------------------------------
    selected_cols = ['Staff', 'Position', 'Kənarlaşma %', 'Score_Val', 'Exam_Result', 'Kr_Yuksek', 'Kr_Diger', 'QKr_Yuksek', 'QKr_Diger']
    final_df = current_stf[selected_cols].copy()
    total_rows = len(final_df)

    rows_html = ""
    for idx, (_, row) in enumerate(final_df.iterrows()):
        is_last_row = (idx == total_rows - 1)
        bottom_left_class = "bottom-left-rounded" if is_last_row else ""

        rows_html += f"""
        <div class="data-cell {bottom_left_class}">
            <div class="emp-name">{row['Staff']}</div>
            <div class="emp-pos">{row['Position']}</div>
        </div>
        <div class="data-cell justify-center">
            <div class="value-bold">{row['Kənarlaşma %']}</div>
        </div>
        <div class="data-cell justify-center">
            <div class="value-bold">{row['Score_Val']}</div>
        </div>
        <div class="data-cell justify-center">
            <div class="exam-container">{row['Exam_Result']}</div>
        </div>
        <div class="data-cell justify-center">
            <div class="compact-stats">
                <span>Y.Risk: <strong>{row['Kr_Yuksek']}</strong></span>
                <span>Digər: <strong>{row['Kr_Diger']}</strong></span>
            </div>
        </div>
        <div class="data-cell justify-center">
            <div class="compact-stats">
                <span>Y.Risk: <strong>{row['QKr_Yuksek']}</strong></span>
                <span>Digər: <strong>{row['QKr_Diger']}</strong></span>
            </div>
        </div>
        """

    custom_component = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{ font-family: 'Segoe UI', -apple-system, sans-serif; background-color: transparent; margin: 0; padding: 4px; }}
        .grid-table {{ display: grid; grid-template-columns: 1.1fr 0.6fr 0.5fr 1.3fr 0.8fr 0.8fr; gap: 5px 8px; width: 100%; box-sizing: border-box; }}
        .header-cell {{ background: linear-gradient(135deg, #4ea3e3 0%, #005bb5 100%); color: #ffffff; font-weight: 700; font-size: 11px; text-transform: uppercase; padding: 8px 3px; text-align: center; letter-spacing: 0.2px; border-top-right-radius: 10px; border-top-left-radius: 10px; display: flex; align-items: center; justify-content: center; }}
        .data-cell {{ background: linear-gradient(180deg, #f0f7fe 0%, #e2f0fc 100%); padding: 6px 8px; display: flex; flex-direction: column; justify-content: center; min-height: 42px; box-sizing: border-box; }}
        .justify-center {{ align-items: center; text-align: center; }}
        .bottom-left-rounded {{ border-bottom-left-radius: 10px; }}
        .emp-name {{ color: #2c3e50; font-weight: 700; font-size: 12px; line-height: 1.1; }}
        .emp-pos {{ color: #64748b; font-size: 10px; margin-top: 1px; }}
        .value-bold {{ color: #2c3e50; font-weight: 800; font-size: 12.5px; }}
        .compact-stats {{ display: flex; flex-direction: column; align-items: center; gap: 1px; font-size: 10.5px; color: #5a6e7f; white-space: nowrap; }}
        .compact-stats strong {{ color: #1e293b; font-weight: 800; }}
        .exam-container {{ font-size: 11px; color: #1e293b; line-height: 1.3; text-align: center; white-space: nowrap; }}
    </style>
    </head>
    <body>
    <div class="grid-table">
        <div class="header-cell">Əməkdaş</div>
        <div class="header-cell">Kənarlaşma</div>
        <div class="header-cell">Score</div>
        <div class="header-cell">İmtahan</div>
        <div class="header-cell">Kredit Nöqsan</div>
        <div class="header-cell">Qeyri-Kredit Nöqsan</div>
        {rows_html}
    </div>
    </body>
    </html>
    """

    calc_height = max(250, (total_rows + 1) * 55)
    components.html(custom_component, height=calc_height, scrolling=True)


# =============================================================
# İMTAHAN NƏTİCƏLƏRİ (BİRBAŞA 'İmtahan' DƏYİŞƏNİ İLƏ)
# =============================================================

if 'İmtahan' in globals() and not İmtahan.empty:
    ex_df = İmtahan.copy()
    ex_df.columns = ex_df.columns.astype(str).str.strip()

    # Sütun adlarını dəqiq tapmaq üçün köməkçi funksiya
    def get_col(possible_names):
        for c in ex_df.columns:
            clean_c = str(c).strip().lower().replace('i̇', 'i').replace('ı', 'i').replace('ə', 'e')
            if any(p in clean_c for p in possible_names):
                return c
        return None

    ex_emp_col = get_col(['employee', 'əməkdaş', 'emekdas', 'user', 'staff'])
    ex_year_col = get_col(['il', 'year'])
    ex_rub_col = get_col(['rub', 'quarter', 'rüb'])
    ex_score_col = get_col(['netice', 'neticə', 'imtahan', 'score', 'bal'])
    ex_mxs_col = get_col(['mxs_name', 'mxs', 'filial', 'branch'])

    if ex_emp_col and ex_year_col and ex_rub_col and ex_score_col:
        
        # Filial filtri (Noqsan koddakı kimi branch_name dəyişəninə əsasən)
        if ex_mxs_col and 'branch_name' in globals() and branch_name:
            e_mask = ex_df[ex_mxs_col].astype(str).str.strip().str.lower() == str(branch_name).strip().lower()
            ex_filtered = ex_df.loc[e_mask].copy()
        else:
            ex_filtered = ex_df.copy()

        # Faiz / Qiyməti formatlayıcı (məsələn: 0.25 -> 25%)
        def format_score(val):
            if pd.isna(val) or str(val).strip() in ['', '-', 'nan', 'None']:
                return ""
            val_str = str(val).strip().replace('%', '').replace(',', '.')
            try:
                num_val = float(val_str)
                if 0 < num_val <= 1.0:
                    return f"{int(round(num_val * 100))}%"
                else:
                    return f"{int(round(num_val))}%"
            except:
                return str(val).strip()

        # Əməkdaş adındakı fərqləri (ə/e, ı/i, oğlu, qızı) aradan qaldıran funksiya
        def clean_name(val):
            if pd.isna(val): return ""
            s = str(val).strip().lower()
            char_map = {'ə': 'e', 'ı': 'i', 'i̇': 'i', 'ö': 'o', 'ü': 'u', 'ğ': 'g', 'ç': 'c', 'ş': 's'}
            for k, v in char_map.items():
                s = s.replace(k, v)
            for stop_word in ['oglu', 'qizi', 'oğlu', 'qızı']:
                s = s.replace(stop_word, '')
            return "".join([c for c in s if c.isalpha()])

        ex_filtered['clean_emp'] = ex_filtered[ex_emp_col].apply(clean_name)
        ex_filtered['clean_yr'] = ex_filtered[ex_year_col].astype(str).str.strip().str.split('.').str[0]
        ex_filtered['clean_rub'] = ex_filtered[ex_rub_col].astype(str).str.strip()
        ex_filtered['clean_score'] = ex_filtered[ex_score_col].apply(format_score)

        # Məlumatları qruplayıb "2025 Q1-25% Q2-15%" şəklində yığırıq
        exam_dict = {}
        for clean_emp, group in ex_filtered.groupby('clean_emp'):
            if not clean_emp: continue
            year_lines = []
            
            for yr, yr_group in group.groupby('clean_yr'):
                q_items = []
                for _, r in yr_group.iterrows():
                    score = r['clean_score']
                    if score:
                        q_items.append(f"{r['clean_rub']}-{score}")
                
                if q_items:
                    year_lines.append(f"{yr} " + " ".join(q_items))
            
            if year_lines:
                exam_dict[clean_emp] = "\n".join(year_lines)

        # Əsas cədvələ uyğunlaşdıran funksiya
        def get_exam_result(staff_name):
            c_name = clean_name(staff_name)
            if not c_name:
                return "-"
            
            # 1. Tam ad uyğunluğu
            if c_name in exam_dict:
                return exam_dict[c_name]
            
            # 2. Hissəvi (Ad/Soyad) uyğunluq
            for dict_emp_name, result_text in exam_dict.items():
                if len(c_name) > 4 and (c_name in dict_emp_name or dict_emp_name in c_name):
                    return result_text
            
            return "-"

        # Nəticəni cədvələ yazırıq
        if 'current_stf' in globals() and 'Staff' in current_stf.columns:
            current_stf['Exam_Result'] = current_stf['Staff'].apply(get_exam_result)
            current_stf['Exam_Result'] = current_stf['Exam_Result'].fillna("-")

branch_review_employee_stats()