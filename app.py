import streamlit as st
import streamlit.string_util as _su
# Allow Private Use Area Unicode (Flaticon uicons) in Streamlit icon parameters
_orig_is_emoji = _su.is_emoji
def _flaticon_is_emoji(s):
    if s and isinstance(s, str) and any(ord(c) >= 0xE000 for c in s):
        return True
    return _orig_is_emoji(s)
_su.is_emoji = _flaticon_is_emoji

import io
import csv
import time
from datetime import datetime, timezone, timedelta
from openpyxl import Workbook, load_workbook
from supabase_service import SupabaseService
from sms_service import SMSService

# Sahifa konfiguratsiyasi
st.set_page_config(
    page_title="Odamlar Ro'yxati va SMS",
    layout="wide"
)

# Mavzu holati (Theme: dark / light)
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

is_dark = (st.session_state.theme == "dark")

theme_bg = "#0f172a" if is_dark else "#f8fafc"
theme_card_bg = "#1e293b" if is_dark else "#ffffff"
theme_border = "#334155" if is_dark else "#e2e8f0"
theme_text = "#f8fafc" if is_dark else "#0f172a"
theme_subtext = "#94a3b8" if is_dark else "#64748b"
theme_accent = "#38bdf8" if is_dark else "#0284c7"
theme_hr = "#334155" if is_dark else "#e2e8f0"
theme_row_hr = "#1e293b" if is_dark else "#f1f5f9"

# Flaticon uicons unicode belgilar (Barcha tugmalar va elementlar uchun)
FLATICON = {
    "user_add": "\uf0a8",         # fi-rr-user-add
    "upload": "\ue42c",           # fi-rr-cloud-upload-alt
    "download": "\ue411",         # fi-rr-cloud-download-alt
    "save": "\ue56a",             # fi-rr-disk
    "close": "\ue4cd",            # fi-rr-cross
    "select_all": "\ue9b2",       # fi-rr-list-check
    "deselect": "\ue364",         # fi-rr-checkbox
    "send": "\ueb50",             # fi-rr-paper-plane
    "chat": "\ue473",             # fi-rr-comment-sms
    "edit": "\ueb87",             # fi-rr-pencil
    "delete": "\uf026",           # fi-rr-trash
    "refresh": "\uece4",          # fi-rr-refresh
    "settings": "\ueda4",         # fi-rr-settings-sliders
    "sun": "\uef1e",              # fi-rr-sun
    "moon": "\uea96",             # fi-rr-moon
    "bolt": "\ue1ce",             # fi-rr-bolt
    "arrow_left": "\ue07a",       # fi-rr-angle-left
    "arrow_right": "\ue07b",      # fi-rr-angle-right
    "check": "\ue35f",            # fi-rr-check
    "badge_check": "\ue135",      # fi-rr-badge-check
    "info": "\ue8ce",             # fi-rr-info
    "warning": "\ue61e",          # fi-rr-exclamation
    "users": "\uf0e0",            # fi-rr-users-alt
    "document": "\ue57f",         # fi-rr-document
}

# Flaticon CDN va Dinamik CSS (Dark / Light)
if is_dark:
    st.markdown("""
<link rel="stylesheet" href="https://cdn-uicons.flaticon.com/2.6.0/uicons-regular-rounded/css/uicons-regular-rounded.css">
<link rel="stylesheet" href="https://cdn-uicons.flaticon.com/2.6.0/uicons-solid-rounded/css/uicons-solid-rounded.css">

<style>
    /* Deploy, menyu va ortiqcha bezaklarni yashirish */
    .stDeployButton,
    [data-testid="stDeployButton"],
    button[title="Deploy this app"],
    [data-testid="stDecoration"],
    #MainMenu,
    [data-testid="stToolbarActions"],
    button[title="View app options"],
    footer {
        visibility: hidden !important;
        display: none !important;
        height: 0 !important;
        opacity: 0 !important;
    }

    /* Asosiy fon */
    .stApp {
        background-color: #0f172a !important;
        color: #f8fafc !important;
    }
    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 0 !important;
        min-height: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        pointer-events: none !important;
    }

    /* Flaticon fontini Streamlit icon elementlariga majburiy qo'llash */
    span[data-testid="stIconEmoji"],
    span[data-testid="stIconMaterial"] {
        font-family: 'uicons-regular-rounded', 'uicons-solid-rounded', sans-serif !important;
        font-style: normal !important;
        font-weight: normal !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        line-height: 1 !important;
        font-size: 1.15rem !important;
    }

    /* Asosiy Tablar uchun Flaticon ikonkalari */
    [data-baseweb="tab"]:nth-child(1) [data-testid="stMarkdownContainer"] p::before {
        font-family: 'uicons-regular-rounded' !important;
        content: "\\f0e0" !important;
        margin-right: 8px !important;
        font-size: 1.15rem !important;
        display: inline-block !important;
        vertical-align: -2px !important;
    }
    [data-baseweb="tab"]:nth-child(2) [data-testid="stMarkdownContainer"] p::before {
        font-family: 'uicons-regular-rounded' !important;
        content: "\\e57f" !important;
        margin-right: 8px !important;
        font-size: 1.15rem !important;
        display: inline-block !important;
        vertical-align: -2px !important;
    }

    /* Odamcha (runner) va Stop tugmasi o'rniga o'ta zamonaviy Neon Spinner loader */
    [data-testid="stStatusWidget"] {
        visibility: visible !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: rgba(30, 41, 59, 0.92) !important;
        backdrop-filter: blur(10px) !important;
        border: 1px solid rgba(56, 189, 248, 0.4) !important;
        border-radius: 24px !important;
        padding: 6px 16px !important;
        position: fixed !important;
        top: 14px !important;
        right: 18px !important;
        z-index: 999999 !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.5), 0 0 12px rgba(56, 189, 248, 0.3) !important;
        pointer-events: none !important;
    }

    [data-testid="stStatusWidget"] * {
        display: none !important;
        visibility: hidden !important;
    }

    [data-testid="stStatusWidget"]::before {
        content: "" !important;
        display: inline-block !important;
        width: 14px !important;
        height: 14px !important;
        margin-right: 8px !important;
        border: 2.5px solid rgba(56, 189, 248, 0.2) !important;
        border-top-color: #38bdf8 !important;
        border-radius: 50% !important;
        animation: runner-spin 0.65s linear infinite !important;
        filter: drop-shadow(0 0 5px #38bdf8) !important;
    }

    [data-testid="stStatusWidget"]::after {
        content: "Yuklanmoqda..." !important;
        display: inline-block !important;
        color: #38bdf8 !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.3px !important;
    }

    @keyframes runner-spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }

    /* Sarlavha */
    .title-box {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 2px;
    }
    .title-icon {
        font-size: 1.8rem;
        color: #38bdf8;
        display: flex;
        align-items: center;
    }
    .title-text {
        font-size: 2rem;
        font-weight: 700;
        color: #f8fafc !important;
        margin: 0;
        padding: 0;
    }
    .sub-text {
        font-size: 0.92rem;
        color: #94a3b8 !important;
        margin-bottom: 1rem;
    }

    /* Telefon raqam uslubi */
    .phone-text {
        font-family: monospace;
        font-size: 1.05rem;
        color: #38bdf8;
        font-weight: 600;
    }

    .sender-badge {
        background: #0284c7;
        color: #ffffff;
        font-family: monospace;
        padding: 3px 10px;
        border-radius: 6px;
        font-size: 0.95rem;
        font-weight: 600;
    }

    /* Amallar tugmalari */
    div[data-testid="stHorizontalBlock"] button[kind="secondary"] {
        padding: 4px 6px !important;
        min-height: 34px !important;
        border: 1px solid #334155 !important;
        background-color: #1e293b !important;
        color: #f8fafc !important;
    }
    div[data-testid="stHorizontalBlock"] button[kind="secondary"]:hover {
        border-color: #38bdf8 !important;
        color: #38bdf8 !important;
    }

    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2.5rem !important;
    }

    /* Standart loader */
    div[data-testid="stSpinner"] {
        padding: 8px 16px !important;
        color: #38bdf8 !important;
    }
</style>
""", unsafe_allow_html=True)
else:
    st.markdown("""
<link rel="stylesheet" href="https://cdn-uicons.flaticon.com/2.6.0/uicons-regular-rounded/css/uicons-regular-rounded.css">
<link rel="stylesheet" href="https://cdn-uicons.flaticon.com/2.6.0/uicons-solid-rounded/css/uicons-solid-rounded.css">

<style>
    /* Deploy, menyu va ortiqcha bezaklarni yashirish */
    .stDeployButton,
    [data-testid="stDeployButton"],
    button[title="Deploy this app"],
    [data-testid="stDecoration"],
    #MainMenu,
    [data-testid="stToolbarActions"],
    button[title="View app options"],
    footer {
        visibility: hidden !important;
        display: none !important;
        height: 0 !important;
        opacity: 0 !important;
    }

    /* Asosiy fon va matnlar (Light Mode) */
    .stApp {
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 0 !important;
        min-height: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        pointer-events: none !important;
    }
    .stApp p, .stApp span, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6 {
        color: #0f172a !important;
    }

    /* Flaticon fontini Streamlit icon elementlariga majburiy qo'llash */
    span[data-testid="stIconEmoji"],
    span[data-testid="stIconMaterial"] {
        font-family: 'uicons-regular-rounded', 'uicons-solid-rounded', sans-serif !important;
        font-style: normal !important;
        font-weight: normal !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        line-height: 1 !important;
        font-size: 1.15rem !important;
    }

    /* Asosiy Tablar uchun Flaticon ikonkalari */
    [data-baseweb="tab"]:nth-child(1) [data-testid="stMarkdownContainer"] p::before {
        font-family: 'uicons-regular-rounded' !important;
        content: "\\f0e0" !important;
        margin-right: 8px !important;
        font-size: 1.15rem !important;
        display: inline-block !important;
        vertical-align: -2px !important;
    }
    [data-baseweb="tab"]:nth-child(2) [data-testid="stMarkdownContainer"] p::before {
        font-family: 'uicons-regular-rounded' !important;
        content: "\\e57f" !important;
        margin-right: 8px !important;
        font-size: 1.15rem !important;
        display: inline-block !important;
        vertical-align: -2px !important;
    }

    /* Loader (Light Mode) */
    [data-testid="stStatusWidget"] {
        visibility: visible !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(10px) !important;
        border: 1px solid rgba(2, 132, 199, 0.3) !important;
        border-radius: 24px !important;
        padding: 6px 16px !important;
        position: fixed !important;
        top: 14px !important;
        right: 18px !important;
        z-index: 999999 !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08) !important;
        pointer-events: none !important;
    }

    [data-testid="stStatusWidget"] * {
        display: none !important;
        visibility: hidden !important;
    }

    [data-testid="stStatusWidget"]::before {
        content: "" !important;
        display: inline-block !important;
        width: 14px !important;
        height: 14px !important;
        margin-right: 8px !important;
        border: 2.5px solid rgba(2, 132, 199, 0.2) !important;
        border-top-color: #0284c7 !important;
        border-radius: 50% !important;
        animation: runner-spin 0.65s linear infinite !important;
        filter: drop-shadow(0 0 3px #0284c7) !important;
    }

    [data-testid="stStatusWidget"]::after {
        content: "Yuklanmoqda..." !important;
        display: inline-block !important;
        color: #0284c7 !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.3px !important;
    }

    @keyframes runner-spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }

    /* Sarlavha */
    .title-box {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 2px;
    }
    .title-icon {
        font-size: 1.8rem;
        color: #0284c7;
        display: flex;
        align-items: center;
    }
    .title-text {
        font-size: 2rem;
        font-weight: 700;
        color: #0f172a !important;
        margin: 0;
        padding: 0;
    }
    .sub-text {
        font-size: 0.92rem;
        color: #64748b !important;
        margin-bottom: 1rem;
    }

    /* Telefon raqam uslubi */
    .phone-text {
        font-family: monospace;
        font-size: 1.05rem;
        color: #0284c7;
        font-weight: 600;
    }

    .sender-badge {
        background: #0284c7;
        color: #ffffff;
        font-family: monospace;
        padding: 3px 10px;
        border-radius: 6px;
        font-size: 0.95rem;
        font-weight: 600;
    }

    /* Kiritish maydonlari (Inputs & Selectboxes) - Toza oq Light Mode */
    [data-testid="stTextInput"] div,
    [data-testid="stTextInput"] input,
    [data-testid="stSelectbox"] div,
    [data-testid="stSelectbox"] [data-baseweb="select"],
    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stSelectbox"] [data-baseweb="select"] div,
    div[data-baseweb="input"],
    div[data-baseweb="input"] > div,
    div[data-baseweb="input"] input,
    div[data-baseweb="base-input"],
    div[data-baseweb="base-input"] > div,
    div[data-baseweb="select"],
    div[data-baseweb="select"] > div,
    div[data-baseweb="select"] div,
    div[data-baseweb="textarea"],
    div[data-baseweb="textarea"] > div,
    div[data-baseweb="textarea"] textarea {
        background-color: #ffffff !important;
        border-color: #cbd5e1 !important;
        color: #0f172a !important;
    }

    input, textarea {
        background-color: #ffffff !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }

    input::placeholder, textarea::placeholder {
        color: #94a3b8 !important;
        -webkit-text-fill-color: #94a3b8 !important;
    }

    div[data-baseweb="select"] span {
        color: #0f172a !important;
    }
    div[data-baseweb="select"] svg {
        fill: #64748b !important;
    }

    div[data-baseweb="popover"], ul[role="listbox"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
    }
    li[role="option"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    li[role="option"]:hover, li[role="option"][aria-selected="true"] {
        background-color: #f1f5f9 !important;
        color: #0284c7 !important;
    }

    /* Borderli konteynerlar (Filtr paneli va h.k.) */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05) !important;
    }

    /* Tablar */
    [data-baseweb="tab-list"] {
        border-bottom: 2px solid #e2e8f0 !important;
    }
    [data-baseweb="tab"] {
        color: #64748b !important;
    }
    [data-baseweb="tab"][aria-selected="true"] {
        color: #0284c7 !important;
        border-bottom-color: #0284c7 !important;
    }
    [data-baseweb="tab-highlight"] {
        background-color: #0284c7 !important;
    }

    /* Tugmalar */
    button[kind="secondary"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        color: #1e293b !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05) !important;
    }
    button[kind="secondary"]:hover {
        background-color: #f0f9ff !important;
        border-color: #0284c7 !important;
        color: #0284c7 !important;
    }
    button[kind="secondary"]:disabled {
        background-color: #f1f5f9 !important;
        border-color: #e2e8f0 !important;
        color: #94a3b8 !important;
    }
    div[data-testid="stHorizontalBlock"] button[kind="secondary"] {
        padding: 4px 6px !important;
        min-height: 34px !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
    }
    div[data-testid="stHorizontalBlock"] button[kind="secondary"]:hover {
        border-color: #0284c7 !important;
        color: #0284c7 !important;
        background-color: #f0f9ff !important;
    }
    button[kind="primary"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
        border: none !important;
    }
    button[kind="primary"]:hover {
        background-color: #0369a1 !important;
        color: #ffffff !important;
    }

    /* Modallar (Dialogs) */
    div[data-testid="stDialog"] > div > div {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        color: #0f172a !important;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-testid="stDialog"] h2,
    div[data-testid="stDialog"] span,
    div[data-testid="stDialog"] p,
    div[data-testid="stDialog"] label {
        color: #0f172a !important;
    }

    /* Checkbox - Toza oq va ko'k belgilangan Light Mode */
    [data-testid="stCheckbox"] span,
    [data-testid="stCheckbox"] [data-baseweb="checkbox"],
    [data-testid="stCheckbox"] [data-baseweb="checkbox"] > span,
    [data-testid="stCheckbox"] [data-baseweb="checkbox"] > div {
        background-color: #ffffff !important;
        border-color: #cbd5e1 !important;
        color: #0f172a !important;
    }
    [data-testid="stCheckbox"] input:checked + span,
    [data-testid="stCheckbox"] [aria-checked="true"] span,
    [data-testid="stCheckbox"] [aria-checked="true"] > div {
        background-color: #0284c7 !important;
        border-color: #0284c7 !important;
    }
    [data-testid="stCheckbox"] label span {
        color: #0f172a !important;
        background: transparent !important;
    }

    /* Fayl yuklash maydoni */
    [data-testid="stFileUploader"] section {
        background-color: #f8fafc !important;
        border: 2px dashed #cbd5e1 !important;
    }
    [data-testid="stFileUploader"] section * {
        color: #475569 !important;
    }

    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2.5rem !important;
    }

    /* Standart loader */
    div[data-testid="stSpinner"] {
        padding: 8px 16px !important;
        color: #0284c7 !important;
    }
</style>
""", unsafe_allow_html=True)

# Servislarni yuklash
@st.cache_resource(show_spinner=False)
def get_service():
    return SupabaseService()

def get_sms_service():
    return SMSService()

service = get_service()
sms_service = get_sms_service()

# ================= TEZKOR KESH (QOTISHNI YO'QOTISH UCHUN) =================
# Ma'lumotlarni RAM keshida saqlaymiz, har bir bosishda Supabase'ga so'rov ketmaydi!
def load_people_data(force_reload=False):
    if "cached_people" not in st.session_state or force_reload:
        st.session_state.cached_people = service.get_people()
    return st.session_state.cached_people

def refresh_cache():
    st.session_state.cached_people = service.get_people()

UZ_TZ = timezone(timedelta(hours=5))

def format_log_timestamp(iso_str: str) -> str:
    if not iso_str:
        return ""
    try:
        dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
        dt_uz = dt.astimezone(UZ_TZ)
        hour12 = dt_uz.strftime("%I").lstrip("0") or "12"
        return f"{dt_uz.month}/{dt_uz.day}/{dt_uz.year}, {hour12}:{dt_uz.strftime('%M:%S')} {dt_uz.strftime('%p')}"
    except Exception:
        return str(iso_str)[:19].replace("T", " ")

def load_message_logs(force_reload=False):
    if "cached_message_logs" not in st.session_state or force_reload:
        st.session_state.cached_message_logs = sms_service.get_message_logs(limit=100)
    return st.session_state.cached_message_logs

# Sessiyada tanlanganlar to'plami
if "selected_ids" not in st.session_state:
    st.session_state.selected_ids = set()

# Joriy sahifa raqami (Paginatsiya uchun)
if "page_number" not in st.session_state:
    st.session_state.page_number = 1

PAGE_SIZE = 15  # Har bir sahifada 15 tadan odam (juda tez ishlashi uchun)

# Checkbox o'zgarganda ishlaydigan tezkor funksiya (Rerun'siz)
def on_check_toggle(pid):
    val = st.session_state.get(f"chk_{pid}", False)
    if val:
        st.session_state.selected_ids.add(pid)
    else:
        st.session_state.selected_ids.discard(pid)


# ================= YORDAMCHI EXCEL FUNKSIYALARI =================

def export_to_excel_bytes(people_data: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Odamlar"
    ws.append(["FullName", "PhoneNumberi"])
    if people_data:
        for p in people_data:
            name_val = p.get("full_name") or p.get("FullName") or ""
            phone_val = p.get("phone_number") or p.get("PhoneNumberi") or ""
            ws.append([str(name_val).strip(), str(phone_val).strip()])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def generate_template_excel() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Shablon"
    ws.append(["FullName", "PhoneNumberi"])
    ws.append(["Karimov Sardor Baxtiyorovich", "+998901234567"])
    ws.append(["Toshmatov Jamshid Olimovich", "+998939876543"])
    ws.append(["Aliyeva Nigora Rustam qizi", "+998971112233"])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# 1. SMS SOZLAMALARI MODAL (TELEFONNI ULASH)
@st.dialog("SMS Gateway Sozlamalari", width="large")
def sms_settings_modal():
    st.markdown(f"""
    Telefoningizdagi **`{sms_service.sender_number}`** SIM kartadan real SMS jo'natish uchun quyidagi sozlamalarni kiriting:
    """)
    
    sender_in = st.text_input("Yuboruvchi SIM karta telefon raqami:", value=sms_service.sender_number, placeholder="+998937078047")
    
    tab_infini, tab_local, tab_eskiz = st.tabs(["InfiniReach (Siz ulagan)", "Lokal IP Gateway", "Eskiz.uz API"])
    
    with tab_infini:
        st.markdown("""
        **Siz telefoningizni InfiniReach-ga muvaffaqiyatli uladingiz!**  
        Endi shunchaki brauzeringizdagi **InfiniReach** oynasidan API kalitini oling:
        1. Chap menyudan **`Integrations`** (yoki **`Settings`**) bo'limiga bosing.
        2. **API Key** (yoki *Create / Generate API Key*) ni ko'chirib oling (Copy).
        3. Quyidagi katakka joylang va **Saqlash** tugmasini bosing:
        """)
        
        inf_key = st.text_input("InfiniReach API Key / Token:", value=sms_service.gateway_token or "", type="password", placeholder="Masalan: ir_live_... yoki token...")
        
        if st.button("InfiniReach Sozlamasini Saqlash", icon=FLATICON["save"], key="btn_save_infini", type="primary", use_container_width=True):
            if inf_key.strip():
                sms_service.save_config(
                    gateway_type="infinireach",
                    url="https://api.infinireach.io",
                    token=inf_key.strip(),
                    sender_number=sender_in.strip() if sender_in.strip() else None
                )
                st.success("InfiniReach sozlamalari saqlandi! Endi bemalol SMS yuborishingiz mumkin.")
                st.rerun()
            else:
                st.error("Iltimos, InfiniReach API Key-ni kiriting!")

    with tab_local:
        st.markdown("""
        **Lokal Wi-Fi orqali to'g'ridan-to'g'ri ulash:**  
        Agar telefonda InfiniReach o'rniga lokal IP ko'rsatilgan bo'lsa (masalan `http://192.168.1.50:8080`), uni kiriting:
        """)
        local_url = st.text_input("Lokal Gateway URL:", value=sms_service.gateway_url if "192.168" in sms_service.gateway_url else "", placeholder="http://192.168.1.50:8080")
        local_user = st.text_input("Lokal Login (ixtiyoriy):", value=sms_service.gateway_login)
        local_pass = st.text_input("Lokal Parol (ixtiyoriy):", value=sms_service.gateway_password, type="password")
        
        if st.button("Lokal Sozlamalarni Saqlash", icon=FLATICON["save"], key="btn_save_local", type="primary", use_container_width=True):
            if local_url.strip():
                sms_service.save_config(
                    gateway_type="android",
                    url=local_url.strip(),
                    login=local_user.strip(),
                    password=local_pass.strip()
                )
                st.success("Lokal Gateway sozlamalari saqlandi!")
                st.rerun()
            else:
                st.error("Lokal URL manzilini kiriting!")

    with tab_eskiz:
        st.markdown("**Eskiz.uz servisi orqali jo'natish:**")
        esk_email = st.text_input("Eskiz Email / Login:", value=sms_service.gateway_login if "@" in sms_service.gateway_login else "")
        esk_pass = st.text_input("Eskiz Parol:", type="password")
        
        if st.button("Eskiz-ni Saqlash", icon=FLATICON["save"], key="btn_save_eskiz", type="primary", use_container_width=True):
            if esk_email.strip() and esk_pass.strip():
                sms_service.save_config(
                    gateway_type="eskiz",
                    url="https://notify.eskiz.uz",
                    login=esk_email.strip(),
                    password=esk_pass.strip()
                )
                st.success("Eskiz sozlamalari saqlandi!")
                st.rerun()
            else:
                st.error("Email va parolni to'ldiring!")


# 2. SMS YUBORISH MODAL
@st.dialog("SMS Xabarnoma Yuborish", width="large")
def send_sms_modal(recipients: list):
    # Holat ko'rsatkichi
    if not sms_service.is_configured():
        st.warning(f"""
        Diqqat: Real telefonga SMS yetib borishi uchun telefoningizdagi SIM karta ulangan bo'lishi kerak!  
        Hozirda kompyuteringiz `{sms_service.sender_number}` SIM kartasi solingan telefoningiz bilan ulanmagan.
        """)
        if st.button("Telefonni ulash (SMS Gateway sozlash)", icon=FLATICON["settings"], type="secondary", use_container_width=True):
            sms_settings_modal()
    else:
        banner_bg = "#064e3b" if is_dark else "#ecfdf5"
        banner_border = "#059669" if is_dark else "#a7f3d0"
        banner_text = "#6ee7b7" if is_dark else "#047857"
        st.markdown(f'<div style="background:{banner_bg}; border:1px solid {banner_border}; padding:8px 14px; border-radius:8px; color:{banner_text}; font-size:0.9rem; margin-bottom:12px;"><i class="fi fi-rr-badge-check" style="font-size:1.05rem; vertical-align:-1px; margin-right:4px;"></i> <b>SMS Gateway ulangan:</b> Xabarlar to\'g\'ridan-to\'g\'ri telefoningizdagi SIM kartadan yetkaziladi.</div>', unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background:{theme_card_bg}; border:1px solid {theme_border}; padding:12px 16px; border-radius:10px; margin-bottom:14px;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <span style="color:{theme_subtext}; font-size:0.9rem;">Yuboruvchi raqam:</span>&nbsp;
                <span class="sender-badge">{sms_service.sender_number}</span>
            </div>
            <div>
                <span style="color:{theme_subtext}; font-size:0.9rem;">Qabul qiluvchilar:</span>&nbsp;
                <span style="color:{theme_accent}; font-weight:700; font-size:1.1rem;">{len(recipients)} ta odam</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    default_sms = st.session_state.get("sms_draft_text", "Assalomu alaykum, {FullName}! Sizga eslatib o'tamiz...")
    sms_text = st.text_area(
        "Xabar matni:",
        value=default_sms,
        height=120,
        placeholder="Masalan: Assalomu alaykum, {FullName}! Xush kelibsiz..."
    )
    st.session_state["sms_draft_text"] = sms_text

    st.markdown(f"""
    <div style="font-size:0.86rem; color:{theme_subtext}; margin-top:-6px; margin-bottom:10px;">
        <i class="fi fi-rr-info" style="font-size:0.95rem; vertical-align:-1px; margin-right:4px; color:{theme_accent};"></i> <b>Tavsiya:</b> Matn ichida <code>{{FullName}}</code> deb yozsangiz, har bir odamga uning o'z ismi qo'yib yuboriladi.
    </div>
    """, unsafe_allow_html=True)

    clean_sender = sms_service.sender_number.replace("+", "").replace(" ", "").replace("-", "")
    self_recipients = [r for r in recipients if str(r.get("phone_number", "")).replace("+", "").replace(" ", "").replace("-", "") == clean_sender]
    if self_recipients:
        names_str = ", ".join(f"**{r.get('full_name')}**" for r in self_recipients[:3])
        st.warning(f"⚠️ **Diqqat:** {names_str} kontaktining raqami yuboruvchi SIM raqamingiz ({sms_service.sender_number}) bilan bir xil! Telefon o'z raqamiga SMS yuborganida xabar qabul qiluvchi shaxsga yetib bormaydi. Iltimos, kontaktning haqiqiy raqamini tekshiring.")

    if recipients and sms_text.strip():
        sample = recipients[0]
        preview_text = sms_service.format_message(sms_text, sample)
        st.markdown(f"##### <i class='fi fi-rr-comment-sms' style='color:{theme_accent}; margin-right:6px;'></i> Jonli ko'rinish (Namuna):", unsafe_allow_html=True)
        st.info(f"**Qabul qiluvchi:** {sample.get('full_name')} ({sample.get('phone_number')})\n\n**SMS matni:**\n> {preview_text}")

        char_count = len(preview_text)
        is_cyrillic = any(ord(c) > 127 for c in preview_text)
        limit = 70 if is_cyrillic else 160
        sms_parts = (char_count // limit) + 1
        st.caption(f"Belgilar: {char_count} ta | Taxminiy SMS qismi: {sms_parts} ta SMS")

    b_send, b_cancel = st.columns([1.2, 1])
    with b_send:
        if st.button(f"SMS Yuborish ({len(recipients)} ta)", type="primary", icon=FLATICON["send"], use_container_width=True):
            if not sms_text.strip():
                st.error("Iltimos, avval xabar matnini yozing!")
            else:
                status_box = st.empty()
                prog_bar = st.progress(0)
                sent_results = []
                total = len(recipients)
                
                for idx, r in enumerate(recipients):
                    cur_name = r.get("full_name") or r.get("phone_number")
                    status_box.markdown(f"⏳ **Yuborilmoqda ({idx + 1}/{total}):** `{cur_name}` ({r.get('phone_number')})...")
                    prog_bar.progress(int((idx / total) * 100))

                    personalized = sms_service.format_message(sms_text, r)
                    ok, stat = sms_service.send_single_sms(r.get("phone_number"), personalized, sender=sms_service.sender_number)
                    sent_results.append({
                        "name": r.get("full_name"),
                        "phone": r.get("phone_number"),
                        "message": personalized,
                        "success": ok,
                        "status_msg": stat
                    })
                    prog_bar.progress(int(((idx + 1) / total) * 100))
                    if idx < total - 1:
                        time.sleep(0.5)

                status_box.empty()
                prog_bar.progress(100)
                all_ok = all(item["success"] for item in sent_results)
                if all_ok:
                    st.success(f"Barcha {len(recipients)} ta odamga {sms_service.sender_number} raqamidan SMS jo'natildi!")
                else:
                    st.warning(f"{sum(1 for i in sent_results if i['success'])} ta xabar jo'natildi, ayrimlarida xatolik bo'ldi.")

                res_df = [
                    {
                        "FullName": item["name"],
                        "Telefon": item["phone"],
                        "Xabar": item["message"],
                        "Holati": "Yuborildi" if item["success"] else "Xato",
                        "Izoh": item["status_msg"]
                    }
                    for item in sent_results
                ]
                st.dataframe(res_df, use_container_width=True)

    with b_cancel:
        if st.button("Yopish", icon=FLATICON["close"], use_container_width=True):
            st.rerun()


# 2. EXCEL FAYLDAN YUKLASH MODAL
@st.dialog("Excel fayl orqali odamlar yuklash", width="large")
def upload_excel_modal():
    st.markdown("Excel fayldagi odamlar ro'yxatini to'g'ridan-to'g'ri Supabase bazasiga yuklang:")

    st.download_button(
        label="Namuna Excel shablonini yuklab olish",
        data=generate_template_excel(),
        file_name="namuna_shablon.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=FLATICON["download"]
    )

    st.markdown("---")
    uploaded_file = st.file_uploader(
        "Excel faylni tanlang (.xlsx yoki .csv):",
        type=["xlsx", "xls", "csv"],
        help="Faylda 2 ta ustun bo'lishi kifoya: FullName va PhoneNumberi"
    )

    if uploaded_file is not None:
        try:
            with st.spinner("Excel fayl o'qilmoqda va tahlil qilinmoqda, iltimos kuting..."):
                rows_data = []
                headers = []

                if uploaded_file.name.endswith(".csv"):
                    content = uploaded_file.getvalue().decode('utf-8-sig', errors='replace')
                    reader = csv.reader(io.StringIO(content))
                    for r in reader:
                        if any(c.strip() for c in r):
                            if not headers:
                                headers = [str(h).strip() if h else f"Ustun_{i+1}" for i, h in enumerate(r)]
                            else:
                                rows_data.append(r)
                else:
                    wb_up = load_workbook(uploaded_file, read_only=True, data_only=True)
                    ws_up = wb_up.active
                    for r in ws_up.iter_rows(values_only=True):
                        if any(v is not None and str(v).strip() != "" for v in r):
                            if not headers:
                                headers = [str(v).strip() if v is not None and str(v).strip() != "" else f"Ustun_{i+1}" for i, v in enumerate(r)]
                            else:
                                rows_data.append(r)
                    wb_up.close()

            if not headers:
                st.error("Fayl ichida ma'lumotlar topilmadi.")
                return

            st.markdown(f"""
            <div style="background:{'rgba(56, 189, 248, 0.1)' if is_dark else '#f0f9ff'}; border:1px solid {'rgba(56, 189, 248, 0.3)' if is_dark else '#bae6fd'}; border-radius:8px; padding:10px 14px; margin-bottom:12px;">
                <span style="font-weight:600; color:{theme_accent}; font-size:1.05rem;">Faylda topilgan qatorlar: {len(rows_data):,} ta</span>
            </div>
            """, unsafe_allow_html=True)

            # Ustunlarni avtomatik aniqlash (O'zbek, Rus va Ingliz tillarida)
            name_aliases = [
                "fullname", "full_name", "f.i.sh", "f.i.o", "fio", "фио", "имя", "фамилия",
                "ism", "nomi", "talaba", "o'quvchi", "abonent", "student", "name", "client"
            ]
            phone_aliases = [
                "phonenumberi", "phone_number", "telefon", "tel", "raqam", "nomer", "телефон",
                "номер", "тел", "номер телефона", "контакты", "mobil", "mobile", "phone"
            ]

            def find_match(hdrs, aliases):
                for idx, h in enumerate(hdrs):
                    clean_h = str(h).strip().lower().replace(" ", "").replace("_", "").replace(".", "")
                    for a in aliases:
                        clean_a = a.lower().replace(" ", "").replace("_", "").replace(".", "")
                        if clean_h == clean_a or clean_a in clean_h:
                            return idx
                return None

            auto_name = find_match(headers, name_aliases)
            auto_phone = find_match(headers, phone_aliases)

            st.markdown("##### Ustunlarni tasdiqlang:")
            c_sel1, c_sel2 = st.columns(2)
            with c_sel1:
                name_idx = st.selectbox(
                    "Ism (FullName) ustuni:",
                    options=list(range(len(headers))),
                    format_func=lambda i: f"{headers[i]} (Ustun #{i+1})",
                    index=auto_name if auto_name is not None else 0
                )
            with c_sel2:
                default_p = auto_phone if auto_phone is not None else (1 if len(headers) > 1 else 0)
                phone_idx = st.selectbox(
                    "Telefon raqam (PhoneNumberi) ustuni:",
                    options=list(range(len(headers))),
                    format_func=lambda i: f"{headers[i]} (Ustun #{i+1})",
                    index=default_p
                )

            # Dastlabki 5 ta qator namunasi
            preview_data = []
            for r in rows_data[:5]:
                n_val = r[name_idx] if name_idx < len(r) and r[name_idx] is not None else ""
                p_val = r[phone_idx] if phone_idx < len(r) and r[phone_idx] is not None else ""
                preview_data.append({
                    "FullName (Ism)": str(n_val).strip(),
                    "PhoneNumberi (Telefon)": str(p_val).strip()
                })

            st.dataframe(preview_data, use_container_width=True)
            st.info(f"💡 Fayldan jami **{len(rows_data):,} ta** odam topildi. Ularni Supabase bazasiga qo'shish uchun quyidagi ko'k tugmani bosing:")

            if st.button("Bazaga Yuklashni Tasdiqlash", icon=FLATICON["save"], type="primary", use_container_width=True):
                with st.spinner("Ma'lumotlar bazaga yuklanmoqda..."):
                    batch_records = []
                    for r in rows_data:
                        raw_n = r[name_idx] if name_idx < len(r) and r[name_idx] is not None else ""
                        raw_p = r[phone_idx] if phone_idx < len(r) and r[phone_idx] is not None else ""
                        n_str = str(raw_n).strip()
                        p_str = str(raw_p).strip()
                        if n_str and p_str and n_str != "None" and p_str != "None":
                            batch_records.append({
                                "full_name": n_str,
                                "phone_number": p_str
                            })

                    if batch_records:
                        count, msg = service.add_people_batch(batch_records)
                        if count > 0:
                            refresh_cache()
                            st.success(f"{count:,} ta odam muvaffaqiyatli bazaga qo'shildi!")
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.warning("Yuklash uchun yaroqli ma'lumotlar topilmadi.")

        except Exception as e:
            st.error(f"Faylni o'qishda xatolik yuz berdi: {e}")


# 3. YANGI ODAM QO'SHISH MODAL
@st.dialog("Yangi odam qo'shish", width="medium")
def add_person_modal():
    st.markdown("Odam ma'lumotlarini kiriting:")
    with st.form("modal_add_form", clear_on_submit=False):
        full_name = st.text_input("To'liq Ism (FullName) *", placeholder="Masalan: Karimov Sardor")
        phone_number = st.text_input("Telefon Raqami (PhoneNumberi) *", placeholder="+998901234567")

        b1, b2 = st.columns([1, 1])
        with b1:
            save_btn = st.form_submit_button("Bazaga Saqlash", icon=FLATICON["save"], use_container_width=True, type="primary")
        with b2:
            cancel_btn = st.form_submit_button("Bekor qilish", icon=FLATICON["close"], use_container_width=True)

        if save_btn:
            if not full_name.strip():
                st.error("To'liq ism (FullName) kiritilishi shart!")
            elif not phone_number.strip():
                st.error("Telefon raqami (PhoneNumberi) kiritilishi shart!")
            else:
                ok, msg = service.add_person({
                    "full_name": full_name.strip(),
                    "phone_number": phone_number.strip()
                })
                if ok:
                    refresh_cache()
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
        elif cancel_btn:
            st.rerun()


# 4. TAHRIRLASH (UPDATE) MODAL
@st.dialog("Odam ma'lumotlarini yangilash", width="medium")
def edit_person_modal(person: dict):
    p_id = person.get("id")
    st.markdown(f"**Tahrirlanmoqda:** `{person.get('full_name')}`")
    
    with st.form(key=f"modal_edit_form_{p_id}", clear_on_submit=False):
        u_full_name = st.text_input("To'liq Ism (FullName) *", value=person.get("full_name", ""))
        u_phone = st.text_input("Telefon Raqami (PhoneNumberi) *", value=person.get("phone_number", ""))

        b1, b2 = st.columns([1, 1])
        with b1:
            update_btn = st.form_submit_button("O'zgarishlarni Saqlash", icon=FLATICON["save"], use_container_width=True, type="primary")
        with b2:
            cancel_btn = st.form_submit_button("Bekor qilish", icon=FLATICON["close"], use_container_width=True)

        if update_btn:
            if not u_full_name.strip():
                st.error("To'liq ism kiritilishi shart!")
            elif not u_phone.strip():
                st.error("Telefon raqami kiritilishi shart!")
            else:
                ok, msg = service.update_person(p_id, {
                    "full_name": u_full_name.strip(),
                    "phone_number": u_phone.strip()
                })
                if ok:
                    refresh_cache()
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
        elif cancel_btn:
            st.rerun()


# 5. O'CHIRISH (DELETE) MODAL
@st.dialog("Odamni o'chirishni tasdiqlash", width="small")
def delete_person_modal(person: dict):
    p_id = person.get("id")
    st.warning(f"Rostdan ham **{person.get('full_name')}** ({person.get('phone_number')}) ni bazadan o'chirmoqchimisiz?")
    st.markdown("<i class='fi fi-rr-exclamation' style='margin-right:4px;'></i> *Bu amalni ortga qaytarib bo'lmaydi.*", unsafe_allow_html=True)
    
    c_yes, c_no = st.columns(2)
    with c_yes:
        if st.button("Ha, o'chirilsin", icon=FLATICON["delete"], type="primary", use_container_width=True, key=f"conf_del_{p_id}"):
            ok, msg = service.delete_person(p_id)
            if ok:
                st.session_state.selected_ids.discard(str(p_id))
                refresh_cache()
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)
    with c_no:
        if st.button("Bekor qilish", icon=FLATICON["close"], use_container_width=True, key=f"canc_del_{p_id}"):
            st.rerun()


# 6. TANLANGANLARNI BIRGALIKDA O'CHIRISH (BATCH DELETE) MODAL
@st.dialog("Tanlangan odamlarni o'chirish", width="small")
def delete_selected_people_modal():
    selected_ids = list(st.session_state.selected_ids)
    count = len(selected_ids)
    if count == 0:
        st.info("Hech qanday odam belgilanmagan.")
        if st.button("Yopish", icon=FLATICON["close"], use_container_width=True):
            st.rerun()
        return

    st.warning(f"Rostdan ham tanlangan **{count} ta** odamni bazadan butunlay o'chirib tashlamoqchimisiz?")
    st.markdown("<i class='fi fi-rr-exclamation' style='margin-right:4px;'></i> *Bu amalni ortga qaytarib bo'lmaydi. Barcha belgilangan odamlar bazadan o'chiriladi.*", unsafe_allow_html=True)
    
    c_yes, c_no = st.columns(2)
    with c_yes:
        if st.button(f"Ha, {count} tasini o'chirish", icon=FLATICON["delete"], type="primary", use_container_width=True, key="conf_batch_del"):
            with st.spinner("O'chirilmoqda..."):
                deleted_count, msg = service.delete_people_batch(selected_ids)
                if deleted_count > 0:
                    for pid in selected_ids:
                        st.session_state.pop(f"chk_{pid}", None)
                    st.session_state.selected_ids = set()
                    st.session_state.page_number = 1
                    refresh_cache()
                    st.success(f"{deleted_count} ta odam muvaffaqiyatli o'chirildi!")
                    time.sleep(0.3)
                    st.rerun()
                else:
                    st.error(msg)
    with c_no:
        if st.button("Bekor qilish", icon=FLATICON["close"], use_container_width=True, key="canc_batch_del"):
            st.rerun()


# ================= ASOSIY OYNA (HEADER) =================

with st.container():
    col_brand, col_actions = st.columns([5.2, 2.2], vertical_alignment="center")
    with col_brand:
        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:12px; padding:2px 0;">
            <div style="width:42px; height:42px; min-width:42px; border-radius:10px; background:linear-gradient(135deg, {theme_accent}, #2563eb); display:flex; align-items:center; justify-content:center; color:#ffffff; font-size:1.35rem; box-shadow:0 0 14px {theme_accent}66;">
                <i class="fi fi-rr-bolt"></i>
            </div>
            <div>
                <div style="font-size:1.35rem; font-weight:700; color:{theme_text}; line-height:1.2; letter-spacing:-0.3px;">
                    SMS Gateway Platform
                </div>
                <div style="display:flex; align-items:center; gap:8px; margin-top:2px;">
                    <span style="font-size:0.82rem; color:{theme_subtext}; font-family:monospace; font-weight:600;">
                        {sms_service.sender_number}
                    </span>
                    <span style="display:inline-flex; align-items:center; gap:5px; background:{'rgba(34,197,94,0.15)' if is_dark else '#dcfce7'}; color:{'#4ade80' if is_dark else '#15803d'}; border:1px solid {'rgba(34,197,94,0.35)' if is_dark else '#bbf7d0'}; border-radius:12px; padding:1px 8px; font-size:0.72rem; font-weight:600;">
                        <span style="width:6px; height:6px; border-radius:50%; background:#22c55e; box-shadow:0 0 6px #22c55e; display:inline-block;"></span>
                        Online
                    </span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_actions:
        btn_c1, btn_c2 = st.columns([1, 1], vertical_alignment="center")
        with btn_c1:
            if st.button("Sozlamalar", icon=FLATICON["settings"], key="btn_header_settings", use_container_width=True, help="SMS Gateway sozlamalarini ko'rish"):
                sms_settings_modal()
        with btn_c2:
            if is_dark:
                if st.button("Kunduzgi", icon=FLATICON["sun"], key="btn_theme_toggle", use_container_width=True, help="Kunduzgi (Light) rejimga o'tish"):
                    st.session_state.theme = "light"
                    st.rerun()
            else:
                if st.button("Tungi", icon=FLATICON["moon"], key="btn_theme_toggle", use_container_width=True, help="Tungi (Dark) rejimga o'tish"):
                    st.session_state.theme = "dark"
                    st.rerun()

st.markdown(f"<hr style='margin: 10px 0 16px 0; border: none; border-top: 1px solid {theme_hr};'/>", unsafe_allow_html=True)

tab_people, tab_logs = st.tabs(["Odamlar Ro'yxati va SMS", "Message Logs (SMS Tarixi)"])

with tab_people:
    # Sarlavha
    st.markdown(f"""
    <div class="title-box">
        <i class="fi fi-rr-users-alt title-icon"></i>
        <span class="title-text">Odamlar Ro'yxati va SMS Xabarnoma</span>
    </div>
    <div class="sub-text">Odamlarni belgilab, <b>{sms_service.sender_number}</b> raqamidan shaxsiy SMS yuborish tizimi</div>
    """, unsafe_allow_html=True)

    # Ma'lumotlarni tezkor xotiradan (RAM) olamiz (0ms kechikish)
    all_people = load_people_data()

    # Toolbar (Harakatlar tugmalari)
    t1, t2, t3 = st.columns([1, 1, 1])
    with t1:
        if st.button("Yangi qo'shish", icon=FLATICON["user_add"], type="primary", use_container_width=True):
            add_person_modal()
    with t2:
        if st.button("Excel yuklash", icon=FLATICON["upload"], use_container_width=True):
            upload_excel_modal()
    with t3:
        excel_bytes = export_to_excel_bytes(all_people)
        st.download_button(
            label="Excelga saqlash",
            data=excel_bytes,
            file_name="odamlar_royxati.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=FLATICON["download"],
            use_container_width=True
        )

    # Qidiruv maydoni (Xotirada bir zumda filtrlaydi)
    search_text = st.text_input(
        "Qidiruv (Ism yoki Telefon raqami):",
        placeholder="Ism yoki telefon raqami bo'yicha qidirish..."
    )

    if search_text.strip():
        sq = search_text.strip().lower()
        people = [
            p for p in all_people
            if sq in str(p.get("full_name", "")).lower() or sq in str(p.get("phone_number", "")).lower()
        ]
    else:
        people = all_people

    st.markdown(f"<hr style='margin: 8px 0 14px 0; border: none; border-top: 1px solid {theme_hr};'/>", unsafe_allow_html=True)

    # ================= BELGILASH VA AMALLAR (SMS / O'CHIRISH) STRIP =================
    col_sel_info, col_sel_actions = st.columns([2.6, 4.4], vertical_alignment="center")

    with col_sel_info:
        sel_count = len(st.session_state.selected_ids)
        if sel_count > 0:
            st.markdown(f"<i class='fi fi-rr-check' style='color:{theme_accent}; margin-right:4px;'></i>**Tanlangan:** <span style='color:{theme_accent}; font-weight:700; font-size:1.15rem;'>{sel_count}</span> ta odam (Jami: {len(people)} ta)", unsafe_allow_html=True)
        else:
            st.markdown(f"Jami: **{len(people)}** ta odam *(SMS yuborish yoki o'chirish uchun quyidagilarni belgilang)*")

    with col_sel_actions:
        a1, a2, a3, a4 = st.columns([1.3, 1.2, 1.8, 1.6])
        with a1:
            if st.button("Hammasini tanlash", icon=FLATICON["select_all"], use_container_width=True, help="Barcha qatorlarni belgilash"):
                st.session_state.selected_ids = set(str(p.get("id")) for p in people)
                for p in people:
                    st.session_state[f"chk_{p.get('id')}"] = True
                st.rerun()
        with a2:
            if st.button("Bekor qilish", icon=FLATICON["deselect"], use_container_width=True, help="Belgilashni bekor qilish"):
                st.session_state.selected_ids = set()
                for p in people:
                    st.session_state[f"chk_{p.get('id')}"] = False
                st.rerun()
        with a3:
            selected_recipients = [p for p in all_people if str(p.get("id")) in st.session_state.selected_ids]
            btn_sms_label = f"SMS Yuborish ({len(selected_recipients)} ta)" if selected_recipients else "SMS Yuborish"
            
            if st.button(btn_sms_label, icon=FLATICON["send"], type="primary", use_container_width=True, disabled=len(selected_recipients) == 0):
                send_sms_modal(selected_recipients)
        with a4:
            btn_del_label = f"O'chirish ({sel_count} ta)" if sel_count > 0 else "O'chirish"
            if st.button(btn_del_label, icon=FLATICON["delete"], use_container_width=True, disabled=sel_count == 0, key="btn_batch_delete_trigger", help="Tanlangan barcha odamlarni o'chirish"):
                delete_selected_people_modal()

    # ================= JADVAL VA PAGINATSIYA =================
    if people:
        # Paginatsiya hisob-kitobi
        total_pages = max(1, (len(people) + PAGE_SIZE - 1) // PAGE_SIZE)
        if st.session_state.page_number > total_pages:
            st.session_state.page_number = total_pages

        start_idx = (st.session_state.page_number - 1) * PAGE_SIZE
        end_idx = min(start_idx + PAGE_SIZE, len(people))
        page_people = people[start_idx:end_idx]

        col_widths = [0.5, 4.2, 3.2, 1.6]

        # Jadval Sarlavhasi
        h = st.columns(col_widths, vertical_alignment="center")
        h[0].markdown("")
        h[1].markdown(f"<i class='fi fi-rr-user' style='color:{theme_accent}; margin-right:6px;'></i>**FullName (To'liq Ism)**", unsafe_allow_html=True)
        h[2].markdown(f"<i class='fi fi-rr-phone-call' style='color:{theme_accent}; margin-right:6px;'></i>**PhoneNumberi (Telefon Raqami)**", unsafe_allow_html=True)
        h[3].markdown(f"<i class='fi fi-rr-settings-sliders' style='color:{theme_accent}; margin-right:6px;'></i>**Amallar**", unsafe_allow_html=True)

        st.markdown(f"<hr style='margin: 4px 0 8px 0; border: none; border-top: 1px solid {theme_hr};'/>", unsafe_allow_html=True)

        # Faqat shu sahifadagi 15 ta odamni chiqaramiz (O'ta tez ishlaydi)
        for idx, p in enumerate(page_people):
            p_id = str(p.get("id"))
            is_checked = p_id in st.session_state.selected_ids

            r = st.columns(col_widths, vertical_alignment="center")
            
            # 1. Belgilash Checkbox (on_change callback bilan, qotishsiz!)
            r[0].checkbox(
                "Tanlash",
                value=is_checked,
                key=f"chk_{p_id}",
                label_visibility="collapsed",
                on_change=on_check_toggle,
                args=(p_id,)
            )

            # 2. Ism
            r[1].markdown(f"**{p.get('full_name', '')}**")
            
            # 3. Telefon raqam
            r[2].markdown(f"<span class='phone-text'>{p.get('phone_number', '')}</span>", unsafe_allow_html=True)

            # 4. Amallar (SMS, Tahrirlash, O'chirish)
            act_col1, act_col2, act_col3 = r[3].columns(3)
            with act_col1:
                if st.button("", icon=FLATICON["chat"], key=f"row_sms_{p_id}_{start_idx+idx}", help=f"{p.get('full_name')} ga SMS yuborish"):
                    send_sms_modal([p])
            with act_col2:
                if st.button("", icon=FLATICON["edit"], key=f"row_edit_{p_id}_{start_idx+idx}", help="Tahrirlash"):
                    edit_person_modal(p)
            with act_col3:
                if st.button("", icon=FLATICON["delete"], key=f"row_del_{p_id}_{start_idx+idx}", help="O'chirish"):
                    delete_person_modal(p)

            st.markdown(f"<hr style='margin: 3px 0 5px 0; border: none; border-top: 1px solid {theme_row_hr};'/>", unsafe_allow_html=True)

        # Agar 1 dan ko'p sahifa bo'lsa, pastda qulay paginatsiya chiqadi
        if total_pages > 1:
            p_prev, p_info, p_next = st.columns([1, 2, 1], vertical_alignment="center")
            with p_prev:
                if st.button("Oldingi", icon=FLATICON["arrow_left"], disabled=st.session_state.page_number <= 1, use_container_width=True):
                    st.session_state.page_number -= 1
                    st.rerun()
            with p_info:
                st.markdown(f"<div style='text-align:center; color:{theme_subtext};'>Sahifa <b>{st.session_state.page_number}</b> / {total_pages} (Jami {len(people)} ta odam)</div>", unsafe_allow_html=True)
            with p_next:
                if st.button("Keyingi", icon=FLATICON["arrow_right"], icon_position="right", disabled=st.session_state.page_number >= total_pages, use_container_width=True):
                    st.session_state.page_number += 1
                    st.rerun()

    else:
        st.info("Hech qanday odam topilmadi. Yuqoridagi 'Yangi qo'shish' yoki 'Excel yuklash' tugmasi orqali odamlarni kiritishingiz mumkin.")


# ================= MESSAGE LOGS SAHIFASI =================
with tab_logs:
    col_l1, col_l2 = st.columns([4, 1.2], vertical_alignment="center")
    with col_l1:
        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:10px; margin-top:6px; margin-bottom:4px;">
            <i class="fi fi-rr-document" style="font-size:1.8rem; color:{theme_accent};"></i>
            <span style="font-size:2rem; font-weight:700; color:{theme_text};">Message Logs</span>
        </div>
        """, unsafe_allow_html=True)
    with col_l2:
        if st.button("Yangilash", icon=FLATICON["refresh"], key="btn_refresh_logs", use_container_width=True):
            load_message_logs(force_reload=True)
            st.rerun()

    logs = load_message_logs()

    # Filtrlar paneli
    with st.container(border=True):
        f1, f2, f3, f4, f5, f6 = st.columns([1.1, 1.1, 1.1, 2.3, 1.1, 1.3])
        with f1:
            flt_channel = st.selectbox("Channel", ["All", "SMS", "WhatsApp"], key="flt_chan")
        with f2:
            flt_direction = st.selectbox("Direction", ["All", "Outbound", "Inbound"], key="flt_dir")
        with f3:
            flt_status = st.selectbox("Status", ["All", "delivered", "queued", "failed", "acked"], key="flt_stat")
        with f4:
            flt_search = st.text_input("Message contains", placeholder="Search in message body...", key="flt_body")
        with f5:
            flt_source = st.selectbox("Source", ["All", "API"], key="flt_src")
        with f6:
            devices_list = ["All Devices"]
            if sms_service.sender_number not in devices_list:
                devices_list.append(sms_service.sender_number)
            for m in logs:
                ef = m.get("e164From")
                if ef and ef not in devices_list:
                    devices_list.append(ef)
            flt_device = st.selectbox("Device", devices_list, key="flt_dev")

    # Filtrlash
    filtered = []
    for m in logs:
        if flt_channel != "All" and m.get("channel", "").lower() != flt_channel.lower():
            continue
        if flt_direction == "Outbound" and m.get("direction") != "out":
            continue
        if flt_direction == "Inbound" and m.get("direction") != "in":
            continue
        if flt_status != "All" and m.get("status", "").lower() != flt_status.lower():
            continue
        if flt_search.strip():
            sq = flt_search.strip().lower()
            body_txt = str(m.get("body", "")).lower()
            to_num = str(m.get("e164To", "")).lower()
            mid = str(m.get("_id", "")).lower()
            if sq not in body_txt and sq not in to_num and sq not in mid:
                continue
        if flt_device != "All Devices" and m.get("e164From") != flt_device:
            continue
        filtered.append(m)

    total_logs = len(filtered)
    st.markdown(f"<div style='color:{theme_subtext}; font-size:0.92rem; margin-bottom:14px; font-weight:500;'>Showing 1–{min(total_logs, 25)} of {total_logs} messages</div>", unsafe_allow_html=True)

    if filtered:
        LOGS_PER_PAGE = 25
        if "logs_page" not in st.session_state:
            st.session_state.logs_page = 1

        total_log_pages = max(1, (total_logs + LOGS_PER_PAGE - 1) // LOGS_PER_PAGE)
        if st.session_state.logs_page > total_log_pages:
            st.session_state.logs_page = total_log_pages

        l_start = (st.session_state.logs_page - 1) * LOGS_PER_PAGE
        l_end = min(l_start + LOGS_PER_PAGE, total_logs)
        page_logs = filtered[l_start:l_end]

        for m in page_logs:
            channel = m.get("channel", "sms").upper()
            direction = "Outbound" if m.get("direction") == "out" else "Inbound"
            status = m.get("status", "unknown").lower()
            msg_id = m.get("_id", "")
            from_num = m.get("e164From", sms_service.sender_number)
            to_num = m.get("e164To", "")
            body = m.get("body", "")
            ts_str = format_log_timestamp(m.get("createdAt"))

            body_len = len(body)
            is_cyrillic = any(ord(c) > 127 for c in body)
            limit = 70 if is_cyrillic else 160
            segments = (body_len // limit) + 1 if body_len > 0 else 1

            if is_dark:
                card_bg = "#1e293b"
                card_border = "#334155"
                card_shadow = "none"
                badge_chan_bg = "rgba(139, 92, 246, 0.15)"
                badge_chan_color = "#c4b5fd"
                badge_chan_border = "rgba(139, 92, 246, 0.3)"
                badge_dir_bg = "rgba(168, 85, 247, 0.15)"
                badge_dir_color = "#d8b4fe"
                badge_dir_border = "rgba(168, 85, 247, 0.3)"
                badge_api_bg = "rgba(6, 182, 212, 0.15)"
                badge_api_color = "#67e8f9"
                badge_api_border = "rgba(6, 182, 212, 0.3)"
                from_color = "#f8fafc"
                to_color = "#38bdf8"
                meta_color = "#94a3b8"
                msg_id_color = "#64748b"
                body_bg = "rgba(15, 23, 42, 0.35)"
                body_color = "#e2e8f0"
                body_border = "#38bdf8"

                if status == "delivered":
                    status_bg = "rgba(34, 197, 94, 0.15)"
                    status_color = "#86efac"
                    status_border = "rgba(34, 197, 94, 0.35)"
                elif status == "sent":
                    status_bg = "rgba(14, 165, 233, 0.15)"
                    status_color = "#38bdf8"
                    status_border = "rgba(14, 165, 233, 0.35)"
                elif status == "queued":
                    status_bg = "rgba(234, 179, 8, 0.15)"
                    status_color = "#fde047"
                    status_border = "rgba(234, 179, 8, 0.35)"
                elif status == "failed":
                    status_bg = "rgba(239, 68, 68, 0.15)"
                    status_color = "#fca5a5"
                    status_border = "rgba(239, 68, 68, 0.35)"
                elif status == "acked":
                    status_bg = "rgba(59, 130, 246, 0.15)"
                    status_color = "#93c5fd"
                    status_border = "rgba(59, 130, 246, 0.35)"
                else:
                    status_bg = "rgba(148, 163, 184, 0.15)"
                    status_color = "#cbd5e1"
                    status_border = "rgba(148, 163, 184, 0.35)"
            else:
                card_bg = "#ffffff"
                card_border = "#e2e8f0"
                card_shadow = "0 2px 8px rgba(0, 0, 0, 0.05)"
                badge_chan_bg = "#f3e8ff"
                badge_chan_color = "#7c3aed"
                badge_chan_border = "#e9d5ff"
                badge_dir_bg = "#f5f3ff"
                badge_dir_color = "#6d28d9"
                badge_dir_border = "#ddd6fe"
                badge_api_bg = "#e0f2fe"
                badge_api_color = "#0369a1"
                badge_api_border = "#bae6fd"
                from_color = "#0f172a"
                to_color = "#0284c7"
                meta_color = "#64748b"
                msg_id_color = "#94a3b8"
                body_bg = "#f8fafc"
                body_color = "#1e293b"
                body_border = "#0284c7"

                if status == "delivered":
                    status_bg = "#dcfce7"
                    status_color = "#15803d"
                    status_border = "#bbf7d0"
                elif status == "sent":
                    status_bg = "#e0f2fe"
                    status_color = "#0284c7"
                    status_border = "#bae6fd"
                elif status == "queued":
                    status_bg = "#fef9c3"
                    status_color = "#a16207"
                    status_border = "#fef08a"
                elif status == "failed":
                    status_bg = "#fee2e2"
                    status_color = "#b91c1c"
                    status_border = "#fecaca"
                elif status == "acked":
                    status_bg = "#dbeafe"
                    status_color = "#1d4ed8"
                    status_border = "#bfdbfe"
                else:
                    status_bg = "#f1f5f9"
                    status_color = "#475569"
                    status_border = "#e2e8f0"

            error_html = ""
            if status == "failed":
                err_msg = m.get("error") or m.get("failedReason") or "Unexpected error while queuing SMS: uid 1010002 does not have android.permission.SEND_SMS."
                err_code = m.get("errorCode") or "SEND_FAILED"
                if is_dark:
                    error_html = f'<div style="background: rgba(239, 68, 68, 0.12); border: 1px solid rgba(239, 68, 68, 0.35); border-radius: 8px; padding: 10px 14px; margin-top: 10px; color: #f87171;"><div style="font-weight: 600; font-size: 0.85rem; margin-bottom: 2px;">Error:</div><div style="font-size: 0.85rem;">{err_msg}</div><div style="font-size: 0.8rem; margin-top: 2px; color: #fca5a5;">Code: {err_code}</div></div>'
                else:
                    error_html = f'<div style="background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; padding: 10px 14px; margin-top: 10px; color: #dc2626;"><div style="font-weight: 600; font-size: 0.85rem; margin-bottom: 2px;">Error:</div><div style="font-size: 0.85rem;">{err_msg}</div><div style="font-size: 0.8rem; margin-top: 2px; color: #ef4444;">Code: {err_code}</div></div>'

            card_html = f'''<div style="background: {card_bg}; border: 1px solid {card_border}; box-shadow: {card_shadow}; border-radius: 10px; padding: 16px 20px; margin-bottom: 12px; transition: all 0.2s;">
<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
<div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
<span style="background: {badge_chan_bg}; color: {badge_chan_color}; border: 1px solid {badge_chan_border}; border-radius: 6px; padding: 2px 9px; font-size: 0.78rem; font-weight: 600; display: inline-flex; align-items: center; gap: 5px;"><i class="fi fi-rr-comment-sms" style="font-size:0.85rem;"></i> {channel}</span>
<span style="background: {badge_dir_bg}; color: {badge_dir_color}; border: 1px solid {badge_dir_border}; border-radius: 6px; padding: 2px 9px; font-size: 0.78rem; font-weight: 600; display: inline-flex; align-items: center; gap: 5px;"><i class="fi fi-rr-paper-plane" style="font-size:0.85rem;"></i> {direction}</span>
<span style="background: {badge_api_bg}; color: {badge_api_color}; border: 1px solid {badge_api_border}; border-radius: 6px; padding: 2px 9px; font-size: 0.78rem; font-weight: 600;">API</span>
<span style="background: {status_bg}; color: {status_color}; border: 1px solid {status_border}; border-radius: 6px; padding: 2px 9px; font-size: 0.78rem; font-weight: 600;">{status}</span>
<span style="color: {meta_color}; font-size: 0.82rem; margin-left: 2px;">{segments} segment</span>
</div>
<div style="color: {meta_color}; font-size: 0.85rem; font-family: monospace;">{ts_str}</div>
</div>
<div style="font-size: 0.95rem; margin-bottom: 2px;">
<span style="color: {meta_color};">From:</span> <b style="color: {from_color}; font-family: monospace;">{from_num}</b> <span style="color: {meta_color}; margin: 0 4px;">&rarr;</span> <span style="color: {meta_color};">To:</span> <b style="color: {to_color}; font-family: monospace;">{to_num}</b>
</div>
<div style="color: {msg_id_color}; font-size: 0.82rem; font-family: monospace; margin-bottom: 8px;">Message ID: {msg_id}</div>
<div style="color: {body_color}; font-size: 0.95rem; line-height: 1.5; white-space: pre-wrap; background: {body_bg}; padding: 8px 12px; border-radius: 6px; border-left: 3px solid {body_border};">{body}</div>
{error_html}
</div>'''
            st.html(card_html)

        if total_log_pages > 1:
            lp_prev, lp_info, lp_next = st.columns([1, 2, 1], vertical_alignment="center")
            with lp_prev:
                if st.button("Oldingi sahifa", icon=FLATICON["arrow_left"], key="lp_btn_prev", disabled=st.session_state.logs_page <= 1, use_container_width=True):
                    st.session_state.logs_page -= 1
                    st.rerun()
            with lp_info:
                st.markdown(f"<div style='text-align:center; color:{theme_subtext};'>Sahifa <b>{st.session_state.logs_page}</b> / {total_log_pages}</div>", unsafe_allow_html=True)
            with lp_next:
                if st.button("Keyingi sahifa", icon=FLATICON["arrow_right"], icon_position="right", key="lp_btn_next", disabled=st.session_state.logs_page >= total_log_pages, use_container_width=True):
                    st.session_state.logs_page += 1
                    st.rerun()
    else:
        st.info("Filtr bo'yicha hech qanday xabar topilmadi.")
