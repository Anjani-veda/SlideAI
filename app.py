import streamlit as st
import os
import time
import random
import uuid
from excel import process_file, THEMES
from excel_database import register_user, verify_user, create_users_table, create_reports_table, save_otp, verify_otp, update_password, get_user_phone

try:
    st.set_page_config(page_title="Excel AI Analyzer", layout="wide")
except Exception:
    # This happens if running with 'python app.py' instead of 'streamlit run app.py'
    print("CRITICAL: This application must be run using Streamlit.")
    print("Please run: streamlit run app.py")
    import sys
    sys.exit(1)

# Initialize database tables (users, reports) if not existing
if 'db_initialized' not in st.session_state:
    try:
        create_users_table()
        create_reports_table()
        st.session_state.db_initialized = True
    except Exception as e:
        st.error(f"❌ Database Connection Error: {e}")
        st.info("Please ensure your MySQL server is running with the credentials specified in excel_database.py")
        st.stop()


# ---------------- SESSION ---------------- #
if "users" not in st.session_state:
    st.session_state.users = {}

if "auth" not in st.session_state:
    st.session_state.auth = None

if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = None

if "last_uploaded" not in st.session_state:
    st.session_state.last_uploaded = None

if "forgot_step" not in st.session_state:
    st.session_state.forgot_step = 1

if "forgot_user" not in st.session_state:
    st.session_state.forgot_user = None


# ---------------- AUTH FUNCTIONS ---------------- #
def auth_page():
    st.markdown("""
        <style>
        div[data-testid="stFormSubmitButton"] > button {
            background-color: #217346;
            color: white;
            border-color: #1e6b40;
        }
        div[data-testid="stFormSubmitButton"] > button:hover {
            background-color: #1e6b40;
            border-color: #1a5c37;
            color: white;
        }
        .stTabs [data-baseweb="tab-highlight"] {
            background-color: #217346;
        }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("<h2 style='text-align: center; color: #217346;'>📊 Excel AI Analyzer</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: gray;'>Please login or create an account</p>", unsafe_allow_html=True)
        st.write("")

        tab1, tab2 = st.tabs(["Login", "Sign Up"])

        with tab1:
            if st.session_state.forgot_step == 1:
                with st.form("login_form"):
                    user = st.text_input("Username", key="login_user")
                    pwd = st.text_input("Password", type="password", key="login_pwd")
                    if st.form_submit_button("Login", use_container_width=True):
                        if verify_user(user, pwd):
                            st.session_state.auth = user
                            st.rerun()
                        else:
                            st.error("Invalid credentials")
                
                if st.button("Forgot Password?", key="btn_forgot"):
                    st.session_state.forgot_step = 2
                    st.rerun()

            elif st.session_state.forgot_step == 2:
                st.info("Enter your username to receive an OTP on your registered phone.")
                user_to_reset = st.text_input("Username", key="reset_user")
                if st.button("Send OTP", key="btn_send_otp"):
                    # First check if user exists and get their phone
                    phone = get_user_phone(user_to_reset)
                    if phone:
                        otp = str(random.randint(100000, 999999))
                        if save_otp(user_to_reset, otp):
                            st.session_state.forgot_user = user_to_reset
                            st.session_state.forgot_step = 3
                            # Mask the phone for security
                            masked_phone = f"*******{phone[-4:]}" if len(phone) >= 4 else phone
                            st.success(f"OTP sent to {masked_phone}! Your demo OTP is: {otp}")
                            st.rerun()
                        else:
                            st.error("Failed to generate OTP.")
                    else:
                        st.error("Username not found.")
                if st.button("Back to Login", key="btn_back1"):
                    st.session_state.forgot_step = 1
                    st.rerun()

            elif st.session_state.forgot_step == 3:
                st.info(f"Resetting password for: {st.session_state.forgot_user}")
                otp_input = st.text_input("Enter 6-Digit OTP", key="reset_otp")
                new_pwd = st.text_input("New Password", type="password", key="reset_new_pwd")
                if st.button("Verify & Reset", key="btn_verify_reset"):
                    if verify_otp(st.session_state.forgot_user, otp_input):
                        if update_password(st.session_state.forgot_user, new_pwd):
                            st.success("Password updated successfully! Please login.")
                            st.session_state.forgot_step = 1
                            st.session_state.forgot_user = None
                            time.sleep(2)
                            st.rerun()
                        else:
                            st.error("Failed to update password.")
                    else:
                        st.error("Invalid OTP.")
                if st.button("Back", key="btn_back2"):
                    st.session_state.forgot_step = 2
                    st.rerun()

        with tab2:
            with st.form("signup_form"):
                new_user = st.text_input("New Username", key="signup_user")
                new_pwd = st.text_input("New Password", type="password", key="signup_pwd")
                new_phone = st.text_input("Phone Number", key="signup_phone", placeholder="e.g. +1234567890")
                if st.form_submit_button("Create Account", use_container_width=True):
                    if not new_phone:
                        st.error("Phone number is required.")
                    else:
                        success, err = register_user(new_user, new_pwd, new_phone)
                        if success:
                            st.success("Account created! Please login.")
                        else:
                            if err == "duplicate":
                                st.warning("User already exists. Choose a different username.")
                            else:
                                st.error(f"Database error during signup: {err}")

# ---------------- LOGOUT ---------------- #
def logout():
    if st.sidebar.button("🚪 Logout"):
        st.session_state.auth = None
        st.rerun()


# ---------------- KPI UI ---------------- #
def show_kpis(kpis):
    if not kpis:
        return
    st.subheader("📊 KPI Dashboard")

    # Inject card CSS once
    st.markdown("""
    <style>
    .kpi-grid { display: flex; flex-wrap: wrap; gap: 14px; margin-bottom: 16px; }
    .kpi-card {
        flex: 1 1 200px;
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
        border: 2px solid #217346;
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: 0 3px 10px rgba(33,115,70,0.12);
        min-width: 180px;
        max-width: 280px;
    }
    .kpi-label {
        font-size: 12px;
        font-weight: 700;
        color: #155e36;
        letter-spacing: 0.4px;
        margin-bottom: 6px;
        line-height: 1.3;
    }
    .kpi-value {
        font-size: 20px;
        font-weight: 800;
        color: #14532d;
        word-break: break-word;
        margin-bottom: 6px;
    }
    .kpi-trend {
        display: inline-block;
        font-size: 10px;
        font-weight: 600;
        background: #217346;
        color: white;
        padding: 2px 8px;
        border-radius: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

    items = list(kpis.items())
    # Render 4 cards per row
    for row_start in range(0, len(items), 4):
        row_items = items[row_start:row_start+4]
        cols = st.columns(len(row_items))
        for col, (k, v) in zip(cols, row_items):
            with col:
                val = v['Current']
                val_str = f"{val:g}" if isinstance(val, (int, float)) else str(val)
                trend = v.get('Trend', '')
                st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-label">{k}</div>
                    <div class="kpi-value">{val_str}</div>
                    {"<span class='kpi-trend'>" + trend + "</span>" if trend else ""}
                </div>
                """, unsafe_allow_html=True)




# ---------------- FONTS ---------------- #
FONTS = ["Arial", "Calibri", "Courier New", "Times New Roman", "Verdana", "Tahoma", "Trebuchet MS"]

# ---------------- PAGES ---------------- #
def upload_page():
    st.title("📊 Excel AI Analyzer Pro")
    st.subheader("Upload & Settings")

    with st.form("upload_form"):
        col1, col2 = st.columns(2)
        with col1:
            theme = st.selectbox("🎨 Theme", list(THEMES.keys()))
        with col2:
            font = st.selectbox("🔤 Font", FONTS)

        uploaded = st.file_uploader("Upload Excel/CSV", type=["csv", "xlsx", "xls"])
        
        st.markdown("---")
        submit = st.form_submit_button("🚀 Analyze & Generate Reports", use_container_width=True)

    if submit:
        if not uploaded:
            st.error("Please upload a file first!")
        else:
            path = f"temp_{uuid.uuid4().hex}_{uploaded.name}"
            with open(path, "wb") as f:
                f.write(uploaded.getbuffer())
            
            with st.spinner("Analyzing data and generating presentation..."):
                options = {
                    "kpis": True, "client": True, "charts": True, "insights": True,
                    "target": True, "variance": True, "category": True,
                    "format_pptx": True, "format_docx": True, "format_pdf": True
                }
                result = process_file(path, theme, font, options=options)
                st.session_state.analysis_result = result
                st.session_state.last_uploaded = uploaded.name

                if result.get("status") == "success":
                    from excel_database import save_report
                    try:
                        sheets = result.get("sheets", {})
                        if sheets:
                            first_sheet = list(sheets.keys())[0]
                            res = sheets[first_sheet]
                            save_report(
                                file_name=f"{uploaded.name} ({first_sheet})",
                                file_path=res.get("ppt_path", ""),
                                user_name=st.session_state.auth,
                                kpis=res.get("overall_metrics", {}),
                                insights=res.get("numeric_insights", []) + res.get("text_insights", []),
                                charts=res.get("charts", {})
                            )
                    except Exception as e:
                        print(f"Failed to save report: {e}")
                
                st.session_state.current_page = "dashboard"
                st.rerun()

def render_sheet_dashboard(result):
    if result.get("top_insights_table"):
        st.subheader("🏆 Key Highlights Summary")
        st.table(result["top_insights_table"])
        st.divider()

    show_kpis(result.get("kpis", {}))
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        if result.get("numeric_insights"):
            st.subheader("📈 Numeric Insights")
            for ins in result["numeric_insights"]:
                st.info(ins)

        if result.get("target_analysis"):
            st.subheader("🎯 Target vs Actual Analysis")
            for ins in result["target_analysis"]:
                st.info(ins)

        if result.get("time_context"):
            st.subheader("⏳ Time Context")
            st.info(result["time_context"])

    with col2:
        if result.get("text_insights"):
            st.subheader("📝 Text Insights")
            for ins in result["text_insights"]:
                st.success(ins)

        if result.get("variance"):
            st.subheader("📊 Variance Analysis")
            for ins in result["variance"]:
                st.success(ins)

        if result.get("category_breakdown"):
            st.subheader("🧩 Category Breakdown")
            for ins in result["category_breakdown"]:
                st.success(ins)

    st.divider()

    if result.get("charts"):
        st.subheader("📉 Generated Charts")
        chart_items = list(result["charts"].items())
        cols = st.columns(min(len(chart_items), 3))
        for i, (k, data) in enumerate(chart_items):
            with cols[i % len(cols)]:
                st.image(data["path"], caption=k, use_container_width=True)
                st.caption(f"💡 {data['conclusion']}")

    st.divider()

    if result.get("full_data"):
        st.subheader("📄 Full Data Preview")
        st.markdown(f"*Total Rows: {len(result['full_data'])} | Total Columns: {len(result['full_data'][0]) if result['full_data'] else 0}*")
        st.dataframe(result["full_data"], use_container_width=True)

    st.divider()

    st.subheader("📥 Download Reports")
    col_d1, col_d2, col_d3 = st.columns(3)

    with col_d1:
        if result.get("ppt_path") and os.path.exists(result["ppt_path"]):
            with open(result["ppt_path"], "rb") as f:
                st.download_button(
                    label="📥 PowerPoint (.pptx)",
                    data=f,
                    file_name=f"Report_{st.session_state.last_uploaded}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    key=f"ppt_{uuid.uuid4()}"
                )

    with col_d2:
        if result.get("docx_path") and os.path.exists(result["docx_path"]):
            with open(result["docx_path"], "rb") as f:
                st.download_button(
                    label="📝 Word Doc (.docx)",
                    data=f,
                    file_name=f"Report_{st.session_state.last_uploaded}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key=f"docx_{uuid.uuid4()}"
                )

    with col_d3:
        if result.get("pdf_path") and os.path.exists(result["pdf_path"]):
            with open(result["pdf_path"], "rb") as f:
                st.download_button(
                    label="📄 PDF Document (.pdf)",
                    data=f,
                    file_name=f"Report_{st.session_state.last_uploaded}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"pdf_{uuid.uuid4()}"
                )

def dashboard_page():
    st.title("📊 Analysis Dashboard")

    result = st.session_state.get("analysis_result")

    if not result:
        st.warning("No analysis results found. Please go back and upload a file.")
        return

    if result.get("status") == "error":
        st.error(result.get("message", "An unknown error occurred."))
    elif "sheets" in result:
        sheet_names = list(result["sheets"].keys())
        if len(sheet_names) > 1:
            st.info(f"📂 Detected {len(sheet_names)} sheets in workbook. Each sheet has its own analysis tab below.")

        tabs = st.tabs([f"📄 {name}" for name in sheet_names])
        for i, name in enumerate(sheet_names):
            with tabs[i]:
                render_sheet_dashboard(result["sheets"][name])
    else:
        # Fallback for old single-sheet result format
        render_sheet_dashboard(result)

# ---------------- HISTORY PAGE ---------------- #
def history_page():
    st.title("📂 Report History")
    
    from excel_database import get_reports
    reports = get_reports()
    
    user_reports = [r for r in reports if r["user_name"] == st.session_state.auth]
    
    if not user_reports:
        st.info("No reports found. Go to Upload to analyze your first file!")
        return

    for r in user_reports:
        with st.expander(f"📄 {r['file_name']} (Analyzed on {r['created_at']})"):
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Value", f"{r['total_value']:g}" if r['total_value'] else "0")
            with col2:
                st.metric("Records Processed", r['records'])
                
            st.write(f"**Max Value:** {r['max_value']:g} | **Min Value:** {r['min_value']:g} | **Average:** {r['avg_value']:g}")
            
            import json
            try:
                insights = json.loads(r['insights'])
                if insights:
                    st.write("**Top Insights:**")
                    for ins in insights[:5]:
                        st.write(f"- {ins}")
            except:
                pass
                
            if os.path.exists(r['file_path']):
                with open(r['file_path'], "rb") as f:
                    st.download_button(
                        label="📥 Download Presentation",
                        data=f,
                        file_name=f"Report_{r['file_name']}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"dl_{r['id']}"
                    )
            else:
                st.warning("Presentation file is no longer available on the server.")

# ---------------- ROUTING ---------------- #
if "current_page" not in st.session_state:
    st.session_state.current_page = "upload"

if st.session_state.auth:
    st.sidebar.success(f"Logged in as: {st.session_state.auth}")
    
    st.sidebar.subheader("Navigation")
    if st.sidebar.button("📤 Upload", use_container_width=True):
        st.session_state.current_page = "upload"
        st.rerun()
    
    if st.session_state.current_page == "dashboard" or st.session_state.get("analysis_result"):
        if st.sidebar.button("📊 Dashboard", use_container_width=True):
            st.session_state.current_page = "dashboard"
            st.rerun()
            
    if st.sidebar.button("📂 History", use_container_width=True):
        st.session_state.current_page = "history"
        st.rerun()
        
    st.sidebar.divider()
    logout()

    if st.session_state.current_page == "upload":
        upload_page()
    elif st.session_state.current_page == "dashboard":
        dashboard_page()
    elif st.session_state.current_page == "history":
        history_page()
else:
    auth_page()