"""
app.py — Factory Production System Pro
Streamlit UI · Google Sheets backend · Session timeout · Audit log · Full permissions
Run: streamlit run app.py
Default login: admin / admin123
"""

import hashlib
import os
import secrets 
import logging
import time
import json
from datetime import datetime
from io import BytesIO

# ======================================================================== #
# FIX #3: DIRECTORY AUTOMATION
# ======================================================================== #
for folder in ['data', 'logs', 'uploads']:
    os.makedirs(folder, exist_ok=True)

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ======================================================================== #
# FIX #1: GOOGLE SHEETS CONNECTION
# ======================================================================== #
import streamlit as st
from streamlit_gsheets import GSheetsConnection
import streamlit as st


from security import SecurityManager, ROLES, ALL_PERMISSIONS, PERMISSION_LABELS, ROLE_DEFAULTS
from data_processor import (
    DataProcessor, DataProcessingError, ProductionDB,
    CANONICAL_COLUMNS, PIECE_WASTE_COLUMNS,
    compute_metrics, generate_sample_data, waste_pct_color, add_section_column,
    export_master_excel, export_pdf_report, REPORTLAB_AVAILABLE,
    DEFAULT_SHIFT_MINUTES, full_machine_label,
)

# ======================================================================== #
# PATHS
# ======================================================================== #
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(BASE_DIR, "data")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
LOG_DIR    = os.path.join(BASE_DIR, "logs")
USER_DB    = os.path.join(DATA_DIR, "users.json")   # Fallback only
MASTER_CSV = os.path.join(DATA_DIR, "master_production_data.csv")
PROD_DB    = os.path.join(DATA_DIR, "production.db")
LOG_PATH   = os.path.join(LOG_DIR,  "processing.log")

# ======================================================================== #
# GOOGLE SHEETS CONNECTION
# ======================================================================== #
@st.cache_resource
def get_gsheets_connection():
    """Initialize Google Sheets connection with caching."""
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        return conn
    except Exception as e:
        logging.error(f"Failed to connect to Google Sheets: {e}")
        return None

# ======================================================================== #
# GOOGLE SHEETS WRAPPER
# ======================================================================== #
class GSheetsWrapper:
    """Wrapper for Google Sheets operations."""
    
    def __init__(self):
        self.conn = get_gsheets_connection()
        self.users_sheet = "users"
        self.production_sheet = "production_data"
        self.audit_sheet = "audit_log"
        self._init_sheets()
    
    def _init_sheets(self):
        """Initialize sheets if they don't exist (checked one by one, since
        GSheetsConnection has no get_sheet_names() method)."""
        if not self.conn:
            return

        def _sheet_exists(name):
            try:
                self.conn.read(worksheet=name, ttl=0)
                return True
            except Exception:
                return False

        try:
            # Create users sheet if it doesn't exist
            if not _sheet_exists(self.users_sheet):
                # Create with default admin user
                admin_data = pd.DataFrame({
                    'username': ['admin'],
                    'password_hash': [''],
                    'salt': [''],
                    'display_name': ['Administrator (Admin)'],
                    'role': ['Admin'],
                    'permissions': [json.dumps(ROLE_DEFAULTS["Admin"])],
                    'is_active': [1],
                    'created_at': [datetime.now().isoformat()],
                    'created_by': ['system']
                })
                self.write_sheet(admin_data, self.users_sheet)

            # Create production sheet if it doesn't exist
            if not _sheet_exists(self.production_sheet):
                prod_data = pd.DataFrame(columns=[
                    "Date", "Shift", "Machine_Name", "Machine_Code", "Operator", "Supervisor",
                    "Output_Quantity", "Wastage_Cigarette", "Wastage_Paper",
                    "Wastage_Tipping_Paper", "Dust", "Stem",
                    "Wastage_Shell_Blanket", "Wastage_Slide_AluFoil",
                    "Wastage_AluFoil_InnerFrame", "Wastage_BOPP",
                    "Stoppage_Reason", "Stoppage_Duration_Min", "Report_Type",
                    "Data_Quality_Flag", "Machine_Label"
                ])
                self.write_sheet(prod_data, self.production_sheet)

            # Create audit log sheet if it doesn't exist
            if not _sheet_exists(self.audit_sheet):
                audit_data = pd.DataFrame(columns=[
                    'timestamp', 'username', 'action', 'table_name',
                    'record_id', 'old_values', 'new_values'
                ])
                self.write_sheet(audit_data, self.audit_sheet)

        except Exception as e:
            logging.error(f"Failed to initialize sheets: {e}")
    
    def read_sheet(self, sheet_name: str) -> pd.DataFrame:
        """Read a sheet from Google Sheets."""
        if not self.conn:
            return pd.DataFrame()
        try:
            return self.conn.read(worksheet=sheet_name, ttl=0)
        except Exception as e:
            logging.error(f"Failed to read sheet {sheet_name}: {e}")
            return pd.DataFrame()

    def write_sheet(self, df: pd.DataFrame, sheet_name: str) -> bool:
        """Write a DataFrame to Google Sheets. Creates the worksheet if it doesn't exist."""
        if not self.conn:
            return False
        try:
            self.conn.update(worksheet=sheet_name, data=df)
        except Exception:
            # Worksheet probably doesn't exist yet — create it instead.
            self.conn.create(worksheet=sheet_name, data=df)
        return True
    
    def append_to_sheet(self, df: pd.DataFrame, sheet_name: str) -> bool:
        """Append data to an existing sheet."""
        if not self.conn:
            return False
        try:
            # Read existing data
            existing = self.read_sheet(sheet_name)
            if existing.empty:
                combined = df
            else:
                combined = pd.concat([existing, df], ignore_index=True)
            return self.write_sheet(combined, sheet_name)
        except Exception as e:
            logging.error(f"Failed to append to sheet {sheet_name}: {e}")
            return False

# Initialize Google Sheets wrapper
gsheets = GSheetsWrapper() if st.secrets.get("connections") else None

# ======================================================================== #
# SECURITY MANAGER (Google Sheets Version)
# ======================================================================== #
class GSheetsSecurityManager:
    """Security manager using Google Sheets backend."""
    
    def __init__(self):
        self.gsheets = gsheets
        self.users_sheet = "users"
        self.audit_sheet = "audit_log"
    
    def _get_users_df(self) -> pd.DataFrame:
        """Get users DataFrame from Google Sheets."""
        if not self.gsheets:
            return pd.DataFrame()
        try:
            df = self.gsheets.read_sheet(self.users_sheet)
        except Exception as e:
            logging.error(f"Failed to read users sheet: {e}")
            return pd.DataFrame()
        df = df.fillna('')
        # NOTE: We deliberately do NOT auto-recreate/overwrite the users
        # sheet here anymore. The sheet is created once with a default
        # admin in _init_sheets() at startup. If a read here comes back
        # empty (e.g. a temporary Google Sheets API hiccup), we simply
        # return the empty result instead of risking overwriting real
        # user data.
        return df
    
    def _save_users_df(self, df: pd.DataFrame):
        """Save users DataFrame to Google Sheets."""
        if self.gsheets:
            self.gsheets.write_sheet(df, self.users_sheet)
    
    def _audit(self, username: str, action: str, table_name: str = None,
               record_id: int = None, old_values=None, new_values=None):
        """Log audit event to Google Sheets."""
        if not self.gsheets:
            return
        audit_row = pd.DataFrame({
            'timestamp': [datetime.now().isoformat()],
            'username': [username],
            'action': [action],
            'table_name': [table_name],
            'record_id': [record_id],
            'old_values': [json.dumps(old_values) if old_values else None],
            'new_values': [json.dumps(new_values) if new_values else None]
        })
        self.gsheets.append_to_sheet(audit_row, self.audit_sheet)
    
    def _hash(self, password: str, salt: str) -> str:
        """Hash password with salt."""
        return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
    
    def authenticate(self, username: str, password: str):
        """Authenticate user."""
        df = self._get_users_df()
        user_rows = df[df['username'] == username]
        if user_rows.empty:
            return False, "Invalid username or password.", None, None
        
        user = user_rows.iloc[0].to_dict()
        if not user.get('is_active', 1):
            return False, "Account deactivated. Contact your admin.", None, None
        
        if self._hash(password, user.get('salt', '')) != user.get('password_hash', ''):
            return False, "Invalid username or password.", None, None
        
        self._audit(username, "LOGIN", "users", None)
        return True, "Login successful.", user.get('role', 'User'), user.get('display_name', username)
    
    def logout_audit(self, username: str):
        """Log logout event."""
        self._audit(username, "LOGOUT", "users")
    
    @staticmethod
    def check_session_timeout(last_activity: float) -> bool:
        """Check if session has timed out."""
        return (time.time() - last_activity) < 3600  # 60 minutes
    
    @staticmethod
    def touch_session() -> float:
        """Touch session to reset timer."""
        return time.time()
    
    def get_user(self, username: str) -> dict:
        """Get user by username."""
        df = self._get_users_df()
        user_rows = df[df['username'] == username]
        if user_rows.empty:
            return None
        user = user_rows.iloc[0].to_dict()
        try:
            user['permissions'] = json.loads(user.get('permissions', '{}'))
        except:
            user['permissions'] = ROLE_DEFAULTS.get(user.get('role', 'User'), {}).copy()
        user['active'] = bool(user.get('is_active', 1))
        return user
    
    def list_users(self) -> list:
        """List all active users."""
        df = self._get_users_df()
        df = df[df['is_active'] == 1]
        users = []
        for _, row in df.iterrows():
            user = row.to_dict()
            try:
                user['permissions'] = json.loads(user.get('permissions', '{}'))
            except:
                user['permissions'] = ROLE_DEFAULTS.get(user.get('role', 'User'), {}).copy()
            user['active'] = True
            users.append(user)
        return users
    
    def add_user(self, username: str, password: str, role: str, display_name: str = "",
                 created_by: str = "admin", custom_perms: dict = None) -> tuple:
        """Add a new user."""
        if role not in ROLES:
            return False, f"Invalid role. Choose from: {', '.join(ROLES)}."
        if not username or not password:
            return False, "Username and password cannot be empty."
        if len(password) < 6:
            return False, "Password must be at least 6 characters."
        
        df = self._get_users_df()
        if username in df['username'].values:
            return False, f"Username '{username}' already exists."
        
        salt = secrets.token_hex(16)
        perms = custom_perms if custom_perms else ROLE_DEFAULTS[role].copy()
        
        new_row = pd.DataFrame({
            'username': [username],
            'password_hash': [self._hash(password, salt)],
            'salt': [salt],
            'display_name': [display_name or username],
            'role': [role],
            'permissions': [json.dumps(perms)],
            'is_active': [1],
            'created_at': [datetime.now().isoformat()],
            'created_by': [created_by]
        })
        
        df = pd.concat([df, new_row], ignore_index=True)
        self._save_users_df(df)
        self._audit(created_by, "CREATE_USER", "users", None,
                    None, {"username": username, "role": role})
        return True, f"User '{username}' created successfully."
    
    def update_user(self, username: str, display_name: str = None, role: str = None,
                    permissions: dict = None, active: bool = None, updated_by: str = "admin") -> tuple:
        """Update an existing user."""
        old = self.get_user(username)
        if not old:
            return False, f"User '{username}' not found."
        
        df = self._get_users_df()
        idx = df[df['username'] == username].index
        if idx.empty:
            return False, f"User '{username}' not found."
        
        if display_name is not None:
            df.loc[idx, 'display_name'] = display_name.strip() or username
        if role is not None:
            if role not in ROLES:
                return False, f"Invalid role '{role}'."
            df.loc[idx, 'role'] = role
        if permissions is not None:
            df.loc[idx, 'permissions'] = json.dumps(permissions)
        if active is not None:
            df.loc[idx, 'is_active'] = 1 if active else 0
        
        self._save_users_df(df)
        self._audit(updated_by, "UPDATE_USER", "users", None,
                    {"display_name": old.get('display_name'), "role": old.get('role'), "is_active": old.get('active', 1)},
                    {"display_name": display_name, "role": role, "is_active": active})
        return True, f"User '{username}' updated."
    
    def delete_user(self, username: str, deleted_by: str = "admin") -> tuple:
        """Soft delete a user."""
        if username == "admin":
            return False, "Cannot deactivate the default admin account."
        
        old = self.get_user(username)
        if not old:
            return False, f"User '{username}' not found."
        
        return self.update_user(username, active=False, updated_by=deleted_by)
    
    def change_password(self, username: str, new_password: str, changed_by: str = None) -> tuple:
        """Change a user's password."""
        if len(new_password) < 6:
            return False, "Password must be at least 6 characters."
        
        user = self.get_user(username)
        if not user:
            return False, f"User '{username}' not found."
        
        df = self._get_users_df()
        idx = df[df['username'] == username].index
        if idx.empty:
            return False, f"User '{username}' not found."
        
        salt = secrets.token_hex(16)
        df.loc[idx, 'password_hash'] = self._hash(new_password, salt)
        df.loc[idx, 'salt'] = salt
        
        self._save_users_df(df)
        self._audit(changed_by or username, "CHANGE_PASSWORD", "users", None)
        return True, "Password changed successfully."
    
    def get_audit_log(self, limit: int = 500) -> list:
        """Get audit log from Google Sheets."""
        if not self.gsheets:
            return []
        df = self.gsheets.read_sheet(self.audit_sheet)
        if df.empty:
            return []
        df = df.sort_values('timestamp', ascending=False).head(limit)
        return df.to_dict('records')
    
    def has_permission(self, username: str, perm: str) -> bool:
        """Check if user has a specific permission."""
        user = self.get_user(username)
        return bool(user.get('permissions', {}).get(perm, False)) if user else False
    
    def can_upload(self, u): return self.has_permission(u, "enter_data")
    def can_delete(self, u): return self.has_permission(u, "delete_data")
    def can_export(self, u): return self.has_permission(u, "export_reports")
    def can_view_reports(self, u): return self.has_permission(u, "view_reports")
    def can_manage_users(self, u): return self.has_permission(u, "manage_users")
    def can_approve(self, u): return self.has_permission(u, "approve_data")


# ======================================================================== #
# PRODUCTION DB (Google Sheets Version)
# ======================================================================== #
class GSheetsProductionDB:
    """Production database using Google Sheets backend."""
    
    def __init__(self):
        self.gsheets = gsheets
        self.sheet_name = "production_data"
    
    def _get_dataframe(self) -> pd.DataFrame:
        """Get production data from Google Sheets."""
        if not self.gsheets:
            return pd.DataFrame()
        return self.gsheets.read_sheet(self.sheet_name)
    
    def _save_dataframe(self, df: pd.DataFrame):
        """Save production data to Google Sheets."""
        if self.gsheets:
            self.gsheets.write_sheet(df, self.sheet_name)
    
    def upsert_from_df(self, df: pd.DataFrame, entered_by: str = "system", source_file: str = "") -> int:
        """Upsert data from DataFrame to Google Sheets."""
        if df.empty or not self.gsheets:
            return 0
        
        # Ensure required columns
        required_cols = ['date', 'machine_code', 'machine_type', 'shift', 'operator',
                         'supervisor', 'output', 'waste_cigarette', 'waste_paper',
                         'waste_tipping', 'dust', 'stem']
        
        for col in required_cols:
            if col not in df.columns:
                df[col] = None
        
        # Add metadata
        df['entered_by'] = entered_by
        df['entered_at'] = datetime.now().isoformat()
        df['source_file'] = source_file
        
        # Read existing data
        existing = self._get_dataframe()
        
        if existing.empty:
            combined = df
        else:
            # Simple merge - remove duplicates by date/machine/shift
            # For production, just append
            combined = pd.concat([existing, df], ignore_index=True)
        
        self._save_dataframe(combined)
        return len(df)
    
    def get_stats(self) -> dict:
        """Get database statistics."""
        df = self._get_dataframe()
        return {
            'active': len(df),
            'deleted': 0,  # No soft delete in GSheets
        }

    def soft_delete_all(self, deleted_by: str = "system") -> int:
        """Clear all production data."""
        count = len(self._get_dataframe())
        empty_df = pd.DataFrame(columns=[
            "Date", "Shift", "Machine_Name", "Machine_Code", "Operator", "Supervisor",
            "Output_Quantity", "Wastage_Cigarette", "Wastage_Paper",
            "Wastage_Tipping_Paper", "Dust", "Stem",
            "Wastage_Shell_Blanket", "Wastage_Slide_AluFoil",
            "Wastage_AluFoil_InnerFrame", "Wastage_BOPP",
            "Stoppage_Reason", "Stoppage_Duration_Min", "Report_Type",
            "Data_Quality_Flag", "Machine_Label"
        ])
        self._save_dataframe(empty_df)
        return count

    def get_audit_log(self, limit: int = 500) -> list:
        """Get audit log from Google Sheets."""
        if not self.gsheets:
            return []
        df = self.gsheets.read_sheet("audit_log")
        if df.empty:
            return []
        df = df.sort_values('timestamp', ascending=False).head(limit)
        return df.to_dict('records')

# ======================================================================== #
# SECURITY SELECTOR
# ======================================================================== #
def get_security_manager():
    """Get the appropriate security manager based on configuration."""
    if gsheets:
        return GSheetsSecurityManager()
    else:
        # Fallback to file-based security
        from security import SecurityManager
        return SecurityManager(USER_DB)

def get_production_db():
    """Get the appropriate production database based on configuration."""
    if gsheets:
        return GSheetsProductionDB()
    else:
        # Fallback to SQLite database
        from data_processor import ProductionDB
        return ProductionDB(PROD_DB)

# ======================================================================== #
# PAGE CONFIG
# ======================================================================== #
st.set_page_config(page_title="Factory Production System — Pro",
                   page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

security = get_security_manager()
prod_db = get_production_db()

# ======================================================================== #
# SESSION
# ======================================================================== #
def init_session():
    defs = {
        "authenticated": False, "username": None, "role": None,
        "display_name": None, "theme": "light", "permissions": {},
        "last_activity": 0.0,
    }
    for k, v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_session()


def check_session_timeout():
    """Auto-logout on inactivity. Call at the top of main()."""
    if st.session_state.authenticated:
        if not security.check_session_timeout(st.session_state.last_activity):
            uname = st.session_state.username
            security.logout_audit(uname)
            for k in ["authenticated", "username", "role", "display_name", "permissions"]:
                st.session_state[k] = False if k == "authenticated" else None
            st.session_state["permissions"] = {}
            st.warning("⏱️ Your session expired after 60 minutes of inactivity. Please log in again.")
            st.rerun()
        else:
            st.session_state["last_activity"] = security.touch_session()


# ======================================================================== #
# THEME / STYLES
# ======================================================================== #
TMPL = "plotly_dark" if st.session_state.theme == "dark" else "plotly_white"
_card_bg = "#1C1F26" if st.session_state.theme == "dark" else "#FFFFFF"
_text_col = "#F0F0F0" if st.session_state.theme == "dark" else "#111111"

st.markdown(f"""
<style>
.main-header{{background:linear-gradient(90deg,#1F3864 0%,#2E5090 60%,#C9A227 100%);
  padding:1.4rem 2rem;border-radius:12px;color:white;margin-bottom:1rem;}}
.main-header h1{{margin:0;font-size:1.75rem;}}
.main-header p{{margin:.2rem 0 0;opacity:.9;}}
.date-badge{{display:inline-block;background:rgba(255,255,255,.18);
  padding:3px 14px;border-radius:20px;font-weight:600;margin-top:6px;}}
.metric-card{{background:{_card_bg};color:{_text_col};border-radius:12px;
  padding:.9rem 1rem;box-shadow:0 2px 10px rgba(0,0,0,.08);
  border-left:5px solid #1F3864;min-height:90px;}}
.metric-card h3{{margin:0;font-size:.72rem;opacity:.65;font-weight:700;
  text-transform:uppercase;letter-spacing:.04em;white-space:nowrap;}}
.metric-card p{{margin:.3rem 0 0;font-size:clamp(.8rem,1.2vw,1.25rem);
  font-weight:700;white-space:normal;word-break:break-all;line-height:1.2;}}
.metric-card small{{font-size:.72rem;opacity:.6;}}
.login-box{{max-width:420px;margin:3rem auto;padding:2rem;
  background:{_card_bg};color:{_text_col};border-radius:16px;
  box-shadow:0 6px 24px rgba(0,0,0,.12);}}
.perm-yes{{color:#27AE60;font-weight:700;}}
.perm-no{{color:#E74C3C;font-weight:700;}}
.audit-row{{font-size:.8rem;font-family:monospace;}}
</style>""", unsafe_allow_html=True)


# ======================================================================== #
# HELPERS
# ======================================================================== #
def load_data() -> pd.DataFrame:
    if gsheets:
        return prod_db._get_dataframe()
    elif os.path.exists(MASTER_CSV):
        df = pd.read_csv(MASTER_CSV, parse_dates=["Date"])
        if "Machine_Label" not in df.columns:
            df["Machine_Label"] = df.apply(
                lambda r: full_machine_label(r.get("Machine_Name"), r.get("Machine_Code")), axis=1
            )
        return df
    return pd.DataFrame(columns=CANONICAL_COLUMNS + ["Data_Quality_Flag", "Machine_Label"])


def save_data(df):
    if gsheets:
        prod_db._save_dataframe(df)
    else:
        df.to_csv(MASTER_CSV, index=False)


def has_perm(perm):
    return bool(st.session_state.permissions.get(perm, False))


def waste_colors(series):
    return [waste_pct_color(p) for p in series]


def metric_card(label, value, sub=""):
    sub_html = f"<small>{sub}</small>" if sub else ""
    st.markdown(
        f'<div class="metric-card"><h3>{label}</h3><p>{value}</p>{sub_html}</div>',
        unsafe_allow_html=True
    )


# ======================================================================== #
# LOGIN
# ======================================================================== #
def login_page():
    st.markdown("""<div class="main-header">
        <h1>🏭 Factory Production System — Pro</h1>
        <p>Secure access · Making & Packing · OEE Dashboards · Full Audit Trail</p>
    </div>""", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.3, 1])
    with col:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.subheader("🔐 Sign In")
        uname = st.text_input("Username")
        pw = st.text_input("Password", type="password")
        if st.button("Log In", use_container_width=True, type="primary"):
            ok, msg, role, dname = security.authenticate(uname, pw)
            if ok:
                user_info = security.get_user(uname)
                st.session_state.update({
                    "authenticated": True,
                    "username": uname,
                    "role": role,
                    "display_name": dname,
                    "permissions": user_info["permissions"] if user_info else {},
                    "last_activity": security.touch_session(),
                })
                st.rerun()
            else:
                st.error(f"❌ {msg}")
        st.caption("Default: **admin / admin123**")
        st.markdown('</div>', unsafe_allow_html=True)


# ======================================================================== #
# EMPLOYEE PERFORMANCE TAB (FIX #4)
# ======================================================================== #
def employee_performance_page():
    st.markdown("""<div class="main-header"><h1>👤 Employee Performance</h1>
        <p>Track production units by employee across all dates</p>
    </div>""", unsafe_allow_html=True)

    if not has_perm("view_reports"):
        st.error("🚫 You don't have permission to view employee performance.")
        return

    df = load_data()
    if df.empty:
        st.info("No production data available yet. Upload data to see employee performance.")
        return

    # Ensure required columns exist
    if 'Operator' not in df.columns:
        st.warning("No 'Operator' column found in data. Employee performance cannot be displayed.")
        return

    # Clean the data
    df = df.copy()
    df['Operator'] = df['Operator'].fillna('Unknown')

    # Compute per-row waste % using the same waste columns used elsewhere in the app
    waste_cols = [c for c in (['Wastage_Cigarette'] + PIECE_WASTE_COLUMNS) if c in df.columns]
    if waste_cols:
        total_waste = df[waste_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
        output_qty = pd.to_numeric(df.get('Output_Quantity', 0), errors='coerce').fillna(0)
        df['Waste_Pct'] = np.where(output_qty > 0, 100 * total_waste / output_qty, 0)
    else:
        df['Waste_Pct'] = 0

    # Get metrics for each employee
    employee_metrics = df.groupby('Operator').agg(
        **{
            'Total Units': ('Output_Quantity', 'sum'),
            'Avg Units': ('Output_Quantity', 'mean'),
            'Days Worked': ('Output_Quantity', 'count'),
            'Avg Waste %': ('Waste_Pct', 'mean'),
        }
    ).round(2)

    employee_metrics = employee_metrics.reset_index()
    employee_metrics = employee_metrics.sort_values('Total Units', ascending=False)

    # Display metrics
    total_employees = len(employee_metrics)
    total_units = employee_metrics['Total Units'].sum()
    avg_units = employee_metrics['Avg Units'].mean()

    col1, col2, col3 = st.columns(3)
    with col1:
        metric_card("Total Employees", total_employees)
    with col2:
        metric_card("Total Production", f"{total_units:,.0f}", "sticks")
    with col3:
        metric_card("Avg per Employee", f"{avg_units:,.0f}", "sticks")

    st.divider()

    # FIX #4: Employee Performance Chart
    st.subheader("📊 Employee Production Units")
    st.caption("Total units produced by each employee")

    # Create bar chart with data labels
    fig = px.bar(
        employee_metrics,
        x='Operator',
        y='Total Units',
        text='Total Units',
        color='Total Units',
        color_continuous_scale=['#C9A227', '#1F3864'],
        template=TMPL,
        labels={'Operator': 'Employee', 'Total Units': 'Total Units (sticks)'}
    )
    fig.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Total Units: %{y:,.0f} sticks<extra></extra>"
    )
    fig.update_layout(
        showlegend=False,
        coloraxis_showscale=False,
        xaxis_tickangle=-40,
        xaxis=dict(tickfont=dict(size=10)),
        margin=dict(b=160),
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # Show detailed table
    with st.expander("📋 Employee Performance Details"):
        display_df = employee_metrics.copy()
        display_df['Total Units'] = display_df['Total Units'].apply(lambda x: f"{x:,.0f}")
        display_df['Avg Units'] = display_df['Avg Units'].apply(lambda x: f"{x:,.0f}")
        display_df['Avg Waste %'] = display_df['Avg Waste %'].apply(lambda x: f"{x:.2f}%" if x > 0 else "N/A")
        st.dataframe(display_df, use_container_width=True)

    # Export option
    if has_perm("export_reports"):
        st.download_button(
            "⬇️ Export Employee Performance CSV",
            employee_metrics.to_csv(index=False).encode(),
            f"employee_performance_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )


# ======================================================================== #
# CHARTS (reusable)
# ======================================================================== #
def render_charts(data: pd.DataFrame, label: str):
    if data.empty:
        st.info(f"No {label} data for the selected filters.")
        return
    m = compute_metrics(data)
    kpi = m["kpis"]
    bm = m["by_machine"]

    # KPI row
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    with c1: metric_card("Total Output", f"{kpi['total_output']:,.0f}", "sticks")
    with c2: metric_card("Avg / Day", f"{kpi['avg_output_per_day']:,.0f}", "sticks/day")
    with c3: metric_card("Waste %", f"{kpi['overall_waste_pct']:.2f}%", "🟢<3 🟡3-5 🟠5-8 🔴>8")
    with c4: metric_card("Cig. Waste", f"{kpi['total_cig_waste_g']:,.0f}", "gm")
    with c5: metric_card("Avg OEE", f"{kpi['avg_oee']:.1f}%")
    with c6: metric_card("Machines", str(kpi["active_machines"]))
    date_min = kpi.get("date_min")
    date_max = kpi.get("date_max")
    if date_min and date_max:
        span = (date_max.date() - date_min.date()).days + 1
        recorded = kpi["days_tracked"]
        days_val = str(recorded) if recorded == span else f"{recorded} of {span}"
        days_sub = f"{date_min.strftime('%d %b')} – {date_max.strftime('%d %b')}"
    else:
        days_val = str(kpi["days_tracked"])
        days_sub = ""
    with c7: metric_card("Days", days_val, days_sub)
    st.markdown("<br>", unsafe_allow_html=True)

    # Row 1
    r1a, r1b = st.columns(2)
    with r1a:
        st.subheader("Production by Machine")
        if not bm.empty:
            bms = bm.sort_values("Total_Output", ascending=False)
            fig = px.bar(bms, x="Machine_Label", y="Total_Output", text="Total_Output",
                         color="Total_Output", color_continuous_scale=["#C9A227", "#1F3864"],
                         template=TMPL,
                         labels={"Machine_Label": "Machine", "Total_Output": "Output (sticks)"})
            fig.update_traces(texttemplate="%{text:,.0f} sticks", textposition="outside",
                              hovertemplate="<b>%{x}</b><br>Output: %{y:,.0f} sticks<extra></extra>")
            fig.update_layout(showlegend=False, coloraxis_showscale=False,
                              xaxis_tickangle=-40,
                              xaxis=dict(tickfont=dict(size=10)),
                              margin=dict(b=160))
            st.plotly_chart(fig, use_container_width=True)

    with r1b:
        st.subheader("Waste Breakdown (piece-count)")
        wb_df = m["waste_breakdown"]
        if not wb_df.empty:
            fig = px.pie(wb_df, names="Waste_Type", values="Total", hole=.4,
                         color_discrete_sequence=px.colors.sequential.Sunset, template=TMPL)
            fig.update_traces(
                texttemplate="%{label}<br>%{percent}<br>%{value:,.0f} sticks",
                hovertemplate="<b>%{label}</b><br>%{value:,.0f} sticks (%{percent})<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.caption("No waste data.")

    # Row 2
    r2a, r2b = st.columns(2)
    with r2a:
        st.subheader("Production Trend")
        bd = m["by_day"]
        if not bd.empty:
            bd = bd.sort_values("Date")
            fig = px.line(bd, x="Date", y="Total_Output", markers=True, text="Total_Output",
                          template=TMPL, labels={"Total_Output": "Output (sticks)"})
            fig.update_traces(line_color="#1F3864",
                              texttemplate="%{text:,.0f}", textposition="top center",
                              hovertemplate="<b>%{x}</b><br>%{y:,.0f} sticks<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.caption("No dated records.")

    with r2b:
        st.subheader("Waste % by Machine")
        if not bm.empty:
            bmw = bm.sort_values("Waste_Pct", ascending=False)
            fig = go.Figure(go.Bar(
                x=bmw["Machine_Label"], y=bmw["Waste_Pct"],
                text=[f"{v:.2f}%" if pd.notna(v) else "N/A" for v in bmw["Waste_Pct"]],
                textposition="outside", marker_color=waste_colors(bmw["Waste_Pct"]),
                hovertemplate="<b>%{x}</b><br>Waste: %{y:.2f}%<extra></extra>"))
            fig.update_layout(template=TMPL, yaxis_title="Waste %",
                              xaxis=dict(tickangle=-40, tickfont=dict(size=10)),
                              margin=dict(b=160),
                annotations=[dict(x=.5, y=-0.55, xref="paper", yref="paper", showarrow=False,
                                  text="🟢 <3% · 🟡 3–5% · 🟠 5–8% · 🔴 >8%", font=dict(size=11))])
            st.plotly_chart(fig, use_container_width=True)

    # Row 3
    r3a, r3b = st.columns(2)
    with r3a:
        st.subheader("OEE Ranking by Machine")
        if not bm.empty:
            fig = px.bar(bm, x="Machine_Label", y="OEE_%", text="OEE_%",
                         color="OEE_%", color_continuous_scale=["#E74C3C", "#F1C40F", "#2ECC71"],
                         template=TMPL, labels={"Machine_Label": "Machine", "OEE_%": "OEE %"})
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_layout(coloraxis_showscale=False, yaxis_title="OEE %",
                              xaxis=dict(tickangle=-40, tickfont=dict(size=10)),
                              margin=dict(b=160))
            st.caption("⚠️ Performance = relative to best machine in range (directional).")
            st.plotly_chart(fig, use_container_width=True)

    with r3b:
        st.subheader("Shift A vs B Comparison")
        bs = m["by_shift"]
        if not bs.empty:
            fig = px.bar(bs, x="Shift", y="Total_Output", text="Total_Output",
                         color="Shift", color_discrete_map={"A": "#1F3864", "B": "#C9A227"},
                         template=TMPL, labels={"Total_Output": "Output (sticks)"})
            fig.update_traces(texttemplate="%{text:,.0f} sticks", textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

    # Downtime by Machine + Reason
    st.subheader("Downtime by Machine & Reason")
    dt = m["downtime"]
    if not dt.empty:
        fig = px.bar(dt.head(15), x="Machine_Label", y="Total_Downtime_Hr",
                     color="Stoppage_Reason", text="Total_Downtime_Hr",
                     template=TMPL,
                     labels={"Machine_Label": "Machine", "Total_Downtime_Hr": "Downtime (hr)",
                             "Stoppage_Reason": "Reason"},
                     title="Top Downtime Events — hover for reason detail")
        fig.update_traces(texttemplate="%{text:.1f} hr", textposition="outside")
        fig.update_layout(xaxis=dict(tickangle=-40, tickfont=dict(size=10)),
                          margin=dict(b=180),
                          legend=dict(orientation="h", y=-0.6))
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("📋 Downtime Detail Table"):
            dt_show = dt[["Machine_Label", "Stoppage_Reason", "Total_Downtime_Hr", "Occurrences"]].copy()
            dt_show = dt_show.rename(columns={
                "Machine_Label": "Machine (Full Code)",
                "Stoppage_Reason": "Reason",
                "Total_Downtime_Hr": "Downtime (hr)",
            })
            dt_show["Downtime (hr)"] = dt_show["Downtime (hr)"].apply(lambda x: f"{x:.2f} hr")
            st.dataframe(dt_show, use_container_width=True)
    else:
        st.caption("No stoppage data recorded.")

    # Full OEE Table
    with st.expander("📋 Full Machine OEE Table"):
        oee_show = bm[["Rank", "Machine_Label", "Total_Output", "Waste_Pct",
                        "Availability_%", "Performance_%", "Quality_%", "OEE_%"]].copy()
        oee_show["Total_Output"] = oee_show["Total_Output"].apply(lambda x: f"{x:,.0f} sticks")
        oee_show["Waste_Pct"] = oee_show["Waste_Pct"].apply(
            lambda x: f"{x:.2f}%" if pd.notna(x) else "N/A")
        for col in ["Availability_%", "Performance_%", "Quality_%", "OEE_%"]:
            oee_show[col] = oee_show[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "N/A")
        oee_show = oee_show.rename(columns={"Machine_Label": "Machine (Full Code)",
                                            "Total_Output": "Output"})
        st.dataframe(oee_show, use_container_width=True)


# ======================================================================== #
# DASHBOARD
# ======================================================================== #
def dashboard_page():
    if not has_perm("view_dashboard"):
        st.error("🚫 You don't have permission to view the dashboard.")
        return

    df = load_data()
    if df.empty:
        st.markdown("""<div class="main-header"><h1>📊 Production Dashboard</h1>
            <p>Live overview · Making & Packing · OEE · Waste %</p></div>""",
            unsafe_allow_html=True)
        st.info("No data yet. Upload reports in **Data Upload**, or try sample data below.")
        if st.button("🎲 Generate Sample Data"):
            sdf = generate_sample_data()
            save_data(sdf)
            st.success("Sample data loaded!")
            st.rerun()
        return

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    df = add_section_column(df)
    vd = df["Date"].dropna()
    if vd.empty:
        st.warning("No valid dates in data.")
        return
    min_d, max_d = vd.min().date(), vd.max().date()

    date_str = f"{min_d.strftime('%d.%m.%y')} → {max_d.strftime('%d.%m.%y')}"
    st.markdown(f"""<div class="main-header">
        <h1>📊 Production Dashboard</h1>
        <p>Live overview · Making &amp; Packing · OEE · Waste %</p>
        <span class="date-badge">📅 {date_str}</span>
    </div>""", unsafe_allow_html=True)

    with st.sidebar:
        st.divider()
        st.markdown("### 🔍 Dashboard Filters")
        sections = sorted(df["Section"].dropna().unique().tolist())
        sec_filter = st.selectbox("Section", ["All"] + sections)
        shifts = sorted(df["Shift"].dropna().unique().tolist())
        shift_filter = st.multiselect("Shift", shifts, default=shifts)

    if min_d == max_d:
        sel = (min_d, max_d)
        st.caption(f"Single date: {min_d.strftime('%d %b %Y')}")
    else:
        sel = st.slider("🗓️ Date Range", min_value=min_d, max_value=max_d,
                        value=(min_d, max_d), format="DD.MM.YY")

    mask = (df["Date"].dt.date >= sel[0]) & (df["Date"].dt.date <= sel[1])
    filtered = df[mask | df["Date"].isna()]
    if shift_filter:
        filtered = filtered[filtered["Shift"].isin(shift_filter) | filtered["Shift"].isna()]
    if filtered.empty:
        st.warning("No records match the selected filters.")
        return

    # Use tabs for sections
    tabs = ["🏭 All Sections", "🚬 Making", "📦 Packing", "👤 Employee Performance"]
    if sec_filter != "All":
        tabs = ["🚬 Making" if sec_filter == "Making" else "📦 Packing", "👤 Employee Performance"]

    tab_objects = st.tabs(tabs)

    for tab, name in zip(tab_objects, tabs):
        with tab:
            if "Employee Performance" in name:
                employee_performance_page()
            elif "All" in name:
                render_charts(filtered, "All")
            elif "Making" in name:
                render_charts(filtered[filtered["Section"] == "Making"], "Making")
            elif "Packing" in name:
                render_charts(filtered[filtered["Section"] == "Packing"], "Packing")

    # Exports
    st.divider()
    st.subheader("📥 Export Reports")
    exp_df = filtered if sec_filter == "All" else filtered[filtered["Section"] == sec_filter]
    exp_m = compute_metrics(exp_df)

    if has_perm("export_reports"):
        ec1, ec2, ec3 = st.columns(3)
        with ec1:
            if st.button("📊 Generate Excel"):
                tmp = os.path.join(DATA_DIR, "report_export.xlsx")
                export_master_excel(exp_df, exp_m, tmp)
                with open(tmp, "rb") as f:
                    st.download_button("⬇️ Download Excel", f.read(),
                        f"production_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with ec2:
            if REPORTLAB_AVAILABLE:
                if st.button("📄 Generate PDF"):
                    tmp = os.path.join(DATA_DIR, "report_export.pdf")
                    export_pdf_report(exp_m, tmp)
                    with open(tmp, "rb") as f:
                        st.download_button("⬇️ Download PDF", f.read(),
                            f"production_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf")
            else:
                st.caption("PDF unavailable — `pip install reportlab`")
        with ec3:
            st.download_button("⬇️ Export CSV", exp_df.to_csv(index=False).encode(),
                f"production_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
    else:
        st.info("🔒 Export requires **Export Reports** permission.")


# ======================================================================== #
# DATA UPLOAD
# ======================================================================== #
def upload_page():
    st.markdown("""<div class="main-header"><h1>📤 Data Upload</h1>
        <p>Upload Excel shift reports — all sheets processed automatically</p>
    </div>""", unsafe_allow_html=True)

    if not has_perm("enter_data"):
        st.error("🚫 You don't have permission to upload data.")
        return

    uploaded = st.file_uploader("Select Excel file(s)", type=["xlsx", "xls"],
                                accept_multiple_files=True)
    if uploaded:
        saved = []
        for f in uploaded:
            dest = os.path.join(UPLOAD_DIR, f.name)
            with open(dest, "wb") as out:
                out.write(f.read())
            saved.append(dest)
        st.success(f"✅ {len(saved)} file(s) ready.")

        if st.button("⚙️ Process Files", type="primary"):
            proc = DataProcessor()
            existing = load_data()
            with st.spinner("Processing..."):
                try:
                    combined = proc.process_files(saved, existing_df=existing)
                    save_data(combined)
                except DataProcessingError as e:
                    st.error(f"❌ {e}")
                    return
                except Exception as e:
                    st.exception(e)
                    return
                st.success(f"✅ Done. Master dataset: {len(combined)} records.")
                if proc.warnings:
                    with st.expander(f"⚠️ {len(proc.warnings)} warning(s)"):
                        for w in proc.warnings:
                            st.write(f"- {w}")
                if proc.errors:
                    with st.expander(f"❌ {len(proc.errors)} error(s)"):
                        for e in proc.errors:
                            st.write(f"- {e}")
                st.dataframe(combined.tail(30), use_container_width=True)

    # Delete Data (soft delete)
    if has_perm("delete_data"):
        st.divider()
        st.subheader("🗑️ Archive / Delete Production Data")
        db_stats = prod_db.get_stats()
        col1, col2, col3 = st.columns(3)
        col1.metric("Active Records (GSheets)", f"{db_stats['active']:,}")
        col2.metric("Archived Records", f"{db_stats['deleted']:,}")
        col3.metric("CSV Records", f"{len(load_data()):,}")

        st.info(
            "**Soft delete** — records are archived, not permanently removed. "
            "The audit trail is preserved. Admins can inspect archived data via the Audit Log."
        )
        st.warning("⚠️ This archives ALL current production data and resets the dashboard.")
        confirm = st.checkbox(
            "I understand this archives all data (soft delete — data is preserved)."
        )
        if confirm:
            if st.button("🗑️ ARCHIVE ALL DATA", type="primary"):
                archived = prod_db.soft_delete_all(deleted_by=st.session_state.username)
                if os.path.exists(MASTER_CSV):
                    os.remove(MASTER_CSV)
                for fname in os.listdir(UPLOAD_DIR):
                    fpath = os.path.join(UPLOAD_DIR, fname)
                    try:
                        os.remove(fpath)
                    except Exception:
                        pass
                st.success(f"✅ {archived} records archived (soft delete). Dashboard reset.")
                st.rerun()
    elif os.path.exists(MASTER_CSV):
        st.divider()
        st.info("🔒 Data deletion requires **Delete Data** permission (Admin or Manager).")


# ======================================================================== #
# REPORTS
# ======================================================================== #
def reports_page():
    st.markdown("""<div class="main-header"><h1>📋 Reports</h1>
        <p>Filter, review, and export production records</p></div>""", unsafe_allow_html=True)

    if not has_perm("view_reports"):
        st.error("🚫 You don't have permission to view reports.")
        return

    df = load_data()
    if df.empty:
        st.info("No data available yet.")
        return
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)

    with st.expander("🔎 Filters", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            machines = sorted(df["Machine_Label"].dropna().unique().tolist())
            sel_m = st.multiselect("Machine", machines)
        with c2:
            shifts = sorted(df["Shift"].dropna().unique().tolist())
            sel_s = st.multiselect("Shift", shifts)
        with c3:
            sections = sorted(df["Section"].dropna().unique().tolist()) if "Section" in df.columns else []
            if not sections:
                df = add_section_column(df)
                sections = sorted(df["Section"].dropna().unique().tolist())
            sel_sec = st.multiselect("Section", sections)
        with c4:
            vd = df["Date"].dropna()
            if not vd.empty:
                dr = st.date_input("Date Range", value=(vd.min().date(), vd.max().date()))
            else:
                dr = None

    filtered = df.copy()
    if "Section" not in filtered.columns:
        filtered = add_section_column(filtered)
    if sel_m:
        filtered = filtered[filtered["Machine_Label"].isin(sel_m)]
    if sel_s:
        filtered = filtered[filtered["Shift"].isin(sel_s)]
    if sel_sec:
        filtered = filtered[filtered["Section"].isin(sel_sec)]
    if dr and isinstance(dr, tuple) and len(dr) == 2:
        filtered = filtered[
            (filtered["Date"].dt.date >= dr[0]) &
            (filtered["Date"].dt.date <= dr[1]) | filtered["Date"].isna()
        ]

    st.write(f"**{len(filtered)} record(s)** match your filters.")

    disp = filtered[["Date", "Shift", "Machine_Label", "Section", "Operator", "Supervisor",
                      "Output_Quantity", "Wastage_Cigarette", "Stoppage_Reason",
                      "Stoppage_Duration_Min", "Data_Quality_Flag"]].copy()
    disp["Output_Quantity"] = disp["Output_Quantity"].apply(
        lambda x: f"{x:,.0f} sticks" if pd.notna(x) else "")
    disp["Wastage_Cigarette"] = disp["Wastage_Cigarette"].apply(
        lambda x: f"{x:,.0f} gm" if pd.notna(x) else "")
    disp["Stoppage_Duration_Min"] = disp["Stoppage_Duration_Min"].apply(
        lambda x: f"{x:.0f} min" if pd.notna(x) else "")
    disp = disp.rename(columns={
        "Machine_Label": "Machine (Full Code)",
        "Output_Quantity": "Output",
        "Wastage_Cigarette": "Cig Waste",
        "Stoppage_Duration_Min": "Downtime",
    })
    st.dataframe(disp, use_container_width=True, height=420)

    if has_perm("export_reports"):
        st.download_button("⬇️ Export CSV", filtered.to_csv(index=False).encode(),
            f"report_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
    else:
        st.info("🔒 Export requires **Export Reports** permission.")


# ======================================================================== #
# AUDIT LOG
# ======================================================================== #
def audit_log_page():
    st.markdown("""<div class="main-header"><h1>📜 Audit Log</h1>
        <p>Every login, data change, and permission update — timestamped</p></div>""",
        unsafe_allow_html=True)

    if not has_perm("manage_users"):
        st.error("🚫 Admin access required.")
        return

    tab_users, tab_data = st.tabs(["👤 User / Auth Events", "📦 Data Events"])

    with tab_users:
        rows = security.get_audit_log(limit=500)
        if rows:
            audit_df = pd.DataFrame(rows)[["timestamp", "username", "action",
                                           "table_name", "record_id",
                                           "old_values", "new_values"]]
            audit_df = audit_df.rename(columns={
                "timestamp": "Timestamp",
                "username": "User",
                "action": "Action",
                "table_name": "Table",
                "record_id": "Record ID",
                "old_values": "Before",
                "new_values": "After",
            })
            st.write(f"**{len(audit_df)} events** (most recent first)")
            col1, col2 = st.columns(2)
            with col1:
                users_f = ["All"] + sorted(audit_df["User"].dropna().unique().tolist())
                sel_user = st.selectbox("Filter by user", users_f, key="au_user")
            with col2:
                actions_f = ["All"] + sorted(audit_df["Action"].dropna().unique().tolist())
                sel_action = st.selectbox("Filter by action", actions_f, key="au_action")
            show = audit_df.copy()
            if sel_user != "All":
                show = show[show["User"] == sel_user]
            if sel_action != "All":
                show = show[show["Action"] == sel_action]
            st.dataframe(show, use_container_width=True, height=450)
            if has_perm("export_reports"):
                st.download_button("⬇️ Export Audit CSV", show.to_csv(index=False).encode(),
                    f"audit_log_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv")
        else:
            st.info("No audit events recorded yet.")

    with tab_data:
        rows = prod_db.get_audit_log(limit=500)
        if rows:
            dlog = pd.DataFrame(rows)[["timestamp", "username", "action",
                                       "record_id", "old_values", "new_values"]]
            dlog = dlog.rename(columns={
                "timestamp": "Timestamp",
                "username": "User",
                "action": "Action",
                "record_id": "Record ID",
                "old_values": "Before",
                "new_values": "After",
            })
            st.write(f"**{len(dlog)} data events**")
            st.dataframe(dlog, use_container_width=True, height=450)
        else:
            st.info("No data events recorded yet.")


# ======================================================================== #
# ADMIN PANEL
# ======================================================================== #
def admin_page():
    st.markdown("""<div class="main-header"><h1>🛠️ Admin Panel</h1>
        <p>User management · Permissions · System status</p></div>""", unsafe_allow_html=True)

    if not has_perm("manage_users"):
        st.error("🚫 Admins only.")
        return

    tab1, tab2, tab3 = st.tabs(["👥 User Management", "🔐 My Account", "📈 System Status"])

    # Tab 1: User Management
    with tab1:
        st.subheader("All Users")
        users = security.list_users()
        tbl = []
        for u in users:
            tbl.append({
                "Display Name": u["display_name"],
                "Username": u["username"],
                "Role": u["role"],
                "Active": "✅" if u["active"] else "❌",
                "Permissions": " | ".join(
                    PERMISSION_LABELS[p] for p in ALL_PERMISSIONS if u["permissions"].get(p)
                ),
            })
        st.dataframe(pd.DataFrame(tbl), use_container_width=True)
        st.divider()

        # Add user
        st.subheader("➕ Add New User")
        with st.form("add_user", clear_on_submit=True):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                nu = st.text_input("Username")
            with c2:
                npw = st.text_input("Password", type="password")
            with c3:
                ndn = st.text_input("Display Name")
            with c4:
                nr = st.selectbox("Role", ROLES)
            st.caption("Permissions (leave blank to use role defaults):")
            pc = st.columns(4)
            custom_perms = {}
            defaults = ROLE_DEFAULTS[nr]
            for i, perm in enumerate(ALL_PERMISSIONS):
                with pc[i % 4]:
                    custom_perms[perm] = st.checkbox(
                        PERMISSION_LABELS[perm], value=defaults.get(perm, False),
                        key=f"new_{perm}")
            if st.form_submit_button("➕ Create User", type="primary"):
                ok, msg = security.add_user(nu, npw, nr, display_name=ndn,
                                            created_by=st.session_state.username,
                                            custom_perms=custom_perms)
                st.success(f"✅ {msg}") if ok else st.error(f"❌ {msg}")
                if ok:
                    st.rerun()

        st.divider()

        # Edit user
        st.subheader("✏️ Edit Existing User")
        usernames = [u["username"] for u in users]
        target = st.selectbox("Select user to edit", usernames)
        tinfo = next((u for u in users if u["username"] == target), None)
        if tinfo:
            with st.form("edit_user"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    new_dn = st.text_input("Display Name", value=tinfo["display_name"])
                with c2:
                    new_role = st.selectbox("Role", ROLES, index=ROLES.index(tinfo["role"]))
                with c3:
                    new_act = st.checkbox("Active", value=tinfo["active"])
                st.caption("Permissions:")
                pc2 = st.columns(4)
                new_perms = {}
                for i, perm in enumerate(ALL_PERMISSIONS):
                    with pc2[i % 4]:
                        new_perms[perm] = st.checkbox(
                            PERMISSION_LABELS[perm],
                            value=tinfo["permissions"].get(perm, False),
                            key=f"edit_{perm}")
                if st.form_submit_button("💾 Save Changes", type="primary"):
                    ok, msg = security.update_user(
                        target, display_name=new_dn, role=new_role,
                        permissions=new_perms, active=new_act,
                        updated_by=st.session_state.username)
                    st.success(f"✅ {msg}") if ok else st.error(f"❌ {msg}")
                    if ok:
                        st.rerun()

            with st.form("reset_pw", clear_on_submit=True):
                rp = st.text_input(f"New password for '{target}'", type="password")
                if st.form_submit_button("🔑 Reset Password"):
                    if rp:
                        ok, msg = security.change_password(
                            target, rp, changed_by=st.session_state.username)
                        st.success(msg) if ok else st.error(msg)

            if target != "admin":
                st.divider()
                if st.button(f"🗑️ Deactivate '{target}' (soft delete)", type="primary"):
                    ok, msg = security.delete_user(
                        target, deleted_by=st.session_state.username)
                    st.success(f"✅ {msg}") if ok else st.error(f"❌ {msg}")
                    if ok:
                        st.rerun()

    # Tab 2: My Account
    with tab2:
        me = security.get_user(st.session_state.username)
        st.subheader(f"Account: {me['display_name']} ({st.session_state.username})")
        st.write(f"**Role:** {me['role']}")
        with st.form("my_account"):
            new_dn_me = st.text_input("Display Name", value=me["display_name"])
            c1, c2 = st.columns(2)
            with c1:
                pw1 = st.text_input("New Password", type="password")
            with c2:
                pw2 = st.text_input("Confirm Password", type="password")
            if st.form_submit_button("💾 Update My Account"):
                changed = False
                if new_dn_me != me["display_name"]:
                    ok, msg = security.update_user(
                        st.session_state.username, display_name=new_dn_me,
                        updated_by=st.session_state.username)
                    if ok:
                        st.session_state.display_name = new_dn_me
                        st.success(f"✅ Display name updated to '{new_dn_me}'")
                        changed = True
                    else:
                        st.error(f"❌ {msg}")
                if pw1:
                    if pw1 != pw2:
                        st.error("❌ Passwords don't match.")
                    else:
                        ok, msg = security.change_password(
                            st.session_state.username, pw1,
                            changed_by=st.session_state.username)
                        st.success(f"✅ {msg}") if ok else st.error(f"❌ {msg}")
                        changed = True
                if changed:
                    st.rerun()

    # Tab 3: System Status
    with tab3:
        df = load_data()
        db_stats = prod_db.get_stats()
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("CSV Records", f"{len(df):,}")
        c2.metric("GSheets Active", f"{db_stats['active']:,}")
        c3.metric("GSheets Archived", f"{db_stats['deleted']:,}")
        c4.metric("Registered Users", len(security.list_users()))
        if "Machine_Label" in df.columns:
            c5.metric("Machines", df["Machine_Label"].nunique())
        st.divider()
        c6, c7 = st.columns(2)
        c6.metric("CSV Present", "✅ Yes" if os.path.exists(MASTER_CSV) else "❌ No")
        c7.metric("Google Sheets", "✅ Connected" if gsheets else "❌ Not Connected")
        if os.path.exists(LOG_PATH):
            with st.expander("📜 Recent Log (last 60 lines)"):
                with open(LOG_PATH) as f:
                    st.code("".join(f.readlines()[-60:]) or "Empty.")


# ======================================================================== #
# MAIN
# ======================================================================== #
def main():
    check_session_timeout()

    if not st.session_state.authenticated:
        login_page()
        return

    with st.sidebar:
        dname = st.session_state.display_name or st.session_state.username
        st.markdown(f"### 👤 {dname}")
        st.caption(f"Role: **{st.session_state.role}**  |  @{st.session_state.username}")
        st.divider()

        theme = st.radio("Theme", ["Light", "Dark"],
                         index=0 if st.session_state.theme == "light" else 1, horizontal=True)
        if theme.lower() != st.session_state.theme:
            st.session_state.theme = theme.lower()
            st.rerun()

        st.divider()
        pages = []
        if has_perm("view_dashboard"):
            pages.append("📊 Dashboard")
        if has_perm("enter_data"):
            pages.append("📤 Data Upload")
        if has_perm("view_reports"):
            pages.append("📋 Reports")
        if has_perm("manage_users"):
            pages.append("🛠️ Admin Panel")
        if has_perm("manage_users"):
            pages.append("📜 Audit Log")

        if not pages:
            st.warning("No pages available for your role.")
            if st.button("🚪 Log Out"):
                security.logout_audit(st.session_state.username)
                for k in ["authenticated", "username", "role", "display_name", "permissions"]:
                    st.session_state[k] = None if k != "authenticated" else False
                st.rerun()
            return

        choice = st.radio("Navigation", pages, label_visibility="collapsed")
        st.divider()
        st.caption(f"⏱️ Session active")
        if st.button("🚪 Log Out", use_container_width=True):
            security.logout_audit(st.session_state.username)
            for k in ["authenticated", "username", "role", "display_name", "permissions"]:
                st.session_state[k] = None if k != "authenticated" else False
            st.session_state["authenticated"] = False
            st.rerun()

    if choice == "📊 Dashboard":
        dashboard_page()
    elif choice == "📤 Data Upload":
        upload_page()
    elif choice == "📋 Reports":
        reports_page()
    elif choice == "🛠️ Admin Panel":
        admin_page()
    elif choice == "📜 Audit Log":
        audit_log_page()


if __name__ == "__main__":
    main()