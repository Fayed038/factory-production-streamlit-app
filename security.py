"""
security.py — Factory Production System Pro
SQLite backend · SHA-256 hashing (NO SALT) · Session timeout · RBAC · Audit trail · Soft delete
"""

import json, hashlib, sqlite3, os, time
from datetime import datetime
from pathlib import Path

ROLES = ["Admin", "Manager", "Supervisor", "User"]

ROLE_DEFAULTS = {
    "Admin":      {"view_dashboard": True, "enter_data": True, "view_reports": True,
                   "export_reports": True, "manage_users": True, "delete_data": True,
                   "approve_data": True, "change_settings": True},
    "Manager":    {"view_dashboard": True, "enter_data": True, "view_reports": True,
                   "export_reports": True, "manage_users": False, "delete_data": True,
                   "approve_data": True, "change_settings": False},
    "Supervisor": {"view_dashboard": True, "enter_data": True, "view_reports": True,
                   "export_reports": False, "manage_users": False, "delete_data": False,
                   "approve_data": False, "change_settings": False},
    "User":       {"view_dashboard": True, "enter_data": True, "view_reports": False,
                   "export_reports": False, "manage_users": False, "delete_data": False,
                   "approve_data": False, "change_settings": False},
}

ALL_PERMISSIONS = list(ROLE_DEFAULTS["Admin"].keys())

PERMISSION_LABELS = {
    "view_dashboard":  "View Dashboard",
    "enter_data":      "Enter / Upload Data",
    "view_reports":    "View Reports",
    "export_reports":  "Export Reports",
    "manage_users":    "Manage Users",
    "delete_data":     "Delete Data",
    "approve_data":    "Approve Data",
    "change_settings": "Change Settings",
}

SESSION_TIMEOUT_SECONDS = 3600  # 60 minutes


class SecurityManager:
    def __init__(self, db_path: str):
        p = Path(db_path)
        self.db_path = str(p.parent / "users.db")
        self._init_db()
        # One-time migration from users.json if it exists
        json_path = p.parent / "users.json"
        if json_path.exists():
            self._migrate_from_json(str(json_path))

    # ------------------------------------------------------------------ #
    # DB helpers
    # ------------------------------------------------------------------ #
    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    def _init_db(self):
        with self._conn() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS users (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    username     TEXT UNIQUE NOT NULL,
                    password     TEXT NOT NULL,
                    display_name TEXT NOT NULL DEFAULT '',
                    role         TEXT NOT NULL DEFAULT 'User',
                    permissions  TEXT NOT NULL DEFAULT '{}',
                    is_active    INTEGER NOT NULL DEFAULT 1,
                    created_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
                    created_by   TEXT DEFAULT 'system'
                );

                CREATE TABLE IF NOT EXISTS audit_log (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    username   TEXT NOT NULL,
                    action     TEXT NOT NULL,
                    table_name TEXT,
                    record_id  INTEGER,
                    old_values TEXT,
                    new_values TEXT,
                    timestamp  DATETIME DEFAULT CURRENT_TIMESTAMP
                );

                CREATE INDEX IF NOT EXISTS idx_audit_ts   ON audit_log(timestamp);
                CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_log(username);
            """)
            count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
            if count == 0:
                # Create default admin with clean SHA-256 (NO SALT)
                pw_hash = self._hash("admin123")
                perms = json.dumps(ROLE_DEFAULTS["Admin"])
                conn.execute(
                    "INSERT INTO users (username,password,display_name,role,permissions,created_by)"
                    " VALUES (?,?,?,?,?,?)",
                    ("admin", pw_hash, "Administrator (Admin)", "Admin", perms, "system")
                )

    def _migrate_from_json(self, json_path: str):
        """Import users.json into SQLite (runs once, then renames the file)."""
        try:
            with open(json_path, "r") as f:
                data = json.load(f)
        except Exception:
            return
        users = data.get("users", {})
        if not users:
            return
        with self._conn() as conn:
            for uname, u in users.items():
                if conn.execute("SELECT id FROM users WHERE username=?", (uname,)).fetchone():
                    continue
                role = u.get("role", "User")
                if role == "Viewer":
                    role = "User"
                perms = u.get("permissions", ROLE_DEFAULTS.get(role, ROLE_DEFAULTS["User"]).copy())
                
                # Handle password: if it's already hashed, use it; otherwise hash it
                password = u.get("password", u.get("password_hash", ""))
                if password and not self._is_hashed(password):
                    password = self._hash(password)
                elif not password:
                    password = self._hash("password123")
                
                conn.execute(
                    "INSERT INTO users (username,password,display_name,role,permissions,created_by)"
                    " VALUES (?,?,?,?,?,?)",
                    (uname, password, u.get("display_name", uname.capitalize()),
                     role, json.dumps(perms), u.get("created_by", "migration"))
                )
        try:
            os.rename(json_path, json_path + ".migrated")
        except Exception:
            pass

    @staticmethod
    def _is_hashed(password: str) -> bool:
        """Check if password looks like a SHA-256 hash (64 hex chars)."""
        return len(password) == 64 and all(c in '0123456789abcdef' for c in password.lower())

    # ------------------------------------------------------------------ #
    # 🔐 FIXED: CLEAN SHA-256 HASHING (NO SALT)
    # ------------------------------------------------------------------ #
    @staticmethod
    def _hash(password: str) -> str:
        """Simple SHA-256 hashing - NO SALT."""
        return hashlib.sha256(password.encode('utf-8')).hexdigest()

    # ------------------------------------------------------------------ #
    # Audit log
    # ------------------------------------------------------------------ #
    def _audit(self, username: str, action: str, table_name: str = None,
               record_id: int = None, old_values=None, new_values=None):
        with self._conn() as conn:
            conn.execute(
                "INSERT INTO audit_log (username,action,table_name,record_id,old_values,new_values)"
                " VALUES (?,?,?,?,?,?)",
                (username, action, table_name, record_id,
                 json.dumps(old_values) if old_values is not None else None,
                 json.dumps(new_values) if new_values is not None else None)
            )

    def get_audit_log(self, limit: int = 500) -> list:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM audit_log ORDER BY timestamp DESC LIMIT ?", (limit,)
            ).fetchall()
        return [dict(r) for r in rows]

    # ------------------------------------------------------------------ #
    # 🔐 FIXED: AUTHENTICATION (NO SALT)
    # ------------------------------------------------------------------ #
    def authenticate(self, username: str, password: str):
        """Authenticate user - clean SHA-256 comparison (NO SALT)."""
        with self._conn() as conn:
            row = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        
        if not row:
            return False, "Invalid username or password.", None, None
        
        user = dict(row)
        
        if not user.get("is_active", 1):
            return False, "Account deactivated. Contact your admin.", None, None
        
        # 🔑 FIXED: Direct SHA-256 comparison - NO SALT
        hashed_input = self._hash(password)
        
        if hashed_input != user["password"]:
            return False, "Invalid username or password.", None, None
        
        self._audit(username, "LOGIN", "users", user["id"])
        return True, "Login successful.", user["role"], user["display_name"]

    def logout_audit(self, username: str):
        self._audit(username, "LOGOUT", "users")

    # ------------------------------------------------------------------ #
    # Session helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def check_session_timeout(last_activity: float) -> bool:
        return (time.time() - last_activity) < SESSION_TIMEOUT_SECONDS

    @staticmethod
    def touch_session() -> float:
        return time.time()

    # ------------------------------------------------------------------ #
    # 🔐 FIXED: USER CRUD (NO SALT)
    # ------------------------------------------------------------------ #
    def add_user(self, username, password, role, display_name="",
                 created_by="admin", custom_perms=None):
        if role not in ROLES:
            return False, f"Invalid role. Choose from: {', '.join(ROLES)}."
        if not username or not password:
            return False, "Username and password cannot be empty."
        if len(password) < 6:
            return False, "Password must be at least 6 characters."
        
        perms = custom_perms if custom_perms else ROLE_DEFAULTS[role].copy()
        pw_hash = self._hash(password)  # 🔑 NO SALT
        
        try:
            with self._conn() as conn:
                conn.execute(
                    "INSERT INTO users (username,password,display_name,role,permissions,created_by)"
                    " VALUES (?,?,?,?,?,?)",
                    (username, pw_hash, display_name or username, role, json.dumps(perms), created_by)
                )
        except sqlite3.IntegrityError:
            return False, f"Username '{username}' already exists."
        
        self._audit(created_by, "CREATE_USER", "users", None,
                    None, {"username": username, "role": role})
        return True, f"User '{username}' created successfully."

    def update_user(self, username, display_name=None, role=None,
                    permissions=None, active=None, updated_by="admin"):
        old = self.get_user(username)
        if not old:
            return False, f"User '{username}' not found."
        
        sets, vals = [], []
        if display_name is not None:
            sets.append("display_name=?"); vals.append(display_name.strip() or username)
        if role is not None:
            if role not in ROLES:
                return False, f"Invalid role '{role}'."
            sets.append("role=?"); vals.append(role)
        if permissions is not None:
            sets.append("permissions=?"); vals.append(json.dumps(permissions))
        if active is not None:
            sets.append("is_active=?"); vals.append(1 if active else 0)
        if not sets:
            return True, "Nothing to update."
        
        vals.append(username)
        with self._conn() as conn:
            conn.execute(f"UPDATE users SET {','.join(sets)} WHERE username=?", vals)
        
        self._audit(updated_by, "UPDATE_USER", "users", old.get("id"),
                    {k: old.get(k) for k in ("display_name", "role", "is_active")},
                    {"display_name": display_name, "role": role, "is_active": active})
        return True, f"User '{username}' updated."

    def delete_user(self, username, deleted_by="admin"):
        if username == "admin":
            return False, "Cannot deactivate the default admin account."
        old = self.get_user(username)
        if not old:
            return False, f"User '{username}' not found."
        with self._conn() as conn:
            conn.execute("UPDATE users SET is_active=0 WHERE username=?", (username,))
        self._audit(deleted_by, "SOFT_DELETE_USER", "users", old.get("id"),
                    {"is_active": 1}, {"is_active": 0})
        return True, f"User '{username}' deactivated (soft delete — data preserved)."

    # 🔐 FIXED: CHANGE PASSWORD (NO SALT)
    def change_password(self, username, new_password, changed_by=None):
        if len(new_password) < 6:
            return False, "Password must be at least 6 characters."
        user = self.get_user(username)
        if not user:
            return False, f"User '{username}' not found."
        
        pw_hash = self._hash(new_password)  # 🔑 NO SALT
        
        with self._conn() as conn:
            conn.execute(
                "UPDATE users SET password=? WHERE username=?",
                (pw_hash, username)
            )
        self._audit(changed_by or username, "CHANGE_PASSWORD", "users", user.get("id"))
        return True, "Password changed successfully."

    def list_users(self) -> list:
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT * FROM users WHERE is_active=1 ORDER BY role, username"
            ).fetchall()
        result = []
        for row in rows:
            u = dict(row)
            try:
                u["permissions"] = json.loads(u.get("permissions", "{}"))
            except Exception:
                u["permissions"] = ROLE_DEFAULTS.get(u.get("role", "User"), {}).copy()
            u["active"] = bool(u.get("is_active", 1))
            result.append(u)
        return result

    def get_user(self, username) -> dict | None:
        with self._conn() as conn:
            row = conn.execute(
                "SELECT * FROM users WHERE username=?", (username,)
            ).fetchone()
        if not row:
            return None
        u = dict(row)
        try:
            u["permissions"] = json.loads(u.get("permissions", "{}"))
        except Exception:
            u["permissions"] = ROLE_DEFAULTS.get(u.get("role", "User"), {}).copy()
        u["active"] = bool(u.get("is_active", 1))
        return u

    # ------------------------------------------------------------------ #
    # Permission shortcuts
    # ------------------------------------------------------------------ #
    def has_permission(self, username, perm) -> bool:
        info = self.get_user(username)
        return bool(info["permissions"].get(perm, False)) if info else False

    def can_upload(self, u):       return self.has_permission(u, "enter_data")
    def can_delete(self, u):       return self.has_permission(u, "delete_data")
    def can_export(self, u):       return self.has_permission(u, "export_reports")
    def can_view_reports(self, u): return self.has_permission(u, "view_reports")
    def can_manage_users(self, u): return self.has_permission(u, "manage_users")
    def can_approve(self, u):      return self.has_permission(u, "approve_data")
    