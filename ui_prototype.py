import json
import sqlite3
import tkinter as tk
import hashlib
import hmac
import os
import secrets
import socket
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import *
from contract_generator import CONFIG_PATH, append_generated_rows, batch_convert_docx_to_pdf, generate_contract_from_gui, load_config



def _resolve_db_path():
    configured = (os.getenv("AQ_DB_PATH") or "").strip()
    if not configured:
        return Path(__file__).with_name("contract_suite.db")
    path = Path(configured).expanduser()
    if path.exists() and path.is_dir():
        return path / "contract_suite.db"
    return path


DB_PATH = _resolve_db_path()
APP_BRAND = "AQ Creativity"

SPACE_XS = 4
SPACE_SM = 8
SPACE_MD = 16
SPACE_LG = 24
SPACE_XL = 32
CORNER_RADIUS = 8

FONT_FAMILY = "Segoe UI"

COLOR_BG = "#0f172a"
COLOR_SIDEBAR = "#111827"
COLOR_CARD = "#1e293b"
COLOR_CARD_2 = "#0b1220"
COLOR_PRIMARY = "#3b82f6"
COLOR_PRIMARY_HOVER = "#2563eb"
COLOR_TEXT = "#e5e7eb"
COLOR_MUTED = "#94a3b8"
COLOR_BORDER = "#1f2937"
LOGO_FILE_NAME = "logo.png"
LOGO_CANDIDATES = (LOGO_FILE_NAME, "aq_creativity_logo.gif", "aq_creativity_logo.pgm", "aq_creativity_logo.ppm")

LIGHT_STATUS_COLORS = {
    "NEW": "#eeeeee",
    "DRAFT": "#cfe2ff",
    "SENT": "#e0ccff",
    "SIGNED": "#d4edda",
    "PAID": "#c6f6d5",
}

DARK_STATUS_COLORS = {
    "NEW": "#22252d",
    "DRAFT": "#1f314f",
    "SENT": "#352752",
    "SIGNED": "#1d3b2b",
    "PAID": "#204235",
}

STATUS_OPTIONS = list(LIGHT_STATUS_COLORS.keys())

FALLBACK_VENDORS = ["Ali Tech", "Sara Beauty", "Omar Food", "Lama Travel"]


class SearchableCombo(ttk.Combobox):
    def __init__(self, master, values, **kwargs):
        super().__init__(master, values=values, **kwargs)
        self.full = list(values)
        self.bind("<KeyRelease>", self.filter)

    def set_values(self, values):
        self.full = list(values)
        self["values"] = self.full

    def filter(self, _):
        txt = self.get().lower()
        self["values"] = [v for v in self.full if txt in v.lower()] or self.full


class ContractSuiteDB:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.conn = self._open_connection(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self._configure_runtime()
        self._init_schema()

    def _open_connection(self, path: Path):
        candidate = path.expanduser()
        try:
            candidate.parent.mkdir(parents=True, exist_ok=True)
            return sqlite3.connect(candidate, timeout=15)
        except sqlite3.Error:
            fallback = Path(__file__).with_name("contract_suite.db")
            fallback.parent.mkdir(parents=True, exist_ok=True)
            self.db_path = fallback
            return sqlite3.connect(fallback, timeout=15)

    def _configure_runtime(self):
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.execute("PRAGMA busy_timeout = 15000")
        self.conn.execute("PRAGMA journal_mode = WAL")
        self.conn.execute("PRAGMA synchronous = NORMAL")

    def _init_schema(self):
        self.conn.executescript(
            """
            PRAGMA foreign_keys = ON;
            CREATE TABLE IF NOT EXISTS tasks (
                id TEXT PRIMARY KEY, brand TEXT NOT NULL DEFAULT '', amount TEXT NOT NULL DEFAULT '',
                contract_type TEXT NOT NULL DEFAULT 'auto', status TEXT NOT NULL DEFAULT 'NEW',
                created_at TEXT NOT NULL, updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS subtasks (
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 task_id TEXT NOT NULL,

                  vendor TEXT,
                  license_number TEXT,
                  iban TEXT,

                  channel TEXT,
                  platforms TEXT,
                  ad_type TEXT,
                  qty TEXT,
                  details TEXT,
                  price TEXT,

                  created_at TEXT,
                  updated_at TEXT,

                  FOREIGN KEY(task_id) REFERENCES tasks(id)
            );
            CREATE TABLE IF NOT EXISTS vendors (
                id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
                license_number TEXT NOT NULL UNIQUE, created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS bank_accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT, vendor_id INTEGER NOT NULL,
                bank_name TEXT NOT NULL, account_name TEXT NOT NULL, iban TEXT NOT NULL UNIQUE,
                account_number TEXT NOT NULL, swift_code TEXT, created_at TEXT NOT NULL,
                FOREIGN KEY(vendor_id) REFERENCES vendors(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS brands (
                id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE, created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS generated_contracts (
                contract_id TEXT PRIMARY KEY, task_id TEXT NOT NULL, brand_name TEXT NOT NULL,
                amount TEXT NOT NULL, contract_type TEXT NOT NULL, generated_at TEXT NOT NULL,
                FOREIGN KEY(task_id) REFERENCES tasks(id)
            );
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY, value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                email TEXT NOT NULL DEFAULT '',
                full_name TEXT NOT NULL DEFAULT '',
                profile_color TEXT NOT NULL DEFAULT '#4f8cff',
                profile_icon TEXT NOT NULL DEFAULT '👤',
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'member',
                created_at TEXT NOT NULL,
                last_login TEXT
            );
            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                actor_username TEXT NOT NULL,
                action TEXT NOT NULL,
                entity_type TEXT NOT NULL,
                entity_id TEXT,
                details TEXT,
                created_at TEXT NOT NULL
            );
            """
        )
        cols = [r["name"] for r in self.conn.execute("PRAGMA table_info(subtasks)").fetchall()]
        if "price" not in cols:
            self.conn.execute("ALTER TABLE subtasks ADD COLUMN price TEXT NOT NULL DEFAULT '0'")

        user_cols = [r["name"] for r in self.conn.execute("PRAGMA table_info(users)").fetchall()]
        if "email" not in user_cols:
            self.conn.execute("ALTER TABLE users ADD COLUMN email TEXT NOT NULL DEFAULT ''")
        if "full_name" not in user_cols:
            self.conn.execute("ALTER TABLE users ADD COLUMN full_name TEXT NOT NULL DEFAULT ''")
        if "profile_color" not in user_cols:
            self.conn.execute("ALTER TABLE users ADD COLUMN profile_color TEXT NOT NULL DEFAULT '#4f8cff'")
        if "profile_icon" not in user_cols:
            self.conn.execute("ALTER TABLE users ADD COLUMN profile_icon TEXT NOT NULL DEFAULT '👤'")

        self.conn.commit()


    def _legacy_password_hash(self, password: str):
        return hashlib.sha256(password.encode("utf-8")).hexdigest()

    def _password_hash(self, password: str):
        salt = secrets.token_hex(16)
        iterations = 200_000
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), iterations).hex()
        return f"pbkdf2_sha256${iterations}${salt}${dk}"

    def _password_matches(self, password: str, stored_hash: str):
        raw = (stored_hash or "").strip()
        if not raw:
            return False
        if raw.startswith("pbkdf2_sha256$"):
            parts = raw.split("$", 3)
            if len(parts) != 4:
                return False
            try:
                iterations = int(parts[1])
            except ValueError:
                return False
            salt, expected = parts[2], parts[3]
            candidate = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), iterations).hex()
            return hmac.compare_digest(candidate, expected)
        return hmac.compare_digest(raw, self._legacy_password_hash(password))

    def now(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def create_task(self, task_id: str, contract_type: str = "auto", status: str = "NEW"):
        now = self.now()
        self.conn.execute("INSERT INTO tasks(id, brand, amount, contract_type, status, created_at, updated_at) VALUES (?, '', '', ?, ?, ?, ?)", (task_id, contract_type, status, now, now))
        self.conn.commit()

    def delete_task(self, task_id: str):
        # Clean dependent rows first to avoid FK failures on existing DBs.
        self.conn.execute("DELETE FROM generated_contracts WHERE task_id = ?", (task_id,))
        self.conn.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
        self.conn.commit()

    def upsert_task(self, task_id: str, brand: str, amount: str, contract_type: str, status: str):
        self.conn.execute("UPDATE tasks SET brand=?, amount=?, contract_type=?, status=?, updated_at=? WHERE id=?", (brand, amount, contract_type, status, self.now(), task_id))
        self.conn.commit()

    def list_tasks(self):
        return self.conn.execute(
            """
            SELECT t.id, t.brand, t.status, t.amount, t.contract_type,
                   COUNT(s.id) AS vendor_count, DATE(t.created_at) AS created_date
            FROM tasks t LEFT JOIN subtasks s ON s.task_id=t.id
            GROUP BY t.id, t.brand, t.status, t.amount, t.contract_type, t.created_at
            ORDER BY t.created_at DESC
            """
        ).fetchall()

    def get_task(self, task_id: str):
        return self.conn.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)).fetchone()

    def list_subtasks(self, task_id: str):
        return self.conn.execute(
            "SELECT id, vendor, license_number, iban, channel, platforms, ad_type, qty, details, price FROM subtasks WHERE task_id=? ORDER BY id",
            (task_id,)
        ).fetchall()

    def create_subtask(
            self,
            task_id,
            vendor,
            license_number,
            iban,
            channel,
            platforms,
            ad_type,
            qty,
            details,
            price,
    ):
        now = self.now()
        self.conn.execute(
            "INSERT INTO subtasks(task_id,vendor,license_number,iban,channel,platforms,ad_type,qty,details,price,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
(task_id, vendor, license_number, iban, channel, platforms, ad_type, qty, details, price, now, now),
        )
        self.conn.commit()

    def update_subtask(self, subtask_id: int, vendor: str, channel: str, platforms: str, ad_type: str, qty: str, details: str, price: str):
        self.conn.execute(
            "UPDATE subtasks SET vendor=?, channel=?, platforms=?, ad_type=?, qty=?, details=?, price=?, updated_at=? WHERE id=?",
            (vendor, channel, platforms, ad_type, qty, details, price, self.now(), subtask_id),
        )
        self.conn.commit()

    def delete_subtask(self, subtask_id: int):
        self.conn.execute("DELETE FROM subtasks WHERE id=?", (subtask_id,))
        self.conn.commit()

    def log_generated_contract(self, contract_id: str, task_id: str, brand_name: str, amount: str, contract_type: str):
        self.conn.execute(
            "INSERT OR REPLACE INTO generated_contracts(contract_id,task_id,brand_name,amount,contract_type,generated_at) VALUES (?,?,?,?,?,?)",
            (contract_id, task_id, brand_name, amount, contract_type, self.now()),
        )
        self.conn.commit()

    def list_generated_contracts_for_task(self, task_id: str):
        return self.conn.execute(
            "SELECT contract_id, generated_at FROM generated_contracts WHERE task_id=? ORDER BY generated_at DESC",
            (task_id,),
        ).fetchall()

    def get_setting(self, key: str, default: str = ""):
        row = self.conn.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str):
        self.conn.execute(
            "INSERT INTO app_settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )
        self.conn.commit()

    def count_users(self):
        row = self.conn.execute("SELECT COUNT(*) AS total FROM users").fetchone()
        return int(row["total"]) if row else 0

    def get_runtime_snapshot(self):
        task_total = self.conn.execute("SELECT COUNT(*) AS total FROM tasks").fetchone()
        subtask_total = self.conn.execute("SELECT COUNT(*) AS total FROM subtasks").fetchone()
        return {
            "db_path": str(self.db_path),
            "tasks": int(task_total["total"]) if task_total else 0,
            "subtasks": int(subtask_total["total"]) if subtask_total else 0,
            "users": self.count_users(),
        }

    def create_user(self, username: str, email: str, full_name: str, password: str, role: str = "member"):
        clean_username = username.strip().lower()
        clean_email = email.strip().lower()
        clean_name = full_name.strip()
        if not clean_name:
            raise ValueError("Full name is required")
        if not clean_username:
            raise ValueError("Username is required")
        if "@" not in clean_email:
            raise ValueError("Valid email is required")
        if len(password) < 4:
            raise ValueError("Password must be at least 4 characters")

        existing_email = self.conn.execute("SELECT 1 FROM users WHERE lower(email)=?", (clean_email,)).fetchone()
        if existing_email:
            raise ValueError("Email is already in use")

        self.conn.execute(
            "INSERT INTO users(username, email, full_name, password_hash, role, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (clean_username, clean_email, clean_name, self._password_hash(password), role, self.now()),
        )
        self.conn.commit()

    def authenticate_user(self, identifier: str, password: str):
        clean_identifier = identifier.strip().lower()
        row = self.conn.execute(
            "SELECT id, username, email, full_name, profile_color, profile_icon, role, password_hash FROM users WHERE lower(username)=? OR lower(email)=?",
            (clean_identifier, clean_identifier),
        ).fetchone()
        if not row or not self._password_matches(password, row["password_hash"]):
            return None
        if not str(row["password_hash"]).startswith("pbkdf2_sha256$"):
            self.conn.execute("UPDATE users SET password_hash=? WHERE id=?", (self._password_hash(password), row["id"]))
        self.conn.execute("UPDATE users SET last_login=? WHERE id=?", (self.now(), row["id"]))
        self.conn.commit()
        return {
            "id": row["id"],
            "username": row["username"],
            "email": row["email"],
            "full_name": row["full_name"],
            "profile_color": row["profile_color"],
            "profile_icon": row["profile_icon"],
            "role": row["role"],
        }

    def get_user_by_username(self, username: str):
        clean_username = username.strip().lower()
        if not clean_username:
            return None
        return self.conn.execute(
            "SELECT id, username, email, full_name, profile_color, profile_icon, role FROM users WHERE lower(username)=? OR lower(email)=?",
            (clean_username, clean_username),
        ).fetchone()

    def update_user_profile(self, user_id: int, full_name: str, username: str, email: str, profile_color: str, profile_icon: str):
        clean_name = full_name.strip()
        clean_username = username.strip().lower()
        clean_email = email.strip().lower()
        if not clean_name:
            raise ValueError("Full name is required")
        if not clean_username:
            raise ValueError("Username is required")
        if "@" not in clean_email:
            raise ValueError("Valid email is required")
        dup = self.conn.execute(
            "SELECT id FROM users WHERE (lower(username)=? OR lower(email)=?) AND id<>?",
            (clean_username, clean_email, user_id),
        ).fetchone()
        if dup:
            raise ValueError("Username or email already in use")
        self.conn.execute(
            "UPDATE users SET full_name=?, username=?, email=?, profile_color=?, profile_icon=? WHERE id=?",
            (clean_name, clean_username, clean_email, profile_color.strip() or "#4f8cff", profile_icon.strip() or "👤", user_id),
        )
        self.conn.commit()

    def change_user_password(self, user_id: int, current_password: str, new_password: str):
        row = self.conn.execute("SELECT password_hash FROM users WHERE id=?", (user_id,)).fetchone()
        if not row or not self._password_matches(current_password, row["password_hash"]):
            raise ValueError("Current password is incorrect")
        if len(new_password) < 4:
            raise ValueError("Password must be at least 4 characters")
        self.conn.execute("UPDATE users SET password_hash=? WHERE id=?", (self._password_hash(new_password), user_id))
        self.conn.commit()

    def reset_password_by_email(self, email: str, new_password: str, recovery_key: str):
        clean_email = email.strip().lower()
        if not clean_email:
            raise ValueError("Email is required")
        if len(new_password) < 4:
            raise ValueError("Password must be at least 4 characters")
        if not self.verify_recovery_key(recovery_key):
            raise ValueError("Recovery key is invalid")

        cur = self.conn.execute("UPDATE users SET password_hash=? WHERE lower(email)=?", (self._password_hash(new_password), clean_email))
        self.conn.commit()
        return cur.rowcount

    def reset_all_accounts(self):
        self.conn.execute("DELETE FROM users")
        self.conn.commit()

    def set_recovery_key(self, recovery_key: str):
        clean_key = recovery_key.strip()
        if len(clean_key) < 6:
            raise ValueError("Recovery key must be at least 6 characters")
        self.set_setting("recovery_key_hash", self._password_hash(clean_key))

    def has_recovery_key(self):
        return bool(self.get_setting("recovery_key_hash", ""))

    def verify_recovery_key(self, recovery_key: str):
        stored_hash = self.get_setting("recovery_key_hash", "")
        if not stored_hash:
            return False
        clean = recovery_key.strip()
        ok = self._password_matches(clean, stored_hash)
        if ok and not str(stored_hash).startswith("pbkdf2_sha256$"):
            self.set_setting("recovery_key_hash", self._password_hash(clean))
        return ok

    def clear_recovery_key(self):
        self.conn.execute("DELETE FROM app_settings WHERE key='recovery_key_hash'")
        self.conn.commit()

    def log_audit(self, actor_username: str, action: str, entity_type: str, entity_id: str = "", details: str = ""):
        self.conn.execute(
            "INSERT INTO audit_logs(actor_username, action, entity_type, entity_id, details, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (actor_username or "unknown", action, entity_type, entity_id, details, self.now()),
        )
        self.conn.commit()

    def list_recent_audit(self, limit: int = 100):
        return self.conn.execute(
            "SELECT actor_username, action, entity_type, entity_id, details, created_at FROM audit_logs ORDER BY id DESC LIMIT ?",
            (limit,),
        ).fetchall()

    def get_license_numbers(self):
        return [r["license_number"] for r in self.conn.execute("SELECT license_number FROM vendors ORDER BY license_number").fetchall()]

    def get_vendor_name_by_license(self, license_number: str):
        row = self.conn.execute("SELECT name FROM vendors WHERE license_number=?", (license_number,)).fetchone()
        if row and row["name"] and row["name"].strip():
            return row["name"].strip()
        return ""

    def get_vendor_profile_by_name(self, vendor_name: str):
        return self.conn.execute(
            """
            SELECT v.name, v.license_number, ba.bank_name, ba.account_name, ba.iban, ba.account_number, COALESCE(ba.swift_code, '') AS swift_code
            FROM vendors v
            LEFT JOIN bank_accounts ba ON ba.vendor_id=v.id
            WHERE v.name=?
            ORDER BY ba.id DESC
            LIMIT 1
            """,
            (vendor_name,),
        ).fetchone()

    def get_vendor_names(self):
        return [r["name"] for r in self.conn.execute("SELECT DISTINCT name FROM vendors ORDER BY name").fetchall()]

    def upsert_brand(self, name: str):
        if not name:
            return
        self.conn.execute("INSERT INTO brands(name,created_at) VALUES (?,?) ON CONFLICT(name) DO NOTHING", (name, self.now()))
        self.conn.commit()

    def get_all_brands(self):
        rows = self.conn.execute("SELECT name AS brand FROM brands UNION SELECT DISTINCT brand FROM tasks WHERE brand!='' ORDER BY brand").fetchall()
        return [r["brand"] for r in rows if r["brand"]]

    def get_ibans_for_license(self, license_number: str):
        rows = self.conn.execute(
            "SELECT ba.iban FROM bank_accounts ba JOIN vendors v ON ba.vendor_id=v.id WHERE v.license_number=? ORDER BY ba.iban",
            (license_number,),
        ).fetchall()
        return [r["iban"] for r in rows]

    def get_bank_info_by_iban(self, iban: str):
        return self.conn.execute(
            "SELECT bank_name,account_name,account_number,COALESCE(swift_code,'') AS swift_code FROM bank_accounts WHERE iban=?",
            (iban,),
        ).fetchone()

    def upsert_vendor_bank(self, vendor_name: str, license_number: str, bank_name: str, account_name: str, iban: str, account_number: str, swift_code: str):
        if not license_number:
            return
        now = self.now()
        self.conn.execute(
            "INSERT INTO vendors(name,license_number,created_at) VALUES (?,?,?) ON CONFLICT(license_number) DO UPDATE SET name=excluded.name",
            (vendor_name or "", license_number, now),
        )
        vendor_id = self.conn.execute("SELECT id FROM vendors WHERE license_number=?", (license_number,)).fetchone()["id"]
        if iban:
            self.conn.execute(
                """
                INSERT INTO bank_accounts(vendor_id,bank_name,account_name,iban,account_number,swift_code,created_at)
                VALUES (?,?,?,?,?,?,?)
                ON CONFLICT(iban) DO UPDATE SET vendor_id=excluded.vendor_id, bank_name=excluded.bank_name,
                    account_name=excluded.account_name, account_number=excluded.account_number, swift_code=excluded.swift_code
                """,
                (vendor_id, bank_name or "", account_name or "", iban, account_number or "", swift_code or "", now),
            )
        self.conn.commit()

    def close(self):
        self.conn.close()


def create_task_id(existing_ids):
    nums = [int(i[1:]) for i in existing_ids if i.startswith("C") and i[1:].isdigit()]
    return f"C{(max(nums)+1 if nums else 1000):04d}"


class ContractSuiteApp:
    def __init__(self):
        self.db = ContractSuiteDB(DB_PATH)
        self.editing_subtask_id = None
        self.dashboard_window = None
        self.dashboard_rows = {}
        self.current_user = None

        self.root = tk.Tk()

        self.app_icon_image = None
        for icon_name in ("icon.png",) + LOGO_CANDIDATES:
            icon_path = Path(__file__).with_name(icon_name)
            if not icon_path.exists():
                continue
            try:
                icon = tk.PhotoImage(file=str(icon_path))
                if icon.width() > 48:
                    icon = icon.subsample(max(1, icon.width() // 48))
                if icon.height() > 48:
                    icon = icon.subsample(1, max(1, icon.height() // 48))
                self.app_icon_image = icon
                self.root.iconphoto(True, icon)
                break
            except tk.TclError:
                continue

        self.app_title_var = tk.StringVar(value=self.db.get_setting("app_title", APP_BRAND))
        self.default_type_setting = tk.StringVar(value=self.db.get_setting("default_contract_type", "auto"))
        self.default_status_setting = tk.StringVar(value=self.db.get_setting("default_status", "NEW"))
        self.theme_setting = tk.StringVar(value=self.db.get_setting("theme_mode", "dark"))

        self.root.title(f"{self.app_title_var.get() or APP_BRAND} – Dashboard")
        self.logo_image = None
        self.apply_theme(self.theme_setting.get())
        self.root.geometry("1300x860")
        self.root.minsize(980, 680)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.current_task = tk.StringVar()
        self.brand_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.status_var = tk.StringVar(value="NEW")
        self.template_key_var = tk.StringVar(value="auto")
        self.available_template_keys = []

        self.search_var = tk.StringVar()
        self.status_filter = tk.StringVar(value="ALL")
        self.brand_filter = tk.StringVar(value="ALL")

        self.license_var = tk.StringVar()
        self.iban_var = tk.StringVar()
        self.vendor_name_auto_var = tk.StringVar()  # read-only autofill
        self.bank_name_var = tk.StringVar()
        self.account_name_var = tk.StringVar()
        self.account_number_var = tk.StringVar()
        self.swift_code_var = tk.StringVar()

        self.vendor_var = tk.StringVar()
        self.channel_var = tk.StringVar()
        self.type_sub_var = tk.StringVar(value="Store Visit")
        self.qty_var = tk.StringVar(value="1")
        self.detail_var = tk.StringVar()
        self.vendor_price_var = tk.StringVar(value="0")
        self.platform_vars = {}
        self.platform_channel_vars = {}
        self.quick_panel = None
        self.quick_panel_overlay = None
        self.popup_locks = {}
        self.dashboard_search_var = tk.StringVar()
        self.runtime_badge_var = tk.StringVar()

        self._build_layout()
        self._bind_events()
        self.refresh_task_rows()
        self.refresh_brand_combo()
        self.refresh_license_combo()
        self.refresh_vendor_combo()
        self.refresh_template_dropdown()
        self.refresh_runtime_badge()
        self.show_vendor_panel()
        self._set_editor_empty_state(True)
        self.require_login()


    def apply_theme(self, mode: str):
        self.current_theme = mode
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        if mode == "light":
            bg = "#f6f9ff"
            fg = "#1f2937"
            control_bg = "#ffffff"
            control_fg = "#111827"
            border = "#d6deeb"
            tree_bg = "#f8fbff"
            tree_fg = "#111827"
            heading_bg = "#edf3ff"
            sidebar_bg = "#edf2fb"
        else:
            bg = COLOR_BG
            fg = COLOR_TEXT
            control_bg = COLOR_CARD
            control_fg = COLOR_TEXT
            border = COLOR_BORDER
            tree_bg = COLOR_CARD_2
            tree_fg = COLOR_TEXT
            heading_bg = "#1a2437"
            sidebar_bg = COLOR_SIDEBAR

        self.ui_colors = {
            "bg": bg,
            "fg": fg,
            "control_bg": control_bg,
            "control_fg": control_fg,
            "border": border,
            "accent": COLOR_PRIMARY if mode == "light" else "#4f8cff",
            "accent_soft": "#dbeafe" if mode == "light" else "#1b2a46",
            "muted": COLOR_MUTED,
            "danger": "#ef4444",
            "success": "#22c55e",
        }

        self.root.configure(bg=bg)
        style.configure("TFrame", background=bg)
        style.configure("Card.TFrame", background=control_bg)
        style.configure("Sidebar.TFrame", background=sidebar_bg)
        right_panel_bg = "#17253d" if mode == "dark" else "#eef4ff"
        style.configure("RightPanel.TFrame", background=right_panel_bg)
        style.configure("TLabelframe", background=bg, foreground=fg, borderwidth=0)
        style.configure("TLabelframe.Label", background=bg, foreground=fg, font=(FONT_FAMILY, 11, "bold"))
        style.configure("TLabel", background=bg, foreground=fg, font=(FONT_FAMILY, 11))
        style.configure("Header.TLabel", background=bg, foreground=fg, font=(FONT_FAMILY, 24, "bold"))
        style.configure("SubHeader.TLabel", background=bg, foreground=fg, font=(FONT_FAMILY, 18, "bold"))
        style.configure("Muted.TLabel", background=bg, foreground="#8ea0c2", font=(FONT_FAMILY, 10))
        style.configure(
            "TButton",
            padding=10,
            background=control_bg,
            foreground=control_fg,
            bordercolor=border,
            lightcolor=control_bg,
            darkcolor=control_bg,
            relief="flat",
        )
        style.map("TButton", background=[("active", "#2f4f8f"), ("pressed", "#3b5ea8")], foreground=[("disabled", "#7c8696")])
        style.configure(
            "TEntry",
            fieldbackground=control_bg,
            foreground=control_fg,
            insertcolor=control_fg,
            bordercolor=border,
            lightcolor=control_bg,
            darkcolor=control_bg,
        )
        style.map("TEntry", bordercolor=[("focus", self.ui_colors["accent"])], lightcolor=[("focus", self.ui_colors["accent"])], darkcolor=[("focus", self.ui_colors["accent"])])
        style.configure(
            "TCombobox",
            fieldbackground=control_bg,
            background=control_bg,
            foreground=control_fg,
            arrowcolor=control_fg,
            bordercolor=border,
            lightcolor=control_bg,
            darkcolor=control_bg,
        )
        style.map("TCombobox", fieldbackground=[("readonly", control_bg)], background=[("readonly", control_bg)], foreground=[("readonly", control_fg)])
        style.configure("Treeview", fieldbackground=tree_bg, background=tree_bg, foreground=tree_fg, rowheight=34, borderwidth=0, relief="flat", font=(FONT_FAMILY, 10))
        style.configure("Treeview.Heading", background=heading_bg, foreground=fg, font=(FONT_FAMILY, 10, "bold"))
        style.map("Treeview", background=[("selected", "#3f5d8a")], foreground=[("selected", "#ffffff")])
        style.configure("TCheckbutton", background=bg, foreground=fg)
        style.configure("Platform.TCheckbutton", background=bg, foreground=fg)

        self._apply_task_tree_status_colors()

        if hasattr(self, "logo_button"):
            self.logo_button.configure(bg=bg, activebackground=bg, highlightbackground=bg)
        if hasattr(self, "logo_image_label"):
            self.logo_image_label.configure(bg=self.ui_colors["control_bg"])
        if hasattr(self, "search_entry"):
            self.search_entry.configure(
                bg="#ffffff" if mode == "light" else "#0f1b33",
                fg="#111827" if mode == "light" else "#e8eefc",
                insertbackground="#111827" if mode == "light" else "#e8eefc",
                highlightbackground=self.ui_colors["accent_soft"],
                highlightcolor=self.ui_colors["accent"],
            )

    def _apply_task_tree_status_colors(self):
        if not hasattr(self, "task_tree"):
            return
        dark_mode = getattr(self, "current_theme", "dark") == "dark"
        zebra_a = "#131a2c" if dark_mode else "#f8faff"
        zebra_b = "#0f1626" if dark_mode else "#eef3fb"
        self.task_tree.tag_configure("zebra_even", background=zebra_a)
        self.task_tree.tag_configure("zebra_odd", background=zebra_b)
        status_fg = {
            "NEW": "#8fb3ff",
            "DRAFT": "#5db0ff",
            "SENT": "#c193ff",
            "SIGNED": "#5fd48f",
            "PAID": "#62e0a8",
        } if dark_mode else {
            "NEW": "#3b6fd8",
            "DRAFT": "#2f85c8",
            "SENT": "#7c4ed0",
            "SIGNED": "#2c8b58",
            "PAID": "#2b9c67",
        }
        for status in STATUS_OPTIONS:
            self.task_tree.tag_configure(status, foreground=status_fg.get(status, "#dbe7ff" if dark_mode else "#1f2430"))
        self.task_tree.tag_configure("hover", background="#1a2742" if dark_mode else "#e4eefc")

    def _animate_widget_tap(self, widget, base_bg="#2a2d34", tap_bg="#424754"):
        try:
            widget.configure(bg=tap_bg)
            widget.after(120, lambda: widget.winfo_exists() and widget.configure(bg=base_bg))
        except Exception:
            return

    def _create_animated_button(self, parent, text, command, **kwargs):
        btn = tk.Button(parent, text=text, bg=kwargs.pop("bg", "#2a2d34"), fg=kwargs.pop("fg", "#f4f6fb"), activebackground=kwargs.pop("activebackground", "#3a3e48"), activeforeground=kwargs.pop("activeforeground", "#ffffff"), relief="flat", bd=0, padx=kwargs.pop("padx", 8), pady=kwargs.pop("pady", 3), cursor="hand2", **kwargs)

        def _run():
            self._animate_widget_tap(btn, base_bg=btn.cget("bg"))
            command()

        btn.configure(command=_run)
        return btn

    def _parse_price(self, raw_value: str):
        text = (raw_value or "").strip().replace(",", "")
        if not text:
            return 0.0
        try:
            return float(text)
        except ValueError:
            return 0.0

    def _recalculate_task_amount(self):
        tid = self.current_task.get()
        if not tid:
            self.amount_var.set("0.00")
            return
        total = sum(self._parse_price(sub["price"]) for sub in self.db.list_subtasks(tid))
        self.amount_var.set(f"{total:.2f}")

    def require_login(self):
        remembered_user = self.db.get_setting("remember_device_user", "")
        if remembered_user:
            remembered_row = self.db.get_user_by_username(remembered_user)
            if remembered_row:
                self.current_user = dict(remembered_row)
                self.user_badge_var.set(f"Signed in: {self.current_user.get('full_name') or self.current_user['username']} ({self.current_user.get('role','member')})")
                if hasattr(self, "profile_menu_btn"):
                    self.profile_menu_btn.configure(text=f"{self.current_user.get('profile_icon','👤')} {(self.current_user.get('full_name') or self.current_user['username']).split()[0]} ▾")
                self._audit("auto_login", "profile", self.current_user['username'])
                self.root.lift()
                return

        authenticated = self.open_auth_popup(required=True)
        if authenticated:
            self.root.lift()
            return
        self.root.destroy()

    def _current_username(self):
        if self.current_user and self.current_user.get("username"):
            return self.current_user["username"]
        return "unknown"

    def _has_permission(self, action: str):
        if not self.current_user:
            return False
        role = self.current_user.get("role", "member")
        allowed = {
            "admin": {"import_export", "settings", "manage_profiles", "delete", "edit", "generate"},
            "member": {"delete", "edit", "generate"},
        }
        return action in allowed.get(role, set())

    def _require_permission(self, action: str, feature_name: str):
        if self._has_permission(action):
            return True
        role = self.current_user.get("role", "member") if self.current_user else "none"
        messagebox.showerror("Permission denied", f"{feature_name} requires permission. Current role: {role}")
        return False

    def _audit(self, action: str, entity_type: str, entity_id: str = "", details: str = ""):
        self.db.log_audit(self._current_username(), action, entity_type, entity_id, details)

    def refresh_runtime_badge(self):
        if not hasattr(self, "runtime_badge_var"):
            return
        snap = self.db.get_runtime_snapshot()
        host = socket.gethostname()
        self.runtime_badge_var.set(
            f"Host: {host}  •  Users: {snap['users']}  •  Tasks: {snap['tasks']}  •  DB: {Path(snap['db_path']).name}"
        )

    def _build_layout(self):
        main = ttk.Frame(self.root, padding=SPACE_MD)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(main, style="Sidebar.TFrame", padding=SPACE_MD)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        sidebar.columnconfigure(0, weight=1)

        ttk.Frame(sidebar, style="Sidebar.TFrame", height=6).grid(row=0, column=0, sticky="ew")
        nav_items = [
            ("Dashboard", self.open_asana_dashboard),
            ("Contracts", lambda: None),
            ("Vendors", self.open_vendor_master_popup),
            ("Templates", self.manage_templates_popup),
            ("Reports", self.open_reports_popup),
            ("Calendar", lambda: None),
            ("Audit Log", self.open_audit_log_popup),
            ("Settings", self.open_settings_popup),
        ]
        self.sidebar_active = tk.StringVar(value="Contracts")
        self.sidebar_buttons = {}
        for idx, (label, command) in enumerate(nav_items, start=1):
            row = tk.Frame(sidebar, bg=self.ui_colors["control_bg"], highlightthickness=0)
            row.grid(row=idx, column=0, sticky="ew", pady=3)
            row.grid_columnconfigure(1, weight=1)
            accent = tk.Frame(row, width=3, bg=self.ui_colors["control_bg"])
            accent.grid(row=0, column=0, sticky="ns")
            btn = tk.Button(
                row,
                text=label,
                command=lambda l=label, c=command: self._activate_sidebar_item(l, c),
                anchor="w",
                relief="flat",
                bd=0,
                padx=10,
                pady=9,
                bg=self.ui_colors["control_bg"],
                fg=self.ui_colors["muted"],
                activebackground=self.ui_colors["accent_soft"],
                activeforeground=self.ui_colors["fg"],
                cursor="hand2",
                font=(FONT_FAMILY, 10, "bold"),
            )
            btn.grid(row=0, column=1, sticky="ew")
            btn.bind("<Enter>", lambda _e, l=label: self._on_sidebar_hover(l, True))
            btn.bind("<Leave>", lambda _e, l=label: self._on_sidebar_hover(l, False))
            self.sidebar_buttons[label] = (row, accent, btn)
        self._refresh_sidebar_state()

        workspace = ttk.Frame(main)
        workspace.grid(row=0, column=1, sticky="nsew")
        workspace.columnconfigure(0, weight=3)
        workspace.columnconfigure(1, weight=2)
        workspace.rowconfigure(1, weight=1)

        topbar = ttk.Frame(workspace, style="Card.TFrame", padding=(SPACE_MD, SPACE_MD))
        topbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        topbar.columnconfigure(1, weight=1)

        left_tools = ttk.Frame(topbar)
        left_tools.grid(row=0, column=0, sticky="w", padx=(0, 12))
        logo_image = self._load_top_left_logo()
        if logo_image:
            self.logo_image_label = tk.Label(left_tools, image=logo_image, bg=self.ui_colors["control_bg"])
            self.logo_image_label.pack(side="left", padx=(0, 8))
        ttk.Label(left_tools, textvariable=self.app_title_var, style="Header.TLabel").pack(side="left")

        self.search_entry = tk.Entry(
            topbar,
            textvariable=self.search_var,
            width=44,
            bg="#ffffff" if self.current_theme == "light" else "#0f1b33",
            fg="#111827" if self.current_theme == "light" else "#e8eefc",
            insertbackground="#111827" if self.current_theme == "light" else "#e8eefc",
            relief="flat",
            bd=0,
            highlightthickness=2,
            highlightbackground=self.ui_colors["accent_soft"],
            highlightcolor=self.ui_colors["accent"],
            font=(FONT_FAMILY, 11),
        )
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12), ipady=6)

        right_tools = ttk.Frame(topbar)
        right_tools.grid(row=0, column=2, sticky="e")
        self.new_task_btn = tk.Button(
            right_tools,
            text="New Contract",
            command=self.new_task,
            relief="flat",
            bd=0,
            padx=14,
            pady=9,
            bg=self.ui_colors["accent"],
            fg="#ffffff",
            activebackground="#2f6fe4",
            activeforeground="#ffffff",
            cursor="hand2",
            font=(FONT_FAMILY, 10, "bold"),
        )
        self.new_task_btn.pack(side="left", padx=(0, 6))
        export_btn = tk.Button(
            right_tools,
            text="Export",
            command=self.open_export_popup,
            relief="flat",
            bd=0,
            padx=12,
            pady=9,
            bg=self.ui_colors["control_bg"],
            fg=self.ui_colors["fg"],
            activebackground=self.ui_colors["accent_soft"],
            activeforeground=self.ui_colors["fg"],
            cursor="hand2",
        )
        export_btn.pack(side="left", padx=(0, 6))
        settings_btn = tk.Button(
            right_tools,
            text="Settings",
            command=self.open_settings_popup,
            relief="flat",
            bd=0,
            padx=10,
            pady=8,
            bg=self.ui_colors["control_bg"],
            fg=self.ui_colors["muted"],
            activebackground=self.ui_colors["accent_soft"],
            activeforeground=self.ui_colors["fg"],
            cursor="hand2",
        )
        settings_btn.pack(side="left", padx=(0, 10))

        self.profile_menu_btn = tk.Button(
            right_tools,
            text=f"{(self.current_user.get('profile_icon','👤') if self.current_user else '👤')} Profile ▾",
            command=self.open_profile_menu,
            relief="flat",
            bd=0,
            padx=10,
            pady=8,
            bg=self.ui_colors["control_bg"],
            fg=self.ui_colors["fg"],
            activebackground=self.ui_colors["accent_soft"],
            activeforeground=self.ui_colors["fg"],
            cursor="hand2",
            font=(FONT_FAMILY, 10, "bold"),
        )
        self.profile_menu_btn.pack(side="left", padx=(0, 2))

        self.user_badge_var = tk.StringVar(value="Not signed in")

        left = ttk.Frame(workspace, style="Card.TFrame", padding=SPACE_MD)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        right = ttk.Frame(workspace, style="RightPanel.TFrame", padding=SPACE_MD)
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        list_filters = ttk.Frame(left)
        list_filters.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Combobox(list_filters, values=["ALL"] + STATUS_OPTIONS, textvariable=self.status_filter, width=12, state="readonly").pack(side="left", padx=(0, 6))
        self.brand_filter_combo = ttk.Combobox(list_filters, values=["ALL"], textvariable=self.brand_filter, width=16, state="readonly")
        self.brand_filter_combo.pack(side="left", padx=(0, 6))
        ttk.Label(list_filters, textvariable=self.runtime_badge_var, style="Muted.TLabel").pack(side="right")

        cols = ("ID", "Brand", "Status", "Amount", "Type", "Vendors", "Created")
        task_table_card = ttk.Frame(left, style="Card.TFrame", padding=8)
        task_table_card.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        task_table_card.columnconfigure(0, weight=1)
        task_table_card.rowconfigure(0, weight=1)

        self.task_tree = ttk.Treeview(task_table_card, columns=cols, show="headings", selectmode="extended")
        self.task_tree.heading("ID", text="ID")
        self.task_tree.heading("Brand", text="Brand")
        self.task_tree.heading("Status", text="Status")
        self.task_tree.heading("Amount", text="Amount")
        self.task_tree.heading("Type", text="Type")
        self.task_tree.heading("Vendors", text="Vendors")
        self.task_tree.heading("Created", text="Created")
        self.task_tree.column("ID", width=82, anchor="w")
        self.task_tree.column("Brand", width=180, anchor="w")
        self.task_tree.column("Status", width=110, anchor="center")
        self.task_tree.column("Amount", width=110, anchor="e")
        self.task_tree.column("Type", width=110, anchor="w")
        self.task_tree.column("Vendors", width=90, anchor="center")
        self.task_tree.column("Created", width=120, anchor="w")
        self.task_tree.grid(row=0, column=0, sticky="nsew")
        self._apply_task_tree_status_colors()

        ttk.Separator(left, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=(8, 6))
        sign_in_status = ttk.Frame(left)
        sign_in_status.grid(row=3, column=0, sticky="ew", pady=(2, 0))
        ttk.Label(sign_in_status, textvariable=self.user_badge_var, foreground="#6f7a8d").pack(side="left")
        ttk.Label(sign_in_status, text="•", foreground="#6f7a8d").pack(side="left", padx=8)
        ttk.Label(sign_in_status, textvariable=self.runtime_badge_var, foreground="#6f7a8d").pack(side="left")

        self.editor_empty_state = ttk.Frame(right, style="Card.TFrame", padding=28)
        self.editor_empty_state.grid(row=0, column=0, rowspan=3, sticky="nsew")
        center = ttk.Frame(self.editor_empty_state, style="Card.TFrame", padding=26)
        center.place(relx=0.5, rely=0.45, anchor="center")
        ttk.Label(center, text="📄", font=("Segoe UI Emoji", 26), style="SubHeader.TLabel").pack(pady=(0, 8))
        ttk.Label(center, text="Select a contract or create a new one", style="SubHeader.TLabel").pack(anchor="center", pady=(0, 6))
        ttk.Label(center, text="Choose a row on the left, or use New Contract to get started.", style="Muted.TLabel").pack(anchor="center")
        tk.Button(
            center,
            text="➕ New Contract",
            command=self.new_task,
            relief="flat",
            bd=0,
            padx=SPACE_MD,
            pady=SPACE_SM,
            bg=self.ui_colors["accent"],
            fg="#ffffff",
            activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="#ffffff",
            cursor="hand2",
            font=(FONT_FAMILY, 10, "bold"),
        ).pack(pady=(SPACE_MD, 0))

        editor_header = ttk.Frame(right)
        editor_header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(editor_header, text="Editor", style="SubHeader.TLabel").pack(anchor="w")
        ttk.Label(editor_header, text="Client details and vendor line items", style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        info = ttk.LabelFrame(right, text="Client Info", padding=16)
        info.grid(row=1, column=0, sticky="ew")
        for c in range(10):
            info.columnconfigure(c, weight=1)

        ttk.Button(info, text="Client", command=self.show_client_panel).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(2, 8))
        ttk.Button(info, text="Vendor", command=self.show_vendor_panel).grid(row=0, column=1, sticky="w", padx=(0, 12), pady=(2, 8))

        self.brand_combo = SearchableCombo(info, [], textvariable=self.brand_var, width=20)
        self.brand_combo.grid(row=0, column=2, sticky="ew", padx=(0, 12), pady=(2, 8))
        ttk.Label(info, text="Total Price (Auto)").grid(row=0, column=3, sticky="e", padx=(0, 8), pady=(2, 8))
        ttk.Label(info, textvariable=self.amount_var, font=(FONT_FAMILY, 10, "bold")).grid(row=0, column=4, sticky="w", padx=(4, 12), pady=(2, 8))
        ttk.Combobox(info, values=STATUS_OPTIONS, textvariable=self.status_var, width=12, state="readonly").grid(row=0, column=5, padx=(0, 10), pady=(2, 8))
        ttk.Label(info, text="Template").grid(row=0, column=6, sticky="e", padx=(6, 2))
        self.template_combo = ttk.Combobox(info, textvariable=self.template_key_var, width=22, state="readonly")
        self.template_combo.grid(row=0, column=7, columnspan=3, sticky="w")

        btns = ttk.Frame(info)
        btns.grid(row=1, column=0, columnspan=10, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Add Template", command=self.add_template_popup).pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="Manage Templates", command=self.manage_templates_popup).pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="Save", command=self.save_task).pack(side="left", padx=(0, 6))

        self.vendor_panel = ttk.Frame(right)
        self.vendor_panel.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        self.vendor_panel.columnconfigure(0, weight=1)
        self.vendor_panel.rowconfigure(0, weight=1)

        self.client_panel = ttk.Frame(right)
        self.client_panel.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        self.client_panel.columnconfigure(0, weight=1)
        ttk.Label(self.client_panel, text="Client subtasks panel (you said you will do this later)", foreground="#666").grid(row=0, column=0, sticky="nw")

        sub_frame = ttk.LabelFrame(self.vendor_panel, text="Vendor Subtasks", padding=14)
        sub_frame.grid(row=0, column=0, sticky="nsew")
        sub_frame.columnconfigure(0, weight=1)
        sub_frame.rowconfigure(0, weight=1)

        sub_cols = ("Vendor", "Channel", "Platforms", "Type", "Qty", "Price", "Details")
        self.sub_tree = ttk.Treeview(sub_frame, columns=sub_cols, show="headings", height=8, selectmode="extended")
        for c in sub_cols:
            self.sub_tree.heading(c, text=c)
            self.sub_tree.column(c, width=100, anchor="center")
        self.sub_tree.grid(row=0, column=0, sticky="nsew")

        sub_btns = ttk.Frame(sub_frame)
        sub_btns.grid(row=1, column=0, sticky="e", pady=(6, 0))
        ttk.Label(sub_btns, text="Tip: Right-click vendor rows for quick delete", foreground="#707070").pack(side="right")

        editor = ttk.LabelFrame(self.vendor_panel, text="Vendor + Bank", padding=14)
        editor.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        for c in range(10):
            editor.columnconfigure(c, weight=1)

        top_form = ttk.Frame(editor)
        top_form.grid(row=0, column=0, columnspan=10, sticky="ew")
        for c in range(6):
            top_form.columnconfigure(c, weight=1)

        ttk.Label(top_form, text="Name").grid(row=0, column=0, sticky="w")
        self.vendor_combo = SearchableCombo(top_form, FALLBACK_VENDORS, textvariable=self.vendor_var, width=18)
        self.vendor_combo.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(top_form, text="Platforms").grid(row=0, column=1, sticky="w")
        platform_box = ttk.Frame(top_form)
        platform_box.grid(row=1, column=1, sticky="w", padx=(0, 8))
        for i, p in enumerate(["TikTok", "Instagram", "Snapchat", "YouTube"]):
            v = tk.BooleanVar()
            self.platform_vars[p] = v
            self.platform_channel_vars[p] = tk.StringVar()
            r, c = divmod(i, 2)
            ttk.Checkbutton(platform_box, text=p, variable=v, style="Platform.TCheckbutton", command=self.refresh_platform_detail_inputs).grid(row=r, column=c, sticky="w", padx=(0, 8), pady=1)

        ttk.Label(top_form, text="Visit Forms").grid(row=0, column=2, sticky="w")
        ad_combo = ttk.Combobox(top_form, values=["Store Visit", "Home Ad", "Multi Service"], textvariable=self.type_sub_var, width=15, state="readonly")
        ad_combo.grid(row=1, column=2, sticky="ew", padx=(0, 8))

        ttk.Label(top_form, text="Qty").grid(row=0, column=3, sticky="w")
        ttk.Entry(top_form, textvariable=self.qty_var, width=8).grid(row=1, column=3, sticky="ew", padx=(0, 8))

        ttk.Label(top_form, text="Details").grid(row=0, column=4, sticky="w")
        self.detail_entry = ttk.Entry(top_form, textvariable=self.detail_var, width=20)
        self.detail_entry.grid(row=1, column=4, sticky="ew", padx=(0, 8))

        self.add_vendor_btn = ttk.Button(top_form, text="Add Vendor", command=self.add_or_update_subtask)
        self.add_vendor_btn.grid(row=1, column=5, sticky="e")

        ttk.Label(editor, text="Platform Details").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.platform_detail_box = ttk.Frame(editor)
        self.platform_detail_box.grid(row=2, column=0, columnspan=10, sticky="ew")
        self.platform_detail_box.columnconfigure(0, weight=1)
        self.refresh_platform_detail_inputs()

        bank_top = ttk.Frame(editor)
        bank_top.grid(row=3, column=0, columnspan=10, sticky="ew", pady=(10, 0))
        for c in range(5):
            bank_top.columnconfigure(c, weight=1)

        ttk.Label(bank_top, text="License").grid(row=0, column=0, sticky="w")
        self.license_combo = SearchableCombo(bank_top, [], textvariable=self.license_var, width=18)
        self.license_combo.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(bank_top, text="IBAN").grid(row=0, column=1, sticky="w")
        self.iban_combo = ttk.Combobox(bank_top, textvariable=self.iban_var, width=22, state="readonly")
        self.iban_combo.grid(row=1, column=1, sticky="ew", padx=(0, 8))

        ttk.Label(bank_top, text="Vendor Price").grid(row=0, column=2, sticky="w")
        ttk.Entry(bank_top, textvariable=self.vendor_price_var, width=10).grid(row=1, column=2, sticky="ew", padx=(0, 8))

        ttk.Button(bank_top, text="Vendor Master Popup", command=self.open_vendor_master_popup).grid(row=1, column=3, sticky="w")

        bank_bottom = ttk.Frame(editor)
        bank_bottom.grid(row=4, column=0, columnspan=10, sticky="ew", pady=(10, 0))
        for c in range(4):
            bank_bottom.columnconfigure(c, weight=1)

        ttk.Label(bank_bottom, text="Bank Name").grid(row=0, column=0, sticky="w")
        ttk.Entry(bank_bottom, textvariable=self.bank_name_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))

        ttk.Label(bank_bottom, text="Account Name").grid(row=0, column=1, sticky="w")
        ttk.Entry(bank_bottom, textvariable=self.account_name_var).grid(row=1, column=1, sticky="ew", padx=(0, 8))

        ttk.Label(bank_bottom, text="Account #").grid(row=0, column=2, sticky="w")
        ttk.Entry(bank_bottom, textvariable=self.account_number_var).grid(row=1, column=2, sticky="ew", padx=(0, 8))

        ttk.Label(bank_bottom, text="SWIFT").grid(row=0, column=3, sticky="w")
        ttk.Entry(bank_bottom, textvariable=self.swift_code_var).grid(row=1, column=3, sticky="ew")

        ad_combo.bind("<<ComboboxSelected>>", self.toggle_details)
        self.toggle_details()

    def _activate_sidebar_item(self, label, command):
        self.sidebar_active.set(label)
        self._refresh_sidebar_state()
        if callable(command):
            command()

    def _refresh_sidebar_state(self):
        if not hasattr(self, "sidebar_buttons"):
            return
        active = self.sidebar_active.get()
        for label, (row, accent, btn) in self.sidebar_buttons.items():
            is_active = label == active
            accent.configure(bg=self.ui_colors["accent"] if is_active else self.ui_colors["control_bg"])
            btn.configure(
                bg=self.ui_colors["accent_soft"] if is_active else self.ui_colors["control_bg"],
                fg=self.ui_colors["fg"] if is_active else self.ui_colors["muted"],
            )

    def _on_sidebar_hover(self, label, entering: bool):
        if not hasattr(self, "sidebar_buttons") or label not in self.sidebar_buttons:
            return
        if self.sidebar_active.get() == label:
            return
        _row, _accent, btn = self.sidebar_buttons[label]
        btn.configure(bg=self.ui_colors["accent_soft"] if entering else self.ui_colors["control_bg"])

    def _status_badge_text(self, status: str):
        badge = {
            "NEW": "[NEW]",
            "DRAFT": "[DRAFT]",
            "SENT": "[SENT]",
            "SIGNED": "[SIGNED]",
            "PAID": "[PAID]",
        }
        return badge.get((status or "").strip().upper(), f"[{status}]" if status else "[UNKNOWN]")

    def _set_editor_empty_state(self, show: bool):
        if not hasattr(self, "editor_empty_state"):
            return
        if show:
            self.editor_empty_state.lift()
        else:
            self.editor_empty_state.lower()

    def show_vendor_panel(self):
        self.client_panel.grid_remove()
        self.vendor_panel.grid()

    def show_client_panel(self):
        self.vendor_panel.grid_remove()
        self.client_panel.grid()


    def _load_top_left_logo(self):
        for logo_name in LOGO_CANDIDATES:
            logo_path = Path(__file__).with_name(logo_name)
            if logo_path.exists():
                try:
                    img = tk.PhotoImage(file=str(logo_path))
                    if img.width() > 52:
                        img = img.subsample(max(1, img.width() // 52))
                    if img.height() > 52:
                        img = img.subsample(1, max(1, img.height() // 52))
                    self.logo_image = img
                    return img
                except tk.TclError:
                    continue
        return None

    def open_asana_dashboard(self):
        if self.dashboard_window and self.dashboard_window.winfo_exists():
            self.dashboard_window.lift()
            self.refresh_asana_dashboard()
            return

        self.dashboard_window = tk.Toplevel(self.root)
        self.dashboard_window.title(f"{APP_BRAND} • Dashboard")
        self.dashboard_window.configure(bg="#0b1020")
        self._lock_popup_position(self.dashboard_window, 1360, 820)
        self._animate_popup_in(self.dashboard_window)

        shell = tk.Frame(self.dashboard_window, bg="#0b1020")
        shell.pack(fill="both", expand=True)
        shell.grid_columnconfigure(1, weight=4)
        shell.grid_columnconfigure(2, weight=2)
        shell.grid_rowconfigure(1, weight=1)

        header = tk.Frame(shell, bg="#121a30", height=72)
        header.grid(row=0, column=0, columnspan=3, sticky="nsew")
        header.grid_columnconfigure(1, weight=1)
        tk.Label(header, text=f"{APP_BRAND} Workspace Dashboard", bg="#121a30", fg="#ffffff", font=("Segoe UI", 20, "bold")).grid(row=0, column=0, padx=16, pady=12, sticky="w")
        self.dashboard_search_var.set("")
        tk.Entry(header, textvariable=self.dashboard_search_var, bg="#0c1730", fg="#dce7ff", insertbackground="#dce7ff", relief="flat", font=("Segoe UI", 11), width=62, bd=0, highlightthickness=1, highlightbackground="#29406d", highlightcolor="#4d83ff").grid(row=0, column=1, padx=12, pady=14, sticky="ew")
        tk.Button(header, text="Refresh", command=self.refresh_asana_dashboard, bg="#2d6bff", fg="#ffffff", relief="flat", padx=16, pady=6, bd=0, activebackground="#4d83ff", activeforeground="#ffffff", cursor="hand2").grid(row=0, column=2, padx=16, pady=14)

        nav = tk.Frame(shell, bg="#0f1730", width=220)
        nav.grid(row=1, column=0, sticky="nsew")
        nav.grid_propagate(False)
        tk.Label(nav, text="Views", bg="#0f1730", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 10))

        self.dashboard_view_var = tk.StringVar(value="Overview")
        self.dashboard_nav_buttons = {}
        for item in ("Overview", "List", "Timeline", "Charts", "Progress", "Dashboard", "Calendar"):
            btn = tk.Button(
                nav,
                text=item,
                command=lambda i=item: self._set_dashboard_view(i),
                bg="#172444",
                fg="#ffffff",
                activebackground="#213968",
                activeforeground="#ffffff",
                relief="flat",
                anchor="w",
                padx=14,
                pady=8,
                bd=0,
                cursor="hand2",
            )
            btn.pack(fill="x", padx=12, pady=3)
            self.dashboard_nav_buttons[item] = btn

        center = tk.Frame(shell, bg="#0b1020")
        center.grid(row=1, column=1, sticky="nsew", padx=(10, 8), pady=(10, 10))
        center.grid_rowconfigure(1, weight=1)
        center.grid_columnconfigure(0, weight=1)

        self.dashboard_main = tk.Frame(center, bg="#0b1020")
        self.dashboard_main.grid(row=0, column=0, sticky="nsew")
        self.dashboard_main.grid_rowconfigure(0, weight=1)
        self.dashboard_main.grid_columnconfigure(0, weight=1)

        right = tk.Frame(shell, bg="#121b33")
        right.grid(row=1, column=2, sticky="nsew", padx=(0, 10), pady=(10, 10))
        right.grid_rowconfigure(4, weight=1)
        right.grid_columnconfigure(0, weight=1)
        tk.Label(right, text="Details", bg="#121b33", fg="#ffffff", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        self.dashboard_detail_title = tk.Label(right, text="Pick a contract ID", bg="#121b33", fg="#b6b6c3", font=("Segoe UI", 12))
        self.dashboard_detail_title.grid(row=1, column=0, sticky="w", padx=12)
        self.dashboard_detail_stats = tk.Label(right, text="", bg="#121b33", fg="#ffffff", justify="left", font=("Segoe UI", 12))
        self.dashboard_detail_stats.grid(row=2, column=0, sticky="w", padx=12, pady=(8, 10))
        tk.Label(right, text="Generated contracts", bg="#121b33", fg="#ffffff", font=("Segoe UI", 13, "bold")).grid(row=3, column=0, sticky="w", padx=12)
        self.dashboard_contracts = tk.Listbox(right, bg="#0c1327", fg="#ffffff", selectbackground="#2b64ff", height=14, borderwidth=0, highlightthickness=0)
        self.dashboard_contracts.grid(row=4, column=0, sticky="nsew", padx=12, pady=(8, 12))

        self.dashboard_search_var.trace_add("write", lambda *_: self.refresh_asana_dashboard())
        self.refresh_asana_dashboard()

    def _set_dashboard_view(self, view_name: str):
        self.dashboard_view_var.set(view_name)
        self.refresh_asana_dashboard()

    def _refresh_dashboard_nav_styles(self):
        current = self.dashboard_view_var.get()
        for name, btn in getattr(self, "dashboard_nav_buttons", {}).items():
            active = name == current
            btn.configure(bg="#2b3c63" if active else "#172444")

    def _filter_dashboard_rows(self):
        query = self.dashboard_search_var.get().strip().lower() if hasattr(self, "dashboard_search_var") else ""
        rows = list(self.db.list_tasks())
        if query:
            rows = [
                r for r in rows
                if query in r["id"].lower()
                or query in (r["brand"] or "").lower()
                or query in (r["status"] or "").lower()
                or query in (r["contract_type"] or "").lower()
                or query in (r["created_date"] or "").lower()
            ]
        return rows

    def refresh_asana_dashboard(self):
        rows = self._filter_dashboard_rows()
        self.dashboard_rows = {row["id"]: row for row in rows}
        self._refresh_dashboard_nav_styles()
        self._render_dashboard_view(rows)
        if rows:
            self._update_dashboard_details_from_task(rows[0]["id"])

    def _render_dashboard_view(self, rows):
        for child in self.dashboard_main.winfo_children():
            child.destroy()

        view = self.dashboard_view_var.get().lower()
        if view in {"overview", "dashboard"}:
            self._render_overview_panel(rows)
        elif view == "list":
            self._render_list_panel(rows)
        elif view == "timeline":
            self._render_timeline_panel(rows)
        elif view == "calendar":
            self._render_calendar_panel(rows)
        elif view == "charts":
            self._render_analytics_panel(rows)
        elif view == "progress":
            self._render_board_panel(rows)
        else:
            self._render_overview_panel(rows)

    def _draw_contract_table(self, parent, rows, height=11):
        box = tk.Frame(parent, bg="#151a22")
        box.pack(fill="both", expand=True)
        cols = ("ID", "Brand", "Status", "Type", "Amount", "Vendors", "Created")
        tree = ttk.Treeview(box, columns=cols, show="headings", height=height)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, anchor="w", width=120)
        tree.pack(fill="both", expand=True)
        for row in rows:
            tree.insert("", "end", iid=row["id"], values=(row["id"], row["brand"], row["status"], row["contract_type"], row["amount"], row["vendor_count"], row["created_date"]))
        tree.bind("<<TreeviewSelect>>", self.on_dashboard_task_select)
        self.dashboard_tree = tree
        if rows:
            tree.selection_set(rows[0]["id"])
        return tree

    def _render_overview_panel(self, rows):
        top = tk.Frame(self.dashboard_main, bg="#0b1020")
        top.pack(fill="both", expand=True)

        tk.Label(top, text="Contract list", bg="#0b1020", fg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))
        table_wrap = tk.Frame(top, bg="#0b1020")
        table_wrap.pack(fill="x")
        self._draw_contract_table(table_wrap, rows, height=10)

        kpi = tk.Label(top, text="", bg="#121b33", fg="#ffffff", justify="left", anchor="w", font=("Segoe UI", 12), padx=10, pady=8)
        kpi.pack(fill="x", pady=(8, 8))

        charts_row = tk.Frame(top, bg="#0b1020")
        charts_row.pack(fill="both", expand=True)
        left = tk.Canvas(charts_row, width=420, height=220, bg="#121b33", highlightthickness=0)
        left.pack(side="left", fill="both", expand=True, padx=(0, 4))
        right = tk.Canvas(charts_row, width=420, height=220, bg="#121b33", highlightthickness=0)
        right.pack(side="left", fill="both", expand=True, padx=(4, 0))

        timeline = tk.Canvas(top, width=860, height=190, bg="#121b33", highlightthickness=0)
        timeline.pack(fill="x", pady=(8, 0))

        self._render_dashboard_charts(rows, kpi_label=kpi, pie_canvas=left, bar_canvas=right, timeline_canvas=timeline)

    def _render_list_panel(self, rows):
        wrap = tk.Frame(self.dashboard_main, bg="#0b1020")
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="List view", bg="#0b1020", fg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))

        cols = ("Name", "Status", "Created", "Sales", "Brand", "Amount", "Platform", "AD Type", "Vendor")
        tree = ttk.Treeview(wrap, columns=cols, show="headings", height=18)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, anchor="w", width=120)
        tree.pack(fill="both", expand=True)

        for row in rows:
            subs = self.db.list_subtasks(row["id"])
            if not subs:
                tree.insert("", "end", iid=row["id"], values=(row["id"], row["status"], row["created_date"], "-", row["brand"], row["amount"], "-", "-", "-"))
                continue
            for i, sub in enumerate(subs):
                iid = f"{row['id']}::{i}"
                tree.insert("", "end", iid=iid, values=(row["id"], row["status"], row["created_date"], "-", row["brand"], row["amount"], sub["platforms"], sub["ad_type"], sub["vendor"]))

        def _pick(_event=None):
            sel = tree.selection()
            if not sel:
                return
            task_id = sel[0].split("::")[0]
            self._update_dashboard_details_from_task(task_id)

        tree.bind("<<TreeviewSelect>>", _pick)

    def _render_board_panel(self, rows):
        wrap = tk.Frame(self.dashboard_main, bg="#0b1020")
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="Progress board", bg="#0b1020", fg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))

        sc = tk.Canvas(wrap, bg="#0b1020", highlightthickness=0)
        sc.pack(fill="both", expand=True)
        content = tk.Frame(sc, bg="#0b1020")
        sid = sc.create_window((0, 0), window=content, anchor="nw")

        statuses = ["NEW", "DRAFT", "SENT", "SIGNED", "PAID"]
        grouped = {s: [r for r in rows if r["status"] == s] for s in statuses}
        for ci, status in enumerate(statuses):
            col = tk.Frame(content, bg="#1a1f29", width=220)
            col.grid(row=0, column=ci, padx=6, sticky="n")
            tk.Label(col, text=f"{status}  {len(grouped[status])}", bg="#1a1f29", fg="#ffffff", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(10, 8))
            for row in grouped[status][:10]:
                card = tk.Frame(col, bg="#222938", padx=8, pady=6)
                card.pack(fill="x", padx=8, pady=4)
                tk.Label(card, text=f"{row['id']} • {row['brand'] or '-'}", bg="#222938", fg="#ffffff", anchor="w").pack(fill="x")
                tk.Label(card, text=f"{row['amount']} SAR", bg="#222938", fg="#a8bad9", anchor="w", font=("Segoe UI", 9)).pack(fill="x")
                card.bind("<Button-1>", lambda _e, tid=row["id"]: self._update_dashboard_details_from_task(tid))

        def _sync(_event=None):
            sc.configure(scrollregion=sc.bbox("all"))
            sc.itemconfigure(sid, width=sc.winfo_width())
        content.bind("<Configure>", _sync)

    def _render_timeline_panel(self, rows):
        wrap = tk.Frame(self.dashboard_main, bg="#0b1020")
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="Timeline", bg="#0b1020", fg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))

        canvas = tk.Canvas(wrap, bg="#121b33", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        dates = sorted({(r["created_date"] or "-")[:10] for r in rows})
        if not dates:
            canvas.create_text(40, 40, text="No timeline data", fill="#ffffff", anchor="nw")
            return

        lane_h = 38
        for idx, date in enumerate(dates[:18]):
            y = 40 + idx * lane_h
            canvas.create_text(20, y + 14, text=date[5:], fill="#8fb3ff", anchor="w")
            day_rows = [r for r in rows if (r["created_date"] or "")[:10] == date]
            for j, row in enumerate(day_rows[:8]):
                x = 120 + j * 150
                canvas.create_rectangle(x, y, x + 140, y + 26, fill="#223b6f", outline="#33538f")
                canvas.create_text(x + 6, y + 13, text=f"{row['id']} {row['brand'] or ''}"[:20], fill="#ffffff", anchor="w")
                canvas.tag_bind("all", "<Button-1>", lambda _e, tid=row["id"]: self._update_dashboard_details_from_task(tid))

    def _render_calendar_panel(self, rows):
        wrap = tk.Frame(self.dashboard_main, bg="#0b1020")
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="Calendar", bg="#0b1020", fg="#ffffff", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))

        canvas = tk.Canvas(wrap, bg="#121b33", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        cols, rows_count = 7, 5
        w, h = 200, 120
        week = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
        for i, d in enumerate(week):
            canvas.create_text(12 + i*w, 18, text=d, fill="#8fb3ff", anchor="nw", font=(FONT_FAMILY, 10, "bold"))
        task_index = 0
        for r in range(rows_count):
            for c in range(cols):
                x1, y1 = 10 + c*w, 36 + r*h
                x2, y2 = x1 + w - 8, y1 + h - 8
                canvas.create_rectangle(x1, y1, x2, y2, outline="#2e3a52", fill="#101a31")
                canvas.create_text(x1 + 8, y1 + 6, text=str((r*cols+c+1)), fill="#8f9fbe", anchor="nw")
                if task_index < len(rows):
                    row = rows[task_index]
                    canvas.create_rectangle(x1 + 8, y1 + 28, x2 - 8, y1 + 56, fill="#2d6bff", outline="")
                    canvas.create_text(x1 + 12, y1 + 42, text=f"{row['id']} {row['brand'] or ''}"[:20], fill="#ffffff", anchor="w")
                    task_index += 1

    def _render_analytics_panel(self, rows):
        wrap = tk.Frame(self.dashboard_main, bg="#0b1020")
        wrap.pack(fill="both", expand=True)

        cards = tk.Frame(wrap, bg="#0b1020")
        cards.pack(fill="x", pady=(0, 8))
        stats = self._collect_dashboard_stats(rows)
        cards_data = [
            ("Total completed tasks", str(stats["done"])),
            ("Total incomplete tasks", str(max(0, stats["total"] - stats["done"]))),
            ("Total overdue tasks", str(stats["legal"])),
            ("Total tasks", str(stats["total"])),
        ]
        for title, value in cards_data:
            card = tk.Frame(cards, bg="#13203d", padx=16, pady=12)
            card.pack(side="left", fill="x", expand=True, padx=6)
            tk.Label(card, text=title, bg="#13203d", fg="#e6eaf3", font=(FONT_FAMILY, 11, "bold")).pack(anchor="w")
            tk.Label(card, text=value, bg="#13203d", fg="#ffffff", font=("Segoe UI", 28)).pack(anchor="w", pady=(10, 0))

        body = tk.Frame(wrap, bg="#0b1020")
        body.pack(fill="both", expand=True)
        left = tk.Canvas(body, width=500, height=280, bg="#13203d", highlightthickness=0)
        left.pack(side="left", fill="both", expand=True, padx=(0, 4))
        right = tk.Canvas(body, width=500, height=280, bg="#13203d", highlightthickness=0)
        right.pack(side="left", fill="both", expand=True, padx=(4, 0))
        bottom = tk.Canvas(wrap, width=1000, height=220, bg="#13203d", highlightthickness=0)
        bottom.pack(fill="x", pady=(8, 0))
        kpi = tk.Label(wrap, text="", bg="#0b1020", fg="#ffffff")

        self._render_dashboard_charts(rows, kpi_label=kpi, pie_canvas=left, bar_canvas=right, timeline_canvas=bottom)

    def _collect_dashboard_stats(self, rows):
        status_counts = {}
        for row in rows:
            status_counts[row["status"]] = status_counts.get(row["status"], 0) + 1
        return {
            "total": len(rows),
            "paid": status_counts.get("PAID", 0),
            "sent": status_counts.get("SENT", 0),
        }

    def _render_dashboard_charts(self, rows, kpi_label, pie_canvas, bar_canvas, timeline_canvas):
        total_contracts = len(rows)
        total_amount = sum(self._parse_price(r["amount"]) for r in rows)
        status_counts = {}
        brand_amounts = {}
        timeline_counts = {}
        for r in rows:
            status_counts[r["status"]] = status_counts.get(r["status"], 0) + 1
            brand = r["brand"] or "Unknown"
            brand_amounts[brand] = brand_amounts.get(brand, 0.0) + self._parse_price(r["amount"])
            date_key = (r["created_date"] or "-")[:10]
            timeline_counts[date_key] = timeline_counts.get(date_key, 0) + 1

        kpi_label.configure(text=f"Contracts: {total_contracts}    Total Spent: {total_amount:.2f}\nProjects: {len(brand_amounts)}    Status buckets: {len(status_counts)}")

        pie_canvas.delete("all")
        pie_canvas.create_text(12, 12, text="Status Split (Pie)", anchor="nw", fill="#ffffff", font=("Segoe UI", 12, "bold"))
        colors = ["#7ea7e2", "#8fd8c3", "#f2cb7a", "#c291ef", "#f18a8a", "#9ea8bd"]
        cx, cy, rad = 118, 120, 74
        start = 0
        total = max(1, sum(status_counts.values()))
        for i, (status, count) in enumerate(sorted(status_counts.items())):
            extent = 360 * count / total
            pie_canvas.create_arc(cx-rad, cy-rad, cx+rad, cy+rad, start=start, extent=extent, fill=colors[i % len(colors)], outline="#13203d")
            pie_canvas.create_text(250, 44 + i*22, text=f"{status}: {count}", anchor="w", fill="#d8dce6", font=(FONT_FAMILY, 10))
            start += extent

        bar_canvas.delete("all")
        bar_canvas.create_text(12, 12, text="Top Projects by Spend", anchor="nw", fill="#ffffff", font=("Segoe UI", 12, "bold"))
        top = sorted(brand_amounts.items(), key=lambda x: x[1], reverse=True)[:6]
        max_amt = max([v for _, v in top], default=1.0)
        for i, (brand, amount) in enumerate(top):
            y = 48 + i * 28
            w = int((amount / max_amt) * 240)
            bar_canvas.create_text(12, y+9, text=brand[:16], anchor="w", fill="#dce5f7", font=(FONT_FAMILY, 10))
            bar_canvas.create_rectangle(160, y, 160+w, y+18, fill="#6ca6f0", outline="")
            bar_canvas.create_text(430, y+9, text=f"{amount:.0f}", anchor="e", fill="#ffffff", font=(FONT_FAMILY, 10))

        timeline_canvas.delete("all")
        timeline_canvas.create_text(12, 12, text="Timeline (Contracts by Date)", anchor="nw", fill="#ffffff", font=("Segoe UI", 12, "bold"))
        ordered = sorted(timeline_counts.items())[-10:]
        if ordered:
            max_c = max(c for _, c in ordered)
            x0, y0, step = 40, 160, 68
            pts = []
            for i, (d, c) in enumerate(ordered):
                x = x0 + i * step
                y = y0 - int((c / max_c) * 110)
                pts.extend([x, y])
                timeline_canvas.create_oval(x-4, y-4, x+4, y+4, fill="#6ad2bf", outline="")
                timeline_canvas.create_text(x, 182, text=d[5:], fill="#b8c2d6", font=("Segoe UI", 9))
            if len(pts) >= 4:
                timeline_canvas.create_line(*pts, fill="#6ad2bf", width=3, smooth=True)

    def _update_dashboard_details_from_task(self, task_id):
        row = self.dashboard_rows.get(task_id)
        if not row:
            return

        subtasks = self.db.list_subtasks(task_id)
        unique_vendors = sorted({s["vendor"] for s in subtasks if s["vendor"].strip()})
        self.dashboard_detail_title.configure(text=f"{task_id} • {row['brand'] or 'No brand'}")
        self.dashboard_detail_stats.configure(
            text=(
                f"Status: {row['status']}\n"
                f"Type: {row['contract_type']}\n"
                f"Amount: {row['amount'] or '-'}\n"
                f"Subtasks: {len(subtasks)}\n"
                f"Vendors: {', '.join(unique_vendors[:5]) or '-'}"
            )
        )

        self.dashboard_contracts.delete(0, tk.END)
        contracts = self.db.list_generated_contracts_for_task(task_id)
        if not contracts:
            self.dashboard_contracts.insert(tk.END, "No generated contracts yet")
        for c in contracts:
            self.dashboard_contracts.insert(tk.END, f"{c['generated_at']}  •  {c['contract_id']}")

    def on_dashboard_task_select(self, _event=None):
        if not getattr(self, "dashboard_tree", None):
            return
        selected = self.dashboard_tree.selection()
        if not selected:
            return
        self._update_dashboard_details_from_task(selected[0])

    def _lock_popup_position(self, pop, width, height, x=None, y=None):
        pop.resizable(False, False)
        sw = pop.winfo_screenwidth()
        sh = pop.winfo_screenheight()
        x = (sw - width) // 2 if x is None else x
        y = (sh - height) // 2 if y is None else y
        pop.geometry(f"{width}x{height}+{x}+{y}")
        pop.update_idletasks()
        pop._locked_geometry = (width, height, x, y)

        def _enforce(_event=None):
            if not pop.winfo_exists():
                return
            w, h, px, py = pop._locked_geometry
            if (pop.winfo_x(), pop.winfo_y()) != (px, py):
                pop.geometry(f"{w}x{h}+{px}+{py}")

        pop.bind("<Configure>", _enforce)

    def _animate_popup_in(self, pop, steps=8):
        pop.attributes("-alpha", 0.0)
        def _step(i=0):
            if not pop.winfo_exists():
                return
            pop.attributes("-alpha", min(1.0, i / steps))
            if i < steps:
                pop.after(18, lambda: _step(i + 1))
        _step()

    def _quick_actions(self):
        # Keep this panel focused on actions not already present in the main sidebar/topbar.
        return [
            ("Profiles (Login / Signup)", self.open_auth_popup),
            ("Open Contract Maker", self.root.lift),
        ]

    def toggle_quick_actions_panel(self):
        if hasattr(self, "logo_button"):
            self._animate_widget_tap(self.logo_button, base_bg=self.logo_button.cget("bg"), tap_bg="#25252c")
        if self.quick_panel and self.quick_panel.winfo_exists():
            self.close_quick_actions_panel()
            return

        self.quick_panel_overlay = tk.Toplevel(self.root)
        self.quick_panel_overlay.overrideredirect(True)
        self.quick_panel_overlay.attributes("-alpha", 0.25)
        self.quick_panel_overlay.configure(bg="#000000")
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        self.quick_panel_overlay.geometry(f"{rw}x{rh}+{rx}+{ry}")
        self.quick_panel_overlay.lift(self.root)
        self.quick_panel_overlay.bind("<Button-1>", lambda _e: self.close_quick_actions_panel())

        self.quick_panel = tk.Toplevel(self.root)
        self.quick_panel.overrideredirect(True)
        self.quick_panel.configure(bg="#1f2430")
        self.quick_panel.lift(self.quick_panel_overlay)
        target_w = 320
        self.quick_panel.geometry(f"1x{rh}+{rx}+{ry}")

        shell = tk.Frame(self.quick_panel, bg="#1f2430", padx=18, pady=18)
        shell.pack(fill="both", expand=True)
        tk.Label(shell, text="Quick Actions", bg="#1f2430", fg="#f4f6fb", font=("Arial", 15, "bold")).pack(anchor="w", pady=(0, 12))
        for label, cmd in self._quick_actions():
            tk.Button(shell, text=label, command=lambda c=cmd: (self.close_quick_actions_panel(), c()), bg="#2d3444", fg="#f4f6fb", activebackground="#3b4356", activeforeground="#ffffff", relief="flat", anchor="w", padx=14, pady=8).pack(fill="x", pady=4)

        def _grow(i=1):
            if not self.quick_panel or not self.quick_panel.winfo_exists():
                return
            w = min(target_w, int(target_w * i / 10))
            self.quick_panel.geometry(f"{w}x{rh}+{rx}+{ry}")
            if i < 10:
                self.quick_panel.after(12, lambda: _grow(i + 1))
        _grow()

    def close_quick_actions_panel(self):
        if self.quick_panel and self.quick_panel.winfo_exists():
            self.quick_panel.destroy()
        if self.quick_panel_overlay and self.quick_panel_overlay.winfo_exists():
            self.quick_panel_overlay.destroy()


    def open_subtask_context_menu(self, event):
        row_id = self.sub_tree.identify_row(event.y)
        if row_id:
            self.sub_tree.selection_set(row_id)
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="⚡ Generate This Vendor Contract", command=self.generate_selected_vendor_contract)
        menu.add_separator()
        menu.add_command(label="✎ Edit Vendor", command=self.open_edit_subtask_popup)
        menu.add_command(label="🗑 Delete Vendor", command=self.delete_selected_subtask)
        menu.tk_popup(event.x_root + 18, event.y_root)

    def open_task_context_menu(self, event):
        row_id = self.task_tree.identify_row(event.y)
        if row_id:
            self.task_tree.selection_set(row_id)
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="⚡ Generate All Contracts in Task", command=self.generate_selected_task_contracts)
        menu.add_separator()
        menu.add_command(label="🗑 Delete AQ Creativity", command=self.delete_selected_task)
        menu.tk_popup(event.x_root + 18, event.y_root)

    def _on_task_tree_hover(self, event):
        row_id = self.task_tree.identify_row(event.y)
        if getattr(self, "_hover_task_row", None) and self.task_tree.exists(self._hover_task_row):
            self.task_tree.item(self._hover_task_row, tags=tuple(t for t in self.task_tree.item(self._hover_task_row, "tags") if t != "hover"))
        self._hover_task_row = row_id if row_id else None
        if row_id and self.task_tree.exists(row_id):
            tags = list(self.task_tree.item(row_id, "tags"))
            if "hover" not in tags:
                tags.append("hover")
            self.task_tree.item(row_id, tags=tuple(tags))

    def _clear_task_tree_hover(self, _event=None):
        if getattr(self, "_hover_task_row", None) and self.task_tree.exists(self._hover_task_row):
            self.task_tree.item(self._hover_task_row, tags=tuple(t for t in self.task_tree.item(self._hover_task_row, "tags") if t != "hover"))
        self._hover_task_row = None

    def _on_button_hover(self, event, enter=True):
        widget = event.widget
        if not isinstance(widget, tk.Button):
            return
        if enter:
            widget.configure(highlightthickness=1, highlightbackground=self.ui_colors["accent"], highlightcolor=self.ui_colors["accent"])
        else:
            widget.configure(highlightthickness=0)

    def _bind_events(self):
        self.search_var.trace_add("write", self.apply_filters)
        self.status_filter.trace_add("write", self.apply_filters)
        self.brand_filter.trace_add("write", self.apply_filters)
        self.license_var.trace_add("write", self.load_ibans_for_license)
        self.iban_var.trace_add("write", self.load_bank_info)
        self.vendor_var.trace_add("write", self._load_vendor_profile_from_name)
        self.task_tree.bind("<<TreeviewSelect>>", self.load_task)
        self.task_tree.bind("<Button-3>", self.open_task_context_menu)
        self.sub_tree.bind("<Double-1>", self.edit_subtask)
        self.sub_tree.bind("<Button-3>", self.open_subtask_context_menu)
        self.license_combo.bind("<<ComboboxSelected>>", self.load_ibans_for_license)
        self.task_tree.bind("<Motion>", self._on_task_tree_hover)
        self.task_tree.bind("<Leave>", self._clear_task_tree_hover)

    def open_auth_popup(self, required=False):
        pop = tk.Toplevel(self.root)
        pop.title("Profiles • Login / Sign Up / Reset")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 560, 560)
        self._animate_popup_in(pop)
        pop.lift()
        pop.focus_force()

        notebook = ttk.Notebook(pop)
        notebook.pack(fill="both", expand=True, padx=14, pady=14)

        login_tab = ttk.Frame(notebook, padding=10)
        signup_tab = ttk.Frame(notebook, padding=10)
        notebook.add(login_tab, text="Login")
        notebook.add(signup_tab, text="Sign Up")

        reset_tab = ttk.Frame(notebook, padding=10)
        notebook.add(reset_tab, text="Reset Password")

        is_first_account = self.db.count_users() == 0

        def _add_labeled_entry(tab, row_idx, label_text, var, *, secret=False):
            ttk.Label(tab, text=label_text).grid(row=row_idx, column=0, sticky="w", pady=(0, 4))
            entry = ttk.Entry(tab, textvariable=var, show="*" if secret else "")
            entry.grid(row=row_idx + 1, column=0, sticky="ew", pady=(0, 10))
            return entry

        login_user = tk.StringVar(value=self.db.get_setting("remember_device_user", ""))
        login_pass = tk.StringVar()
        remember_me = tk.BooleanVar(value=bool(login_user.get().strip()))

        ttk.Label(login_tab, text="Welcome back", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w", pady=(2, 10))
        login_user_entry = _add_labeled_entry(login_tab, 1, "Username or email", login_user)
        _add_labeled_entry(login_tab, 3, "Password", login_pass, secret=True)
        ttk.Checkbutton(login_tab, text="Remember me on this device", variable=remember_me).grid(row=5, column=0, sticky="w", pady=(4, 2))
        login_tab.columnconfigure(0, weight=1)

        signup_name = tk.StringVar()
        signup_user = tk.StringVar()
        signup_email = tk.StringVar()
        signup_pass = tk.StringVar()
        signup_pass2 = tk.StringVar()

        ttk.Label(signup_tab, text="Create account", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w", pady=(2, 10))
        _add_labeled_entry(signup_tab, 1, "Full name", signup_name)
        _add_labeled_entry(signup_tab, 3, "Username", signup_user)
        _add_labeled_entry(signup_tab, 5, "Email", signup_email)
        _add_labeled_entry(signup_tab, 7, "Password", signup_pass, secret=True)
        _add_labeled_entry(signup_tab, 9, "Confirm password", signup_pass2, secret=True)

        signup_recovery = tk.StringVar()
        signup_recovery2 = tk.StringVar()
        recovery_box = ttk.LabelFrame(signup_tab, text="Admin recovery key", padding=8)
        recovery_box.grid(row=11, column=0, sticky="ew", pady=(6, 2))
        ttk.Label(recovery_box, text="Required only for first account").grid(row=0, column=0, sticky="w", pady=(0, 6))
        signup_recovery_entry = ttk.Entry(recovery_box, textvariable=signup_recovery, show="*")
        signup_recovery_entry.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        signup_recovery2_entry = ttk.Entry(recovery_box, textvariable=signup_recovery2, show="*")
        signup_recovery2_entry.grid(row=2, column=0, sticky="ew")
        if not is_first_account:
            signup_recovery_entry.configure(state="disabled")
            signup_recovery2_entry.configure(state="disabled")
        recovery_box.columnconfigure(0, weight=1)
        signup_tab.columnconfigure(0, weight=1)

        reset_email = tk.StringVar()
        reset_pass = tk.StringVar()
        reset_pass2 = tk.StringVar()
        reset_recovery_key = tk.StringVar()

        ttk.Label(reset_tab, text="Recover account", style="SubHeader.TLabel").grid(row=0, column=0, sticky="w", pady=(2, 10))
        _add_labeled_entry(reset_tab, 1, "Email", reset_email)
        _add_labeled_entry(reset_tab, 3, "New password", reset_pass, secret=True)
        _add_labeled_entry(reset_tab, 5, "Confirm password", reset_pass2, secret=True)
        _add_labeled_entry(reset_tab, 7, "Recovery key", reset_recovery_key, secret=True)
        reset_tab.columnconfigure(0, weight=1)

        if is_first_account:
            status_text = "No profiles yet. Create the first admin account in Sign Up tab."
            notebook.select(signup_tab)
        else:
            status_text = "Log in to continue."
        status_var = tk.StringVar(value=status_text)
        ttk.Label(pop, textvariable=status_var, foreground="#5d6878").pack(anchor="w", padx=14, pady=(0, 8))

        result = {"ok": False}

        def _finish_login(user_row):
            self.current_user = dict(user_row)
            self.user_badge_var.set(f"Signed in: {self.current_user.get('full_name') or self.current_user['username']} ({self.current_user.get('role','member')})")
            if remember_me.get():
                self.db.set_setting("remember_device_user", self.current_user["username"])
            else:
                self.db.set_setting("remember_device_user", "")
            if hasattr(self, "profile_menu_btn"):
                self.profile_menu_btn.configure(text=f"{self.current_user.get('profile_icon','👤')} {(self.current_user.get('full_name') or self.current_user['username']).split()[0]} ▾")
            self._audit("login", "profile", self.current_user['username'])
            result["ok"] = True
            pop.destroy()

        def _login(*_):
            user_row = self.db.authenticate_user(login_user.get().strip(), login_pass.get())
            if not user_row:
                status_var.set("Invalid username or password.")
                return
            _finish_login(user_row)

        def _signup():
            if signup_pass.get() != signup_pass2.get():
                status_var.set("Passwords do not match.")
                return
            try:
                clean_username = signup_user.get().strip().lower()
                clean_email = signup_email.get().strip().lower()
                clean_name = signup_name.get().strip()
                if not clean_name:
                    status_var.set("Full name is required.")
                    return
                if not clean_username:
                    status_var.set("Username is required.")
                    return
                if "@" not in clean_email:
                    status_var.set("Enter a valid email address.")
                    return
                first_user = self.db.count_users() == 0
                role = "admin" if first_user else "member"
                recovery_key_entered = signup_recovery.get().strip() or signup_recovery2.get().strip()

                if first_user:
                    if signup_recovery.get() != signup_recovery2.get():
                        status_var.set("Recovery keys do not match.")
                        return
                    self.db.set_recovery_key(signup_recovery.get())
                elif recovery_key_entered and not self.db.has_recovery_key():
                    if signup_recovery.get() != signup_recovery2.get():
                        status_var.set("Recovery keys do not match.")
                        return
                    self.db.set_recovery_key(signup_recovery.get())

                self.db.create_user(clean_username, clean_email, clean_name, signup_pass.get(), role=role)
                self._audit("create_profile", "profile", clean_username, details=f"role={role};email={clean_email}")
                status_var.set("Profile created. Log in now.")
                login_user.set(clean_username)
                login_pass.set(signup_pass.get())
                notebook.select(login_tab)
            except sqlite3.IntegrityError:
                status_var.set("Username or email already exists.")
            except ValueError as exc:
                status_var.set(str(exc))

        def _reset_password():
            if reset_pass.get() != reset_pass2.get():
                status_var.set("Reset passwords do not match.")
                return
            try:
                email = reset_email.get().strip().lower()
                if "@" not in email:
                    status_var.set("Enter a valid email address.")
                    return
                updated = self.db.reset_password_by_email(email, reset_pass.get(), reset_recovery_key.get())
                if updated == 0:
                    status_var.set("No profile found for that email.")
                    return
                self._audit("reset_password", "profile", email)
                status_var.set("Password reset. You can now log in.")
                login_user.set(email)
                login_pass.set(reset_pass.get())
                notebook.select(login_tab)
            except ValueError as exc:
                status_var.set(str(exc))

        recovery = ttk.LabelFrame(pop, text="Emergency: Reset all profiles", padding=8)
        recovery.pack(fill="x", padx=14, pady=(0, 8))
        ttk.Label(
            recovery,
            text="Enter your recovery key to reset all accounts and re-create admin.",
            foreground="#a84d00",
            wraplength=470,
            justify="left",
        ).pack(anchor="w")

        recovery_key_var = tk.StringVar()
        ttk.Entry(recovery, textvariable=recovery_key_var, show="*").pack(fill="x", pady=(6, 4))

        def _reset_all_accounts():
            if not self.db.has_recovery_key():
                status_var.set("No recovery key is set yet. Create/set one from Sign Up first.")
                return
            if not self.db.verify_recovery_key(recovery_key_var.get()):
                status_var.set("Invalid recovery key.")
                return
            if not messagebox.askyesno(
                "Reset all accounts",
                "This will DELETE all profiles and logins. Continue?",
                parent=pop,
            ):
                return
            self.db.reset_all_accounts()
            self.db.clear_recovery_key()
            self.current_user = None
            self.user_badge_var.set("Not signed in")
            status_var.set("All accounts removed. Create a new admin in Sign Up tab and set a new recovery key.")
            notebook.select(signup_tab)
            self._audit("reset_all_accounts", "profile")

        ttk.Button(recovery, text="Reset Using Recovery Key", command=_reset_all_accounts).pack(anchor="e", pady=(2, 0))

        btns = ttk.Frame(pop)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(btns, text="Log In", command=_login).pack(side="right")
        ttk.Button(btns, text="Sign Up", command=_signup).pack(side="right", padx=(0, 8))
        ttk.Button(btns, text="Reset Password", command=_reset_password).pack(side="right", padx=(0, 8))

        if required:
            ttk.Button(btns, text="Exit app", command=pop.destroy).pack(side="left")
        else:
            ttk.Button(btns, text="Close", command=pop.destroy).pack(side="left")

        def _submit_by_tab(_event=None):
            tab = notebook.select()
            if tab == str(login_tab):
                _login()
            elif tab == str(signup_tab):
                _signup()
            else:
                _reset_password()

        pop.bind("<Return>", _submit_by_tab)
        pop.protocol("WM_DELETE_WINDOW", pop.destroy)
        login_user_entry.focus_set()
        pop.wait_window()
        return result["ok"]

    def open_audit_log_popup(self):
        if not self._require_permission("settings", "Audit Log"):
            return
        pop = tk.Toplevel(self.root)
        pop.title("Audit Log")
        pop.transient(self.root)
        self._lock_popup_position(pop, 900, 420)
        self._animate_popup_in(pop)

        cols = ("When", "User", "Action", "Entity", "ID", "Details")
        tree = ttk.Treeview(pop, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=130 if col != "Details" else 250, anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for row in self.db.list_recent_audit(300):
            tree.insert(
                "",
                "end",
                values=(row["created_at"], row["actor_username"], row["action"], row["entity_type"], row["entity_id"], row["details"]),
            )

    def open_vendor_master_popup(self):
        if not self._require_permission("edit", "Vendor master update"):
            return
        pop = tk.Toplevel(self.root)
        pop.title("Vendor Master")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 520, 260)
        self._animate_popup_in(pop)

        name_v = tk.StringVar(value=self.vendor_name_auto_var.get() or self.vendor_var.get())
        lic_v = tk.StringVar(value=self.license_var.get())
        iban_v = tk.StringVar(value=self.iban_var.get())
        bank_v = tk.StringVar(value=self.bank_name_var.get())
        accn_v = tk.StringVar(value=self.account_name_var.get())
        acct_v = tk.StringVar(value=self.account_number_var.get())
        swift_v = tk.StringVar(value=self.swift_code_var.get())

        frame = ttk.Frame(pop, padding=12)
        frame.pack(fill="both", expand=True)
        labels = [
            ("Vendor Name", name_v), ("License", lic_v), ("IBAN", iban_v),
            ("Bank Name", bank_v), ("Account Name", accn_v), ("Account #", acct_v), ("SWIFT", swift_v),
        ]
        for i, (lab, var) in enumerate(labels):
            ttk.Label(frame, text=lab).grid(row=i, column=0, sticky="w", pady=2)
            ttk.Entry(frame, textvariable=var, width=36).grid(row=i, column=1, sticky="ew", pady=2)

        frame.columnconfigure(1, weight=1)

        def _save():
            if not lic_v.get().strip():
                messagebox.showwarning("Missing license", "License is required", parent=pop)
                return
            self.db.upsert_vendor_bank(name_v.get().strip(), lic_v.get().strip(), bank_v.get().strip(), accn_v.get().strip(), iban_v.get().strip(), acct_v.get().strip(), swift_v.get().strip())
            self.license_var.set(lic_v.get().strip())
            self.vendor_name_auto_var.set(name_v.get().strip())
            self.vendor_var.set(name_v.get().strip())
            self.iban_var.set(iban_v.get().strip())
            self.bank_name_var.set(bank_v.get().strip())
            self.account_name_var.set(accn_v.get().strip())
            self.account_number_var.set(acct_v.get().strip())
            self.swift_code_var.set(swift_v.get().strip())
            self.refresh_license_combo()
            self.refresh_vendor_combo()
            self._audit("upsert_vendor_master", "vendor", lic_v.get().strip(), details=f"vendor={name_v.get().strip()}")
            pop.destroy()

        ttk.Button(frame, text="Save", command=_save).grid(row=len(labels), column=1, sticky="e", pady=(10, 0))

    def import_excel_to_db(self):
        if not self._require_permission("import_export", "Import Excel"):
            return
        file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            import pandas as pd
        except Exception:
            messagebox.showerror("Missing dependency", "Please install pandas/openpyxl first: pip install pandas openpyxl")
            return

        try:
            df = pd.read_excel(file_path, dtype=str).fillna("")
            df = df.rename(columns={c: c.strip().lower().replace(" ", "_").replace("(", "").replace(")", "") for c in df.columns})

            def col(*names):
                for n in names:
                    if n in df.columns:
                        return n
                return None

            brand_c = col("brand", "brand_name")
            vendor_c = col("vendor", "vendor_name", "name", "influencer_name_as_per_license", "influencer_name_as_per_license_")
            lic_c = col("license", "license_number")
            iban_c = col("iban")
            bank_c = col("bank_name")
            accn_c = col("account_name")
            acct_c = col("account_number")
            swift_c = col("swift", "swift_code")

            imported = 0
            for _, row in df.iterrows():
                brand = row[brand_c].strip() if brand_c else ""
                vendor = row[vendor_c].strip() if vendor_c else ""
                lic = row[lic_c].strip() if lic_c else ""
                iban = row[iban_c].strip() if iban_c else ""
                bank = row[bank_c].strip() if bank_c else ""
                accn = row[accn_c].strip() if accn_c else ""
                acct = row[acct_c].strip() if acct_c else ""
                swift = row[swift_c].strip() if swift_c else ""
                if brand:
                    self.db.upsert_brand(brand)
                if lic:
                    self.db.upsert_vendor_bank(vendor, lic, bank, accn, iban, acct, swift)
                    imported += 1

            self.refresh_brand_combo()
            self.refresh_license_combo()
            self.refresh_vendor_combo()
            self._audit("import_excel", "database", details=f"rows={imported} file={file_path}")
            messagebox.showinfo("Import complete", f"Imported/updated {imported} vendor license records from Excel.")
        except Exception as exc:
            messagebox.showerror("Import failed", str(exc))

    def export_db_to_excel(self):
        if not self._require_permission("import_export", "Export Excel"):
            return
        file_path = filedialog.asksaveasfilename(title="Export to Excel", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            import pandas as pd
        except Exception:
            messagebox.showerror("Missing dependency", "Please install pandas/openpyxl first: pip install pandas openpyxl")
            return

        try:
            tasks = [dict(r) for r in self.db.list_tasks()]
            vendors = [
                dict(r)
                for r in self.db.conn.execute(
                    """
                    SELECT v.name, v.license_number, ba.bank_name, ba.account_name, ba.iban, ba.account_number, COALESCE(ba.swift_code,'') AS swift_code
                    FROM vendors v
                    LEFT JOIN bank_accounts ba ON ba.vendor_id=v.id
                    ORDER BY v.name, v.license_number
                    """
                ).fetchall()
            ]
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                pd.DataFrame(tasks).to_excel(writer, index=False, sheet_name="contracts")
                pd.DataFrame(vendors).to_excel(writer, index=False, sheet_name="vendors")
            self._audit("export_excel", "database", details=f"tasks={len(tasks)} vendors={len(vendors)} file={file_path}")
            messagebox.showinfo("Export complete", "Data exported to Excel successfully.")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_json(self):
        if not self._require_permission("import_export", "Export JSON"):
            return

        selected_task_ids = list(self.task_tree.selection())
        if not selected_task_ids:
            selected_task_ids = [row["id"] for row in self.db.list_tasks()]

        if not selected_task_ids:
            messagebox.showwarning("No data", "There are no tasks to export.")
            return

        file_path = filedialog.asksaveasfilename(title="Export to JSON", defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not file_path:
            return

        try:
            export_rows = []
            for task_id in selected_task_ids:
                task = self.db.get_task(task_id)
                if not task:
                    continue

                subtasks = [dict(row) for row in self.db.list_subtasks(task_id)]
                export_rows.append({
                    "task": dict(task),
                    "subtasks": subtasks,
                })

            payload = {
                "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "task_count": len(export_rows),
                "tasks": export_rows,
            }

            Path(file_path).write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
            self._audit("export_json", "database", details=f"tasks={len(export_rows)} file={file_path}")
            messagebox.showinfo("Export complete", f"Exported {len(export_rows)} task(s) to JSON.")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def _locate_contract_pdf(self, contract_id: str):
        try:
            conf = load_config()
            output_dir = Path(conf.get("output_dir", "")).expanduser()
        except Exception:
            return None
        if not output_dir.exists():
            return None
        matches = list(output_dir.rglob(f"{contract_id}.pdf"))
        if not matches:
            return None
        return max(matches, key=lambda p: p.stat().st_mtime)

    def export_contracts_to_pdf(self):
        if not self._require_permission("import_export", "Export PDF"):
            return
        selected_ids = list(self.task_tree.selection()) or ([self.current_task.get()] if self.current_task.get() else [])
        selected_ids = [t for t in selected_ids if t]
        if not selected_ids:
            messagebox.showwarning("No task", "Select at least one contract first.")
            return

        export_dir = filedialog.askdirectory(title="Select folder for PDF export")
        if not export_dir:
            return
        target_dir = Path(export_dir)
        target_dir.mkdir(parents=True, exist_ok=True)

        copied = 0
        missing = []
        for task_id in selected_ids:
            contract_rows = self.db.list_generated_contracts_for_task(task_id)
            if not contract_rows:
                missing.append(task_id)
                continue
            for idx, c in enumerate(contract_rows, start=1):
                src = self._locate_contract_pdf(c["contract_id"])
                if not src:
                    continue
                file_name = f"{task_id}.pdf" if len(contract_rows) == 1 else f"{task_id}_{idx}.pdf"
                dst = target_dir / file_name
                dst.write_bytes(src.read_bytes())
                copied += 1
        self._audit("export_pdf", "database", details=f"tasks={len(selected_ids)} copied={copied}")
        if copied:
            messagebox.showinfo("Export complete", f"Exported {copied} PDF file(s) to:\n{target_dir}")
        else:
            messagebox.showwarning("No PDFs found", "No generated PDFs were found for selected contracts.")

    def open_export_popup(self):
        pop = tk.Toplevel(self.root)
        pop.title("Export")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 340, 180)
        self._animate_popup_in(pop)
        frame = ttk.Frame(pop, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Choose export format", style="SubHeader.TLabel").pack(anchor="w", pady=(0, 12))
        ttk.Button(frame, text="Export PDF", command=lambda: (pop.destroy(), self.export_contracts_to_pdf())).pack(fill="x", pady=(0, 8))
        ttk.Button(frame, text="Export Excel", command=lambda: (pop.destroy(), self.export_db_to_excel())).pack(fill="x", pady=(0, 8))
        ttk.Button(frame, text="Export JSON", command=lambda: (pop.destroy(), self.export_json())).pack(fill="x")

    def open_reports_popup(self):
        rows = [dict(r) for r in self.db.list_tasks()]
        pop = tk.Toplevel(self.root)
        pop.title("Reports")
        pop.transient(self.root)
        self._lock_popup_position(pop, 540, 380)
        self._animate_popup_in(pop)

        total_revenue = sum(self._parse_price(r.get("amount", "0")) for r in rows)
        vendor_payments = {}
        monthly_contracts = {}
        for task in rows:
            month_key = (task.get("created_date") or "Unknown")[:7]
            monthly_contracts[month_key] = monthly_contracts.get(month_key, 0) + 1
            for sub in self.db.list_subtasks(task["id"]):
                vendor = (sub["vendor"] or "Unknown").strip() or "Unknown"
                vendor_payments[vendor] = vendor_payments.get(vendor, 0.0) + self._parse_price(sub["price"])

        frame = ttk.Frame(pop, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Reports", style="SubHeader.TLabel").pack(anchor="w")
        ttk.Label(frame, text=f"Total Revenue: {total_revenue:.2f}").pack(anchor="w", pady=(10, 4))
        ttk.Label(frame, text=f"Vendor Payments Total: {sum(vendor_payments.values()):.2f}").pack(anchor="w", pady=4)

        ttk.Label(frame, text="Top Vendor Payments", style="SubHeader.TLabel").pack(anchor="w", pady=(12, 4))
        for vendor, amount in sorted(vendor_payments.items(), key=lambda x: x[1], reverse=True)[:8]:
            ttk.Label(frame, text=f"• {vendor}: {amount:.2f}").pack(anchor="w")

        ttk.Label(frame, text="Monthly Contracts", style="SubHeader.TLabel").pack(anchor="w", pady=(12, 4))
        for month, count in sorted(monthly_contracts.items()):
            ttk.Label(frame, text=f"• {month}: {count}").pack(anchor="w")

    def refresh_template_dropdown(self):
        try:
            conf = load_config()
            keys = sorted(conf.get("template_map", {}).keys())
        except Exception:
            keys = []

        self.available_template_keys = keys
        values = ["auto"] + keys
        self.template_combo["values"] = values
        if self.template_key_var.get() not in values:
            self.template_key_var.set(self.default_type_setting.get() if self.default_type_setting.get() in values else "auto")

    def add_template_popup(self):
        if not self._require_permission("settings", "Add template"):
            return

        pop = tk.Toplevel(self.root)
        pop.title("Add Template")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 540, 180)
        self._animate_popup_in(pop)

        frame = ttk.Frame(pop, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        key_var = tk.StringVar()
        path_var = tk.StringVar()

        ttk.Label(frame, text="Template key").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=key_var).grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Label(frame, text="DOCX template file").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=path_var).grid(row=1, column=1, sticky="ew", pady=4)

        def _browse():
            fp = filedialog.askopenfilename(title="Select template", filetypes=[("Word template", "*.docx")])
            if fp:
                path_var.set(fp)

        ttk.Button(frame, text="Browse", command=_browse).grid(row=1, column=2, padx=(8, 0))

        def _save():
            key = key_var.get().strip().lower().replace(" ", "_")
            path = path_var.get().strip()
            if not key or not path:
                messagebox.showwarning("Missing data", "Template key and path are required")
                return
            if not Path(path).exists():
                messagebox.showerror("Invalid path", "Selected template file does not exist")
                return

            conf = load_config()
            conf.setdefault("template_map", {})[key] = path
            CONFIG_PATH.write_text(json.dumps(conf, indent=2, ensure_ascii=False), encoding="utf-8")
            self.refresh_template_dropdown()
            self.template_key_var.set(key)
            self._audit("add_template", "template", key, details=path)
            pop.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=pop.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Save", command=_save).pack(side="left")

    def _save_template_config(self, conf):
        CONFIG_PATH.write_text(json.dumps(conf, indent=2, ensure_ascii=False), encoding="utf-8")

    def manage_templates_popup(self):
        if not self._require_permission("settings", "Manage templates"):
            return

        pop = tk.Toplevel(self.root)
        pop.title("Manage Templates")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 680, 360)
        self._animate_popup_in(pop)

        frame = ttk.Frame(pop, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        ttk.Label(list_frame, text="Template keys").pack(anchor="w")
        key_list = tk.Listbox(list_frame, width=24, exportselection=False)
        key_list.pack(fill="y", expand=True, pady=(6, 0))

        edit_frame = ttk.LabelFrame(frame, text="Template Details", padding=10)
        edit_frame.grid(row=0, column=1, sticky="nsew")
        edit_frame.columnconfigure(1, weight=1)

        key_var = tk.StringVar()
        path_var = tk.StringVar()
        rename_var = tk.StringVar()

        ttk.Label(edit_frame, text="Current key").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(edit_frame, textvariable=key_var, state="readonly").grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Label(edit_frame, text="Rename key to").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(edit_frame, textvariable=rename_var).grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Label(edit_frame, text="Template path").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(edit_frame, textvariable=path_var).grid(row=2, column=1, sticky="ew", pady=4)

        def _browse_path():
            fp = filedialog.askopenfilename(title="Select template", filetypes=[("Word template", "*.docx")])
            if fp:
                path_var.set(fp)

        ttk.Button(edit_frame, text="Browse", command=_browse_path).grid(row=2, column=2, padx=(8, 0))

        def _load_keys(selected_key=""):
            conf = load_config()
            keys = sorted(conf.get("template_map", {}).keys())
            key_list.delete(0, tk.END)
            for k in keys:
                key_list.insert(tk.END, k)
            target = selected_key if selected_key in keys else (keys[0] if keys else "")
            if target:
                idx = keys.index(target)
                key_list.selection_clear(0, tk.END)
                key_list.selection_set(idx)
                key_list.activate(idx)
                _load_selected()
            else:
                key_var.set("")
                rename_var.set("")
                path_var.set("")

        def _selected_key():
            sel = key_list.curselection()
            if not sel:
                return ""
            return key_list.get(sel[0])

        def _load_selected(_event=None):
            skey = _selected_key()
            conf = load_config()
            tmap = conf.get("template_map", {})
            key_var.set(skey)
            rename_var.set(skey)
            path_var.set(tmap.get(skey, ""))

        key_list.bind("<<ListboxSelect>>", _load_selected)

        def _update_template():
            old_key = _selected_key()
            if not old_key:
                messagebox.showwarning("No selection", "Select a template key first")
                return

            new_key = rename_var.get().strip().lower().replace(" ", "_")
            new_path = path_var.get().strip()
            if not new_key or not new_path:
                messagebox.showwarning("Missing data", "Rename key and template path are required")
                return
            if not Path(new_path).exists():
                messagebox.showerror("Invalid path", "Template file path does not exist")
                return

            conf = load_config()
            tmap = conf.setdefault("template_map", {})
            if old_key not in tmap:
                messagebox.showerror("Missing key", "Selected template key no longer exists")
                return

            # rename/update
            del tmap[old_key]
            tmap[new_key] = new_path

            if conf.get("standard_template_key") == old_key:
                conf["standard_template_key"] = new_key
            if self.default_type_setting.get() == old_key:
                self.default_type_setting.set(new_key)
                self.db.set_setting("default_contract_type", new_key)

            self._save_template_config(conf)
            self.refresh_template_dropdown()
            self.template_key_var.set(new_key)
            self._audit("update_template", "template", old_key, details=f"renamed_to={new_key};path={new_path}")
            _load_keys(new_key)

        def _delete_template():
            del_key = _selected_key()
            if not del_key:
                messagebox.showwarning("No selection", "Select a template key first")
                return
            if not messagebox.askyesno("Delete template", f"Delete template '{del_key}'?"):
                return

            conf = load_config()
            tmap = conf.setdefault("template_map", {})
            if del_key not in tmap:
                messagebox.showerror("Missing key", "Selected template key no longer exists")
                return

            if len(tmap) <= 1:
                messagebox.showerror("Blocked", "At least one template must remain")
                return

            del tmap[del_key]
            remaining_keys = sorted(tmap.keys())
            if conf.get("standard_template_key") == del_key:
                conf["standard_template_key"] = remaining_keys[0]
            if self.default_type_setting.get() == del_key:
                self.default_type_setting.set("auto")
                self.db.set_setting("default_contract_type", "auto")
            if self.template_key_var.get() == del_key:
                self.template_key_var.set("auto")

            self._save_template_config(conf)
            self.refresh_template_dropdown()
            self._audit("delete_template", "template", del_key)
            _load_keys()

        btns = ttk.Frame(edit_frame)
        btns.grid(row=3, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Update", command=_update_template).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Delete", command=_delete_template).pack(side="left")

        _load_keys(self.template_key_var.get())

    def open_profile_menu(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Profile", command=self.open_profile_settings_popup)
        menu.add_command(label="Settings", command=self.open_settings_popup)
        menu.add_separator()
        menu.add_command(label="Log out", command=self.logout_current_user)
        x = self.profile_menu_btn.winfo_rootx() if hasattr(self, "profile_menu_btn") else self.root.winfo_rootx() + 100
        y = self.profile_menu_btn.winfo_rooty() + self.profile_menu_btn.winfo_height() if hasattr(self, "profile_menu_btn") else self.root.winfo_rooty() + 100
        menu.tk_popup(x, y)

    def logout_current_user(self):
        self.current_user = None
        self.db.set_setting("remember_device_user", "")
        self.user_badge_var.set("Not signed in")
        if hasattr(self, "profile_menu_btn"):
            self.profile_menu_btn.configure(text=f"{(self.current_user.get('profile_icon','👤') if self.current_user else '👤')} Profile ▾")
        self.require_login()

    def open_profile_settings_popup(self):
        if not self.current_user:
            messagebox.showwarning("Not signed in", "Please sign in first.")
            return

        pop = tk.Toplevel(self.root)
        pop.title("Profile Settings")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 520, 360)
        self._animate_popup_in(pop)

        frame = ttk.Frame(pop, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        full_name_v = tk.StringVar(value=self.current_user.get("full_name", ""))
        username_v = tk.StringVar(value=self.current_user.get("username", ""))
        email_v = tk.StringVar(value=self.current_user.get("email", ""))
        color_name_map = {"Blue": "#4f8cff", "Green": "#22c55e", "Orange": "#f59e0b", "Red": "#ef4444", "Purple": "#a855f7"}
        reverse_color_map = {v: k for k, v in color_name_map.items()}
        color_v = tk.StringVar(value=reverse_color_map.get(self.current_user.get("profile_color", "#4f8cff"), "Blue"))
        icon_v = tk.StringVar(value=self.current_user.get("profile_icon", "👤"))
        cur_pass_v = tk.StringVar()
        new_pass_v = tk.StringVar()

        ttk.Label(frame, text="Full name").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=full_name_v).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Username").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=username_v).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Email").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=email_v).grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Profile color").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=color_v, values=list(color_name_map.keys()), state="readonly", width=14).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Label(frame, text="Profile icon").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=icon_v, values=["👤", "🧑", "👩", "👨", "🦊", "🐼"], state="readonly", width=8).grid(row=4, column=1, sticky="w", pady=4)

        sep = ttk.Separator(frame, orient="horizontal")
        sep.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(10, 10))

        ttk.Label(frame, text="Current password").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=cur_pass_v, show="*").grid(row=6, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="New password").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=new_pass_v, show="*").grid(row=7, column=1, sticky="ew", pady=4)

        def _save_profile():
            try:
                self.db.update_user_profile(
                    self.current_user["id"],
                    full_name_v.get(),
                    username_v.get(),
                    email_v.get(),
                    color_name_map.get(color_v.get(), "#4f8cff"),
                    icon_v.get(),
                )
                self.current_user["full_name"] = full_name_v.get().strip()
                self.current_user["username"] = username_v.get().strip().lower()
                self.current_user["email"] = email_v.get().strip().lower()
                self.current_user["profile_color"] = color_v.get().strip() or "#4f8cff"
                self.current_user["profile_icon"] = icon_v.get().strip() or "👤"
                self.user_badge_var.set(f"Signed in: {self.current_user.get('full_name') or self.current_user['username']} ({self.current_user.get('role','member')})")
                if hasattr(self, "profile_menu_btn"):
                    self.profile_menu_btn.configure(text=f"{self.current_user.get('profile_icon','👤')} {(self.current_user.get('full_name') or self.current_user['username']).split()[0]} ▾")
                if cur_pass_v.get().strip() or new_pass_v.get().strip():
                    self.db.change_user_password(self.current_user["id"], cur_pass_v.get(), new_pass_v.get())
                self._audit("update_profile", "profile", self.current_user["username"])
                messagebox.showinfo("Saved", "Profile updated successfully.", parent=pop)
                pop.destroy()
            except ValueError as exc:
                messagebox.showerror("Profile update failed", str(exc), parent=pop)

        btns = ttk.Frame(frame)
        btns.grid(row=8, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=pop.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Save", command=_save_profile).pack(side="left")

    def open_settings_popup(self):
        if not self._require_permission("settings", "Settings"):
            return
        pop = tk.Toplevel(self.root)
        pop.title("App Settings")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 620, 420)
        self._animate_popup_in(pop)

        frame = ttk.Frame(pop, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        title_v = tk.StringVar(value=self.app_title_var.get())
        default_type_v = tk.StringVar(value=self.default_type_setting.get())
        default_status_v = tk.StringVar(value=self.default_status_setting.get())
        theme_v = tk.StringVar(value=self.theme_setting.get())
        db_path_v = tk.StringVar(value=str(self.db.db_path))

        ttk.Label(frame, text="App title").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=title_v).grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Label(frame, text="Default template").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=default_type_v, values=["auto"] + self.available_template_keys, state="readonly", width=20).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(frame, text="Default status").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=default_status_v, values=STATUS_OPTIONS, state="readonly", width=12).grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(frame, text="Theme").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=theme_v, values=["dark", "light"], state="readonly", width=12).grid(row=3, column=1, sticky="w", pady=4)

        ttk.Label(frame, text="These defaults are used when creating a new AQ Creativity task.", foreground="#666").grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Label(frame, text="Shared DB path", foreground="#666").grid(row=5, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frame, state="readonly", width=56, justify="left", textvariable=db_path_v).grid(row=5, column=1, sticky="ew", pady=(8, 0))
        ttk.Label(frame, text="For office multi-user setup, point AQ_DB_PATH to a shared server path and keep regular backups.", foreground="#666").grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))

        tools = ttk.LabelFrame(frame, text="Tools", padding=8)
        tools.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(tools, text="Import Excel", command=self.import_excel_to_db).pack(side="left", padx=(0, 6))
        ttk.Button(tools, text="Export Excel", command=self.export_db_to_excel).pack(side="left", padx=(0, 6))
        ttk.Button(tools, text="Audit Log", command=self.open_audit_log_popup).pack(side="left", padx=(0, 6))
        ttk.Button(tools, text="Profiles", command=self.open_auth_popup).pack(side="left")

        def _save():
            app_title = (title_v.get().strip() or APP_BRAND)
            self.db.set_setting("app_title", app_title)
            self.db.set_setting("default_contract_type", default_type_v.get().strip() or "auto")
            self.db.set_setting("default_status", default_status_v.get().strip() or "NEW")
            self.db.set_setting("theme_mode", theme_v.get().strip() or "dark")

            self.app_title_var.set(app_title)
            self.default_type_setting.set(default_type_v.get().strip() or "auto")
            self.default_status_setting.set(default_status_v.get().strip() or "NEW")
            self.theme_setting.set(theme_v.get().strip() or "dark")
            self.apply_theme(self.theme_setting.get())
            self.root.title(f"{self.app_title_var.get()} – Dashboard")
            self._audit("update_settings", "app_settings", details=f"title={app_title};theme={self.theme_setting.get()}")
            pop.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=8, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Cancel", command=pop.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Save", command=_save).pack(side="left")

    def refresh_brand_combo(self):
        brands = self.db.get_all_brands()
        self.brand_combo.set_values(brands)
        self.brand_filter_combo["values"] = ["ALL"] + brands
        if self.brand_filter.get() not in ["ALL"] + brands:
            self.brand_filter.set("ALL")

    def refresh_license_combo(self):
        self.license_combo.set_values(self.db.get_license_numbers())

    def refresh_vendor_combo(self):
        self.vendor_combo.set_values(self.db.get_vendor_names() or FALLBACK_VENDORS)

    def load_ibans_for_license(self, *_):
        lic = self.license_var.get().strip()
        if not lic:
            self.vendor_name_auto_var.set("")
            self.iban_combo["values"] = []
            self.iban_var.set("")
            self.bank_name_var.set("")
            self.account_name_var.set("")
            self.account_number_var.set("")
            self.swift_code_var.set("")
            return

        vendor_name = self.db.get_vendor_name_by_license(lic)
        if vendor_name:
            self.vendor_name_auto_var.set(vendor_name)
            self.vendor_var.set(vendor_name)
        else:
            # keep editable vendor name if DB has no explicit name for this license
            self.vendor_name_auto_var.set(self.vendor_var.get().strip())

        ibans = self.db.get_ibans_for_license(lic)
        self.iban_combo["values"] = ibans
        if ibans and not self.iban_var.get().strip():
            self.iban_var.set(ibans[0])
        elif not ibans:
            self.iban_var.set("")

    def load_bank_info(self, *_):
        iban = self.iban_var.get().strip()
        if not iban:
            return
        row = self.db.get_bank_info_by_iban(iban)
        if row:
            self.bank_name_var.set(row["bank_name"])
            self.account_name_var.set(row["account_name"])
            self.account_number_var.set(row["account_number"])
            self.swift_code_var.set(row["swift_code"])

    def refresh_task_rows(self):
        self.task_tree.delete(*self.task_tree.get_children())
        for idx, row in enumerate(self.db.list_tasks()):
            zebra_tag = "zebra_even" if idx % 2 == 0 else "zebra_odd"
            status_display = self._status_badge_text(row["status"])
            self.task_tree.insert(
                "",
                "end",
                iid=row["id"],
                values=(row["id"], row["brand"], status_display, row["amount"], row["contract_type"], row["vendor_count"], row["created_date"]),
                tags=(zebra_tag, row["status"]),
            )
        self.apply_filters()
        self.refresh_runtime_badge()

    def apply_filters(self, *_):
        text, status, brand = self.search_var.get().lower(), self.status_filter.get(), self.brand_filter.get()
        for row in self.db.list_tasks():
            iid = row["id"]
            visible = not (
                (status != "ALL" and row["status"] != status)
                or (brand != "ALL" and row["brand"] != brand)
                or (text and text not in iid.lower() and text not in row["brand"].lower())
            )
            if visible and self.task_tree.exists(iid):
                self.task_tree.reattach(iid, "", "end")
            elif self.task_tree.exists(iid):
                self.task_tree.detach(iid)

    def new_task(self):
        if not self._require_permission("edit", "Create task"):
            return
        if self.current_task.get():
            self.save_task(silent=True)
        task_id = create_task_id([r["id"] for r in self.db.list_tasks()])
        self.db.create_task(task_id, self.default_type_setting.get() or "auto", self.default_status_setting.get() or "NEW")
        self.refresh_task_rows()
        self.current_task.set(task_id)
        self.brand_var.set("")
        self.amount_var.set("")
        self.template_key_var.set(self.default_type_setting.get() or "auto")
        self.status_var.set(self.default_status_setting.get() or "NEW")
        self.sub_tree.delete(*self.sub_tree.get_children())
        self.clear_subtask_editor()
        self._set_editor_empty_state(False)
        self._audit("create_task", "task", task_id)

    def delete_selected_task(self):
        if not self._require_permission("delete", "Delete tasks"):
            return
        selected = list(self.task_tree.selection())
        if not selected:
            messagebox.showwarning("No task", "Select one or more tasks to delete")
            return
        if not messagebox.askyesno("Delete task(s)", f"Delete {len(selected)} selected task(s) and all their vendors?"):
            return

        deleted = 0
        for tid in selected:
            if self.db.get_task(tid):
                self.db.delete_task(tid)
                self._audit("delete_task", "task", tid)
                deleted += 1

        if self.current_task.get() in selected:
            self.current_task.set("")
            self._set_editor_empty_state(True)
            self.brand_var.set("")
            self.amount_var.set("")
            self.sub_tree.delete(*self.sub_tree.get_children())
            self.clear_subtask_editor()

        self.refresh_task_rows()
        self.refresh_brand_combo()
        messagebox.showinfo("Deleted", f"Deleted {deleted} task(s).")

    def save_task(self, silent=False):
        if not self._require_permission("edit", "Save task"):
            return
        tid = self.current_task.get()
        if not tid:
            return
        self._recalculate_task_amount()
        self.db.upsert_task(tid, self.brand_var.get().strip(), self.amount_var.get().strip(), self.template_key_var.get().strip() or "auto", self.status_var.get().strip() or "NEW")
        self.db.upsert_brand(self.brand_var.get().strip())
        self._audit("update_task", "task", tid, details=f"brand={self.brand_var.get().strip()}")
        self.refresh_task_rows()
        self.refresh_brand_combo()
        if not silent:
            messagebox.showinfo("Saved", f"Task {tid} saved.")

    def load_task(self, _event=None):
        sel = self.task_tree.selection()
        if not sel:
            return
        tid = sel[0]
        row = self.db.get_task(tid)
        if not row:
            return
        self.current_task.set(tid)
        self.brand_var.set(row["brand"])
        self.amount_var.set(row["amount"] or "0.00")
        self.template_key_var.set(row["contract_type"] or "auto")
        self.status_var.set(row["status"])
        self.refresh_subtasks(tid)
        subtask_rows = self.db.list_subtasks(tid)
        if subtask_rows:
            first_id = str(subtask_rows[0]["id"])
            if self.sub_tree.exists(first_id):
                self.sub_tree.selection_set(first_id)
            self._load_subtask_into_editor(subtask_rows[0])
        else:
            self.clear_subtask_editor()
        self._set_editor_empty_state(False)

    def refresh_subtasks(self, task_id):
        self.sub_tree.delete(*self.sub_tree.get_children())
        for sub in self.db.list_subtasks(task_id):
            self.sub_tree.insert("", "end", iid=str(sub["id"]), values=(sub["vendor"], sub["channel"], sub["platforms"], sub["ad_type"], sub["qty"], sub["price"], sub["details"]))
        self._recalculate_task_amount()

    def clear_subtask_editor(self):
        self.editing_subtask_id = None
        self.vendor_var.set("")
        self.channel_var.set("")
        for var in self.platform_channel_vars.values():
            var.set("")
        self.type_sub_var.set("Store Visit")
        self.qty_var.set("1")
        self.detail_var.set("")
        self.vendor_price_var.set("0")
        self.license_var.set("")
        self.vendor_name_auto_var.set("")
        self.iban_var.set("")
        self.bank_name_var.set("")
        self.account_name_var.set("")
        self.account_number_var.set("")
        self.swift_code_var.set("")
        for v in self.platform_vars.values():
            v.set(False)
        self.refresh_platform_detail_inputs()
        self.toggle_details()
        if hasattr(self, "add_vendor_btn"):
            self.add_vendor_btn.configure(text="Add Vendor")

    def _load_subtask_into_editor(self, subtask_row):
        self.editing_subtask_id = int(subtask_row["id"])
        self.vendor_var.set(subtask_row["vendor"])
        self.channel_var.set(subtask_row["channel"])
        raw_platforms = [p.strip() for p in (subtask_row["platforms"] or "").split(",") if p.strip()]
        for p in self.platform_vars:
            self.platform_vars[p].set(p in raw_platforms)
        self._load_platform_channels(subtask_row["channel"])
        self.refresh_platform_detail_inputs()
        self.type_sub_var.set(subtask_row["ad_type"] or "Store Visit")
        self.qty_var.set(subtask_row["qty"] or "1")
        self.vendor_price_var.set(subtask_row["price"] or "0")
        self.detail_var.set(subtask_row["details"] or "")
        self.toggle_details()
        self._load_vendor_profile_from_name()
        if hasattr(self, "add_vendor_btn"):
            self.add_vendor_btn.configure(text="Add Vendor")

    def _load_vendor_profile_from_name(self, *_):
        profile = self.db.get_vendor_profile_by_name(self.vendor_var.get().strip())
        if not profile:
            return
        self.license_var.set(profile["license_number"] or "")
        self.vendor_name_auto_var.set(profile["name"] or "")
        self.iban_var.set(profile["iban"] or "")
        self.bank_name_var.set(profile["bank_name"] or "")
        self.account_name_var.set(profile["account_name"] or "")
        self.account_number_var.set(profile["account_number"] or "")
        self.swift_code_var.set(profile["swift_code"] or "")

    def add_or_update_subtask(self):
        if not self._require_permission("edit", "Edit vendor subtasks"):
            return
        tid = self.current_task.get()
        if not tid:
            messagebox.showwarning("Select task first", "Pick a task on the left")
            return
        plats = [p for p, v in self.platform_vars.items() if v.get()]
        price = f"{self._parse_price(self.vendor_price_var.get()):.2f}"
        channel_payload = self._serialize_platform_channels(plats)
        payload = (self.vendor_var.get().strip(), channel_payload, ", ".join(plats), self.type_sub_var.get().strip() or "Store Visit", self.qty_var.get().strip() or "1", self.detail_var.get().strip(), price)
        self.db.create_subtask(
            tid,
            self.vendor_var.get().strip(),
            self.license_var.get().strip(),
            self.iban_var.get().strip(),
            channel_payload,
            ", ".join(plats),
            self.type_sub_var.get(),
            self.qty_var.get(),
            self.detail_var.get(),
            self.vendor_price_var.get(),
        )
        self._audit("create_subtask", "subtask", tid, details=f"vendor={payload[0]}")
        self.clear_subtask_editor()
        self.refresh_subtasks(tid)
        self.save_task(silent=True)
        self.refresh_task_rows()

    def delete_selected_subtask(self):
        if not self._require_permission("delete", "Delete vendor subtasks"):
            return
        selected = list(self.sub_tree.selection())
        if not selected:
            messagebox.showwarning("No vendor", "Select one or more vendor rows to remove")
            return
        if not messagebox.askyesno("Remove vendors", f"Delete {len(selected)} selected vendor row(s)?"):
            return
        for sid in selected:
            self.db.delete_subtask(int(sid))
            self._audit("delete_subtask", "subtask", str(sid))
        tid = self.current_task.get()
        if tid:
            self.refresh_subtasks(tid)
            self.save_task(silent=True)
        self.refresh_task_rows()

    def edit_subtask(self, _event):
        self.open_edit_subtask_popup()

    def open_edit_subtask_popup(self):
        sel = self.sub_tree.selection()
        if not sel:
            messagebox.showwarning("No vendor", "Select a vendor row first")
            return
        if not self._require_permission("edit", "Edit vendor subtasks"):
            return

        sid = int(sel[0])
        tid = self.current_task.get()
        if not tid:
            messagebox.showwarning("No task", "Select a task first")
            return

        subtask = next((row for row in self.db.list_subtasks(tid) if row["id"] == sid), None)
        if not subtask:
            messagebox.showerror("Missing vendor", "Selected vendor row was not found")
            return

        pop = tk.Toplevel(self.root)
        pop.title("Edit Vendor")
        pop.transient(self.root)
        pop.grab_set()
        self._lock_popup_position(pop, 560, 330)
        self._animate_popup_in(pop)

        frame = ttk.Frame(pop, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        vendor_v = tk.StringVar(value=subtask["vendor"])
        channel_v = tk.StringVar(value=subtask["channel"])
        platforms_v = tk.StringVar(value=subtask["platforms"])
        type_v = tk.StringVar(value=subtask["ad_type"])
        qty_v = tk.StringVar(value=subtask["qty"])
        price_v = tk.StringVar(value=subtask["price"])
        details_v = tk.StringVar(value=subtask["details"])

        fields = [
            ("Vendor", vendor_v),
            ("Channel", channel_v),
            ("Platforms", platforms_v),
            ("Type", type_v),
            ("Qty", qty_v),
            ("Price", price_v),
            ("Details", details_v),
        ]
        for i, (label, var) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", pady=4, padx=(0, 8))
            ttk.Entry(frame, textvariable=var).grid(row=i, column=1, sticky="ew", pady=4)

        def _save_edit():
            self.db.update_subtask(
                sid,
                vendor_v.get().strip(),
                channel_v.get().strip(),
                platforms_v.get().strip(),
                type_v.get().strip() or "Store Visit",
                qty_v.get().strip() or "1",
                details_v.get().strip(),
                f"{self._parse_price(price_v.get()):.2f}",
            )
            self._audit("update_subtask", "subtask", str(sid), details=f"vendor={vendor_v.get().strip()}")
            self.refresh_subtasks(tid)
            self.save_task(silent=True)
            self.refresh_task_rows()
            pop.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=pop.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Save", command=_save_edit).pack(side="left")


    def refresh_platform_detail_inputs(self):
        if not hasattr(self, "platform_detail_box"):
            return
        for child in self.platform_detail_box.winfo_children():
            child.destroy()

        selected = [p for p, v in self.platform_vars.items() if v.get()]
        if not selected:
            ttk.Label(self.platform_detail_box, text="Select one or more platforms to add platform names.").grid(row=0, column=0, sticky="w")
            return

        for i, platform in enumerate(selected):
            row = ttk.Frame(self.platform_detail_box)
            row.grid(row=i, column=0, sticky="ew", pady=(0, 4))
            row.columnconfigure(1, weight=1)
            ttk.Label(row, text=f"{platform} Name").grid(row=0, column=0, sticky="w", padx=(0, 6))
            ttk.Entry(row, textvariable=self.platform_channel_vars[platform]).grid(row=0, column=1, sticky="ew")

    def _serialize_platform_channels(self, selected_platforms):
        parts = []
        for p in selected_platforms:
            val = self.platform_channel_vars.get(p, tk.StringVar()).get().strip() if p in self.platform_channel_vars else ""
            if val:
                parts.append(f"{p}: {val}")
        return " | ".join(parts)

    def _load_platform_channels(self, channel_text: str):
        for var in self.platform_channel_vars.values():
            var.set("")
        raw = (channel_text or "").strip()
        if not raw:
            return
        for part in raw.split("|"):
            part = part.strip()
            if ":" not in part:
                continue
            platform, value = [x.strip() for x in part.split(":", 1)]
            if platform in self.platform_channel_vars:
                self.platform_channel_vars[platform].set(value)

    def toggle_details(self, *_):
        if self.type_sub_var.get() == "Multi Service":
            self.detail_entry.grid()
        else:
            self.detail_entry.grid_remove()
            self.detail_var.set("")

    def _validate_generation_context(self, task, subtask, context):
        required = {
            "brand": task["brand"].strip(),
            "amount": task["amount"].strip(),
            "vendor": subtask["vendor"].strip(),
            "channel": subtask["channel"].strip(),
            "platform": subtask["platforms"].strip(),
            "license": context["license_number"].strip(),
            "iban": context["iban"].strip(),
        }
        missing = [k for k, v in required.items() if not v]
        if missing:
            return f"Vendor '{subtask['vendor'] or '-'}' is missing: {', '.join(missing)}"
        return None

    def _context_for_subtask(self, task, subtask):
        platforms = [p.strip() for p in subtask["platforms"].split(",") if p.strip()]
        raw_channel = (subtask["channel"] or "").strip()
        channel_payload = raw_channel
        if not channel_payload and platforms:
            channel_payload = ", ".join(platforms)

        bank_info = self.db.get_bank_info_by_iban(subtask["iban"])

        if bank_info:
            bank_name = bank_info["bank_name"]
            account_name = bank_info["account_name"]
            account_number = bank_info["account_number"]
            swift_code = bank_info["swift_code"]
        else:
            bank_name = ""
            account_name = ""
            account_number = ""
            swift_code = ""

        ad_types = f"{subtask['ad_type']}, {subtask['qty']}" if subtask["qty"] else subtask["ad_type"]

        return {
            "brand_name": task["brand"],
            "amount": subtask["price"],
            "contract_type": (task["contract_type"].strip() if task["contract_type"].strip() not in {"", "auto"} else "after_pay"),
            "channel_name": channel_payload,
            "platform": ", ".join(sorted(set(platforms))),
            "ad_types": ad_types,
            "influencer_name_as_per_license": subtask["vendor"],
            "license_number": subtask["license_number"],
            "city_as_per_license": "",
            "neighbourhood_as_per_license": "",
            "bank_name": bank_name,
            "account_name": account_name,
            "iban" : subtask["iban"],
            "account_number": account_number,
            "swift_code": swift_code,
        }

    def _generate_contracts_for_task(self, task_id: str, selected_subtask_ids=None, generate_all=False):

        self.save_task(silent=True)

        task = self.db.get_task(task_id)
        subtasks = self.db.list_subtasks(task_id)

        if not task:
            messagebox.showerror("Missing task", "Task data was not found in database")
            return

        if not subtasks:
            messagebox.showwarning("No vendors", "Add vendors before generating contracts")
            return

        selected_ids = {int(i) for i in (selected_subtask_ids or []) if str(i).isdigit()}

        if generate_all:
            selected_ids = set()

        seen_ids = set()
        target_subtasks = []

        for row in subtasks:

            sid = int(row["id"])

            if sid in seen_ids:
                continue

            if selected_ids and sid not in selected_ids:
                continue

            seen_ids.add(sid)
            target_subtasks.append(row)

        if not target_subtasks:
            messagebox.showwarning("No vendors", "No matching vendor rows were selected for generation")
            return

        progress_popup = tk.Toplevel(self.root)
        progress_popup.title("Generating contracts")
        progress_popup.transient(self.root)
        progress_popup.grab_set()

        self._lock_popup_position(progress_popup, 420, 140)

        progress_frame = ttk.Frame(progress_popup, padding=12)
        progress_frame.pack(fill="both", expand=True)

        progress_label = ttk.Label(progress_frame, text=f"Generating 0/{len(target_subtasks)}...")
        progress_label.pack(anchor="w", pady=(0, 8))

        progress = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            maximum=len(target_subtasks),
        )

        progress.pack(fill="x")

        progress_popup.update_idletasks()

        failures = []
        generated_ids = []
        generated_pdf_paths = []
        generated_rows = []
        generated_docx_paths = []
        output_folder = None

        use_batch_pdf = len(target_subtasks) > 1

        try:

            for idx, subtask in enumerate(target_subtasks, start=1):

                progress_label.configure(
                    text=f"Generating {idx}/{len(target_subtasks)}: {subtask['vendor']}"
                )

                progress.configure(value=idx)
                progress_popup.update_idletasks()

                context = self._context_for_subtask(task, subtask)

                error = self._validate_generation_context(task, subtask, context)

                if error:
                    failures.append(error)
                    continue

                try:

                    result = generate_contract_from_gui(
                        context,
                        append_excel=False,
                        return_metadata=True,
                        convert_pdf=not use_batch_pdf,
                    )

                    cid = result["contract_id"]

                    self.db.log_generated_contract(
                        cid,
                        task_id,
                        task["brand"],
                        task["amount"],
                        task["contract_type"],
                    )

                    self._audit(
                        "generate_contract",
                        "contract",
                        cid,
                        details=f"task={task_id};vendor={subtask['vendor']}",
                    )

                    generated_ids.append(cid)

                    generated_rows.append(result["row"])

                    output_folder = output_folder or result.get("today_folder")

                    if result.get("docx_path"):
                        generated_docx_paths.append(result["docx_path"])

                    if result.get("pdf_path"):
                        generated_pdf_paths.append(Path(result["pdf_path"]))


                except Exception as exc:

                    failures.append(
                        f"Vendor '{subtask['vendor']}' failed: {exc}"
                    )

            if generated_rows:
                append_generated_rows(generated_rows)

            if use_batch_pdf and generated_docx_paths and output_folder:
                progress_label.configure(text="Converting PDFs...")

                progress_popup.update_idletasks()

                generated_pdf_paths.extend(
                    batch_convert_docx_to_pdf(
                        generated_docx_paths,
                        output_folder,
                    )
                )


        finally:

            if progress_popup.winfo_exists():
                progress_popup.destroy()

        if generated_ids:
            summary = "\n".join(generated_ids[:10])

            messagebox.showinfo(
                "Success",
                f"Generated {len(generated_ids)} contract(s):\n{summary}",
            )

        if failures:
            messagebox.showwarning(
                "Generation issues",
                "\n".join(failures[:10]),
            )


    def generate_selected_task_contracts(self):
        if not self._require_permission("generate", "Generate contract"):
            return
        selected = self.task_tree.selection()
        if not selected:
            messagebox.showwarning("No task", "Right-click a task first")
            return
        task_id = selected[0]
        self._generate_contracts_for_task(task_id, generate_all=True)

    def generate_selected_vendor_contract(self):
        if not self._require_permission("generate", "Generate contract"):
            return
        tid = self.current_task.get()
        if not tid:
            messagebox.showwarning("No task", "Select a task first")
            return
        selected_subtasks = list(self.sub_tree.selection())
        if not selected_subtasks:
            messagebox.showwarning("No vendor", "Right-click a vendor row first")
            return
        self._generate_contracts_for_task(tid, selected_subtask_ids=selected_subtasks)

    def generate_contract(self):
        if not self._require_permission("generate", "Generate contract"):
            return
        tid = self.current_task.get()
        if not tid:
            messagebox.showwarning("No task", "Select a task first")
            return
        selected_subtasks = list(self.sub_tree.selection())
        self._generate_contracts_for_task(tid, selected_subtask_ids=selected_subtasks)


    def on_close(self):
        self.db.close()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    ContractSuiteApp().run()
