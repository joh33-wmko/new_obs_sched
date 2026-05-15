# New Semester Observing Schedule
# Step 11: Staff Schedule Uploader
# This script converts an Excel file to CSV format and optionally uploads it to the staff schedule site.

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import datetime
import importlib
import ast
import html
import re
import shutil
import socket
import ssl
import subprocess
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.poolmanager import PoolManager

def load_live_config(config_path=None):
    config_file = Path(config_path) if config_path else Path.cwd() / "config.live.ini"
    if not config_file.exists():
        return {}

    content = config_file.read_text(encoding="utf-8").strip()
    if not content:
        return {}

    try:
        parsed = ast.literal_eval(content)
    except (SyntaxError, ValueError):
        parsed = None

    if isinstance(parsed, dict):
        return parsed

    namespace = {}
    try:
        exec(compile(content, str(config_file), "exec"), {"__builtins__": {}}, namespace)
    except Exception:
        return {}

    return {k: v for k, v in namespace.items() if not k.startswith("_")}


LIVE_CONFIG = load_live_config()
SSH_PASSWORD_CACHE = {}


def get_config(section, key, default=None):
    """Get a configuration value from LIVE_CONFIG."""
    return LIVE_CONFIG.get(section, {}).get(key, default)


def get_ssh_password(settings):
    """Get SSH password from cache or prompt user for it."""
    cache_key = (settings.get("ssh_host"), settings.get("ssh_user"), settings.get("ssh_port"))
    cached = SSH_PASSWORD_CACHE.get(cache_key)
    if cached:
        return cached

    parent = getattr(tk, "_default_root", None)
    display_host = settings.get("ssh_hostname") or settings.get("ssh_host")
    prompt = (
        f"Enter SSH password for {settings.get('ssh_user')}@{display_host}"
        #f" (port {settings.get('ssh_port')})"
    )
    password = simpledialog.askstring("SSH Password Required", prompt, show="*", parent=parent)
    if not password:
        raise RuntimeError("SSH password entry was cancelled.")

    SSH_PASSWORD_CACHE[cache_key] = password
    return password


class LegacyTLSAdapter(HTTPAdapter):
    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        context = ssl.create_default_context()
        context.set_ciphers("DEFAULT@SECLEVEL=1")
        self.poolmanager = PoolManager(
            num_pools=connections,
            maxsize=maxsize,
            block=block,
            ssl_context=context,
            **pool_kwargs,
        )


def _set_dialog_geometry(dialog, width, height, min_width=420, min_height=220):
    """Clamp dialog size to screen and center it so controls stay visible."""
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()

    max_width = max(min_width, screen_width - 80)
    max_height = max(min_height, screen_height - 120)

    final_width = max(min_width, min(int(width), int(max_width)))
    final_height = max(min_height, min(int(height), int(max_height)))

    pos_x = max(0, (screen_width - final_width) // 2)
    pos_y = max(0, (screen_height - final_height) // 3)
    dialog.geometry(f"{final_width}x{final_height}+{pos_x}+{pos_y}")


def _estimate_message_dialog_size(text, base_width=640, base_height=280):
    lines = text.splitlines() or [""]
    longest_line = max((len(line) for line in lines), default=0)
    line_count = len(lines)

    # Grow width for long lines and height for multiline messages, capped by screen clamp.
    width = base_width + min(320, max(0, longest_line - 70) * 4)
    height = base_height + min(300, max(0, line_count - 10) * 14)
    return width, height


def show_scrollable_text_dialog(title, text):
    dialog = tk.Toplevel()
    dialog.title(title)
    _set_dialog_geometry(dialog, 1000, 620, min_width=760, min_height=420)
    dialog.resizable(True, True)

    frame = tk.Frame(dialog)
    frame.pack(fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side="right", fill="y")

    text_widget = tk.Text(frame, wrap="none", yscrollcommand=scrollbar.set)
    text_widget.insert("1.0", text)
    text_widget.config(state="disabled")
    text_widget.pack(side="left", fill="both", expand=True)

    scrollbar.config(command=text_widget.yview)

    btn = tk.Button(dialog, text="Close", command=dialog.destroy)
    btn.pack(pady=8)

    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    dialog.grab_set()
    dialog.wait_window()


def show_focused_info_dialog(title, text, parent=None):
    owner = parent or getattr(tk, "_default_root", None)
    dialog = tk.Toplevel(owner)
    dialog.title(title)
    est_width, est_height = _estimate_message_dialog_size(text, base_width=660, base_height=300)
    _set_dialog_geometry(dialog, est_width, est_height, min_width=520, min_height=260)
    dialog.resizable(True, True)

    wrap_length = max(460, min(900, est_width - 40))
    tk.Label(
        dialog,
        text=text,
        padx=16,
        pady=16,
        justify="left",
        anchor="w",
        wraplength=wrap_length,
    ).pack(fill="both", expand=True)

    tk.Button(dialog, text="OK", command=dialog.destroy).pack(pady=(0, 12))

    dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    dialog.grab_set()
    dialog.wait_window()


def show_focused_yes_no_dialog(title, text, parent=None):
    owner = parent or getattr(tk, "_default_root", None)
    dialog = tk.Toplevel(owner)
    dialog.title(title)
    est_width, est_height = _estimate_message_dialog_size(text, base_width=660, base_height=300)
    _set_dialog_geometry(dialog, est_width, est_height, min_width=520, min_height=260)
    dialog.resizable(True, True)

    answer = {"value": False}

    wrap_length = max(460, min(900, est_width - 40))
    tk.Label(
        dialog,
        text=text,
        padx=16,
        pady=16,
        justify="left",
        anchor="w",
        wraplength=wrap_length,
    ).pack(fill="both", expand=True)

    button_row = tk.Frame(dialog)
    button_row.pack(pady=(0, 12))

    def choose_yes():
        answer["value"] = True
        dialog.destroy()

    def choose_no():
        answer["value"] = False
        dialog.destroy()

    tk.Button(button_row, text="Yes", command=choose_yes).pack(side="left", padx=6)
    tk.Button(button_row, text="No", command=choose_no).pack(side="left", padx=6)

    dialog.protocol("WM_DELETE_WINDOW", choose_no)
    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    dialog.grab_set()
    dialog.wait_window()
    return answer["value"]


def apply_insert_statements_to_database(connection, sql_file):
    """Execute INSERT statements from SQL file and track results."""
    stats = {
        'total': 0,
        'success': 0,
        'failed': 0,
        'already_exists': 0,
        'statements': []  # List of (statement, status, error) tuples
    }
    
    sql_text = sql_file.read_text(encoding="utf-8")
    statements = extract_insert_statements(sql_text)
    
    if not statements:
        stats['total'] = 0
        return stats
    
    cursor = connection.cursor()
    try:
        for stmt in statements:
            stats['total'] += 1
            status = None
            error_msg = None
            
            try:
                cursor.execute(stmt)
                connection.commit()
                stats['success'] += 1
                status = 'SUCCESS'
            except Exception as e:
                error_msg = str(e)
                # Check if it's a duplicate key error (already exists)
                if 'Duplicate entry' in error_msg or 'duplicate' in error_msg.lower():
                    stats['already_exists'] += 1
                    status = 'DUPLICATE'
                else:
                    stats['failed'] += 1
                    status = 'FAILED'
            
            stats['statements'].append((stmt, status, error_msg))
    finally:
        cursor.close()
    
    return stats


def write_insert_results_to_file(stats, sql_file, data_dir):
    """Write INSERT execution results to a detailed report file."""
    result_file = data_dir / f"{sql_file.stem}_results.txt"
    
    with open(result_file, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("DATABASE INSERT EXECUTION RESULTS\n")
        f.write("=" * 80 + "\n\n")
        
        f.write("SUMMARY\n")
        f.write("-" * 80 + "\n")
        f.write(f"Total Statements: {stats['total']}\n")
        f.write(f"✓ Inserted Successfully: {stats['success']}\n")
        f.write(f"⊘ Already Existed (Skipped): {stats['already_exists']}\n")
        f.write(f"✗ Failed: {stats['failed']}\n\n")
        
        f.write("DETAILED RESULTS\n")
        f.write("-" * 80 + "\n\n")
        
        for idx, (stmt, status, error) in enumerate(stats['statements'], 1):
            f.write(f"[{idx}/{stats['total']}] {status}\n")
            f.write(f"  {stmt[:100]}{'...' if len(stmt) > 100 else ''}\n")
            if error:
                f.write(f"  ERROR: {error[:150]}{'...' if len(error) > 150 else ''}\n")
            f.write("\n")
    
    return result_file


def show_insert_statements_preview():
    # Find the last generated SQL file
    data_dir = ensure_data_dir()
    processed_type = detect_processed_report_type(data_dir)
    tracker_file = data_dir / "last_generated_csv_filename.txt"
    if not tracker_file.exists():
        show_focused_info_dialog("No SQL file", "No last_generated_csv_filename.txt found.")
        return
    csv_stem = tracker_file.read_text(encoding="utf-8").strip().rsplit(".csv", 1)[0]
    sql_file = data_dir / f"{csv_stem}.sql"
    if not sql_file.exists():
        show_focused_info_dialog("No SQL file", f"No SQL file found: {sql_file}")
        return
    sql_text = sql_file.read_text(encoding="utf-8")
    # Extract INSERT statements
    statements = extract_insert_statements(sql_text)
    if not statements:
        show_focused_info_dialog("No INSERTs found", f"No INSERT statements found in {sql_file.name}.")
        return

    # Show in a scrollable dialog
    title = "Preview INSERT Statements"
    if processed_type and processed_type != "UNKNOWN":
        title = f"Preview {processed_type} INSERT Statements"
    show_scrollable_text_dialog(title, "\n\n".join(statements))


def _clean_header_value(value):
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text.startswith("Unnamed:"):
        return ""

    return text


def _flatten_headers(columns):
    flattened = []
    current_prefix = ""

    for index, column in enumerate(columns, start=1):
        if isinstance(column, tuple):
            parts = [_clean_header_value(part) for part in column]
        else:
            parts = [_clean_header_value(column)]

        if parts and parts[0]:
            current_prefix = parts[0]
        elif current_prefix and parts:

            parts[0] = current_prefix

        combined = []
        for part in parts:
            if part and (not combined or combined[-1] != part):
                combined.append(part)

        flattened.append(" ".join(combined) or f"column_{index}")

    return flattened


def _normalize_semester_token(stem):
    match = re.search(r"(?<!\d)(\d{2}|\d{4})\s*([abAB])(?![A-Za-z0-9])", stem)
    if not match:
        return None

    year = match.group(1)
    semester = match.group(2).upper()
    if len(year) == 2:
        year = f"20{year}"

    return f"{year}{semester}"


def _normalize_staff_type(stem):
    match = re.search(r"(?<![A-Za-z0-9])(OA|NA|SA)(?![A-Za-z0-9])", stem, flags=re.IGNORECASE)
    if not match:
        return None

    return match.group(1).upper()


def _normalize_month_range(stem):
    month_pattern = (
        r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
    )
    match = re.search(rf"{month_pattern}\s*-\s*{month_pattern}", stem, flags=re.IGNORECASE)
    if not match:
        return None

    month_map = {
        "jan": "Jan",
        "feb": "Feb",
        "mar": "Mar",
        "apr": "Apr",
        "may": "May",
        "jun": "Jun",
        "jul": "Jul",
        "aug": "Aug",
        "sep": "Sep",
        "oct": "Oct",
        "nov": "Nov",
        "dec": "Dec",
    }

    start = month_map[match.group(1)[:3].lower()]
    end = month_map[match.group(2)[:3].lower()]
    return f"{start}-{end}"


def build_output_filename(file_path):
    stem = Path(file_path).stem
    semester = _normalize_semester_token(stem)
    staff_type = _normalize_staff_type(stem)
    month_range = _normalize_month_range(stem)

    if semester and staff_type:
        filename = f"{semester}_{staff_type}_Schedule"
        if month_range:
            filename = f"{filename}_{month_range}"
        return f"{filename}.csv"

    return f"{Path(file_path).stem}.csv"


def ensure_data_dir():
    data_dir = Path.cwd() / "data"
    data_dir.mkdir(exist_ok=True)
    return data_dir


def write_db_report(report_text, report_type, data_dir):
    """Append a database report to Step11_db_report.txt with type and timestamp."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_file = data_dir / "Step11_db_report.txt"
    with open(report_file, "a", encoding="utf-8") as handle:
        handle.write(f"[{report_type}] {timestamp}\n")
        handle.write(report_text.strip() + "\n\n")
    return report_file


def detect_processed_report_type(data_dir):
    """Infer processed report type from last generated CSV filename."""
    tracker_file = data_dir / "last_generated_csv_filename.txt"
    if not tracker_file.exists():
        return "UNKNOWN"

    filename = tracker_file.read_text(encoding="utf-8").strip()
    stem = Path(filename).stem
    match = re.search(
        r"(?<![A-Za-z0-9])(OA|NA|SA|SWOC|EEOC)(?![A-Za-z0-9])",
        stem,
        flags=re.IGNORECASE,
    )
    return match.group(1).upper() if match else "UNKNOWN"


def fetch_db_snapshot(cursor, main_table, type_columns, staff_types):
    """Fetch total rows and per-type counts for a table snapshot."""
    cursor.execute(f"SELECT COUNT(*) FROM `{main_table}`;")
    total_rows = cursor.fetchone()[0]

    type_counts = {}
    cursor.execute(f"SHOW COLUMNS FROM `{main_table}`;")
    available_columns = {row[0].lower() for row in cursor.fetchall()}
    type_column = next(
        (name for name in type_columns if name.lower() in available_columns),
        None,
    )

    if type_column:
        for staff_type in staff_types:
            cursor.execute(
                f"SELECT COUNT(*) FROM `{main_table}` WHERE LOWER(`{type_column}`) LIKE %s;",
                (f"%{staff_type}%",),
            )
            type_counts[staff_type] = cursor.fetchone()[0]

    return total_rows, type_counts


def copy_source_to_data_dir(file_path, data_dir):
    source = Path(file_path)
    destination = data_dir / source.name

    if source.resolve() != destination.resolve():
        shutil.copy2(source, destination)

    return destination


def store_last_csv_filename(output_path, data_dir):
    """Record the last generated CSV filename for reference."""
    tracker_file = data_dir / "last_generated_csv_filename.txt"
    tracker_file.write_text(f"{output_path.name}\n", encoding="utf-8")


def infer_upload_type(xlsx_file_path=None, csv_file_path=None):
    """Detect staff type from filename or return default."""
    if xlsx_file_path:
        xlsx_staff = _normalize_staff_type(Path(xlsx_file_path).stem)
        if xlsx_staff:
            return xlsx_staff.lower()

    if csv_file_path:
        csv_staff = _normalize_staff_type(Path(csv_file_path).stem)
        if csv_staff:
            return csv_staff.lower()

    return get_config("NEW_OBS_SEM", "DEFAULT_UPLOAD_TYPE", "eeoc")


def get_mysql_connection_settings(section_name):
    db_config = LIVE_CONFIG.get(section_name, {})

    host = db_config.get("DB_HOST")
    user = db_config.get("DB_USER")
    password = db_config.get("DB_PASS")
    database = db_config.get("DB_NAME")

    port_value = db_config.get("DB_PORT") or 3306
    try:
        port = int(port_value)
    except (TypeError, ValueError):
        port = 3306

    ssh_port_value = db_config.get("SSH_PORT", 22)
    try:
        ssh_port = int(ssh_port_value)
    except (TypeError, ValueError):
        ssh_port = 22

    mysql_remote_port_value = db_config.get("MYSQL_REMOTE_PORT", 3306)
    try:
        mysql_remote_port = int(mysql_remote_port_value)
    except (TypeError, ValueError):
        mysql_remote_port = 3306

    return {
        "section": section_name,
        "host": host,
        "user": user,
        "password": password,
        "database": database,
        "port": port,
        "ssh_tunnel_enabled": bool(db_config.get("SSH_HOST") and db_config.get("SSH_USER")),
        "ssh_host": db_config.get("SSH_HOST"),
        "ssh_hostname": db_config.get("SSH_HOSTNAME"),
        "ssh_user": db_config.get("SSH_USER"),
        "ssh_port": ssh_port,
        "ssh_key_file": db_config.get("SSH_KEY_FILE") or LIVE_CONFIG.get("SSH_KEY_FILE"),
        "mysql_remote_host": db_config.get("MYSQL_REMOTE_HOST") or "127.0.0.1",
        "mysql_remote_port": mysql_remote_port,
    }


def _is_tcp_port_open(host, port, timeout=1.0):
    try:
        with socket.create_connection((host, int(port)), timeout=timeout):
            return True
    except OSError:
        return False


def _start_ssh_tunnel(settings):
    if not settings.get("ssh_tunnel_enabled"):
        return None

    local_host = settings.get("host") or "127.0.0.1"
    local_port = settings.get("port") or 3306

    # If something is already listening on the configured local endpoint,
    # assume an existing tunnel is active and reusable.
    if _is_tcp_port_open(local_host, local_port):
        return None

    ssh_target = f"{settings['ssh_user']}@{settings['ssh_host']}"
    bind_spec = (
        f"{local_host}:{local_port}:"
        f"{settings['mysql_remote_host']}:{settings['mysql_remote_port']}"
    )
    
    # Try to use sshpass if available (for automated password auth)
    sshpass_available = shutil.which("sshpass") is not None
    
    if sshpass_available:
        # Use sshpass with a runtime password prompt.
        ssh_password = get_ssh_password(settings)
        
        command = [
            "sshpass",
            "-p",
            ssh_password,
            "ssh",
            "-N",
            "-L",
            bind_spec,
            "-p",
            str(settings["ssh_port"]),
            "-o", "StrictHostKeyChecking=accept-new",
            "-o", "ExitOnForwardFailure=yes",
            "-o", "ConnectTimeout=10",
            "-o", "UserKnownHostsFile=/dev/null",
            ssh_target,
        ]
    else:
        # Fallback: provide manual instructions
        manual_cmd = (
            f"ssh -N -L {bind_spec} -p {settings['ssh_port']} {ssh_target}"
        )
        raise RuntimeError(
            "SSH tunnel requires sshpass for automated password authentication.\n\n"
            "Install sshpass:\n"
            "  brew install sshpass\n\n"
            "Or set up manually in another terminal:\n"
            f"  {manual_cmd}\n"
            "Then re-run this script."
        )

    # Start tunnel process
    try:
        process = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as error:
        raise RuntimeError(
            f"Failed to start SSH tunnel with sshpass: {error}"
        )

    # Wait for tunnel to become ready (up to 90 seconds)
    max_wait_iterations = 180
    for iteration in range(max_wait_iterations):
        # Check if process exited early (indicates failure)
        if process.poll() is not None:
            exit_code = process.returncode
            if exit_code != 0:
                raise RuntimeError(
                    f"SSH tunnel exited with code {exit_code}.\n"
                    "Possible causes:\n"
                    "  - Wrong SSH password in config\n"
                    "  - Remote host unreachable\n"
                    "  - SSH key authorization issue\n\n"
                    "Verify credentials and connectivity, then retry."
                )
            break
        
        # Check if port is now open
        if _is_tcp_port_open(local_host, local_port, timeout=0.5):
            return process
        
        time.sleep(0.5)

    # Timeout reached
    process.terminate()
    try:
        process.wait(timeout=2)
    except subprocess.TimeoutExpired:
        process.kill()
    
    raise RuntimeError(
        f"SSH tunnel did not open local port {local_port} within {max_wait_iterations * 0.5:.0f}s.\n"
        "The SSH process may have encountered an authentication issue.\n"
        "Verify SSH password and connectivity, then retry."
    )


def _stop_ssh_tunnel(process):
    if process is None:
        return
    if process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=3)
        except subprocess.TimeoutExpired:
            process.kill()


def get_mysql_connection_candidates():
    candidates = []
    settings = get_mysql_connection_settings("DB_SERVER")
    if settings.get("host") and settings.get("user") and settings.get("password"):
        candidates.append(settings)
    return candidates


def connect_to_remote_mysql_db():
    candidates = get_mysql_connection_candidates()
    if not candidates:
        show_focused_info_dialog(
            "Missing MySQL settings",
            "No MySQL connection sections found in config.live.ini.\n"
            "Expected section: DB_SERVER."
        )
        return False

    try:
        pymysql = importlib.import_module("pymysql")
    except ImportError:
        show_focused_info_dialog(
            "Missing MySQL client",
            "PyMySQL is not installed. Install it before trying to connect to the remote MySQL database."
        )
        return False

    failures = []
    for settings in candidates:
        tunnel_process = None
        missing = [name for name in ("host", "user", "password") if not settings.get(name)]
        if missing:
            failures.append(
                f"[{settings.get('section')}] missing required keys: {', '.join(missing)}"
            )
            continue

        connection = None
        try:
            tunnel_process = _start_ssh_tunnel(settings)
            db_name = settings["database"]
            connect_timeout = get_config("DB_SERVER", "CONNECT_TIMEOUT", 10)
            
            connection = pymysql.connect(
                host=settings["host"],
                user=settings["user"],
                password=settings["password"],
                database=db_name,
                port=settings["port"],
                connect_timeout=connect_timeout,
            )
            type_columns = get_config("DB_SERVER", "TYPE_COLUMNS", [])
            staff_types = get_config("DB_SERVER", "STAFF_TYPES", [])
            with connection.cursor() as cursor:
                cursor.execute("SHOW TABLES;")
                tables = [row[0] for row in cursor.fetchall()]
                
                # Find main table from preferred list
                preferred = get_config("DB_SERVER", "PREFERRED_TABLES", [])
                main_table = next((t for t in tables if t.lower() in [p.lower() for p in preferred]), None)
                if not main_table and tables:
                    main_table = tables[0]
                if not main_table:
                    show_focused_info_dialog("No tables found", f"No tables found in {db_name} database.")
                    return False

                total_rows, type_counts = fetch_db_snapshot(
                    cursor,
                    main_table,
                    type_columns,
                    staff_types,
                )

                # Get MySQL version
                cursor.execute("SELECT VERSION()")
                version_row = cursor.fetchone()
                mysql_version = version_row[0] if version_row else "unknown"

            type_summary = (
                "\n".join(f"{k.upper()}: {v}" for k, v in type_counts.items())
                if type_counts
                else "Type breakdown unavailable (no type-like column found)."
            )

            msg = (
                f"Connected using {settings['section']} tunnel {settings['host']}:{settings['port']}\n"
                f"MySQL Server version: {mysql_version}\n\n"
                f"Database: {db_name}\nTable: {main_table}\n\n"
                f"Total rows: {total_rows}\n"
                + type_summary
            )
            show_focused_info_dialog("Current MySQL Table Stats", msg, parent=getattr(tk, "_default_root", None))
            write_db_report(msg, "BEFORE_INSERT", ensure_data_dir())

            # Show the generated INSERT statements from the last SQL file
            show_insert_statements_preview()
            
            # Ask user if they want to apply the inserts
            apply_choice = show_focused_yes_no_dialog(
                "Apply INSERTs Info",
                "Apply the INSERT statements to the database?\n\n"
                "This will:\n"
                "• Insert new records\n"
                "• Skip records that already exist (duplicate dates)\n"
                "• Report statistics"
            )
            
            if apply_choice:
                # Get the SQL file path
                data_dir = ensure_data_dir()
                tracker_file = data_dir / "last_generated_csv_filename.txt"
                if tracker_file.exists():
                    csv_stem = tracker_file.read_text(encoding="utf-8").strip().rsplit(".csv", 1)[0]
                    sql_file = data_dir / f"{csv_stem}.sql"
                    if sql_file.exists():
                        # Apply the inserts
                        stats = apply_insert_statements_to_database(connection, sql_file)
                        processed_type = detect_processed_report_type(data_dir)

                        with connection.cursor() as snapshot_cursor:
                            post_total_rows, post_type_counts = fetch_db_snapshot(
                                snapshot_cursor,
                                main_table,
                                type_columns,
                                staff_types,
                            )
                        post_type_summary = (
                            "\n".join(f"{k.upper()}: {v}" for k, v in post_type_counts.items())
                            if post_type_counts
                            else "Type breakdown unavailable (no type-like column found)."
                        )
                        end_delimiter = (
                            "\n" + ("=" * 72) + "\n"
                            + f"END OF PROCESSING FOR TYPE: {processed_type}\n"
                            + ("=" * 72)
                        )
                        
                        # Write results to file
                        result_file = write_insert_results_to_file(stats, sql_file, data_dir)
                        
                        # Show results
                        result_msg = (
                            f"Database Insert Results for {processed_type}\n\n"
                            f"Total statements: {stats['total']}\n"
                            f"✓ Inserted successfully: {stats['success']}\n"
                            f"⊘ Already existed (skipped): {stats['already_exists']}\n"
                            f"✗ Failed: {stats['failed']}\n\n"
                            f"After Insert Snapshot\n"
                            f"Table: {main_table}\n\n"
                            f"Total rows: {post_total_rows}\n"
                            f"{post_type_summary}\n\n"
                            f"Detailed results saved to:\n{result_file.name}\n"
                            f"{end_delimiter}"
                        )
                        
                        show_focused_info_dialog("Insert Results", result_msg)
                        write_db_report(result_msg, "AFTER_INSERT", ensure_data_dir())
            
            return True
        except Exception as error:
            failures.append(
                f"[{settings.get('section')}] {settings.get('user')}@{settings.get('host')}:{settings.get('port')} -> {error}"
            )
        finally:
            if connection is not None:
                connection.close()
            _stop_ssh_tunnel(tunnel_process)

    show_focused_info_dialog(
        "Database connection failed",
        "Could not connect to the remote MySQL database with any configured profile.\n\n"
        + "\n".join(failures)
    )
    return False


def preflight_mysql_connection_check():
    candidates = get_mysql_connection_candidates()
    if not candidates:
        show_focused_info_dialog(
            "Missing MySQL settings",
            "No MySQL connection sections found in config.live.ini.\n"
            "Expected section: DB_SERVER."
        )
        return False

    try:
        pymysql = importlib.import_module("pymysql")
    except ImportError:
        show_focused_info_dialog(
            "Missing MySQL client",
            "PyMySQL is not installed. Install it before trying database import."
        )
        return False

    failures = []
    for settings in candidates:
        tunnel_process = None
        missing = [name for name in ("host", "user", "password") if not settings.get(name)]
        if missing:
            failures.append(
                f"[{settings.get('section')}] missing required keys: {', '.join(missing)}"
            )
            continue

        connection = None
        try:
            tunnel_process = _start_ssh_tunnel(settings)
            query_timeout = get_config("DB_SERVER", "QUERY_TIMEOUT", 5)
            connection = pymysql.connect(
                host=settings["host"],
                user=settings["user"],
                password=settings["password"],
                port=settings["port"],
                connect_timeout=query_timeout,
            )
            with connection.cursor() as cursor:
                cursor.execute("SELECT 1")
                cursor.fetchone()
            return True
        except Exception as error:
            failures.append(
                f"[{settings.get('section')}] {settings.get('user')}@{settings.get('host')}:{settings.get('port')} -> {error}"
            )
        finally:
            if connection is not None:
                connection.close()
            _stop_ssh_tunnel(tunnel_process)

    show_focused_info_dialog(
        "Database preflight failed",
        "Database authentication failed before import.\n\n"
        "Fix credentials/grants, then rerun this step.\n\n"
        + "\n".join(failures)
    )
    return False


def upload_csv_to_staff_site(csv_path, upload_type, verbose_enabled):
    """Upload CSV file to staff schedule service."""
    session = requests.Session()
    session.mount("https://", LegacyTLSAdapter())

    form_data = {"type": upload_type, "submit": "Submit"}
    if verbose_enabled:
        form_data["verbose"] = "on"

    upload_url = get_config("NEW_OBS_SEM", "STAFF_UPLOAD_URL")
    with open(csv_path, "rb") as file_handle:
        response = session.post(
            str(upload_url),
            data=form_data,
            files={"file": (Path(csv_path).name, file_handle, "text/csv")},
            timeout=60,
        )

    response.raise_for_status()
    return response


def open_file_in_integrated_browser(output_file):
    # Keep focus in Finder on macOS instead of foregrounding VS Code.
    if sys.platform == "darwin":
        try:
            subprocess.Popen(["open", "-R", str(output_file)])
            return output_file
        except Exception:
            subprocess.Popen(["open", str(output_file)])
            return output_file

    return output_file


def open_response_in_integrated_browser(response_text, output_file):
    output_file.write_text(response_text, encoding="utf-8")
    return open_file_in_integrated_browser(output_file)


def extract_insert_statements(response_text):
    normalized = response_text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
    plain_text = html.unescape(re.sub(r"<[^>]+>", "\n", normalized))
    statements = re.findall(r"(?is)\binsert\b.*?;", plain_text)
    cleaned = [re.sub(r"\s+", " ", stmt).strip() for stmt in statements]
    return cleaned


def save_insert_statements(response_text, csv_output_path, data_dir):
    sql_output = data_dir / f"{Path(csv_output_path).stem}.sql"
    statements = extract_insert_statements(response_text)

    if statements:
        sql_output.write_text("\n".join(statements) + "\n", encoding="utf-8")
    else:
        sql_output.write_text("-- No INSERT statements found in upload response.\n", encoding="utf-8")

    return sql_output


def pick_file(root):
    root.deiconify()
    root.lift()
    root.focus_force()
    root.attributes("-topmost", True)
    root.update_idletasks()

    try:
        return filedialog.askopenfilename(
            parent=root,
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
    finally:
        root.attributes("-topmost", False)
        root.withdraw()

def get_sheet_names(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    return xls.sheet_names

def choose_sheet(sheet_names, file_path, default_sheet=None, parent=None):
    dialog = tk.Toplevel(parent)
    dialog.title("Select Sheet to Convert")
    _set_dialog_geometry(dialog, 820, 360, min_width=660, min_height=300)
    dialog.resizable(True, True)

    default_name = default_sheet or sheet_names[0]
    selected = tk.StringVar(master=dialog, value=default_name)
    cancelled = [False]

    tk.Label(
        dialog,
        text=(
            f"Excel (xlsx) file:\n{file_path}\n\n"
            f"Accept default or select Sheet name from the dropdown list, below\n\n"
            f"Then OK to continue..."
        ),
        padx=16,
        pady=16,
        justify="left",
        wraplength=660
    ).pack()

    dropdown = tk.OptionMenu(dialog, selected, *sheet_names)
    dropdown.pack(pady=(0, 8))

    def confirm():
        dialog.destroy()

    def cancel():
        cancelled[0] = True
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="OK", command=confirm).pack(side="left", padx=6)
    tk.Button(btn_frame, text="Cancel", command=cancel).pack(side="left", padx=6)

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    dialog.grab_set()
    dialog.wait_window()
    return None if cancelled[0] else selected.get()

def confirm_staff_type_dialog(staff_type, file_path, parent=None):
    """Show detected staff type and let user confirm before proceeding."""
    dialog = tk.Toplevel(parent)
    dialog.title("Confirm Staff Type")
    _set_dialog_geometry(dialog, 560, 280, min_width=500, min_height=240)
    dialog.resizable(True, True)

    staff_type_display = staff_type.upper() if staff_type else "UNKNOWN"
    cancelled = [False]

    tk.Label(
        dialog,
        text=(
            f"File: {Path(file_path).name}\n\n"
            f"Detected staff type: {staff_type_display}\n\n"
            "Is this correct?"
        ),
        padx=16,
        pady=20,
        justify="left",
        wraplength=460
    ).pack()

    def confirm():
        dialog.destroy()

    def cancel():
        cancelled[0] = True
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="Yes, Continue", command=confirm).pack(side="left", padx=6)
    tk.Button(btn_frame, text="No, Cancel", command=cancel).pack(side="left", padx=6)

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    dialog.grab_set()
    dialog.wait_window()
    return not cancelled[0]


def apply_staff_type_processing(df, staff_type):
    """Apply type-specific data transformations and validations."""
    if not staff_type:
        return df

    staff_type_lower = staff_type.lower()

    # Type-specific processing
    if staff_type_lower == "na":
        # NA-specific: ensure Date column exists (schedule format)
        if "Date" not in df.columns:
            show_focused_info_dialog(
                "Invalid NA file",
                "NA schedule files must contain a 'Date' column."
            )
    elif staff_type_lower == "oa":
        # OA-specific: validation/processing (to be defined)
        pass
    elif staff_type_lower == "sa":
        # SA-specific: validation/processing (to be defined)
        pass

    return df

def open_in_calc(file_path, sheet_name):
    launcher = shutil.which("libreoffice") or shutil.which("soffice")
    if not launcher:
        show_focused_info_dialog(
            "LibreOffice not found",
            "Could not find LibreOffice.\nInstall LibreOffice or add it to PATH."
        )
        return None

    return subprocess.Popen([
        launcher,
        "--calc",
        file_path
    ])


def close_calc_process(process):
    if process is None:
        return

    if process.poll() is not None:
        return

    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()


def focus_converted_file(file_path):
    if sys.platform == "darwin":
        target = Path(file_path)
        if target.is_dir():
            subprocess.Popen(["open", str(target)])
        else:
            subprocess.Popen(["open", "-R", str(target)])


def show_final_completion_dialog(output=None, upload_message=None, sql_file=None, parent=None):
    dialog = tk.Toplevel(parent)
    dialog.title("Conversions Complete")
    dialog.resizable(True, True)

    action = [""]

    def continue_to_database():
        action[0] = "ok"
        dialog.destroy()

    def cancel():
        action[0] = "cancel"
        dialog.destroy()

    button_frame = tk.Frame(dialog)
    button_frame.pack(side="bottom", fill="x", pady=(0, 12))
    tk.Button(button_frame, text="OK", command=continue_to_database).pack(side="left", padx=6)
    tk.Button(button_frame, text="Cancel", command=cancel).pack(side="left", padx=6)

    export_section = ""
    if output and upload_message:
        export_section = f"{upload_message}\n\n"
        export_section += f"Converted CSV file:\n{output}\n\n"
        if sql_file:
            export_section += f"Saved SQL file:\n{sql_file}\n\nand is ready for database update.\n\n"

    est_width, est_height = _estimate_message_dialog_size(export_section or "Conversions complete.", base_width=700, base_height=340)
    _set_dialog_geometry(dialog, est_width, est_height, min_width=620, min_height=320)

    if export_section:
        wrap_length = max(520, min(980, est_width - 40))
        tk.Label(
            dialog,
            text=export_section.rstrip(),
            padx=16,
            pady=20,
            justify="left",
            wraplength=wrap_length
        ).pack()

    # When completion is shown, immediately reveal/select the generated SQL file.
    if sql_file:
        focus_converted_file(sql_file)

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    dialog.grab_set()
    dialog.wait_window()
    return action[0]


def export_sheet(file_path, sheet_name, data_dir, upload_enabled, staff_type=None):
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=[0, 1],
        engine="openpyxl"
    )
    df.columns = _flatten_headers(df.columns)

    # Apply staff-type-specific processing
    df = apply_staff_type_processing(df, staff_type)

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%m/%d/%y")

    output = data_dir / build_output_filename(file_path)
    df.to_csv(output, index=False)
    store_last_csv_filename(output, data_dir)
    sql_file = None

    if upload_enabled:
        upload_url = get_config("NEW_OBS_SEM", "STAFF_UPLOAD_URL")
        if not upload_url:
            show_focused_info_dialog(
                "Invalid upload URL",
                "STAFF_UPLOAD_URL is missing in config.live.ini"
            )
            return False

        upload_type = infer_upload_type(xlsx_file_path=file_path, csv_file_path=output)
        try:
            verbose_response = upload_csv_to_staff_site(output, upload_type, verbose_enabled=True)
            verbose_file = data_dir / "last_upload_verbose.html"
            open_response_in_integrated_browser(verbose_response.text, verbose_file)

            final_response = upload_csv_to_staff_site(output, upload_type, verbose_enabled=False)
            final_results_file = data_dir / "last_upload_sql_results.html"
            open_response_in_integrated_browser(final_response.text, final_results_file)
            sql_file = save_insert_statements(final_response.text, output, data_dir)
            upload_message = (
                f"===== Processing staff type: {upload_type.upper()} =====\n"
                #f"Verbose HTTP {verbose_response.status_code}\nFinal HTTP {final_response.status_code}\n\n"
            )
        except Exception as error:
            upload_message = f"Upload failed: {error}"
    else:
        upload_message = "Upload skipped (Automatic upload is disabled)."

    return True, output, upload_message, sql_file

def main():
    root = tk.Tk()
    root.title("Step 11: Publish Staff Schedule")
    root.geometry("640x280")
    root.resizable(False, False)
    startup_action = tk.StringVar(value="")

    tk.Label(
        root,
        text="Generates two types of SQL command files:\n\n > From Excel Sheets provided for SA, OA, and NA\n> Continuing sequence rotation for SWOC staff\n\nThen updates the Night Staff Schedule database\n\nSelect the Excel file to process",
        padx=16,
        pady=16
    ).pack()

    def select_file():
        startup_action.set("select")

    def cancel_startup():
        startup_action.set("cancel")

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="OK", command=select_file).pack(side="left", padx=6)
    tk.Button(button_frame, text="Cancel", command=cancel_startup).pack(side="left", padx=6)

    root.protocol("WM_DELETE_WINDOW", cancel_startup)

    root.lift()
    root.focus_force()
    root.attributes("-topmost", True)
    root.update()

    root.wait_variable(startup_action)
    if startup_action.get() != "select":
        root.destroy()
        return

    upload_enabled = True
    root.destroy()

    root = tk.Tk()
    root.withdraw()
    file_path = pick_file(root)
    if not file_path:
        root.destroy()
        return

    data_dir = ensure_data_dir()
    copied_file_path = copy_source_to_data_dir(file_path, data_dir)

    # Detect and confirm staff type
    detected_staff_type = _normalize_staff_type(Path(copied_file_path).stem)
    if not confirm_staff_type_dialog(detected_staff_type or "unknown", copied_file_path, parent=root):
        root.destroy()
        return

    sheets = get_sheet_names(copied_file_path)

    # Open spreadsheet visually
    calc_process = open_in_calc(copied_file_path, sheets[0])

    # Let user choose
    chosen_sheet = choose_sheet(
        sheets,
        file_path=copied_file_path,
        default_sheet=sheets[0],
        parent=root,
    )
    if chosen_sheet is None:
        root.destroy()
        return

    # Export selected sheet
    result = export_sheet(copied_file_path, chosen_sheet, data_dir, upload_enabled, staff_type=detected_staff_type)
    if not result:
        root.destroy()
        return
    _, output, upload_message, sql_file = result

    completion_action = show_final_completion_dialog(
        output=output,
        upload_message=upload_message,
        sql_file=sql_file,
        parent=root,
    )
    close_calc_process(calc_process)

    if completion_action != "ok":
        root.destroy()
        return

    if not preflight_mysql_connection_check():
        root.destroy()
        return

    if not connect_to_remote_mysql_db():
        root.destroy()
        return

    if sql_file:
        focus_converted_file(sql_file)
    root.destroy()

if __name__ == "__main__":
    main()