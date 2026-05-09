# New Semester Observing Schedule
# Step 11: Staff Schedule Uploader
# This script converts an Excel file to CSV format and optionally uploads it to the staff schedule site.

# Usage with Python 3.13:
# % python xlsx2csv.py
# Follow dialog prompts to select the Excel file, choose the sheet, and review upload results.

import tkinter as tk
from tkinter import filedialog, messagebox   #, simpledialog
import pandas as pd
import subprocess
import shutil 
import sys
#import time
import re
import ssl
import html
import ast
import requests
from pathlib import Path
from urllib.parse import urlparse
from requests.adapters import HTTPAdapter
from urllib3.poolmanager import PoolManager

LAST_GENERATED_CSV_FILENAME = None


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

    namespace: dict = {}
    try:
        exec(compile(content, str(config_file), "exec"), {"__builtins__": {}}, namespace)
    except Exception:
        return {}

    return {k: v for k, v in namespace.items() if not k.startswith("_")}


LIVE_CONFIG = load_live_config()
STAFF_UPLOAD_URL = LIVE_CONFIG.get("NEW_OBS_SEM", {}).get("STAFF_UPLOAD_TOOL", None)


def is_valid_http_url(value):
    if not value or not isinstance(value, str):
        return False

    parsed = urlparse(value.strip())
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)

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


def copy_source_to_data_dir(file_path, data_dir):
    source = Path(file_path)
    destination = data_dir / source.name

    if source.resolve() != destination.resolve():
        shutil.copy2(source, destination)

    return destination


def store_last_csv_filename(output_path, data_dir):
    global LAST_GENERATED_CSV_FILENAME
    LAST_GENERATED_CSV_FILENAME = output_path.name
    tracker_file = data_dir / "last_generated_csv_filename.txt"
    tracker_file.write_text(f"{LAST_GENERATED_CSV_FILENAME}\n", encoding="utf-8")


def infer_upload_type(xlsx_file_path=None, csv_file_path=None):
    if xlsx_file_path:
        xlsx_staff = _normalize_staff_type(Path(xlsx_file_path).stem)
        if xlsx_staff:
            return xlsx_staff.lower()

    if csv_file_path:
        csv_staff = _normalize_staff_type(Path(csv_file_path).stem)
        if csv_staff:
            return csv_staff.lower()

    # Default to eeoc for generated swoc-like uploads without detectable staff type.
    return "eeoc"


def upload_csv_to_staff_site(csv_path, upload_type, verbose_enabled):
    session = requests.Session()
    session.mount("https://", LegacyTLSAdapter())

    form_data = {
        "type": upload_type,
        "submit": "Submit",
    }
    if verbose_enabled:
        form_data["verbose"] = "on"

    with open(csv_path, "rb") as file_handle:
        response = session.post(
            str(STAFF_UPLOAD_URL),
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
    return filedialog.askopenfilename(
        parent=root,
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

def get_sheet_names(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    return xls.sheet_names

def choose_sheet(sheet_names, file_path, default_sheet=None, parent=None):
    dialog = tk.Toplevel(parent)
    dialog.title("Select Sheet to Convert")
    dialog.geometry("700x320")
    dialog.resizable(False, False)

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

def open_in_calc(file_path, sheet_name):
    launcher = shutil.which("libreoffice") or shutil.which("soffice")
    if not launcher:
        messagebox.showerror(
            "LibreOffice not found",
            "Could not find LibreOffice.\nInstall LibreOffice or add it to PATH."
        )
        return

    subprocess.Popen([
        launcher,
        "--calc",
        file_path
    ])


def focus_converted_file(file_path):
    if sys.platform == "darwin":
        target = Path(file_path)
        if target.is_dir():
            subprocess.Popen(["open", str(target)])
        else:
            subprocess.Popen(["open", "-R", str(target)])


def show_quick_view_instructions(output_path):
    dialog = tk.Toplevel()
    dialog.title("Quick View Instructions")
    dialog.geometry("400x120")
    dialog.resizable(False, False)
    cancelled = [False]

    tk.Label(
        dialog,
        text="Press space bar for quick view and press again to close quick view",
        padx=16,
        pady=24,
        wraplength=350
    ).pack()

    def close_and_refocus():
        dialog.destroy()
        focus_converted_file(output_path)

    def cancel():
        cancelled[0] = True
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="OK", command=close_and_refocus).pack(side="left", padx=6)
    tk.Button(btn_frame, text="Cancel", command=cancel).pack(side="left", padx=6)

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    # dialog.grab_set()
    dialog.wait_window()
    return cancelled[0]


def show_final_completion_dialog(output=None, upload_message=None, sql_file=None):
    dialog = tk.Toplevel()
    dialog.title("Conversions Complete")
    dialog.geometry("640x380")
    dialog.resizable(False, False)

    button_frame = tk.Frame(dialog)
    button_frame.pack(side="bottom", fill="x", pady=(0, 12))
    tk.Button(button_frame, text="OK", command=dialog.destroy).pack()

    export_section = ""
    if output and upload_message:
        export_section = f"{upload_message}\n\n"
        export_section += f"Converted CSV file:\n{output}\n\n"
        if sql_file:
            export_section += f"Saved SQL file:\n{sql_file}\n\n"

    tk.Label(
        dialog,
        text=f"{export_section}Keep this window open while you:\n > Use the Space bar for quick view (toggle)\n > Edit to remove any overlapping dates in the SQL file and Save\n\nThen return here and click OK to proceed to run the database imports.",
        padx=16,
        pady=20,
        justify="left",
        wraplength=600
    ).pack()

    dialog.lift()
    dialog.focus_force()
    dialog.attributes("-topmost", True)
    dialog.grab_set()
    dialog.wait_window()


def export_sheet(file_path, sheet_name, data_dir, upload_enabled):
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=[0, 1],
        engine="openpyxl"
    )
    df.columns = _flatten_headers(df.columns)

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%m/%d/%y")

    output = data_dir / build_output_filename(file_path)
    df.to_csv(output, index=False)
    store_last_csv_filename(output, data_dir)
    sql_file = None

    if upload_enabled:
        if not is_valid_http_url(STAFF_UPLOAD_URL):
            messagebox.showerror(
                "Invalid upload URL",
                f"STAFF_UPLOAD_TOOL is missing or invalid in config.live.ini.\nCurrent value: {STAFF_UPLOAD_URL}"
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
            # if show_quick_view_instructions(sql_file)\
            #     return False
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

    sheets = get_sheet_names(copied_file_path)

    # Open spreadsheet visually
    open_in_calc(copied_file_path, sheets[0])

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
    result = export_sheet(copied_file_path, chosen_sheet, data_dir, upload_enabled)
    if not result:
        root.destroy()
        return
    _, output, upload_message, sql_file = result

    show_final_completion_dialog(output=output, upload_message=upload_message, sql_file=sql_file)
    if sql_file:
        focus_converted_file(sql_file)
    root.destroy()

if __name__ == "__main__":
    main()