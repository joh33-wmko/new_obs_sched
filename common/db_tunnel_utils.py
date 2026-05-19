"""Database and SSH tunnel utilities for remote database access.

Pure utilities for:
- SSH tunnel management (creation/cleanup)
- Database connectivity checks
- Database operations (INSERT execution, result tracking)
- SQL parsing and file operations

These are independent of Tkinter and suitable for use by any step.
"""

import html
import os
import re
import socket
import subprocess
import shutil
import time
from pathlib import Path


def _is_tcp_port_open(host, port, timeout=1.0):
    """Check if a TCP port is open and accepting connections."""
    try:
        with socket.create_connection((host, int(port)), timeout=timeout):
            return True
    except OSError:
        return False


def _resolve_ssh_password(settings, password_provider=None, password_env_var=None):
    """Resolve SSH password from a provider callback or environment variable.

    Args:
        settings: SSH settings dictionary passed to the provider.
        password_provider: Optional callable ``password_provider(settings) -> str``.
        password_env_var: Optional env var name for password fallback.

    Returns:
        str | None: Resolved password if available, else None.
    """
    if password_provider:
        password = password_provider(settings)
        if password:
            return password

    env_name = password_env_var or "SSH_PASSWORD"
    return os.getenv(env_name) or None


def _start_ssh_tunnel(
    settings,
    password_provider=None,
    *,
    get_password_func=None,
    password_env_var=None,
):
    """Start an SSH tunnel to access remote MySQL.
    
    Args:
        settings: Dict with SSH connection settings (host, user, port, etc.)
        password_provider: Callable that returns SSH password when needed.
                   If not provided, env var fallback and/or SSH key auth is used.
        get_password_func: Backward-compatible alias for password_provider.
        password_env_var: Optional env var name used as password fallback.
    
    Returns:
        subprocess.Popen: The tunnel process, or None if already listening.
    
    Raises:
        RuntimeError: If tunnel setup or startup fails.
    """
    if not settings.get("ssh_tunnel_enabled"):
        return None

    # Backward compatibility for prior call sites.
    if password_provider is None and get_password_func is not None:
        password_provider = get_password_func

    local_host = settings.get("host") or "127.0.0.1"
    local_port = settings.get("port") or 3306
    ssh_connect_timeout = int(settings.get("ssh_connect_timeout", 10))
    ssh_tunnel_wait_seconds = float(settings.get("ssh_tunnel_wait_seconds", 20.0))
    poll_interval = float(settings.get("ssh_tunnel_poll_interval", 0.25))
    password_prompt_limit = int(settings.get("ssh_password_prompt_limit", 1))

    # If something is already listening on the configured local endpoint,
    # assume an existing tunnel is active and reusable.
    if _is_tcp_port_open(local_host, local_port):
        return None

    ssh_target = f"{settings['ssh_user']}@{settings['ssh_host']}"
    bind_spec = (
        f"{local_host}:{local_port}:"
        f"{settings['mysql_remote_host']}:{settings['mysql_remote_port']}"
    )
    
    # Prefer sshpass + password when available; otherwise try key-based auth.
    sshpass_available = shutil.which("sshpass") is not None
    ssh_password = _resolve_ssh_password(
        settings,
        password_provider=password_provider,
        password_env_var=password_env_var,
    )
    ssh_key_file = settings.get("ssh_key_file")

    base_ssh_command = [
        "ssh",
        "-N",
        "-L",
        bind_spec,
        "-p",
        str(settings["ssh_port"]),
        "-o",
        "StrictHostKeyChecking=accept-new",
        "-o",
        "ExitOnForwardFailure=yes",
        "-o",
        f"ConnectTimeout={ssh_connect_timeout}",
        "-o",
        "UserKnownHostsFile=/dev/null",
        "-o",
        f"NumberOfPasswordPrompts={password_prompt_limit}",
    ]
    if ssh_key_file:
        base_ssh_command.extend(["-i", str(ssh_key_file)])

    if ssh_password:
        if not sshpass_available:
            raise RuntimeError(
                "SSH password was provided, but sshpass is not installed.\n\n"
                "Install sshpass:\n"
                "  brew install sshpass\n\n"
                "Or configure SSH key authentication and retry."
            )
        command = ["sshpass", "-p", ssh_password, *base_ssh_command, ssh_target]
    else:
        # Force non-interactive key auth path for faster failures.
        base_ssh_command.extend(["-o", "BatchMode=yes"])
        command = [*base_ssh_command, ssh_target]

    # Start tunnel process
    try:
        process = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as error:
        raise RuntimeError(
            f"Failed to start SSH tunnel: {error}"
        )

    # Wait for tunnel to become ready (configurable timeout)
    max_wait_iterations = max(1, int(ssh_tunnel_wait_seconds / poll_interval))
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
        if _is_tcp_port_open(local_host, local_port, timeout=poll_interval):
            return process
        
        time.sleep(poll_interval)

    # Timeout reached
    process.terminate()
    try:
        process.wait(timeout=2)
    except subprocess.TimeoutExpired:
        process.kill()
    
    raise RuntimeError(
        f"SSH tunnel did not open local port {local_port} within {max_wait_iterations * poll_interval:.0f}s.\n"
        "The SSH process may have encountered an authentication issue.\n"
        "Verify SSH password and connectivity, then retry."
    )


def _stop_ssh_tunnel(process):
    """Cleanly terminate an SSH tunnel subprocess.
    
    Args:
        process: subprocess.Popen object or None.
    """
    if process is None:
        return
    if process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=3)
        except subprocess.TimeoutExpired:
            process.kill()


def fetch_db_snapshot(cursor, main_table, type_columns, staff_types):
    """Fetch total rows and per-type counts for a table snapshot.
    
    Args:
        cursor: Database cursor.
        main_table: Name of the table to query.
        type_columns: List of column names that might contain type info.
        staff_types: List of staff type values to count (e.g., ['OA', 'NA', 'SA']).
    
    Returns:
        tuple: (total_rows, type_counts_dict)
    """
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


def apply_insert_statements_to_database(connection, sql_file):
    """Execute INSERT statements from SQL file and track results.
    
    Args:
        connection: Active database connection.
        sql_file: Path to SQL file containing INSERT statements.
    
    Returns:
        dict: Statistics including total, success, failed, already_exists counts,
              and detailed statement execution results.
    """
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
    """Write INSERT execution results to a detailed report file.
    
    Args:
        stats: Results dict from apply_insert_statements_to_database().
        sql_file: Path to the source SQL file.
        data_dir: Directory where results file will be written.
    
    Returns:
        Path: Path to the written results file.
    """
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


def extract_insert_statements(response_text):
    """Extract INSERT statements from HTML response text.
    
    Args:
        response_text: Response text potentially containing HTML markup.
    
    Returns:
        list: Cleaned INSERT statements.
    """
    normalized = response_text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
    plain_text = html.unescape(re.sub(r"<[^>]+>", "\n", normalized))
    statements = re.findall(r"(?is)\binsert\b.*?;", plain_text)
    cleaned = [re.sub(r"\s+", " ", stmt).strip() for stmt in statements]
    return cleaned


def save_insert_statements(response_text, csv_output_path, data_dir):
    """Parse INSERT statements from response and save to SQL file.
    
    Args:
        response_text: Response text from server (may contain HTML).
        csv_output_path: Base filename for the SQL output (stem used).
        data_dir: Directory where SQL file will be written.
    
    Returns:
        Path: Path to the written SQL file.
    """
    sql_output = data_dir / f"{Path(csv_output_path).stem}.sql"
    statements = extract_insert_statements(response_text)

    if statements:
        sql_output.write_text("\n".join(statements) + "\n", encoding="utf-8")
    else:
        sql_output.write_text("-- No INSERT statements found in upload response.\n", encoding="utf-8")

    return sql_output
