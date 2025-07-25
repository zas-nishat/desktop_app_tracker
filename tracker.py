import time
import json
import os
from datetime import datetime
import win32gui
import win32process
import psutil

SAVE_FILE = "usage_sessions.json"
HTML_FILE = "usage_report.html"
SAVE_INTERVAL = 10  # seconds - data and HTML update interval

def get_active_app_name():
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return process.name().lower().replace(".exe", "")
    except Exception:
        return "unknown"

def format_duration(seconds):
    seconds = int(seconds)
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return f"{h}h {m}m {s}s"

def save_sessions(sessions, filename=SAVE_FILE):
    with open(filename, "w") as f:
        json.dump(sessions, f, indent=4)

def load_sessions(filename=SAVE_FILE):
    if os.path.exists(filename):
        try:
            with open(filename, "r") as f:
                return json.load(f)
        except:
            return []
    return []

def generate_html(sessions):
    # Calculate total usage per app
    totals = {}
    for s in sessions:
        totals[s['app']] = totals.get(s['app'], 0) + s['duration_sec']

    # Sort apps alphabetically for overview
    sorted_totals = sorted(totals.items(), key=lambda x: x[0])

    # Build overview HTML table
    overview_html = """
    <h2>App Usage Overview (Total Time per App)</h2>
    <table>
      <thead><tr><th>App</th><th>Total Duration</th></tr></thead>
      <tbody>
    """
    for app, total_sec in sorted_totals:
        overview_html += f"<tr><td>{app}</td><td>{format_duration(total_sec)}</td></tr>"
    overview_html += "</tbody></table><br><br>"

    # Build detailed session HTML table
    detail_html = """
    <h2>Detailed Sessions</h2>
    <table>
      <thead><tr><th>App</th><th>Start Time</th><th>End Time</th><th>Duration</th></tr></thead>
      <tbody>
    """
    for session in sessions:
        detail_html += f"<tr><td>{session['app']}</td><td>{session['start']}</td><td>{session['end']}</td><td>{format_duration(session['duration_sec'])}</td></tr>"
    detail_html += "</tbody></table>"

    # Final full HTML with some basic styling
    full_html = f"""
    <html><head>
    <title>App Usage Report</title>
    <style>
      body {{ font-family: Arial, sans-serif; background:#f9f9f9; padding:20px; }}
      table {{ border-collapse: collapse; width: 100%; background:#fff; margin-bottom: 30px; }}
      th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
      th {{ background: #f2f2f2; }}
      h2 {{ color: #333; }}
    </style>
    </head><body>
    <h1>App Usage Report</h1>
    {overview_html}
    {detail_html}
    </body></html>
    """
    return full_html

def tracking_loop():
    sessions = load_sessions()
    current_app = get_active_app_name()
    start_time = datetime.now()

    last_save_time = time.time()

    while True:
        time.sleep(1)
        new_app = get_active_app_name()

        if new_app != current_app:
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            sessions.append({
                "app": current_app,
                "start": start_time.strftime("%Y-%m-%d %H:%M:%S"),
                "end": end_time.strftime("%Y-%m-%d %H:%M:%S"),
                "duration_sec": duration
            })
            current_app = new_app
            start_time = datetime.now()

        # Regularly save & update HTML
        if time.time() - last_save_time > SAVE_INTERVAL:
            save_sessions(sessions)
            html_content = generate_html(sessions)
            with open(HTML_FILE, "w", encoding="utf-8") as f:
                f.write(html_content)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Data saved & HTML report updated.")
            last_save_time = time.time()

if __name__ == "__main__":
    print("Starting app usage tracker... (Press Ctrl+C to stop)")
    try:
        tracking_loop()
    except KeyboardInterrupt:
        print("\nTracking stopped by user. Goodbye!")
