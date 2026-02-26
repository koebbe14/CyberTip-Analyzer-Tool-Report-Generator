"""Main application window for CAT-RG.

Thread-safe UI updates, progress bar, drag-and-drop support.
"""

from __future__ import annotations

import os
import queue
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk
from typing import Optional

from PIL import Image, ImageTk

from catrg.models.config import ConfigManager
from catrg.models.statements import StatementManager
from catrg.core.parser import (
    load_json,
    parse_cybertip,
    extract_ip_addresses,
    ParsedCyberTip,
)
from catrg.core.ip_lookup import IpLookupService
from catrg.core.report_generator import (
    generate_police_report,
    generate_ip_report,
    ReportTemplate,
)
from catrg.core.docx_formatter import save_docx
from catrg.core.excel_exporter import export_ip_data, export_evidence
from catrg.utils.validators import validate_cybertip_json
from catrg.utils.date_utils import get_base_path
from catrg.utils.logger import get_logger
from catrg.gui.dialogs import (
    show_investigator_dialog,
    show_maxmind_dialog,
    show_arin_dialog,
    show_customize_statements_dialog,
    show_choose_statements_dialog,
    show_template_dialog,
    show_comparison_dialog,
)

log = get_logger(__name__)

# Optional drag-and-drop support
_DND_AVAILABLE = False
try:
    import tkinterdnd2
    _DND_AVAILABLE = True
except (ImportError, RuntimeError):
    log.info("tkinterdnd2 not available; drag-and-drop disabled")


VERSION = "2.1"


class CyberTipAnalyzer:
    """Main application controller."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("CAT-RG")
        self.root.geometry("800x650")

        self.json_file_path: Optional[str] = None
        self.config = ConfigManager()
        self.stmts = StatementManager()
        self.template = ReportTemplate()
        self.ip_service: Optional[IpLookupService] = None
        self._current_tip: Optional[ParsedCyberTip] = None

        # Load persisted state
        self.config.load_recent_files()
        self.config.load_investigator()
        self.stmts.load()

        # On first launch, walk the user through all setup dialogs with blank fields
        first_launch = not self.config.investigator.name or not self.config.investigator.title

        if first_launch:
            show_investigator_dialog(self.root, self.config)
            d = show_maxmind_dialog(self.root, self.config)
            if d:
                self.root.wait_window(d)
            d = show_arin_dialog(self.root, self.config)
            if d:
                self.root.wait_window(d)
        else:
            self.config.load_maxmind()
            self.config.load_arin()
            if not self.config.maxmind.is_configured:
                d = show_maxmind_dialog(self.root, self.config)
                if d:
                    self.root.wait_window(d)
            if not self.config.arin.is_configured:
                d = show_arin_dialog(self.root, self.config)
                if d:
                    self.root.wait_window(d)

        # Logo
        self.logo_image = None
        logo_path = get_base_path() / "logo.jpg"
        try:
            img = Image.open(logo_path)
            img = img.resize((200, 100), Image.Resampling.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(img)
        except Exception as e:
            log.warning("Could not load logo: %s", e)

        self._create_menu()
        self._create_widgets()
        self._setup_dnd()

    # ── Menu ──────────────────────────────────────────────────────

    def _create_menu(self) -> None:
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open File", command=self.open_file, accelerator="Ctrl+O")
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Recent Files", menu=self.recent_menu)
        file_menu.add_command(label="Save Report As", command=self.save_report_as, accelerator="Ctrl+S")
        file_menu.add_command(label="Export as Text", command=self.export_as_text)
        file_menu.add_command(label="Export IP Data to Excel", command=self.export_ip_excel)
        file_menu.add_command(label="Export Evidence to Excel", command=self.export_evidence_excel)
        file_menu.add_command(label="Clear Output", command=self.clear_output)
        file_menu.add_command(label="New Analysis", command=self.new_analysis, accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_app)
        menubar.add_cascade(label="File", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Compare Reports", command=lambda: show_comparison_dialog(self.root))
        tools_menu.add_command(label="Report Template", command=lambda: show_template_dialog(self.root, self.template))
        tools_menu.add_command(label="Customize Statements", command=lambda: show_customize_statements_dialog(self.root, self.stmts))
        menubar.add_cascade(label="Tools", menu=tools_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        help_menu.add_command(label="Help", command=self._show_help)
        help_menu.add_command(label="Update MaxMind Credentials", command=lambda: show_maxmind_dialog(self.root, self.config))
        help_menu.add_command(label="Update Investigator Info", command=lambda: show_investigator_dialog(self.root, self.config))
        help_menu.add_command(label="Update ARIN API Key", command=lambda: show_arin_dialog(self.root, self.config))
        help_menu.add_command(label="Check for Updates", command=self._check_for_updates)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_report_as())
        self.root.bind("<Control-n>", lambda e: self.new_analysis())

        self._update_recent_menu()

    def _show_about(self) -> None:
        messagebox.showinfo(
            "About",
            f"CyberTip Analysis Tool & Report Generator\nVersion {VERSION}\n"
            "Developed by Patrick Koebbe\n\n"
            "This tool analyzes CyberTipline JSON reports and generates formatted DOCX reports.",
        )

    def _show_help(self) -> None:
        messagebox.showinfo(
            "Help",
            "To use the CyberTip Analysis Tool & Report Generator:\n\n"
            "1. Select a JSON file via 'File > Open File' or the 'Browse' button.\n"
            "   You can also drag-and-drop a JSON file onto the window.\n\n"
            "2. Click 'Choose Statements' to select which statements to include.\n\n"
            "3. Click 'Analyze CyberTip' to process and generate a report.\n\n"
            "4. Save with 'File > Save Report As' (Ctrl+S) or export as text/Excel.\n\n"
            "5. Use 'Tools > Compare Reports' to find common data across tips.\n\n"
            "6. Use 'Tools > Report Template' to configure section order/visibility.\n\n"
            "HOT KEYS\n"
            "- Ctrl+O: Open new .json file\n"
            "- Ctrl+S: Save report as\n"
            "- Ctrl+N: Start new analysis\n\n"
            "Supported ESPs:\n"
            "Discord, Dropbox, Facebook, Google, Imgur, Instagram, Kik, MeetMe,\n"
            "Microsoft, Reddit, Roblox, Snapchat, Sony, Synchronoss, TikTok,\n"
            "WhatsApp, X (Twitter), Yahoo\n\n"
            "For support, contact Patrick Koebbe - Patrick.Koebbe@gmail.com",
        )

    # ── Widgets ───────────────────────────────────────────────────

    def _create_widgets(self) -> None:
        tk.Label(self.root, text="CAT-RG", font=("Arial", 18, "bold")).pack(pady=5)
        tk.Label(self.root, text="CyberTip Analysis Tool & Report Generator", font=("Arial", 11, "bold", "underline")).pack(pady=5)

        if self.logo_image:
            tk.Label(self.root, image=self.logo_image).pack(pady=10)

        # Drop zone / file selection area
        self.drop_frame = tk.Frame(
            self.root, bd=2, relief="groove", bg="#f0f4f8",
            highlightbackground="#a0b4c8", highlightthickness=1,
        )
        self.drop_frame.pack(fill="x", padx=40, pady=8)

        dnd_hint = "Drag & drop a .json file here" if _DND_AVAILABLE else "Select a .json file below"
        self.drop_label = tk.Label(
            self.drop_frame, text=dnd_hint,
            font=("Arial", 10, "italic"), fg="#6688aa", bg="#f0f4f8", pady=10,
        )
        self.drop_label.pack()

        self.file_entry = tk.Entry(self.drop_frame, width=50, justify="center")
        self.file_entry.pack(pady=(0, 5))

        tk.Button(self.drop_frame, text="Browse to .json", command=self.open_file).pack(pady=(0, 8))

        self.choose_btn = tk.Button(self.root, text="Choose Statements", command=self._choose_statements, state="disabled")
        self.choose_btn.pack(pady=5)

        tk.Button(self.root, text="Analyze CyberTip", command=self.analyze_report).pack(pady=5)

        # Progress area
        prog_frame = ttk.Frame(self.root)
        prog_frame.pack(fill="x", padx=20, pady=5)
        self.status_label = tk.Label(prog_frame, text="Ready", font=("Arial", 10), fg="blue")
        self.status_label.pack(side=tk.LEFT)
        self.progress_bar = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self.progress_bar.pack(side=tk.RIGHT, padx=10)

        self.output_text = scrolledtext.ScrolledText(self.root, width=90, height=25)
        self.output_text.pack(pady=10)

    def _setup_dnd(self) -> None:
        """Enable drag-and-drop if tkinterdnd2 is available."""
        if not _DND_AVAILABLE:
            return
        try:
            self.root.drop_target_register(tkinterdnd2.DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
            self.root.dnd_bind("<<DragEnter>>", self._on_drag_enter)
            self.root.dnd_bind("<<DragLeave>>", self._on_drag_leave)
            log.info("Drag-and-drop enabled")
        except Exception as e:
            log.warning("Could not enable drag-and-drop: %s", e)

    def _on_drag_enter(self, event) -> None:
        self.drop_frame.config(bg="#d6e8f7", highlightbackground="#3388cc", highlightthickness=2)
        self.drop_label.config(bg="#d6e8f7", fg="#2266aa", text="Drop file here...")

    def _on_drag_leave(self, event) -> None:
        self.drop_frame.config(bg="#f0f4f8", highlightbackground="#a0b4c8", highlightthickness=1)
        self.drop_label.config(bg="#f0f4f8", fg="#6688aa", text="Drag & drop a .json file here")

    def _on_drop(self, event) -> None:
        self._on_drag_leave(event)
        path = event.data.strip("{}")
        if path.lower().endswith(".json"):
            self.json_file_path = path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)
            self.config.add_recent_file(path)
            self._update_recent_menu()
            self.choose_btn.config(state="normal")
            self.drop_label.config(fg="#228833", text="File loaded successfully")
            self._set_status("File loaded via drag-and-drop")
        else:
            messagebox.showwarning("Warning", "Please drop a .json file.")

    # ── Thread-safe status updates ────────────────────────────────

    def _set_status(self, text: str) -> None:
        """Schedule a status-label update on the main thread."""
        self.root.after(0, lambda: self.status_label.config(text=text))

    def _set_progress(self, completed: int, total: int) -> None:
        """Schedule a progress-bar update on the main thread."""
        def _update():
            if total > 0:
                self.progress_bar["maximum"] = total
                self.progress_bar["value"] = completed
            self.status_label.config(text=f"Querying IPs: {completed}/{total}")
        self.root.after(0, _update)

    # ── File operations ───────────────────────────────────────────

    def open_file(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if path:
            self.json_file_path = path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)
            self.config.add_recent_file(path)
            self._update_recent_menu()
            self.choose_btn.config(state="normal")

    def _update_recent_menu(self) -> None:
        self.recent_menu.delete(0, tk.END)
        if not self.config.recent_files:
            self.recent_menu.add_command(label="No recent files", state="disabled")
        else:
            for fp in self.config.recent_files:
                self.recent_menu.add_command(label=fp, command=lambda f=fp: self._load_recent(f))

    def _load_recent(self, path: str) -> None:
        if os.path.exists(path):
            self.json_file_path = path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)
            self.choose_btn.config(state="normal")
        else:
            messagebox.showwarning("Warning", f"File not found: {path}")
            self.config.recent_files.remove(path)
            self.config.save_recent_files()
            self._update_recent_menu()

    def _choose_statements(self) -> None:
        if not self.json_file_path:
            messagebox.showwarning("Warning", "Please select a JSON file first.")
            return

        def preview_fn():
            data = load_json(self.json_file_path)
            if data is None:
                return None
            tip = parse_cybertip(data)
            return generate_police_report(
                tip, data, self.stmts,
                self.config.investigator.title,
                self.config.investigator.name,
                self.template,
            )

        show_choose_statements_dialog(self.root, self.stmts, self.json_file_path, preview_fn)

    # ── Export ────────────────────────────────────────────────────

    def save_report_as(self) -> None:
        current = self.output_text.get(1.0, tk.END).strip()
        if not current:
            messagebox.showwarning("Warning", "No report available to save.")
            return

        if "IP ADDRESS ANALYSIS:" in current:
            police_report, ip_report = current.split("IP ADDRESS ANALYSIS:", 1)
            police_report += "IP ADDRESS ANALYSIS:"
            ip_report = "IP ADDRESS ANALYSIS:" + ip_report
        else:
            police_report = current
            ip_report = ""

        try:
            filename = self._save_docx(police_report + ip_report)
            messagebox.showinfo("Success", f"Report saved as {filename}")
        except ValueError as e:
            messagebox.showwarning("Warning", str(e))

    def export_as_text(self) -> None:
        current = self.output_text.get(1.0, tk.END).strip()
        if not current:
            messagebox.showwarning("Warning", "No report available to export.")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            initialfile=f"cybertip_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        )
        if filename:
            with open(filename, "w", encoding="utf-8") as f:
                f.write(current)
            messagebox.showinfo("Success", f"Report exported as {filename}")

    def export_ip_excel(self) -> None:
        if not self._current_tip or not self._current_tip.all_ip_data:
            messagebox.showwarning("Warning", "No IP data available to export.")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"cybertip_ip_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )
        if not filename:
            return
        try:
            export_ip_data(filename, self._current_tip.all_ip_data, self.ip_service)
            messagebox.showinfo("Success", f"IP data exported as {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export IP data: {e}")

    def export_evidence_excel(self) -> None:
        if not self._current_tip:
            messagebox.showwarning("Warning", "No JSON file loaded to export evidence data.")
            return
        if not self._current_tip.evidence_files:
            messagebox.showwarning("Warning", "No evidence data available to export.")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"cybertip_evidence_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )
        if not filename:
            return
        try:
            export_evidence(filename, self._current_tip)
            messagebox.showinfo("Success", f"Evidence summary exported as {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export evidence summary: {e}")

    def _save_docx(self, full_report: str) -> str:
        default_name = f"cybertip_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        filename = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialfile=default_name,
            title="Save CyberTipline Report As",
        )
        if not filename:
            raise ValueError("Save operation cancelled by user")

        ip_intro = self.stmts.get_text("ip_intro") if "ip_intro" in self.stmts.selected else ""
        meta = self.stmts.get_text("meta") if "meta" in self.stmts.selected else ""

        return save_docx(full_report, filename, self.stmts, ip_intro, meta, template=self.template)

    # ── Analysis ──────────────────────────────────────────────────

    def analyze_report(self) -> None:
        if not self.json_file_path:
            messagebox.showerror("Error", "Please select a JSON file")
            return

        self._set_status("Starting analysis...")
        self.progress_bar["value"] = 0

        self.ip_service = IpLookupService(
            maxmind_id=self.config.maxmind.account_id,
            maxmind_key=self.config.maxmind.license_key,
            arin_key=self.config.arin.api_key,
        )

        q: queue.Queue = queue.Queue()

        def run():
            try:
                data = load_json(self.json_file_path)
                if data is None:
                    q.put(("error", "Failed to load JSON"))
                    return

                valid, reason = validate_cybertip_json(data)
                if not valid:
                    q.put(("error", f"Invalid CyberTip JSON: {reason}"))
                    return

                self._set_status("Parsing CyberTip data...")
                tip = parse_cybertip(data)
                self._current_tip = tip

                # Determine which IPs to query
                max_ips = 50
                all_ips = list(tip.all_ip_data.keys())
                num_unique = len(all_ips)

                if num_unique > max_ips:
                    self.root.after(0, lambda: self._prompt_ip_limit(num_unique, max_ips, data, tip, q))
                    return

                self._run_analysis_phase2(data, tip, all_ips, q)
            except Exception as e:
                log.exception("Analysis failed")
                q.put(("error", str(e)))

        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        self._poll_result(thread, q)

    def _prompt_ip_limit(self, total: int, limit: int, data: dict, tip: ParsedCyberTip, q: queue.Queue) -> None:
        process_all = messagebox.askyesno(
            "IP Limit Exceeded",
            f"{total} unique IPs found (more than {limit}).\n"
            "Query all IPs for geolocation/WHOIS? (Yes: All -- may be slow; No: First 50 -- faster)\n"
            "All IPs will be listed in the report regardless.",
        )
        ips_to_query = list(tip.all_ip_data.keys()) if process_all else list(tip.all_ip_data.keys())[:limit]

        def resume():
            try:
                self._run_analysis_phase2(data, tip, ips_to_query, q)
            except Exception as e:
                log.exception("Analysis phase 2 failed")
                q.put(("error", str(e)))

        thread = threading.Thread(target=resume, daemon=True)
        thread.start()
        self._poll_result(thread, q)

    def _run_analysis_phase2(self, data: dict, tip: ParsedCyberTip, ips_to_query: list, q: queue.Queue) -> None:
        self._set_status(f"Generating report ({len(ips_to_query)} IPs to query)...")

        police_report = generate_police_report(
            tip, data, self.stmts,
            self.config.investigator.title,
            self.config.investigator.name,
            self.template,
            self.ip_service,
        )

        ip_report = generate_ip_report(
            tip.all_ip_data,
            self.ip_service,
            queried_ips=set(ips_to_query),
            progress_callback=self._set_progress,
        )

        q.put(("success", police_report, ip_report))

    def _poll_result(self, thread: threading.Thread, q: queue.Queue) -> None:
        if thread.is_alive() and q.empty():
            self.root.after(100, lambda: self._poll_result(thread, q))
            return

        try:
            result = q.get_nowait()
        except queue.Empty:
            self.root.after(100, lambda: self._poll_result(thread, q))
            return

        if result[0] == "error":
            messagebox.showerror("Error", result[1])
            self._set_status("Analysis failed")
            return

        police_report, ip_report = result[1], result[2]
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, police_report + ip_report)

        try:
            filename = self._save_docx(police_report + ip_report)
            messagebox.showinfo("Success", f"Report saved as {filename}")
            self._set_status("Analysis complete")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {e}")
            self._set_status("Save failed")

    # ── Utility ───────────────────────────────────────────────────

    def clear_output(self) -> None:
        self.output_text.delete(1.0, tk.END)
        self.file_entry.delete(0, tk.END)
        self.json_file_path = None
        self._current_tip = None
        if self.ip_service:
            self.ip_service.clear_cache()
        self.choose_btn.config(state="disabled")
        self.stmts.selected = set(self.stmts.statements.keys())
        self.progress_bar["value"] = 0
        self._set_status("Ready")

    def new_analysis(self) -> None:
        self.clear_output()
        self.open_file()
        if self.json_file_path:
            self.analyze_report()

    def exit_app(self) -> None:
        self.config.save_recent_files()
        self.root.destroy()

    def _check_for_updates(self) -> None:
        import requests
        repo_api = "https://api.github.com/repos/koebbe14/CyberTip-Analyzer-Tool-Report-Generator/releases/latest"
        repo_dl = "https://github.com/koebbe14/CyberTip-Analyzer-Tool-Report-Generator/releases/latest"

        def ver_tuple(v):
            return tuple(int(x) for x in v.split("."))

        try:
            resp = requests.get(repo_api, timeout=10)
            if resp.status_code == 200:
                latest = resp.json().get("tag_name", "0.0").lstrip("v")
                if ver_tuple(latest) > ver_tuple(VERSION):
                    if messagebox.askyesno("Update Available", f"A new version {latest} is available.\n\nOpen download page?"):
                        webbrowser.open(repo_dl)
                else:
                    messagebox.showinfo("Up to Date", "You are running the latest version.")
            else:
                messagebox.showerror("Error", f"Failed to check for updates: HTTP {resp.status_code}")
        except Exception as e:
            messagebox.showerror("Error", f"Error checking for updates: {e}")
