"""Main application window for CAT-RG.

Thread-safe UI updates, progress bar, drag-and-drop support.
"""

from __future__ import annotations

import os
import queue
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageTk

from catrg.models.config import ConfigManager
from catrg.models.statements import StatementManager
from catrg.core.parser import (
    load_json,
    parse_cybertip,
    extract_ip_addresses,
    IpOccurrence,
    ParsedCyberTip,
)
from catrg.core.ip_lookup import IpLookupService
from catrg.core.report_generator import (
    generate_police_report,
    generate_multi_tip_report,
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
    _collect_json_paths,
    _parse_dnd_file_list,
)

log = get_logger(__name__)

# Optional drag-and-drop support
_DND_AVAILABLE = False
try:
    import tkinterdnd2
    _DND_AVAILABLE = True
except (ImportError, RuntimeError):
    log.info("tkinterdnd2 not available; drag-and-drop disabled")


VERSION = "2.2"


def apply_window_icon(root: tk.Misc) -> None:
    """Apply ``app_icon.ico`` to the title bar and taskbar (Windows + frozen-friendly).

    Uses an absolute path (required for ``iconbitmap`` in many Tk builds) and
    ``iconphoto`` as a fallback because taskbar behavior differs by Tk version.
    """
    base = get_base_path()
    candidates = [
        (base / "app_icon.ico").resolve(),
    ]
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append((exe_dir / "app_icon.ico").resolve())

    ico_path = next((p for p in candidates if p.is_file()), None)
    if ico_path is None:
        log.debug("app_icon.ico not found (tried: %s)", candidates)
        return

    path_str = os.path.normpath(os.path.abspath(str(ico_path)))

    try:
        root.iconbitmap(path_str)
    except tk.TclError as exc:
        log.debug("iconbitmap failed for %s: %s", path_str, exc)

    try:
        pil_img = Image.open(ico_path)
        if pil_img.mode != "RGBA":
            pil_img = pil_img.convert("RGBA")
        pil_img.thumbnail((256, 256), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(pil_img)
        root.iconphoto(True, photo)
        root._catrg_icon_photo = photo  # noqa: SLF001 — prevent GC
    except Exception as exc:
        log.debug("iconphoto failed: %s", exc)


class CyberTipAnalyzer:
    """Main application controller."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("CAT-RG")
        self.root.geometry("800x650")
        apply_window_icon(self.root)

        self.json_file_paths: List[str] = []
        self.config = ConfigManager()
        self.stmts = StatementManager()
        self.template = self._load_default_template()
        self.ip_service: Optional[IpLookupService] = None
        self._current_tips: List[ParsedCyberTip] = []
        self._merged_ip_data: Dict[str, List[IpOccurrence]] = {}

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
        file_menu.add_command(label="Open Files", command=self.open_files, accelerator="Ctrl+O")
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
        tools_menu.add_command(label="Customize Statements", command=lambda: show_customize_statements_dialog(self.root, self.stmts, template=self.template))
        menubar.add_cascade(label="Tools", menu=tools_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        help_menu.add_command(label="Help", command=self._show_help)
        help_menu.add_command(label="Update MaxMind Credentials", command=lambda: show_maxmind_dialog(self.root, self.config))
        help_menu.add_command(label="Update Investigator Info", command=lambda: show_investigator_dialog(self.root, self.config))
        help_menu.add_command(label="Update ARIN API Key", command=lambda: show_arin_dialog(self.root, self.config))
        help_menu.add_command(label="Check for Updates", command=self._check_for_updates)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.bind("<Control-o>", lambda e: self.open_files())
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
            "1. Add one or more .json files via 'Add Files', 'Add Folder',\n"
            "   or drag-and-drop files/folders onto the window.\n\n"
            "2. Click 'Choose Statements' to select which statements to include.\n\n"
            "3. Click 'Analyze CyberTips' to process and generate a combined report.\n\n"
            "4. Save with 'File > Save Report As' (Ctrl+S) or export as text/Excel.\n\n"
            "5. Use 'Tools > Compare Reports' to find common data across tips.\n\n"
            "6. Use 'Tools > Report Template' to configure section order/visibility.\n\n"
            "HOT KEYS\n"
            "- Ctrl+O: Open .json files\n"
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

        # Drop zone
        self._dnd_hint = (
            "Drag & drop .json files or folders here"
            if _DND_AVAILABLE else
            "Use the buttons to add .json files or folders"
        )
        self.drop_frame = tk.Frame(
            self.root, bd=2, relief="groove", bg="#f0f4f8",
            highlightbackground="#a0b4c8", highlightthickness=1,
        )
        self.drop_frame.pack(fill="x", padx=40, pady=8)

        self.drop_label = tk.Label(
            self.drop_frame, text=self._dnd_hint,
            font=("Arial", 10, "italic"), fg="#6688aa", bg="#f0f4f8", pady=8,
        )
        self.drop_label.pack()

        # File list with scrollbar
        files_outer = ttk.Frame(self.drop_frame)
        files_outer.pack(fill="x", padx=10, pady=(0, 5))

        list_frame = ttk.Frame(files_outer)
        list_frame.pack(side=tk.LEFT, fill="both", expand=True)
        files_sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self.files_listbox = tk.Listbox(
            list_frame, width=60, height=5,
            yscrollcommand=files_sb.set,
        )
        files_sb.config(command=self.files_listbox.yview)
        self.files_listbox.pack(side=tk.LEFT, fill="both", expand=True)
        files_sb.pack(side=tk.LEFT, fill=tk.Y)

        btn_frame = ttk.Frame(files_outer)
        btn_frame.pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="Add Files", command=self.open_files, width=10).pack(pady=2)
        tk.Button(btn_frame, text="Add Folder", command=self._add_folder, width=10).pack(pady=2)
        tk.Button(btn_frame, text="Remove", command=self._remove_file, width=10).pack(pady=2)
        tk.Button(btn_frame, text="Clear All", command=self._clear_files, width=10).pack(pady=2)

        self.file_count_label = tk.Label(self.drop_frame, text="Files loaded: 0", font=("Arial", 9))
        self.file_count_label.pack(pady=(0, 5))

        self.choose_btn = tk.Button(self.root, text="Choose Statements", command=self._choose_statements, state="disabled")
        self.choose_btn.pack(pady=5)

        tk.Button(self.root, text="Analyze CyberTips", command=self.analyze_report).pack(pady=5)

        # Progress area
        prog_frame = ttk.Frame(self.root)
        prog_frame.pack(fill="x", padx=20, pady=5)
        self.status_label = tk.Label(prog_frame, text="Ready", font=("Arial", 10), fg="blue")
        self.status_label.pack(side=tk.LEFT)
        self.progress_bar = ttk.Progressbar(prog_frame, mode="determinate", length=200)
        self.progress_bar.pack(side=tk.RIGHT, padx=10)

        self.output_text = scrolledtext.ScrolledText(self.root, width=90, height=25)
        self.output_text.pack(pady=10)

    # ── File list helpers ─────────────────────────────────────────

    def _update_file_count(self) -> None:
        self.file_count_label.config(text=f"Files loaded: {len(self.json_file_paths)}")
        if self.json_file_paths:
            self.choose_btn.config(state="normal")
        else:
            self.choose_btn.config(state="disabled")

    def _ingest_paths(self, raw_paths) -> int:
        added = 0
        for raw in raw_paths:
            for jp in _collect_json_paths(raw):
                if jp not in self.json_file_paths:
                    self.json_file_paths.append(jp)
                    self.files_listbox.insert(tk.END, os.path.basename(jp))
                    added += 1
        self._update_file_count()
        return added

    def _remove_file(self) -> None:
        sel = self.files_listbox.curselection()
        if sel:
            self.json_file_paths.pop(sel[0])
            self.files_listbox.delete(sel[0])
            self._update_file_count()

    def _clear_files(self) -> None:
        self.json_file_paths.clear()
        self.files_listbox.delete(0, tk.END)
        self._update_file_count()

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder to scan for .json files")
        if folder:
            added = self._ingest_paths([folder])
            if added == 0:
                messagebox.showinfo("Add Folder", "No .json files found in the selected folder.")

    # ── Drag-and-drop ─────────────────────────────────────────────

    def _setup_dnd(self) -> None:
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
        self.drop_label.config(bg="#d6e8f7", fg="#2266aa", text="Drop files or folders here...")

    def _on_drag_leave(self, event) -> None:
        self.drop_frame.config(bg="#f0f4f8", highlightbackground="#a0b4c8", highlightthickness=1)
        self.drop_label.config(bg="#f0f4f8", fg="#6688aa", text=self._dnd_hint)

    def _on_drop(self, event) -> None:
        self._on_drag_leave(event)
        paths = _parse_dnd_file_list(self.root, getattr(event, "data", "") or "")
        if not paths:
            return
        added = self._ingest_paths(paths)
        if added:
            self.drop_label.config(fg="#228833", text=f"{added} file(s) added")
            self._set_status(f"{added} file(s) added via drag-and-drop")
            self.root.after(2500, lambda: self.drop_label.config(fg="#6688aa", text=self._dnd_hint))

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

    def open_files(self) -> None:
        paths = filedialog.askopenfilenames(filetypes=[("JSON files", "*.json")])
        if paths:
            self._ingest_paths(paths)
            for p in paths:
                self.config.add_recent_file(p)
            self._update_recent_menu()

    def _update_recent_menu(self) -> None:
        self.recent_menu.delete(0, tk.END)
        if not self.config.recent_files:
            self.recent_menu.add_command(label="No recent files", state="disabled")
        else:
            for fp in self.config.recent_files:
                self.recent_menu.add_command(label=fp, command=lambda f=fp: self._load_recent(f))

    def _load_recent(self, path: str) -> None:
        if os.path.exists(path):
            self._ingest_paths([path])
        else:
            messagebox.showwarning("Warning", f"File not found: {path}")
            self.config.recent_files.remove(path)
            self.config.save_recent_files()
            self._update_recent_menu()

    def _choose_statements(self) -> None:
        if not self.json_file_paths:
            messagebox.showwarning("Warning", "Please add a JSON file first.")
            return

        preview_path = self.json_file_paths[0]

        def preview_fn():
            data = load_json(preview_path)
            if data is None:
                return None
            tip = parse_cybertip(data)
            return generate_police_report(
                tip, data, self.stmts,
                self.config.investigator.title,
                self.config.investigator.name,
                self.template,
            )

        show_choose_statements_dialog(self.root, self.stmts, preview_path, preview_fn)

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
        if not self._merged_ip_data:
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
            export_ip_data(filename, self._merged_ip_data, self.ip_service)
            messagebox.showinfo("Success", f"IP data exported as {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export IP data: {e}")

    def export_evidence_excel(self) -> None:
        if not self._current_tips:
            messagebox.showwarning("Warning", "No data available to export evidence.")
            return
        has_evidence = any(t.evidence_files for t in self._current_tips)
        if not has_evidence:
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
            export_evidence(filename, self._current_tips[0], tips=self._current_tips)
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
        if not self.json_file_paths:
            messagebox.showerror("Error", "Please add at least one JSON file.")
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
                tips_and_data: List[Tuple[ParsedCyberTip, dict]] = []
                merged_ips: Dict[str, List[IpOccurrence]] = {}

                for i, path in enumerate(self.json_file_paths):
                    self._set_status(f"Loading file {i + 1}/{len(self.json_file_paths)}...")
                    data = load_json(path)
                    if data is None:
                        q.put(("error", f"Failed to load: {os.path.basename(path)}"))
                        return

                    valid, reason = validate_cybertip_json(data)
                    if not valid:
                        q.put(("error", f"Invalid JSON ({os.path.basename(path)}): {reason}"))
                        return

                    tip = parse_cybertip(data)
                    tips_and_data.append((tip, data))

                    for ip, occs in tip.all_ip_data.items():
                        merged_ips.setdefault(ip, []).extend(occs)

                self._current_tips = [t for t, _ in tips_and_data]
                self._merged_ip_data = merged_ips

                max_ips = 50
                all_ips = list(merged_ips.keys())
                num_unique = len(all_ips)

                if num_unique > max_ips:
                    self.root.after(0, lambda: self._prompt_ip_limit(
                        num_unique, max_ips, tips_and_data, merged_ips, q,
                    ))
                    return

                self._run_analysis_phase2(tips_and_data, merged_ips, all_ips, q)
            except Exception as e:
                log.exception("Analysis failed")
                q.put(("error", str(e)))

        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        self._poll_result(thread, q)

    def _prompt_ip_limit(
        self, total: int, limit: int,
        tips_and_data: List[Tuple[ParsedCyberTip, dict]],
        merged_ips: Dict[str, List[IpOccurrence]],
        q: queue.Queue,
    ) -> None:
        process_all = messagebox.askyesno(
            "IP Limit Exceeded",
            f"{total} unique IPs found across all CyberTips (more than {limit}).\n"
            "Query all IPs for geolocation/WHOIS? (Yes: All -- may be slow; No: First 50 -- faster)\n"
            "All IPs will be listed in the report regardless.",
        )
        ips_to_query = list(merged_ips.keys()) if process_all else list(merged_ips.keys())[:limit]

        def resume():
            try:
                self._run_analysis_phase2(tips_and_data, merged_ips, ips_to_query, q)
            except Exception as e:
                log.exception("Analysis phase 2 failed")
                q.put(("error", str(e)))

        thread = threading.Thread(target=resume, daemon=True)
        thread.start()
        self._poll_result(thread, q)

    def _run_analysis_phase2(
        self,
        tips_and_data: List[Tuple[ParsedCyberTip, dict]],
        merged_ips: Dict[str, List[IpOccurrence]],
        ips_to_query: list,
        q: queue.Queue,
    ) -> None:
        self._set_status(f"Generating report ({len(ips_to_query)} IPs to query)...")

        if len(tips_and_data) == 1:
            tip, data = tips_and_data[0]
            police_report = generate_police_report(
                tip, data, self.stmts,
                self.config.investigator.title,
                self.config.investigator.name,
                self.template,
                self.ip_service,
            )
        else:
            police_report = generate_multi_tip_report(
                tips_and_data, self.stmts,
                self.config.investigator.title,
                self.config.investigator.name,
                self.template,
                self.ip_service,
            )

        ip_report = generate_ip_report(
            merged_ips,
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
        self._clear_files()
        self._current_tips = []
        self._merged_ip_data = {}
        if self.ip_service:
            self.ip_service.clear_cache()
        self.choose_btn.config(state="disabled")
        self.progress_bar["value"] = 0
        self._set_status("Ready")

    def new_analysis(self) -> None:
        self.clear_output()
        self.open_files()
        if self.json_file_paths:
            self.analyze_report()

    @staticmethod
    def _load_default_template() -> ReportTemplate:
        try:
            profiles = ReportTemplate.list_profiles()
            if "Default" in profiles:
                return ReportTemplate.load_profile("Default")
        except Exception:
            pass
        return ReportTemplate()

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
