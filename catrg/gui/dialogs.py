"""All dialog windows for CAT-RG.

Credentials, investigator info, statement customisation,
statement chooser, report-template editor, and report comparison.
"""

from __future__ import annotations

import os
import re
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import TYPE_CHECKING, Dict, List, Optional, Set

from catrg.models.config import ConfigManager
from catrg.models.statements import (
    StatementManager,
    DEFAULT_STATEMENTS,
    DEFAULT_FORMATTING,
    PLACEMENT_PREFIXES,
    PREFIX_TO_PLACEMENT,
)
from catrg.core.parser import load_json, extract_comparison_data
from catrg.core.report_generator import (
    ReportTemplate,
    DEFAULT_SECTIONS,
    DEFAULT_SECTION_NAMES,
)
from catrg.utils.logger import get_logger

if TYPE_CHECKING:
    from catrg.gui.main_window import CyberTipAnalyzer

log = get_logger(__name__)

ESP_VALUES = [
    "Custom", "Discord", "Dropbox", "Facebook", "Google", "Imgur",
    "Instagram", "Kik", "MeetMe", "Microsoft", "Reddit", "Roblox",
    "Snapchat", "Sony", "Synchronoss", "TikTok", "WhatsApp",
    "X (Twitter)", "Yahoo",
]

_ESP_LIST = [
    "Discord", "Dropbox", "Facebook", "Google", "Imgur",
    "Instagram", "Kik", "MeetMe", "Microsoft", "Reddit", "Roblox",
    "Snapchat", "Sony", "Synchronoss", "TikTok", "WhatsApp",
    "X (Twitter)", "Yahoo",
]


# ── Investigator Info ─────────────────────────────────────────────

def show_investigator_dialog(root: tk.Tk, config: ConfigManager) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Investigator Information")
    dialog.geometry("300x200")
    dialog.transient(root)
    dialog.grab_set()

    tk.Label(dialog, text="Enter Your Name:").pack(pady=5)
    name_entry = tk.Entry(dialog, width=30)
    name_entry.pack(pady=5)
    name_entry.insert(0, config.investigator.name)

    tk.Label(dialog, text="Enter Your Title:").pack(pady=5)
    title_entry = tk.Entry(dialog, width=30)
    title_entry.pack(pady=5)
    title_entry.insert(0, config.investigator.title)

    def save():
        n = name_entry.get().strip()
        t = title_entry.get().strip()
        if not n or not t:
            messagebox.showwarning("Warning", "Please enter both name and title.")
            return
        config.investigator.name = n
        config.investigator.title = t
        config.save_investigator()
        dialog.destroy()

    tk.Button(dialog, text="Save", command=save).pack(pady=10)
    dialog.protocol("WM_DELETE_WINDOW", lambda: root.quit())
    root.wait_window(dialog)


# ── MaxMind Credentials ──────────────────────────────────────────

def show_maxmind_dialog(root: tk.Tk, config: ConfigManager) -> Optional[tk.Toplevel]:
    dialog = tk.Toplevel(root)
    dialog.title("MaxMind Credentials")
    dialog.geometry("400x350")
    dialog.transient(root)
    dialog.grab_set()

    tk.Label(dialog, text="For IP Geo Lookup, enter your MaxMind credentials below (optional).\n\nCreate new account at ").pack(pady=5)
    link = tk.Label(dialog, text="this link", fg="blue", cursor="hand2")
    link.pack(pady=5)
    link.bind("<Button-1>", lambda e: webbrowser.open_new(
        "https://www.maxmind.com/en/geolite2/signup?utm_source=kb&utm_medium=kb-link&utm_campaign=kb-create-account"
    ))

    tk.Label(dialog, text="MaxMind Account ID:").pack(pady=5)
    id_entry = tk.Entry(dialog, width=30)
    id_entry.pack(pady=5)
    id_entry.insert(0, config.maxmind.account_id)

    tk.Label(dialog, text="MaxMind License Key:").pack(pady=5)
    key_entry = tk.Entry(dialog, width=30, show="*")
    key_entry.pack(pady=5)
    key_entry.insert(0, config.maxmind.license_key)

    def save_and_close():
        config.maxmind.account_id = id_entry.get().strip()
        config.maxmind.license_key = key_entry.get().strip()
        config.save_maxmind()
        dialog.destroy()
        if config.maxmind.is_configured:
            messagebox.showinfo("Success", "MaxMind credentials updated successfully.")
        else:
            messagebox.showinfo("Info", "MaxMind credentials skipped -- geolocation will not be available.")

    tk.Button(dialog, text="Save", command=save_and_close).pack(pady=10)
    tk.Button(dialog, text="Skip (No Geolocation)", command=dialog.destroy).pack(pady=5)
    return dialog


# ── ARIN Credentials ─────────────────────────────────────────────

def show_arin_dialog(root: tk.Tk, config: ConfigManager) -> Optional[tk.Toplevel]:
    dialog = tk.Toplevel(root)
    dialog.title("ARIN API Key")
    dialog.geometry("400x300")
    dialog.transient(root)
    dialog.grab_set()

    tk.Label(
        dialog,
        text="Enter your ARIN API key (optional).\n\n"
             "Adding an ARIN API key increases daily query limits.\n\n"
             "Daily query limits WITHOUT ARIN API key - 15 per min and 256 per day.\n"
             "Daily query limits WITH ARIN API Key - 60 per min and 1,000 per day.\n\n"
             "Get an API key at:",
    ).pack(pady=5)
    link = tk.Label(dialog, text="this link", fg="blue", cursor="hand2")
    link.pack(pady=5)
    link.bind("<Button-1>", lambda e: webbrowser.open_new("https://account.arin.net/public/account-setup"))

    tk.Label(dialog, text="ARIN API Key:").pack(pady=5)
    key_entry = tk.Entry(dialog, width=50)
    key_entry.pack(pady=5)
    key_entry.insert(0, config.arin.api_key)

    def save_and_close():
        config.arin.api_key = key_entry.get().strip()
        if config.arin.is_configured:
            config.save_arin()
            messagebox.showinfo("Success", "ARIN API key saved.")
        dialog.destroy()

    tk.Button(dialog, text="Save", command=save_and_close).pack(pady=10)
    tk.Button(dialog, text="Skip", command=dialog.destroy).pack(pady=5)
    return dialog


# ── Customise Statements ─────────────────────────────────────────

def show_customize_statements_dialog(root: tk.Tk, stmts: StatementManager) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Customize Report Statements")
    dialog.geometry("1000x800")
    dialog.transient(root)
    dialog.grab_set()

    notebook = ttk.Notebook(dialog)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    for key, default_text in DEFAULT_STATEMENTS.items():
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=key.capitalize())
        tk.Label(frame, text=f"Edit {key} statement:").pack(pady=5)
        if key == "intro":
            tk.Label(frame, text="Tip: Use [CURRENT_DATE], [INVESTIGATOR_NAME], [INVESTIGATOR_TITLE], "
                     "[CYBERTIP_NUMBER], [REPORT_DATE_RECEIVED] placeholders.",
                     fg="gray", wraplength=700).pack(pady=2)
        tw = tk.Text(frame, width=80, height=20, wrap="word")
        tw.pack(pady=5)
        current_val = stmts.statements.get(key, default_text)
        display_text = current_val if isinstance(current_val, str) else current_val.get("text", str(default_text))
        tw.insert(tk.END, display_text)

        def save_tab(k=key, widget=tw):
            stmts.statements[k] = widget.get(1.0, tk.END).strip()

        def reset_tab(k=key, widget=tw, dt=default_text):
            widget.delete(1.0, tk.END)
            widget.insert(tk.END, dt)

        btn_row = ttk.Frame(frame)
        btn_row.pack(pady=5)
        tk.Button(btn_row, text="Save This Statement", command=save_tab).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_row, text="Reset to Default", command=reset_tab).pack(side=tk.LEFT, padx=5)

    _build_add_new_tab(notebook, stmts)
    _build_manage_tab(notebook, stmts)

    bottom = ttk.Frame(dialog)
    bottom.pack(fill="x", padx=10, pady=5)

    def close_and_save():
        stmts.save()
        dialog.destroy()

    ttk.Button(bottom, text="Save All & Close", command=close_and_save).pack(side=tk.RIGHT, padx=5)
    ttk.Button(bottom, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)


# ── Statement Templates ──────────────────────────────────────────

STATEMENT_TEMPLATES: Dict[str, Dict[str, str]] = {
    "Legal Disclaimer": {
        "name": "legal_disclaimer",
        "text": "This report is for official use only and may contain sensitive "
                "information protected by law. Unauthorized disclosure, copying, or "
                "distribution of this report or its contents is strictly prohibited.",
        "desc": "Standard law-enforcement confidentiality notice.",
    },
    "Investigator Note": {
        "name": "investigator_note",
        "text": "Investigator's additional observations:\n\n[Enter details here]",
        "desc": "Placeholder for free-form investigator observations.",
    },
    "Chain of Custody Notice": {
        "name": "chain_of_custody",
        "text": "This report and all associated digital evidence have been handled in "
                "accordance with department chain-of-custody procedures. All files were "
                "downloaded directly from the NCMEC CyberTipline portal and stored on "
                "a forensic workstation.",
        "desc": "Documents how evidence was received and stored.",
    },
    "Evidence Handling Note": {
        "name": "evidence_handling",
        "text": "Digital evidence referenced in this report has been preserved in its "
                "original form. Hash values (MD5) listed for each file can be used to "
                "verify that evidence has not been altered.",
        "desc": "Explains evidence integrity verification.",
    },
    "ESP Response Disclaimer": {
        "name": "esp_response_disclaimer",
        "text": "Information provided by the Electronic Service Provider (ESP) in this "
                "CyberTip has not been independently verified by the investigating agency "
                "unless otherwise noted. Subscriber information and account details are "
                "as reported by the ESP.",
        "desc": "Notes that ESP-provided data is unverified.",
    },
    "Jurisdiction Statement": {
        "name": "jurisdiction_statement",
        "text": "Based on the IP address geolocation and/or subscriber information "
                "contained in this CyberTip, the reported activity appears to have "
                "occurred within the jurisdiction of [AGENCY_NAME].",
        "desc": "Establishes jurisdictional basis for the investigation.",
    },
    "Sensitivity Warning": {
        "name": "sensitivity_warning",
        "text": "WARNING: This report contains references to child sexual abuse material "
                "(CSAM). The content described herein is disturbing in nature. This report "
                "should only be reviewed by authorized personnel.",
        "desc": "Content sensitivity warning for report readers.",
    },
    "Case Assignment Note": {
        "name": "case_assignment",
        "text": "This CyberTip has been assigned to [INVESTIGATOR_NAME] for follow-up "
                "investigation. Case number: [CASE_NUMBER]. Date assigned: [CURRENT_DATE].",
        "desc": "Records case assignment and tracking information.",
    },
    "Duplicate from Existing...": {
        "name": "",
        "text": "",
        "desc": "Copy text from an existing custom statement.",
    },
}

PLACEMENT_DESCRIPTIONS: Dict[str, str] = {
    "At Beginning of Report": "Appears before the introduction paragraph at the very top of the report.",
    "Before Incident Summary": "Appears just above the INCIDENT SUMMARY section header.",
    "After Incident Summary": "Appears directly below the incident details (type, date, ESP).",
    "Before Suspect Information": "Appears just above the SUSPECT INFORMATION section header.",
    "After Suspect Information": "Appears after all suspect/person details.",
    "Before Evidence Summary": "Appears just above the EVIDENCE SUMMARY section header.",
    "After Evidence Summary": "Appears after all evidence file details.",
    "Before IP Address Analysis": "Appears just above the IP ADDRESS ANALYSIS section header.",
    "After IP Address Analysis": "Appears after all IP lookup results.",
    "At End of Report": "Appears in a CUSTOM STATEMENTS block at the very end of the report.",
}

HIGHLIGHT_OPTIONS = ["(none)", "yellow", "green", "cyan", "magenta", "red", "blue"]


# ── Add New Tab ───────────────────────────────────────────────────

def _build_add_new_tab(notebook: ttk.Notebook, stmts: StatementManager) -> None:
    outer = ttk.Frame(notebook)
    notebook.add(outer, text="Add New")
    canvas = tk.Canvas(outer, highlightthickness=0)
    vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    add_frame = ttk.Frame(canvas)
    add_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=add_frame, anchor="nw")
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel, add="+")

    # ── Step 1: Starting Point ────────────────────────────────────
    step1 = ttk.LabelFrame(add_frame, text="  Step 1 -- Choose a Starting Point  ", padding=10)
    step1.pack(fill="x", padx=10, pady=(10, 5))

    template_names = ["None (blank)"] + list(STATEMENT_TEMPLATES.keys())
    tmpl_var = tk.StringVar(value=template_names[0])

    tk.Label(step1, text="Start with a template to pre-fill common text, or begin from scratch:").pack(anchor="w")
    tmpl_frame = ttk.Frame(step1)
    tmpl_frame.pack(fill="x", pady=5)
    tk.Label(tmpl_frame, text="Template:").pack(side=tk.LEFT, padx=(0, 5))
    tmpl_cb = ttk.Combobox(tmpl_frame, textvariable=tmpl_var, values=template_names,
                           state="readonly", width=35)
    tmpl_cb.pack(side=tk.LEFT)

    tmpl_desc_var = tk.StringVar(value="")
    tmpl_desc_label = tk.Label(step1, textvariable=tmpl_desc_var, fg="gray",
                               wraplength=700, justify=tk.LEFT)
    tmpl_desc_label.pack(anchor="w", pady=(2, 0))

    # ── Step 2: Statement Identity ────────────────────────────────
    step2 = ttk.LabelFrame(add_frame, text="  Step 2 -- Name & Placement  ", padding=10)
    step2.pack(fill="x", padx=10, pady=5)

    name_row = ttk.Frame(step2)
    name_row.pack(fill="x", pady=3)
    tk.Label(name_row, text="Statement Name:").pack(side=tk.LEFT, padx=(0, 5))
    key_entry = tk.Entry(name_row, width=40)
    key_entry.pack(side=tk.LEFT)
    tk.Label(name_row, text="(unique identifier, e.g. 'chain_of_custody')", fg="gray").pack(side=tk.LEFT, padx=5)

    place_row = ttk.Frame(step2)
    place_row.pack(fill="x", pady=3)
    tk.Label(place_row, text="Insert Location:").pack(side=tk.LEFT, padx=(0, 5))
    placement_opts = list(PLACEMENT_PREFIXES.keys())
    placement_var = tk.StringVar(value="At End of Report")
    place_cb = ttk.Combobox(place_row, textvariable=placement_var, values=placement_opts,
                            state="readonly", width=35)
    place_cb.pack(side=tk.LEFT)

    place_desc_var = tk.StringVar(value=PLACEMENT_DESCRIPTIONS.get("At End of Report", ""))
    place_desc_label = tk.Label(step2, textvariable=place_desc_var, fg="gray",
                                wraplength=700, justify=tk.LEFT)
    place_desc_label.pack(anchor="w", pady=(2, 0))

    # Sort order
    order_row = ttk.Frame(step2)
    order_row.pack(fill="x", pady=3)
    tk.Label(order_row, text="Display Order:").pack(side=tk.LEFT, padx=(0, 5))
    order_var = tk.IntVar(value=10)
    order_spin = tk.Spinbox(order_row, from_=1, to=999, textvariable=order_var, width=5)
    order_spin.pack(side=tk.LEFT)
    tk.Label(order_row, text="(lower = appears first when multiple statements share the same placement)", fg="gray").pack(side=tk.LEFT, padx=5)

    def _on_placement_change(*_args):
        place_desc_var.set(PLACEMENT_DESCRIPTIONS.get(placement_var.get(), ""))
    placement_var.trace_add("write", _on_placement_change)

    # ── Step 3: Conditional Display ───────────────────────────────
    step3 = ttk.LabelFrame(add_frame, text="  Step 3 -- When Should This Statement Appear? (Optional)  ", padding=10)
    step3.pack(fill="x", padx=10, pady=5)

    cond_mode = tk.StringVar(value="always")
    tk.Radiobutton(step3, text="Include in ALL reports (no condition)",
                   variable=cond_mode, value="always").pack(anchor="w")
    tk.Radiobutton(step3, text="Only include when the report is from specific ESP(s):",
                   variable=cond_mode, value="specific").pack(anchor="w", pady=(5, 0))

    esp_check_frame = ttk.Frame(step3)
    esp_check_frame.pack(fill="x", padx=25, pady=5)
    esp_vars: Dict[str, tk.BooleanVar] = {}

    col_count = 4
    for i, esp in enumerate(_ESP_LIST):
        var = tk.BooleanVar(value=False)
        esp_vars[esp] = var
        r, c = divmod(i, col_count)
        cb = ttk.Checkbutton(esp_check_frame, text=esp, variable=var)
        cb.grid(row=r, column=c, sticky="w", padx=5, pady=1)

    custom_esp_row = ttk.Frame(step3)
    custom_esp_row.pack(fill="x", padx=25, pady=(0, 5))
    tk.Label(custom_esp_row, text="Other ESP (type name):").pack(side=tk.LEFT, padx=(0, 5))
    custom_esp_entry = tk.Entry(custom_esp_row, width=25)
    custom_esp_entry.pack(side=tk.LEFT)

    cond_summary_var = tk.StringVar(value="Condition: None (appears in all reports)")
    cond_summary = tk.Label(step3, textvariable=cond_summary_var, font=("Arial", 9, "italic"), fg="blue")
    cond_summary.pack(anchor="w", pady=(5, 0))

    def _update_cond_summary(*_args):
        if cond_mode.get() == "always":
            cond_summary_var.set("Condition: None (appears in all reports)")
            for w in esp_check_frame.winfo_children():
                w.configure(state="disabled")
            custom_esp_entry.configure(state="disabled")
        else:
            for w in esp_check_frame.winfo_children():
                w.configure(state="normal")
            custom_esp_entry.configure(state="normal")
            selected = [name for name, var in esp_vars.items() if var.get()]
            custom = custom_esp_entry.get().strip()
            if custom:
                selected.append(custom)
            if not selected:
                cond_summary_var.set("Condition: (select at least one ESP above)")
            elif len(selected) == 1:
                cond_summary_var.set(f'Condition: Only when ESP is "{selected[0]}"')
            else:
                cond_summary_var.set(f"Condition: Only when ESP is one of: {', '.join(selected)}")

    cond_mode.trace_add("write", _update_cond_summary)
    for var in esp_vars.values():
        var.trace_add("write", _update_cond_summary)
    custom_esp_entry.bind("<KeyRelease>", lambda e: _update_cond_summary())
    _update_cond_summary()

    def _get_condition() -> str:
        if cond_mode.get() == "always":
            return ""
        selected = [name for name, var in esp_vars.items() if var.get()]
        custom = custom_esp_entry.get().strip()
        if custom:
            selected.append(custom)
        if not selected:
            return ""
        if len(selected) == 1:
            return f'esp_name == "{selected[0]}"'
        quoted = ", ".join('"' + v + '"' for v in selected)
        return f"esp_name in [{quoted}]"

    # ── Step 4: Statement Text ────────────────────────────────────
    step4 = ttk.LabelFrame(add_frame, text="  Step 4 -- Write Your Statement  ", padding=10)
    step4.pack(fill="x", padx=10, pady=5)

    toolbar = ttk.Frame(step4)
    toolbar.pack(fill="x", pady=(0, 5))
    tk.Label(toolbar, text="Insert:").pack(side=tk.LEFT, padx=(0, 5))

    add_text = tk.Text(step4, width=90, height=12, wrap="word", font=("Consolas", 10))
    add_text.pack(fill="x", pady=5)

    def _insert_snippet(snippet: str):
        add_text.insert(tk.INSERT, snippet)
        add_text.focus_set()

    snippets = [
        ("Today's Date", "[CURRENT_DATE]"),
        ("Investigator", "[INVESTIGATOR_NAME]"),
        ("Title", "[INVESTIGATOR_TITLE]"),
        ("Case #", "[CASE_NUMBER]"),
        ("Agency", "[AGENCY_NAME]"),
        ("CyberTip #", "[CYBERTIP_NUMBER]"),
        ("ESP", "[ESP_NAME]"),
    ]
    for label, snip in snippets:
        ttk.Button(toolbar, text=label,
                   command=lambda s=snip: _insert_snippet(s)).pack(side=tk.LEFT, padx=2)

    char_count_var = tk.StringVar(value="0 characters")
    tk.Label(step4, textvariable=char_count_var, fg="gray").pack(anchor="e")

    def _update_char_count(*_args):
        length = len(add_text.get(1.0, tk.END).strip())
        char_count_var.set(f"{length} character{'s' if length != 1 else ''}")
    add_text.bind("<KeyRelease>", _update_char_count)

    # ── Step 5: Formatting (Optional) ─────────────────────────────
    step5 = ttk.LabelFrame(add_frame, text="  Step 5 -- DOCX Formatting (Optional)  ", padding=10)
    step5.pack(fill="x", padx=10, pady=5)

    fmt_row1 = ttk.Frame(step5)
    fmt_row1.pack(fill="x", pady=3)

    tk.Label(fmt_row1, text="Font Size:").pack(side=tk.LEFT, padx=(0, 3))
    font_size_var = tk.IntVar(value=12)
    tk.Spinbox(fmt_row1, from_=8, to=28, textvariable=font_size_var, width=4).pack(side=tk.LEFT, padx=(0, 15))

    bold_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(fmt_row1, text="Bold", variable=bold_var).pack(side=tk.LEFT, padx=5)

    italic_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(fmt_row1, text="Italic", variable=italic_var).pack(side=tk.LEFT, padx=5)

    fmt_row2 = ttk.Frame(step5)
    fmt_row2.pack(fill="x", pady=3)

    tk.Label(fmt_row2, text="Left Indent (inches):").pack(side=tk.LEFT, padx=(0, 3))
    indent_var = tk.DoubleVar(value=0.0)
    tk.Spinbox(fmt_row2, from_=0.0, to=3.0, increment=0.25, textvariable=indent_var, width=5).pack(side=tk.LEFT, padx=(0, 15))

    tk.Label(fmt_row2, text="Highlight:").pack(side=tk.LEFT, padx=(0, 3))
    highlight_var = tk.StringVar(value="(none)")
    ttk.Combobox(fmt_row2, textvariable=highlight_var, values=HIGHLIGHT_OPTIONS,
                 state="readonly", width=12).pack(side=tk.LEFT)

    def _get_formatting() -> dict:
        return {
            "font_size": font_size_var.get(),
            "bold": bold_var.get(),
            "italic": italic_var.get(),
            "indent": indent_var.get(),
            "highlight": "" if highlight_var.get() == "(none)" else highlight_var.get(),
        }

    # ── Live Preview ──────────────────────────────────────────────
    step6 = ttk.LabelFrame(add_frame, text="  Preview  ", padding=10)
    step6.pack(fill="x", padx=10, pady=5)

    preview_text = tk.Text(step6, width=90, height=6, wrap="word", state="disabled",
                           bg="#f5f5f5", font=("Times New Roman", 10))
    preview_text.pack(fill="x")

    def _refresh_preview(*_args):
        name = key_entry.get().strip() or "(unnamed)"
        placement = placement_var.get()
        body = add_text.get(1.0, tk.END).strip() or "(no text entered)"
        cond = cond_summary_var.get()
        fmt = _get_formatting()
        fmt_desc_parts = []
        if fmt["bold"]:
            fmt_desc_parts.append("Bold")
        if fmt["italic"]:
            fmt_desc_parts.append("Italic")
        fmt_desc_parts.append(f"{fmt['font_size']}pt")
        if fmt["indent"] > 0:
            fmt_desc_parts.append(f"Indent {fmt['indent']}in")
        if fmt["highlight"]:
            fmt_desc_parts.append(f"Highlight: {fmt['highlight']}")
        fmt_desc = ", ".join(fmt_desc_parts)

        preview_text.configure(state="normal")
        preview_text.delete(1.0, tk.END)
        preview_text.insert(tk.END, f"--- {placement} ---\n\n")
        preview_text.insert(tk.END, f"{name.upper()}: {body}\n\n")
        preview_text.insert(tk.END, f"[{cond}]  |  Format: {fmt_desc}")
        preview_text.configure(state="disabled")

    add_text.bind("<KeyRelease>", lambda e: (_update_char_count(), _refresh_preview()))
    placement_var.trace_add("write", _refresh_preview)
    cond_summary_var.trace_add("write", _refresh_preview)
    key_entry.bind("<KeyRelease>", lambda e: _refresh_preview())
    for fv in (bold_var, italic_var):
        fv.trace_add("write", _refresh_preview)
    font_size_var.trace_add("write", _refresh_preview)
    highlight_var.trace_add("write", _refresh_preview)

    # ── Template application logic ────────────────────────────────
    def _on_template_change(event=None):
        chosen = tmpl_var.get()
        if chosen == "None (blank)":
            tmpl_desc_var.set("")
            return

        tmpl = STATEMENT_TEMPLATES.get(chosen)
        if not tmpl:
            return

        tmpl_desc_var.set(tmpl["desc"])

        if chosen == "Duplicate from Existing...":
            _show_duplicate_picker()
            return

        key_entry.delete(0, tk.END)
        key_entry.insert(0, tmpl["name"])
        add_text.delete(1.0, tk.END)
        add_text.insert(tk.END, tmpl["text"])
        _update_char_count()
        _refresh_preview()

    tmpl_cb.bind("<<ComboboxSelected>>", _on_template_change)

    def _show_duplicate_picker():
        custom = {k: v for k, v in stmts.statements.items() if k not in DEFAULT_STATEMENTS}
        if not custom:
            messagebox.showinfo("Info", "No custom statements to duplicate from.")
            tmpl_var.set(template_names[0])
            return

        picker = tk.Toplevel(add_frame.winfo_toplevel())
        picker.title("Duplicate From Existing Statement")
        picker.geometry("500x350")
        picker.transient(add_frame.winfo_toplevel())
        picker.grab_set()

        tk.Label(picker, text="Select a statement to copy:").pack(pady=5)
        lb = tk.Listbox(picker, width=60, height=12)
        lb.pack(fill="both", expand=True, padx=10, pady=5)
        keys = sorted(custom.keys())
        for k in keys:
            placement = stmts.get_placement_label(k)
            lb.insert(tk.END, f"{k}  ({placement})")

        def _do_duplicate():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("Warning", "Select a statement to duplicate.")
                return
            orig_key = keys[sel[0]]
            val = stmts.statements[orig_key]
            text = val["text"] if isinstance(val, dict) else val
            key_entry.delete(0, tk.END)
            key_entry.insert(0, f"{orig_key}_copy")
            add_text.delete(1.0, tk.END)
            add_text.insert(tk.END, text)
            if isinstance(val, dict) and val.get("condition"):
                _apply_condition_from_string(val["condition"])
            if isinstance(val, dict) and val.get("formatting"):
                fmt = val["formatting"]
                font_size_var.set(fmt.get("font_size", 12))
                bold_var.set(fmt.get("bold", False))
                italic_var.set(fmt.get("italic", False))
                indent_var.set(fmt.get("indent", 0.0))
                highlight_var.set(fmt.get("highlight", "") or "(none)")
            _update_char_count()
            _refresh_preview()
            picker.destroy()

        tk.Button(picker, text="Duplicate", command=_do_duplicate).pack(pady=5)
        tk.Button(picker, text="Cancel", command=picker.destroy).pack(pady=5)

    def _apply_condition_from_string(cond: str):
        if not cond:
            cond_mode.set("always")
            return
        cond_mode.set("specific")
        values = re.findall(r'"([^"]*)"', cond)
        for name, var in esp_vars.items():
            var.set(name in values)
        remaining = [v for v in values if v not in esp_vars]
        if remaining:
            custom_esp_entry.delete(0, tk.END)
            custom_esp_entry.insert(0, remaining[0])

    # ── Action Buttons ────────────────────────────────────────────
    btn_frame = ttk.Frame(add_frame)
    btn_frame.pack(fill="x", padx=10, pady=(5, 15))

    def _add_statement():
        name = key_entry.get().strip()
        if not name:
            messagebox.showwarning("Missing Name", "Please enter a statement name in Step 2.")
            key_entry.focus_set()
            return
        if not name.replace("_", "").replace("-", "").isalnum():
            messagebox.showwarning(
                "Invalid Name",
                "Statement name should only contain letters, numbers, underscores, or hyphens.",
            )
            key_entry.focus_set()
            return

        prefix = PLACEMENT_PREFIXES[placement_var.get()]
        full_key = f"{prefix}{name}" if prefix else name
        if full_key in stmts.statements:
            messagebox.showwarning(
                "Duplicate Name",
                f"A statement named '{name}' already exists at '{placement_var.get()}'.\n"
                "Choose a different name or placement.",
            )
            return

        text = add_text.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("Missing Text", "Please write your statement text in Step 4.")
            add_text.focus_set()
            return

        cond = _get_condition()
        if cond_mode.get() == "specific" and not cond:
            messagebox.showwarning(
                "No ESPs Selected",
                "You chose to limit this statement to specific ESPs but didn't select any.\n"
                "Either select ESP(s) in Step 3 or switch to 'Include in ALL reports'.",
            )
            return

        stmt_data = {
            "text": text,
            "condition": cond,
            "order": order_var.get(),
            "formatting": _get_formatting(),
            "enabled": True,
        }
        stmts.statements[full_key] = stmt_data
        stmts.selected.add(full_key)

        placement_label = placement_var.get()
        cond_label = cond_summary_var.get()
        messagebox.showinfo(
            "Statement Added",
            f"Successfully added:\n\n"
            f"Name: {name}\n"
            f"Location: {placement_label}\n"
            f"Order: {order_var.get()}\n"
            f"{cond_label}",
        )
        _clear_form()

    def _clear_form():
        tmpl_var.set(template_names[0])
        tmpl_desc_var.set("")
        key_entry.delete(0, tk.END)
        placement_var.set("At End of Report")
        order_var.set(10)
        cond_mode.set("always")
        for var in esp_vars.values():
            var.set(False)
        custom_esp_entry.delete(0, tk.END)
        add_text.delete(1.0, tk.END)
        font_size_var.set(12)
        bold_var.set(False)
        italic_var.set(False)
        indent_var.set(0.0)
        highlight_var.set("(none)")
        _update_char_count()
        _refresh_preview()

    ttk.Button(btn_frame, text="Add Statement", command=_add_statement).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="Clear Form", command=_clear_form).pack(side=tk.LEFT, padx=5)

    _refresh_preview()


# ── Manage Custom Statements Tab (full overhaul) ─────────────────

def _build_manage_tab(notebook: ttk.Notebook, stmts: StatementManager) -> None:
    manage = ttk.Frame(notebook)
    notebook.add(manage, text="Manage Custom Statements")

    # Top: search
    top = ttk.Frame(manage)
    top.pack(fill="x", padx=10, pady=5)
    tk.Label(top, text="Search:").pack(side=tk.LEFT, padx=(0, 5))
    search_var = tk.StringVar()
    tk.Entry(top, textvariable=search_var, width=40).pack(side=tk.LEFT)

    # Main pane
    pane = ttk.PanedWindow(manage, orient=tk.HORIZONTAL)
    pane.pack(fill="both", expand=True, padx=10, pady=5)

    # Left: statement list with columns
    left = ttk.Frame(pane)
    pane.add(left, weight=1)

    columns = ("key", "placement", "order", "enabled")
    tree = ttk.Treeview(left, columns=columns, show="headings", height=12)
    tree.heading("key", text="Name")
    tree.heading("placement", text="Placement")
    tree.heading("order", text="Order")
    tree.heading("enabled", text="Enabled")
    tree.column("key", width=180)
    tree.column("placement", width=160)
    tree.column("order", width=50, anchor="center")
    tree.column("enabled", width=60, anchor="center")
    tree.pack(fill="both", expand=True)

    tree_sb = ttk.Scrollbar(left, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_sb.set)
    tree_sb.pack(side="right", fill="y")

    # Right: editing panel
    right = ttk.Frame(pane)
    pane.add(right, weight=2)

    tk.Label(right, text="Statement Text:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
    edit_text = tk.Text(right, width=55, height=8, wrap="word", font=("Consolas", 10))
    edit_text.pack(fill="x", pady=2)

    # Condition editor
    cond_frame = ttk.LabelFrame(right, text="  Condition  ", padding=5)
    cond_frame.pack(fill="x", pady=5)
    tk.Label(cond_frame, text="Raw condition string:").pack(anchor="w")
    cond_entry = tk.Entry(cond_frame, width=50)
    cond_entry.pack(fill="x", pady=2)
    tk.Label(cond_frame, text='Leave blank for "all reports". e.g. esp_name == "Facebook"', fg="gray").pack(anchor="w")

    # Placement reassignment
    place_frame = ttk.LabelFrame(right, text="  Placement  ", padding=5)
    place_frame.pack(fill="x", pady=5)
    manage_place_var = tk.StringVar(value="At End of Report")
    ttk.Combobox(place_frame, textvariable=manage_place_var, values=list(PLACEMENT_PREFIXES.keys()),
                 state="readonly", width=35).pack(fill="x", pady=2)

    # Order
    order_frame = ttk.Frame(right)
    order_frame.pack(fill="x", pady=3)
    tk.Label(order_frame, text="Display Order:").pack(side=tk.LEFT, padx=(0, 5))
    manage_order_var = tk.IntVar(value=10)
    tk.Spinbox(order_frame, from_=1, to=999, textvariable=manage_order_var, width=5).pack(side=tk.LEFT)

    # Formatting
    fmt_frame = ttk.LabelFrame(right, text="  Formatting  ", padding=5)
    fmt_frame.pack(fill="x", pady=5)
    fmt_r1 = ttk.Frame(fmt_frame)
    fmt_r1.pack(fill="x", pady=2)
    tk.Label(fmt_r1, text="Font Size:").pack(side=tk.LEFT, padx=(0, 3))
    m_fontsize = tk.IntVar(value=12)
    tk.Spinbox(fmt_r1, from_=8, to=28, textvariable=m_fontsize, width=4).pack(side=tk.LEFT, padx=(0, 10))
    m_bold = tk.BooleanVar(value=False)
    ttk.Checkbutton(fmt_r1, text="Bold", variable=m_bold).pack(side=tk.LEFT, padx=3)
    m_italic = tk.BooleanVar(value=False)
    ttk.Checkbutton(fmt_r1, text="Italic", variable=m_italic).pack(side=tk.LEFT, padx=3)

    fmt_r2 = ttk.Frame(fmt_frame)
    fmt_r2.pack(fill="x", pady=2)
    tk.Label(fmt_r2, text="Indent:").pack(side=tk.LEFT, padx=(0, 3))
    m_indent = tk.DoubleVar(value=0.0)
    tk.Spinbox(fmt_r2, from_=0.0, to=3.0, increment=0.25, textvariable=m_indent, width=5).pack(side=tk.LEFT, padx=(0, 10))
    tk.Label(fmt_r2, text="Highlight:").pack(side=tk.LEFT, padx=(0, 3))
    m_highlight = tk.StringVar(value="(none)")
    ttk.Combobox(fmt_r2, textvariable=m_highlight, values=HIGHLIGHT_OPTIONS, state="readonly", width=12).pack(side=tk.LEFT)

    _selected_key = tk.StringVar(value="")

    def _refresh_tree(q=""):
        tree.delete(*tree.get_children())
        custom = {k: v for k, v in stmts.statements.items() if k not in DEFAULT_STATEMENTS}
        for key in sorted(custom, key=lambda k: (stmts.get_order(k), k)):
            val = custom[key]
            text_val = val["text"] if isinstance(val, dict) else val
            if q and q.lower() not in key.lower() and q.lower() not in text_val.lower():
                continue
            placement = stmts.get_placement_label(key)
            order = stmts.get_order(key)
            enabled = key in stmts.selected
            tree.insert("", tk.END, iid=key, values=(key, placement, order, "Yes" if enabled else "No"))

    def _on_tree_select(event=None):
        sel = tree.selection()
        if not sel:
            return
        key = sel[0]
        _selected_key.set(key)
        val = stmts.statements.get(key)
        if val is None:
            return

        edit_text.delete(1.0, tk.END)
        if isinstance(val, dict):
            edit_text.insert(tk.END, val.get("text", ""))
            cond_entry.delete(0, tk.END)
            cond_entry.insert(0, val.get("condition", ""))
            manage_order_var.set(val.get("order", 10))
            fmt = val.get("formatting", {})
            m_fontsize.set(fmt.get("font_size", 12))
            m_bold.set(fmt.get("bold", False))
            m_italic.set(fmt.get("italic", False))
            m_indent.set(fmt.get("indent", 0.0))
            m_highlight.set(fmt.get("highlight", "") or "(none)")
        else:
            edit_text.insert(tk.END, val)
            cond_entry.delete(0, tk.END)
            manage_order_var.set(10)
            m_fontsize.set(12)
            m_bold.set(False)
            m_italic.set(False)
            m_indent.set(0.0)
            m_highlight.set("(none)")

        current_placement = stmts.get_placement_label(key)
        manage_place_var.set(current_placement)

    tree.bind("<<TreeviewSelect>>", _on_tree_select)
    search_var.trace("w", lambda *_: _refresh_tree(search_var.get()))

    # Action buttons
    actions = ttk.Frame(right)
    actions.pack(fill="x", pady=5)

    def _save_edits():
        key = _selected_key.get()
        if not key or key not in stmts.statements:
            messagebox.showwarning("Warning", "Select a statement first.")
            return

        new_text = edit_text.get(1.0, tk.END).strip()
        if not new_text:
            messagebox.showwarning("Warning", "Statement text cannot be empty.")
            return

        new_placement_label = manage_place_var.get()
        new_prefix = PLACEMENT_PREFIXES.get(new_placement_label, "")

        # Extract the bare name (strip old prefix)
        bare_name = key
        for pfx in sorted(PLACEMENT_PREFIXES.values(), key=len, reverse=True):
            if pfx and key.startswith(pfx):
                bare_name = key[len(pfx):]
                break

        new_key = f"{new_prefix}{bare_name}" if new_prefix else bare_name

        new_val = {
            "text": new_text,
            "condition": cond_entry.get().strip(),
            "order": manage_order_var.get(),
            "formatting": {
                "font_size": m_fontsize.get(),
                "bold": m_bold.get(),
                "italic": m_italic.get(),
                "indent": m_indent.get(),
                "highlight": "" if m_highlight.get() == "(none)" else m_highlight.get(),
            },
            "enabled": key in stmts.selected,
        }

        if new_key != key:
            del stmts.statements[key]
            stmts.selected.discard(key)
            stmts.statements[new_key] = new_val
            if new_val["enabled"]:
                stmts.selected.add(new_key)
            _selected_key.set(new_key)
        else:
            stmts.statements[key] = new_val

        messagebox.showinfo("Saved", f"Statement '{bare_name}' updated.")
        _refresh_tree(search_var.get())

    def _toggle_enabled():
        key = _selected_key.get()
        if not key:
            return
        if key in stmts.selected:
            stmts.selected.discard(key)
        else:
            stmts.selected.add(key)
        val = stmts.statements.get(key)
        if isinstance(val, dict):
            val["enabled"] = key in stmts.selected
        _refresh_tree(search_var.get())

    def _delete_stmt():
        key = _selected_key.get()
        if not key:
            return
        if messagebox.askyesno("Confirm", f"Delete statement '{key}'?"):
            del stmts.statements[key]
            stmts.selected.discard(key)
            _selected_key.set("")
            edit_text.delete(1.0, tk.END)
            cond_entry.delete(0, tk.END)
            _refresh_tree(search_var.get())

    def _move_order(delta: int):
        key = _selected_key.get()
        if not key:
            return
        val = stmts.statements.get(key)
        if isinstance(val, dict):
            val["order"] = max(1, val.get("order", 10) + delta)
            manage_order_var.set(val["order"])
        _refresh_tree(search_var.get())

    ttk.Button(actions, text="Save Changes", command=_save_edits).pack(side=tk.LEFT, padx=3)
    ttk.Button(actions, text="Toggle Enabled", command=_toggle_enabled).pack(side=tk.LEFT, padx=3)
    ttk.Button(actions, text="Delete", command=_delete_stmt).pack(side=tk.LEFT, padx=3)
    ttk.Button(actions, text="Move Up", command=lambda: _move_order(-1)).pack(side=tk.LEFT, padx=3)
    ttk.Button(actions, text="Move Down", command=lambda: _move_order(1)).pack(side=tk.LEFT, padx=3)

    _refresh_tree()


# ── Choose Statements ────────────────────────────────────────────

def show_choose_statements_dialog(
    root: tk.Tk,
    stmts: StatementManager,
    json_path: str,
    generate_preview_fn=None,
) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Choose Statements for Report")
    dialog.geometry("800x600")
    dialog.transient(root)
    dialog.grab_set()

    tk.Label(dialog, text="Select statements to include in the report (all checked by default):").pack(pady=5)
    frame = ttk.Frame(dialog)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    canvas = tk.Canvas(frame)
    sb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    sf = ttk.Frame(canvas)
    sf.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=sf, anchor="nw")
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    check_vars: Dict[str, tk.BooleanVar] = {}
    for key in sorted(stmts.statements.keys()):
        var = tk.BooleanVar(value=key in stmts.selected)
        check_vars[key] = var
        placement = stmts.get_placement_label(key)
        val = stmts.statements[key]
        cond = val.get("condition", "") if isinstance(val, dict) else ""
        ttk.Checkbutton(sf, text=f"{key} ({placement}) [Condition: {cond or 'None'}]", variable=var).pack(anchor="w", pady=2)

    def save_selection():
        stmts.selected = {k for k, v in check_vars.items() if v.get()}
        messagebox.showinfo("Success", "Statement selection saved.")
        dialog.destroy()

    def preview():
        if not generate_preview_fn:
            return
        text = generate_preview_fn()
        if not text:
            return
        pw = tk.Toplevel(dialog)
        pw.title("Report Preview")
        pw.geometry("600x400")
        pw.transient(dialog)
        pw.grab_set()
        tk.Label(pw, text="Report Preview (first 1000 characters):").pack(pady=5)
        ta = scrolledtext.ScrolledText(pw, width=80, height=20, wrap="word")
        ta.pack(pady=5)
        ta.insert(tk.END, text[:1000])
        ta.config(state="disabled")
        tk.Button(pw, text="Close", command=pw.destroy).pack(pady=5)

    tk.Button(dialog, text="Preview Report", command=preview).pack(pady=5)
    tk.Button(dialog, text="Save Selection", command=save_selection).pack(pady=5)
    tk.Button(dialog, text="Cancel", command=dialog.destroy).pack(pady=5)


# ── Report Template Editor (full overhaul) ────────────────────────

SECTION_LABELS = {
    "incident_summary": "Incident Summary",
    "suspect_information": "Suspect Information",
    "evidence_summary": "Evidence Summary",
    "ip_analysis": "IP Address Analysis",
}


def show_template_dialog(root: tk.Tk, template: ReportTemplate) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Report Template Settings")
    dialog.geometry("750x700")
    dialog.transient(root)
    dialog.grab_set()

    notebook = ttk.Notebook(dialog)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    # ── Tab 1: Section Order & Visibility ─────────────────────────
    tab_order = ttk.Frame(notebook)
    notebook.add(tab_order, text="Section Order & Visibility")

    tk.Label(tab_order, text="Configure which sections appear and in what order.",
             wraplength=600).pack(pady=5)

    order_frame = ttk.Frame(tab_order)
    order_frame.pack(fill="both", expand=True, padx=10, pady=5)

    listbox = tk.Listbox(order_frame, width=45, height=8, font=("Arial", 11))
    listbox.pack(side=tk.LEFT, fill="both", expand=True)

    def _refresh_listbox():
        listbox.delete(0, tk.END)
        for s in template.section_order:
            label = template.get_section_name(s)
            prefix = "[x]" if template.section_visible.get(s, True) else "[ ]"
            listbox.insert(tk.END, f"{prefix}  {label}")

    _refresh_listbox()

    btn_frame = ttk.Frame(order_frame)
    btn_frame.pack(side=tk.RIGHT, padx=10)

    def move_up():
        sel = listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        template.section_order[idx], template.section_order[idx - 1] = (
            template.section_order[idx - 1], template.section_order[idx]
        )
        _refresh_listbox()
        listbox.selection_set(idx - 1)

    def move_down():
        sel = listbox.curselection()
        if not sel or sel[0] >= len(template.section_order) - 1:
            return
        idx = sel[0]
        template.section_order[idx], template.section_order[idx + 1] = (
            template.section_order[idx + 1], template.section_order[idx]
        )
        _refresh_listbox()
        listbox.selection_set(idx + 1)

    def toggle_visible():
        sel = listbox.curselection()
        if not sel:
            return
        section = template.section_order[sel[0]]
        template.section_visible[section] = not template.section_visible[section]
        _refresh_listbox()
        listbox.selection_set(sel[0])

    tk.Button(btn_frame, text="Move Up", command=move_up, width=14).pack(pady=3)
    tk.Button(btn_frame, text="Move Down", command=move_down, width=14).pack(pady=3)
    tk.Button(btn_frame, text="Toggle Visible", command=toggle_visible, width=14).pack(pady=3)

    # ── Tab 2: Section Renaming ───────────────────────────────────
    tab_rename = ttk.Frame(notebook)
    notebook.add(tab_rename, text="Section Names")

    tk.Label(tab_rename, text="Customize the header text for each report section.",
             wraplength=600).pack(pady=10)

    name_entries: Dict[str, tk.Entry] = {}
    for section_id in DEFAULT_SECTIONS:
        row = ttk.Frame(tab_rename)
        row.pack(fill="x", padx=20, pady=4)
        default_label = SECTION_LABELS.get(section_id, section_id)
        tk.Label(row, text=f"{default_label}:", width=22, anchor="w").pack(side=tk.LEFT)
        ent = tk.Entry(row, width=40)
        ent.pack(side=tk.LEFT, padx=5)
        ent.insert(0, template.section_names.get(section_id, DEFAULT_SECTION_NAMES.get(section_id, "")))
        name_entries[section_id] = ent

        def _reset_name(sid=section_id, entry=ent):
            entry.delete(0, tk.END)
            entry.insert(0, DEFAULT_SECTION_NAMES.get(sid, ""))
        tk.Button(row, text="Reset", command=_reset_name).pack(side=tk.LEFT, padx=5)

    def _apply_names():
        for sid, ent in name_entries.items():
            template.section_names[sid] = ent.get().strip() or DEFAULT_SECTION_NAMES.get(sid, "")
        _refresh_listbox()
        messagebox.showinfo("Applied", "Section names updated.")

    tk.Button(tab_rename, text="Apply Names", command=_apply_names).pack(pady=10)

    # ── Tab 3: Custom Sections ────────────────────────────────────
    tab_custom = ttk.Frame(notebook)
    notebook.add(tab_custom, text="Custom Sections")

    tk.Label(tab_custom, text="Add custom named sections with free-form body text.\n"
             "These appear after the standard sections.",
             wraplength=600).pack(pady=10)

    cs_list = tk.Listbox(tab_custom, width=60, height=6)
    cs_list.pack(fill="x", padx=20, pady=5)

    def _refresh_cs_list():
        cs_list.delete(0, tk.END)
        for cs in template.custom_sections:
            cs_list.insert(tk.END, cs.get("name", "(unnamed)"))

    _refresh_cs_list()

    cs_name_row = ttk.Frame(tab_custom)
    cs_name_row.pack(fill="x", padx=20, pady=3)
    tk.Label(cs_name_row, text="Section Name:").pack(side=tk.LEFT, padx=(0, 5))
    cs_name_entry = tk.Entry(cs_name_row, width=30)
    cs_name_entry.pack(side=tk.LEFT)

    tk.Label(tab_custom, text="Section Body (supports placeholders):").pack(anchor="w", padx=20)
    cs_body = tk.Text(tab_custom, width=60, height=6, wrap="word")
    cs_body.pack(fill="x", padx=20, pady=5)

    def _add_custom_section():
        name = cs_name_entry.get().strip()
        body = cs_body.get(1.0, tk.END).strip()
        if not name:
            messagebox.showwarning("Warning", "Enter a section name.")
            return
        template.custom_sections.append({"name": name, "body": body})
        cs_name_entry.delete(0, tk.END)
        cs_body.delete(1.0, tk.END)
        _refresh_cs_list()

    def _remove_custom_section():
        sel = cs_list.curselection()
        if not sel:
            return
        template.custom_sections.pop(sel[0])
        _refresh_cs_list()

    cs_btns = ttk.Frame(tab_custom)
    cs_btns.pack(fill="x", padx=20, pady=5)
    ttk.Button(cs_btns, text="Add Section", command=_add_custom_section).pack(side=tk.LEFT, padx=3)
    ttk.Button(cs_btns, text="Remove Selected", command=_remove_custom_section).pack(side=tk.LEFT, padx=3)

    # ── Tab 4: Preview ────────────────────────────────────────────
    tab_preview = ttk.Frame(notebook)
    notebook.add(tab_preview, text="Preview")

    tk.Label(tab_preview, text="Skeleton preview of the report layout based on current settings.",
             wraplength=600).pack(pady=5)

    preview_text = scrolledtext.ScrolledText(tab_preview, width=70, height=20, wrap="word",
                                              font=("Consolas", 10), state="disabled", bg="#fafafa")
    preview_text.pack(fill="both", expand=True, padx=10, pady=5)

    def _refresh_preview():
        lines = []
        lines.append("=" * 50)
        lines.append("REPORT SKELETON PREVIEW")
        lines.append("=" * 50)
        lines.append("")
        lines.append("[Introduction Paragraph]")
        lines.append("On [CURRENT_DATE], I, [INVESTIGATOR_TITLE] [INVESTIGATOR_NAME], ...")
        lines.append("")

        for section in template.section_order:
            visible = template.section_visible.get(section, True)
            name = template.get_section_name(section)
            marker = "" if visible else "  [HIDDEN]"
            lines.append(f"{'─' * 40}")
            lines.append(f"{name}:{marker}")
            lines.append(f"  (content for {SECTION_LABELS.get(section, section)})")
            lines.append("")

        for cs in template.custom_sections:
            lines.append(f"{'─' * 40}")
            lines.append(f"{cs.get('name', 'CUSTOM').upper()}:")
            body_preview = cs.get("body", "")[:80]
            lines.append(f"  {body_preview}{'...' if len(cs.get('body', '')) > 80 else ''}")
            lines.append("")

        lines.append("─" * 40)
        lines.append("[Custom Statements at End of Report]")
        lines.append("=" * 50)

        preview_text.configure(state="normal")
        preview_text.delete(1.0, tk.END)
        preview_text.insert(tk.END, "\n".join(lines))
        preview_text.configure(state="disabled")

    ttk.Button(tab_preview, text="Refresh Preview", command=_refresh_preview).pack(pady=5)
    _refresh_preview()

    # ── Bottom buttons ────────────────────────────────────────────
    bottom = ttk.Frame(dialog)
    bottom.pack(fill="x", padx=10, pady=10)

    def reset():
        template.section_order = list(DEFAULT_SECTIONS)
        template.section_visible = {s: True for s in DEFAULT_SECTIONS}
        template.section_names = dict(DEFAULT_SECTION_NAMES)
        template.custom_sections = []
        _refresh_listbox()
        for sid, ent in name_entries.items():
            ent.delete(0, tk.END)
            ent.insert(0, DEFAULT_SECTION_NAMES.get(sid, ""))
        _refresh_cs_list()
        _refresh_preview()

    def save_and_close():
        _apply_names()
        messagebox.showinfo("Success", "Template settings saved.")
        dialog.destroy()

    ttk.Button(bottom, text="Reset to Default", command=reset).pack(side=tk.LEFT, padx=5)
    ttk.Button(bottom, text="Save and Close", command=save_and_close).pack(side=tk.RIGHT, padx=5)
    ttk.Button(bottom, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)


# ── Report Comparison ─────────────────────────────────────────────

def show_comparison_dialog(root: tk.Tk) -> None:
    dialog = tk.Toplevel(root)
    dialog.title("Compare CyberTip Reports")
    dialog.geometry("900x700")
    dialog.transient(root)
    dialog.grab_set()

    tk.Label(dialog, text="Load multiple CyberTip JSON files to find common identifiers.", font=("Arial", 11)).pack(pady=10)

    files_frame = ttk.Frame(dialog)
    files_frame.pack(fill="x", padx=10, pady=5)

    files_list = tk.Listbox(files_frame, width=80, height=5)
    files_list.pack(side=tk.LEFT, fill="both", expand=True)
    loaded_files: List[str] = []

    btn_frame = ttk.Frame(files_frame)
    btn_frame.pack(side=tk.RIGHT, padx=5)

    def add_file():
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if path and path not in loaded_files:
            loaded_files.append(path)
            files_list.insert(tk.END, os.path.basename(path))

    def remove_file():
        sel = files_list.curselection()
        if sel:
            loaded_files.pop(sel[0])
            files_list.delete(sel[0])

    tk.Button(btn_frame, text="Add File", command=add_file, width=12).pack(pady=3)
    tk.Button(btn_frame, text="Remove", command=remove_file, width=12).pack(pady=3)

    results_text = scrolledtext.ScrolledText(dialog, width=100, height=30, wrap="word")
    results_text.pack(fill="both", expand=True, padx=10, pady=10)

    def run_comparison():
        if len(loaded_files) < 2:
            messagebox.showwarning("Warning", "Load at least 2 files to compare.")
            return

        results_text.delete(1.0, tk.END)
        all_data: Dict[str, dict] = {}

        for path in loaded_files:
            data = load_json(path)
            if data is None:
                results_text.insert(tk.END, f"ERROR: Could not load {os.path.basename(path)}\n")
                continue
            comparison = extract_comparison_data(data)
            report_id = data.get("reportId", os.path.basename(path))
            all_data[report_id] = comparison

        if len(all_data) < 2:
            results_text.insert(tk.END, "Not enough valid files to compare.\n")
            return

        results_text.insert(tk.END, f"Comparing {len(all_data)} CyberTip reports\n")
        results_text.insert(tk.END, "=" * 60 + "\n\n")

        report_ids = list(all_data.keys())

        for category in ("ips", "emails", "phones", "screen_names", "user_ids", "hashes"):
            label = category.replace("_", " ").title()
            sets = [all_data[rid].get(category, set()) for rid in report_ids]
            if not any(sets):
                continue

            common = sets[0]
            for s in sets[1:]:
                common = common & s

            if common:
                results_text.insert(tk.END, f"COMMON {label.upper()} ({len(common)}):\n")
                for val in sorted(common):
                    present_in = [rid for rid in report_ids if val in all_data[rid].get(category, set())]
                    results_text.insert(tk.END, f"  {val}  (in tips: {', '.join(present_in)})\n")
                results_text.insert(tk.END, "\n")

        results_text.insert(tk.END, "\nPER-REPORT SUMMARY:\n" + "-" * 40 + "\n")
        for rid, comp in all_data.items():
            results_text.insert(tk.END, f"\nTip #{rid}:\n")
            for cat, vals in comp.items():
                if vals:
                    results_text.insert(tk.END, f"  {cat.replace('_',' ').title()}: {len(vals)}\n")

    tk.Button(dialog, text="Compare", command=run_comparison).pack(pady=5)
    tk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=5)
