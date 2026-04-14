
# CAT-RG: CyberTip Analysis Tool & Report Generator

###  Click on "Releases" (to the right) and download the .exe file

## Introduction:

CAT-RG (CyberTip Analysis Tool & Report Generator) is a standalone Windows application designed to analyze CyberTipline JSON reports from the National Center for Missing and Exploited Children (NCMEC). It extracts key information such as suspect details, incident summaries, evidence (e.g., uploaded files), and IP addresses, then generates formatted reports in DOCX format. The tool also supports optional exports to text and Excel files for IP data and evidence summaries.

This tool supports various Electronic Service Providers (ESPs) and uses external services like MaxMind for geolocation and ARIN for WHOIS data on IP addresses.

Version: 2.2
Release Date: April 7, 2026
Platform: Windows

## What's new in 2.2 (since 2.1)

- **Version bump:** CAT-RG is now **2.2** everywhere it is shown (Help > About, `catrg.__version__`, GitHub update checks).
- **Compare Reports** feature (`load_json` / `extract_comparison_data` from `catrg.core.parser`). Includes drag-and-drop, DOCX and plain-text export

## Features:

- **JSON Report Analysis:** Parses CyberTipline JSON files to extract incident summaries, suspect information, evidence details (e.g., files, URLs, chats), and IP addresses.

- **Drag & Drop Support:** Drag and drop a .json file directly onto the application window to load it (requires tkinterdnd2).

- **Report Generation:** Automatically creates a detailed DOCX report with sections for Incident Summary, Suspect Information, Evidence Summary, and IP Address Analysis.
  - "Customize Report Statements" feature provides the ability to add/modify/or delete statements that are exported in the report.
  - Allows user to edit the pre-formatted report statements or create new statements to be included in the report.
  - Allows for conditional formatting of customized report statements (for example: if ESP is Facebook, add your own custom statement to a specific section of the report).

- **Report Templates:** Save, load, and manage report template profiles that control section visibility, ordering, and naming. Templates can be imported/exported for sharing between users.

- **IP Address Analysis:** Extracts unique IPs, queries MaxMind for geolocation (city/country) and ARIN for ownership (organization). Handles IPv4/IPv6, with optional query capping for large reports (>50 IPs).

- **Exports:**
  - DOCX: Formatted police-style "initial" report.
  - Text (.txt): Plain text version of the full report.
  - Excel (.xlsx) for IP Data: Spreadsheet with IP, Date/Time, Port, Event, MaxMind Country/City, ARIN Organization.
  - Excel (.xlsx) for Evidence: Spreadsheet with file details like Name, Hash, Upload Time, Tags, IP, and webpage info (e.g., URLs for X/Twitter or Reddit).

- **Credential Management:**
  - Required: MaxMind Account ID and License Key for geolocation.
  - Optional: ARIN API Key for higher WHOIS query limits (increases from 15/min, 256/day to 60/min, 1,000/day).
 

- **Investigator Info:** Prompts for name and title on first run; used in reports.

- **UI Elements:**
  - Browse button to select JSON files.
  - Drag & drop zone for loading .json files.
  - Analyze button to process and generate reports.
  - Output text area to preview reports.
  - Status label and progress bar for progress updates.
  - Menu bar with File (Open, Save, Export, etc.) and Help options.
  - Recent Files submenu for quick access (up to 5 files).
  - Keyboard shortcuts: Ctrl+O (Open), Ctrl+S (Save), Ctrl+N (New Analysis).

- **Supported ESPs:** Discord, Dropbox, Facebook, Google, Imgur, Instagram, Kik, MeetMe, Microsoft (including BingImage), Reddit, Roblox, Snapchat, Sony, Synchronoss, TikTok, WhatsApp, X (Twitter), Yahoo. Other ESPs may work but are unverified.

- **Additional Notes:**
  - All user data (credentials, investigator info, recent files, custom statements, report templates, logs) is stored in `%LOCALAPPDATA%\CATRG` (typically `C:\Users\[username]\AppData\Local\CATRG\`).
  - Threaded analysis for responsiveness during long queries.
  - Error handling for invalid JSON, missing credentials, or query failures.
  - "Check for Updates" feature checks version against latest release in GitHub Repo.
  - Multi-file comparison: Analyze and compare multiple CyberTip JSON files side by side.

## Installation and Setup

Since CAT-RG is packaged as a standalone .exe using PyInstaller, no Python or additional libraries are needed on the end-user's machine. All dependencies (e.g., tkinter, openpyxl, python-docx, requests, Pillow) are bundled.

1. **Download the .exe:**
   - Obtain the CAT-RG.exe file from the Releases page or distribution source.

2. **Run the Application:**
   - Double-click CAT-RG.exe to launch.
   - On first run, the application will walk you through three setup dialogs in order:
     1. Investigator name and title (required for reports).
     2. MaxMind credentials (see "Creating API Keys" below). You may skip this, but geolocation will not be available.
     3. ARIN API Key (optional, recommended for high-volume use). You may skip this.
   - On subsequent launches, you will only be prompted for credentials that are still missing.

3. **Data Storage:**
   - All configuration and user data is stored in `%LOCALAPPDATA%\CATRG\` (typically `C:\Users\[username]\AppData\Local\CATRG\`).
   - This includes:
     - `investigator_info.json` -- your name and title
     - `maxmind_credentials.json` -- MaxMind credentials (if not using keyring)
     - `arin_credentials.json` -- ARIN API key (if not using keyring)
     - `custom_statements.json` -- customized report statements
     - `recent_files.json` -- recently opened files
     - `report_templates/` -- saved report template profiles
     - `catrg.log` -- application log
   - Do not delete these files unless you want to reset your configuration.

4. **Dependencies (for running from source):**
   - If running from source (not the .exe), open a terminal in the **`V2.2`** folder, ensure Python 3.13+, and install packages via `pip install -r requirements.txt`.


# How to Use

## Starting the Application

- Launch CAT-RG.exe.
- The main window shows:
  - "CAT-RG" title and optional logo.
  - Drag & drop zone for .json files.
  - File entry field and "Browse to .json" button.
  - "Analyze CyberTip" button.
  - Status label and progress bar (starts as "Ready").
  - Output text area for report previews.

## Analyzing a CyberTip Report

1. **Select a JSON File:**
   - Click "Browse to .json" or use File > Open File (Ctrl+O).
   - Or drag and drop a .json file onto the application window.
   - Choose a CyberTipline JSON file.
   - Recent files appear under File > Recent Files for quick access.

2. **Run Analysis:**
   - Click "Analyze CyberTip" or use File > New Analysis (Ctrl+N).
   - The tool:
     - Extracts data (incident, suspects, evidence, IPs).
     - Queries MaxMind/ARIN for IPs (with prompt if >50 IPs).
     - Generates a report in the output text area.
     - Automatically prompts to save the DOCX report.

3. **View and Save Report:**
   - Preview in the output text area.
   - Save via File > Save Report As (Ctrl+S) -- creates a DOCX with formatted sections, bold headers, and highlights (e.g., red for Investigator's Description).

4. **Exports:**
   - Text: File > Export as Text -- saves the full report as .txt.
   - IP Data Excel: File > Export IP Data to Excel -- saves spreadsheet with IP details (re-queries if needed).
   - Evidence Excel: File > Export Evidence to Excel -- saves spreadsheet with file/webpage details.

5. **Other Actions:**
   - File > Clear Output: Clears the text area and resets file path.
   - File > New Analysis: Clears, opens file dialog, and analyzes.
   - File > Exit: Closes the app (saves recent files).
   - Help > About: Shows version and developer info.
   - Help > Help: Detailed usage instructions.
   - Help > Update MaxMind Credentials: Re-prompt for MaxMind keys.
   - Help > Update Investigator Info: Re-prompt for name/title.
   - Help > Update ARIN API Key: Re-prompt for ARIN key.
   - Help > Check for Updates: Check for newer versions on GitHub.

## Report Structure

- **Incident Summary:** Type, date/time, reported by ESP, description.
- **Suspect Information:** Name, DOB, email, phone, address, screen name, etc. (ESP-specific details like MeetMe profile).
- **Evidence Summary:** File details (name, hash, upload time, tags, IP), viewed by ESP, investigator description. Includes webpage/URL info for X/Twitter or Reddit.
- **IP Address Analysis:** Unique IPs with occurrences (date/time, port, event), MaxMind geolocation, ARIN ownership.

## Creating API Keys

### MaxMind Credentials (Required)

MaxMind provides geolocation data. Credentials are free for basic use.

1. **Sign Up:**
   - Go to https://www.maxmind.com/en/geolite2/signup.
   - Create a free account (email verification required).

2. **Generate License Key:**
   - Log in to your MaxMind account.
   - Under "Account" > "Services" > "License Key," generate a new key.
   - SAVE your Account ID (displayed on the dashboard) and License Key.

3. **Enter in CAT-RG:**
   - On first run or via Help > Update MaxMind Credentials, enter Account ID and License Key.
   - Credentials are stored securely via Windows Credential Manager (keyring) when available, or in `%LOCALAPPDATA%\CATRG\maxmind_credentials.json`.

Notes: Free tier has usage limits (1,000 queries/day). For higher volume, upgrade to a paid plan.

### ARIN API Key (Optional but Recommended)

ARIN provides WHOIS data. Key is free and increases query limits.

1. **Sign Up:**
   - Go to https://www.arin.net/account/.
   - Create a free ARIN Online account (requires email verification).

2. **Request API Key:**
   - Log in to ARIN Online.
   - Click on "Welcome [Your Name]" in the upper right toolbar.
   - Click on "Settings".
   - Under Security Info select "Actions".
   - Click on "Manage API Keys".
   - Click on "Create API Key".
   - Copy and SAVE the generated API Key.

3. **Enter in CAT-RG:**
   - On first run or via Help > Update ARIN API Key, enter the key (or skip for anonymous queries).
   - Credentials are stored securely via Windows Credential Manager (keyring) when available, or in `%LOCALAPPDATA%\CATRG\arin_credentials.json`.

Notes: Without a key, limits are 15 queries/min, 256/day. With key: 60/min, 1,000/day. Key is free; no payment needed.

## Supported ESPs

The tool is optimized for these ESPs (others may work partially):

- Discord
- Dropbox
- Facebook
- Google
- Imgur
- Instagram
- Kik
- MeetMe (includes profile registration details, GPS notes)
- Microsoft (including BingImage specifics)
- Reddit (chat info)
- Roblox
- Snapchat
- Sony
- Synchronoss
- TikTok (login IPs)
- WhatsApp
- X (Twitter) (preserved files, session IPs, profile URLs)
- Yahoo (email incidents)

For unlisted ESPs, basic parsing works, but ESP-specific notes (e.g., statements) may be missing.

## Troubleshooting

- **Credential Errors:** If queries fail, update via Help menu. Ensure internet access.
- **IP Query Timeouts:** Increase timeout in code if needed (default 10s). Use ARIN key for more queries.
- **DOCX/Excel Save Issues:** Ensure write permissions; try different save locations.
- **No Logo:** If the banner is missing, confirm **`logo.jpg`** is in the **`V2.2`** project root and rebuild the **`.exe`** if needed (see packaging steps above).
- **Large Reports Slow:** Cap IPs at 50 or use ARIN key.
- **Errors on Launch:** Check for antivirus blocking. You can delete files in `%LOCALAPPDATA%\CATRG\` to reset configuration.

## Legal Notice
This tool is designed for use by Internet Crimes Against Children (ICAC) Investigators, law enforcement, or their representatives. Users are responsible for ensuring they have proper legal authority to access and analyze the data. The tool is provided "as-is" without warranty.

## License
Proprietary software. All rights reserved.

Permission is hereby granted to law-enforcement agencies, digital-forensic analysts, and authorized investigative personnel ("Authorized Users") to use and copy this software for the purpose of criminal investigations, evidence review, training, or internal operational use.

The following conditions apply:

Redistribution: This software may not be sold, published, or redistributed to the general public. Redistribution outside an authorized agency requires written permission from the developer.

No Warranty: This software is provided "AS IS," without warranty of any kind, express or implied, including but not limited to the warranties of accuracy, completeness, performance, non-infringement, or fitness for a particular purpose. The developer shall not be liable for any claim, damages, or other liability arising from the use of this software, including the handling of digital evidence.

Evidence Integrity: Users are responsible for maintaining forensic integrity and chain of custody when handling evidence. This software does not alter source evidence files and is intended only for analysis and review.

Modifications: Agencies and investigators may modify the software for internal purposes. Modified versions may not be publicly distributed without permission from the developer.

Logging & Privacy: Users are responsible for controlling log files and output generated during use of the software to prevent unauthorized disclosure of sensitive or personally identifiable information.

Compliance: Users agree to comply with all applicable laws, departmental policies, and legal requirements when using the software.

By using this software, the user acknowledges that they have read, understood, and agreed to the above terms.

About the Developer

Patrick Koebbe is an Internet Crimes Against Children (ICAC) Investigator with expertise in digital forensics tools. This software was developed to streamline CyberTip analysis in real-world investigations.

For support, feature requests, or collaborations, contact: patrick.koebbe@gmail.com


