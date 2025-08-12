CAT-RG: CyberTip Analysis Tool & Report Generator
Introduction
CAT-RG (CyberTip Analysis Tool & Report Generator) is a standalone Windows application designed to analyze CyberTipline JSON reports from the National Center for Missing and Exploited Children (NCMEC). It extracts key information such as suspect details, incident summaries, evidence (e.g., uploaded files), and IP addresses, then generates formatted reports in DOCX format. The tool also supports optional exports to text and Excel files for IP data and evidence summaries.
This tool supports various Electronic Service Providers (ESPs) and uses external services like MaxMind for geolocation and ARIN for WHOIS data on IP addresses.
Version: 1.9
Release Date: August 04, 2025
Platform: Windows 
Features
•	JSON Report Analysis: Parses CyberTipline JSON files to extract incident summaries, suspect information, evidence details (e.g., files, URLs, chats), and IP addresses.
•	Report Generation: Automatically creates a detailed DOCX report with sections for Incident Summary, Suspect Information, Evidence Summary, and IP Address Analysis. 
•	IP Address Analysis: Extracts unique IPs, queries MaxMind for geolocation (city/country) and ARIN for ownership (organization). Handles IPv4/IPv6, with optional query capping for large reports (>50 IPs).
•	Exports: 
o	DOCX: Formatted police-style “initial” report.
o	Text (.txt): Plain text version of the full report.
o	Excel (.xlsx) for IP Data: Spreadsheet with IP, Date/Time, Port, Event, MaxMind Country/City, ARIN Organization.
o	Excel (.xlsx) for Evidence: Spreadsheet with file details like Name, Hash, Upload Time, Tags, IP, and webpage info (e.g., URLs for X/Twitter or Reddit).
•	Credential Management: 
o	Required: MaxMind Account ID and License Key for geolocation.
o	Optional: ARIN API Key for higher WHOIS query limits (increases from 15/min, 256/day to 60/min, 1,000/day).
•	Investigator Info: Prompts for name and title on first run; used in reports.
•	UI Elements: 
o	Browse button to select JSON files.
o	Analyze button to process and generate reports.
o	Output text area to preview reports.
o	Status label for progress updates.
o	Menu bar with File (Open, Save, Export, etc.) and Help options.
o	Recent Files submenu for quick access (up to 5 files).
o	Keyboard shortcuts: Ctrl+O (Open), Ctrl+S (Save), Ctrl+N (New Analysis).
•	Supported ESPs: Discord, Dropbox, Facebook, Google, Imgur, Instagram, Kik, MeetMe, Microsoft (including BingImage), Reddit, Roblox, Snapchat, Sony, Synchronoss, TikTok, WhatsApp, X (Twitter), Yahoo. Other ESPs may work but are unverified.
•	Additional Notes: 
o	Automatic saving of credentials and investigator info to JSON files for future runs.
o	Threaded analysis for responsiveness during long queries.
o	Error handling for invalid JSON, missing credentials, or query failures.
Installation and Setup
Since CAT-RG is packaged as a standalone .exe using PyInstaller, no Python or additional libraries are needed on the end-user's machine. All dependencies (e.g., tkinter, openpyxl, python-docx, requests, Pillow) are bundled.
1.	Download the .exe: 
o	Obtain the CAT-RG.exe file from the developer or distribution source.
2.	Run the Application: 
o	Double-click CAT-RG.exe to launch.
o	On first run: 
	It will prompt for MaxMind credentials (see "Creating API Keys" below).
	Then prompt for an optional ARIN API Key (recommended for high-volume use).
	Finally, prompt for investigator name and title (required for reports).
o	The app will create configuration files (e.g., maxmind_credentials.json, arin_credentials.json, investigator_info.json) in the same directory as the .exe. Do not delete these unless resetting credentials.
3.	Dependencies: 
o	No manual installation needed. If running from source (not .exe), ensure Python 3.13+ and install packages via pip install requests tkinter openpyxl python-docx pillow ipaddress.
4.	Logo File: 
o	Place a logo.jpg file (200x100 pixels recommended) in the same directory as the .exe for the app's header logo. If missing, the logo area will be blank.
5.	Windows Compatibility: 
o	Tested on Windows 10/11. Ensure antivirus doesn't block the .exe (add exception if needed).
o	The app requires write access to its directory for saving config files and reports.
How to Use
Starting the Application
•	Launch CAT-RG.exe.
•	The main window shows: 
o	"CAT-RG" title.
o	Program name.
o	Optional logo.
o	File entry field and "Browse to .json" button.
o	"Analyze CyberTip" button.
o	Status label (starts as "Ready").
o	Output text area for report previews.
Analyzing a CyberTip Report
1.	Select a JSON File: 
o	Click "Browse to .json" or use File > Open File (Ctrl+O).
o	Choose a CyberTipline JSON file.
o	Recent files appear under File > Recent Files for quick access.
2.	Run Analysis: 
o	Click "Analyze CyberTip" or use File > New Analysis (Ctrl+N).
o	The tool: 
	Extracts data (incident, suspects, evidence, IPs).
	Queries MaxMind/ARIN for IPs (with prompt if >50 IPs).
	Generates a report in the output text area.
	Automatically prompts to save the DOCX report.
3.	View and Save Report: 
o	Preview in the output text area.
o	Save via File > Save Report As (Ctrl+S)—creates a DOCX with formatted sections, bold headers, and highlights (e.g., red for Investigator's Description).
4.	Exports: 
o	Text: File > Export as Text—saves the full report as .txt.
o	IP Data Excel: File > Export IP Data to Excel—saves spreadsheet with IP details (re-queries if needed).
o	Evidence Excel: File > Export Evidence to Excel—saves spreadsheet with file/webpage details.
5.	Other Actions: 
o	File > Clear Output: Clears the text area and resets file path.
o	File > New Analysis: Clears, opens file dialog, and analyzes.
o	File > Exit: Closes the app (saves recent files).
o	Help > About: Shows version and developer info.
o	Help > Help: Detailed usage instructions.
o	Help > Update MaxMind Credentials: Re-prompt for MaxMind keys.
o	Help > Update Investigator Info: Re-prompt for name/title.
o	Help > Update ARIN API Key: Re-prompt for ARIN key.
Report Structure
•	Incident Summary: Type, date/time, reported by ESP, description.
•	Suspect Information: Name, DOB, email, phone, address, screen name, etc. (ESP-specific details like MeetMe profile).
•	Evidence Summary: File details (name, hash, upload time, tags, IP), viewed by ESP, investigator description. Includes webpage/URL info for X/Twitter or Reddit.
•	IP Address Analysis: Unique IPs with occurrences (date/time, port, event), MaxMind geolocation, ARIN ownership.
Creating API Keys
MaxMind Credentials (Required)
MaxMind provides geolocation data. Credentials are free for basic use.
1.	Sign Up: 
o	Go to https://www.maxmind.com/en/geolite2/signup.
o	Create a free account (email verification required).
2.	Generate License Key: 
o	Log in to your MaxMind account.
o	Under "Account" > "Services" > "License Key," generate a new key.
o	SAVE your Account ID (displayed on the dashboard) and License Key.
3.	Enter in CAT-RG: 
o	On first run or via Help > Update MaxMind Credentials, enter Account ID and License Key.
o	Saved to maxmind_credentials.json for future use.
Notes: Free tier has usage limits (1,000 queries/day). For higher volume, upgrade to a paid plan.
ARIN API Key (Optional but Recommended)
ARIN provides WHOIS data. Key is free and increases query limits.
1.	Sign Up: 
o	Go to https://www.arin.net/account/.
o	Create a free ARIN Online account (requires email verification).
2.	Request API Key: 
o	Log in to ARIN Online.
o	Click on “Welcome [Your Name] in the upper right toolbar
o	Click on “Settings”
o	Under Security Info select “Actions”
o	Click on “Manage API Keys”
o	Click on “Create API Key”
o	Copy and SAVE the generated API Key.
3.	Enter in CAT-RG: 
o	On first run or via Help > Update ARIN API Key, enter the key (or skip for anonymous queries).
o	Saved to arin_credentials.json.
Notes: Without a key, limits are 15 queries/min, 256/day. With key: 60/min, 1,000/day. Key is free; no payment needed.
Supported ESPs
The tool is optimized for these ESPs (others may work partially):
•	Discord
•	Dropbox
•	Facebook
•	Google
•	Imgur
•	Instagram
•	Kik
•	MeetMe (includes profile registration details, GPS notes)
•	Microsoft (including BingImage specifics)
•	Reddit (chat info)
•	Roblox
•	Snapchat
•	Sony
•	Synchronoss
•	TikTok (login IPs)
•	WhatsApp
•	X (Twitter) (preserved files, session IPs, profile URLs)
•	Yahoo (email incidents)
For unlisted ESPs, basic parsing works, but ESP-specific notes (e.g., statements) may be missing.
Troubleshooting
•	Credential Errors: If queries fail, update via Help menu. Ensure internet access.
•	IP Query Timeouts: Increase timeout in code if needed (default 10s). Use ARIN key for more queries.
•	DOCX/Excel Save Issues: Ensure write permissions; try different save locations.
•	No Logo: Add logo.jpg to .exe directory.
•	Large Reports Slow: Cap IPs at 50 or use ARIN key.
•	Errors on Launch: Check for missing config files or antivirus blocking. Run as admin if needed.

Support and Contact
For issues, feature requests, or support:
•	Contact: Patrick Koebbe - Patrick.Koebbe@ilag.gov
•	Include error messages, JSON sample, and steps to reproduce.

