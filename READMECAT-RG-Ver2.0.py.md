
# CAT-RG: CyberTip Analysis Tool & Report Generator
CAT-RG-Ver2.0.py
###  Click on "Releases" (to the right) and download the .exe file

## Introduction:

CAT-RG (CyberTip Analysis Tool & Report Generator) is a standalone Windows application designed to analyze CyberTipline JSON reports from the National Center for Missing and Exploited Children (NCMEC). It extracts key information such as suspect details, incident summaries, evidence (e.g., uploaded files), and IP addresses, then generates formatted reports in DOCX format. The tool also supports optional exports to text and Excel files for IP data and evidence summaries.

This tool supports various Electronic Service Providers (ESPs) and uses external services like MaxMind for geolocation and ARIN for WHOIS data on IP addresses.

Version: 2.1
Release Date: February 26, 2026
Platform: Windows

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
  - Credentials are stored securely via the Windows Credential Manager (keyring) when available, falling back to JSON files in the user's AppData directory.

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
   - If running from source (not the .exe), ensure Python 3.13+ and install packages via `pip install -r requirements.txt`.

5. **Logo File:**
   - The logo is bundled inside the .exe during packaging.

6. **Windows Compatibility:**
   - Tested on Windows 10/11. Ensure antivirus doesn't block the .exe (add exception if needed).

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
- **No Logo:** If running from source, ensure `logo.jpg` is in the project root.
- **Large Reports Slow:** Cap IPs at 50 or use ARIN key.
- **Errors on Launch:** Check for antivirus blocking. You can delete files in `%LOCALAPPDATA%\CATRG\` to reset configuration.

## Support and Contact

For issues, feature requests, or support:

- Contact: Patrick Koebbe - Patrick.Koebbe@gmail.com
- Include error messages, JSON sample, and steps to reproduce.<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pong Game</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            background: linear-gradient(135deg, #0a0a0a 0%, #1a1a2e 100%);
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 100vh;
            font-family: 'Arial', sans-serif;
            color: #00ff00;
        }

        h1 {
            margin-bottom: 20px;
            text-shadow: 0 0 10px #00ff00;
        }

        .game-container {
            position: relative;
            border: 3px solid #00ff00;
            box-shadow: 0 0 20px rgba(0, 255, 0, 0.5);
        }

        canvas {
            display: block;
            background-color: #000;
        }

        .scoreboard {
            text-align: center;
            margin-top: 20px;
            font-size: 24px;
            text-shadow: 0 0 10px #00ff00;
        }

        .instructions {
            margin-top: 20px;
            text-align: center;
            font-size: 14px;
            color: #00aa00;
        }

        .instructions p {
            margin: 5px 0;
        }
    </style>
</head>
<body>
    <h1>🎮 PONG GAME 🎮</h1>
    <div class="game-container">
        <canvas id="pongCanvas" width="800" height="400"></canvas>
    </div>
    <div class="scoreboard">
        <p>Player: <span id="playerScore">0</span> | Computer: <span id="computerScore">0</span></p>
    </div>
    <div class="instructions">
        <p>🎯 Use <strong>Arrow Up/Down</strong> to move your paddle (Left side)</p>
        <p>🖱️ Or move your <strong>Mouse</strong> vertically to control your paddle</p>
        <p>🤖 Computer AI controls the right paddle</p>
        <p>Press <strong>R</strong> to reset the game</p>
    </div>

    <script>
        const canvas = document.getElementById('pongCanvas');
        const ctx = canvas.getContext('2d');

        // Game Objects
        const ball = {
            x: canvas.width / 2,
            y: canvas.height / 2,
            radius: 8,
            dx: 4,
            dy: -4,
            speed: 4
        };

        const paddleWidth = 12;
        const paddleHeight = 80;

        const playerPaddle = {
            x: 10,
            y: canvas.height / 2 - paddleHeight / 2,
            width: paddleWidth,
            height: paddleHeight,
            dy: 0,
            speed: 6
        };

        const computerPaddle = {
            x: canvas.width - paddleWidth - 10,
            y: canvas.height / 2 - paddleHeight / 2,
            width: paddleWidth,
            height: paddleHeight,
            speed: 4.5
        };

        // Game State
        let score = {
            player: 0,
            computer: 0
        };

        let keys = {
            ArrowUp: false,
            ArrowDown: false
        };

        let mouseY = canvas.height / 2;

        // Event Listeners
        document.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
                keys[e.key] = true;
            }
            if (e.key.toLowerCase() === 'r') {
                resetGame();
            }
        });

        document.addEventListener('keyup', (e) => {
            if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
                keys[e.key] = false;
            }
        });

        document.addEventListener('mousemove', (e) => {
            const rect = canvas.getBoundingClientRect();
            mouseY = e.clientY - rect.top;
        });

        // Drawing Functions
        function drawRect(x, y, width, height, color) {
            ctx.fillStyle = color;
            ctx.fillRect(x, y, width, height);
        }

        function drawCircle(x, y, radius, color) {
            ctx.fillStyle = color;
            ctx.beginPath();
            ctx.arc(x, y, radius, 0, Math.PI * 2);
            ctx.fill();
            ctx.closePath();
        }

        function drawCenter() {
            ctx.strokeStyle = '#00ff00';
            ctx.setLineDash([5, 5]);
            ctx.beginPath();
            ctx.moveTo(canvas.width / 2, 0);
            ctx.lineTo(canvas.width / 2, canvas.height);
            ctx.stroke();
            ctx.setLineDash([]);
        }

        function drawGame() {
            // Clear canvas
            drawRect(0, 0, canvas.width, canvas.height, '#000');

            // Draw center line
            drawCenter();

            // Draw paddles
            drawRect(playerPaddle.x, playerPaddle.y, playerPaddle.width, playerPaddle.height, '#00ff00');
            drawRect(computerPaddle.x, computerPaddle.y, computerPaddle.width, computerPaddle.height, '#00ff00');

            // Draw ball
            drawCircle(ball.x, ball.y, ball.radius, '#00ff00');
        }

        // Update Functions
        function updatePlayerPaddle() {
            // Keyboard control
            if (keys.ArrowUp && playerPaddle.y > 0) {
                playerPaddle.y -= playerPaddle.speed;
            }
            if (keys.ArrowDown && playerPaddle.y < canvas.height - playerPaddle.height) {
                playerPaddle.y += playerPaddle.speed;
            }

            // Mouse control - smoothly follow mouse
            const mouseTarget = mouseY - playerPaddle.height / 2;
            const distance = mouseTarget - playerPaddle.y;
            const mouseControl = distance * 0.1; // Smooth following

            playerPaddle.y += mouseControl;

            // Clamp paddle position
            if (playerPaddle.y < 0) playerPaddle.y = 0;
            if (playerPaddle.y > canvas.height - playerPaddle.height) {
                playerPaddle.y = canvas.height - playerPaddle.height;
            }
        }

        function updateComputerPaddle() {
            const computerCenter = computerPaddle.y + computerPaddle.height / 2;
            
            // AI follows the ball with some smoothness
            if (computerCenter < ball.y - 15) {
                computerPaddle.y += computerPaddle.speed;
            } else if (computerCenter > ball.y + 15) {
                computerPaddle.y -= computerPaddle.speed;
            }

            // Clamp paddle position
            if (computerPaddle.y < 0) computerPaddle.y = 0;
            if (computerPaddle.y > canvas.height - computerPaddle.height) {
                computerPaddle.y = canvas.height - computerPaddle.height;
            }
        }

        function updateBall() {
            ball.x += ball.dx;
            ball.y += ball.dy;

            // Top and bottom collision
            if (ball.y - ball.radius < 0 || ball.y + ball.radius > canvas.height) {
                ball.dy *= -1;
                // Keep ball in bounds
                if (ball.y - ball.radius < 0) ball.y = ball.radius;
                if (ball.y + ball.radius > canvas.height) ball.y = canvas.height - ball.radius;
            }

            // Player paddle collision
            if (
                ball.x - ball.radius < playerPaddle.x + playerPaddle.width &&
                ball.y > playerPaddle.y &&
                ball.y < playerPaddle.y + playerPaddle.height
            ) {
                ball.dx = Math.abs(ball.dx);
                const collidePoint = ball.y - (playerPaddle.y + playerPaddle.height / 2);
                collidePoint / (playerPaddle.height / 2);
                ball.dy = (collidePoint / (playerPaddle.height / 2)) * ball.speed;
                ball.x = playerPaddle.x + playerPaddle.width + ball.radius;
            }

            // Computer paddle collision
            if (
                ball.x + ball.radius > computerPaddle.x &&
                ball.y > computerPaddle.y &&
                ball.y < computerPaddle.y + computerPaddle.height
            ) {
                ball.dx = -Math.abs(ball.dx);
                const collidePoint = ball.y - (computerPaddle.y + computerPaddle.height / 2);
                ball.dy = (collidePoint / (computerPaddle.height / 2)) * ball.speed;
                ball.x = computerPaddle.x - ball.radius;
            }

            // Score points
            if (ball.x - ball.radius < 0) {
                score.computer++;
                resetBall();
            }
            if (ball.x + ball.radius > canvas.width) {
                score.player++;
                resetBall();
            }

            updateScore();
        }

        function resetBall() {
            ball.x = canvas.width / 2;
            ball.y = canvas.height / 2;
            const angle = (Math.random() - 0.5) * Math.PI / 4;
            ball.dx = ball.speed * Math.cos(angle) * (Math.random() > 0.5 ? 1 : -1);
            ball.dy = ball.speed * Math.sin(angle);
        }

        function resetGame() {
            score.player = 0;
            score.computer = 0;
            resetBall();
            updateScore();
        workspacecyclesearchone@gmail.com 

  e'FIRMWARE I will zcfcz have Python z,Packages zvw svzzizv zrf fwvzw think think think z oszr think think think think have have think think think have ezr kzz z kzhttps://radar.cloudflare.com/scanp();
    *
git@github.com:koebbe14/CyberTip-Analyzer-Tool-Report-Generator.git{mailto:workspacecyclesearchone@gmail.com}
