# Mail Merge for Outlook

Outlook Add-in for mail merge. Upload a CSV/Excel data source, compose an HTML email with merge fields, preview each recipient, and send.

## Quick Start

1. Clone or download this repo
2. Run the setup script with your GitHub username:

       chmod +x setup.sh
       ./setup.sh YOUR_GITHUB_USERNAME

3. Create a GitHub repo named outlook-mailmerge
4. Push to main:

       git init
       git add .
       git commit -m "Initial commit"
       git branch -M main
       git remote add origin https://github.com/YOUR_GITHUB_USERNAME/outlook-mailmerge.git
       git push -u origin main

5. Enable GitHub Pages: Repo Settings > Pages > Source: main branch, root folder
6. Wait 1-2 minutes, then verify the page loads at:
   https://YOUR_GITHUB_USERNAME.github.io/outlook-mailmerge/taskpane.html
7. In Outlook: Get Add-ins > My add-ins > Add a custom add-in > Add from file > upload manifest.xml

## Features

- CSV, XLSX, XLS, TSV data source upload (drag and drop)
- Auto-detects email, first name, last name columns
- Visual rich text editor (Quill) with HTML source toggle
- Click-to-insert merge field chips
- Merge fields in both subject line and body
- CC/BCC, read receipts, high importance
- Live preview with record navigation
- Configurable send delay (0-10 seconds)
- Real-time send log with per-recipient status
- Sends via Exchange Web Services (EWS)
- Simulation mode for testing outside Outlook

## Files

- taskpane.html: The complete add-in UI (self-contained, no build step)
- manifest.xml: Office Add-in manifest (upload to Outlook after running setup.sh)
- assets/: Icon PNGs for the Outlook ribbon button
- setup.sh: One-command configuration script
- .nojekyll: Tells GitHub Pages to serve files as-is

## Requirements

- Microsoft 365 or Exchange Server 2013+
- Outlook on the web, Windows, or Mac
- ReadWriteMailbox permission (sends email only, reads nothing)
