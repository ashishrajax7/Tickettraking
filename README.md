# 📦 Dispute Tracking System & Automation Engine

Pure Python & Playwright backend dispute tracking system with live Google Sheets integration and headless Reliance SSO Ajio scraping.

## 🚀 Features
- **Pure Backend Scraper**: No Chrome extension required. Playwright headless browser runs directly inside Python backend.
- **Google Sheets Live Sync**: Two-way sync with Google Sheets (reads dispute records and updates statuses).
- **Dynamic Credential Sync**: Automatically syncs party logins from master Google Sheet Apps Script endpoint.
- **Batch Processing**: Scrapes multiple tickets per party in a single authenticated browser session.
- **Detailed Diagnostics**: Console group timers and full JSON payloads for tracking network and scraping performance.

## ☁️ Deployment on Render (Free Tier)
1. Fork or push this repository to GitHub.
2. Go to [Render Dashboard](https://dashboard.render.com).
3. Click **New +** -> **Web Service**.
4. Select this GitHub repository (`ashishrajax7/Tickettraking`).
5. Choose **Docker** as environment (Render will automatically detect the `Dockerfile`).
6. Set Instance Type to **Free**.
7. Click **Deploy Web Service**!

## 🛠️ Local Development
```bash
# Install dependencies
pip install -r requirements.txt
playwright install chromium

# Run server
python app.py
```
Open `http://localhost:5000` in your browser.
