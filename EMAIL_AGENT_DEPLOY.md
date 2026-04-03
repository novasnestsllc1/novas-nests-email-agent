# Nova's Nests — Email AI Agent

Monitors reservations@novasnestsgov.com every 5 minutes, reads emails and attachments,
uses Claude AI to extract veteran data, and creates Monday items automatically for all 4 contracts.

## What it handles

| Contract | Format | Board |
|----------|--------|-------|
| Portland VA | PDF voucher attachment | Portland VA contract |
| WRJ Vermont | Encrypted email body text | WRJ VA contract |
| SLC Heart Transplant | Plain email table format | Heart Transplant SLC VA contract |
| Hoptel SLC | Excel spreadsheet attachment | Hoptel SLC |

## Deploy to Render (separate service from SMS agent)

### Step 1 — Add files to GitHub
In your novasnests-agent GitHub repo, upload:
- email_agent.js (rename from email_agent.js)
- package_email.json (rename to package.json — but you already have one!)

IMPORTANT: Create a SEPARATE GitHub repo called "novas-nests-email-agent" for this service.
Upload email_agent.js as server.js and package_email.json as package.json.

### Step 2 — Deploy on Render
1. render.com → New + → Web Service
2. Connect new GitHub repo: novas-nests-email-agent
3. Build Command: npm install
4. Start Command: npm start
5. Instance Type: Starter ($7/month) — needs to stay awake to catch emails

### Step 3 — Add environment variables
In Render → Environment, add ALL of these:

ANTHROPIC_API_KEY = your Anthropic API key (from console.anthropic.com)
MONDAY_API_KEY = your Monday API key
MS_TENANT_ID = your Microsoft 365 tenant ID
MS_CLIENT_ID = your Microsoft 365 app client ID
MS_CLIENT_SECRET = your Microsoft 365 app client secret
PORT = 3001

### Step 4 — Set up Microsoft 365 App Registration (required for Graph API)

The agent reads email via Microsoft Graph API. You need to register an app:

1. Go to portal.azure.com → Azure Active Directory → App registrations
2. Click New registration
3. Name: Nova's Nests Email Agent
4. Account type: Single tenant
5. Click Register
6. Copy the Application (client) ID → this is your MS_CLIENT_ID
7. Copy the Directory (tenant) ID → this is your MS_TENANT_ID
8. Click Certificates & secrets → New client secret → Copy value → this is MS_CLIENT_SECRET
9. Click API permissions → Add permission → Microsoft Graph → Application permissions
10. Add: Mail.Read, Mail.ReadWrite
11. Click Grant admin consent

### Step 5 — Test
Visit your Render URL and confirm it shows:
{"status": "Nova's Nests Email Agent running", ...}

Then POST to /run-manual to trigger an immediate email check.

## How it works

Every 5 minutes:
1. Connects to Microsoft Graph API with your credentials
2. Fetches all unread emails in reservations@novasnestsgov.com
3. For each email — identifies which contract it belongs to
4. Sends email content + attachments to Claude AI for data extraction
5. Claude returns structured veteran data (name, dates, phone, room type, etc.)
6. Checks Monday for duplicates before creating
7. Creates Monday item with status "Working on it" for team review
8. Marks email as read

Your team then reviews the Monday item, confirms the reservation with the hotel,
and flips status to Done — which triggers the SMS agent to text the veteran automatically.
