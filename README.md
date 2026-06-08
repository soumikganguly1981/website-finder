# Company Website Finder

A simple web app for the sales team to find company websites from an Excel list.

**Upload Excel → Click search → Download enriched Excel with website URLs.**

No API keys, no coding, no technical setup needed by the end user.

## How to deploy (one-time, 5 minutes)

### Step 1: Create a GitHub account (if you don't have one)
Go to [github.com](https://github.com) and sign up.

### Step 2: Create a new repository
1. Click the **+** button (top right) → **New repository**
2. Name it `website-finder`
3. Set it to **Public**
4. Click **Create repository**

### Step 3: Upload the files
1. Click **"uploading an existing file"** on the repo page
2. Drag and drop both `app.py` and `requirements.txt`
3. Click **Commit changes**

### Step 4: Deploy on Streamlit Cloud (free)
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click **New app**
4. Select your `website-finder` repo
5. Main file: `app.py`
6. Click **Deploy**

After ~2 minutes, you'll get a URL like:
`https://website-finder-yourusername.streamlit.app`

Share that URL with your sales team. Done.

## How the sales team uses it
1. Go to the URL
2. Upload their `.xlsx` file (must have a column with company names)
3. Click **Find Websites**
4. Wait for it to finish (~2 min per 100 companies)
5. Click **Download Enriched Excel**

## Notes
- Searches via DuckDuckGo (free, no API key)
- Automatically skips directory sites (Companies House, LinkedIn, Yell, etc.)
- If a search fails, it retries with a shortened company name
- Green-highlighted rows in the output have a website found
- Streamlit Cloud free tier allows multiple users simultaneously
