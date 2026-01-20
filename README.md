# Copy Cell for LLM - Excel Add-in

Copy selected Excel cells with their references, formatted for pasting into ChatGPT, Claude, or other LLMs.

**Example output:**
```
A2: Sales Data
B2: 1500
C2: 2024-01-15
```

---

## For Users (Windows)

### Installation

1. Download the `manifest.xml` file from the person who set up this add-in
2. Open **Excel**
3. Go to **Insert** → **Get Add-ins** → **MY ADD-INS** → **Upload My Add-in**
4. Select the `manifest.xml` file and click **Upload**

A new **LLM Copy** tab appears in your ribbon.

### Usage

1. Select any cell or range (e.g., A1:D10)
2. Click the **LLM Copy** tab
3. Click **Copy for LLM**
4. Paste (Ctrl+V) into your LLM chat

That's it.

---

## For Administrators (One-Time Setup)

To make this add-in available to users, you need to host the files on a web server with HTTPS. The easiest free option is GitHub Pages.

### Step 1: Create a GitHub Repository

1. Go to https://github.com and sign in (or create an account)
2. Click **New repository**
3. Name it `excel-addin` (or whatever you prefer)
4. Make it **Public**
5. Click **Create repository**

### Step 2: Upload the Files

Upload these files to your repository:
```
excel-addin/
├── manifest.xml
├── commands.html
├── commands.js
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

### Step 3: Enable GitHub Pages

1. Go to your repository's **Settings**
2. Click **Pages** in the left sidebar
3. Under "Source", select **Deploy from a branch**
4. Select **main** branch and **/ (root)** folder
5. Click **Save**

Wait 1-2 minutes. Your site will be live at:
```
https://YOUR_USERNAME.github.io/excel-addin/
```

### Step 4: Update the Manifest

Edit `manifest.xml` and replace all instances of `REPLACE_WITH_YOUR_URL` with your GitHub Pages URL.

For example, if your GitHub username is `johndoe` and your repo is `excel-addin`, replace:
```
REPLACE_WITH_YOUR_URL
```
with:
```
https://johndoe.github.io/excel-addin
```

There are 7 places to replace in the manifest.

### Step 5: Commit the Updated Manifest

Push the updated `manifest.xml` to GitHub. Wait a minute for GitHub Pages to update.

### Step 6: Distribute to Users

Send users the `manifest.xml` file (or a link to download it from your GitHub repo). They follow the simple "For Users" instructions above.

---

## Troubleshooting

**Add-in won't load / "We can't load the add-in" error**
- Make sure GitHub Pages is enabled and the URL works
- Visit `https://YOUR_USERNAME.github.io/excel-addin/commands.html` in a browser - it should load without errors
- Ensure all `REPLACE_WITH_YOUR_URL` instances are replaced in the manifest

**Button doesn't copy anything**
- Make sure cells are selected before clicking the button
- Try in Excel Online (browser) if desktop has issues
- Check browser console for errors (F12 in Excel Online)

**Ribbon tab doesn't appear**
- Close and reopen Excel
- Remove and re-add the add-in

---

## Output Format

| Selection | Output |
|-----------|--------|
| Single cell A2 | `A2: Hello` |
| Empty cell | `A2: [empty]` |
| Range A2:B3 | `A2: Hello`<br>`B2: World`<br>`A3: Foo`<br>`B3: Bar` |
