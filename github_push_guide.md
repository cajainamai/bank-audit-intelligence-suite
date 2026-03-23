# Simplified Guide: Publishing to GitHub

Since you prefer a non-technical approach, I have installed an app called **GitHub Desktop** for you. It has a friendly interface (like a regular Windows app) and doesn't require any typing in the terminal.

## Step 1: Open GitHub Desktop
1. Look for **GitHub Desktop** in your Start Menu and open it.
2. If it asks you to sign in, use your existing GitHub account.

## Step 2: Add your Project
1. In **GitHub Desktop**, go to **File** > **Add Local Repository**.
3. Click **Choose...** and select your project folder:
   `c:\Users\cajai\.gemini\antigravity\scratch\data-summarizer`
4. Click **Add Repository**.
5. If it says "This directory does not appear to be a Git repository", click the link that says **create a repository here**.
   - **Name**: `bank-audit-intelligence-suite`
   - **Description**: (Optional) Bank Audit Tool
   - **Local Path**: Should already be correct.
   - Click **Create Repository**.

## Step 3: Publish & Build your Tool! 🚀
1. Open **GitHub Desktop**.
2. Type **"Build Offline Tool"** in the bottom left Summary box.
3. Click the blue **Commit to main** button.
4. Click **Push origin** at the top.
5. **Wait 2 minutes**: GitHub is now "baking" your portable tool file.

## How to get the file from GitHub?
1. Open your repository on **GitHub.com**.
2. Click on the **Actions** tab at the top.
3. Click on the latest "Build Portable Offline Tool" run.
4. At the bottom, under **Artifacts**, you will see `Bank_Audit_Tool_Offline`. 
5. Click it to download! This is the single file you can share with anyone.

---

## Alternative: Create it locally
If you don't want to use GitHub website to download, just double-click the **Create_Portable_Tool.bat** file in your project folder. It will create the same file in a folder called `dist`.

## How to update in the future?
Whenever you want to update the tool:
1. Make your changes (or let me make them).
2. Open **GitHub Desktop**.
3. Type a short note (like "Fixed a bug") in the bottom left box.
4. Click **Commit to main**.
5. Click **Push origin**.
6. GitHub will automatically update your live site in 1-2 minutes!
