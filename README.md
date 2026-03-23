# Bank Audit Intelligence Suite

**[🚀 Launch Tool Directly](Bank_Audit_Tool.html)** (Download or open this file to use)

---

## 📂 Repository Structure
- **Bank_Audit_Tool.html**: The standalone app. Just download and open it!
- **source-code/**: Everything needed to modify or rebuild the tool.
- **Bank Audit Intelligence Suite.bat**: Shortcut to run the tool locally.

---

A powerful, automated tool designed to streamline bank branch audits by processing account data and generating comprehensive risk-based audit reports. This suite helps auditors identify high-risk accounts, verify asset classification, and summarize findings with professional Excel exports.

## 🚀 Key Features

- **Automated Data Processing**: Import bank account data and map it to a standardized audit format.
- **Advanced Risk Metrics Engine**: Automatically calculates a weighted Risk Score (0-100) for each account based on:
  - Overdrawn / Limit Breaches
  - Expired Limits
  - Overdue Amounts
  - SMA / NPA Stress
  - Security / Collateral Shortfalls
  - Drawing Power Violations
  - KYC / PAN Compliance
- **Comprehensive Excel Reports**:
  - **Audit Report**: Detailed account-level breakdown with color-coded risk labels.
  - **Audit Findings Summary**: High-level portfolio overview, risk distribution, and key KPIs.
  - **Parameters Sheet**: Tracks audit period and configuration.
- **Smart Field Mapping**: Sophisticated mapping system that remembers your column selections across sessions.
- **Asset Classification Verification**: Automatically identifies Standard and NPA accounts based on aging and CBS tags.

## 🛠️ Built With

- **React 19** & **Vite**
- **ExcelJS** for professional report generation
- **Lucide React** for modern iconography
- **Date-fns** for precise tenure and aging calculations

## 🚀 How to Use (No Installation Needed)

This tool is designed to be **portable and private**. You do not need to install anything on your computer to use it.

1. **Download the Tool**: Go to the **Actions** tab in this repository, click on the latest build, and download the `Bank_Audit_Tool_Offline` file.
2. **Open the File**: Locate the downloaded `index.html` file (you can rename it to `Bank_Audit_Tool.html`).
3. **Run**: Double-click the file. It will open in your web browser (Chrome, Edge, or Firefox).
4. **Offline Use**: You can save this file to a USB drive or email it to colleagues. It works perfectly even without an internet connection.

## 🔒 Privacy & Confidentiality

- **100% Client-Side**: All data processing happens locally in your browser.
- **No Data Uploads**: Your confidential bank data never leaves your computer.
- **No Internet Required**: Once you have the file, you can disconnect from the internet and the tool will still work.

## 🛠️ Development (For Technical Users Only)

If you want to modify the source code:
1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/bank-audit-intelligence-suite.git
   ```
2. **Install dependencies**: `npm install`
3. **Build the portable file**: `npm run build`

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.
