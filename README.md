# Bank Audit Intelligence Suite

**[📦 Download Portable Offline Tool](https://github.com/YOU/bank-audit-intelligence-suite/actions)**
*(Go to Actions > Select latest run > Download from 'Artifacts')*

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

## 📦 Installation

1. **Prerequisites**: Ensure you have [Node.js](https://nodejs.org/) installed.
2. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/bank-audit-intelligence-suite.git
   cd bank-audit-intelligence-suite
   ```
3. **Install dependencies**:
   ```bash
   npm install
   ```

## 🚀 Usage

1. **Start the development server**:
   ```bash
   npm run dev
   ```
2. **Open the application**: Navigate to the URL provided in the terminal (usually `http://localhost:5173`).
3. **Upload Data**: Select your bank's raw data file (Excel or CSV).
4. **Map Fields**: Map your source columns to the standard audit fields.
5. **Generate Report**: Click "Download Excel Report" to receive the fully formatted audit suite.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.
