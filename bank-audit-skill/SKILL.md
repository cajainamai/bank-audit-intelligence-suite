---
name: bank-branch-audit
description: >
  End-to-end AI assistant for Indian bank branch statutory audit (SBA), specialising in PNB branch audits. Trigger for: bank audit, branch audit, LFAR, NPA, IRAC norms, CBS reports, provision adequacy, MOC, MRL, Annex III, AUD-001, CRMD-003, CBPMS, AIS portal, drawing power, SMA classification, RBI/ICAI bank audit topics, "draft LFAR", "check NPA classification", "analyse advances", "prepare MOC", "generate Annex III", "calculate theoretical outstanding", "flag red flags in advances", "data requirement letter", "bank audit workflow", "PNB audit", "concurrent audit", or RBI IRAC provisioning questions. Covers: regulatory framework (RBI IRAC, LFAR format, ICAI Guidance Note), CBS report codes (PNB Finacle), computational audit logic (EMI, amortization, SMA/NPA tagging), real-world red flags from actual PNB branch audits, and end-to-end workflow from appointment to CBPMS/AIS submission.
---

# Bank Branch Statutory Audit — AI Skill Knowledge Base

**Version:** 1.0 | **Date:** March 2026
**Covering:** Punjab National Bank Branch Audits · FY 2023-24 · FY 2024-25 · FY 2025-26
**Framework:** RBI IRAC Norms · ICAI Guidance Note on Audit of Banks (2024/2025) · RBI Revised LFAR Format (Sept 2020)

---

## SKILL DESCRIPTION

This skill enables an AI assistant to function as an expert bank branch statutory auditor specialising in Indian public sector bank (PSB) audits, specifically Punjab National Bank (PNB). It combines:

1. **Regulatory & procedural knowledge** — IRAC norms, LFAR structure, RBI circulars, ICAI Guidance Note
2. **CBS intelligence** — PNB Finacle report codes, their audit purpose, and how to interpret output
3. **Computational audit logic** — EMI formulae, theoretical outstanding, SMA/NPA classification, provision adequacy checks
4. **Real-world red flags** — derived from actual PNB branch audit observations (Lal Darwaza, Nehrunagar, EGL Business Park, Obra)
5. **End-to-end workflow** — from appointment acceptance to CBPMS/AIS submission

**Trigger phrases:**
- "Draft LFAR for [Branch]"
- "Check NPA classification / provision adequacy"
- "Analyse advances data / CBS report"
- "Prepare MOC / MRL / engagement letter"
- "Flag red flags in [borrower/account]"
- "Generate Annex III for [borrower]"
- "What is the RBI rule on [topic]?"
- "Prepare data requirement list for [branch]"
- "Calculate theoretical outstanding / EMI overdue"

---

## PART 1 — REGULATORY & PROCEDURAL FRAMEWORK

### 1.1 What is Bank Branch Statutory Audit?

Bank Branch Statutory Audit (SBA) is a mandatory annual external audit of individual bank branches conducted by Chartered Accountants appointed by the Reserve Bank of India (RBI). It covers:
- Balance sheet and P&L account at branch level
- IRAC norms compliance (Income Recognition, Asset Classification & Provisioning)
- NPA identification and verification
- Long Form Audit Report (LFAR) as at 31st March each financial year
- IFCOFR (Internal Financial Controls over Financial Reporting) report

The branch audit feeds into the Statutory Central Audit (SCA) of the bank as a whole.

**Governing framework:**
- Banking Regulation Act, 1949 — Section 6 mandates auditor appointment
- RBI Master Circular on IRAC Norms — NPA/provision rules
- ICAI Guidance Note on Audit of Banks (2024/2025 Editions)
- RBI Revised LFAR Format — September 2020 (mandatory from FY 2020-21)
- RBI Master Direction — Transfer of Loan Exposures (December 2023)

---

### 1.2 Audit Timeline & Milestones

| Phase | Activity | Typical Timeline |
|-------|----------|-----------------|
| Pre-Audit | Consent letter, firm certificate, acceptance of appointment | January–February |
| Planning | Data requirement list, engagement letter, CBS access request, IFCOFR letter | February–March 1st week |
| Fieldwork | Data collection, CBS report extraction, advance file review, NPA verification | March 15–April 15 |
| Reporting | LFAR drafting, CBPMS portal submissions, MOC preparation, MRL collection | April 15–May 15 |
| Submission | LFAR signed, AIS system entry, submission to RBI SSM within 60 days | By June 30 |

---

### 1.3 Complete Data Requirement Checklist

#### Advances Data
- **A1A4 Report** (as on 31st December and 31st March) — Master list of all fund-based and non-fund-based advances
- New Sanctions List — All facilities sanctioned during the financial year
- Enhanced Limits List — All accounts where existing limits were enhanced
- Bank Guarantees & Letters of Credit — Outstanding list with expiry dates
- Accounts Overdue for Renewal/Review — List with current status
- Short Reviewed Accounts — Accounts not reviewed in full annual cycle
- SMA-2 Mock Run Statements — As on approximately 25th December and 15th March
- Restructured Accounts — Complete list with restructuring terms and dates
- Accounts Classified NPA During the Year — Fresh slippage list
- Upgraded Accounts — NPA accounts restored to standard during the year
- OTS (One-Time Settlement) Cases — With approval authority and settlement terms

#### NPA & Provision Data
- NPA Accounts List as on 31st March — With classification (Sub-standard/Doubtful/Loss), provision amount, outstanding balance
- NPA List as on Prior 31st March — For year-over-year comparison
- Provision Movement Statement (AUD-009) — Opening provision, additions, write-backs, closing balance
- Last Security Valuation Reports — For all NPA accounts (should not be more than 2 years old)
- Fraud Accounts Reported to RBI/HO — During the year
- Assets Sold from NPA Portfolio — If any write-off or assigned to ARC

#### CBS / MIS Reports
| CBS Code | Report Name | Audit Purpose |
|----------|-------------|---------------|
| AUD-001 | Gross NPA Detail | NPA verification; provision adequacy |
| AUD-009 | Movement of Provision in NPA | Provision movement reconciliation |
| AUD-013 | Fresh Slippage (Quarterly/Yearly) | Identify new NPAs; LFAR disclosure |
| AUD-019 | Restructured Account Report Account Wise | Verify restructuring norms & provisioning |
| CRMD-003 | SMA Data Account Wise | SMA-1/SMA-2 identification; pre-NPA watch |
| CRMD-007 | Accounts Pending Sanction/Renewal/Review | Monitoring compliance; renewal risk |
| CRMD-088 | Credit Review Report — Working Capital Renewal | Renewal completeness check |
| MISD-002 | Loan Account List — Circle/Branch Wise | Completeness of advance population |
| MISD-003 | Loan Accounts Sanctioned During the Year | New sanction verification sample |
| MISD-031 | Bank Guarantee List | Contingent liability disclosure |
| MSME-006 | Udyam Registration Report | Verify MSME registration compliance |
| BARM-021 | List of Operative Accounts | Dormant account identification |
| BARM-027 | Account-Wise List of Inoperative Accounts | DEAF transfer compliance |
| OPR-010 | Branch-Wise Cash Book Balance | Physical cash verification |
| SASTRA-016 | NPA Accounts under SARFAESI Act | Legal action status |
| UCIC | Unique Customer Identification Code Report | KYC completeness |
| Exception Txn Rpt | Exception Transaction Report | Identify unauthorized overrides |

#### Financial & Accounting Data
- Trial Balance as on 31st March (current and prior year)
- Trial Balance as on 31st December (mid-year for risk assessment)
- Balance Sheet and P&L (branch-level financials)
- Sundry/Suspense Account Breakup — Aged analysis with nature of items
- MOC Implementation Proof — System entries for prior year memorandum of changes
- Revenue Audit Report — Latest findings and rectifications

#### Inspection & Compliance Reports
- Annual Inspection Report (RBIA / internal inspection)
- RBIA Audit Report & Rectification Certificate
- Concurrent Audit Reports — Last 4 quarters
- Stock Audit Reports — For applicable CC/WC accounts
- RBI Inspection Comments
- Vigilance / Special Investigation Reports

---

## PART 2 — IRAC NORMS: NPA CLASSIFICATION & PROVISIONING

### 2.1 Asset Classification Framework

| Category | Overdue Threshold | Description |
|----------|-------------------|-------------|
| **Standard Asset** | 0–29 days | Normal income-generating asset |
| **SMA-0** | 30–59 days overdue | Special Mention; early stress |
| **SMA-1** | 30–59 days | Enhanced monitoring required |
| **SMA-2** | 60–89 days | Elevated stress; close to NPA |
| **Sub-standard NPA** | 90 days – 12 months | Non-Performing; income not recognised |
| **Doubtful NPA** | 12–36 months | Graded provisioning based on security |
| **Loss Asset** | >36 months | Irrecoverable; 100% provision; write off |

**SMA Tags used in the Bank Audit Intelligence Suite tool:**

```
irregularity_months ≤ 1  → Standard
irregularity_months ≤ 2  → SMA-0
irregularity_months ≤ 3  → SMA-1
irregularity_months ≤ 4  → SMA-2
irregularity_months > 4  → NPA
```

### 2.2 NPA Identification — Product-Wise Criteria

**Term Loans:**
- Interest or instalment overdue for more than 90 days from due date
- EMI system: recovery first applied against earliest dues (first due, first served — per ICAI opinion)

**Cash Credit / Overdraft:**
- Account "out of order" if outstanding continuously exceeds sanctioned limit for 90 days
- OR credits in the account not sufficient to cover interest debited during preceding 90 days
- OR no credits continuously for 90 days

**Bills Purchased / Discounted:**
- Bill overdue for more than 90 days

**Agricultural Advances:**
- Short duration crop: overdue for 2 crop seasons beyond due date
- Long duration crop: overdue for 1 crop season beyond due date
- Minimum 1-year moratorium for natural calamity (33%+ crop loss); rescheduling allowed

**Credit Cards:**
- Minimum amount unpaid for 90 days from statement date

### 2.3 Provisioning Norms

| Asset Classification | Provision Rate | Key Criteria |
|---------------------|----------------|--------------|
| Standard Asset | 0.40% | General provision on total outstanding |
| SMA-1 | +0.25% (total 0.65%) | Additional over standard provision |
| SMA-2 | +0.40% (total 0.80%) | Further additional; high-stress pre-NPA |
| Sub-standard (Secured) | 15% | On total outstanding minus realisable security |
| Sub-standard (Unsecured) | 25% | Higher rate; no security |
| Doubtful D1 (0–1 yr) | 25% secured + 100% unsecured | |
| Doubtful D2 (1–3 yrs) | 40% secured + 100% unsecured | |
| Doubtful D3 (>3 yrs) | 100% | Full provision |
| Loss Asset | 100% | Should be written off from books |

### 2.4 Key IRAC Principles
- **Objective, uniform criteria** — same rules apply across all accounts regardless of account history
- **System-based automation** — CBS flags SMA/NPA at day-end automatically
- **Income de-recognition** — Interest on NPA accounts must be reversed; not recognised as income
- **Provisioning on net basis** — Computed on outstanding minus realisable value of tangible security
- **Valuation freshness** — Security valuation must not be older than 2 years
- **Evergreening prohibited** — Routing repayments through new loans to avoid NPA classification is prohibited

---

## PART 3 — COMPUTATIONAL AUDIT LOGIC (TERM LOAN ANALYSIS)

This section documents the exact mathematical logic embedded in the Bank Audit Intelligence Suite tool.

### 3.1 EMI Calculation

Standard reducing balance formula:

```
EMI = P × R × (1+R)^N / ((1+R)^N - 1)

Where:
  P = Sanctioned Principal (Loan Amount)
  R = Monthly Interest Rate = (Annual ROI% / 100) / 12
  N = Loan Tenure in Months = DATEDIF(Sanction Date, Expiry Date, months)
```

If R = 0 (zero interest), then: `EMI = P / N`

### 3.2 Tenure Calculation

```
Tenure (Months) = DATEDIF(Sanction Date, Expiry Date) in months
               = (Year(Expiry) - Year(Sanction)) × 12 + Month(Expiry) - Month(Sanction)
```

### 3.3 Instalments Due (as on Audit Date)

```
Instalments Due = (Year(Audit Date) - Year(Sanction Date)) × 12
                + Month(Audit Date) - Month(Sanction Date) + 1
```

This uses **inclusive counting** — the month of sanction counts as instalment #1.

### 3.4 Theoretical Outstanding Balance (Amortization)

After `n` instalments have been serviced:

```
Theo O/s Balance = P × [(1+R)^N - (1+R)^n] / [(1+R)^N - 1]
```

At zero rate: `Theo O/s = P - (P/N × n)`

### 3.5 Irregularity / Overdue EMI Count

```
O/s Difference = MAX(0, CBS Net Outstanding - Theo O/s Balance)
Outstanding EMIs = FLOOR(O/s Difference / Theo EMI, 1)
```

This is **Method A (Outstanding Difference)** — the stricter, auditor-preferred method.

### 3.6 Theoretical vs CBS Asset Classification

Using Outstanding EMI count:
```
≤1 EMI overdue  → Standard
≤2 EMIs overdue → SMA-0
≤3 EMIs overdue → SMA-1
≤4 EMIs overdue → SMA-2
>4 EMIs overdue → NPA
```

**Key audit flag:** If Bank's classification = Standard BUT Theoretical classification = SMA-x or NPA → **REPORT IN LFAR Section I.3(e)(vi) — potential underclassification**

### 3.7 Credit Balance Netting

For CC/OD accounts: `Net Effective Exposure = Net O/s - Credit Balance`

This netting should be done before computing overdrawn status and before comparing with Theo O/s.

### 3.8 Overdrawn Detection (CC/OD Accounts)

```
Overdrawn = YES if (Net O/s - Credit Balance) > Sanction Limit
```

Overdrawn by >110%: `Net O/s > Sanction Limit × 1.10` → Flag for Overdrawn Accounts sheet

### 3.9 PAN-Based Entity Type Derivation

The 4th character of PAN reveals entity type:
| PAN 4th Char | Entity Type |
|---|---|
| P | Individual (Person) |
| C | Company (Private/Public Ltd) |
| H | HUF (Hindu Undivided Family) |
| F | Firm (Partnership) |
| A | Association of Persons (AOP) |
| T | Trust |
| B | Body of Individuals (BOI) |
| L | Local Authority |
| J | Artificial Juridical Person |
| G | Government |

### 3.10 Sanction Type (New vs Old)

```
New = If Sanction Date falls within the Audit Period (From Date to To Date)
Old = If Sanction Date is before the Audit Period start date
```

### 3.11 Excel Formulas Embedded in Tool (TL Accounts Sheet)

| Column | Formula |
|--------|---------|
| Tenure (Months) | `=MAX(0, (YEAR(ExpiryDate)-YEAR(SanctionDate))*12 + MONTH(ExpiryDate)-MONTH(SanctionDate))` |
| Instalments Due | `=MAX(0, (YEAR(AuditDate)-YEAR(SanctionDate))*12 + MONTH(AuditDate)-MONTH(SanctionDate) + 1)` |
| Theo EMI | `=IFERROR(PMT(ROI/12, Tenure, -SanctionLimit), 0)` |
| Theo Principal Paid | `=IFERROR(IF(ROI=0, (SanctionLimit/Tenure)*MIN(InstDue,Tenure), -CUMPRINC(ROI/12,Tenure,SanctionLimit,1,MIN(InstDue,Tenure),0)), 0)` |
| Theo Interest Charged | `=IFERROR(IF(ROI=0, 0, -CUMIPMT(ROI/12,Tenure,SanctionLimit,1,MIN(InstDue,Tenure),0)), 0)` |
| Theo O/s Balance | `=MAX(0, SanctionLimit - TheoPrincipalPaid)` |
| O/s Difference | `=MAX(0, NetOs - TheoOs)` |
| Outstanding EMIs | `=IF(AND(OsDiff>0,TheoEMI>0), FLOOR(OsDiff/TheoEMI,1), 0)` |
| Asset Class (Theo Bal) | `=IF(OutstandingEMIs<=1,"Standard",IF(OutstandingEMIs<=2,"SMA-0",IF(OutstandingEMIs<=3,"SMA-1",IF(OutstandingEMIs<=4,"SMA-2","NPA"))))` |

---

## PART 4 — LFAR STRUCTURE (RBI REVISED FORMAT 2020)

The Long Form Audit Report (LFAR) is the primary deliverable. It follows Annex II of the RBI revised format (Sept 2020).

### 4.1 Section I — Assets

**I.1 Cash**
- Joint custody compliance; surprise checks conducted; insurance cover adequacy; CBS vs physical variance

**I.2 Balances with RBI/SBI/Other Banks**
- Inter-bank reconciliation completeness; aging of outstanding items; adverse financial impact of unreconciled items

**I.3 Advances (Most Critical Section)**

ANNEX III required for all large accounts (>10% of aggregate advances OR above ₹10 crore).

Key sub-clauses:
- `(a)(i)` — Large accounts examined; credit appraisal quality
- `(a)(ii)` — Sanction authority compliance; delegated authority
- `(b)(i)` — Documentation completeness; charges registered; securities obtained
- `(b)(iii)` — Documentation deficiencies; waivers; missing guarantees
- `(c)(ii)` — Drawing Power computation; stock statement regularity
- `(d)(i)` — Review and renewal compliance; short-reviewed accounts
- `(e)(i)` — NPA identification: fresh NPAs; upgraded accounts
- `(e)(ii)` — NPA classification accuracy: sub-standard, doubtful, loss
- `(e)(vi)` — **Accounts where auditor's classification differs from bank's classification**
- `(e)(vii)` — Provision adequacy; comparison with bank's computed provision
- `(e)(ix)` — Income recognition on NPA accounts; unrealised interest reversal
- `(f)(i)` — Fraud accounts: reported to RBI; provision status
- `(f)(v)` — Restructured accounts; compliance with restructuring norms
- `(f)(vii)` — SARFAESI/legal proceedings status
- `(f)(x)` — OTS cases; adequacy of settlement terms
- `(f)(xii)` — Wilful defaulters; reporting compliance

### 4.2 Section II — Liabilities

- `II.1(b)` — Deposit classification accuracy (SA/CA/TD/RD); interest accrual; DEAF transfers
- `II.2(a)` — Nature and aging of sundry creditors; items >3 years outstanding
- `II.3` — LC/BG completeness; bills for collection; derivative MTM valuation

### 4.3 Section III — Profit & Loss

- Interest income recognition per IRAC norms
- Investment income, fees, other income — correctly recorded
- Interest paid on deposits — accuracy of rate and accrual
- NPA provision, depreciation on investments

### 4.4 Section IV — Governance & Compliance

- **IV.1 Fraud Cases** — All fraud identified, reported, provided for
- **IV.2 Regulatory Compliance** — KYC/AML, PSL targets, government schemes (PMEGP, Padho Pradesh, CSIS)
- **IV.3 Internal Controls** — ICFR/IFCOFR; Risk Control Matrix (RCM); exception transaction monitoring
- **IV.4 Books & Records** — Banking Regulation Act compliance; subsidiary registers vs GL reconciliation; audit trail in CBS
- **IV.5 Inter-Branch Reconciliation** — Aging of unreconciled items; items >6 months

---

## PART 5 — CBPMS PORTAL & AIS PORTAL REFERENCE

### 5.1 CBPMS Portal Reports (Final Submission)

| CBPMS Code | Description | Audit Purpose |
|-----------|-------------|---------------|
| BRANCH_AUDITOR_REPORT | Branch Auditor Report (primary statutory report) | Core audit report |
| MOCBSPL / MOC_LOANFUND | MOC — Loan Fund/Non-Fund | MOC disclosures; BS & PL impact |
| BRANCH_MOC_SUMMARY_ANNUAL | Annual MOC Summary | Consolidated for SCA |
| BRANCH_PREVIOUSYEAR_MOC | Prior Year MOC Implementation | Verify prior year MOC effected |
| BRANCH_RECONCILIATION_STATEMENT | Branch Reconciliation Statement | Inter-branch/HO reconciliation |
| INCOME_EXPENDITURE | Income & Expenditure Statement | P&L verification |
| BSPL_Unaudited_Report | Unaudited BS & P&L | Opening pre-audit financials |
| BRANCH_DICGC_SUMMARY_NEW | DICGC Insurance Summary | Deposit insurance compliance |
| RISK_CONTROL_MATRIX_BRANCH | Risk Control Matrix — Branch | ICFR testing |
| IAD_BRANCH | Internal Audit Department Report | Cross-reference findings |
| CSIS_BRANCH | Credit Score — Branch | Credit risk scoring |
| PADHO_PRADESH_BRANCH | Padho Pradesh Education Loan Scheme | Govt scheme compliance |

### 5.2 AIS Portal — RBI Auditor Information System (Annual Forms)

| AIS Form | Description |
|----------|-------------|
| FORM1_ANNUAL | Auditor particulars, firm details, branch details, audit period |
| FORM2_ANNUAL | Summary of large advance accounts reviewed |
| FORM3_ANNUAL | NPA accounts details and classification |
| FORM4_ANNUAL | Provision adequacy statement |
| FORM5_ANNUAL | MOC details — adjustments to branch financials |
| FORM6_ANNUAL | Fraud accounts and irregularities |
| FORM7_ANNUAL | Status of compliance with prior year observations |
| FORM8_ANNUAL | SMA accounts position |
| FORM9_ANNUAL | Restructured accounts details |
| FORM10_ANNUAL | Overall auditor observations and recommendations |

---

## PART 6 — DETAILED AUDIT PROCEDURES

### 6.1 Pre-Audit Planning
1. Obtain A1A4 report; stratify advances by outstanding amount, type (CC/TL/OD), and risk grade
2. Identify large advances — accounts exceeding 10% of aggregate FB+NFB limits or ₹10 crore
3. Review prior year LFAR observations and MOC implementation
4. Review RBIA and concurrent audit reports for known deficiencies
5. Request CBS system access; verify report generation capability
6. Prepare audit program based on branch profile: CC vs TL mix, NPA level, industry concentration

### 6.2 Credit Appraisal Procedures
1. Obtain Credit Appraisal Memorandum (CAM); review against sanctioned terms
2. Verify borrower financials — last 2–3 years audited accounts, ITR, projected financials
3. Check: debt-equity ratio, interest coverage, DSCR, current ratio
4. Confirm credit rating from RBI-accredited agency (for accounts above rating threshold)
5. Review promoter profile, management background, and track record
6. Assess industry risk — sector-specific risks noted in appraisal
7. Check for "quick mortality" — accounts classified NPA within 12 months of first sanction

### 6.3 Sanctions & Disbursements
8. Verify sanction is within delegated authority of sanctioning official
9. Confirm all pre-sanction conditions satisfied before disbursement
10. Check disbursement in line with purpose — funds not diverted/misdirected
11. For Term Loans: verify stage-wise disbursement matches project milestones
12. Obtain Chartered Engineer/CA certificates for construction/project loans
13. Verify end-use: purchase invoices, stock statements, supplier payments match disbursement

### 6.4 Documentation Procedures
14. Confirm mortgage/pledge/lien/hypothecation created per sanction terms
15. Verify charge registration (EquiMortgage, SARFAESI registration) is completed
16. Check all guarantees (personal/corporate) are obtained and alive
17. Verify insurance policies are in force, adequate, and assigned to bank
18. Confirm title deeds/property documents in bank's custody
19. Check Deposit Receipts have bank lien marked in CBS
20. Note and report documentation deficiencies in LFAR

### 6.5 Stock Statement / Drawing Power (CC/WC Accounts)
- Verify stock statements received at required frequency (monthly for CC accounts)
- Check stock statement is on borrower's letterhead with authorised signature
- Compute Drawing Power (DP) from latest stock statement; verify CBS DP matches
- Identify accounts where DP < outstanding — these are "out of order"
- Note delayed submission or gaps in stock statement submissions

### 6.6 Periodic Review & Renewal
- All CC/WC accounts must be reviewed/renewed annually — check compliance
- Obtain CRMD-007 for pending renewal/review accounts; quantify risk exposure
- Verify latest audited financial statements obtained before renewal
- Confirm credit rating is current and not expired

### 6.7 NPA Monitoring
- Extract NPA list from AUD-001; cross-check against SMA-2 list from CRMD-003
- Verify NPA date is correct — overdue date vs date classified NPA in CBS
- Check income de-recognition: no interest income on NPA accounts
- Verify unrealised interest reversed and not carried in books

### 6.8 Annex III Working Paper for Large Accounts

For each large advance account, prepare Annex III covering:

| Annex III Item | What to Document |
|----------------|-----------------|
| Borrower & Facility Details | Name, loan a/c no., type, sanctioned limit, outstanding, security |
| Sanction History | Original sanction date, renewals, enhancements, sanctioning authority |
| Financial Performance | Latest P&L ratios, current ratio, DSCR, turnover |
| Collateral / Security | Primary security (type, value, valuation date); collateral (type, value, date) |
| NPA Status | Classification; NPA date; provision held; comparison with bank's provision |
| Observation | Deficiency in documentation, monitoring, NPA classification, end-use, fraud |
| Management Response | Branch Manager's response; action taken or planned |

---

## PART 7 — RED FLAGS DATABASE

These red flags are derived from real PNB branch audit observations and broader ICAI guidance. Use these as the core detection logic.

### 7.1 Credit Appraisal Red Flags
- Borrower financials not audited or more than 12 months old at time of sanction
- Promoter net worth negative or insufficient relative to loan size
- **Quick mortality**: NPA within 12 months of first disbursement — mis-selling or fraud indicator
- Credit rating not obtained or expired at time of renewal
- Industry-specific risks not assessed in credit appraisal
- DSCR/ICR below RBI minimum thresholds despite loan sanction

### 7.2 Disbursement & End-Use Red Flags
- Large portion of CC disbursement transferred to sister concerns/related parties immediately
- Funds routed through sub-accounts or third-party accounts with no business nexus
- Invoices from new vendors with recently registered GST (low credibility)
- E-way bills absent for claimed goods movements
- Purchase invoices for amounts higher than market rates (over-invoicing)
- Advance payments to contractors outstanding indefinitely (diversion risk)
- Payments to entities not appearing in borrower's own purchase ledger

### 7.3 Documentation Red Flags
- Mortgage not created despite immovable property as security
- Charge not registered with ROC/Sub-Registrar within prescribed time
- Guarantor's net worth not verified or guarantee letter not obtained
- Insurance policy expired or not assigned to bank
- Title deeds for property not in bank's safe custody
- Document waiver obtained without appropriate authority sanction

### 7.4 Monitoring & Drawing Power Red Flags
- Stock statements not received for more than 2 consecutive months for CC account
- Stock statement received but Drawing Power not recomputed in CBS
- Outstanding in CC account exceeds DP continuously — "out of order" but not flagged NPA
- Debtor list submitted is fictitious or contains round-number figures
- Stock statements on plain paper without borrower letterhead or authorised signature
- Large reduction in stock statement values in final quarter before audit date

### 7.5 NPA Classification Red Flags (Highest Audit Risk)
- SMA-2 account not classified NPA despite 90+ days overdue — system override or manual intervention
- Interest being recognised on account that should be classified NPA
- NPA classification date later than actual overdue date — provision under-stated
- Accounts showing regular credits but credits are from related parties ("evergreening")
- Restructured account not complying with restructuring terms — should revert to NPA
- OTS settlement less than outstanding + accrued interest — inadequate provision
- **Computational mismatch**: Tool shows NPA/SMA-x but bank's Asset Classification = Standard → LFAR Section I.3(e)(vi) disclosure required

### 7.6 Fraud Indicators (CBS-Based)
- Exception transaction report shows repeated limit overrides without system sanction
- Large withdrawals just before period-end with no supporting business justification
- Multiple accounts of same entity with simultaneous high-value debits
- Same invoice number submitted against multiple accounts/borrowers
- Credit in one account matched by debit in another without genuine economic transaction
- Unexplained credits from unknown counterparties just before quarter-end to avoid SMA classification

### 7.7 Sector-Specific — Warehouse / Commodity Loans (Obra Branch Pattern)
- Verify physical stock existence via independent warehouse inspection
- Crop season alignment: ensure advance for crop X is given in correct season window
- Apply agricultural IRAC criteria (crop season overdue) — NOT the 90-day rule
- Verify warehouse receipt registration and insurance on stored commodity
- Check market price vs loan value — margin erosion from price volatility
- Warehouse loan: verify collateral manager's independence and accreditation

---

## PART 8 — MOC (MEMORANDUM OF CHANGES)

The MOC documents all corrections and adjustments required to be made in branch books.

### 8.1 Typical MOC Items

| MOC Category | Direction | Typical Scenario |
|---|---|---|
| NPA Provision Shortfall | Dr. P&L / Cr. Provision A/c | Bank's provision < RBI norm |
| Unrealised Interest Reversal | Dr. Interest Income / Cr. NPA A/c | Interest booked on NPA must be reversed |
| Income Recognition Adjustment | Dr. Interest Income / Cr. Suspense | Interest accrued but not realised |
| Standard Asset Provision | Dr. P&L / Cr. Provision A/c | 0.40% general provision not fully provided |
| Suspense Account Clearance | Various | Old debit items in suspense not cleared |
| Depreciation on Investments | Dr. P&L / Cr. Investments | MTM depreciation not accounted |

### 8.2 MOC Format Requirements
Each MOC item must specify:
- Account/GL number
- Nature of change
- Amount
- Debit/Credit side
- Impact on P&L / Balance Sheet
- System Entry Proof (CBS screenshot or batch log)
- Authority Approval (BM / Regional Manager)

### 8.3 Prior Year MOC
Check BRANCH_PREVIOUSYEAR_MOC report to verify that prior year MOC items were actually entered into CBS.

---

## PART 9 — WRITTEN REPRESENTATION LETTER (MRL)

Required under SA 580 (Written Representations). PNB calls this the MRL (Management Representation Letter).

### Mandatory Contents
- Compliance with applicable accounting standards (IndAS / Banking Regulation Act)
- Completeness: all advances, NPAs, provisions, and contingent liabilities disclosed
- Accuracy: amounts correctly computed and recorded
- NPA Classification: all accounts classified per RBI IRAC norms
- Security Valuation: all security valuations current (not more than 2 years old)
- Fraud: all known fraud cases reported to RBI and provided for
- Litigation: all pending litigations and contingent liabilities disclosed
- Going Concern: no events casting doubt on going concern
- Audit Access: all requested information and access provided
- Subsequent Events: no events after balance sheet date requiring adjustment

**Risk Note:** Non-receipt or qualified MRL may require a qualification or emphasis of matter paragraph in the main audit report.

---

## PART 10 — IFCOFR (Internal Financial Controls)

Since FY 2019-20, branch auditors issue a separate report on IFCOFR.

Requirements:
- Separate Engagement Letter for IFCOFR
- Testing of key controls: CBS access controls, limit override authorization, NPA auto-classification
- Risk Control Matrix (RCM) review at branch level
- CBPMS submission: RISK_CONTROL_MATRIX_BRANCH report
- Opinion: whether ICFR is adequate and operating effectively

---

## PART 11 — END-TO-END AUDIT WORKFLOW

### Pre-Audit Stage
- STEP 1: Receive RBI appointment letter — review branch assignments and confirm acceptance
- STEP 2: Submit Consent Letter (with Firm Certificate — Form 135113W) to bank HO
- STEP 3: Sign and return Letter of Acceptance to the bank
- STEP 4: Prepare and send Data Requirement List + Engagement Letter to branch
- STEP 5: Send IFCOFR Engagement Letter separately
- STEP 6: Request CBS user credentials and report server access

### Fieldwork Stage
- STEP 7: Extract CBS reports (A1A4, SMA, NPA lists) and build advance population
- STEP 8: Select sample of large advances for detailed review (Annex III accounts)
- STEP 9: Review credit files — appraisal, sanctions, disbursements, documentation
- STEP 10: Verify NPA classification and provision adequacy for all NPA accounts
- STEP 11: Perform stock statement / drawing power procedures for CC accounts
- STEP 12: Review RBIA, concurrent audit, and inspection reports for cross-reference
- STEP 13: Verify trial balance, suspense accounts, and inter-branch reconciliation
- STEP 14: Check CBPMS MOCBSPL and other MOC reports for prior year effects

### Reporting Stage
- STEP 15: Prepare LFAR draft with all section-wise observations
- STEP 16: Prepare MOC list — quantify adjustments to branch financials
- STEP 17: Obtain Written Representation Letter (MRL) from Branch Manager
- STEP 18: Submit MOC for system implementation — collect CBS entry proof
- STEP 19: Finalise LFAR and Annex III for each large advance
- STEP 20: Submit reports through CBPMS Portal (all relevant forms)
- STEP 21: Generate UDIN and submit AIS forms (FORM1 through FORM10)
- STEP 22: Collect signed LFAR and Annex III from branch
- STEP 23: Forward to SCA (Statutory Central Auditor) for bank-level consolidation
- STEP 24: Final submission to RBI SSM within 60 days

### PNB Reimbursement Claim (Annexure T-3 Format)
- A — Travelling Expenses (boarding passes and flight/train invoices)
- B — Lodging Charges (hotel invoices)
- C — Boarding Charges (meal/food invoices where applicable)
- D — Local Conveyance Charges (cab/auto receipts)

---

## PART 12 — BRANCH FILE / FOLDER STRUCTURE (REFERENCE)

| Folder Path | Contents |
|-------------|----------|
| FY 20XX-XX / Branch (Code) / Appointment & Engagement | Consent letter, firm certificate, data requirement, engagement letters (main + IFCOFR), letter of acceptance, NOC |
| FY 20XX-XX / Branch (Code) / Branch Data / CBS Reports | AUD, CRMD, MISD, MSME, OPR, BARM, SASTRA report extracts |
| FY 20XX-XX / Branch (Code) / Branch Data / NPA & Provision Data | NPA list, provision movement, OTS details, upgraded/downgraded accounts |
| FY 20XX-XX / Branch (Code) / Branch Data / SMA Data | CRMD-003 SMA reports (December and March dates) |
| FY 20XX-XX / Branch (Code) / Branch Data / RBIA Reports | Risk-based internal audit report and closure certificate |
| FY 20XX-XX / Branch (Code) / Branch Data / Account Statements | Individual borrower account statements (CBS text exports) |
| FY 20XX-XX / Branch (Code) / Advances Working Papers / [Borrower Name] | Stock statements, invoices, ledgers, sanction letter, credit evaluation, site visit photos, audit queries |
| FY 20XX-XX / Branch (Code) / LFAR Annexures | Excel LFAR annexure files (Assets, Liabilities, etc.) |
| FY 20XX-XX / Branch (Code) / MOC & Written Representation | MOC working, MRL, IRAC reference |
| FY 20XX-XX / Branch (Code) / RBI Notifications & Circulars | RBI circulars relevant to branch issues |
| FY 20XX-XX / Branch (Code) / Final Reports / LFAR | Signed LFAR, Annex III |
| FY 20XX-XX / Branch (Code) / Final Reports / CBPMS Reports | All CBPMS portal PDFs |
| FY 20XX-XX / Branch (Code) / Final Reports / AIS Data | AIS FORM1 through FORM10 annual PDFs |
| FY 20XX-XX / Bank Audit Material | ICAI Guidance Note, Finacle CBS Handbook, reference PPTs |
| FY 20XX-XX / ICAI Guidance Notes | ICAI GN Part B and Part C2 for the year |
| FY 20XX-XX / Reimbursement Details / [Location] | Annexure T-3 claim forms with invoices |

---

## PART 13 — TOOL INTEGRATION: BANK AUDIT INTELLIGENCE SUITE

The Bank Audit Intelligence Suite (React/Vite app) automates the computational part of the audit. Here's how it integrates with the skill:

### What the Tool Does
1. Accepts CBS extract (A1A4-equivalent) as Excel input
2. Maps columns to standard audit fields (CIF, A/c No., Borrower Name, ROI, Sanction Limit, Net O/s, Sanction Date, Expiry Date, etc.)
3. Auto-calculates: Loan Type, Overdrawn, Sanction Type (New/Old), Entity Type (from PAN), Asset Classification
4. For Term Loans: computes Tenure, Instalments Due, Theo EMI, Theo Principal Paid, Theo Interest Charged, Theo O/s Balance, O/s Difference, Outstanding EMIs, Asset Class (Theo Bal)
5. Exports multi-sheet Excel workbook: Audit Report, ROI Analysis, CC Accounts, TL Accounts, NPA Accounts, Overdrawn Accounts (>110%)

### Key Output Sheets for Audit Evidence
| Sheet | Audit Use |
|-------|-----------|
| Audit Report | Master population; completeness check |
| TL Accounts | Independent NPA classification via amortization; pink-highlighted rows = misclassification risk |
| CC Accounts | Out-of-order CC accounts; drawing power analysis |
| NPA Accounts | NPA population with provision data |
| Overdrawn (>110%) | Standard accounts exceeding sanctioned limit significantly |
| ROI Analysis | Interest rate variance; potential rate manipulation |

### Pink Highlighting Rule (Critical)
Rows highlighted in light pink on TL Accounts sheet mean:
- Bank's `Asset Classification` = Standard (or STD)
- Tool's `Asset Class (Theo Bal)` = SMA-0/SMA-1/SMA-2/NPA

→ **These accounts require mandatory LFAR disclosure under Section I.3(e)(vi)**

---

## PART 14 — QUICK REFERENCE: KEY REGULATORY NUMBERS

| Parameter | Value | Source |
|-----------|-------|--------|
| NPA threshold (standard advances) | 90 days overdue | RBI IRAC Master Circular |
| Agricultural NPA — short duration crop | 2 crop seasons | RBI IRAC |
| Agricultural NPA — long duration crop | 1 crop season | RBI IRAC |
| Standard asset provision | 0.40% | RBI IRAC |
| Sub-standard (secured) provision | 15% | RBI IRAC |
| Sub-standard (unsecured) provision | 25% | RBI IRAC |
| Doubtful D1 provision (secured) | 25% | RBI IRAC |
| Doubtful D2 provision (secured) | 40% | RBI IRAC |
| Doubtful D3/Loss provision | 100% | RBI IRAC |
| Security valuation max age | 2 years | RBI IRAC |
| Large advance threshold | >10% aggregate OR >₹10 cr | ICAI Guidance Note |
| Annexure III requirement | All large advances | LFAR Format |
| LFAR submission deadline | 60 days from audit date | RBI circular |
| Inoperative account DEAF transfer | After 10 years | RBI circular |
| IFCOFR effective from | FY 2019-20 | RBI/ICAI |
| Concurrent audit overlap period | Last 4 quarters | Best practice |
| Stock statement frequency (CC) | Monthly | RBI/bank policy |

---

*This knowledge base was compiled from: PNB Branch Audit Practice Files (FY 2023-24, FY 2024-25), Bank Audit Intelligence Suite codebase, ICAI Guidance Note on Audit of Banks (2024/2025), RBI IRAC Master Circular, and RBI Revised LFAR Format (Sept 2020).*
