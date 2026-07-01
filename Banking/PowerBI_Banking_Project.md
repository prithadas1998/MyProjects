# 🏦 Banking & Finance — Power BI Project

## Overview
A complete Power BI dashboard covering four banking domains:
- **Loan Portfolio & Credit Risk**
- **Transactions & Cash Flow**
- **Customer Accounts & Deposits**
- **Branch Performance & KPIs**

---

## 📁 Data Model — Tables & Relationships

### Tables (all tabs in `banking_data.xlsx`)
| Table | Rows | Key Column | Description |
|---|---|---|---|
| Customers | 500 | CustomerID | Master customer data |
| Accounts | 800 | AccountID / CustomerID | Deposit accounts |
| Transactions | 5,000 | TransactionID / AccountID | All financial transactions |
| Loans | 300 | LoanID / CustomerID | Loan portfolio |
| Branches | 20 | BranchCode | Branch metadata & targets |

### Relationships (set in Power BI Model View)
```
Customers (CustomerID)  ──1:M──▶  Accounts (CustomerID)
Customers (CustomerID)  ──1:M──▶  Loans (CustomerID)
Accounts  (AccountID)   ──1:M──▶  Transactions (AccountID)
Accounts  (BranchCode)  ──M:1──▶  Branches (BranchCode)
```

---

## 🔧 Power Query — Data Load Steps

### Step 1 — Connect to Excel
1. **Home → Get Data → Excel Workbook**
2. Select `banking_data.xlsx`
3. Check all 5 tables → **Load**

### Step 2 — Transform Each Table
In Power Query Editor, apply these changes:

**Transactions:**
```
= Table.TransformColumnTypes(Source, {
    {"TransactionDate", type date},
    {"Amount", type number}
  })
```

**Loans:**
```
= Table.TransformColumnTypes(Source, {
    {"DisbursalDate", type date},
    {"LoanAmount", Currency.Type},
    {"OutstandingAmt", Currency.Type}
  })
```

**Accounts:**
```
= Table.TransformColumnTypes(Source, {
    {"Balance", Currency.Type},
    {"OpenDate", type date}
  })
```

---

## 📅 Date Table (Create in Power BI)

```dax
DateTable = 
ADDCOLUMNS(
    CALENDAR(DATE(2023,1,1), DATE(2024,12,31)),
    "Year",          YEAR([Date]),
    "Month",         FORMAT([Date], "MMM YYYY"),
    "MonthNum",      MONTH([Date]),
    "Quarter",       "Q" & QUARTER([Date]) & " " & YEAR([Date]),
    "WeekNum",       WEEKNUM([Date]),
    "DayOfWeek",     FORMAT([Date], "DDD"),
    "IsWeekend",     IF(WEEKDAY([Date],2) >= 6, TRUE, FALSE),
    "FiscalYear",    IF(MONTH([Date]) >= 4, "FY" & YEAR([Date]), "FY" & YEAR([Date])-1)
)
```

**Link:**  `DateTable[Date]` → `Transactions[TransactionDate]`  (1:M)

---

## 📐 DAX Measures Library

### ── SECTION 1: TRANSACTION KPIs ──

```dax
Total Transaction Volume = 
SUMX(FILTER(Transactions, Transactions[Status] = "Completed"), Transactions[Amount])

Total Transactions Count = 
COUNTROWS(FILTER(Transactions, Transactions[Status] = "Completed"))

Avg Transaction Value = 
DIVIDE([Total Transaction Volume], [Total Transactions Count], 0)

Transaction Failure Rate = 
DIVIDE(
    COUNTROWS(FILTER(Transactions, Transactions[Status] = "Failed")),
    COUNTROWS(Transactions),
    0
)

MoM Transaction Growth = 
VAR CurrMonth = [Total Transaction Volume]
VAR PrevMonth = CALCULATE([Total Transaction Volume], DATEADD(DateTable[Date], -1, MONTH))
RETURN DIVIDE(CurrMonth - PrevMonth, PrevMonth, 0)

YTD Transaction Volume = 
TOTALYTD([Total Transaction Volume], DateTable[Date])

Deposit Volume = 
CALCULATE([Total Transaction Volume], Transactions[TransactionType] = "Deposit")

Withdrawal Volume = 
CALCULATE([Total Transaction Volume], Transactions[TransactionType] = "Withdrawal")

Net Cash Flow = [Deposit Volume] - [Withdrawal Volume]
```

---

### ── SECTION 2: LOAN PORTFOLIO KPIs ──

```dax
Total Loans Disbursed = 
SUM(Loans[LoanAmount])

Total Outstanding = 
SUMX(FILTER(Loans, Loans[LoanStatus] = "Active"), Loans[OutstandingAmt])

NPA Amount = 
CALCULATE(SUM(Loans[OutstandingAmt]), Loans[LoanStatus] = "NPA")

NPA Ratio = 
DIVIDE([NPA Amount], [Total Outstanding], 0)

Gross NPA % = 
DIVIDE(
    CALCULATE(SUM(Loans[LoanAmount]), Loans[LoanStatus] = "NPA"),
    SUM(Loans[LoanAmount]),
    0
)

Active Loan Count = 
CALCULATE(COUNTROWS(Loans), Loans[LoanStatus] = "Active")

Avg Loan Size = 
DIVIDE([Total Loans Disbursed], COUNTROWS(Loans), 0)

Avg Interest Rate = 
AVERAGEX(FILTER(Loans, Loans[LoanStatus] = "Active"), Loans[InterestRate])

DPD 90+ Count = 
CALCULATE(COUNTROWS(Loans), Loans[DPD] >= 90)

DPD 90+ Ratio = 
DIVIDE([DPD 90+ Count], [Active Loan Count], 0)

Loan Collection Efficiency = 
1 - DIVIDE([NPA Amount], [Total Outstanding], 0)
```

---

### ── SECTION 3: CUSTOMER & ACCOUNT KPIs ──

```dax
Total Customers = DISTINCTCOUNT(Customers[CustomerID])

Active Customers = 
CALCULATE(DISTINCTCOUNT(Customers[CustomerID]), Customers[IsActive] = "Yes")

New Customers MTD = 
CALCULATE(
    COUNTROWS(Customers),
    DATESMTD(DateTable[Date]),
    USERELATIONSHIP(DateTable[Date], Customers[JoinDate])
)

Avg Credit Score = AVERAGE(Customers[CreditScore])

High Risk Customers = 
CALCULATE(COUNTROWS(Customers), Customers[CreditScore] < 650)

Total Deposits = 
CALCULATE(SUM(Accounts[Balance]), Accounts[AccountType] IN {"Savings","Fixed Deposit","Recurring Deposit"})

Total CASA = 
CALCULATE(SUM(Accounts[Balance]), Accounts[AccountType] IN {"Savings","Current"})

CASA Ratio = 
DIVIDE([Total CASA], SUM(Accounts[Balance]), 0)

Avg Account Balance = AVERAGE(Accounts[Balance])

Active Accounts = 
CALCULATE(COUNTROWS(Accounts), Accounts[Status] = "Active")

Dormant Accounts = 
CALCULATE(COUNTROWS(Accounts), Accounts[Status] = "Dormant")

Customer Lifetime Value = 
AVERAGEX(
    Customers,
    CALCULATE(SUM(Transactions[Amount]), Transactions[Status] = "Completed")
)
```

---

### ── SECTION 4: BRANCH KPIs ──

```dax
Branch Deposit Achievement = 
DIVIDE(
    CALCULATE(SUM(Accounts[Balance])),
    SUM(Branches[TargetDeposits]),
    0
)

Branch Loan Achievement = 
DIVIDE(
    CALCULATE(SUM(Loans[LoanAmount])),
    SUM(Branches[TargetLoans]),
    0
)

Top Branch by Deposits = 
TOPN(1, VALUES(Branches[BranchName]), [Total Deposits])

Branches Hitting Target = 
COUNTROWS(
    FILTER(
        VALUES(Branches[BranchCode]),
        [Branch Deposit Achievement] >= 1
    )
)

Avg Staff per Branch = AVERAGE(Branches[Staff])

Revenue per Staff = 
DIVIDE([Total Transaction Volume], SUM(Branches[Staff]), 0)
```

---

## 📊 Dashboard Pages — Layout Guide

### Page 1 — Executive Summary
| Visual | Type | Fields |
|---|---|---|
| Total Deposits | Card | Total Deposits |
| Total Loans | Card | Total Loans Disbursed |
| NPA Ratio | Card (Red if >5%) | NPA Ratio |
| CASA Ratio | Card | CASA Ratio |
| Monthly Transaction Trend | Line Chart | Month, Transaction Volume |
| Loan vs Deposit by Segment | Clustered Bar | Segment, Loans, Deposits |
| Region Map | Filled Map | Region, Transaction Volume |

### Page 2 — Loan Portfolio & Credit Risk
| Visual | Type | Fields |
|---|---|---|
| NPA by Loan Type | Donut | Loan Type, NPA Amount |
| DPD Distribution | Histogram | DPD buckets, Count |
| Credit Score Distribution | Histogram | CreditScore |
| Outstanding by Status | Stacked Bar | LoanStatus, OutstandingAmt |
| Loan Disbursals Over Time | Area Chart | DisbursalDate, LoanAmount |
| Top 10 Risky Accounts | Table | CustomerID, DPD, OutstandingAmt |

### Page 3 — Transactions & Cash Flow
| Visual | Type | Fields |
|---|---|---|
| Net Cash Flow | Waterfall | Month, Deposits, Withdrawals |
| Transactions by Channel | Pie | Channel, Count |
| Txn Type Breakdown | Bar | TransactionType, Amount |
| Failure Rate Trend | Line | Month, Failure Rate |
| Hourly Heatmap (optional) | Matrix | DayOfWeek, Hour, Volume |

### Page 4 — Customer & Accounts
| Visual | Type | Fields |
|---|---|---|
| Customer by Segment | Donut | Segment, Count |
| Customer by Region | Bar | Region, Count |
| CASA vs FD Balance | Stacked Column | Month, CASA, FD |
| Avg Credit Score by Segment | Bar | Segment, Avg Credit Score |
| Account Status Breakdown | Donut | Status, Count |
| New Customers Trend | Line | Month, New Customers |

### Page 5 — Branch Performance
| Visual | Type | Fields |
|---|---|---|
| Target Achievement | Bullet/Gauge | BranchName, Achievement % |
| Deposits by Branch | Bar | BranchName, Total Deposits |
| Loans by Branch | Bar | BranchName, Total Loans |
| Branch Map | Bubble Map | Region, Deposits |
| Revenue per Staff | Bar | BranchName, Revenue/Staff |

---

## 🎨 Theme & Formatting

### Color Palette
```json
{
  "name": "Banking Finance",
  "dataColors": [
    "#1F3864", "#2E75B6", "#4472C4",
    "#70AD47", "#FFC000", "#FF0000",
    "#A9D18E", "#9DC3E6"
  ],
  "background": "#FFFFFF",
  "foreground": "#1F3864",
  "tableAccent": "#2E75B6"
}
```

Save as `BankingTheme.json` and load via **View → Themes → Browse for themes**

### Conditional Formatting Rules
- **NPA Ratio**: Red if > 5%, Yellow if 3–5%, Green if < 3%
- **DPD**: Red if ≥ 90, Yellow if 30–89, Green if 0
- **Branch Achievement**: Green if ≥ 100%, Yellow if 80–99%, Red if < 80%
- **Credit Score**: Red if < 650, Yellow if 650–700, Green if > 700

---

## 🔐 Row-Level Security (RLS)

```dax
-- Role: Region Manager
[Region] = USERNAME()

-- Role: Branch Manager  
[BranchCode] = LOOKUPVALUE(Branches[BranchCode], Branches[BranchName], USERNAME())
```

Set up via **Modeling → Manage Roles**

---

## 📤 Publishing & Sharing

1. **Save** the `.pbix` file
2. **Publish** → Select workspace → Confirm
3. In Power BI Service: **Schedule Refresh** (if using gateway)
4. **Share** → Enter email or embed in SharePoint/Teams

---

## ✅ Quick Setup Checklist

- [ ] Load `banking_data.xlsx` into Power BI Desktop
- [ ] Set all column data types (date, currency, text)
- [ ] Create `DateTable` using DAX above
- [ ] Set all 4 table relationships in Model View
- [ ] Copy all DAX measures into a dedicated `_Measures` table
- [ ] Build 5 dashboard pages per layout guide
- [ ] Apply `BankingTheme.json`
- [ ] Set up conditional formatting on KPI cards
- [ ] Configure RLS roles
- [ ] Publish to Power BI Service
