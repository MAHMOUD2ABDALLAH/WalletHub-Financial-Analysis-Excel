# WalletHub Case Study - Data Analyst Assessment

## Project Overview

This repository contains my completed solutions for the **WalletHub Data Analyst assessment**, covering three evaluation components: a **financial data research test (DA Test)**, an **Excel functions test**, and a **grammar review test**. Each component was designed to assess the practical skills required for data-driven work: accurate research, mathematical validation, advanced Excel proficiency, and professional written communication.

The project demonstrates the ability to collect and verify financial data from primary sources, apply formulas with mathematical precision, use Excel for data transformation and business logic, and respond constructively to reviewer feedback.

### Business Value

- Verify financial product data directly from official institutional sources
- Validate APY calculations using standard compound interest formulas
- Demonstrate scalable, dynamic Excel solutions over manual data entry
- Show responsiveness to feedback through documented improvements

---

## Key Features

### Task 1 - Financial Product Research (DA Test)
- Comprehensive categorization of OneAZ Credit Union consumer products (banking, savings, lending, and services)
- Confirmation of auto loan availability for new and used vehicles
- Tiered rate documentation for the Premium Checking Account
- Fee structure and restriction analysis for Wells Fargo Way2Save
- Government insurance (FDIC/NCUA) verification across institutions
- Fixed vs. variable rate product distinction (Union Savings Bank)
- Overdraft fee limit documentation (BMO)

### Task 2 - Mathematical APY Validation
- Applied the standard compound interest formula: `APY = (1 + i_nom / N)^N - 1`
- Inputs: `i_nom = 0.0134`, `N = 365`
- Derived result: **1.3490% APY (~1.35%)**
- Full formula and input documentation provided for transparency

### Task 3 - Excel Functions Suite (9 Tasks)
- **Task 1:** `AVERAGE` for dynamic mean calculation
- **Task 2:** AutoFilter for non-destructive record filtering
- **Task 3:** `&` concatenation for phrase construction
- **Task 4:** `COUNTIF` for single-formula duplicate detection
- **Task 5:** PivotTable for aggregated totals across duplicate entries
- **Task 6:** `IF` statements for FHA loan calculations
- **Task 7:** `RIGHT` / `LEN` / `FIND` for dynamic text extraction after "X"
- **Task 8:** `VLOOKUP` for automatic record retrieval based on lookup value
- **Task 9:** `UNIQUE` + `AVERAGEIF` for dynamic unique lists with averages

### Task 4 - Grammar Review (Reviewer Test)
- Three reviews corrected (Wirtz Insurance, Ashford Insurance, Capital One Credit Card)
- Track changes applied to every correction
- Minimal-change editing principle maintained throughout

---

## Technology Stack

| Component | Technology |
|-----------|------------|
| Primary Tool | Microsoft Excel (Office 365) |
| Data Processing | Excel Formulas (VLOOKUP, INDEX/MATCH, COUNTIF, SUMIF) |
| Dynamic Arrays | UNIQUE, FILTER, LET |
| Visualization | PivotTables, AutoFilter, Conditional Formatting |
| Research Sources | Official institutional websites, rate disclosure pages, Wikipedia |
| Documentation | Microsoft Word (track changes), PDF summary reports |

---

## Sheet & File Descriptions

### Excel Functions Test Workbook

| Sheet Name | Purpose | Key Functions |
|------------|---------|---------------|
| Instructions | Test overview and methodology requirements | — |
| Task 1 | Average calculation | `AVERAGE` |
| Task 2 | Record filtering | AutoFilter |
| Task 3 | Phrase concatenation | `&` operator |
| Task 4 | Duplicate detection | `COUNTIF` |
| Task 5 | PivotTable aggregation | PivotTable |
| Task 6 | Conditional loan logic | `IF` |
| Task 7 | Text extraction after delimiter | `RIGHT`, `LEN`, `FIND` |
| Task 8 | Dynamic record retrieval | `VLOOKUP` |
| Task 9 | Unique list + averages | `UNIQUE`, `AVERAGEIF` |

### DA Test & Reviewer Test Documents

| Document | Purpose |
|----------|---------|
| Mahmoud Abdallah_DA Test_07.28.2026.docx | Full responses to all financial research questions |
| DA test summary.pdf | Executive summary of methodology, achievements, and feedback response |
| Mahmoud Abdallah_Reviewer Test_07.28.2026.docx | Grammar corrections with track changes |
| Excel_Functions_Test_Interview_Summary.pdf | One-line rationale for each Excel function used |

---

## Key Formulas Reference

### 1. APY Calculation (Compound Interest)

```
APY = (1 + i_nom / N)^N - 1
```

Inputs: `i_nom = 0.0134`, `N = 365`
Result: `1.3490% APY`

### 2. Average Calculation (Task 1)

```
=AVERAGE(D2:D5)
```

### 3. Duplicate Detection (Task 4)

```
=IF(COUNTIF($B$2:B2, B2) > 1, B2, "N/A")
```

### 4. Text Extraction After "X" (Task 7)

```
=RIGHT(B2, LEN(B2) - FIND("X", B2))
```

### 5. Conditional Loan Logic (Task 6)

```
E3: =IF(C2="FHA", C4*0.0175, 0)
E4: =IF(C2="FHA", C4*1.0175, C4)
E5: =IF(C2="FHA", C3*0.78, 0)
```

### 6. Dynamic Record Retrieval (Task 8)

```
=VLOOKUP($C$11, $B$2:$E$5, 2, FALSE)   'Date Paid
=VLOOKUP($C$11, $B$2:$E$5, 3, FALSE)   'Amount
=VLOOKUP($C$11, $B$2:$E$5, 4, FALSE)   'State
```

### 7. Unique List + Average Price (Task 9)

```
E2: =UNIQUE(A2:A16)
F2: =AVERAGEIF(A2:A16, E2, B2:B16)
```

---

## Research Sources

All financial data was verified from official primary sources:

| Institution | Data Verified |
|-------------|---------------|
| OneAZ Credit Union (oneazcu.com) | Product suite, Premium Checking tiers, auto loans |
| Wikipedia | Compound interest formula verification |
| Capital One (capitalone.com) | Money Market interest rate and APY |
| Wells Fargo (wellsfargo.com) | Way2Save restrictions and requirements |
| First American Bank (firstambank.com) | FDIC-insured account types |
| Union Savings Bank (unionsavings.com) | Fixed-rate savings products |
| BMO (bmo.com) | Overdraft fee limits |

---

## Key Results Summary

| Metric | Result |
|--------|--------|
| APY calculation accuracy | 1.3490% (verified against formula) |
| Excel tasks completed | 9 of 9 |
| Financial institutions researched | 6 |
| Feedback items addressed | 5 of 5 (Q1a, 1c, 3a, 4a, 4b) |
| Grammar reviews corrected | 3 with track changes |

---

## Response to Feedback

**Original feedback from WalletHub HR (Lotika Lawrence):**
> "You did well but there is room for improvement... Please take another look at your solutions for the DA test (for questions 1a, 1c, 3a, 4a and 4b in particular)."

**Improvements made:**

| Question | Original | Improved |
|----------|----------|----------|
| 1a | Brief product list | Full categorization with names, features, rate ranges |
| 1c | Incomplete rates | Exact tiered structure with balance ranges and source URL |
| 3a | Partial answer | Official disclosure document with exact rate and APY |
| 4a | Missing details | All four insured account types plus coverage limits |
| 4b | Vague products | Named fixed-rate products with terms and FDIC confirmation |

---

## Skills Demonstrated

### Research & Analysis
- Web research and source verification
- Financial product knowledge
- Rate structure and fee interpretation
- Cross-referencing across multiple pages

### Technical
- APY calculation and mathematical validation
- Advanced Excel formulas and dynamic arrays
- PivotTable and filter-based reporting
- Data standardization across institutions

### Professional
- Clear written communication (B2+ English)
- Track-changes editing and grammar correction
- Response to structured feedback
- Documentation and transparency

---

## Setup Instructions

### Prerequisites
- Microsoft Excel 2019 or later (Office 365 recommended for `UNIQUE` and dynamic arrays)
- Microsoft Word for reviewing tracked changes
- PDF reader for summary documents

### Viewing the Solutions

1. Clone the repository
   ```
   git clone https://github.com/yourusername/WalletHub-Case-Study.git
   ```

2. Open the Excel workbook
   ```
   start Mahmoud_Abdallah_Excel_Test_07.28.2026.xlsx
   ```

3. Select any task sheet and click a result cell to view the formula in the formula bar

4. Open the DA Test document for written research answers
   ```
   start Mahmoud_Abdallah_DA_Test_07.28.2026.docx
   ```

5. Open the Reviewer Test with track changes enabled
   ```
   start Mahmoud_Abdallah_Reviewer_Test_07.28.2026.docx
   ```

---

## Troubleshooting

| Issue | Cause | Solution |
|-------|-------|----------|
| `#NAME?` on Task 9 | `UNIQUE` not supported | Use Excel 365 or Office 2021+ |
| `#N/A` on Task 8 | Name not found in lookup range | Verify C11 matches an exact name in column B |
| `#VALUE!` on Task 7 | Missing "X" delimiter in string | Confirm all strings contain "X" |
| Track changes not visible | Review mode disabled | Review → Track Changes → All Markups |
| Formula shows as text | Cell formatted as Text | Reformat cell to General and re-enter |

---

## Challenges Overcome

1. **Missing public data** — some rates required accessing disclosure documents rather than public landing pages
2. **Variable vs. fixed rates** — distinguishing rate structures across different institution types
3. **Institutional variation** — standardizing metrics between banks and credit unions
4. **Mathematical precision** — ensuring accurate APY calculation with correct formula application

---

## File Structure

```
WalletHub-Case-Study/
│
├── README.md                                      # Project documentation
│
├── excel/
│   └── Mahmoud_Abdallah_Excel_Test_07.28.2026.xlsx   # 9-task Excel functions test
│
├── da_test/
│   ├── Mahmoud_Abdallah_DA_Test_07.28.2026.docx      # Full research responses
│   └── DA_test_summary.pdf                           # Methodology & achievements summary
│
├── reviewer_test/
│   └── Mahmoud_Abdallah_Reviewer_Test_07.28.2026.docx  # Grammar corrections (track changes)
│
├── docs/
│   ├── Excel_Functions_Test_Interview_Summary.pdf    # Rationale for each function
│   ├── methodology.md                                # Research and validation approach
│   └── sources.md                                    # Verified source URLs
│
└── assets/
    └── Innovative_logo.png                           # Project branding
```

---

## Author

**Mahmoud Abdallah**

### Any Questions
**Mahmoud_Abdallah20@outlook.com**

<img width="350" height="200" alt="WalletHub" src="https://github.com/user-attachments/assets/c8b122a5-40c3-4194-945b-8192119ee162" />
