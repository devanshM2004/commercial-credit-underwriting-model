# Commercial Credit Underwriting Model

An employer-facing portfolio project that underwrites a sample commercial loan
request end-to-end — from financial spreading through DSCR, collateral, risk
rating, and a written credit memo recommendation. Built as a single, clean,
interview-friendly Excel workbook generated from Python (`openpyxl`).

## The Case

| | |
|---|---|
| **Borrower** | Northstar Equipment Services LLC |
| **Industry** | Commercial equipment repair & maintenance |
| **Request** | $750,000 senior secured term loan |
| **Purpose** | Purchase service vehicles and shop equipment |
| **Term / Rate** | 60 months / 8.25%, fully amortizing |
| **Relationship** | New borrower |
| **Question** | *Should the bank approve the loan?* |

## What the Workbook Contains

`Commercial_Credit_Underwriting_Model.xlsx` has eight tabs that flow into one
another:

1. **01 Borrower Overview** — borrower facts and the proposed loan terms (the
   loan amount, term and rate live here as inputs and drive the rest of the model).
2. **02 Income Statement** — revenue, COGS, gross profit, operating expenses,
   EBITDA, D&A, interest, taxes and net income for 2023–2025.
3. **03 Balance Sheet** — cash, AR, inventory, fixed assets, payables, short-
   and long-term debt, equity, plus an automatic balance check.
4. **04 Ratio Analysis** — revenue growth, gross & EBITDA margins, current
   ratio, debt-to-equity, debt-to-EBITDA and working capital, each with a plain-
   English explanation.
5. **05 DSCR Analysis** — proposed-loan annual debt service (Excel `PMT`),
   cash flow available for debt service (CFADS), and both proposed-loan and
   global DSCR vs. a policy minimum.
6. **06 Collateral LTV** — vehicles and equipment with advance rates, lendable
   value, LTV and collateral-coverage ratio.
7. **07 Risk Rating** — a weighted 1–5 scorecard (repayment capacity, leverage,
   liquidity, collateral, management, industry, trends) mapped to a risk grade.
8. **08 Credit Memo** — recommendation, strengths, weaknesses, risks,
   mitigants and approval conditions; key metrics pull live from the other tabs.

### Conventions
- **Blue font = input/assumption cells** you can change.
- **Black font = formulas/outputs** driven by those inputs.
- Live Excel formulas are used throughout so a reviewer can audit and flex the
  model. No macros.

## Headline Result

**Recommendation: Approve with conditions.**

- 2025 **global DSCR ≈ 1.33x** (and proposed-loan DSCR ≈ 2.6x) vs. a 1.20x floor.
- **Collateral coverage ≈ 1.09x** of lendable value; LTV ≈ 92% on discounted value.
- Modest leverage (Debt/EBITDA < 2x, D/E ≈ 1.0x), strong liquidity (current
  ratio > 2x), and three years of steady growth with stable ~39% margins.

## Banking / Underwriting Relevance

This model demonstrates the core skills of a commercial credit analyst:

- **Financial spreading** — normalizing three years of statements into
  comparable, ratio-ready form.
- **Repayment capacity (DSCR)** — the single most important question in C&I
  lending: does cash flow cover debt service with cushion?
- **Leverage & liquidity analysis** — how much debt the business carries and
  whether it can meet near-term obligations.
- **Collateral / LTV** — sizing the secondary (recovery) source of repayment
  using realistic liquidation haircuts.
- **Risk rating** — translating qualitative and quantitative factors into an
  internal grade that drives pricing and approval authority.
- **Credit memo writing** — communicating a defensible recommendation with
  conditions, the deliverable a credit committee actually reads.

## How to Build

```bash
pip install -r requirements.txt
python build_workbook.py
```

This regenerates `Commercial_Credit_Underwriting_Model.xlsx` from scratch.

## Disclaimer

All figures are **synthetic** and created for illustration. This is an
educational portfolio project, not a real credit decision or financial advice.
