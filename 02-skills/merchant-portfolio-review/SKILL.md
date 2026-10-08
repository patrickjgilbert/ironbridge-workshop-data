---
name: merchant-portfolio-review
description: Monthly review of an Ironbridge ACH merchant portfolio from a folder of raw activity exports. Organizes the inbox, builds ironbridge-portfolio-review.xlsx (Summary, By Merchant, Flags, Sources) with live formulas, builds a self-contained branded ironbridge-dashboard.html, and names the three things to deal with this week. Use when pointed at a new folder of monthly ACH activity exports plus the merchant master, pricing schedule and risk policy memo, or when asked to "run the portfolio review", "review the merchant inbox", or "rebuild the portfolio dashboard".
---

# Merchant portfolio review

One line to run it: "Run the merchant portfolio review on this folder." Everything below is the contract. Do not ask the user to explain the files; ask only the interview questions in step 2.

## 1. Read everything first

Open every file in the folder before moving anything. Filenames are not trustworthy; the contents are. Identify each file by what is inside it:

- **Activity export**: CSV with the header `month,merchant_id,merchant_name,vertical,debit_count,debit_volume,credit_count,credit_volume,returns_total,returns_unauthorized,returns_administrative,returns_nsf,card_volume`. One row per merchant per month. Date it by the `month` column, never the filename.
- **Merchant master**: xlsx with merchant_id, merchant_name, vertical, onboarded, account_manager, pricing_tier, contract_renewal, status.
- **Pricing schedule**: PDF with tiers (per debit item, basis points on debit volume, volume band), credit per-item price, card basis points, and the fee formula. Note its effective date.
- **Risk policy memo**: docx with network thresholds, internal warning lines, the action ladder and the concentration rule.
- **Correspondence**: saved emails (`From:`/`Date:`/`Subject:` headers) and chat exports. Date an email by its Date header, a chat export by its last message.
- **Noise**: anything with no merchant content or no date (lunch orders, blank scans). Leave it.

Check duplicates by hash, not by name. Two files with the same MD5 are one file.

## 2. Organize (do not delete, do not edit any file's contents)

```
activity/YYYY-MM_ach-activity.csv        one per month, named from the month column
reference/<merchant master>.xlsx, reference/<pricing schedule>.pdf
correspondence/YYYY-MM-DD_short-description.<original ext>
policy/<risk policy memo>.docx
_review/<original name>                  duplicates, byte-identical copies, superseded re-pulls
(root)                                   anything unidentifiable or undated stays put
```

Write `WHAT-I-FOUND.md`: one table (original filename, where it went, how identified) and a **Needs a human** section listing the undated/unidentifiable files, every place a reference file disagrees with correspondence, and anything in the chat that nobody answered.

Then interview the user one question at a time. Always ask these, in this order, unless the files already settle them:
1. Confirm that any duplicate month is counted once.
2. Any merchant whose tier in the master conflicts with a signed amendment: which tier to bill in the workbook. Default: workbook revenue at the master tier, flag the difference at the amendment tier with the dollars.
3. If the pricing schedule's effective date falls inside the trailing twelve months and the prior schedule is absent: confirm the current rates are applied to the earlier months.
4. Any merchant with status other than Active (Notice given, Onboarding, Suspended): confirm how to treat it.
5. Any merchant with fewer than three months of data: confirm it is excluded from quarter-on-quarter judgments.

## 3. Calculation rules

- Never compute in your head. Every number is a spreadsheet formula over an embedded data block, or code over the files, and traces to a file and a row.
- Trailing twelve months (TTM) = the twelve distinct months in `activity/`, each counted once. Q2 = Apr-Jun of the latest year, Q3 = Jul-Sep. Adjust the quarter definitions if the twelve months end in a different month.
- Use merchant names exactly as the master spells them. Join exports to the master on merchant_name (or merchant_id, and confirm they agree).
- **Return rates**: the exports carry no rate column. Each rate = returns of that type / `debit_count`, same merchant, same month. Unauthorized = `returns_unauthorized` (R05, R07, R10, R11, R29). Administrative = `returns_administrative` (R02, R03, R04). Overall = `returns_total`. Portfolio rate = sum of returns / sum of debit_count.
- **Monthly revenue per merchant** = `debit_count` x tier per-item price + `debit_volume` x tier bps / 10000 + `credit_count` x credit per-item price + `card_volume` x card bps / 10000. Tier comes from the master. Prices come from the pricing schedule. Read all four from the files; do not assume the 2026 values hold.
- **Blended basis points** = TTM revenue / TTM debit volume x 10000.
- **Share** = merchant TTM debit volume / portfolio TTM debit volume.
- **Q3 vs Q2** = Q3 debit volume / Q2 debit volume - 1. Blank (not zero) when Q2 is empty.
- **Billing error dollars** = for each affected month, `debit_count` x (correct per-item - billed per-item) + `debit_volume` x (correct bps - billed bps) / 10000, from the amendment's effective date to the latest month.

## 4. Risk lines (read them from the memo each time; these are the 2026 values for checking)

| Measure | Ironbridge internal line | Nacha network threshold |
|---|---|---|
| Unauthorized return rate | 0.40% | 0.5% |
| Administrative return rate | 2.5% | 3.0% |
| Overall return rate | 12% | 15% |

Applied to each merchant, each month. Action ladder: one month over an internal line = warning letter within 5 business days of close; two consecutive months = written remediation plan due 10 business days after the letter; over a network threshold = formal notice plus a 5% rolling reserve on monthly debit volume held 90 days; two consecutive months over a network threshold = CRO termination review, decision within 10 business days of the second close. Concentration: any merchant over 25% of TTM debit volume is reported to the board with renewal date, return rates and retention steps (not a breach). A portfolio rate drifting toward a line over several months is a board-report trend even when no merchant crosses it.

## 5. Workbook: `ironbridge-portfolio-review.xlsx`, exactly four tabs

- **Summary**: TTM debit volume, TTM revenue, TTM debit count, blended bps, Q3 vs Q2 portfolio change; portfolio return rates by type (TTM, latest month, first month, internal line, Nacha threshold); top-5 merchants by TTM volume with share, revenue and a 25% test; a monthly trend table (volume, revenue, three rates, count of merchants over each internal line); a one-line note on any drifting portfolio rate.
- **By Merchant**: one row per merchant in the master: tier, status, renewal, account manager, TTM debit volume, share, TTM revenue at the master tier, Q2 volume, Q3 volume, Q3 vs Q2, latest-month unauthorized rate, latest-month administrative rate, overall TTM return rate, note. A portfolio total row.
- **Flags**: one row per merchant that crosses a Risk Policy line or shows a problem in the correspondence. Columns: priority, merchant, flag, line or rule, evidence (computed, as a formula string), dollars at stake (formula), what the dollars are, policy consequence / action, source files. Priority order: unauthorized over a line, concentration over 25% plus any retention risk, billing errors, merchants leaving, then trends and watch items. A second block lists every merchant at or over the administrative internal line in the latest month with its last three rates.
- **Sources**: the file list with what was taken from each, then the pricing block, the master block, and the full activity data block (all rows, with helper columns for tier, monthly revenue and quarter). Every formula on the other tabs points here.

Build it in one pass with openpyxl (`uv run --with openpyxl`). Do not test it unless asked; the user opens it.

Then say, in plain English, the three things to deal with this week with the dollars attached, and name anything that matters but is not a this-week item (a merchant already leaving, a trend).

## 6. Dashboard: `ironbridge-dashboard.html` (and `ironbridge-dashboard-branded.html`)

Self-contained: embed the activity rows, master, pricing and policy lines as JSON in the page and compute everything in JavaScript. No hardcoded figures, no external scripts; the only permitted external link is Google Fonts. Never name a top-level JavaScript variable `top`.

Sections in order: a short note at the top with the three things for the week; a KPI row (TTM debit volume, TTM revenue, blended bps, portfolio unauthorized return rate); monthly debit volume and monthly revenue trend; merchant concentration as horizontal bars sorted largest first with the 25% line drawn and any merchant over it flagged; return rates by merchant for the latest month against the internal and Nacha lines; the unauthorized-rate series for any merchant over a line with the two lines drawn; a monthly portfolio rate table. A source line under every chart naming the file.

For the branded version follow `brand-guidelines.md` exactly: palette tokens (Forge Green structure, Verdigris as the single data color, Amber Lamp only on flagged items, Pewter for everything else, Limestone page, Paper cards, River Slate secondary text, Hairline rules), Playfair Display / Inter / JetBrains Mono with fallback stacks, left-aligned 12-column layout, square flat cards, 4px Forge Green top rule, wordmark with the truss glyph, headlines that state the finding as a sentence, number formats ($2.07B / $38,496.55 / 0.52% / 28.4% / 11 bps / Sep 2026), no pies, no gradients, no legends when direct labels work. Same numbers as the generic dashboard.

## 7. Benchmarks (only when web search is on)

Append a Benchmarks section to the branded dashboard: portfolio unauthorized, administrative and overall rates next to Nacha's thresholds, and TTM debit volume growth (latest quarter vs the same quarter a year earlier, from the exports) next to the most recent quarterly network growth Nacha has published. Cite each source with a link. Change no other number on the page.

## 8. Checks before handing over

- TTM volume counts each month once (a duplicate month inflates it by roughly one twelfth).
- Revenue uses the master tier, and any amendment that disagrees with the master is a Flags row with dollars.
- Every Flags row carries dollars and a source file.
- Return rates came from counts divided out, not from a column that does not exist.
- The brand pass changed no number.
