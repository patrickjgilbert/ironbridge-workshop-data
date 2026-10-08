# The merchant portfolio inbox

Point Cowork at this folder and paste the six prompts below, in order, in one session. Step 1 runs on Sonnet 5.5; switch to Opus 5.5 for step 2 and stay there.

This folder is what Maya Okafor's analyst left her before going on leave: twelve months of ACH activity exports (October 2025 through September 2026) with whatever names the export tool and the analyst gave them, the merchant master, the 2026 pricing schedule, the risk policy memo, three saved emails that matter, one that does not, a Teams chat export, and a blank scan. 22 files, plus this README and `brand-guidelines.md`.

The point of the exercise is that nothing can be identified from its filename and nothing has been calculated. Return rates are not in the exports; they are returns divided by debit count. Revenue is not in the exports; it is the pricing schedule applied to each merchant's tier in the master. Claude has to open every file, build the numbers, and notice what does not add up.

## The prompts

**Step 1, organize. Sonnet 5.5.**

```
This folder is the merchant portfolio inbox my analyst left me and it is a mess. Read every file, then organize it: put monthly activity exports in activity/ named YYYY-MM_ach-activity.csv, the merchant list and pricing in reference/, saved emails and chat exports in correspondence/ named YYYY-MM-DD_short-description with the original extension, and policy documents in policy/. Anything that is a duplicate goes in _review/ under its original name. Anything you cannot identify or date stays where it is. Do not delete anything and do not change what is inside any file. When you are done, summarize here in the thread: a table of every original filename, where it went and how you identified it, then a short list of anything that needs a human. Then interview me, one question at a time, about anything in these files you need me to clarify before we analyze them.
```

**Step 2, the workbook. Switch to Opus 5.5, same session.**

```
Now build ironbridge-portfolio-review.xlsx from the activity exports, the merchant master and the pricing schedule. Rules: do every calculation with formulas or code, never in your head; count May once; every number must trace to a file and a row; use merchant names exactly as the master spells them; if something is missing or can be read two ways, stop and ask me before you build. Four tabs: Summary (TTM debit volume, TTM revenue, portfolio return rates by type, the top-5 merchants by volume with their share), By Merchant (one row per merchant: TTM debit volume, TTM revenue at the tier in the master, Q3 vs Q2 volume change, Sep unauthorized return rate, Sep administrative return rate, overall return rate), Flags (one row per merchant that crosses a line in the Risk Policy memo or shows a problem in the correspondence, with the dollars at stake and the source), and Sources (what you read and what you took from each). Then tell me in plain English the three things I should deal with this week, with the dollars attached.
```

**Step 3, the dashboard. Opus 5.5, same session.**

```
Now turn that workbook into one self-contained HTML dashboard that I can open offline and send to my CEO: a KPI row (TTM debit volume, TTM revenue, blended basis points, portfolio unauthorized return rate), monthly volume and revenue trend, merchant concentration with any merchant over 25% flagged, return rates by merchant against the lines in the Risk Policy memo, and a short note at the top naming the three things to deal with this week. Compute everything from the data; do not hardcode numbers.
```

**Step 4, the skill. Opus 5.5, same session.**

```
Wait. I might have to do this again. Turn what we just did in this session into a skill called merchant-portfolio-review, so next month it runs from one line on a new folder of exports. Put the folder layout, the calculation rules, the risk lines, the tab layout and the dashboard spec inside the skill so it works without me explaining anything. Keep it to one SKILL.md file with the steps and the rules, no scripts. Then read it back to me in five lines.
```

**Step 5, the brand pass. Opus 5.5, same session.**

```
Now reference brand-guidelines.md in this folder and follow those guidelines exactly as written. Rebuild the dashboard so it looks like Ironbridge made it. Do not change a single number. Save it as a new file next to the first one.
```

**Step 6, research. Opus 5.5, same session, web search on.**

```
Search the web for Nacha's current ACH return rate thresholds and the most recent quarterly ACH network volume statistics Nacha has published. Add a Benchmarks section to the bottom of the branded dashboard that shows our portfolio return rates next to the Nacha thresholds and our TTM debit volume growth next to the network's growth, and cite each source with a link. Do not change any other number on the page.
```

## The files

| File | What it really is |
|---|---|
| `ach_activity_2025-10.csv` | October 2025 activity export |
| `ach_activity_2025-11.csv` | November 2025 |
| `export (1).csv` | December 2025 |
| `export (2).csv` | January 2026 |
| `ach-activity-feb.csv` | February 2026 |
| `ach activity march 2026.csv` | March 2026 |
| `activity_april_FINAL.csv` | April 2026 |
| `activity_may_v2.csv` | May 2026 |
| `Copy of activity_may_v2.csv` | May 2026 again, byte-for-byte identical |
| `june.csv` | June 2026 |
| `ach_activity_2026-07.csv` | July 2026 |
| `aug export.csv` | August 2026 |
| `Sept_export (1).csv` | September 2026 |
| `merchant master.xlsx` | 24 merchants: id, name, vertical, onboarded, account manager, pricing tier, renewal, status |
| `Pricing Schedule 2026.pdf` | Three tiers effective Jan 1 2026, credits, cards, and how a monthly fee is calculated |
| `Risk Policy memo.docx` | Nacha thresholds, Ironbridge's internal warning lines, what happens at each, the 25% concentration rule |
| `RE_ return rate notice.txt` | Summit Ridge Lending's COO disputing the September return-rate notice (Sep 24 2026) |
| `FW_ pricing amendment - Harrow Finch.txt` | Countersigned amendment moving Harrow & Finch from Tier B to Tier C effective Apr 1 2026 (Mar 27 2026) |
| `Northgate - processor review.txt` | Northgate BNPL is running a processor review, decision by Nov 15 2026 (Sep 18 2026) |
| `lunch order thursday.txt` | Eight sandwich orders. Not work. No date. |
| `teams-chat-merchant-ops-sept.txt` | Merchant Ops channel, September 2026, the four account managers and the risk analyst |
| `scan_0097.pdf` | A blank page |
| `brand-guidelines.md` | Ironbridge's brand, for step 5 |

Every export has the same header row:

```
month,merchant_id,merchant_name,vertical,debit_count,debit_volume,credit_count,credit_volume,returns_total,returns_unauthorized,returns_administrative,returns_nsf,card_volume
```

One row per merchant per month. `returns_unauthorized` is R05, R07, R10, R11 and R29; `returns_administrative` is R02, R03 and R04; `returns_nsf` is R01 and R09; `returns_total` includes a small remainder of other codes. A return rate is the return count divided by `debit_count` for the same month. Monthly revenue for a merchant is `debit_count` times the per-item fee for its tier, plus `debit_volume` times the tier's basis points, plus `credit_count` times $0.10, plus `card_volume` times 22 basis points.

## What's planted

Check Claude's summary in the thread, the workbook and the dashboard against this list. TTM is October 2025 through September 2026, counting May once. Q2 is April to June 2026, Q3 is July to September 2026.

| Finding | The fact | Where | What a correct run does with it |
|---|---|---|---|
| Summit Ridge Lending is over the unauthorized line | Unauthorized return rate **0.31%** in June, **0.44%** in July, **0.52%** in August, **0.61%** in September 2026. Over Ironbridge's 0.40% internal line three months running, over Nacha's 0.50% two months running. Administrative and overall rates stay normal. Every other merchant is under 0.40% all year. | the exports, divided out; the Risk Policy memo; the Sep 24 email; the Teams chat | Flags tab row for Summit Ridge with the four rates and the policy consequence (second month over the network threshold starts a termination review). Top of the "three things this week". The COO's reply says they changed debit timing in June, which is where the climb starts. |
| Northgate BNPL is more than a quarter of the book and may leave | **28.4%** of TTM debit volume ($588,304,698.74); the next largest merchant, Blue Heron Consumer Finance, is 10.6%. Northgate's TTM revenue ($584,194.01) is the single largest line. The Sep 18 email: processor review against two other providers, decision by Nov 15, best-and-final pricing and a return-handling SLA due Oct 16. | the exports; the master (Tier A, renewal Dec 31 2026); `Northgate - processor review.txt` | Flagged on the concentration chart (over the memo's 25% line) and on the Flags tab with the revenue at stake. In the three things this week, with the Oct 16 date. |
| Harrow & Finch Receivables is billed at the wrong tier | The master says Tier B. The countersigned amendment moves them to Tier C effective April 1 2026. Billed at B for April through September, they were under-billed by **$38,496.55** (for each month, `debit_count` x $0.03 plus `debit_volume` x 4 bps). Their volume has been under $5M a month since March. | `FW_ pricing amendment - Harrow Finch.txt`; `merchant master.xlsx`; the exports | Flags tab row with the dollar figure. In the three things this week. The Teams chat has someone asking whether the new pricing "ever made it into the master" and nobody answering. A careful run asks in the step 1 interview which tier to use. |
| May is in the folder twice | `Copy of activity_may_v2.csv` is byte-identical to `activity_may_v2.csv`. Counting May twice inflates TTM volume by about 8.4%. | the two files; the Teams chat ("pulled May twice when the report timed out, one of them is just a copy") | Step 1 moves the copy to `_review/` and says why. Step 2 counts May once. TTM debit volume should be **$2,071,495,418.10** and TTM revenue at the tiers in the master **$2,954,335.45**. |
| Cedar Falls Collections is leaving | Q3 debit volume ($5,989,590.07) is **38%** under Q2 ($9,660,629.15), falling every month July, August, September. Status in the master is `Notice given`. There is no email about it; the Teams chat says they "went quiet since July". | the exports; the master; the Teams chat | A By Merchant row showing the Q3 vs Q2 drop, a Flags row citing the status and the chat. Not one of the three things this week (the decision has already been made), but it should be named. |
| Administrative returns are drifting up | The portfolio administrative return rate goes from **1.9%** in October 2025 to **2.6%** in September 2026. Under Nacha's 3.0% every month. The memo's 2.5% internal line applies to each merchant, and by September a number of merchants sit between 2.5% and 3.0% while none crosses Nacha's line. | the exports, divided out; the Teams chat (the risk analyst mentions it) | Named as a trend in the Summary or dashboard, not a top-three item. A run that lists every merchant over 2.5% in September on the Flags tab is not wrong; a run that calls it the week's priority has missed the point. |
| Two files need a human | `scan_0097.pdf` is a blank page with no text. `lunch order thursday.txt` is sandwich orders with no date. | the two files | Both stay where they are and appear under "Needs a human" in the summary Claude posts in the thread. |

Other things a good run notices: Redfern Apparel onboarded August 11 2026 and has only two months of data (status `Onboarding`), so it should not be judged on Q3 vs Q2. Three merchants renew October 31 2026 (Pinecrest, Stonebridge, Brightwater). The portfolio unauthorized rate in September is well under 0.40% even with Summit Ridge in it, which is why the merchant-level view matters.

What a correct step 1 leaves behind:

```
activity/2025-10_ach-activity.csv ... activity/2026-09_ach-activity.csv   12 files
reference/merchant master.xlsx, reference/Pricing Schedule 2026.pdf
correspondence/2026-03-27_harrow-finch-pricing-amendment.txt
correspondence/2026-09-18_northgate-processor-review.txt
correspondence/2026-09-24_summit-ridge-return-rate-reply.txt
correspondence/2026-09-30_teams-merchant-ops-september.txt   (or dated 2026-10-06, the export date)
policy/Risk Policy memo.docx
_review/Copy of activity_may_v2.csv
scan_0097.pdf, lunch order thursday.txt                      left in place
README.md, brand-guidelines.md
```

The short descriptions in the correspondence names will vary. The dates should not.

## Things to watch for in Claude's answer

- Did it compute return rates from the counts, or look for a rate column that is not there?
- Did it count May once? TTM volume near $2.24B means it did not.
- Did it use the tier in the master for revenue, and then separately flag that the master is wrong for Harrow & Finch? Both halves matter.
- Did it ask about the Harrow & Finch tier in the step 1 interview, or discover it in step 2?
- Did the Flags tab carry dollars (Northgate's TTM revenue, Harrow & Finch's $38,496.55) and sources, or just names?
- Did the brand pass change any number? It should not have.
