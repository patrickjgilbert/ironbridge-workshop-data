# Ironbridge Payments brand guidelines

For anything Ironbridge puts in front of a merchant, the board, a bank partner or our own team: dashboards, reports, notices, proposals. One page. If something is not covered here, keep it plain and use Forge Green on Limestone.

Ironbridge is a Dayton, Ohio company that moves other people's money. The look is the old iron truss bridge over the Great Miami River: dark green paint, grey river stone, a lot of straight lines, and one amber warning lamp at the end of the span. Serious, settled, nothing shiny.

## Palette

| Token | Hex | Role |
|---|---|---|
| Forge Green | `#1E4D3B` | All headings, body text, axis lines, the wordmark. The structure of the page. |
| Verdigris | `#2E8B6F` | The single data color. The series or bar that carries the point. Nothing decorative is ever Verdigris. |
| Amber Lamp | `#D98A14` | Flags only: a merchant over a return-rate line, a merchant over 25% of volume, a billing error. If nothing is wrong, there is no amber on the page. |
| Limestone | `#F3F1EC` | Page background. |
| Paper | `#FFFFFF` | Card and table background. |
| River Slate | `#4A5A66` | Secondary text: labels, captions, axis text, footnotes, source lines. |
| Pewter | `#B8BFC4` | Every data series that is not the one that matters. |
| Hairline | `#E4E7E9` | Card borders, table rules and gridlines. |

Area fills may use Verdigris at 15% opacity. No other tints, no gradients.

## Type

| Use | Font | Fallback stack | Size and weight |
|---|---|---|---|
| H1 (page title) | Playfair Display | Georgia, "Times New Roman", serif | 32px, 600 |
| H2 (section and chart headlines) | Playfair Display | Georgia, "Times New Roman", serif | 20px, 600 |
| Body and notes | Inter | -apple-system, "Segoe UI", Helvetica, Arial, sans-serif | 16px, 400, line height 1.6 |
| Labels (KPI names, axis titles, table headers) | Inter | same as body | 12px, 600, uppercase, letter spacing 0.06em, River Slate |
| Numbers (KPIs, table figures, chart labels) | JetBrains Mono | ui-monospace, "SF Mono", Menlo, Consolas, monospace | KPI values 32px, 500. Table and chart numbers 13 to 14px, 400 |

The three fonts may be linked from Google Fonts:

```
https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600&family=Inter:wght@400;600&family=JetBrains+Mono:wght@400;500&display=swap
```

Always include the fallback stack, so a file opened offline still reads correctly.

## Layout

- Left-aligned. Headlines, text and KPI values all start on the left edge. Nothing centered except a lone number inside a small card.
- 12-column grid, 24px gutters, max content width 1200px, 32px outer padding. The KPI row is 4 cards of 3 columns each.
- Generous whitespace: 48px between sections, 24px inside cards.
- Cards: Paper background, 1px Hairline border, no shadow. Corner radius 0. Square corners, like the bridge.
- A 4px Forge Green rule across the top of the page, full width, above the wordmark.
- No icons beside KPIs, no emoji, no background images.

## Charts

- One accent color. The series that carries the point is Verdigris; everything else is Pewter.
- The flagged item (the merchant over 25%, the month a rate crossed a line) is Amber Lamp, and it is the only amber on the chart.
- Return-rate lines from the risk policy are drawn as 1px dashed River Slate rules with the value labeled at the right end, for example "Nacha 0.50%" and "Ironbridge 0.40%".
- Label lines and bars directly at their ends. Use a legend only when direct labels would overlap.
- No pie or donut charts. Shares are horizontal bars, sorted largest first.
- No gradients, no 3D, no drop shadows, no animation.
- Gridlines: horizontal only, 1px Hairline. No chart border.
- Bar axes start at zero. Line charts may start above zero if the axis says so.
- Lines are 2px. Mark only the last point, with its value.
- Mix over time is a stacked bar per month in Verdigris, Pewter and River Slate, not a stacked area rainbow.

## Numbers

- Large money: `$2.07B`, `$588M`, `$38.5K`. Table money: `$38,496.55`, thousands separators always.
- Return rates to two decimals: `0.52%`, `0.40%`. Shares and changes to one decimal: `28.4%`, `-38.0%`.
- Basis points as `11 bps`, never `0.11%`.
- Numbers right-aligned in tables, set in JetBrains Mono with tabular figures (`font-variant-numeric: tabular-nums`) so columns line up.
- Months as `Sep 2026` on screen. Never show `2026-09` or a raw date to a reader.

## Headline voice

Every section and chart headline states the finding as a sentence. A reader who reads only the headlines should get the story.

- Yes: "Summit Ridge has been over the unauthorized line since July"
- Yes: "Northgate is 28.4% of debit volume"
- No: "Return Rate Trend", "Merchant Concentration", "Portfolio Overview"

Sentence case. No exclamation marks. No "Insights" or "Key Takeaways" labels.

## Wordmark

`IRONBRIDGE` in Playfair Display 600, all caps, letter spacing 0.12em, Forge Green, with a small truss glyph to its left: three short vertical strokes joined by a diagonal, drawn in inline SVG, 18 to 20px, Forge Green. Under it, in Inter 12px River Slate: "Payments, Dayton, Ohio". No image files needed.

## Do / Don't

| Do | Don't |
|---|---|
| One Verdigris series per chart, everything else Pewter | A different color for every merchant or vertical |
| Amber Lamp on the one thing that needs action | Amber as decoration, or red and green traffic lights |
| Headlines that state the finding | Headlines that name the chart type |
| 1px borders, square corners, flat Paper cards | Rounded cards with shadows, glass effects, gradients |
| Direct labels at the end of a line or bar | Legends in a box below the chart |
| Horizontal bars for shares | Pie or donut charts |
| A source line under each chart in River Slate 12px naming the file | Unlabeled numbers with no file behind them |
