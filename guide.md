# MICS Equity Research Platform: Features and Methodology

This guide explains every feature of the MICS screener and deep analysis app, with the exact rules and formulas the code uses. Use the index to jump to a topic; section numbers match the headings below.

## Index

| No. | Section | Look here for |
| --- | --- | --- |
| 1 | The platform at a glance | The three parts, how they connect, the nightly cycle |
| 2 | Data and the nightly workflow | TradingView export, Dataroma scrape, file cleanup, ISIN memory, renamed columns |
| 3 | Screener: sectors and markets | GICS mapping, REITs, unclassified companies, market and search filters |
| 4 | Screener: rule sets | Sector-tuned bars, the same-for-all set, which rules apply to which sector, data coverage |
| 5 | Screener: criteria reference | All 15 criteria with data field, direction and formula |
| 6 | Screener: strictness slider | How tightening and loosening move every bar |
| 7 | Screener: momentum stages | The six early-signal conditions, the five stages, the momentum score |
| 8 | Screener: results and exports | Summary, table, Excel, PowerPoint, CSV, saved settings |
| 9 | Deep analysis: company page | Every card on the page and the formulas behind it |
| 10 | Deep analysis: verdict scorecard | Pillars, scoring formula, Buy, Hold or Sell, red flags |
| 11 | Deep analysis: index page | The 70 indexes, local and US dollar returns, fund data |
| 12 | Deep analysis: exports | Excel, PowerPoint, printable page, the opinion switch |
| 13 | Settings reference | Every setting you can change, its default and its file |
| 14 | Known limits | What the numbers cannot tell you |
| 15 | Glossary | Plain meanings of the terms used |

## 1. The platform at a glance

The platform has three parts: a nightly build that screens about 30,000 listed companies, a web screener to explore the result, and a deep analysis app for one company or one index at a time.

| Part | Where it lives | What it does |
| --- | --- | --- |
| Nightly build | GitHub repo wealthmics/Stock-Screener, file marquee_screener.py, run by the workflow build-screener.yml | Reads the TradingView export and the Dataroma superinvestor list, screens every company, writes the workbook, the web screener and the data feed for the deep analysis app |
| Web screener | GitHub Pages site of the same repo, index.html | Interactive screening by sector, market, rules and momentum, with Excel, PowerPoint and CSV export |
| Deep analysis app | Streamlit app from repo wealthmics/mics-deep-analysis | Full analysis of one company or one index, with a verdict and Excel, PowerPoint and HTML export |

### The nightly cycle

The workflow runs every night at 18:30 UTC, which is 00:00 IST and 22:30 UAE time, and can also be started by hand from the Actions tab.

1. Scrape the Dataroma superinvestor holdings, unless a file dated today already exists.
2. Delete every older Dataroma file, so only one remains in the data folder.
3. Read the newest TradingView export in the data folder, whatever its file name.
4. Fill in missing ISINs from the ISIN memory and handle renamed columns.
5. Match the superinvestor holdings to companies in the export.
6. Compute the sector-tuned bars from the day's data, then screen every company.
7. Write the Excel workbook, the web screener (index.html), stocks.json and sector_stats.json.
8. Commit everything back to the repo. GitHub Pages serves the new screener. The deep analysis app picks up the new company feed within 15 minutes and the new sector statistics within an hour.

The deep analysis app fetches its company and index data live from Yahoo when a page is opened, and combines it with the nightly TradingView feed.

## 2. Data and the nightly workflow

Two files feed everything: a TradingView screener export of all listed companies, and a list of superinvestor holdings scraped from Dataroma. File names never matter; the build recognises each file by its columns.

### 2.1 The TradingView export

Put the export (a CSV) in the data folder of the Stock-Screener repo. A file counts as the company export when its column names contain at least 3 of: market capitalization, price to earnings, return on equity, net margin, country or region, sector. If several qualify, the one with the latest date in its file name wins; the file's modified time only breaks ties.

The build reads ratios, scores, returns, prices and averages from this file. Section 5 lists every column a screening rule uses. Four optional columns switch on extra momentum signals when present: Performance 1 month, Performance 3 months, 52 week high and Relative volume.

### 2.2 Dataroma superinvestor holdings

The build scrapes dataroma.com with a headless Chrome browser, page by page (up to 15 pages, 3 tries per page, stopping at the first empty page). A scrape returning fewer than 100 holdings counts as failed.

- When it scrapes: on every run, unless a Dataroma file dated today already exists (setting `SCRAPE_DATAROMA = 'auto'`, `DATAROMA_MAX_AGE_DAYS = 0`).
- File name: `dataroma_portfolio_YYYY-MM-DD.csv`, dated the day of the scrape. That date is the superinvestor date shown everywhere.
- Only one file kept: after choosing the newest Dataroma file, the build deletes every other Dataroma file in the data folder (`KEEP_ONLY_LATEST_DATAROMA = True`). The workflow commits the deletion.
- If the scrape fails: the build keeps the newest existing file and shows that file's own date. Nothing is deleted.

### 2.3 Matching holdings to companies

Each Dataroma holding is matched to a company in the export in this order. Fuzzy name matching is deliberately not used, because short legal names produce false matches.

1. Verified alias: a hand-checked list for ADRs and other share classes, for example TSM to 2330 in Taiwan.
2. Exact ticker. If the ticker exists on several exchanges, the listing priced within 2% of the Dataroma quote wins, then the US listing, then the largest market cap.
3. Ticker variant: a legacy -OLD suffix removed, or dot and underscore swapped.
4. Company name, after removing legal suffixes such as Inc, Ltd, Holdings and class labels.

A match is High confidence when an exact ticker, alias or name match agrees on price within 3%, or when an exact ticker matches a US dollar listing. Everything else is Medium, including every ticker variant match. Unmatched holdings go to the Unidentified sheet with a reason: an ETF or fund, a legacy -OLD ticker, a warrant or unit, or a company not in the export.

### 2.4 ISIN memory

The deep analysis app is keyed by ISIN, and TradingView exports do not always include an ISIN column. The build therefore remembers every ISIN it sees in `data/isin_map.json`, keyed by ticker and country.

- It learns from any export that carries ISINs, and from the last published stocks.json.
- Numeric tickers are normalised, so 005930 and 5930.0 are the same company.
- On an export without ISINs, every remembered ISIN is filled back in. A company never seen with an ISIN still goes into the feed, under a key built from its ticker and currency (TV:TICKER:CCY).

### 2.5 Renamed or missing columns

TradingView sometimes renames columns. For each column the build needs, it looks for, in order:

1. The exact name.
2. The same name with different spacing or capitals.
3. A known synonym, for example Ticker for Symbol, or Name for Description.
4. The same measure for a different reporting period, preferring trailing twelve months, then annual, then quarterly. Return periods such as 1 month or year to date are never swapped for each other.

The run log lists every substitution. If an essential column is missing (Symbol, Description, Country, Sector, Industry, Price, Market capitalization), the build stops with a clear message. Any other missing measure is left blank and its rule is switched off for that run.

### 2.6 Files produced

| File | What it holds | Used by |
| --- | --- | --- |
| Marquee_Investor_Portfolio_Consolidated.xlsx | Read Me, Marquee Portfolio, Unidentified, Sector Summary, All Stocks, Screen Settings and one screen sheet per sector plus All Sectors | You, and the screener's Full screener workbook button |
| index.html | The web screener, with all company data inside | GitHub Pages |
| stocks.json | Every screened or superinvestor-held company with its TradingView figures, superinvestor data and the period of each figure | Deep analysis app |
| sector_stats.json | Lower quartile, median and upper quartile of 15 ratios for each GICS sector, across the whole universe | Deep analysis app |
| data/isin_map.json | The ISIN memory | The build itself |

## 3. Screener: sectors and markets

Every company is placed in one of the 11 GICS sectors by folding TradingView's sectors together; companies that fit none are shown as Unclassified.

### 3.1 How TradingView sectors map to GICS

| GICS sector | TradingView sectors folded into it |
| --- | --- |
| Information Technology | Electronic technology, Technology services |
| Communication Services | Communications |
| Consumer Discretionary | Consumer durables, Consumer services, Retail trade |
| Consumer Staples | Consumer non-durables, Distribution services |
| Health Care | Health technology, Health services |
| Financials | Finance |
| Energy | Energy minerals |
| Materials | Non-energy minerals, Process industries |
| Industrials | Producer manufacturing, Industrial services, Commercial services, Transportation |
| Utilities | Utilities |
| Real Estate | None directly; see 3.2 |

### 3.2 Real estate and REITs

TradingView has no real estate sector. Any company whose industry name contains "real estate" or "REIT" (any capitals) is moved to Real Estate, whatever its TradingView sector. Most come out of Finance.

### 3.3 Unclassified companies

About 800 companies sit in TradingView's Miscellaneous or Government sectors, which map to no GICS sector. In the web screener they appear as Unclassified, a separate tick under the sector list. They are judged on the All Sectors rules (section 4).

### 3.4 Filters in the sidebar

| Filter | What it does |
| --- | --- |
| Sectors | Tick any mix of the 11 sectors and Unclassified. The number beside each is how many companies the sector holds in total. Select all and Clear are at the top. |
| Listed in the US | Keeps companies whose country of registration is the United States. |
| Listed outside the US | Keeps every other company. Untick both and nothing shows. |
| Only companies superinvestors hold | Keeps only companies matched to a Dataroma holding (marked with a gold star in the table). |
| Search | Keeps companies whose name or ticker contains the text typed. |

The numbers beside the market ticks count the companies in the ticked sectors, before any rule is applied.

## 4. Screener: rule sets

By default every company must beat the median company of its own sector on each measure that fits its business. On the 21 October 2026 data that leaves 1,611 of 30,000 companies.

### 4.1 The two threshold sets

| Set | Button | How the bars are set |
| --- | --- | --- |
| Sector-tuned (default) | Sector-tuned | Computed from the data on every run: each bar is the 50th percentile (the median) of that sector, using every company in the sector, not only those that pass. Each sector uses only the rules that fit it (4.3). |
| Same for all | Same for all | One fixed set of bars for every sector (4.4). Like for like, but banks and utilities rarely pass it. |

### 4.2 How a sector-tuned bar is calculated

For a rule where higher is better, the bar is the value at percentile t of the sector. For a rule where lower is better (P/E, debt to equity), it is the value at percentile 100 minus t, so the company still has to be better than t% of its sector.

```latex
\text{higher is better:}\quad \text{bar} = Q_{\text{sector}}\left(\tfrac{t}{100}\right) \qquad \text{lower is better:}\quad \text{bar} = Q_{\text{sector}}\left(1-\tfrac{t}{100}\right) \qquad t = 50 + \text{shift}
```

Q is the sector's quantile, read from all companies in the sector that report the measure. Further rules:

- P/E and debt to equity use only positive values, and a company with a negative P/E (a loss) or negative equity never passes those rules, in either set.
- Energy, Materials and Real Estate have a P/E shift of minus 15, so t = 35 and the P/E bar is the sector's 65th percentile. Their earnings swing with the cycle, so a strict P/E bar would reject them at the wrong moments.
- The Piotroski bar is rounded up to a whole score.
- Market cap is a fixed floor of $300 million (US dollars), the same for every sector. It never moves with the data or the slider.
- A measure reported by fewer than 20 companies in a sector gets no bar for that sector.
- Hand overrides go in `TUNED_OVERRIDES` in marquee_screener.py and replace the computed value.

### 4.3 Which rules apply to which sector

Each sector is judged only on the measures that describe its kind of business. In the web screener, Rules by sector (top of the Criteria list) shows this grid with the live bar in every cell, and the reason in every not-used cell when you hover over it.

| Rule | Information Technology, Communication Services, Consumer Discretionary, Consumer Staples, Health Care, Energy, Materials, Industrials, Unclassified | Financials | Real Estate | Utilities |
| --- | --- | --- | --- | --- |
| P/E | Yes (Energy, Materials looser) | Yes | Yes, looser | Yes |
| Return on invested capital | Yes | No | No | No |
| Return on equity | No | Yes | Yes | Yes |
| Return on assets | No | Yes | No | No |
| Return on capital employed | No | No | No | No |
| Gross margin | No | No | No | No |
| Operating margin | Yes | No | Yes | Yes |
| Net margin | No | Yes | No | No |
| Free cash flow margin | Yes | No | No | No |
| Current ratio | No | No | No | No |
| Debt to equity | Yes | No | Yes | No |
| Interest cover | Yes | No | Yes | Yes |
| Altman Z-score | Yes | No | No | No |
| Piotroski F-score | Yes | Yes | Yes | Yes |
| Market cap at least $300m | Yes | Yes | Yes | Yes |

Why a rule is left out:

| Rule | Reason shown in the screener |
| --- | --- |
| Return on equity (ordinary companies) | ROIC is used instead; borrowing can inflate ROE |
| Return on assets | Used only for banks and insurers |
| Return on capital employed | ROIC is used instead; the two say the same thing |
| Return on invested capital (financials, property, utilities) | Not meaningful for lenders, landlords and regulated utilities |
| Gross margin, net margin (most sectors) | Operating and free cash flow margin say more |
| Operating margin (financials) | A bank has no operating margin in the usual sense |
| Free cash flow margin (financials, property, utilities) | Capital spending or lending makes it meaningless here |
| Current ratio | Interest cover and debt tell the story better |
| Debt to equity (financials, utilities) | Debt is the raw material of a bank, and normal for a utility |
| Interest cover (financials) | A bank's interest is its cost of goods |
| Altman Z-score (financials, property, utilities) | The model was not built for this kind of balance sheet |

### 4.4 The same-for-all set

| Rule | Bar |
| --- | --- |
| P/E | at most 20 |
| ROE, ROIC, ROCE, ROA | at least 10% each |
| Gross, operating, net and free cash flow margin | at least 10% each |
| Altman Z-score | at least 2.99 |
| Piotroski F-score | at least 7 |
| Current ratio | at least 1 |
| Debt to equity | at most 2 |
| Interest cover | at least 3 |
| Market cap | no floor |

### 4.5 One sector, several sectors, and the per-sector switch

- One sector ticked: the criteria list shows that sector's own bars as editable boxes. Change any value and the screen updates at once.
- Several sectors ticked, with Judge each company on its own sector's rules on (the default): every company is judged on its own sector's bars. The criteria list shows the range of bars across the ticked sectors and how many sectors use each rule.
- Several sectors ticked, with that switch off: one rule set (the All Sectors bars of the chosen threshold set) applies to every company, and its boxes become editable.

The All Sectors sheet of the workbook is the union of the 11 sector screens, plus Unclassified companies judged on the All Sectors rules.

### 4.6 Missing figures and the coverage guard

By default a blank figure is ignored, so a company is judged on what it reports. To stop a company passing on almost no data, it must report at least 60% of the rules applied to it (`MIN_DATA_COVERAGE = 0.6`). Tick Exclude companies with missing figures to reject any company with a blank in any applied rule.

## 5. Screener: criteria reference

The screener has 15 criteria in six groups. TradingView calculates every figure; the formulas below are the standard definitions, so treat them as what the number means rather than a recalculation.

| Group | Criterion | TradingView column | Passes when | Period | Formula, in words |
| --- | --- | --- | --- | --- | --- |
| Valuation | P/E, trailing | Price to earnings ratio | At most the bar, and above zero | Trailing 12 months | Share price ÷ earnings per share |
| Returns | Return on equity (ROE) | Return on equity %, Trailing 12 months | At least the bar | Trailing 12 months | Net income ÷ shareholders' equity |
| Returns | Return on assets (ROA) | Return on assets %, Trailing 12 months | At least the bar | Trailing 12 months | Net income ÷ total assets |
| Returns | Return on invested capital (ROIC) | Return on invested capital %, Trailing 12 months | At least the bar | Trailing 12 months | After-tax operating profit ÷ (debt + equity, less cash) |
| Returns | Return on capital employed (ROCE) | Return on capital employed %, Trailing 12 months | At least the bar | Trailing 12 months | Operating profit (EBIT) ÷ (total assets − current liabilities) |
| Margins | Gross margin | Gross margin %, Trailing 12 months | At least the bar | Trailing 12 months | Gross profit ÷ revenue |
| Margins | Operating margin | Operating margin %, Trailing 12 months | At least the bar | Trailing 12 months | Operating profit ÷ revenue |
| Margins | Net margin | Net margin %, Trailing 12 months | At least the bar | Trailing 12 months | Net income ÷ revenue |
| Margins | Free cash flow margin | Free cash flow margin %, Trailing 12 months | At least the bar | Trailing 12 months | (Operating cash flow − capital spending) ÷ revenue |
| Balance sheet | Current ratio | Current ratio, Quarterly | At least the bar | Latest quarter | Current assets ÷ current liabilities |
| Balance sheet | Debt to equity | Debt to equity ratio, Quarterly | At most the bar, and above zero | Latest quarter | Total debt ÷ shareholders' equity |
| Balance sheet | Interest cover | Interest coverage, Trailing 12 months | At least the bar | Trailing 12 months | Operating profit (EBIT) ÷ interest expense |
| Quality scores | Altman Z-score | Altman Z-score, Trailing 12 months | At least the bar | Trailing 12 months | Weighted mix of five balance sheet and profit ratios (below) |
| Quality scores | Piotroski F-score | Piotroski F-score, Trailing 12 months | At least the bar | Trailing 12 months | Count of nine tests passed, 0 to 9 (below) |
| Size | Market cap | Market capitalization | At least the bar | Current | Share price × shares in issue, in US dollars |

Percentages are stored as reported, so 12.5 means 12.5%. In the sidebar the market cap box is in billions of US dollars.

### 5.1 Altman Z-score

A bankruptcy-risk score. In the classic form for listed manufacturers:

```latex
Z = 1.2\,\frac{\text{working capital}}{\text{total assets}} + 1.4\,\frac{\text{retained earnings}}{\text{total assets}} + 3.3\,\frac{\text{EBIT}}{\text{total assets}} + 0.6\,\frac{\text{market value of equity}}{\text{total liabilities}} + 1.0\,\frac{\text{sales}}{\text{total assets}}
```

Below 1.8 is the distress zone, 1.8 to 3.0 the grey zone, above 3.0 the safe zone. The bands do not apply to banks, insurers or REITs, which is why those sectors do not use it.

### 5.2 Piotroski F-score

One point for each test passed, comparing the latest year with the year before:

1. Net income is positive.
2. Operating cash flow is positive.
3. Return on assets has risen.
4. Operating cash flow exceeds net income.
5. Long-term debt to assets has fallen.
6. The current ratio has risen.
7. No new shares were issued.
8. Gross margin has risen.
9. Asset turnover (revenue ÷ assets) has risen.

7 to 9 is strong, 4 to 6 middling, 0 to 3 weak.

## 6. Screener: strictness slider

The slider moves every bar by the same number of percentile points within its own sector, so each step admits or removes roughly the same share of companies on every rule, whatever its units.

### 6.1 Range and labels

The slider runs from 20 points stricter to 30 points looser, in steps of 5. With the sector-tuned set, the label above it reads Bars at the Nth percentile of each sector: 50th as set, down to the 20th at the loosest and up to the 70th at the strictest.

### 6.2 How a bar is moved

For every sector and measure, the build stores 21 points of the distribution: the 0th, 5th, 10th and so on up to the 100th percentile. To move a bar by L points:

1. Find the percentile p where the current bar sits in the sector, by straight-line interpolation between the stored points.
2. Move it: for a higher-is-better rule p' = p − L; for a lower-is-better rule (P/E, debt to equity) p' = p + L. L is positive when looser. p' is held between 1 and 99.
3. The new bar is the value at p', again by interpolation.

```latex
\text{higher is better: } p' = p - L \qquad \text{lower is better: } p' = p + L \qquad \text{new bar} = Q_{\text{sector}}(p')
```

- The Piotroski bar is rounded to the nearest whole score after moving.
- The $300 million market cap floor never moves.
- The same method works on the same-for-all set and on bars you type in: the bar is placed in the sector's distribution and moved from there. With one rule set for several sectors, the distribution of the whole universe is used.
- When a bar has moved, the criteria list shows the value actually applied, for example applied as 6.59%.

### 6.3 Worked examples, 21 October 2026 data

| Sector and rule | Stricter by 10 (60th) | As set (50th) | Looser by 10 (40th) | Looser by 20 (30th) |
| --- | --- | --- | --- | --- |
| Financials, return on equity at least | 12.61% | 10.77% | 8.72% | 6.59% |
| Information Technology, P/E at most | 21.0x | 27.3x | 36.4x | 47.2x |

For P/E, stricter means a lower cap, so the 60th percentile setting uses the sector's 40th percentile P/E, and looser by 20 uses its 70th.

### 6.4 What it does to the count

All sectors ticked, sector-tuned set, each company on its own sector's rules:

| Bars at | Companies passing |
| --- | --- |
| 70th percentile | 278 |
| 60th percentile | 792 |
| 50th percentile (as set) | 1,608 |
| 40th percentile | 3,159 |
| 30th percentile | 5,102 |
| 20th percentile | 7,891 |

## 7. Screener: momentum stages

Momentum is a second filter for timing: it sorts the companies that already pass the quality rules into five trend stages, so you can keep only those just starting to move up.

### 7.1 The data it uses

All from the TradingView export, as of its date: price, the 50 day and 200 day exponential moving averages (EMA), the 14 day RSI, and the 6 month and 1 year returns. The 1 month and 3 month returns, the 52 week high and relative volume are used when the export carries them.

### 7.2 The six conditions for an early stage

| Condition | Test | What it catches |
| --- | --- | --- |
| Trend turning up | Price above the 50 day EMA, and the 50 day EMA between 0.97 and 1.10 times the 200 day EMA | The 50 day has just crossed above the 200 day, or is about to |
| Not yet stretched | Price no more than 15% above the 50 day EMA | The move has started but has not run away |
| RSI between 50 and 70 | 14 day RSI from 50 to 70 inclusive | Strength without being overbought |
| Gaining pace | With a 3 month return: 3 month return above zero and above half the 6 month return. Without it: 6 month return above zero and above half the 1 year return | The recent part of the move is faster than the longer part |
| Room below its 52 week high | Price 5% to 25% below the 52 week high | Space to rise before the old high |
| Volume confirms | Relative volume at least 1.2 | More trading than usual behind the move |

How the conditions combine:

- Every ticked condition must hold (AND). Untick a condition to drop it.
- A condition whose column is not in the export is greyed out (needs data) and ignored for every company.
- If a company is missing the figure for a ticked condition, that condition fails for that company.
- If no condition is ticked, no company is Early.

### 7.3 The five stages

Each company gets exactly one stage, checked in this order; the first that fits wins.

1. Early, just turning up: every ticked condition in 7.2 holds.
2. Extended: price more than 25% above the 50 day EMA, or RSI above 75 with the price above the 50 day EMA.
3. Downtrend: price below the 50 day EMA and the 50 day EMA below the 200 day EMA.
4. Established uptrend: price above the 50 day EMA and the 50 day EMA more than 10% above the 200 day EMA.
5. No clear trend: anything else.

A company without a price or either EMA gets no stage and is hidden while the momentum filter is on. Keep Not yet stretched ticked: because Early is checked first, unticking it lets a stretched stock count as Early.

### 7.4 Using the filter

- The number beside each stage counts the companies that pass the quality rules and fall in that stage.
- Tick Filter by momentum stage, then the stages you want. Ticked stages combine with OR: Early and Established together show both.
- Ticking a stage switches the filter on by itself.

### 7.5 Worked example

Price 105, 50 day EMA 100, 200 day EMA 96, RSI 61, 6 month return +12%, 1 year return +14%, no 3 month column.

1. Trend turning up: price is above 100, and 100 ÷ 96 = 1.04 is between 0.97 and 1.10. Holds.
2. Not yet stretched: 105 is 5% above 100. Holds.
3. RSI 61 is between 50 and 70. Holds.
4. Gaining pace: 12 is above zero and above 14 ÷ 2 = 7. Holds.

All four available conditions hold, so the stage is Early.

### 7.6 Momentum score

The Momentum column shows the stage; sorting it uses a score from 0 to 100. Each condition that holds earns its weight, and the total is divided by the weights of the conditions that have data:

| Condition | Weight |
| --- | --- |
| Trend turning up | 25 |
| Gaining pace | 25 |
| RSI between 50 and 70 | 20 |
| Not yet stretched | 15 |
| Room below its 52 week high | 10 |
| Volume confirms | 5 |

```latex
\text{score} = 100 \times \frac{\sum \text{weights of conditions that hold}}{\sum \text{weights of conditions with data}}
```

## 8. Screener: results and exports

Everything updates the moment a tick, value or slider changes, and every export carries all companies that pass, not just the page on screen.

### 8.1 The summary card

- The count of companies that pass, with a one-line description of the rules in force (threshold set, how rules are applied, strictness, momentum).
- US listed, Outside the US, Superinvestor held, and Median market cap of the passing companies.
- A chip per sector with its passing count, shown when more than one sector is represented.

### 8.2 The results table

| Column | Shows |
| --- | --- |
| Company | Ticker (a link to the deep analysis page, sent with the ISIN and currency when known), company name, and a gold star when superinvestors hold it. Stays in place when scrolling sideways. |
| Sector, Country | GICS sector and country of registration |
| Market cap | In US dollars: $5.2tn, $668bn, $814m |
| P/E, Fwd P/E, Int. cover | Multiples, for example 27.4x |
| ROE, ROA, ROIC, Op. margin, Net margin, FCF margin | Percentages; negatives in red |
| Debt / equity, Current ratio, Altman Z | Two decimals |
| Piotroski | Whole score out of 9 |
| 1 year, 5 yr a year | Price return; 5 yr a year is the annualised 5 year return |
| Momentum | The stage from section 7; sorts by momentum score |
| Price | Last price in the listing currency |

Click a heading to sort; click again to reverse. Choose 25, 50, 100 or 250 rows per page. The 5 year a year figure is:

```latex
\text{5 yr a year} = \left(1 + \tfrac{\text{5 year return \%}}{100}\right)^{1/5} - 1
```

It is left blank when the 5 year loss is 100% or worse.

### 8.3 Excel workbook export

Three sheets, in navy and gold, ready to print landscape and one page wide:

- Summary: the settings used, then a table by sector of names, US, non-US, superinvestor held, median P/E, median ROE, median operating margin and total market cap, with a total row. Counts and market cap are live formulas on the Results sheet; medians are fixed at export.
- Results: every passing company, grouped headings, sorted by sector then market cap, with filters and frozen headings. Includes momentum stage and score.
- Thresholds: the exact bar applied for every rule in every rule set used, after strictness, with off where a rule is not used.

### 8.4 PowerPoint export

A 16:9 deck in navy and gold, content in tables:

1. Cover: how many companies passed, across which sectors, and the data date.
2. Summary: a table by sector (names, US, non-US, superinvestor held, median P/E, ROE, operating margin, market cap) and up to five of the most widely held by superinvestors.
3. How the names were selected: the settings and the bars for every rule set.
4. One slide per sector with its 12 largest companies by market cap: ticker, company, country, market cap, P/E, ROE, operating margin, net margin, momentum, Piotroski and superinvestor holding. With one sector ticked, up to three slides (36 names).
5. Notes on the data, ending with a plain statement that a screen is not a recommendation.

### 8.5 CSV export

One row per passing company: ticker, company, sector, country, market cap in US dollars, P/E, forward P/E, ROE, ROA, ROIC, operating, net and FCF margin, debt to equity, interest cover, current ratio, Altman Z, Piotroski, 1 year return, 5 year a year, momentum stage, momentum score, price, currency, superinvestor held and ISIN.

File names follow `MICS_screen_<sectors>_<data date>`, with the sector names when three or fewer are ticked.

### 8.6 Export libraries

The Excel and PowerPoint exports are built in the browser with ExcelJS and PptxGenJS, loaded from cdn.jsdelivr.net (cdnjs as a backup) on the first export. That first export needs internet access; if a firewall blocks the library, a message appears next to the buttons. CSV needs nothing.

### 8.7 Saved settings and the script

- The browser remembers your ticks, values, slider, momentum choices, sort and rows per page. Reset everything returns to the defaults.
- Copy rules for the script copies the current bars, after strictness, as a ready block: paste it into `TUNED_OVERRIDES` (sector-tuned) or `SECTOR_OVERRIDES` (same for all) in marquee_screener.py to make them permanent.
- Index analysis in the header opens the index page of the deep analysis app. Guide opens this guide.

## 9. Deep analysis: company page

The company page combines the nightly TradingView figures (so it agrees with the screener) with live Yahoo data for statements, price history, valuation, ownership and news, and ends in one verdict (section 10).

### 9.1 Opening a company

- From the screener: click a ticker. The link carries the ticker, the ISIN when known and the currency.
- From the app's front page: search by name, ticker or ISIN, or pick from four lists of 15: Most widely held (most superinvestors), Highest conviction (largest single superinvestor position), Cheapest that cleared (lowest P/E of at least 2 among screened names over $300m with positive ROE) and Highest quality (highest ROE among screened names over $300m with Piotroski 8 or 9).
- Finding the company in the feed: by its key (ISIN, or TV:TICKER:CCY), else by ticker; when a ticker exists on several exchanges, the listing in the expected currency wins.
- Finding it on Yahoo: with an ISIN, Yahoo's search resolves it and the listing in the expected currency is taken. Without one, Yahoo is searched by ticker, then by company name, again preferring the expected currency. A US ticker is used as is, with BRK.B written BRK-B.

### 9.2 Masthead and the business

Name, sector, industry, country, ISIN, last price and change, and the position within the 52 week range. The 52 week range and analyst targets are compared with Yahoo's own price, so both sides are in the same currency even when Yahoo resolved a different listing. The business card is Yahoo's company description.

### 9.3 Quality and forensic strip

| Score | Source | Bands | What it measures |
| --- | --- | --- | --- |
| Altman Z | TradingView | Distress below 1.8, grey zone 1.8 to 3, safe above 3. Not applicable to Financials, Real Estate and Utilities | Bankruptcy risk |
| Piotroski F | TradingView | Strong 7 to 9, middling 4 to 6, weak 0 to 3 | Nine tests of improving fundamentals |
| Beneish M | Calculated from Yahoo statements | Flagged above −1.78, clean below | Possible earnings manipulation |
| Distress probability (Ohlson O) | Calculated from Yahoo statements | High above 10%, watch above 5%, low otherwise | Chance of financial distress within two years |
| Accruals (Sloan) | Calculated from Yahoo statements | Profit ahead above 5%, cash ahead below −5%, neutral between | Share of profit not backed by cash |
| Cash conversion | Calculated from Yahoo statements | Backed at 1.1x or more, in line at 0.9x to 1.1x, thin below 0.9x. Not shown for Financials and Real Estate | Operating cash flow over net profit, averaged over profitable years |

```latex
M = -4.84 + 0.92\,\text{DSRI} + 0.528\,\text{GMI} + 0.404\,\text{AQI} + 0.892\,\text{SGI} + 0.115\,\text{DEPI} - 0.172\,\text{SGAI} + 4.679\,\text{TATA} - 0.327\,\text{LVGI}
```

The eight Beneish inputs compare this year with last year: receivables to sales (DSRI), gross margin (GMI), asset quality (AQI), sales growth (SGI), depreciation (DEPI), overheads to sales (SGAI), total accruals to assets (TATA) and leverage (LVGI).

```latex
O = -1.32 - 0.407\,\text{size} + 6.03\,\tfrac{TL}{TA} - 1.43\,\tfrac{WC}{TA} + 0.076\,\tfrac{CL}{CA} - 2.37\,\tfrac{NI}{TA} - 1.83\,\tfrac{FFO}{TL} + 0.285\,\text{loss2y} - 1.72\,\text{negEq} - 0.521\,\Delta NI \qquad P(\text{distress}) = \frac{1}{1+e^{-O}}
```

```latex
\text{Sloan accruals} = \frac{\text{net income} - \text{operating cash flow}}{\text{total assets}} \qquad \text{cash conversion} = \text{average of } \frac{\text{operating cash flow}}{\text{net income}} \text{ over years with a profit}
```

### 9.4 Price, drawdown and risk

The chart shows the adjusted daily close in navy, the 50 day average in gold (solid) and the 200 day average in blue-grey (dashed), with a legend above. The lower panel is the drawdown: how far the price sat below its highest level so far in the window. Hovering shows the date, close and both averages.

Range buttons (3M, 6M, 1Y, 2Y, 5Y, 10Y, Max) switch instantly; every range is drawn when the page opens. A range is offered when at least 60% of its days exist. Under each chart, measured on the full daily series of that window:

| Figure | Calculation |
| --- | --- |
| High and low | Highest and lowest close in the window |
| Change over the window | Last close ÷ first close − 1 |
| S&P 500, same window | The same, for the S&P 500 index (^GSPC) over the same dates |
| Ahead of or behind the S&P 500 | Stock change minus S&P 500 change, in percentage points |
| Worst fall from a high | Largest drop from a running high to a later low, with both dates, and when (or whether) the old high was regained |
| Volatility, annualised | Standard deviation of daily returns × √252 |
| Beta against the S&P 500 | Covariance of stock and index returns ÷ variance of index returns |
| Correlation with the index | Correlation of the same returns |

```latex
\beta = \frac{\operatorname{cov}(r_{\text{stock}},\, r_{\text{S\&P 500}})}{\operatorname{var}(r_{\text{S\&P 500}})}
```

Beta rules:

- Weekly returns (Friday to Friday) for windows of one year or longer, and for every window on a listing outside the US, because markets in different time zones close at different times. Daily returns otherwise.
- At least 12 weekly or 40 daily returns are needed; below 26 returns the text warns that it is a rough reading.
- With a correlation below 0.3 the text says the beta is weak.
- For a listing not in US dollars the text notes that currency moves are part of the figure.

Below the ranges: the deepest fall in the full record with its dates and whether it was regained, and the price against its 200 day average.

### 9.5 Against its sector

Bars compare six ratios with the sector median from sector_stats.json: P/E, return on equity, operating margin, net margin, interest cover, debt to equity. The gold marker is the median; a bar's length is the company's value divided by the median. A negative value gets no bar. A P/E on a loss is left out. For Financials, debt to equity and interest cover are left out.

The reading counts the measures on the favourable side of the median (higher for returns, margins and interest cover; lower for P/E and debt to equity) and names those that trail.

The rest of the ratios table places each TradingView ratio against its sector quartiles: best quartile, better than median, worse than median, worst quartile. The words always mean better or worse for the holder, so a low P/E counts as better.

### 9.6 What drives the return on equity (DuPont)

```latex
\text{ROE} = \underbrace{\frac{\text{net income}}{\text{revenue}}}_{\text{net margin}} \times \underbrace{\frac{\text{revenue}}{\text{assets}}}_{\text{asset turnover}} \times \underbrace{\frac{\text{assets}}{\text{equity}}}_{\text{leverage}}
```

Up to four annual periods from Yahoo, oldest first. The reading names the main driver of the latest year: leverage when assets are 3 or more times equity and the net margin is under 10%; net margin when it is 10% or more; asset turnover when it is 1 or more; otherwise no single driver. A loss year gets no decomposition.

### 9.7 Does the profit arrive as cash, and where does it go

- Each year: net income, operating cash flow, and their ratio (only for profitable years).
- Reading: backed by cash at an average of 1.1x or more, tracking at 0.9x to 1.1x, profit ahead of cash below 0.9x. It counts the years cash exceeded profit and says how many loss years were left out.
- Uses of cash: capital spending, buybacks, dividends, debt repaid and acquisitions, each summed over the years shown and as a share of total operating cash flow.

### 9.8 Growth the holder actually received

Compound annual growth, first year to last year shown:

```latex
\text{CAGR} = \left(\frac{\text{last}}{\text{first}}\right)^{1/\text{years}} - 1
```

Shown for revenue, net income and share count, and per share (revenue per share, earnings per share). The gap column is total growth minus per share growth: a positive gap is dilution from new shares, a negative gap is the lift from buybacks. Share count trend reads falling (buybacks) below −0.5% a year, rising (new shares) above +0.5%, otherwise broadly flat.

### 9.9 The rail cards

| Card | What it shows and how it reads |
| --- | --- |
| Valuation | Yahoo's EV to EBITDA, price to book, price to sales, EV to revenue and PEG. P/E against the sector median: materially cheaper below 0.7 times the median, cheaper below 0.95, in line to 1.15, dearer below 1.6, well above otherwise. EV to EBITDA: low below 8, moderate below 15, high otherwise. A loss-making P/E or negative EBITDA is called out; EV to EBITDA and price to book are not judged for Financials and Real Estate. |
| Returns | TradingView year to date, 6 month, 1 year and 5 year returns, and 5 years annualised. The last year is called well ahead of or behind the 5 year pace when the gap exceeds 10 points. |
| Trend | Price against the 50 and 200 day averages and the 14 day RSI (TradingView). RSI above 70 reads stretched, below 30 washed out. |
| The analyst view | Yahoo's consensus, mean, low and high target, and the change to the mean target from Yahoo's own price. Fewer than 4 analysts, or a high to low target spread above half the price, is called out. |
| Superinvestor ownership | How many Dataroma-tracked managers hold it, and whether it cleared the screen. |
| Positioning | Institutional ownership (above 80%: fresh demand must come from outside that base; 40% to 80%: normal; below 40%: lightly held) and short interest (high above 8% of float, low below 3%). |
| Largest holders, Recent coverage | Yahoo's top institutional holders and latest news. |

## 10. Deep analysis: verdict scorecard

The verdict turns every figure into a 0 to 100 score against the company's own sector, rolls them into five weighted pillars, and maps the total to Buy (70 and above), Hold (45 to 69) or Sell (below 45). It is a mechanical rating for internal research, not advice.

### 10.1 Scoring one figure against its sector

The score is the figure's approximate percentile in its sector, from the sector's lower quartile (p25), median (p50) and upper quartile (p75):

```latex
\text{score} = \begin{cases} 25 - 22\cdot\min\!\left(1, \frac{p_{25}-v}{p_{75}-p_{25}}\right) & v \le p_{25} \\ 25 + 25\cdot\frac{v-p_{25}}{p_{50}-p_{25}} & p_{25} < v \le p_{50} \\ 50 + 25\cdot\frac{v-p_{50}}{p_{75}-p_{50}} & p_{50} < v \le p_{75} \\ 75 + 22\cdot\min\!\left(1, \frac{v-p_{75}}{p_{75}-p_{25}}\right) & v > p_{75} \end{cases}
```

The score therefore runs from 3 to 97. For a lower-is-better measure (P/E, forward P/E, debt to equity) the score is 100 minus this. The words used: best quarter of its sector (75 and above), better than its sector median (50 to 75), worse than its sector median (25 to 50), worst quarter (below 25). For P/E the words are cheapest quarter, cheaper than median, dearer than median, dearest quarter.

### 10.2 The five pillars

| Pillar | Weight | Figures (scored against the sector unless a fixed score is given) |
| --- | --- | --- |
| Quality | 30% | Most sectors: ROIC, operating margin, free cash flow margin. Financials: ROE, ROA, net margin. Real Estate and Utilities: ROE, operating margin. Plus cash conversion, except Financials and Real Estate: average 1.2x or more scores 85, 1.0x to 1.2x 70, 0.8x to 1.0x 45, below 0.8x 20. |
| Financial strength | 20% | Debt to equity (not Financials or Utilities; negative equity scores 10), interest cover (not Financials), Altman Z (not Financials, Real Estate or Utilities: above 3 scores 85, 1.8 to 3 scores 50, below 1.8 scores 15), Piotroski (score × 100 ÷ 9). |
| Valuation | 25% | Trailing P/E (a loss scores 15). Forward P/E when positive, plus 5 points (up to 97) when it is below 0.9 times the trailing P/E, meaning earnings are expected to grow. |
| Momentum | 15% | The trend stage from the price history (10.4): Early 80, Established 70, Extended 50, No clear trend 45, Downtrend 20; plus 5 when the last 3 months beat half the 6 month return (not for Extended or Downtrend); plus half the 1 year lead over the S&P 500 in points, limited to ±15. Held between 3 and 97. |
| Signals | 10% | Held by superinvestors: 2 or more scores 70, one scores 62. With 4 or more analysts, change to the mean target: above 20% scores 80, above 5% 65, above −5% 50, otherwise 30. Short interest above 8% of float scores 25, below 3% scores 60. |

A pillar's score is the simple average of its figures. The total is the weighted average of the pillars that have data, with the weights rescaled to add to 100%:

```latex
\text{total} = \frac{\sum_{\text{pillars with data}} w_i \, s_i}{\sum_{\text{pillars with data}} w_i}
```

### 10.3 From score to opinion

1. Coverage: the share of expected figures (outside Signals) that were available. Below 50%, the opinion is Insufficient data.
2. Bands: total of 70 or more is Buy, 45 to 69 is Hold, below 45 is Sell.
3. Red flags cap a Buy at Hold; they never change a Hold or a Sell:
    1. Beneish M-score above −1.78.
    2. Altman Z below 1.8 (sectors where Altman applies).
    3. Average cash conversion below 0.7x over two or more years (not Financials or Real Estate).
    4. The latest annual year was a loss.

### 10.4 Momentum from the full price history

On the company page momentum uses the daily adjusted close from Yahoo (at least 220 days), with 50 and 200 day simple moving averages (SMA) and a 14 day RSI. Because the full history is known, the date of the last golden cross (50 day crossing above the 200 day) is exact. Checked in this order:

1. Early: price above the 50 day, the 50 day above the 200 day, the cross happened within the last 90 days, and the price no more than 15% above the 50 day.
2. Extended: price more than 25% above the 50 day, or RSI above 75 with the price above the 50 day.
3. Established: price above the 50 day and the 50 day above the 200 day.
4. Downtrend: price below the 50 day and the 50 day below the 200 day.
5. No clear trend: anything else.

The screener (section 7) uses TradingView's exponential averages and has no history, so the two can disagree for a stock close to a crossover.

### 10.5 How the verdict reads on the page

- The opinion and total, with the share of figures available.
- Any red flags, and whether they capped the opinion.
- One row per pillar: weight, score, the strongest figure, and the weakest figure when it scores below 50.
- Lifting the score: up to 3 figures scoring 65 or more. Holding it back: up to 3 scoring 35 or less.
- Check before acting: up to 3 checks drawn from the figures, for example a price still more than 40% below an old high that was never regained, heavy debt, high short interest, a price that trailed the S&P 500 by more than 25 points over a year, and always the latest quarterly results and guidance, which the page does not carry.
- The analyst's own note, when typed into the box under the page, appears at the top of the verdict.

## 11. Deep analysis: index page

The index page analyses a market rather than a company: its returns in local currency and to a US dollar investor, its falls, how closely it moves with the S&P 500, and what the fund that tracks it holds, costs and is priced at. It covers 70 indexes in 6 regions.

### 11.1 Opening an index

- Front page of the app: Analyse an index. Pick a region, then an index, then Analyse.
- The index home page (Index analysis button in the screener header) shows the same picker and a table of every index with its symbols and funds.
- Any other Yahoo symbol can be typed in: one starting with ^ is treated as an index, anything else as a fund.

### 11.2 The indexes covered

| Region | Indexes |
| --- | --- |
| Global (4) | MSCI ACWI, MSCI World, MSCI EAFE, MSCI Emerging Markets |
| United States (8) | S&P 500, Nasdaq 100, Nasdaq Composite, Dow Jones Industrial Average, Russell 1000, Russell 2000, Russell 3000, S&P MidCap 400 |
| Americas (8) | S&P Latin America 40, S&P/TSX Composite (Canada), S&P/BMV IPC (Mexico), Ibovespa (Brazil), S&P Merval (Argentina), MSCI Chile, MSCI All Colombia, MSCI All Peru |
| Europe (20) | STOXX Europe 600, EURO STOXX 50, FTSE 100, FTSE 250, DAX, CAC 40, AEX, IBEX 35, FTSE MIB, Swiss Market Index, OMX Stockholm 30, OMX Copenhagen 25, OMX Helsinki 25, MSCI Norway, BEL 20, ATX, MSCI Ireland, MSCI Poland, BIST 100 (Turkey), MSCI All Greece |
| Asia Pacific (24) | MSCI AC Asia ex Japan, Nikkei 225, TOPIX, CSI 300, SSE Composite, Shenzhen Component, MSCI China, Hang Seng, Hang Seng China Enterprises, TAIEX, KOSPI, KOSDAQ, Nifty 50, BSE Sensex, Nifty Bank, Jakarta Composite, FTSE Bursa Malaysia KLCI, Straits Times Index, SET Index, PSEi, MarketVector Vietnam Local, S&P/ASX 200, All Ordinaries, S&P/NZX 50 |
| Middle East and Africa (6) | Tadawul All Share, MSCI UAE, MSCI All Qatar, TA-35, MSCI South Africa, EGX 30 |

Each index has a Yahoo symbol for the index level and a tracking fund (ETF), preferring the US-listed fund, plus alternatives. For 20 indexes Yahoo has no reliable index series, so the page works from the fund alone: all MSCI indexes, S&P Latin America 40, S&P MidCap 400, TOPIX and MarketVector Vietnam. Three have no tracking fund: KOSDAQ, Nifty Bank and EGX 30. The list lives in indexes.py; a renamed symbol is a one-line fix there.

### 11.3 What the page shows

| Part | What it shows |
| --- | --- |
| Masthead | Index level (or the fund price when there is no index series), change year to date, distance below the 52 week high |
| What it covers | A one-line description of the index |
| Strip | One year return, five years a year, one year volatility, distance below the record high, five year beta against the S&P 500, fund cost a year |
| Level, drawdown and risk | The same chart and per-window figures as the company page (section 9.4), with beta and correlation against the S&P 500. For the S&P 500 itself, beta is 1 by definition. |
| Returns over time | 1, 3 and 6 months, year to date, 1 year, and 3, 5 and 10 years annualised, for the index, the fund and the S&P 500 side by side |
| What the fund holds | The ten largest holdings with their weights, the sector split, and the weight of the top ten |
| The fund that tracks it | Name, ticker, manager, cost a year, assets, distribution yield and alternative funds |
| Valuation | The fund's trailing P/E and price to book against the S&P 500's P/E (from SPY): a clear discount below 0.75 times, close to US equities up to 1.15 times, above otherwise |
| Trend | Against the 50 and 200 day averages and the 52 week high |

### 11.4 How the returns are measured

```latex
\text{return} = \frac{\text{last close}}{\text{close at the start}} - 1 \qquad \text{annualised (3 years and longer)} = (1+\text{return})^{12/\text{months}} - 1
```

- The start is the last close on or before the date that many months back; year to date starts from the last close of the previous year.
- The index is price only, in its own currency. The fund is adjusted for dividends, in its own currency. The S&P 500 is price only, in US dollars.
- When the fund's currency differs from the index's, the gap between the two columns is mostly the currency, plus dividends, less the fund's costs. When both are in the same currency, the fund normally runs ahead by about the dividend yield, less costs.

### 11.5 The reading

The Where the numbers point section lists up to four points per column, each triggered only by the data:

| Trigger | Column |
| --- | --- |
| Five years ahead of the S&P 500 / behind it (with a currency note for non-dollar indexes) | Supported / cuts against |
| Level above / below its 200 day average | Supported / cuts against |
| Five year correlation with the S&P 500 below 0.5 (diversifies) / above 0.85 (adds little) | Supported / cuts against |
| One year fund return in another currency at least 5 points above / below the index return | Supported / cuts against |
| Fund P/E below 0.75 times the S&P 500's (and why it is cheaper) / above 1.15 times | Supported and verify / cuts against |
| Top ten holdings above 45% of the fund | Cuts against, and verify the largest holdings |
| Fund cost above 0.5% a year | Cuts against |
| Still more than 15% below a high it never regained | Verify |
| Argentina, Turkey or Egypt: local returns inflated by inflation | Verify |
| Always: earnings revisions and the policy outlook | Verify |

## 12. Deep analysis: exports

Every company and index page has three buttons under it: Excel workbook, PowerPoint and Printable page (HTML). All three are built from the same figures as the page, in the same navy and gold design, with content in tables.

### 12.1 The opinion switch

On a company page, the box Include the Buy, Hold or Sell opinion in the Excel and PowerPoint files is off by default, because a rating sent to a client can count as an investment recommendation. With it off, the files carry the total score and the pillars but not the opinion. The opinion always shows on the page itself.

### 12.2 Company files

| Excel sheet | Contents |
| --- | --- |
| Overview | Snapshot, quality and forensic scores, verdict (score, opinion when switched on, pillar table, lifting, holding back, red flags, checks) |
| Against the sector | Every TradingView ratio against the sector's quartiles, with its position |
| Price and risk | One row per window: change, S&P 500, difference, worst fall with dates, recovery, volatility, beta, correlation, returns used. A native line chart of the stock against the S&P 500. |
| Price history | Weekly closes for 5 years, stock and S&P 500 rebased to 100 (feeds the chart) |
| Financials | DuPont, profit against cash, uses of cash, growth per share |
| Valuation and ownership | Yahoo multiples, ownership, short interest, analyst target |
| Notes | Sources, beta method, use |

PowerPoint, 9 slides: cover; at a glance (snapshot and scores); against the sector; price against the S&P 500 over 5 years (a native chart of month-end closes, with the window's figures beside it); return, drawdown and beta by window; how the return is earned and whether it arrives as cash; valuation, ownership and the analyst view; verdict (pillar table, lifting, holding back, checks); notes and sources.

### 12.3 Index files

- Excel: Overview (snapshot and the reading), Returns, Risk by window (with a native chart), Price history (weekly, rebased), Holdings (top holdings and sectors).
- PowerPoint, 7 slides: cover; at a glance with the reading; returns over time; the index against the S&P 500 (month-end, rebased); return, drawdown and beta by window; what the fund holds; notes and sources.

### 12.4 Printable page and file names

The printable page is one self-contained HTML file. Open it and press Ctrl+P to save a PDF; untick Headers and footers in the print dialog so the browser does not print a file path down the edges.

File names: `MICS_<ticker>_analysis_<date>` for a company and `MICS_<index name>_index_<date>` for an index.

## 13. Settings reference

Every setting below is a line near the top of its file; change it, commit, and the next run uses it.

### 13.1 Screener build (marquee_screener.py, Stock-Screener repo)

| Setting | Default | What it controls |
| --- | --- | --- |
| `SCREEN_UNIVERSE` | `'all'` | `'all'` screens every company; `'marquee'` screens only superinvestor holdings |
| `SCRAPE_DATAROMA` | `'auto'` | `'never'`, `'auto'` (scrape unless today's file exists) or `'always'` |
| `DATAROMA_MAX_AGE_DAYS` | `0` | Age in days before `'auto'` scrapes again; 0 means every new day |
| `DATAROMA_PAGES` | `15` | Most Dataroma pages read |
| `KEEP_ONLY_LATEST_DATAROMA` | `True` | Delete every Dataroma file except the newest |
| `THRESHOLD_MODE` | `'tuned'` | `'tuned'` (sector-tuned, section 4.2) or `'base'` (same for all) |
| `TUNED_PERCENTILE` | `50` | Percentile each tuned bar is set at |
| `MCAP_FLOOR` | `300e6` | Market cap floor in US dollars |
| `SECTOR_RULES` | See section 4.3 | Which rules each sector uses |
| `RULE_TARGET_SHIFT` | P/E −15 for Energy, Materials, Real Estate | Per-rule percentile shifts |
| `TUNED_OVERRIDES` | Empty | Hand-set bars that replace computed ones |
| `SECTOR_OVERRIDES`, `DEFAULT_FILTERS` | See section 4.4 | The same-for-all set |
| `ALL_SECTORS_USES_SECTOR_RULES` | `True` | The All Sectors sheet judges each company on its own sector's rules |
| `BLANK_FAILS` | `False` | `True` rejects any company with a blank screened figure |
| `MIN_DATA_COVERAGE` | `0.6` | Share of applied rules a company must report |
| `SECTOR_MAP`, `PULL_REITS_OUT_OF_FINANCE` | Section 3 | Sector mapping and the REIT move |
| `DEEP_APP_URL` | The Streamlit app address | Where ticker links in the screener point |
| `DATE_IN_FILENAME` | `True` | Date in output file names; the workflow switches it off with `--stable-names` |
| `GITHUB_UPLOAD` | `False` | Upload through the GitHub API; not needed, the workflow commits |

The workflow runs: `python marquee_screener.py --data-dir data --out-dir output --publish-dir . --scrape auto --stable-names --no-upload`. Other options: `--no-html` (skip the web page) and `--web` (split the page into index.html and data.json).

### 13.2 Deep analysis app (mics-deep-analysis repo)

| File and setting | Default | What it controls |
| --- | --- | --- |
| `scoring.py`: `WEIGHTS` | Quality 30, Financial strength 20, Valuation 25, Momentum 15, Signals 10 | Pillar weights |
| `scoring.py`: `BUY_AT`, `SELL_BELOW` | 70, 45 | Opinion bands |
| `risk.py`: `BENCH_SYMBOL` | `^GSPC` | Benchmark for beta and relative return |
| `risk.py`: `SPANS` | 3M 63, 6M 126, 1Y 252, 2Y 504, 5Y 1,260, 10Y 2,520 trading days, Max | Chart windows |
| `indexes.py`: `INDEXES` | 70 entries | The index list, symbols and funds |
| `app.py`: `FEED_BASE` | https://wealthmics.github.io/Stock-Screener | Where stocks.json and sector_stats.json are read |

How fresh the app's data is:

| Data | Kept for |
| --- | --- |
| Company feed (stocks.json) | 15 minutes |
| Sector statistics | 1 hour |
| Company data from Yahoo | 15 minutes |
| S&P 500 and index price histories | 1 hour in the app, 12 hours on disk |
| Fund profiles | 6 hours in the app, 12 hours on disk |

## 14. Known limits

The platform narrows the field for research; these are the places where its numbers need care.

| Area | Limit | What to do |
| --- | --- | --- |
| Screener counts | The web page stores figures to 4 decimals, so a few companies sitting exactly on a bar can differ from the workbook (1,608 against 1,611 on 21 October 2026) | Treat the workbook as the record |
| Momentum in the screener | A one-day snapshot: it cannot know when the 50 day crossed the 200 day, and it uses TradingView's exponential averages while the company page uses simple ones | Confirm the stage on the company page |
| Momentum signals | Room below the 52 week high and volume stay off until those columns are added to the TradingView export; gaining pace uses 6 months against 1 year until the 3 month return is added | Add 1 month, 3 month, 52 week high and relative volume to the export |
| Thresholds and weights | Momentum tests, verdict weights and opinion bands are standard rules of thumb, not backtested | Read the score as a ranking aid |
| Mixed dates | Ratios and screener returns are from the TradingView export date; prices, statements and charts on the deep analysis pages are live from Yahoo | Check the dates shown on each card |
| Statements | Yahoo returns at most four years, and restatements overwrite earlier years | Fine for analysis, not for backtesting |
| Superinvestor data | Scraped daily, but managers file holdings quarterly, up to 45 days after the quarter ends | Read ownership as of the last filing |
| Beta for non-US listings | Measured in local currency against the S&P 500 in US dollars, so currency moves are part of it | Use the correlation alongside it |
| Index symbols | Yahoo occasionally renames symbols; the page says which one failed | Correct it in indexes.py |
| Sector medians | Computed over every company in the sector, micro caps included | A median company is not a quality company; the bar is a floor |
| The verdict | Reads only the figures on the page; no news, guidance or management view | Use Check before acting, and add your own note |
| Exports in the screener | Need internet once to load the Excel and PowerPoint libraries | Allow cdn.jsdelivr.net, or use CSV |

## 15. Glossary

| Term | Plain meaning |
| --- | --- |
| ADR | A US-listed certificate representing shares of a foreign company |
| Annualised | A multi-year return expressed as the steady yearly rate that would give the same result |
| Beta | How much a stock has moved for each 1% move in the S&P 500; 1.2 means about 1.2% |
| CAGR | Compound annual growth rate: the steady yearly growth from a first value to a last value |
| Correlation | How closely two series move together, from −1 (opposite) through 0 (unrelated) to 1 (in step) |
| Coverage guard | The rule that a company must report at least 60% of the applied rules to pass |
| Dataroma | A website tracking the 13F filings of well-known value investors |
| Drawdown | How far a price sits below its highest level so far |
| DuPont | Splitting return on equity into net margin, asset turnover and leverage |
| EMA, SMA | Exponential and simple moving averages. An EMA gives more weight to recent days |
| EV to EBITDA | Enterprise value (market cap plus debt, less cash) divided by operating profit before depreciation |
| GICS | Global Industry Classification Standard: the 11 sectors used here |
| Golden cross | The 50 day average crossing above the 200 day average |
| ISIN | A 12-character code that identifies one security worldwide |
| Marquee or superinvestor holding | A company held by at least one manager tracked on Dataroma |
| Median | The middle value: half the companies are above it, half below |
| Percentile | The share of companies below a value; the 60th percentile beats 60% of them |
| Quartile | The 25th (lower), 50th (median) and 75th (upper) percentiles |
| Relative volume | Today's trading volume divided by its recent average |
| RSI | Relative strength index, 0 to 100; above 70 is often called overbought, below 30 oversold |
| Short interest | Shares sold short, as a share of the shares available to trade |
| Trailing twelve months (TTM) | The latest four quarters added together |
| Volatility | How widely returns swing, as an annualised standard deviation |
