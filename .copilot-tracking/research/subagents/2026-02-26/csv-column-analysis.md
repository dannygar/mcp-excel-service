# CSV Column Analysis: ORB TRANSACTION TRACKER

**Source**: `docs/samples/2026 GY Capital Group LLC Trade Tracker.csv`
**Date**: 2026-02-26
**Status**: Complete

---

## 1. Row layout

| Row(s) | Content |
|--------|---------|
| 1–3 | Empty |
| 4 | Title: "Daily Trade Log" (cell A4) |
| 5–11 | Empty (row 6 has a stray backtick in column N) |
| 12 | Summary section header: Metrics / Values / Strategy / Profit / Profit per contract / Metrics / Values |
| 13–16 | Summary metrics data (YTD Count, YTD Occurence, Win Rate, P/L) and strategy breakdown + roll accounting |
| 17 | Estimated EOM Target (value is `#REF!`) and Roll 3 |
| 18–20 | Roll 4–6 rows (all `$0.00`) |
| 21 | VIX sub-table headers: "Column1", "Column2", "Column3" |
| 22 | VIX Previous Close: `$17.93` |
| 23 | VIX Current Price: `$18.66` |
| 24 | Empty |
| 25 | Section label: "ORB TRANSACTION TRACKER" (cell A25) |
| 26 | Empty |
| **27** | **Column headers for the trade table** |
| **28–56** | **Trade data rows (29 rows)** |
| 57+ | Empty trailing rows |

---

## 2. Summary/metrics section (rows 12–20)

### Left block (columns A–C)

| Row | Cell A | Cell C | Data type |
|-----|--------|--------|-----------|
| 12 | Metrics | Values | Headers |
| 13 | YTD Count | 28 | Integer |
| 14 | YTD Occurence | 29 | Integer |
| 15 | Win Rate | 97% | Percentage |
| 16 | P/L | $8,088.54 | Currency |
| 17 | Estimated EOM Target: | #REF! | Formula (broken ref) |

### Strategy block (columns E–G)

| Row | Cell E (Strategy) | Cell F (Profit) | Cell G (Profit per contract) |
|-----|-------------------|-----------------|------------------------------|
| 12 | Strategy | Profit | Profit per contract |
| 13 | IC | $952.43 | $38 |
| 14 | VPCS | $3,497.06 | $11 |
| 15 | VCCS | $3,639.05 | $14 |
| 16 | VIX | $0.00 | $0 |

### Right block (columns H–I)

| Row | Cell H | Cell I |
|-----|--------|--------|
| 12 | Metrics | Values |
| 13 | Profit (without fees): | $10,350.00 |
| 14 | Fees: | -$2,261.46 |
| 15 | Roll 1: | -$3,653.84 |
| 16 | Roll 2: | $511.06 |
| 17 | Roll 3: | $0.00 |
| 18 | Roll 4: | $0.00 |
| 19 | Roll 5: | $0.00 |
| 20 | Roll 6: | $0.00 |

### VIX reference (rows 21–23)

| Row | Cell A | Cell B | Cell C |
|-----|--------|--------|--------|
| 21 | Column1 | Column2 | Column3 |
| 22 | CBOE MKT VOLATILITY IDX | Previous Close | $17.93 |
| 23 | CBOE MKT VOLATILITY IDX | Current Price | $18.66 |

---

## 3. Complete column mapping (row 27 headers, 29 columns A–AC)

| Col | Header | Data type | Format | Sample values | Notes |
|-----|--------|-----------|--------|---------------|-------|
| A | Date of Purchase | Date | M/D/YYYY | `2/2/2026`, `2/13/2026` | Trade open date |
| B | Open Day | Text | 3-letter weekday | `MON`, `TUE`, `WED`, `THU`, `FRI` | Derived from Date of Purchase |
| C | Time IN | Time | H:MM AM/PM | `10:09 AM`, `2:02 PM` | Entry time |
| D | Close Day | Date | M/D/YYYY | `2/2/2026`, `2/5/2026`, `2/20/2026` | May differ from Open date for rolled trades |
| E | Time OUT | Time | H:MM AM/PM | `4:00 PM`, `10:24 AM`, `9:33 AM` | Exit time |
| F | Trade Time | Duration | HH:MM:SS | `00:05:51`, `04:02:20`, `07:04:05` | Computed from Time IN/OUT (spans days) |
| G | Strategy | Text | Enum | `VCCS`, `VPCS`, `IC` | Three distinct values |
| H | Credit | Currency | $X.XX | `$0.20`, `$0.30`, `$0.45` | Per-contract credit received |
| I | Debit | Currency | $X.XX | `$1.95` (only 1 occurrence) | Per-contract debit paid; empty for most trades |
| J | Contracts | Integer | Plain number | `25`, `5`, `2` | Number of contracts |
| K | Net Profit | Currency | $X,XXX.XX | `$500.00`, `$750.00`, `-$4,125.00`, `$1,125.00` | = (Credit - Debit) × Contracts × 100 |
| L | Open Fees | Currency | $XX.XX | `$86.26`, `$17.26`, `$122.56`, `$172.57` | Fees at open |
| M | Close Fees | Currency | $XX.XX | `$36.24` (only 1 occurrence) | Usually empty; filled when debit close occurs |
| N | Profit | Currency | $XXX.XX | `$413.74`, `-$4,283.76`, `$952.43` | = Net Profit - Open Fees - Close Fees |
| O | ATR | Number | X.XX (decimal) | `3.88`, `6.38`, `2.90` | Average True Range at time of trade |
| P | Short Calls | Currency | $X,XXX | `$7,015`, `$6,960`, `$6,975` | Strike price (no decimals); empty for VPCS |
| Q | Short Puts | Currency | $X,XXX | `$6,890`, `$6,860`, `$6,855` | Strike price (no decimals); empty for VCCS |
| R | Spread Width | Currency | $XXX | `$125`, `$155`, `$180`, `$120` | Difference between strikes; empty when only one side |
| S | Width | Currency | $XX | `$15`, `$20` | Option spread width in points |
| T | C Delta at 10am | Decimal | 0.XX | `0.04`, `0.03`, `0.02` | Call delta at 10:00 AM; empty for VPCS |
| U | P Delta at 10am | Decimal | 0.XX | `0.06`, `0.04`, `0.05` | Put delta at 10:00 AM; empty for VCCS |
| V | C Delta at 11:30 am | Decimal | 0.XX | `0.03`, `0.02`, `0.01` | Call delta at 11:30 AM; empty for VPCS |
| W | P Delta at 11:30am | Decimal | 0.XX | `0.03`, `0.05`, `0.55` | Put delta at 11:30 AM; empty for VCCS |
| X | C Delta at 1:30pm | Decimal | 0.XX | `0.03`, `0.02`, `0.01` | Call delta at 1:30 PM |
| Y | P Delta at 1:30pm | Decimal | 0.XX | `0.02`, `0.28`, `0.22` | Put delta at 1:30 PM |
| Z | C Delta at 3:00pm | Decimal | 0.XX | `0.02`, `0.01`, `0.00` | Call delta at 3:00 PM |
| AA | P Delta at 3:00pm | Decimal | 0.XX | `0.01`, `0.02`, `0.00` | Put delta at 3:00 PM |
| AB | Rolls | Integer | Plain number | `1`, `2` | Roll count; empty if no rolls |
| AC | Notes | Text | Free-form | `rolled out 6860->6835 (15)`, `rolled out 6845->6845 (20)` | Roll details; empty for most trades |

---

## 4. Data row count and range

- **Header row**: Row 27
- **First data row**: Row 28 (`2/2/2026, MON, 10:09 AM, ...`)
- **Last data row**: Row 56 (`2/26/2026, THU, 10:11 AM, ...`)
- **Total trade rows**: 29
- **Date range**: 2/2/2026 through 2/26/2026

---

## 5. Columns filled at open vs. close

### Filled at trade open time (by `logTrades`)

| Column | Header | Reasoning |
|--------|--------|-----------|
| A | Date of Purchase | Known at open |
| B | Open Day | Derived from date at open |
| C | Time IN | Known at open |
| G | Strategy | Selected at open |
| H | Credit | Known at open |
| J | Contracts | Known at open |
| L | Open Fees | Incurred at open |
| O | ATR | Market data at open |
| P | Short Calls | Strike selected at open (VCCS, IC only) |
| Q | Short Puts | Strike selected at open (VPCS, IC only) |
| R | Spread Width | Computed from strikes at open |
| S | Width | Selected at open |

### Filled at trade close time (by `closeTrade`)

| Column | Header | Reasoning |
|--------|--------|-----------|
| D | Close Day | Known at close |
| E | Time OUT | Known at close |
| F | Trade Time | Computed from open/close times |
| I | Debit | Only present if bought back (loss/roll) |
| K | Net Profit | Computed at close |
| M | Close Fees | Incurred at close (usually empty = expired worthless) |
| N | Profit | Computed at close |

### Filled progressively during trade lifetime (by `updateTradeWithDelta`)

| Column | Header | Time checkpoint |
|--------|--------|-----------------|
| T | C Delta at 10am | 10:00 AM |
| U | P Delta at 10am | 10:00 AM |
| V | C Delta at 11:30 am | 11:30 AM |
| W | P Delta at 11:30am | 11:30 AM |
| X | C Delta at 1:30pm | 1:30 PM |
| Y | P Delta at 1:30pm | 1:30 PM |
| Z | C Delta at 3:00pm | 3:00 PM |
| AA | P Delta at 3:00pm | 3:00 PM |

### Filled on roll events

| Column | Header | Reasoning |
|--------|--------|-----------|
| AB | Rolls | Incremented on roll |
| AC | Notes | Roll details appended |

---

## 6. Currency and number formatting patterns

| Pattern | Columns | Examples |
|---------|---------|----------|
| `$X.XX` (per-contract price) | H (Credit), I (Debit) | `$0.20`, `$1.95` |
| `$X,XXX.XX` (total dollar amount) | K (Net Profit), N (Profit) | `$500.00`, `-$4,125.00`, `$1,125.00` |
| `$XX.XX` (fee amount) | L (Open Fees), M (Close Fees) | `$86.26`, `$36.24`, `$172.57` |
| `$X,XXX` (strike price, no decimals) | P (Short Calls), Q (Short Puts) | `$7,015`, `$6,890` |
| `$XXX` (spread, no decimals) | R (Spread Width) | `$125`, `$180` |
| `$XX` (width, no decimals) | S (Width) | `$15`, `$20` |
| `X.XX` (plain decimal) | O (ATR) | `3.88`, `6.38` |
| `0.XX` (decimal 0–1) | T–AA (Deltas) | `0.04`, `0.55`, `0.00` |
| Integer | J (Contracts), AB (Rolls) | `25`, `5`, `2`, `1` |
| Negative currency: `-$X,XXX.XX` | K, N | `-$4,125.00`, `-$4,283.76` |

---

## 7. Strategy value inventory

| Strategy | Full name (inferred) | Count | Trades |
|----------|---------------------|-------|--------|
| **VCCS** | Vertical Credit Call Spread | 11 | Rows 28, 31, 35, 37, 39, 42, 44, 49, 51, 54, 56 |
| **VPCS** | Vertical Credit Put Spread | 17 | Rows 29, 30, 32, 33, 34, 36, 38, 40, 41, 43, 45, 46, 47, 48, 50, 52, 55 |
| **IC** | Iron Condor | 1 | Row 54 |
| **VIX** | VIX-related (no trades yet) | 0 | (referenced in summary only) |

### Strategy behavior by column:

- **VCCS**: Short Calls filled, Short Puts empty, Spread Width empty (only 1 side), Call deltas filled, Put deltas empty
- **VPCS**: Short Puts filled, Short Calls empty, Spread Width empty (only 1 side), Put deltas filled, Call deltas empty
- **IC**: Both Short Calls AND Short Puts filled, Spread Width filled, both Call and Put deltas filled

---

## 8. Key observations and patterns

1. **Most trades expire worthless**: Close Fees (M) is empty for 28 of 29 rows, meaning trades expired at `4:00 PM` with no buy-back. Only the losing trade on 2/5/2026 has `Close Fees: $36.24` AND `Debit: $1.95`.

2. **Rolled trades span multiple days**: Trades with `Rolls >= 1` have `Close Day` different from `Date of Purchase`. Close Day can be several trading days later (e.g., 2/13 → 2/20).

3. **Standard contract size is 25**: Most trades use 25 contracts. Reduced sizes (5, 2) appear after losses or during cautious periods.

4. **Width values**: Almost all trades use `$15` width. After rolls, some shift to `$20` width.

5. **Trade Time calculation**: Durations exceeding market hours (e.g., `04:02:20`, `07:04:05`, `00:20:17`, `00:23:09`) represent multi-day held positions, suggesting the computation spans calendar time, not just trading hours.

6. **IC row (2/25/2026)**: The single Iron Condor fills both P and Q columns, has the highest Open Fees ($172.57 ≈ 2× normal), and contains deltas for both call and put sides at all time checkpoints.

---

## 9. Clarifying questions / recommended next research

1. **Computed columns**: Which columns are Excel formulas vs. manually entered? Likely candidates for formulas: F (Trade Time), K (Net Profit), N (Profit), and possibly B (Open Day). The CSV cannot distinguish.
2. **Exact Excel cell addresses**: The CSV maps to Excel starting at row 1 column A. Confirm whether the actual workbook has the same starting position or if there are offsets.
3. **`#REF!` in row 17**: The "Estimated EOM Target" formula is broken. This probably references cells not present in the CSV export.
4. **VIX strategy**: The summary references VIX with `$0.00` profit. No VIX trades exist in the data; this may be a placeholder for future trades.
5. **logTrades vs. updateRange**: Clarify whether `logTrades` writes all open-time columns in one call, or if some (like ATR, deltas) require separate `updateRange` calls.
6. **Read the server.py implementation**: Cross-reference these column positions with the actual MCP tool implementations to validate column mapping matches code expectations.
