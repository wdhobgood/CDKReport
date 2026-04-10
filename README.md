# FTZ Duty Calculator

VBA-based Excel tool for processing Ship-PC customs entry files. Calculates import duties using CBP's ES003 entry summary methodology — aggregating transactions into ES003 lines, rounding entered values to whole dollars, and applying per-ordinal duty rounding.

## CBP ES003-Aligned Calculation (v2026-04-10)

This tool now mirrors how CBP/ACE actually calculates duty on entry summaries:

1. **Aggregate** transactions into ES003 lines by grouping key: `MFN HTS + Manufacturer ID + Country of Origin + Chapter 99 codes`
2. **Entered Value** = `Round(SUM(qty × sum_of_all_sequence_values), 0)` — whole dollars
3. **Per-ordinal duty** = `WorksheetFunction.Round(EnteredValue × rate, 2)` — arithmetic rounding to cents
4. **Line total** = MFN Duty + S301 Duty + S338 Duty + S122 Duty
5. **Entry total** = SUM of all line totals

### Key Differences from Previous Version

| Aspect | Previous (per-transaction) | Current (CBP-aligned) |
|--------|---------------------------|----------------------|
| Calculation level | Per Ship-PC row | Per ES003 aggregated line |
| Entered value | MFN sequence value only | SUM of ALL sequence values |
| Value rounding | None | Rounded to whole dollars |
| Duty rounding | None (Excel default) | `WorksheetFunction.Round(, 2)` per ordinal |
| HTS lookups | `IsHTSMatch()` linear scan | `Scripting.Dictionary` O(1) |
| Cell I/O | Per-cell (~4.7M reads) | Bulk array (1 read + 1 write) |
| Fee distribution | Even per row | Proportional by entered value |

## Ship-PC Format Specification

### Input File Column Structure

| Column | Field | Notes |
|--------|-------|-------|
| 2 | Receipt Date | Used for IEEPA matrix lookup |
| 3 | Entry Number | Primary identifier |
| 6 | NGC# | Internal reference |
| 7 | Material Number | Part identifier |
| 8 | PO Number | Purchase order |
| 11 | Quantity | Item quantity |
| 12 | Manufacturer ID | Used for ES003 line grouping |
| 13 | Country of Origin | ISO 2-letter code |
| **14-16** | **Seq 1: HTS/Value/Rate** | First sequence |
| **17-19** | **Seq 2: HTS/Value/Rate** | Second sequence |
| **20-22** | **Seq 3: HTS/Value/Rate** | Third sequence |
| **23-25** | **Seq 4: HTS/Value/Rate** | Fourth sequence |
| 26 | Status | D = Deleted (filtered out) |

### Processing Algorithm (CBP ES003 Method)

**Phase 1 — Scan & Aggregate** (all input rows read into memory array):
```
For each input row:
    1. Sum ALL sequence values: sumAllSeqValues = val1 + val2 + val3 + val4
    2. Identify MFN (first non-Chapter-99 HTS code) for rate
    3. Categorize Chapter 99 codes (S301/S338/S122) via Dictionary lookup
    4. Build grouping key: MFN_HTS | ManufacturerID | COO | sorted_ch99_codes
    5. Aggregate: group.TotalEnteredValue += qty × sumAllSeqValues
```

**Phase 2 — Calculate & Output** (one output row per aggregated line):
```
For each ES003 line:
    enteredValue = WorksheetFunction.Round(TotalEnteredValue, 0)
    mfnDuty  = WorksheetFunction.Round(enteredValue × mfnRate, 2)
    s301Duty = WorksheetFunction.Round(enteredValue × s301Rate, 2)
    s338Duty = WorksheetFunction.Round(enteredValue × s338Rate, 2)
    s122Duty = WorksheetFunction.Round(enteredValue × s122Rate, 2)
    totalDuty = mfnDuty + s301Duty + s338Duty + s122Duty
```

### Why SUM of ALL Sequence Values?

When a Chapter 99 tariff is present, the Ship-PC may report the same unit value on BOTH the Chapter 99 and MFN sequences. CBP's entered value includes both. Using only the MFN value underestimates by ~50% on affected lines.

**Verified against ES003 for entry CDK-6000743-9:**
- 319 of 331 ES003 lines matched on grouping key
- 256 lines: penny-perfect duty match
- Overall: 87.6% of ES003 total duty (up from 64.3% with old per-transaction method)
- Remaining gap is due to customs additions (freight, assists, packing) not reflected in Ship-PC

## Configuration

### Settings Sheet

| Cell/Column | Purpose |
|-------------|---------|
| Column A | S301 HTS codes |
| Column B | S338 HTS codes |
| Column C | S122 HTS codes |
| E2 | Total Cotton Fee |
| E3 | Total MPF |
| H2 | Total Duty Upper Bound |
| H3 | Total Duty Lower Bound |
| I2 | Duty Matrix file path |
| N2 | Unit Value Upper Bound |
| N3 | Unit Value Lower Bound |

### Duty Matrix Format
- Column A: HTS Code
- Column B: ISO Country Code
- Columns C+: Date columns (header = date, cells = rates as decimals)

## Output Structure

### Details Sheet Columns (Aggregated ES003 Lines)
```
A=Entry Number      I=MFN Duty
B=MFN HTS Code      J=S301 Duty
C=Manufacturer ID   K=S338 Duty
D=Country of Origin  L=S122 Duty
E=Ch99 Codes        M=99 Value
F=Total Qty         O=Cotton Fee
G=Txn Count         P=MPF
H=Entered Value     Q=Total Duty
                    R=Total Fees
                    S=Total
```

### Summary Sheet Cells
```
B2  = Entry Number
B5  = Total Duties Paid
B6  = Total Duties + Fees
B9  = MFN Duties
B10 = S301 Duties
B11 = S338 Duties
B12 = S122 Duties
B15 = 99 Value Total
```

## Core VBA Functions

**MasterProcessAllFiles()** — Main entry, batch processing
**ProcessSingleFile()** — Per-file orchestration
**ProcessDutyCalculations()** — Two-phase ES003-aligned engine (aggregate → calculate)
**BuildGroupKey()** — Constructs ES003 line grouping key
**SortCh99Codes()** — Sorts Chapter 99 codes for consistent grouping
**LoadHTSDictionary()** — Pre-loads Settings HTS lists into Dictionary for O(1) lookup
**ValidateIEEPADutyRate()** — Date-based matrix validation
**DistributeFees()** — Proportional fee distribution by entered value
**CalculateSummaryTotals()** — Aggregates for Summary sheet

## Performance

Optimized for large Ship-PC files (189K+ rows):
- Bulk array I/O: one `Range.Value` read/write instead of millions of cell accesses
- `Scripting.Dictionary` for HTS category lookups (O(1) vs O(n) per lookup)
- `Application.ScreenUpdating = False` and `xlCalculationManual` during processing
- Batched log writes
- Expected: seconds instead of minutes for 189K rows

## Validation Features

**IEEPA Duty Rate**: Validates Chapter 99 rates against date-based matrix (most recent date ≤ receipt date)
**Rate Consistency**: Warns if transactions in the same ES003 group have different MFN rates
**Unit Value Bounds**: Flags values outside N2/N3 thresholds
**Total Duty Bounds**: Alerts when total duty outside H2/H3 range
**99 Value Detection**: Flags non-zero Chapter 99 values

## Installation

1. Import `Module1.bas` into VBA (Alt+F11)
2. Create UserForm `ProgressForm` with controls: `lblStatus`, `lblPhase`, `lblEntry`, `frameProgress`, `ProgressBar1`
3. Configure Settings sheet (columns A/B/C, cells E2/E3/H2/H3/I2/N2/N3)
4. Create folders: `Input/`, `Input/Archive/`, `Output/`
5. Run `MasterProcessAllFiles` (Alt+F8)

## Common Issues

**"Method or data member not found"** — ProgressForm missing controls or ProgressBar1 not inside frameProgress
**"SharePoint/OneDrive web"** — Copy project to local drive
**Calculations = $0** — Check Settings columns A/B/C have correct HTS codes
**Duty Matrix not importing** — Verify path in Settings I2, file exists, data starts at A1
**Duty doesn't match ES003 exactly** — Ship-PC may not include customs additions (freight, assists, packing). The calculation method matches CBP; differences are input data limitations.

## Repository

https://github.com/wdhobgood/CDKReport

MIT License — See LICENSE file
