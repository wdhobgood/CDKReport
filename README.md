# FTZ Duty Calculator

VBA-based Excel tool for processing Ship-PC customs entry files. Calculates import duties across multiple tariff schedules (MFN, Section 301, Section 338, Section 122) with IEEPA validation.

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
| 13 | Country of Origin | ISO 2-letter code |
| **14-16** | **Seq 1: HTS/Value/Rate** | First sequence |
| **17-19** | **Seq 2: HTS/Value/Rate** | Second sequence |
| **20-22** | **Seq 3: HTS/Value/Rate** | Third sequence |
| **23-25** | **Seq 4: HTS/Value/Rate** | Fourth sequence |
| 26 | Status | D = Deleted (filtered out) |

### Multi-Sequence HTS Processing

Each line has up to 4 HTS sequences. Each sequence = HTS code + Value + Rate.

**Critical: Position Independence**
- MFN can be in ANY sequence position (1, 2, 3, or 4)
- Chapter 99 codes can be in ANY sequence positions
- Code scans all 4 sequences to dynamically identify MFN (first non-Chapter-9) and Chapter 99 codes

**Processing Algorithm:**

**Pass 1** - Find MFN Base:
```
For sequences in [14, 17, 20, 23]:
    If HTS does not start with "9" AND mfnValue not yet set:
        mfnValue = sequence.Value
        mfnRate = sequence.Rate
```

**Pass 2** - Accumulate Chapter 99 Rates:
```
For sequences in [14, 17, 20, 23]:
    If HTS matches Settings Column A (S301):
        s301Rate += sequence.Rate
    If HTS matches Settings Column B (S338):
        s338Rate += sequence.Rate
    If HTS matches Settings Column C (S122):
        s122Rate += sequence.Rate
```

**Calculate Duties:**
```
mfnDuty  = Qty × (mfnValue × mfnRate)
s301Duty = Qty × (mfnValue × s301Rate)
s338Duty = Qty × (mfnValue × s338Rate)
s122Duty = Qty × (mfnValue × s122Rate)
```

### Chapter 99 Special Handling

Chapter 99 codes (starting with "99") are trade remedy tariffs:
- Have Value = $0 (no independent value)
- Carry Rate information only
- Apply rates to MFN base value

**Example:**
```
Seq 1: 99030125 (S338), Value=$0, Rate=10%
Seq 2: 99030251 (S338), Value=$0, Rate=35%
Seq 3: 6204628011 (MFN), Value=$7.32, Rate=16.6%
Qty: 2

Processing:
  MFN Value = $7.32 (from Seq 3)
  Total S338 Rate = 10% + 35% = 45%
  
  MFN Duty  = 2 × ($7.32 × 0.166) = $2.430
  S338 Duty = 2 × ($7.32 × 0.45) = $6.588   ← Uses MFN value, not $0
  Total = $9.018
```

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

### Details Sheet Columns
```
A=Entry Number    J=MFN Duty      
B=Receipt Date    K=S301 Duty     
C=Txn Code        L=S338 Duty     
D=PO              M=S122 Duty     
E=NGC#            N=Cotton Fee    
F=Material Number O=MPF           
G=Qty             P=Total Fees    
H=Unit Value      Q=Total Duty    
I=Total Value     R=99 Value      
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

**MasterProcessAllFiles()** - Main entry, batch processing  
**ProcessSingleFile()** - Per-file orchestration  
**ProcessDutyCalculations()** - Loops through input rows  
**ProcessRow()** - Two-pass sequence scan (cols 14,17,20,23), duty calculations  
**IsHTSMatch()** - Checks if HTS exists in Settings columns A/B/C  
**ValidateIEEPADutyRate()** - Date-based matrix validation  
**DistributeFees()** - Distributes Cotton Fee and MPF across rows  
**CalculateSummaryTotals()** - Aggregates for Summary sheet  

## Critical Code Pattern

```vba
' PASS 1: Find MFN base (any sequence position)
For Each col In Array(14, 17, 20, 23)
    htsNum = Trim(CStr(ws_Input.Cells(inputRow, col).Value))
    If htsNum <> "" Then
        htsValue = Val(ws_Input.Cells(inputRow, col + 1).Value)
        htsRate = Val(ws_Input.Cells(inputRow, col + 2).Value)
        
        If Left(htsNum, 1) <> "9" And mfnHTSCode = "" Then
            mfnValue = htsValue  ' Base for ALL duty calculations
            mfnRate = htsRate
        End If
    End If
Next

' PASS 2: Accumulate Chapter 99 rates
For Each col In Array(14, 17, 20, 23)
    htsNum = Trim(CStr(ws_Input.Cells(inputRow, col).Value))
    If htsNum <> "" Then
        htsRate = Val(ws_Input.Cells(inputRow, col + 2).Value)
        
        If IsHTSMatch(htsNum, ws_Settings, 1) Then s301Rate = s301Rate + htsRate
        If IsHTSMatch(htsNum, ws_Settings, 2) Then s338Rate = s338Rate + htsRate
        If IsHTSMatch(htsNum, ws_Settings, 3) Then s122Rate = s122Rate + htsRate
    End If
Next

' Calculate using MFN base (NOT Chapter 99 values)
s301Duty = qty * (mfnValue * s301Rate)
s338Duty = qty * (mfnValue * s338Rate)
s122Duty = qty * (mfnValue * s122Rate)
```

## Validation Features

**IEEPA Duty Rate**: Validates Chapter 99 rates against date-based matrix (most recent date ≤ receipt date)  
**Unit Value Bounds**: Flags values outside H2/H3 thresholds  
**Total Duty Bounds**: Alerts when total duty outside N2/N3 range  
**99 Value Detection**: Flags non-zero Chapter 99 values (indicates potential overpayment)

## Installation

1. Import `FTZ_Duty_Calculator.bas` into VBA (Alt+F11)
2. Create UserForm `ProgressForm` with controls: `lblStatus`, `lblPhase`, `lblEntry`, `frameProgress`, `ProgressBar1`
3. Configure Settings sheet (columns A/B/C, cells E2/E3/H2/H3/I2/N2/N3)
4. Create folders: `Input/`, `Input/Archive/`, `Output/`
5. Run `MasterProcessAllFiles` (Alt+F8)

## Common Issues

**"Method or data member not found"** - ProgressForm missing controls or ProgressBar1 not inside frameProgress  
**"SharePoint/OneDrive web"** - Copy project to local drive  
**Calculations = $0** - Check Settings columns A/B/C have correct HTS codes  
**Duty Matrix not importing** - Verify path in Settings I2, file exists, data starts at A1

## Repository

https://github.com/wdhobgood/CDKReport

MIT License - See LICENSE file
