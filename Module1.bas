' ============================================================
' 9903 Analyzer - Module1.bas
' Version : 2026-04-10 (CBP ES003-aligned duty calculation)
' Modified: 2026-04-10
' Changes : - Fix entered value: SUM of ALL HTS sequence values, not MFN only
'           - Aggregate transactions into ES003 lines by MFN HTS + MID + COO + Ch99 codes
'           - Round entered value to whole dollars before duty calc
'           - Per-ordinal rounding: WorksheetFunction.Round(EV * rate, 2)
'           - Bulk array I/O for performance (189K rows in seconds)
'           - Dictionary-based HTS category lookups (O(1) vs O(n))
'           - Proportional fee distribution by entered value
' ============================================================
Option Explicit

' === Configuration Flags ===
Const ENABLE_S338 As Boolean = False    ' Set to True to re-enable Section 338 duty calculations

Public g_99ValueWarnings As Collection

' === ES003 Line Aggregation Type ===
Type ES003LineData
    GroupKey As String
    MfnHTSCode As String
    ManufacturerID As String
    CountryOfOrigin As String
    Ch99CodesSorted As String
    MfnRate As Double
    S301Rate As Double
    S338Rate As Double
    S122Rate As Double
    TotalEnteredValue As Double
    TotalQty As Double
    TransactionCount As Long
    EntryNumber As String
    ReceiptDate As Variant
    Ch99Value As Double
End Type

Sub MasterProcessAllFiles()
    Dim startTime As Double, endTime As Double, workbookPath As String
    Dim inputFolderPath As String, archiveFolderPath As String, outputFolderPath As String
    Dim fileName As String, fileList() As String, fileCount As Long, i As Long, mainWb As Workbook
    Dim importResult As String, warningMessage As String, item As Variant

    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    On Error GoTo ErrorHandler

    Set mainWb = ThisWorkbook
    workbookPath = mainWb.Path

    If Left(workbookPath, 4) = "http" Then
        MsgBox "This file is opened from SharePoint/OneDrive web." & vbCrLf & vbCrLf & _
               "Please copy the ENTIRE folder to your Desktop and run from there.", _
               vbExclamation, "Copy Folder to Desktop Required"
        GoTo CleanUp
    End If

    If Right(workbookPath, 1) <> "\" Then workbookPath = workbookPath & "\"
    inputFolderPath = workbookPath & "Input\"
    archiveFolderPath = inputFolderPath & "Archive\"
    outputFolderPath = workbookPath & "Output\"

    If Not FolderExistsLocal(inputFolderPath) Or Not FolderExistsLocal(archiveFolderPath) Or Not FolderExistsLocal(outputFolderPath) Then
        MsgBox "Required folders not found.", vbCritical
        GoTo CleanUp
    End If

    Set g_99ValueWarnings = New Collection

    Load ProgressForm
    ProgressForm.Show vbModeless

    Call UpdateProgress("Importing Duty Matrix...", "", 0, 1)
    importResult = ImportDutyMatrix(mainWb)
    If importResult <> "SUCCESS" Then
        MsgBox "Warning: Could not import Duty Matrix." & vbCrLf & _
               "Error: " & importResult & vbCrLf & vbCrLf & _
               "IEEPA validations will be skipped." & vbCrLf & _
               "Check Settings sheet cell I2 for the correct file path.", vbExclamation
    End If

    ReDim fileList(0)
    fileName = Dir(inputFolderPath & "*.xls*")
    fileCount = 0
    Do While fileName <> ""
        If Not Left(fileName, 1) = "~" Then
            ReDim Preserve fileList(fileCount)
            fileList(fileCount) = fileName
            fileCount = fileCount + 1
        End If
        fileName = Dir()
    Loop

    If fileCount = 0 Then
        MsgBox "No Excel files found in Input folder.", vbInformation
        GoTo CleanUp
    End If

    For i = 0 To fileCount - 1
        Call ProcessSingleFile(inputFolderPath & fileList(i), mainWb, outputFolderPath, archiveFolderPath, i + 1, fileCount)
    Next i

    endTime = Timer
    Call LogMasterExecution(mainWb.Sheets("Log"), startTime, endTime, fileCount)

    Unload ProgressForm

    If g_99ValueWarnings.Count > 0 Then
        warningMessage = "WARNING: The following entries have 99 Value greater than $0:" & vbCrLf & vbCrLf
        For Each item In g_99ValueWarnings
            warningMessage = warningMessage & item & vbCrLf
        Next item
        warningMessage = warningMessage & vbCrLf & "This indicates that Chapter 99 HTS sequences have non-zero values." & vbCrLf & _
                        "Please verify these are correct to avoid overpaying duties."
        MsgBox warningMessage, vbExclamation, "99 Value Warning"
    End If

CleanUp:
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If fileCount > 0 Then MsgBox "Processing completed! Check Log sheet for warnings.", vbInformation
    Exit Sub
ErrorHandler:
    Unload ProgressForm
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub UpdateProgress(phase As String, entryNumber As String, currentFile As Long, totalFiles As Long)
    On Error Resume Next
    If currentFile > 0 And totalFiles > 0 Then
        ProgressForm.lblStatus.Caption = "Processing File " & currentFile & " of " & totalFiles
        ProgressForm.lblPhase.Caption = phase
        ProgressForm.lblEntry.Caption = "Entry: " & entryNumber
        ProgressForm.ProgressBar1.Width = (ProgressForm.frameProgress.Width - 4) * (currentFile / totalFiles)
    Else
        ProgressForm.lblStatus.Caption = phase
        ProgressForm.lblPhase.Caption = ""
        ProgressForm.lblEntry.Caption = ""
        ProgressForm.ProgressBar1.Width = 0
    End If
    DoEvents
    On Error GoTo 0
End Sub

Function FolderExistsLocal(folderPath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        FolderExistsLocal = fso.FolderExists(folderPath)
        Set fso = Nothing
    Else
        FolderExistsLocal = (Dir(folderPath, vbDirectory) <> "")
    End If
    On Error GoTo 0
End Function

Function ImportDutyMatrix(mainWb As Workbook) As String
    Dim dutyMatrixPath As String, sourceWb As Workbook
    Dim ws_DutyMatrix As Worksheet, ws_Settings As Worksheet
    Dim lastRow As Long, lastCol As Long

    On Error GoTo ErrorHandler

    Set ws_Settings = mainWb.Sheets("Settings")
    Set ws_DutyMatrix = mainWb.Sheets("Duty Matrix")

    dutyMatrixPath = Trim(ws_Settings.Range("I2").value)

    If dutyMatrixPath = "" Then
        ImportDutyMatrix = "Path in Settings I2 is empty"
        Exit Function
    End If

    If Dir(dutyMatrixPath) = "" Then
        ImportDutyMatrix = "File not found at: " & dutyMatrixPath
        Exit Function
    End If

    Set sourceWb = Workbooks.Open(dutyMatrixPath, ReadOnly:=True, UpdateLinks:=False)

    With sourceWb.Sheets(1).UsedRange
        lastRow = .Rows.Count
        lastCol = .Columns.Count
    End With

    ws_DutyMatrix.Cells.Clear

    If lastRow >= 1 And lastCol >= 1 Then
        sourceWb.Sheets(1).UsedRange.Copy
        ws_DutyMatrix.Range("A1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        If ws_DutyMatrix.Range("A1").value = "" Then
            sourceWb.Close SaveChanges:=False
            ImportDutyMatrix = "Data paste failed - A1 is empty"
            Exit Function
        End If
    End If

    sourceWb.Close SaveChanges:=False
    ImportDutyMatrix = "SUCCESS"
    Exit Function

ErrorHandler:
    If Not sourceWb Is Nothing Then
        On Error Resume Next
        sourceWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    ImportDutyMatrix = "Error: " & Err.Description
End Function

Function ClearDutyMatrix(mainWb As Workbook)
    Dim ws_DutyMatrix As Worksheet
    On Error Resume Next
    Set ws_DutyMatrix = mainWb.Sheets("Duty Matrix")
    If Not ws_DutyMatrix Is Nothing Then ws_DutyMatrix.Cells.Clear
    On Error GoTo 0
End Function

Sub ProcessSingleFile(sourceFilePath As String, mainWb As Workbook, outputFolderPath As String, archiveFolderPath As String, currentFile As Long, totalFiles As Long)
    Dim sourceWb As Workbook, ws_Input As Worksheet, ws_Output As Worksheet, ws_Summary As Worksheet
    Dim lastRow As Long, entryNumber As String, newWb As Workbook, newFilePath As String
    Dim sourceFileName As String, removedDCount As Long, totalDuty As Double
    Dim total99Value As Double

    On Error GoTo ErrorHandler
    sourceFileName = Dir(sourceFilePath)

    Call UpdateProgress("Opening file...", "", currentFile, totalFiles)
    Set sourceWb = Workbooks.Open(sourceFilePath, ReadOnly:=False, UpdateLinks:=False)

    Set ws_Input = mainWb.Sheets("Input")
    Set ws_Output = mainWb.Sheets("Output")
    Set ws_Summary = mainWb.Sheets("Summary")

    Call UpdateProgress("Loading data...", "", currentFile, totalFiles)
    lastRow = sourceWb.Sheets(1).Cells(sourceWb.Sheets(1).Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        sourceWb.Sheets(1).Rows("2:" & lastRow).Copy
        ws_Input.Rows("2:2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If

    entryNumber = Trim(CStr(ws_Input.Cells(2, 3).value))

    Call UpdateProgress("Validating data...", entryNumber, currentFile, totalFiles)
    removedDCount = RemoveDStatusRows(ws_Input, entryNumber, mainWb.Sheets("Log"))
    ws_Summary.Range("B2").value = entryNumber

    Call UpdateProgress("Calculating duties...", entryNumber, currentFile, totalFiles)
    Call ProcessDutyCalculations(mainWb)
    Call DistributeFees(ws_Output, mainWb.Sheets("Settings"))
    Call CalculateSummaryTotals(ws_Output, ws_Summary)

    totalDuty = Val(ws_Summary.Range("B5").value)
    total99Value = Val(ws_Summary.Range("B15").value)
    Call ValidateTotalDuty(mainWb, entryNumber, totalDuty)

    If total99Value > 0 Then
        g_99ValueWarnings.Add "Entry " & entryNumber & ": $" & Format(total99Value, "#,##0.00")
    End If

    Call UpdateProgress("Creating output file...", entryNumber, currentFile, totalFiles)
    If entryNumber = "" Then entryNumber = "UNKNOWN_" & Format(Now(), "yyyymmdd_hhnnss")
    Set newWb = Workbooks.Add
    Application.DisplayAlerts = False
    Do While newWb.Sheets.Count > 1
        newWb.Sheets(newWb.Sheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    newWb.Sheets(1).Name = "Summary"
    ws_Summary.Range("A1:B15").Copy
    newWb.Sheets("Summary").Range("A1").PasteSpecial Paste:=xlPasteValues
    newWb.Sheets("Summary").Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    newWb.Sheets("Summary").Columns("A:B").AutoFit
    newWb.Sheets.Add After:=newWb.Sheets(newWb.Sheets.Count)
    newWb.Sheets(newWb.Sheets.Count).Name = "Details"
    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 1 Then
        ws_Output.Range("A1").CurrentRegion.Copy
        newWb.Sheets("Details").Range("A1").PasteSpecial Paste:=xlPasteValues
        newWb.Sheets("Details").Range("A1").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        newWb.Sheets("Details").Columns.AutoFit
    End If

    Call UpdateProgress("Saving output file...", entryNumber, currentFile, totalFiles)
    newFilePath = outputFolderPath & entryNumber & "_Weekly_FTZ_ACH_Upload.xlsx"
    On Error Resume Next
    If Dir(newFilePath) <> "" Then Kill newFilePath
    On Error GoTo ErrorHandler
    newWb.SaveAs fileName:=newFilePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Call UpdateProgress("Cleaning up...", entryNumber, currentFile, totalFiles)
    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws_Output.Rows("2:" & lastRow).ClearContents
    ws_Summary.Range("B2,B4:B6,B9:B12,B15").ClearContents
    lastRow = ws_Input.Cells(ws_Input.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws_Input.Rows("2:" & lastRow).ClearContents
    sourceWb.Close SaveChanges:=False
    Set sourceWb = Nothing
    Dim archiveFilePath As String
    archiveFilePath = archiveFolderPath & sourceFileName
    On Error Resume Next
    If Dir(archiveFilePath) <> "" Then Kill archiveFilePath
    On Error GoTo ErrorHandler
    Name sourceFilePath As archiveFilePath
    Exit Sub
ErrorHandler:
    MsgBox "Error processing file: " & Err.Description, vbCritical
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
End Sub

Function RemoveDStatusRows(ws_Input As Worksheet, entryNumber As String, ws_Log As Worksheet) As Long
    Dim lastRow As Long, i As Long, statusValue As String, removedCount As Long
    removedCount = 0
    lastRow = ws_Input.Cells(ws_Input.Rows.Count, 1).End(xlUp).Row
    For i = lastRow To 2 Step -1
        If ws_Input.Cells(i, 1).value <> "" Then
            statusValue = Trim(UCase(CStr(ws_Input.Cells(i, 26).value)))
            If statusValue = "D" Then
                ws_Input.Rows(i).Delete
                removedCount = removedCount + 1
            End If
        End If
    Next i
    If removedCount > 0 Then
        Call LogMessage(ws_Log, entryNumber, "", "", "", "", "", CDbl(removedCount), removedCount & " 'D' status rows removed.")
    End If
    RemoveDStatusRows = removedCount
End Function

' === Fee Distribution (proportional by entered value) ===
Sub DistributeFees(ws_Output As Worksheet, ws_Settings As Worksheet)
    Dim lastRow As Long, i As Long, totalMPF As Double, totalCottonFee As Double
    Dim totalEnteredValue As Double, proportion As Double, totalDuty As Double

    totalCottonFee = Val(ws_Settings.Range("E2").value)
    totalMPF = Val(ws_Settings.Range("E3").value)
    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    ' Read entered values into array for performance
    Dim evData As Variant
    evData = ws_Output.Range(ws_Output.Cells(2, 8), ws_Output.Cells(lastRow, 8)).value

    ' Sum total entered value
    For i = 1 To UBound(evData, 1)
        totalEnteredValue = totalEnteredValue + Val(evData(i, 1))
    Next i

    If totalEnteredValue = 0 Then Exit Sub

    ' Build fee output array: Cotton Fee (col 15/O), MPF (col 16/P), Total Fees (col 18/R), Total (col 19/S)
    Dim rowCount As Long
    rowCount = lastRow - 1
    Dim feeData() As Variant
    ReDim feeData(1 To rowCount, 1 To 4)

    ' Read total duty column (col 17/Q) for grand total calc
    Dim dutyData As Variant
    dutyData = ws_Output.Range(ws_Output.Cells(2, 17), ws_Output.Cells(lastRow, 17)).value

    For i = 1 To rowCount
        proportion = Val(evData(i, 1)) / totalEnteredValue
        feeData(i, 1) = totalCottonFee * proportion      ' Cotton Fee -> col 15 (O)
        feeData(i, 2) = totalMPF * proportion             ' MPF -> col 16 (P)
        feeData(i, 3) = feeData(i, 1) + feeData(i, 2)    ' Total Fees -> col 18 (R)
        feeData(i, 4) = Val(dutyData(i, 1)) + feeData(i, 3) ' Total -> col 19 (S)
    Next i

    ' Write fees in bulk (columns O, P, R, S)
    ws_Output.Range(ws_Output.Cells(2, 15), ws_Output.Cells(lastRow, 15)).value = Application.Index(feeData, 0, 1)
    ws_Output.Range(ws_Output.Cells(2, 16), ws_Output.Cells(lastRow, 16)).value = Application.Index(feeData, 0, 2)
    ws_Output.Range(ws_Output.Cells(2, 18), ws_Output.Cells(lastRow, 18)).value = Application.Index(feeData, 0, 3)
    ws_Output.Range(ws_Output.Cells(2, 19), ws_Output.Cells(lastRow, 19)).value = Application.Index(feeData, 0, 4)
End Sub

' === Summary Totals (reads from aggregated output) ===
Sub CalculateSummaryTotals(ws_Output As Worksheet, ws_Summary As Worksheet)
    Dim lastRow As Long, i As Long
    Dim totalMFN As Double, totalS301 As Double, totalS338 As Double, totalS122 As Double
    Dim totalDuty As Double, totalFees As Double, totalDutyPlusFees As Double, total99Value As Double

    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Read all needed columns into arrays for performance
    Dim outData As Variant
    outData = ws_Output.Range(ws_Output.Cells(2, 1), ws_Output.Cells(lastRow, 19)).value

    For i = 1 To UBound(outData, 1)
        totalMFN = totalMFN + Val(outData(i, 9))        ' Col I: MFN Duty
        totalS301 = totalS301 + Val(outData(i, 10))      ' Col J: S301 Duty
        totalS338 = totalS338 + Val(outData(i, 11))      ' Col K: S338 Duty
        totalS122 = totalS122 + Val(outData(i, 12))      ' Col L: S122 Duty
        totalDuty = totalDuty + Val(outData(i, 17))       ' Col Q: Total Duty
        totalFees = totalFees + Val(outData(i, 18))       ' Col R: Total Fees
        total99Value = total99Value + Val(outData(i, 13)) ' Col M: 99 Value
    Next i

    totalDutyPlusFees = totalDuty + totalFees
    ws_Summary.Range("B9").value = totalMFN
    ws_Summary.Range("B10").value = totalS301
    ws_Summary.Range("B11").value = totalS338
    ws_Summary.Range("B12").value = totalS122
    ws_Summary.Range("B5").value = totalDuty
    ws_Summary.Range("B6").value = totalDutyPlusFees
    ws_Summary.Range("B15").value = total99Value
End Sub

' =============================================================================
' ProcessDutyCalculations — CBP ES003-aligned two-phase aggregation engine
' Phase 1: Scan all input rows, aggregate into ES003 lines by grouping key
' Phase 2: Calculate duty per aggregated line using CBP formula with rounding
' =============================================================================
Sub ProcessDutyCalculations(mainWb As Workbook)
    Dim startTime As Double, endTime As Double
    Dim ws_Input As Worksheet, ws_Output As Worksheet, ws_Settings As Worksheet, ws_Log As Worksheet
    Dim lastRow As Long, i As Long, j As Long, recordCount As Long

    startTime = Timer
    On Error GoTo ErrorHandler

    Set ws_Input = mainWb.Sheets("Input")
    Set ws_Output = mainWb.Sheets("Output")
    Set ws_Settings = mainWb.Sheets("Settings")
    Set ws_Log = mainWb.Sheets("Log")

    Call FormatColumnsAsText(ws_Input, ws_Settings)

    ' Clear existing output
    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws_Output.Rows("2:" & lastRow).ClearContents
    Call EnsureOutputHeaders(ws_Output)

    lastRow = ws_Input.Cells(ws_Input.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Call LogExecution(ws_Log, startTime, Timer, 0)
        Exit Sub
    End If

    ' === PERFORMANCE: Bulk read all input data into array ===
    Dim inputData As Variant
    inputData = ws_Input.Range(ws_Input.Cells(2, 1), ws_Input.Cells(lastRow, 26)).value
    Dim inputRowCount As Long
    inputRowCount = UBound(inputData, 1)

    ' === PERFORMANCE: Pre-load Settings HTS lists into Dictionaries ===
    Dim dictS301 As Object, dictS338 As Object, dictS122 As Object
    Set dictS301 = CreateObject("Scripting.Dictionary")
    Set dictS338 = CreateObject("Scripting.Dictionary")
    Set dictS122 = CreateObject("Scripting.Dictionary")
    Call LoadHTSDictionary(ws_Settings, 1, dictS301)
    Call LoadHTSDictionary(ws_Settings, 2, dictS338)
    Call LoadHTSDictionary(ws_Settings, 3, dictS122)

    ' === PHASE 1: Scan & Aggregate ===
    Dim dict As Object  ' Scripting.Dictionary: groupKey -> index in lines()
    Set dict = CreateObject("Scripting.Dictionary")
    Dim lines() As ES003LineData
    Dim lineCount As Long
    lineCount = 0
    ReDim lines(0 To 999)

    ' Validation log buffer
    Dim logMessages() As String
    Dim logCount As Long
    logCount = 0
    ReDim logMessages(0 To 999)

    Dim htsNum As String, htsValue As Double, htsRate As Double
    Dim mfnHTSCode As String, mfnRate As Double, mfnValue As Double
    Dim s301Rate As Double, s338Rate As Double, s122Rate As Double
    Dim sumAllSeqValues As Double, ch99Value As Double
    Dim qty As Double, entryNumber As String, manufacturerID As String, countryOfOrigin As String
    Dim receiptDate As Variant
    Dim sortedCh99 As String
    Dim poNum As String, ngcNum As String, matNum As String
    Dim ch99HTSCodes(0 To 3) As String, ch99Rates(0 To 3) As Double, chapter99Count As Long
    Dim groupKey As String, idx As Long
    Dim col As Long, colIdx As Long

    ' Unit value validation bounds (read once)
    Dim uvUpperBound As Double, uvLowerBound As Double
    uvUpperBound = Val(ws_Settings.Range("N2").value)
    uvLowerBound = Val(ws_Settings.Range("N3").value)

    recordCount = 0

    For i = 1 To inputRowCount
        If CStr(inputData(i, 1)) <> "" Then
            recordCount = recordCount + 1

            ' --- Extract fields from array ---
            entryNumber = Trim(CStr(inputData(i, 3)))
            manufacturerID = Trim(CStr(inputData(i, 12)))
            countryOfOrigin = Trim(CStr(inputData(i, 13)))
            qty = Val(inputData(i, 11))

            On Error Resume Next
            receiptDate = inputData(i, 2)
            If Not IsDate(receiptDate) Then
                If IsNumeric(receiptDate) And CDbl(receiptDate) > 1 Then
                    receiptDate = CDate(CDbl(receiptDate))
                End If
            End If
            On Error GoTo ErrorHandler

            ' --- Pass 1: Find MFN base and sum ALL sequence values ---
            mfnHTSCode = ""
            mfnRate = 0
            mfnValue = 0
            sumAllSeqValues = 0
            s301Rate = 0
            s338Rate = 0
            s122Rate = 0
            ch99Value = 0
            chapter99Count = 0

            For colIdx = 0 To 3
                col = Choose(colIdx + 1, 14, 17, 20, 23)
                htsNum = Trim(CStr(inputData(i, col)))
                If htsNum <> "" Then
                    htsValue = Val(inputData(i, col + 1))
                    htsRate = Val(inputData(i, col + 2))

                    ' Sum ALL sequence values (Bug 1 fix)
                    sumAllSeqValues = sumAllSeqValues + htsValue

                    ' Identify MFN (first non-chapter-99)
                    If Left(htsNum, 2) <> "99" And mfnHTSCode = "" Then
                        mfnRate = htsRate
                        mfnValue = htsValue
                        mfnHTSCode = htsNum
                    End If

                    ' Categorize chapter 99 codes
                    If dictS301.Exists(htsNum) Then
                        s301Rate = s301Rate + htsRate
                        If htsValue > 0 Then ch99Value = ch99Value + htsValue
                    End If
                    If ENABLE_S338 Then
                        If dictS338.Exists(htsNum) Then
                            s338Rate = s338Rate + htsRate
                            If htsValue > 0 Then ch99Value = ch99Value + htsValue
                        End If
                    End If
                    If dictS122.Exists(htsNum) Then
                        s122Rate = s122Rate + htsRate
                        If htsValue > 0 Then ch99Value = ch99Value + htsValue
                    End If

                    ' Track ch99 codes for grouping key and IEEPA validation
                    If Left(htsNum, 2) = "99" Then
                        ch99HTSCodes(chapter99Count) = htsNum
                        ch99Rates(chapter99Count) = htsRate
                        chapter99Count = chapter99Count + 1
                    End If
                End If
            Next colIdx

            ' --- Build grouping key ---
            sortedCh99 = SortCh99Codes(ch99HTSCodes, chapter99Count)
            groupKey = BuildGroupKey(mfnHTSCode, manufacturerID, countryOfOrigin, sortedCh99)

            ' --- Aggregate into ES003 line ---
            If dict.Exists(groupKey) Then
                idx = dict(groupKey)
                lines(idx).TotalEnteredValue = lines(idx).TotalEnteredValue + (qty * sumAllSeqValues)
                lines(idx).TotalQty = lines(idx).TotalQty + qty
                lines(idx).TransactionCount = lines(idx).TransactionCount + 1
                lines(idx).Ch99Value = lines(idx).Ch99Value + (qty * ch99Value)

                ' Rate consistency check
                If Abs(lines(idx).MfnRate - mfnRate) > 0.0001 Then
                    If logCount > UBound(logMessages) Then ReDim Preserve logMessages(0 To logCount + 999)
                    logMessages(logCount) = "RATE_WARN|" & entryNumber & "||||||" & qty & "|MFN rate mismatch in group " & groupKey & ": " & Format(mfnRate, "0.0000") & " vs " & Format(lines(idx).MfnRate, "0.0000")
                    logCount = logCount + 1
                End If
            Else
                ' New group
                If lineCount > UBound(lines) Then ReDim Preserve lines(0 To lineCount + 999)
                lines(lineCount).GroupKey = groupKey
                lines(lineCount).MfnHTSCode = mfnHTSCode
                lines(lineCount).ManufacturerID = manufacturerID
                lines(lineCount).CountryOfOrigin = countryOfOrigin
                lines(lineCount).Ch99CodesSorted = sortedCh99
                lines(lineCount).MfnRate = mfnRate
                lines(lineCount).S301Rate = s301Rate
                lines(lineCount).S338Rate = s338Rate
                lines(lineCount).S122Rate = s122Rate
                lines(lineCount).TotalEnteredValue = qty * sumAllSeqValues
                lines(lineCount).TotalQty = qty
                lines(lineCount).TransactionCount = 1
                lines(lineCount).EntryNumber = entryNumber
                lines(lineCount).ReceiptDate = receiptDate
                lines(lineCount).Ch99Value = qty * ch99Value
                dict.Add groupKey, lineCount
                lineCount = lineCount + 1
            End If

            ' --- Per-row validation: Unit Value ---
            If mfnValue > uvUpperBound Or mfnValue < uvLowerBound Then
                If logCount > UBound(logMessages) Then ReDim Preserve logMessages(0 To logCount + 999)
                poNum = Trim(CStr(inputData(i, 8)))
                ngcNum = Trim(CStr(inputData(i, 6)))
                matNum = Trim(CStr(inputData(i, 7)))
                logMessages(logCount) = "UV_WARN|" & entryNumber & "|" & poNum & "|" & ngcNum & "|" & matNum & "|" & mfnHTSCode & "|" & countryOfOrigin & "|" & qty & "|The Unit Value for this line is outside the bound - Unit Value: $" & Format(mfnValue, "#,##0.00")
                logCount = logCount + 1
            End If

            ' --- Per-row validation: IEEPA rate ---
            If chapter99Count > 0 And IsDate(receiptDate) Then
                Dim k As Long
                For k = 0 To chapter99Count - 1
                    Call ValidateIEEPADutyRate(mainWb, entryNumber, Trim(CStr(inputData(i, 8))), _
                        Trim(CStr(inputData(i, 6))), Trim(CStr(inputData(i, 7))), _
                        ch99HTSCodes(k), countryOfOrigin, qty, CDate(receiptDate), ch99Rates(k))
                Next k
            End If

            ' Clear ch99 arrays for next row
            For j = 0 To 3
                ch99HTSCodes(j) = ""
                ch99Rates(j) = 0
            Next j
        End If
    Next i

    ' === PHASE 2: Calculate duty per aggregated line & write output ===
    If lineCount = 0 Then
        Call LogExecution(ws_Log, startTime, Timer, recordCount)
        Exit Sub
    End If

    ' Build output array
    Dim numCols As Long
    numCols = 19  ' A through S
    Dim outputData() As Variant
    ReDim outputData(1 To lineCount, 1 To numCols)

    Dim enteredValue As Double, mfnDuty As Double, s301Duty As Double
    Dim s338Duty As Double, s122Duty As Double, totalDuty As Double

    For i = 0 To lineCount - 1
        ' CBP formula: round entered value to whole dollars
        enteredValue = Application.WorksheetFunction.Round(lines(i).TotalEnteredValue, 0)

        ' Per-ordinal duty with arithmetic rounding to cents
        mfnDuty = Application.WorksheetFunction.Round(enteredValue * lines(i).MfnRate, 2)
        s301Duty = Application.WorksheetFunction.Round(enteredValue * lines(i).S301Rate, 2)
        s338Duty = Application.WorksheetFunction.Round(enteredValue * lines(i).S338Rate, 2)
        s122Duty = Application.WorksheetFunction.Round(enteredValue * lines(i).S122Rate, 2)
        totalDuty = mfnDuty + s301Duty + s338Duty + s122Duty

        ' Write to output array
        outputData(i + 1, 1) = lines(i).EntryNumber           ' A: Entry Number
        outputData(i + 1, 2) = lines(i).MfnHTSCode            ' B: MFN HTS Code
        outputData(i + 1, 3) = lines(i).ManufacturerID        ' C: Manufacturer ID
        outputData(i + 1, 4) = lines(i).CountryOfOrigin       ' D: Country of Origin
        outputData(i + 1, 5) = lines(i).Ch99CodesSorted       ' E: Ch99 Codes
        outputData(i + 1, 6) = lines(i).TotalQty              ' F: Total Qty
        outputData(i + 1, 7) = lines(i).TransactionCount      ' G: Txn Count
        outputData(i + 1, 8) = enteredValue                   ' H: Entered Value ($)
        outputData(i + 1, 9) = mfnDuty                        ' I: MFN Duty
        outputData(i + 1, 10) = s301Duty                      ' J: S301 Duty
        outputData(i + 1, 11) = s338Duty                      ' K: S338 Duty
        outputData(i + 1, 12) = s122Duty                      ' L: S122 Duty
        outputData(i + 1, 13) = Application.WorksheetFunction.Round(lines(i).Ch99Value, 2) ' M: 99 Value
        ' N: (reserved)
        ' O-P: Fees (filled by DistributeFees)
        ' Q: Total Fees (filled by DistributeFees)
        outputData(i + 1, 17) = totalDuty                     ' Q: Total Duty
        ' R: (reserved)
        outputData(i + 1, 19) = totalDuty                     ' S: Total (updated by DistributeFees)
    Next i

    ' === PERFORMANCE: Bulk write output array ===
    ws_Output.Range(ws_Output.Cells(2, 1), ws_Output.Cells(lineCount + 1, numCols)).value = outputData

    ' === Write buffered log messages ===
    If logCount > 0 Then
        Dim parts() As String
        For i = 0 To logCount - 1
            parts = Split(logMessages(i), "|")
            If UBound(parts) >= 8 Then
                Call LogMessage(ws_Log, parts(1), parts(2), parts(3), parts(4), parts(5), parts(6), Val(parts(7)), parts(8))
            End If
        Next i
    End If

    endTime = Timer
    Call LogExecution(ws_Log, startTime, endTime, recordCount)
    Exit Sub

ErrorHandler:
    MsgBox "Error in ProcessDutyCalculations: " & Err.Description & " (Line: " & Erl & ")", vbCritical
End Sub

' === Output Headers (CBP ES003-aligned layout) ===
Sub EnsureOutputHeaders(ws_Output As Worksheet)
    If ws_Output.Range("A1").value = "" Then
        ws_Output.Range("A1").value = "Entry Number"
        ws_Output.Range("B1").value = "MFN HTS Code"
        ws_Output.Range("C1").value = "Manufacturer ID"
        ws_Output.Range("D1").value = "Country of Origin"
        ws_Output.Range("E1").value = "Ch99 Codes"
        ws_Output.Range("F1").value = "Total Qty"
        ws_Output.Range("G1").value = "Txn Count"
        ws_Output.Range("H1").value = "Entered Value"
        ws_Output.Range("I1").value = "MFN Duty"
        ws_Output.Range("J1").value = "S301 Duty"
        ws_Output.Range("K1").value = "S338 Duty"
        ws_Output.Range("L1").value = "S122 Duty"
        ws_Output.Range("M1").value = "99 Value"
        ws_Output.Range("N1").value = ""
        ws_Output.Range("O1").value = "Cotton Fee"
        ws_Output.Range("P1").value = "MPF"
        ws_Output.Range("Q1").value = "Total Duty"
        ws_Output.Range("R1").value = "Total Fees"
        ws_Output.Range("S1").value = "Total"
    End If
End Sub

Sub FormatColumnsAsText(ws_Input As Worksheet, ws_Settings As Worksheet)
    Dim htsColumns As Variant, col As Variant
    On Error Resume Next
    htsColumns = Array(14, 17, 20, 23)
    For Each col In htsColumns
        ws_Input.Columns(col).TextToColumns Destination:=ws_Input.Columns(col), DataType:=xlFixedWidth, FieldInfo:=Array(0, xlTextFormat)
    Next col
    ws_Settings.Columns(1).TextToColumns Destination:=ws_Settings.Columns(1), DataType:=xlFixedWidth, FieldInfo:=Array(0, xlTextFormat)
    ws_Settings.Columns(2).TextToColumns Destination:=ws_Settings.Columns(2), DataType:=xlFixedWidth, FieldInfo:=Array(0, xlTextFormat)
    ws_Settings.Columns(3).TextToColumns Destination:=ws_Settings.Columns(3), DataType:=xlFixedWidth, FieldInfo:=Array(0, xlTextFormat)
End Sub

' === Helper: Build ES003 line grouping key ===
Function BuildGroupKey(mfnHTSCode As String, manufacturerID As String, countryOfOrigin As String, ch99CodesSorted As String) As String
    BuildGroupKey = mfnHTSCode & "|" & Trim(UCase(manufacturerID)) & "|" & Trim(UCase(countryOfOrigin)) & "|" & ch99CodesSorted
End Function

' === Helper: Sort up to 4 ch99 codes and return pipe-delimited string ===
Function SortCh99Codes(codes() As String, count As Long) As String
    If count = 0 Then
        SortCh99Codes = ""
        Exit Function
    End If
    If count = 1 Then
        SortCh99Codes = codes(0)
        Exit Function
    End If
    ' Simple bubble sort for 2-4 elements
    Dim i As Long, j As Long, temp As String
    Dim sorted() As String
    ReDim sorted(0 To count - 1)
    For i = 0 To count - 1
        sorted(i) = codes(i)
    Next i
    For i = 0 To count - 2
        For j = 0 To count - i - 2
            If sorted(j) > sorted(j + 1) Then
                temp = sorted(j)
                sorted(j) = sorted(j + 1)
                sorted(j + 1) = temp
            End If
        Next j
    Next i
    Dim result As String
    result = sorted(0)
    For i = 1 To count - 1
        result = result & "|" & sorted(i)
    Next i
    SortCh99Codes = result
End Function

' === Helper: Load HTS codes from a Settings column into a Dictionary ===
Sub LoadHTSDictionary(ws_Settings As Worksheet, col As Long, dict As Object)
    Dim lastRow As Long, i As Long, val As String
    On Error Resume Next
    lastRow = ws_Settings.Cells(ws_Settings.Rows.Count, col).End(xlUp).Row
    On Error GoTo 0
    For i = 1 To lastRow
        val = Trim(CStr(ws_Settings.Cells(i, col).value))
        If val <> "" And Not dict.Exists(val) Then
            dict.Add val, True
        End If
    Next i
End Sub

' === Validation Functions (unchanged) ===

Sub ValidateTotalDuty(mainWb As Workbook, entryNumber As String, totalDuty As Double)
    Dim ws_Settings As Worksheet, ws_Log As Worksheet, upperBound As Double, lowerBound As Double, message As String
    Set ws_Settings = mainWb.Sheets("Settings")
    Set ws_Log = mainWb.Sheets("Log")
    upperBound = Val(ws_Settings.Range("H2").value)
    lowerBound = Val(ws_Settings.Range("H3").value)
    If totalDuty > upperBound Or totalDuty < lowerBound Then
        message = "Warning: Total duty on entry is outside the limit - $" & Format(totalDuty, "#,##0.00")
        Call LogMessage(ws_Log, entryNumber, "", "", "", "", "", 0, message)
    End If
End Sub

Sub ValidateUnitValue(mainWb As Workbook, entryNumber As String, poNumber As String, ngcNumber As String, materialNumber As String, htsCode As String, countryOfOrigin As String, qty As Double, unitValue As Double)
    Dim ws_Settings As Worksheet, ws_Log As Worksheet, upperBound As Double, lowerBound As Double, message As String
    Set ws_Settings = mainWb.Sheets("Settings")
    Set ws_Log = mainWb.Sheets("Log")
    upperBound = Val(ws_Settings.Range("N2").value)
    lowerBound = Val(ws_Settings.Range("N3").value)
    If unitValue > upperBound Or unitValue < lowerBound Then
        message = "The Unit Value for this line is outside the bound - Unit Value: $" & Format(unitValue, "#,##0.00")
        Call LogMessage(ws_Log, entryNumber, poNumber, ngcNumber, materialNumber, htsCode, countryOfOrigin, qty, message)
    End If
End Sub

Sub ValidateIEEPADutyRate(mainWb As Workbook, entryNumber As String, poNumber As String, ngcNumber As String, materialNumber As String, ieepaHTSCode As String, countryOfOrigin As String, qty As Double, receiptDate As Date, ieepaRate As Double)
    Dim ws_DutyMatrix As Worksheet, ws_Log As Worksheet, lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim matrixISOCode As String, matrixIEEPARate As Double, message As String, found As Boolean
    Dim headerValue As Variant, headerDate As Date, bestCol As Long, bestDate As Date
    Dim tempDate As Variant

    Set ws_DutyMatrix = mainWb.Sheets("Duty Matrix")
    Set ws_Log = mainWb.Sheets("Log")

    If ws_DutyMatrix.Range("A1").value = "" Then Exit Sub

    lastRow = ws_DutyMatrix.UsedRange.Rows.Count
    lastCol = ws_DutyMatrix.UsedRange.Columns.Count
    found = False

    bestCol = 0
    bestDate = DateSerial(1900, 1, 1)

    For j = 3 To lastCol
        headerValue = ws_DutyMatrix.Cells(1, j).value

        On Error Resume Next
        tempDate = Null

        If IsDate(headerValue) Then
            tempDate = CDate(headerValue)
        End If

        If IsNull(tempDate) And IsNumeric(headerValue) Then
            tempDate = CDate(headerValue)
        End If

        If IsNull(tempDate) And VarType(headerValue) = vbString Then
            tempDate = DateValue(headerValue)
        End If

        On Error GoTo 0

        If Not IsNull(tempDate) And IsDate(tempDate) Then
            headerDate = CDate(tempDate)
            If headerDate <= receiptDate And headerDate > bestDate Then
                bestDate = headerDate
                bestCol = j
            End If
        End If
    Next j

    If bestCol = 0 Then
        message = "No applicable date column found in Duty Matrix for Receipt Date: " & Format(receiptDate, "mm/dd/yyyy")
        Call LogMessage(ws_Log, entryNumber, poNumber, ngcNumber, materialNumber, ieepaHTSCode, countryOfOrigin, qty, message)
        Exit Sub
    End If

    For i = 2 To lastRow
        matrixISOCode = Trim(UCase(CStr(ws_DutyMatrix.Cells(i, 2).value)))

        If matrixISOCode = UCase(countryOfOrigin) Then
            matrixIEEPARate = Val(ws_DutyMatrix.Cells(i, bestCol).value)
            found = True

            If ieepaRate < matrixIEEPARate - 0.0001 Then
                message = "Warning: IEEPA Duty on Line (" & Format(ieepaRate, "0.0%") & ") is less than IEEPA duty on Duty Matrix Table (" & Format(matrixIEEPARate, "0.0%") & ") - Receipt Date: " & Format(receiptDate, "mm/dd/yyyy")
                Call LogMessage(ws_Log, entryNumber, poNumber, ngcNumber, materialNumber, ieepaHTSCode, countryOfOrigin, qty, message)
            ElseIf ieepaRate > matrixIEEPARate + 0.0001 Then
                message = "Warning: IEEPA Duty on Line (" & Format(ieepaRate, "0.0%") & ") is greater than IEEPA Duty on Duty Matrix Table (" & Format(matrixIEEPARate, "0.0%") & ") - Receipt Date: " & Format(receiptDate, "mm/dd/yyyy")
                Call LogMessage(ws_Log, entryNumber, poNumber, ngcNumber, materialNumber, ieepaHTSCode, countryOfOrigin, qty, message)
            End If
            Exit For
        End If
    Next i

    If Not found And countryOfOrigin <> "" Then
        message = "ISO Code '" & countryOfOrigin & "' not found in Duty Matrix - Receipt Date: " & Format(receiptDate, "mm/dd/yyyy")
        Call LogMessage(ws_Log, entryNumber, poNumber, ngcNumber, materialNumber, ieepaHTSCode, countryOfOrigin, qty, message)
    End If
End Sub

' === Logging Functions ===

Sub LogMessage(ws_Log As Worksheet, entryNumber As String, poNumber As String, ngcNumber As String, materialNumber As String, htsCode As String, countryOfOrigin As String, qty As Double, message As String)
    Dim lastLogRow As Long
    If ws_Log.Cells(1, 1).value = "" Then
        lastLogRow = 1
    Else
        lastLogRow = ws_Log.Cells(ws_Log.Rows.Count, 1).End(xlUp).Row + 1
    End If
    ws_Log.Cells(lastLogRow, 1).value = Date
    ws_Log.Cells(lastLogRow, 2).value = Time
    ws_Log.Cells(lastLogRow, 3).value = entryNumber
    ws_Log.Cells(lastLogRow, 4).value = poNumber
    ws_Log.Cells(lastLogRow, 5).value = ngcNumber
    ws_Log.Cells(lastLogRow, 6).value = materialNumber
    ws_Log.Cells(lastLogRow, 7).value = htsCode
    ws_Log.Cells(lastLogRow, 8).value = countryOfOrigin
    ws_Log.Cells(lastLogRow, 9).value = qty
    ws_Log.Cells(lastLogRow, 10).value = message
End Sub

Sub LogExecution(ws_Log As Worksheet, startTime As Double, endTime As Double, recordCount As Long)
    Dim duration As Double, durationMinutes As String
    duration = (endTime - startTime) / 60
    durationMinutes = Format(duration, "0.00")
    Call LogMessage(ws_Log, "", "", "", "", "", "", CDbl(recordCount), "Program completed successfully within " & durationMinutes & " minutes.")
End Sub

Sub LogMasterExecution(ws_Log As Worksheet, startTime As Double, endTime As Double, fileCount As Long)
    Dim duration As Double, durationMinutes As String
    duration = (endTime - startTime) / 60
    durationMinutes = Format(duration, "0.00")
    Call LogMessage(ws_Log, "", "", "", "", "", "", CDbl(fileCount), "Master process completed: " & fileCount & " file(s) processed within " & durationMinutes & " minutes.")
End Sub
