'V3-5-26
Option Explicit

Public g_99ValueWarnings As Collection

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

Sub DistributeFees(ws_Output As Worksheet, ws_Settings As Worksheet)
    Dim lastRow As Long, i As Long, totalMPF As Double, totalCottonFee As Double
    Dim mpfPerRow As Double, cottonFeePerRow As Double, rowCount As Long, totalDuty As Double
    totalCottonFee = Val(ws_Settings.Range("E2").value)
    totalMPF = Val(ws_Settings.Range("E3").value)
    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row
    rowCount = lastRow - 1
    If rowCount > 0 Then
        cottonFeePerRow = totalCottonFee / rowCount
        mpfPerRow = totalMPF / rowCount
        For i = 2 To lastRow
            ws_Output.Cells(i, 14).value = cottonFeePerRow
            ws_Output.Cells(i, 15).value = mpfPerRow
            ws_Output.Cells(i, 16).value = cottonFeePerRow + mpfPerRow
            totalDuty = Val(ws_Output.Cells(i, 17).value)
            ws_Output.Cells(i, 19).value = totalDuty + cottonFeePerRow + mpfPerRow
        Next i
    End If
End Sub

Sub CalculateSummaryTotals(ws_Output As Worksheet, ws_Summary As Worksheet)
    Dim lastRow As Long, i As Long, totalMFN As Double, totalS301 As Double, totalS338 As Double, totalS122 As Double
    Dim totalDuty As Double, totalFees As Double, totalDutyPlusFees As Double, total99Value As Double
    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        totalMFN = totalMFN + Val(ws_Output.Cells(i, 10).value)
        totalS301 = totalS301 + Val(ws_Output.Cells(i, 11).value)
        totalS338 = totalS338 + Val(ws_Output.Cells(i, 12).value)
        totalS122 = totalS122 + Val(ws_Output.Cells(i, 13).value)
        totalDuty = totalDuty + Val(ws_Output.Cells(i, 17).value)
        totalFees = totalFees + Val(ws_Output.Cells(i, 16).value)
        total99Value = total99Value + Val(ws_Output.Cells(i, 18).value)
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

Sub ProcessDutyCalculations(mainWb As Workbook)
    Dim startTime As Double, endTime As Double
    Dim ws_Input As Worksheet, ws_Output As Worksheet, ws_Settings As Worksheet, ws_Log As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long, recordCount As Long
    startTime = Timer
    On Error GoTo ErrorHandler
    Set ws_Input = mainWb.Sheets("Input")
    Set ws_Output = mainWb.Sheets("Output")
    Set ws_Settings = mainWb.Sheets("Settings")
    Set ws_Log = mainWb.Sheets("Log")
    Call FormatColumnsAsText(ws_Input, ws_Settings)
    lastRow = ws_Output.Cells(ws_Output.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws_Output.Rows("2:" & lastRow).ClearContents
    Call EnsureOutputHeaders(ws_Output)
    lastRow = ws_Input.Cells(ws_Input.Rows.Count, 1).End(xlUp).Row
    outputRow = 2
    recordCount = 0
    For i = 2 To lastRow
        If ws_Input.Cells(i, 1).value <> "" Then
            Call ProcessRow(mainWb, ws_Input, ws_Output, ws_Settings, i, outputRow)
            outputRow = outputRow + 1
            recordCount = recordCount + 1
        End If
    Next i
    endTime = Timer
    Call LogExecution(ws_Log, startTime, endTime, recordCount)
    Exit Sub
ErrorHandler:
    MsgBox "Error in ProcessDutyCalculations: " & Err.Description, vbCritical
End Sub

Sub EnsureOutputHeaders(ws_Output As Worksheet)
    If ws_Output.Range("A1").value = "" Then
        ws_Output.Range("A1").value = "Entry Number"
        ws_Output.Range("B1").value = "Receipt Date"
        ws_Output.Range("C1").value = "Txn Code"
        ws_Output.Range("D1").value = "PO"
        ws_Output.Range("E1").value = "NGC#"
        ws_Output.Range("F1").value = "Material Number"
        ws_Output.Range("G1").value = "Qty"
        ws_Output.Range("H1").value = "Unit Value"
        ws_Output.Range("I1").value = "Total Value"
        ws_Output.Range("J1").value = "MFN Duty"
        ws_Output.Range("K1").value = "S301 Duty"
        ws_Output.Range("L1").value = "S338 Duty"
        ws_Output.Range("M1").value = "S122 Duty"
        ws_Output.Range("N1").value = "Cotton Fee"
        ws_Output.Range("O1").value = "MPF"
        ws_Output.Range("P1").value = "Total Fees"
        ws_Output.Range("Q1").value = "Total Duty"
        ws_Output.Range("R1").value = "99 Value"
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

Sub ProcessRow(mainWb As Workbook, ws_Input As Worksheet, ws_Output As Worksheet, ws_Settings As Worksheet, inputRow As Long, outputRow As Long)
    Dim htsNum As String, htsValue As Double, htsRate As Double, mfnValue As Double, mfnRate As Double
    Dim s301Rate As Double, s338Rate As Double, s122Rate As Double
    Dim qty As Double, mfnDuty As Double, s301Duty As Double, s338Duty As Double, s122Duty As Double, totalDuty As Double
    Dim col As Variant, htsColumns As Variant, receiptDate As Variant, countryOfOrigin As String
    Dim entryNumber As String, poNumber As String, ngcNumber As String, materialNumber As String, unitValue As Double
    Dim mfnHTSCode As String, ch99Value As Double
    Dim chapter99Count As Long
    Dim ch99HTSCodes() As String, ch99Rates() As Double
    Dim i As Long
    
    entryNumber = Trim(CStr(ws_Input.Cells(inputRow, 3).value))
    
    On Error Resume Next
    receiptDate = ws_Input.Cells(inputRow, 2).value
    If Not IsDate(receiptDate) Then receiptDate = Date
    On Error GoTo 0
    
    countryOfOrigin = Trim(CStr(ws_Input.Cells(inputRow, 13).value))
    poNumber = Trim(CStr(ws_Input.Cells(inputRow, 8).value))
    ngcNumber = Trim(CStr(ws_Input.Cells(inputRow, 6).value))
    materialNumber = Trim(CStr(ws_Input.Cells(inputRow, 7).value))
    qty = Val(ws_Input.Cells(inputRow, 11).value)
    htsColumns = Array(14, 17, 20, 23)
    
    mfnHTSCode = ""
    chapter99Count = 0
    ReDim ch99HTSCodes(0 To 3)
    ReDim ch99Rates(0 To 3)
    mfnValue = 0
    mfnRate = 0
    s301Rate = 0
    s338Rate = 0
    s122Rate = 0
    unitValue = 0
    ch99Value = 0
    
    For Each col In htsColumns
        htsNum = Trim(CStr(ws_Input.Cells(inputRow, col).value))
        If htsNum <> "" Then
            htsValue = Val(ws_Input.Cells(inputRow, col + 1).value)
            htsRate = Val(ws_Input.Cells(inputRow, col + 2).value)
            
            If Left(htsNum, 1) <> "9" And mfnHTSCode = "" Then
                mfnValue = htsValue
                mfnRate = htsRate
                unitValue = htsValue
                mfnHTSCode = htsNum
            End If
        End If
    Next col
    
    For Each col In htsColumns
        htsNum = Trim(CStr(ws_Input.Cells(inputRow, col).value))
        If htsNum <> "" Then
            htsValue = Val(ws_Input.Cells(inputRow, col + 1).value)
            htsRate = Val(ws_Input.Cells(inputRow, col + 2).value)
            
            If IsHTSMatch(htsNum, ws_Settings, 1) Then
                s301Rate = s301Rate + htsRate
                If htsValue > 0 Then ch99Value = ch99Value + htsValue
            End If
            
            If IsHTSMatch(htsNum, ws_Settings, 2) Then
                s338Rate = s338Rate + htsRate
                If htsValue > 0 Then ch99Value = ch99Value + htsValue
            End If
            
            If IsHTSMatch(htsNum, ws_Settings, 3) Then
                s122Rate = s122Rate + htsRate
                If htsValue > 0 Then ch99Value = ch99Value + htsValue
            End If
            
            If Left(htsNum, 2) = "99" Then
                ch99HTSCodes(chapter99Count) = htsNum
                ch99Rates(chapter99Count) = htsRate
                chapter99Count = chapter99Count + 1
            End If
        End If
    Next col
    
    mfnDuty = qty * (mfnValue * mfnRate)
    s301Duty = qty * (mfnValue * s301Rate)
    s338Duty = qty * (mfnValue * s338Rate)
    s122Duty = qty * (mfnValue * s122Rate)
    totalDuty = mfnDuty + s301Duty + s338Duty + s122Duty
    
    ws_Output.Cells(outputRow, 1).value = entryNumber
    ws_Output.Cells(outputRow, 2).value = receiptDate
    ws_Output.Cells(outputRow, 3).value = ws_Input.Cells(inputRow, 4).value
    ws_Output.Cells(outputRow, 4).value = poNumber
    ws_Output.Cells(outputRow, 5).value = ngcNumber
    ws_Output.Cells(outputRow, 6).value = materialNumber
    ws_Output.Cells(outputRow, 7).value = qty
    ws_Output.Cells(outputRow, 8).value = mfnValue
    ws_Output.Cells(outputRow, 9).value = qty * mfnValue
    ws_Output.Cells(outputRow, 10).value = mfnDuty
    ws_Output.Cells(outputRow, 11).value = s301Duty
    ws_Output.Cells(outputRow, 12).value = s338Duty
    ws_Output.Cells(outputRow, 13).value = s122Duty
    ws_Output.Cells(outputRow, 17).value = totalDuty
    ws_Output.Cells(outputRow, 18).value = ch99Value
    ws_Output.Cells(outputRow, 19).value = totalDuty
    
    Call ValidateUnitValue(mainWb, entryNumber, poNumber, ngcNumber, materialNumber, mfnHTSCode, countryOfOrigin, qty, unitValue)
    
    For i = 0 To chapter99Count - 1
        Call ValidateIEEPADutyRate(mainWb, entryNumber, poNumber, ngcNumber, materialNumber, ch99HTSCodes(i), countryOfOrigin, qty, CDate(receiptDate), ch99Rates(i))
    Next i
End Sub

Function IsHTSMatch(htsNum As String, ws_Settings As Worksheet, col As Long) As Boolean
    Dim lastRow As Long, i As Long, settingsValue As String
    On Error Resume Next
    lastRow = ws_Settings.Cells(ws_Settings.Rows.Count, col).End(xlUp).Row
    For i = 1 To lastRow
        settingsValue = Trim(CStr(ws_Settings.Cells(i, col).value))
        If settingsValue <> "" And settingsValue = htsNum Then
            IsHTSMatch = True
            Exit Function
        End If
    Next i
End Function

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



