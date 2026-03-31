' Created by Zexin Yu 13/3/2026
' design for reading well report and fill relevant field in WSG template.
' That to be filled file should be open in WELLCAD windown
' this script should be with well report
' e.g. 
' root_path/universal_wr_reader
' root_path/some well report
' be advised that the field of well report is hard coded in this script, shoud be modified if any changes.

Dim RootPath
Dim wellReportPath
Dim FieldMap, TableArea
Dim strBHName




'Definition of the root directory for templates (script folder)
'RootPath = folder containing this script
RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set objFSO = CreateObject("Scripting.FileSystemObject")


'==========================
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' Subfunction Library|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'====================================================
' Hard-coded lookup maps
' be careful to modify
'====================================================
Sub HardCodedDicts(FieldMap, TableArea):
    set FieldMap = CreateObject("Scripting.Dictionary")
    FieldMap.CompareMode = 1   ' TextCompare
    
        ' ----- Single cell fields -----
    FieldMap.Add ":", "LATER"
    FieldMap.Add "#1", "LATER"
    FieldMap.Add "COMP", "D4"
    FieldMap.Add "WELL", "D11"
    FieldMap.Add "LOC",  "D14"
    FieldMap.Add "FLD",  "D13"
    FieldMap.Add "STAT", "D15"
    FieldMap.Add "CNTY", "D16"
    FieldMap.Add "LOGU", "LATER"
    FieldMap.Add "DATE", "LATER"
    FieldMap.Add "DRDP", "D21"
    FieldMap.Add "LOTD", "D22"
    FieldMap.Add "LT", "LATER" 'LT OPTV
    FieldMap.Add "LB", "LATER" 'LBOPTV
    FieldMap.Add "LB#1", "LATER" 'LTABI
    FieldMap.Add "LB#2", "LATER" 'LBABI
    FieldMap.Add "RECB", "LATER"
    FieldMap.Add "PDIP", "J12"
    FieldMap.Add "PAZI", "J13"
    FieldMap.Add "EAST", "L16"
    FieldMap.Add "NRTH", "L17"
    FieldMap.Add "EGL",  "-"
    FieldMap.Add "MAGN", "LATER"
    FieldMap.Add "RIGN", "-"
    FieldMap.Add "BS",   "D20"
    FieldMap.Add "BSU",  "E20"
    FieldMap.Add "CASB", "-"
    FieldMap.Add "CASL", "-"
    FieldMap.Add "CASX", "-"
    FieldMap.Add "CASD", "-"

    ' ----- Table areas -----
    Set TableArea = CreateObject("Scripting.Dictionary")
    TableArea.CompareMode = 1

    TableArea.Add "ToolType",    "B34:B44"
    TableArea.Add "LoggingUnit", "I34:I44"
    TableArea.Add "LeadOperator","H34:H44"
    TableArea.Add "Date",        "G34:G44"
    TableArea.Add "LT",          "L34:L44"
    TableArea.Add "LB",          "N34:N44"

End Sub


Function FindBHFilesInGPX(fso, rootFolder, ByRef wellReportPath, ByRef )
    Dim gpxFolderPath

    wellReportPath = ""

    gpxFolderPath = rootFolder

    If Not fso.FolderExists(gpxFolderPath) Then
        MsgBox "GPX folder not found:" & vbCrLf & gpxFolderPath
        FindBHFilesInGPX = False
        Exit Function
    End If
    Call SearchGPXFolderRecursive(fso, fso.GetFolder(gpxFolderPath), wellReportPath)

    If wellReportPath = "" Then
        MsgBox "Missing WellReport file in current folder."
        FindBHFilesInGPX = False
        Exit Function
    End If


    FindBHFilesInGPX = True
End Function

Sub SearchGPXFolderRecursive(fso, folderObj, ByRef wellReportPath, ByRef)
    Dim fileObj, subFolder
    Dim fileNameUpper, extNameUpper

    For Each fileObj In folderObj.Files
        fileNameUpper = UCase(fileObj.Name)
        extNameUpper = UCase(fso.GetExtensionName(fileObj.Name))

        ' required: WellReport Excel file
        If wellReportPath = "" Then
            If InStr(1, fileNameUpper, "WELLREPORT", vbTextCompare) > 0 Then
                If extNameUpper = "XLSX" Or extNameUpper = "XLSM" Or extNameUpper = "XLS" Then
                    wellReportPath = fileObj.Path
                End If
            End If
        End If

        If wellReportPath <> ""  Then Exit Sub
    Next

    For Each subFolder In folderObj.SubFolders
        Call SearchGPXFolderRecursive(fso, subFolder, wellReportPath)
        If wellReportPath <> ""  Then Exit Sub
    Next
End Sub

Function FindLatestTVRecord(ws, tableArea, byRef, tvInfo)
    Dim rngType, rngDate, rngOp, rngUnit, rngTop, rngBot
    Dim startRow, endRow, r
    Dim typeCol, dateCol, opCol, unitCol, LTPCol, LBTCol
    Dim bestRow, bestDate
    Dim curType, curDate
    Dim foundLatest, foundOBI, foundABI

    Set tvInfo = CreateObject("Scripting.Dictionary")
    
    If Not tableArea.Exists("ToolType") Then Exit Function
    If Not tableArea.Exists("Date") Then Exit Function
    If Not tableArea.Exists("LeadOperator") Then Exit Function
    If Not tableArea.Exists("LoggingUnit") Then Exit Function    

    Set rngType = ws.Range(tableArea("ToolType"))
    Set rngDate = ws.Range(tableArea("Date"))
    Set rngOp   = ws.Range(tableArea("LeadOperator"))
    Set rngUnit = ws.Range(tableArea("LoggingUnit"))
    Set rngTop = ws.Range(tableArea("LT"))
    Set rngTBot = ws.Range(tableArea("BT"))

    startRow = rngType.Row
    endRow = rngType.Row + rngType.Rows.Count - 1

    typeCol = rngType.Column
    dateCol = rngDate.Column
    opCol   = rngOp.Column
    unitCol = rngUnit.Column
    LTPCol = rngTop.Column
    LBTCol = rngBot.Column
    
    foundLatest = False
    foundOBI = False
    foundABI = False
    
    For r = endRow To startRow Step -1
        curType = UCase(Trim(CStr(ws.Cells(r, typeCol).Value)))

        If curType = "OBI" Or curType = "ABI" Then

            If Not foundLatest Then
                tvInfo("ROW") = r
                tvInfo("TYPE") = curType
                tvInfo("DATE_RAW") = ws.Cells(r, dateCol).Value2
                tvInfo("DATE_TEXT") = ws.Cells(r, dateCol).Text
                tvInfo("TOOLTYPE") = Trim(CStr(ws.Cells(r, typeCol).Value))
                tvInfo("LEADOPERATOR") = Trim(CStr(ws.Cells(r, opCol).Value))
                tvInfo("LOGGINGUNIT") = Trim(CStr(ws.Cells(r, unitCol).Value))
                foundLatest = True
            End If

            If curType = "OBI" And Not foundOBI Then
                    tvInfo("OBI_ROW") = r
                    tvInfo("OBI_LT") = ws.Cells(r, LTPCol).Value
                    tvInfo("OBI_LB") = ws.Cells(r, LBTCol).Value
                    foundOBI = True
            End If
            
            If curType = "ABI" And Not foundABI Then
                    tvInfo("ABI_ROW") = r
                    tvInfo("ABI_LT") = ws.Cells(r, ltCol).Value
                    tvInfo("ABI_LB") = ws.Cells(r, lbCol).Value
                    foundABI = True
            End If
            If foundLatest And foundOBI And foundABI Then Exit For
        End If
    Next
    FindLatestTVRecord = foundLatest
End Function

Function NormalizeLookupKey(s)
    Dim t
    t = UCase(Trim(CStr(s)))
    t = Replace(t, " ", "")
    NormalizeLookupKey = t
End Function

Function CloneDict(src)
    Dim d, k
    Set d = CreateObject("Scripting.Dictionary")

    For Each k In src.Keys
        d(k) = src(k)
    Next

    Set CloneDict = d
End Function


Function ProcessWellReport(wb, coverSheetName, byRef)
    Dim wsCover
    Dim fieldMapRaw, fieldMapOut, tableArea
    Dim k, v
    Dim tvInfo, ok

    HardCodedDicts fieldMapRaw, TableArea
    Set wsCover = wb.Worksheets(coverSheetName)
    
    Set fieldMapOut = CloneDict(fieldMapRaw)
        ' 1) replace cell address with actual value from Cover Sheet
    For Each k In fieldMapRaw.Keys
        v = Trim(CStr(fieldMapRaw(k)))

        If UCase(v) = "LATER" Then
            fieldMapOut(k) = "LATER"
        ElseIf v = "-" Then
            fieldMapOut(k) = "-"
        ElseIf v <> "" Then
            fieldMapOut(k) = Trim(CStr(wsCover.Range(v).Text))
        Else
            fieldMapOut(k) = ""
        End If
    Next

    ok = FindLatestTVRecord(wsCover, tableArea, tvInfo)
    If Not ok Then
        MsgBox "Well report wrong: no ABI or OBI record found."
        ProcessWellReport = False
        Exit Function
    End If

    If fieldMapOut("BSU")= "in" Then 
        fieldMapOut("BS") = CStr(CLng(CDbl(fieldMapOut("BS")) * 25.4))
    ElseIf fieldMapOut("BSU")= "cm" Then
        fieldMapOut("BS") = CStr(CLng(CDbl(fieldMapOut("BS")) * 10))
    End if

    fieldMapOut("RECB") = tvInfo("LEADOPERATOR")
    fieldMapOut("LOGU") = tvInfo("LOGGINGUNIT")
    fieldMapOut("DATE") = tvInfo("DATE_TEXT")
    fieldMapOut("LT") = tvInfo("OBI_LT")
    fieldMapOut("LB") = tvInfo("OBI_LB")
    fieldMapOut("LB#1") = tvInfo("ABI_LT")
    fieldMapOut("LB#2") = tvInfo("ABI_LB")

    ProcessWellReport = True
End Function    
