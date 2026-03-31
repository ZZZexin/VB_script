' ================================
' Proc_TV - FMG_TV_LasPrep
' ================================
' Based on Russell Eade 4/2/2016
' Created by Zexin Yu 13/3/2026
' ================================
' This code is designed to take all .wcl files in the RootPath folder and perform the code
' the file is opened in wellcad and the "AMP" image track is copied and edited
' the copied files are split into depth chunks to prepare the files for isi processing
' the wcl files are then saved as with a _sliced.wcl extension 
' please have following structure to better process:
' RootPath\GPX
' RootPath\OTV
' RootPath\ATV
Dim RootPath
Dim wellReportPath, LasFilePath
Dim strBHName
Dim ok
Dim objFSO, obWCAD, obBHDOC, obHeader
Dim Dpos
Dim ReportInfo
Dim finalMap, key
Dim temQ ' temp for skipping importing las file

'Set obWCAD = CreateObject("WellCAD.Application")
'obWCAD.Showwindow()

'Definition of the root directory for templates (script folder)
'RootPath = folder containing this script
RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Read Well Report, field have been fixed. 
' If well report structure is changed
' The according function should be changed accordingly
'
Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()
Set obBHDoc = obWCAD.GetBorehole()

temQ = MsgBox("Is there ATV?", vbYesNoCancel + vbQuestion, "Confirm")
If temQ = vbYes Then
    obBHDOC.ApplyTemplate "C:\Proc_TV\05_FMG_TV_LasPrep\templates\FMGX_Template.wdt", false, true, false, false, True
ElseIf temQ = vbNo Then
    obBHDOC.ApplyTemplate "C:\Proc_TV\05_FMG_TV_LasPrep\templates\FMGX_OTV-ONLY.wdt", false, true, false, false, True
else
    WScript.Quit
End if 

ok = LoadWellReportLookupCSV(ReportInfo)
If Not ok Then
    ReadWellReport = False
    WScript.Quit
End If

ok = FindBHFilesInGPX(objFSO, RootPath, wellReportPath, lasFilePath)
If Not ok Then
    WScript.Quit
End If

Dim xlApp, wb, ws
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False
Set wb = xlApp.Workbooks.Open(wellReportPath)
Set ws = wb.Worksheets("Cover Sheet")
ws.Unprotect "magoo"

ok = ProcessWellReport(wb, "Cover Sheet", ReportInfo)
If Not ok Then
    ReadWellReport = False
    WScript.Quit
End If

Set finalMap = ReportInfo("ProcessedFieldMap")
ReadWellReport = True


Set obHeader = obBHDoc.Header

For Each key In finalMap.Keys
    obHeader.ItemText key, CStr(finalMap(key))
Next
    
wb.Close False
xlApp.Quit
Set wb = Nothing
Set xlApp = Nothing    

'==========================
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' Subfunction Library|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' find the well report under script's directory

Function FindBHFilesInGPX(fso, rootFolder, ByRef wellReportPath, ByRef lasFilePath)
    Dim gpxFolderPath

    wellReportPath = ""
    lasFilePath = ""

    gpxFolderPath = rootFolder & "\GPX"

    If Not fso.FolderExists(gpxFolderPath) Then
        MsgBox "GPX folder not found:" & vbCrLf & gpxFolderPath
        FindBHFilesInGPX = False
        Exit Function
    End If
    Call SearchGPXFolderRecursive(fso, fso.GetFolder(gpxFolderPath), wellReportPath, lasFilePath)

    If wellReportPath = "" Then
        MsgBox "Missing WellReport file in GPX folder."
        FindBHFilesInGPX = False
        Exit Function
    End If

    If lasFilePath = "" Then
        MsgBox "Missing OH LAS file, check OH data. Loading TV LAS"
    End If

    FindBHFilesInGPX = True
End Function

Sub SearchGPXFolderRecursive(fso, folderObj, ByRef wellReportPath, ByRef lasFilePath)
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

        ' optional: any LAS file
        If lasFilePath = "" Then
            If extNameUpper = "LAS" Then
                lasFilePath = fileObj.Path
            End If
        End If

        If wellReportPath <> "" And lasFilePath <> "" Then Exit Sub
    Next

    For Each subFolder In folderObj.SubFolders
        Call SearchGPXFolderRecursive(fso, subFolder, wellReportPath, lasFilePath)
        If wellReportPath <> "" And lasFilePath <> "" Then Exit Sub
    Next
End Sub

' Load config csv file 
Function LoadWellReportLookupCSV(byRef ReportInfo)
    Dim fso, ts
    Dim lookupPath, line, arr, key, val
    Dim section
    Dim fieldMap, tableArea, magDev

    lookupPath = "C:\Proc_TV\05_FMG_TV_LasPrep\templates\WELLREPORT_LOOKUP.csv"


    Set ReportInfo = CreateObject("Scripting.Dictionary")
    Set fieldMap = CreateObject("Scripting.Dictionary")
    Set tableArea = CreateObject("Scripting.Dictionary")
    Set magDev = CreateObject("Scripting.Dictionary")

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(lookupPath) Then
        MsgBox "Lookup file not found:" & vbCrLf & lookupPath
        LoadWellReportLookupCSV = False
        Exit Function
    End If

    Set ts = fso.OpenTextFile(lookupPath, 1)

    section = "FIELD"

    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        If line <> "" Then
            arr = Split(line, ",")

            key = ""
            val = ""

            If UBound(arr) >= 0 Then key = Trim(arr(0))
            If UBound(arr) >= 1 Then val = Trim(arr(1))

            If UCase(key) = "TABLEAREA" Then
                section = "TABLE"
            ElseIf UCase(key) = "MAGDEV" Then
                section = "MAG"
            ElseIf UCase(key) = "#END" Then
                Exit Do
            Else
                Select Case section
                    Case "FIELD"
                        If key <> "" Then fieldMap(key) = val
                    Case "TABLE"
                        If key <> "" Then tableArea(key) = val
                    Case "MAG"
                        If key <> "" Then magDev(key) = val
                End Select
            End If
        End If
    Loop

    ts.Close

    Set ReportInfo("FieldMap") = fieldMap
    Set ReportInfo("TableArea") = tableArea
    Set ReportInfo("MagDev") = magDev

    LoadWellReportLookupCSV = True
End Function





' find latest tv records from Well Report
Function FindLatestTVRecord(ws, tableArea, byRef tvInfo)
    Dim rngType, rngDate, rngOp, rngUnit
    Dim startRow, endRow, r
    Dim typeCol, dateCol, opCol, unitCol
    Dim bestRow, bestDate
    Dim curType, curDate

    Set tvInfo = CreateObject("Scripting.Dictionary")
    If Not tableArea.Exists("ToolType") Then Exit Function
    If Not tableArea.Exists("Date") Then Exit Function
    If Not tableArea.Exists("LeadOperator") Then Exit Function
    If Not tableArea.Exists("LoggingUnit") Then Exit Function
    
    Set rngType = ws.Range(tableArea("ToolType"))
    Set rngDate = ws.Range(tableArea("Date"))
    Set rngOp   = ws.Range(tableArea("LeadOperator"))
    Set rngUnit = ws.Range(tableArea("LoggingUnit"))

    startRow = rngType.Row
    endRow   = rngType.Row + rngType.Rows.Count - 1  
    
    typeCol = rngType.Column
    dateCol = rngDate.Column
    opCol   = rngOp.Column
    unitCol = rngUnit.Column

    For r = endRow To startRow Step -1
        curType = UCase(Trim(CStr(ws.Cells(r, typeCol).Value)))

        If curType = "OBI" Or curType = "ABI" Then
            tvInfo("ROW") = r
            tvInfo("DATE_RAW") = ws.Cells(r, dateCol).Value2
            tvInfo("DATE_TEXT") = ws.Cells(r, dateCol).Text
            tvInfo("TOOLTYPE") = Trim(CStr(ws.Cells(r, typeCol).Value))
            tvInfo("LEADOPERATOR") = Trim(CStr(ws.Cells(r, opCol).Value))
            tvInfo("LOGGINGUNIT") = Trim(CStr(ws.Cells(r, unitCol).Value))

            FindLatestTVRecord = True
            Exit Function
        End If
    Next

    FindLatestTVRecord = False
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

Function FindMagDevValue(locValue, magDevDict)
    Dim k, normLoc, normKey

    normLoc = NormalizeLookupKey(locValue)
    FindMagDevValue = ""

    For Each k In magDevDict.Keys
        normKey = NormalizeLookupKey(k)
        If normKey = normLoc Then
            FindMagDevValue = magDevDict(k)
            Exit Function
        End If
    Next
End Function


' Fetch data from Well Reports

Function ProcessWellReport(wb, coverSheetName, byRef ReportInfo)
    Dim wsCover
    Dim fieldMapRaw, fieldMapOut, tableArea, magDev
    Dim k, v, locValue, magSuggest, magFinal
    Dim tvInfo, ok

    Set wsCover = wb.Worksheets(coverSheetName)
    Set fieldMapRaw = ReportInfo("FieldMap")
    Set tableArea = ReportInfo("TableArea")
    Set magDev = ReportInfo("MagDev")

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

' 2) LOC -> MagDev lookup
    locValue = ""
    If fieldMapOut.Exists("LOC") Then locValue = fieldMapOut("LOC")

    magSuggest = FindMagDevValue(locValue, magDev)

    magFinal = InputBox( _
        "Confirm Mag Dev" & vbCrLf & _
        "LOC: " & locValue & vbCrLf & _
        "Suggested Mag Dev: " & magSuggest, _
        "Mag Dev", _
        magSuggest)

    If Trim(CStr(magFinal)) = "" Then
        magFinal = magSuggest
    End If

    fieldMapOut("MAGN") = magFinal

    ' 3) latest OBI/ABI from table area
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

    Set ReportInfo("ProcessedFieldMap") = fieldMapOut

    ProcessWellReport = True
End Function    