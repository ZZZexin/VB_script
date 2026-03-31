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

'Get the hard coded dictinary
HardCodedDicts FieldMap, TableArea


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
    FieldMap.Add "LTOPTV", "LATER"
    FieldMap.Add "LBOPTV", "LATER"
    FieldMap.Add "LTABI", "LATER"
    FieldMap.Add "LBABI", "LATER"
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


Function FindBHFilesInGPX(fso, rootFolder, ByRef wellReportPath, ByRef lasFilePath)
    Dim gpxFolderPath

    wellReportPath = ""
    lasFilePath = ""

    gpxFolderPath = rootFolder

    If Not fso.FolderExists(gpxFolderPath) Then
        MsgBox "GPX folder not found:" & vbCrLf & gpxFolderPath
        FindBHFilesInGPX = False
        Exit Function
    End If
    Call SearchGPXFolderRecursive(fso, fso.GetFolder(gpxFolderPath), wellReportPath, lasFilePath)

    If wellReportPath = "" Then
        MsgBox "Missing WellReport file in current folder."
        FindBHFilesInGPX = False
        Exit Function
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