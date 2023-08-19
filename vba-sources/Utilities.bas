Attribute VB_Name = "Utilities"
Option Explicit

'Formats
Public Const fmtDate As String = "[$-409]dd-mmm-yyyy;@"
Public Const fmtAccounting As String = "#,##0.00;[Red]-#,##0.00"
Public Const fmtInteger As String = "#,##0_);(#,##0)"

Public tStart As Date

Public bDisplayAlerts As Boolean
Public vCalculation As Variant

'---------------------------------------------------------------------------------------------------------------------------
' Objective:    Prepare appliction for automation execution.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    02-Nov-2016     Eduardo Garcia      Created.
'                    2.0    01-Jun-2016     Eduardo Garcia      Parameter removed.
'                    3.0    13-Mar-2018     Eduardo Garcia      Application variables added.
'---------------------------------------------------------------------------------------------------------------------------
Public Sub StartCode()
    tStart = Time
    bDisplayAlerts = Application.DisplayAlerts
    vCalculation = Application.Calculation
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

'---------------------------------------------------------------------------------------------------------------------------
' Objective:    Converts Brazil (European) dates to English.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    16-Feb-2019     Eduardo Garcia      Created.
'---------------------------------------------------------------------------------------------------------------------------
Public Sub DateBRtoUS(HeaderRow As Long, ColumnByName As String, LastRow As Long)
    Dim yDay As Byte
    Dim yMonth As Byte
    Dim lYear As Long
    
    Dim tHours As Date
    
    Dim rngRange As Range
    
    
    Range(HeaderRow & ":" & HeaderRow).Find(ColumnByName, Cells(HeaderRow, 1), _
        xlValues, xlWhole).Offset(1, 0).Activate
    
    'Range(ActiveCell, Cells(LastRow, ActiveCell.Column)).TextToColumns _
        Destination:=ActiveCell, DataType:=xlDelimited, TextQualifier:=xlNone, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, _
        Space:=False, Other:=False, FieldInfo:=Array(1, xlDMYFormat), _
        TrailingMinusNumbers:=True
    
    For Each rngRange In Range(ActiveCell, Cells(LastRow, ActiveCell.Column))
        
        If rngRange.value = "" Then
            GoTo NextRange
        End If
        
        yDay = Left(rngRange.value, 2)
        yMonth = Mid(rngRange.value, 4, 2)
        lYear = Mid(rngRange.value, 7, 4)
        tHours = 0
        
        If InStr(1, rngRange.value, " ", vbTextCompare) > 0 Then
            tHours = Right(rngRange.value, 8)
        End If
        
        rngRange.value = CDate(yMonth & "/" & yDay & "/" & lYear & " " & tHours)
NextRange:
    Next rngRange
    
    Set rngRange = Nothing
End Sub

'---------------------------------------------------------------------------------------------------------------------------
' Objective:    Finished appliction for automation execution.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    02-Nov-2016     Eduardo Garcia      Created.
'                    2.0    01-Jun-2016     Eduardo Garcia      Parameter removed.
'                    3.0    13-Mar-2018     Eduardo Garcia      Application variables added.
'---------------------------------------------------------------------------------------------------------------------------
Public Sub EndCode()
    Application.DisplayAlerts = bDisplayAlerts
    Application.Calculation = vCalculation
    MsgBox "Finished in " & Format(Time - tStart, "Hh:Nn:Ss"), vbInformation
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Adjust columns width.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    27-Oct-2016     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub AdjustColumnsWidth(ByRef refSheet As Worksheet)
    Dim rngColumn As Range
    
    
    refSheet.Cells.EntireColumn.AutoFit
    
    Set rngColumn = refSheet.Range("A1")
    
    Do Until rngColumn.value = ""
        
        If rngColumn.ColumnWidth > 50 Then
            rngColumn.ColumnWidth = 50
        End If
        
        Set rngColumn = rngColumn.Offset(0, 1)
    Loop
    
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Remove auto filters from active workbook.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    02-May-2018     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub ClearActiveWorkbookFilters()
    Dim wksSheet As Worksheet
    
    For Each wksSheet In ActiveWorkbook.Sheets
        Call ClearFilters(wksSheet)
    Next wksSheet
    
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Remove sheet autofilters.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    26-Mar-2018     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub ClearFilters(ByRef refSheet As Worksheet)
            
    If refSheet.FilterMode Then
        refSheet.ShowAllData
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Collapse all grouped levels rows and columns from all sheets.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    26-Apr-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub ColapseAllGroups()
    Dim wksSheet As Worksheet
    
    
    On Error Resume Next
    
    For Each wksSheet In ThisWorkbook.Sheets
        wksSheet.Activate
        wksSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=0
    Next wksSheet
    
    Set wksSheet = Nothing
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Expand all grouped levels (passed as params) of rows and columns from all sheets.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    26-Apr-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub ExpandAllGroups(Optional ByVal lRowLevels As Long = 1, _
Optional ByVal lColumnLevels As Long = 1)
    Dim wksSheet As Worksheet
    
    On Error Resume Next
    
    For Each wksSheet In ThisWorkbook.Sheets
        wksSheet.Activate
        wksSheet.Outline.ShowLevels RowLevels:=lRowLevels, ColumnLevels:=lColumnLevels
    Next wksSheet
    
    Set wksSheet = Nothing
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Extract only number from a value.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    12-Jul-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function ExtractNumbers(Cell As String) As String
    Dim lCharacter As Long
    Dim lSize As Long
    
    Dim sCharacter As String
    Dim sTemp As String
    
    
    sTemp = ""
    lSize = Len(Cell)
    
    For lCharacter = 1 To lSize
        sCharacter = Mid(Cell, lCharacter, 1)
        
        If IsNumeric(sCharacter) Then
            sTemp = sTemp & sCharacter
        End If
        
    Next lCharacter
    
    ExtractNumbers = sTemp
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Find last row that contains value for activesheet.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    17-Oct-2016     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function FindLastRow(lFirstCol As Long, lLastCol As Long) As Long
    Dim lCol As Long
    Dim lRow As Long
    Dim lTemp As Long
    
    lRow = 0
    
    For lCol = lFirstCol To lLastCol
        lTemp = ActiveSheet.Cells(1048576, lCol).End(xlUp).Row
        lRow = IIf(lTemp > lRow, lTemp, lRow)
    Next lCol
    
    FindLastRow = lRow
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Find last row that contains value for activesheet.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    17-Oct-2016     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function FindNA(ByRef refSheet As Worksheet) As Boolean
    Dim objFound As Object
    
    Set objFound = refSheet.Cells.Find("#N/A", refSheet.Range("A1"), xlValues, XlLookAt.xlWhole, _
        XlSearchOrder.xlByRows)
    
    FindNA = Not (objFound Is Nothing)
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Fix cells contents by just converting them to General format.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    02-Nov-2016     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub FixColumnTextToColumns(ByRef refSheet As Worksheet, ColumnSelected As Long, _
RowStart As Long, RowEnd As Long)
    
    If refSheet.Cells(1048576, ColumnSelected).End(xlUp).Row > 1 Then
        refSheet.Range(Cells(RowStart, ColumnSelected).Address & ":" & Cells(RowEnd, _
            ColumnSelected).Address).TextToColumns Destination:=refSheet.Cells(RowStart, _
            ColumnSelected), DataType:=xlFixedWidth, FieldInfo:=Array(0, 1), _
            TrailingMinusNumbers:=True
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Get all files from a specified folder.
'
' Requires:
'   References: Microsoft Scripting Runtime
'
' History:      Revision    Date            Developer           Notes
'                    1.0    15-Jan-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function GetAllFiles(sPath As String) As Object
'    Dim objFSO As New FileSystemObject
    Dim objFSO As Object
    
    
    Set objFSO = CreateObject("FileSystemObject")
    
    If Not objFSO.FolderExists(sPath) Then
        MsgBox "Folder [" & sPath & "] not found ", vbCritical, "Error"
        Set GetAllFiles = Nothing
    Else
        Set GetAllFiles = objFSO.GetFolder(sPath).Files
    End If
    
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Get file name.
'
' Requires:
'   References: Microsoft Scripting Runtime
'
' History:      Revision    Date            Developer           Notes
'                    1.0    15-Jan-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function GetFileName(objFile As Object) As String
'    Dim objFSO As New FileSystemObject
    Dim objFSO As Object
    
    
    Set objFSO = CreateObject("FileSystemObject")
    
    GetFileName = objFSO.GetFileName(objFile)
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Open a file selection window.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    17-Oct-2016     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function GetFiles(Optional sPath As String, Optional sFilter As String = "All Files,*.*", _
Optional sTitle = "Select Files", Optional Multiple As Boolean = True) As Collection
    Dim sDescription As String
    Dim sExtension As String
    
    Dim vItem As Variant
    Dim vFiles As Variant
    
    Dim cItems As New Collection
    
    Dim fdlg As Office.FileDialog
    
    
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    
    With fdlg
        .AllowMultiSelect = Multiple
        .Filters.Clear
        sDescription = Split(sFilter, ",")(0)
        sExtension = Split(sFilter, ",")(1)
        .Filters.Add sDescription, sExtension
        .InitialFileName = sPath
        .Title = sTitle
        .Show
        
        If .SelectedItems.Count = 0 Then
            Set cItems = Nothing
        End If
        
        For Each vItem In .SelectedItems
            cItems.Add vItem
        Next vItem
        
    End With

    Set fdlg = Nothing
    Set GetFiles = cItems
    Set cItems = Nothing
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Get only one FileName using spefied FileName name.
'
' Requires:
'   References: Microsoft Scripting Runtime
'
' History:      Revision    Date            Developer           Notes
'                    1.0    02-Nov-2016     Eduardo Garcia      Created.
'                    2.0    15-Jan-2017     Eduardo Garcia      File System Object implemented.
'                                                               Returns object now.
'--------------------------------------------------------------------------------------------------
Public Function GetOneFile(sFilePath As String) As Object
'    Dim objFSO As New FileSystemObject
    Dim objFSO As Object
    
    
    Set objFSO = CreateObject("FileSystemObject")
    
    
    If Not objFSO.FileExists(sFilePath) Then
        MsgBox "File [" & sFilePath & "] does not exist", vbCritical, "Error"
        Set GetOneFile = Nothing
    Else
        Set GetOneFile = objFSO.GetFile(sFilePath)
    End If
    
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Get path slash direction.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    22-Mar-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function GetSlash(Path As String)
    GetSlash = IIf(InStr(1, Path, "/") > 0, "/", "\")
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Check if a values is in a pivot field.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    16-Mar-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function HasPivotItem(pField As PivotField, value As String) As Boolean
    On Error Resume Next
    HasPivotItem = False
    HasPivotItem = Not IsNull(pField.PivotItems(value))
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Check if a string is in an array.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    01-Nov-2016     Eduardo Garcia      Created.
'                    2.0    25-Apr-2017     Eduardo Garcia      For Each loop implemented.
'--------------------------------------------------------------------------------------------------
Public Function IsInArray(SearchString As String, ArrayToSearch As Variant) As Boolean
    Dim bTemp As Boolean
    
    Dim vItem As Variant
    
    
    For Each vItem In ArrayToSearch
        bTemp = (vItem = SearchString)
        
        If bTemp Then
            Exit For
        End If
        
    Next vItem
    
    IsInArray = bTemp
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Check if the Cost Certer has a valid format.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    17-Oct-2016     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Function IsValidCostCenter(vCC As Variant)
    IsValidCostCenter = (IsNumeric(vCC) Or (Left(vCC, 3) = "PS-") Or (Right(vCC, 4) = "HQ01")) _
        And (Len(vCC) <= 10)
End Function

'--------------------------------------------------------------------------------------------------
' Objective:    Remove Named Ranges with issues.
'
' Requires:
'   References: Microsoft Scripting Runtime
'
' History:      Revision    Date            Developer           Notes
'                    1.0    17-Oct-2016     Eduardo Garcia      Created.
'                    2.0    27-Apr-2017     Eduardo Garcia      DisplayMSG parameter added.
'                    3.0    14-Jul-2017     Eduardo Garcia      ExportToFile parameter added.
'                                                               Exporting message code implemented
'                                                               stmText -> WriteToCSV.
'--------------------------------------------------------------------------------------------------
Public Sub NamesKiller(Optional ByVal KeepArray As Variant, _
Optional ByVal DisplayMSG As Boolean = True, Optional ByVal ExportToFile As Boolean = False)
    Dim bDelete As Boolean
    
    Dim lCount As Long
    Dim lNames As Long
    Dim lStyles As Long
    
    Dim sMessage As String
    
'    Dim objFSO As New FileSystemObject
    Dim objFSO As Object
    
'    Dim txtStm As TextStream
    Dim txtStm As Object
    
    Dim vItem As Variant
    
    
    On Error Resume Next
    
    Set objFSO = CreateObject("FileSystemObject")
    Set txtStm = CreateObject("TextStream")
    
    lNames = Names.Count
    lCount = lNames - 1
    
    For lCount = lNames To 1 Step -1
        sMessage = "Named Range " & lCount & " of " & lNames & " named " & Names(lCount).Name & _
            ",Kept"
        
        If IsMissing(KeepArray) Then
            
            If (InStr(1, Names(lCount).Name, "xlfn", vbTextCompare) = 0) And (InStr(1, Names(lCount). _
                Name, "Slicer", vbTextCompare) = 0) Then
                
                If (InStr(1, Names(lCount).RefersTo, "#REF!", vbTextCompare) > 0) _
                Or (InStr(1, Names(lCount).RefersTo, "$", vbTextCompare) = 0) Then
                    Names(lCount).Delete
                    sMessage = Replace(sMessage, "Kept", "Deleted")
                End If
                
            End If
                
        Else
            bDelete = True
            
            For Each vItem In KeepArray
                
                If Names(lCount).Name = vItem Then
                    bDelete = False
                    Exit For
                End If
                
            Next
            
            If bDelete Then
                Names(lCount).Delete
                sMessage = Replace(sMessage, "Kept", "Deleted")
            End If
            
        End If
        
        Debug.Print sMessage
        
        ' This piece of code does not work with file opened from OneDrive/SharePoint
        If ExportToFile Then
            
            If txtStm Is Nothing Then
                Set txtStm = objFSO.CreateTextFile(Replace(ThisWorkbook.FullName, ".xlsm", "_Names.csv"), True, False)
            End If
            
            txtStm.WriteLine sMessage
        End If
        
    Next
    
    If Not txtStm Is Nothing Then
        txtStm.Close
    End If
    
    sMessage = ""
    sMessage = sMessage & "Named Ranges Removed:" & vbTab & (lNames - Names.Count) & vbCr
    sMessage = sMessage & "Named Ranges Remaining:" & vbTab & Names.Count & vbCr
    
    If DisplayMSG Then
        MsgBox sMessage, vbInformation
    End If
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------
' Objective:    Converts Brazil (European) numbers to English or just fix them.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    16-Feb-2019     Eduardo Garcia      Created.
'---------------------------------------------------------------------------------------------------------------------------
Public Sub NumberBRtoUS(HeaderRow As Long, ColumnByName As String, LastRow As Long)
    Dim sDecimal As String
    Dim sThousand As String
    
    
    Range(HeaderRow & ":" & HeaderRow).Find(ColumnByName, Cells(HeaderRow, 1), _
        xlValues, xlWhole).Offset(1, 0).Activate
    
    ' Check language (1033 = US, 1046 = BR)
    Select Case Application.SpellingOptions.DictLang
        Case 1033
            sDecimal = "."
            sThousand = ","
        Case 1046
            sDecimal = ","
            sThousand = "."
        Case Else
            MsgBox "Missing defition of separators for Decimal and Thousand " & vbCrLf & _
                "Please contact Eduardo Pereira Garcia at edupgarcia@protonmail.com " & _
                "and ask for this fix.", vbCritical
            Exit Sub
    End Select
    
    Range(ActiveCell, Cells(LastRow, ActiveCell.Column)).TextToColumns _
        Destination:=ActiveCell, DataType:=xlDelimited, TextQualifier:=xlNone, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, _
        Space:=False, Other:=False, FieldInfo:=Array(1, xlGeneralFormat), _
        DecimalSeparator:=sDecimal, ThousandsSeparator:=sThousand, _
        TrailingMinusNumbers:=True
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Rename pivot field name.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    11-May-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Private Sub PivotFixFieldName(OldName As String, NewName As String)
    Dim wksSheet As Worksheet
    Dim pvtPivot As PivotTable
    Dim pflPivotField As PivotField
    
    
    For Each wksSheet In ThisWorkbook.Sheets
    
        For Each pvtPivot In wksSheet.PivotTables
            
            For Each pflPivotField In pvtPivot.PivotFields
                
                If pflPivotField.Name = OldName Then
                    pflPivotField.Name = NewName
                End If
                
            Next pflPivotField
            
        Next pvtPivot
        
    Next wksSheet
    
    MsgBox "Done", vbInformation
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Remove/re-add all filter items every time pivot is refreshed.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    11-May-2017     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub PivotItemsLimitNone()
    Dim shtSheet As Worksheet
    
    Dim pvtPivot As PivotTable
    
    
    For Each shtSheet In ThisWorkbook.Sheets
    
        For Each pvtPivot In shtSheet.PivotTables
            pvtPivot.PivotCache.MissingItemsLimit = xlMissingItemsNone
            Debug.Print shtSheet.Name & "," & pvtPivot.Name
        Next pvtPivot
        
    Next shtSheet
    
    MsgBox "Done", vbInformation
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Update all Pivot tables from active worksheet.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    09-Oct-2017     Eduardo Garcia      Created.
'                    2.0    19-Dec-2017     Eduardo Garcia      Added paramenter to show message or
'                                                               not.
'--------------------------------------------------------------------------------------------------
Public Sub RefreshPivots(Optional ByVal ShowMessage As Boolean = True)
    Dim pvtPivot As PivotTable
    
    
    If ShowMessage Then
        tStart = Now
    End If
    
    For Each pvtPivot In ActiveSheet.PivotTables
        pvtPivot.PivotCache.Refresh
    Next pvtPivot
    
    If ShowMessage Then
        MsgBox "Finished in " & Format(Now - tStart, "HH:NN:SS"), vbInformation
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Reset the calculation mode to automatic and re-activate the display of alerts.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    19-Mar-2018     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Sub RestoreCalculationAndDisplayAlerts()
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Export modules and classes.
'
' History:      Revision    Date            Developer           Notes
'                    1.0    16-Feb-2019     Eduardo Garcia      Created.
'--------------------------------------------------------------------------------------------------
Public Sub SourceExport()
    Dim yModule As Byte
    
    Dim sSourcePath As String
    
    Dim objFSO As Object
    
    
    sSourcePath = ThisWorkbook.Path & "\Source\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    With ThisWorkbook.VBProject.VBComponents
        
        For yModule = 1 To .Count
            
            Select Case Left(.Item(yModule).Name, 3)
                Case "sht", "cls"
                    .Item(yModule).Export sSourcePath & .Item(yModule).Name & ".cls"
                Case "mdl"
                    .Item(yModule).Export sSourcePath & .Item(yModule).Name & ".bas"
                Case "Thi"
                    .Item(yModule).Export sSourcePath & "ThisWorkbook.cls"
            End Select
            
        Next yModule
    
    End With
    
    MsgBox "Modules exported", vbInformation
    
    Set objFSO = Nothing
End Sub

'--------------------------------------------------------------------------------------------------
' Objective:    Remove non-used Styles.
'
' Requires:
'   References: Microsoft Scripting Runtime
'
' History:      Revision    Date            Developer           Notes
'                    1.0    17-Oct-2016     Eduardo Garcia      Created.
'                    2.0    27-Apr-2017     Eduardo Garcia      DisplayMSG parameter added.
'                    3.0    14-Jul-2017     Eduardo Garcia      ExportToFile parameter added.
'                                                               Exporting message code implemented
'                                                               stmText -> WriteToCSV.
'--------------------------------------------------------------------------------------------------
Public Sub StyleKiller(Optional ByVal DisplayMSG As Boolean = True, _
Optional ByVal ExportToFile As Boolean = False)
    Dim lCount As Long
    Dim lNames As Long
    Dim lStyles As Long
    
    Dim sMessage As String
    
'    Dim objFSO As New FileSystemObject
    Dim objFSO As Object
    
'    Dim txtStm As TextStream
    Dim txtStm As Object
    
    Dim vItem As Variant
    
    
    On Error Resume Next
    
    Set objFSO = CreateObject("FileSystemObject")
    Set txtStm = CreateObject("TextStream")
    
    lStyles = ThisWorkbook.Styles.Count
    
    For lCount = ThisWorkbook.Styles.Count - 1 To 1 Step -1
        sMessage = "Style " & lCount & " of " & lStyles & " named " & ThisWorkbook. _
            Styles(lCount).Name & ",Kept"
        
        If Not ThisWorkbook.Styles(lCount).BuiltIn Then
            ThisWorkbook.Styles(lCount).Locked = False
            ThisWorkbook.Styles(lCount).Delete
            sMessage = Replace(sMessage, "Kept", "Deleted")
        End If
        
        Debug.Print sMessage
        
        ' This piece of code does not work with file opened from OneDrive/SharePoint
        If ExportToFile Then
            
            If txtStm Is Nothing Then
                Set txtStm = objFSO.CreateTextFile(Replace(ThisWorkbook.FullName, ".xlsm", "_Styles.csv"), True, False)
            End If
            
            txtStm.WriteLine sMessage
        End If
        
    Next
    
    If Not txtStm Is Nothing Then
        txtStm.Close
    End If
    
    sMessage = ""
    sMessage = sMessage & "Styles Removed:" & vbTab & (lStyles - ThisWorkbook.Styles.Count) & vbCr
    sMessage = sMessage & "Styles Remaining:" & vbTab & ThisWorkbook.Styles.Count & vbCr
    
    If DisplayMSG Then
        MsgBox sMessage, vbInformation
    End If
    
End Sub
