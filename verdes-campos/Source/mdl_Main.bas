Attribute VB_Name = "mdl_Main"
Option Explicit

Const sMainPassword As Long = 1108

Private Sub SetFileAndDate(Row As Byte, FileName As String)
    shtMenu.Activate
    ActiveSheet.Unprotect sMainPassword
    Range("C" & Row).value = FileName
    Range("D" & Row).value = Now
    ActiveSheet.Protect sMainPassword
End Sub

Public Sub UpdateKMMDay()
    Dim lCol As Long
    Dim lLastRow As Long
    
    Dim cltFile As New Collection
    
    Dim wbkInput As Workbook
    Dim wbkTemp As Workbook
    Dim rngLast As Range
    
    Dim vCol As Variant
    Dim vCols As Variant
    
    
    Set cltFile = Utilities.GetFiles(, , "Selecione o arquivo [KMM Dia]", False)
    
    If cltFile.Count = 0 Then
        MsgBox "Nenhum arquivo selecionado", vbCritical
        Exit Sub
    End If
    
    shtMenu.Activate
    SetFileAndDate Cells.Find("Arquivos").Offset(2, 0).Row, cltFile(1)
    
    Set wbkInput = Workbooks.Open(cltFile(1), False, True)
    
    ' Clear old data
    shtKMMDia.Activate
    lLastRow = Cells.Find("Frota").End(xlDown).Row
    Range("A3:A" & lLastRow).EntireRow.ClearContents
    
    ' Copy new data
    wbkInput.Sheets(1).Activate
    lLastRow = Cells.Find("Frota").End(xlDown).Row - 1
    Range("A2:AM" & lLastRow).Copy
    
    ' Paste new data
    shtKMMDia.Activate
    Range("C2").PasteSpecial xlPasteValues
    lLastRow = Cells.Find("Frota").End(xlDown).Row
    
    ' Fill down formulas and formats
    ' Formula range 1
    Range("A2:B2").Copy
    Range("A3:B" & lLastRow).PasteSpecial xlPasteFormulas
    Range("A3:B" & lLastRow).PasteSpecial xlPasteFormats
    
    ' Data range
    Range("C2:AO2").Copy
    Range("C3:AO" & lLastRow).PasteSpecial xlPasteFormats
    
    ' Formula range 2
    Set rngLast = Range("XFD1").End(xlToLeft)
    Range("AP2", Cells(2, rngLast.Column)).Copy
    Range("AP3", Cells(lLastRow, rngLast.Column)).PasteSpecial xlPasteFormulas
    Range("AP3", Cells(lLastRow, rngLast.Column)).PasteSpecial xlPasteFormats
    Set rngLast = Nothing
    
    ' Numeric Columns
    On Error Resume Next
    
    vCols = Array("Frota", "Ton", "Romaneio", "Motorista", "Faturamento", "Km Rodado", _
        "Km Atual", "Próx. Revisão", "Número O.S.", "Revisão")
    
    For Each vCol In vCols
        NumberBRtoUS HeaderRow:=1, ColumnByName:=CStr(vCol), LastRow:=lLastRow
    Next vCol
    
    ' Fix Date Columns
    vCols = Array("Data Status", "Data de Entrada", "Últ. Insp. Pneus Carreta", _
        "Últ. Insp. Pneus Cavalo")
    
    For Each vCol In vCols
        DateBRtoUS HeaderRow:=1, ColumnByName:=CStr(vCol), LastRow:=lLastRow
    Next vCol
    
    On Error GoTo 0
    
    ' Close Input
    Range("A2").Select
    wbkInput.Close False
    
    ' Update Customers
    shtCheckEmails.Range("A3:B1048576").Clear
    
    shtKMMDia.Activate
    lCol = Range("1:1").Find("Cliente", Range("A1"), xlValues, XlLookAt.xlWhole).Column
    Range(Cells(2, lCol), Cells(lLastRow, lCol)).Copy
    Set wbkTemp = Workbooks.Add
    ActiveCell.PasteSpecial xlPasteValues
    
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    
    With wbkTemp.Sheets(1).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Do While Range("A1").value = ""
        Range("A1").EntireRow.Delete
    Loop
    
    Range("A1", Range("A1048576").End(xlUp)).Copy
    shtCheckEmails.Activate
    Range("A1").PasteSpecial xlPasteValues
    
    lLastRow = Range("B1048576").End(xlUp).Row
    
    ' Fix formats and formulas
    ActiveCell.Copy
    Range(ActiveCell, ActiveCell.End(xlDown)).PasteSpecial xlPasteFormats
    
    ActiveCell.Offset(0, 1).Copy
    Range(ActiveCell.Offset(0, 1), ActiveCell.End(xlDown).Offset(0, 1)).FillDown
    Range("A1").Select
    
    ' Close Input
    shtMenu.Activate
    wbkTemp.Close False
    
    MsgBox "[KMM Dia] atualizado", vbInformation
    
    Set wbkTemp = Nothing
    Set wbkInput = Nothing
    Set cltFile = Nothing
End Sub

Public Sub UpdateKMMMonth()
    Dim lLastRow As Long
    
    Dim cltFile As New Collection
    
    Dim wbkInput As Workbook
    
    
    Set cltFile = Utilities.GetFiles(, , "Selecione o arquivo [KMM Mês]", False)
    
    If cltFile.Count = 0 Then
        MsgBox "Nenhum arquivo selecionado", vbCritical
        Exit Sub
    End If
    
    SetFileAndDate shtMenu.Cells.Find("Arquivos").Offset(4, 0).Row, cltFile(1)
    
    Set wbkInput = Workbooks.Open(cltFile(1), False, True)
    shtKmmMes.Activate
    ClearFilters ActiveSheet
    
    ' Clear old data
    lLastRow = Cells.Find("Nº Romaneio").End(xlDown).Row
    Range("A3:A" & lLastRow).EntireRow.ClearContents
    
    ' Copy new data
    wbkInput.Sheets(1).Activate
    lLastRow = Cells.Find("Nº Romaneio").End(xlDown).Row - 1
    Range("A2:CH" & lLastRow).Copy
    
    ' Paste new data
    shtKmmMes.Activate
    Range("A2").PasteSpecial xlPasteValues
    lLastRow = Cells.Find("Nº Romaneio").End(xlDown).Row
    
    ' Fill down formulas
    Range("CI2:CL2").Copy
    Range("CI3:CL" & lLastRow).PasteSpecial xlPasteFormulas
    
    ' Fill down formats
    Range("A2:CL2").Copy
    Range("A3:CL" & lLastRow).PasteSpecial xlPasteFormats
    
    ' Fix Frota
    Range("1:1").Find("Frota").Offset(1, 0).Activate
    Range(ActiveCell, ActiveCell.End(xlDown)).TextToColumns ActiveCell, xlDelimited, _
        xlTextQualifierNone
    
    Range("A2").Select
    
    ' Close Input
    shtMenu.Activate
    wbkInput.Close False
    
    MsgBox "[KMM Mês] atualizado", vbInformation
    
    Set wbkInput = Nothing
    Set cltFile = Nothing
End Sub

Public Sub CreateStatement()
    Dim lRow As Long
    
    Dim vFilePath As Variant
    
    
    shtExtratoFaturamento.Copy
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight).End(xlDown)).Copy
    Range("B2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range(Selection, Selection.End(xlToRight).End(xlDown)).AutoFilter Field:=1, Criteria1:=0
    Range(Selection.Offset(1, 0), Selection.End(xlToRight).End(xlDown)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    ActiveSheet.ShowAllData
    Range("B3").Select
    
    
    vFilePath = Application.GetSaveAsFilename("Extrato" & Format(Now, "DDMM") & ".xlsb", "Excel Binário (*.xlsb), *.xlsb")
    
    If CStr(vFilePath) = CStr(False) Then
        MsgBox "O arquivo não foi salvo", vbCritical
    Else
        ActiveWorkbook.SaveAs vFilePath, FileFormat:=xlExcel12, CreateBackup:=False, ConflictResolution:=True
        MsgBox "Arquivo " & vbCrLf & vFilePath & " salvo", vbInformation
        ActiveWindow.Close
    End If
    
End Sub

Public Sub CreateBilling()
    Dim lLastRowSource As Long
    Dim lFirstRowTarget As Long
    Dim lLastRowPaste As Long
    Dim lFirstRowPivot As Long
    Dim lLastRowPivot As Long
    Dim lRow As Long
    Dim lLastRow As Long
    Dim lCol As Long
    Dim lColChave As Long
    
    Dim wbkTemp As Workbook
    Dim rngRange As Range
    Dim pvtPivot As PivotTable
    
    Dim objMD5 As New clsMD5
    
    Dim vCol As Variant
    Dim vCols As Variant
    
    
    ' Find target row
    shtKmmMes.Activate
    Call ClearFilters(ActiveSheet)
    lFirstRowTarget = Cells.Find("Nº Romaneio").End(xlDown).Row + 1
    
    ' Find source range
    shtKMMDia.Activate
    Call ClearFilters(ActiveSheet)
    lLastRowSource = Cells.Find("Frota").End(xlDown).Row
    
    ' Filter data and copy
    Cells.AutoFilter Range("A1:XFD1").Find("Em KMM Mês").Column, "Não"
    
    ' Check for Não
    If Range("A1048576").End(xlUp).Row = 1 Then
        GoTo SkipFilter
    End If
    
    ' Grupo
    shtKMMDia.Range("E2:E" & lLastRowSource).SpecialCells(xlCellTypeVisible).Copy _
        shtKmmMes.Range("AP" & lFirstRowTarget)
    lLastRowPaste = shtKmmMes.Range("AP" & lLastRowSource).End(xlDown).Row
    
    ' Frota
    shtKMMDia.Range("AW2:AW" & lLastRowSource).SpecialCells(xlCellTypeVisible).Copy _
        shtKmmMes.Range("BP" & lFirstRowTarget)
    
    ' Placa Tração
    shtKMMDia.Range("AX2:AX" & lLastRowSource).SpecialCells(xlCellTypeVisible).Copy _
        shtKmmMes.Range("H" & lFirstRowTarget)
    
    ' Motorista
    shtKMMDia.Range("W2:W" & lLastRowSource).SpecialCells(xlCellTypeVisible).Copy _
        shtKmmMes.Range("AA" & lFirstRowTarget)
    
    ' Frete Total
    shtKmmMes.Range("O" & lFirstRowTarget & ":O" & lLastRowPaste).value = 0
    
    ' N.o Romaneio
    shtKmmMes.Range("A" & lFirstRowTarget & ":A" & lLastRowPaste).value = "KMM Dia"
    
    ' Copy formulas down
    shtKmmMes.Activate
    Range("CI" & lFirstRowTarget - 1 & ":CL" & lFirstRowTarget - 1).Copy
    Range("CI" & lFirstRowTarget & ":CL" & lLastRowPaste).PasteSpecial xlPasteFormulas
    
    ' Copy formats
    Range("A" & lFirstRowTarget - 1 & ":CL" & lFirstRowTarget - 1).Copy
    Range("A" & lFirstRowTarget & ":CL" & lLastRowPaste).PasteSpecial xlPasteFormats
    Range("A2").Activate
    Application.CutCopyMode = False
    
SkipFilter:
    shtKMMDia.Activate
    Call ClearFilters(ActiveSheet)
    
    ' Update pivots
    shtFarolFaturamento.Activate
    
    For Each pvtPivot In ActiveSheet.PivotTables
        pvtPivot.PivotCache.Refresh
    Next pvtPivot
        
    ' Rank data
    Set rngRange = Cells.Find("ID")
    lFirstRowPivot = rngRange.Row + 1
    lLastRowPivot = rngRange.Offset(0, 3).End(xlDown).Row
    
    ' Clear old IDs
    Range(rngRange.Offset(1, 0), rngRange.End(xlDown)).ClearContents
    Range(rngRange.Offset(1, 2), rngRange.Offset(1, 2).End(xlDown)).ClearContents
    
    For lRow = lFirstRowPivot To lLastRowPivot
        Cells(lRow, rngRange.Column).Value2 = _
            objMD5.md5(Cells(lRow, rngRange.Column).Address)
    Next lRow
    
    shtFarolFaturamento.Range(rngRange, rngRange.End(xlToRight).End(xlDown)).Copy
    
    Set wbkTemp = Workbooks.Add
    ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
    
    Cells.AutoFilter
    Cells.EntireColumn.AutoFit
    lLastRow = Cells.SpecialCells(xlCellTypeLastCell).Row
    
    vCols = _
        Array("Grupo", "Soma de Frete Total", "Frota", "Placa Tração")
    
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        
        For Each vCol In vCols
            lCol = Range("1:1").Find(vCol).Column
            
            .SortFields.Add _
                Key:=Range(Cells(2, lCol).Address & ":" & _
                    Cells(lLastRow, lCol).Address), _
                SortOn:=xlSortOnValues, _
                Order:=IIf(InStr(1, vCol, "Soma de Frete Total", vbTextCompare) > 0, _
                    xlDescending, xlAscending), _
                DataOption:=xlSortNormal
        Next vCol
        
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    lCol = Range("1:1").Find("Rank").Column
    lColChave = Range("1:1").Find("Chave").Column
    Range(Cells(2, lCol), Cells(lLastRow, lCol)).FormulaR1C1 = "=IF(RC" & lColChave & "<>R[-1]C" & lColChave & ",1,R[-1]C+1)"
    Range("A2").Select
    ThisWorkbook.Activate
    Range(rngRange.Offset(1, 2), rngRange.End(xlDown).Offset(0, 2)).FormulaR1C1 = "=VLOOKUP(RC[-2],[" & wbkTemp.Name & "]" & wbkTemp.Sheets(1).Name & "!C1:C3,3,0)"
    Range(rngRange.Offset(1, 2), rngRange.End(xlDown).Offset(0, 2)).Copy
    rngRange.Offset(1, 2).PasteSpecial xlPasteValues
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AI5").Select
    Application.CutCopyMode = False
    
    ' Close Input
    shtMenu.Activate
    wbkTemp.Close False

    MsgBox "Dados copiados e ranqueados", vbInformation
    
    Set wbkTemp = Nothing
    Set rngRange = Nothing
    Set pvtPivot = Nothing
    Set objMD5 = Nothing
End Sub

Public Sub CreateEmails()
    Dim lLastRow As Long
    Dim lRow As Long
    Dim lCol As Long
    
    Dim dSecs As Double
    
    Dim sLog As String
    
    Dim rngRange As Range
    Dim rngLast As Range
    
    Dim objWord As Object 'Word.Document
    Dim objOutlook As Object 'Outlook.Application
    Dim objMail As Object 'Outlook.MailItem
    
    Dim wbkTemp As Workbook
    
    Dim cltLog As New Collection
    
    Dim vLog As Variant
    
    
    
    shtFarolEmail.Activate
    
    cltLog.Add Range("UnLoad2").value & vbTab & "E-mail gerado"
    
    For Each rngRange In Range("tblEmails[Cliente]")
    ' Create required objects to build e-mail
        Set objWord = CreateObject("Word.Application")
        Set objOutlook = CreateObject("Outlook.Application")
        
        Range("Region1").value = rngRange.value
        sLog = Range("Region1").value & vbTab & "Não"
        
        If Range("EmailDataBegin").value = "" _
        Or rngRange.Offset(0, 1).value = "" Then
            GoTo NextRange
        End If
        
        Set rngLast = Range("EmailDataHeader").End(xlToRight).End(xlDown)
        
        ' Create e-mail
        Set objMail = objOutlook.CreateItem(0) 'olMailItem
        
        ' Copy range to e-mail and save it in Draft box
        With objMail
            .To = rngRange.Offset(0, 1).value
            .CC = rngRange.Offset(0, 2).value
            .BCC = ""
            .Subject = Range("Subject").value
            .BodyFormat = 2 'olFormatHTML
            Set objWord = .GetInspector.WordEditor
            
            ' Build e-mail in a temporary wokbook
            Set wbkTemp = Workbooks.Add
            
            ThisWorkbook.Activate
            Range("EmailHeader").Copy
            wbkTemp.Sheets(1).Paste
            
            ThisWorkbook.Activate
            Range(Range("EmailBody").Address & ":" & rngLast.Offset(2, 1).Address).Copy
            wbkTemp.Sheets(1).Activate
            Cells(Cells.SpecialCells(xlCellTypeLastCell).Offset(1, 0).Row, 1).Select
            ActiveSheet.Paste
            
            ThisWorkbook.Activate
            Range("EmailFooter").Copy
            wbkTemp.Sheets(1).Activate
            Cells(Cells.SpecialCells(xlCellTypeLastCell).Offset(1, 0).Row, 1).Select
            ActiveSheet.Paste
            
            ' Adjust column width and remove grid lines
            For lCol = 1 To ActiveSheet.UsedRange.Columns.Count
                Columns(lCol).ColumnWidth = _
                    shtFarolEmail.Columns(lCol).ColumnWidth
            Next lCol
            
            Range("A1").Select
            ActiveWindow.DisplayGridlines = False
            
            ' Copy all data
            ActiveSheet.UsedRange.Copy
            objWord.Range.Paste
            
            Application.DisplayAlerts = False
            wbkTemp.Close False
            Application.DisplayAlerts = True
            Set wbkTemp = Nothing
            
            .Display
            Application.CutCopyMode = False
            .Close 0 'olSave
            
            sLog = Range("Region1").value & vbTab & "Sim"
            
            Application.Wait (Now + TimeValue("0:00:05"))
        End With
        
NextRange:
        cltLog.Add sLog
        Set objMail = Nothing
        Set objWord = Nothing
        Set objOutlook = Nothing
    Next rngRange
    
    ' Create log
    Workbooks.Add
    ActiveCell.value = "Log de criação de e-mails"
    
    For Each vLog In cltLog
        ActiveCell.Offset(1, 0).Select
        Range(ActiveCell, ActiveCell.Offset(0, 1)).value = Split(vLog, vbTab)
    Next vLog
    
    Range("A2", Cells.SpecialCells(xlCellTypeLastCell)).AutoFilter Field:=2, Criteria1:="Sim"
    Cells.EntireColumn.AutoFit
    
    MsgBox "Os e-mails foram salvos em Rascunhos do Outlook", vbInformation
    
End Sub
