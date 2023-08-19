VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtFarolEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Target.Address <> Range("UnLoad2").Address _
    And Target.Address <> Range("Region1").Address Then
        Exit Sub
    End If
    
    Dim lColTarget As Long
    
    Dim rngLast As Range
    
    Dim vCol As Variant
    Dim vCols As Variant
    
    
    '"Placa Tração", "Grupo Carreta2", "Origem", "Destinatario", "Referência", "Tempo Parado"
    vCols = Array(8, 5, 19, 46, 11, 45)
    shtFarolEmail.Activate
    ClearFilters ActiveSheet
    
    shtMenu.Range("UnLoad1").value = Range("UnLoad2").value
    
    If Range("EmailDataBegin").value <> "" Then
        Range("EmailDataBegin", Range("EmailDataHeader").End(xlToRight).End(xlDown)).ClearContents
    End If
    
    With shtKMMDia
        .Activate
        ClearFilters ActiveSheet
        .Cells.AutoFilter Field:=2, Criteria1:=shtFarolEmail.Range("EmailFilter").value
        Set rngLast = .Range("A1").End(xlDown)
        
        If rngLast.Row = 1048576 Then
            GoTo EndSub
        End If
        
        lColTarget = 2
        
        For Each vCol In vCols
            .Activate
            .Range(.Cells(2, vCol), _
                .Cells(rngLast.Row, vCol)).SpecialCells(xlCellTypeVisible).Copy
            
            shtFarolEmail.Activate
            shtFarolEmail.Cells(Range("EmailDataBegin").Row, lColTarget).PasteSpecial _
                xlPasteValues
            
            lColTarget = lColTarget + 1
        Next vCol
        
        .Activate
        ClearFilters ActiveSheet
    End With
    
    shtFarolEmail.Activate
    Set rngLast = Range("EmailDataHeader").End(xlToRight).End(xlDown)
    ' Range("EmailDataHeader", rngLast).AutoFilter
    
    With shtFarolEmail.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add _
            Key:=Range("G" & Range("EmailDataBegin").Row & ":G" & rngLast.Row), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SortFields.Add _
            Key:=Range("B" & Range("EmailDataBegin").Row & ":B" & rngLast.Row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add _
            Key:=Range("C" & Range("EmailDataBegin").Row & ":C" & rngLast.Row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add _
            Key:=Range("D" & Range("EmailDataBegin").Row & ":D" & rngLast.Row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add _
            Key:=Range("E" & Range("EmailDataBegin").Row & ":E" & rngLast.Row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add _
            Key:=Range("F" & Range("EmailDataBegin").Row & ":F" & rngLast.Row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("Region1").Select
    
EndSub:
    shtKMMDia.Activate
    ClearFilters ActiveSheet
    
    shtFarolEmail.Activate
    
    Set rngLast = Nothing
End Sub
