VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Target.Address = Range("UnLoad1").Address Then
        
        If shtFarolEmail.Range("UnLoad2").value <> Target.value Then
            shtFarolEmail.Range("UnLoad2").value = Target.value
        End If
        
    End If
    
End Sub
