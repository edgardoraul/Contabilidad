VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Hojas As Object  ' Dictionary global para acceder desde cualquier módulo
Public Planilla As Variant
Public Hojitas As Object


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = False
    For Each Planilla In ThisWorkbook.Worksheets
        Call Desproteger(Planilla.Name)
    Next Planilla
End Sub

Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    ' Recorrer todas las hojas visibles
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Asientos" Then
            Debug.Print ws.Name
            
            ' Coloca la fecha de hoy por defecto
            ws.Range("G3").Value = "=Today()"
            
            ' Coloca el número de asiento
            Call UltimaFila("Diario")
            ws.Range("C3").Value = Worksheets("Diario").Cells(Ultima, 1).Value
        End If
        
        If ws.Visible = xlSheetVisible Then
            Call Proteger(ws.Name)
        End If
    Next ws
    
End Sub

