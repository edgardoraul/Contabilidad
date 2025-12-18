Attribute VB_Name = "Contabilidad"
Option Explicit
Public Ultima As Long
Function UltimaFila(Hoja)
    Ultima = Worksheets(Hoja).Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print "Pestaña: " & Hoja & " - Ultima Fila: " & Ultima
End Function

Function ValidarAsiento()
    ' Se obtienen datos para validar el número de los asientos.
    UltimaFila ("Diario")
    If Ultima = 1 Then
        Ultima = 2
    End If
    NumAsientoOld = Worksheets("Diario").Cells(Ultima, 1).Value

    If NumAsientoOld = 0 Then
        NumAsientoNew = 1
        Worksheets("Asientos").Range("C3").Value = NumAsientoNew
    End If
    NumAsientoNew = Worksheets("Asientos").Range("C3").Value
End Function


