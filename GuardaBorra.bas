Attribute VB_Name = "GuardaBorra"
Option Explicit
Public NumAsientoNew As Long
Public NumAsientoOld As Long

Sub GuardarAsiento()
    Dim i As Integer
    Dim UltimaAsientos As Integer
    Dim UltimaDiario As Long
    UltimaAsientos = Cells(Rows.Count, 1).End(xlUp).Row - 3
    UltimaDiario = Worksheets("Diario").Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print UltimaAsientos, UltimaDiario
    
    ' 1) Validaciones.
    Call ValidarAsiento
    Call Desproteger("Diario")
    
    ' 2) Guarda las cuentas del debe.
    For i = 5 To UltimaAsientos

        If Cells(i, 1).Value <> "" Then
            
            ' Incrementa la última fila
            UltimaDiario = UltimaDiario + 1
        
            ' La fecha del asiento
            With Worksheets("Diario")
                ' A - Copia el número del asiento
                .Cells(UltimaDiario, 1).Value = NumAsientoNew
                
                ' B - Copia la fecha del asiento
                .Cells(UltimaDiario, 2).Value = Worksheets("Asientos").Cells(3, 7).Value
                
                ' C - Copia el código de la cuenta del debe
                .Cells(UltimaDiario, 3).Value = Worksheets("Asientos").Cells(i, 1).Value
                
                ' D - Fórmula para obtener la cuenta
                .Cells(UltimaDiario, 4).FormulaLocal = "=BUSCARV(C" & UltimaDiario & ";Tabla4[#Todo];2;FALSO)"
                
                ' E - Copia el importe de la cuenta del Debe
                .Cells(UltimaDiario, 5).Value = Worksheets("Asientos").Cells(i, 3).Value
                
                ' F - Copia la leyenda
                .Cells(UltimaDiario, 7).Value = Worksheets("Asientos").Cells(UltimaAsientos + 3, 3).Value
            End With
        Else
            Exit For
        End If
    Next i
    
    ' 3) Guarda las cuentas del haber.
    For i = 5 To UltimaAsientos
        If Cells(i, 5).Value <> "" Then
            
            ' Incrementa la última fila
            UltimaDiario = UltimaDiario + 1
        
            ' La fecha del asiento
            With Worksheets("Diario")
                ' A - Copia el número del asiento
                .Cells(UltimaDiario, 1).Value = NumAsientoNew
                
                ' B - Copia la fecha del asiento
                .Cells(UltimaDiario, 2).Value = Worksheets("Asientos").Cells(3, 7).Value
                
                ' C - Copia el código de la cuenta del Haber
                .Cells(UltimaDiario, 3).Value = Worksheets("Asientos").Cells(i, 5).Value
                
                ' D - Fórmula para obtener la cuenta
                .Cells(UltimaDiario, 4).FormulaLocal = "=BUSCARV(C" & UltimaDiario & ";Tabla4[#Todo];2;FALSO)"
                
                ' E - Copia el importe de la cuenta del Haber
                .Cells(UltimaDiario, 6).Value = Worksheets("Asientos").Cells(i, 7).Value
                
                ' F - Copia la leyenda
                .Cells(UltimaDiario, 7).Value = Worksheets("Asientos").Cells(UltimaAsientos + 3, 3).Value

            End With
        End If
    Next i
    
    
    ' Guarda el asiento, siempre con un número incremental
    Debug.Print "Asiento Nº: " & NumAsientoNew & " - Guardado ;-)"
    
    ' Incrementar el número del asiento
    NumAsientoNew = NumAsientoNew + 1
    Worksheets("Asientos").Cells(3, 3).Value = NumAsientoNew
    MsgBox "Asiento Nº " & NumAsientoNew & " guardado."
    
    Call Proteger("Diario")
End Sub
Sub BorrarAsiento()
    Dim NumAsiento As Long
    Dim i As Long
    Dim UltimaFila As Long
    Dim CantEncontrados As Long
    Dim PrimeraFila As Long
    Dim UltimoAsiento As Long
    Dim Rng As Range
    Dim FirstAddress As String
    
    ' 0) Abrir la escritura
    Call Desproteger("Diario")
    
    ' 1) Pedir número de asiento
    NumAsiento = CLng(InputBox("Vas a borrar mediante un contraasiento. " & vbCrLf & "Ingrese el número de asiento a buscar:", "Borrar Asiento"))
    If NumAsiento = 0 Then Exit Sub   ' Cancelado o inválido
    
    ' 2) Determinar última fila de Diario
    UltimaFila = Worksheets("Diario").Cells(Rows.Count, 1).End(xlUp).Row
    UltimoAsiento = Worksheets("Diario").Cells(UltimaFila, 1).Value
    
    ' 3) Buscar en columna A
    CantEncontrados = 0
    PrimeraFila = 0
    
    With Worksheets("Diario").Range("A2:A" & UltimaFila)
        Set Rng = .Find(What:=NumAsiento, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Rng Is Nothing Then
            MsgBox "El asiento Nº " & NumAsiento & " no existe. Intentá con otro."
            GoTo Salida
        End If
        
        ' Primer hallazgo
        PrimeraFila = Rng.Row
        CantEncontrados = 1
        
        FirstAddress = Rng.Address
        
        ' Buscar restantes
        Do
            Set Rng = .FindNext(Rng)
            If Rng Is Nothing Then Exit Do
            If Rng.Address = FirstAddress Then Exit Do
            
            CantEncontrados = CantEncontrados + 1
            
        Loop
        
    End With
   
   
    ' 4) Resultado
    Debug.Print "Encontrados: " & CantEncontrados & vbCrLf & _
           "Primera fila: " & IIf(PrimeraFila = 0, "No encontrado", PrimeraFila)
    
    ' 5) Borra mediante un bucle con los importes al revés
    For i = 1 To CantEncontrados
        With Worksheets("Diario")
            ' Nº contrasiento
            .Cells(i + UltimaFila, 1).Value = UltimoAsiento + 1
            
            ' Fecha del contrasiento
            .Cells(i + UltimaFila, 2).Value = Format(Date, "dd/mm/yyyy")
            
            ' Código original
            .Cells(i + UltimaFila, 3).Value = .Cells(i + PrimeraFila - 1, 3).Value
            
            ' Detalle cuenta
            .Cells(i + PrimeraFila - 1, 4).Copy
            .Cells(i + UltimaFila, 4).PasteSpecial xlPasteAll
            
            ' Monto haber
            .Cells(i + UltimaFila, 6).Value = .Cells(i + PrimeraFila - 1, 5).Value
            
            ' Monto debe
            .Cells(i + UltimaFila, 5).Value = .Cells(i + PrimeraFila - 1, 6).Value
          
            ' Leyenda contraseña
            .Cells(i + UltimaFila, 7).Value = "Contrasiento Nº " & NumAsiento
            
            ' Limpiar la selección
            Application.CutCopyMode = False
        End With
    Next i
    
    ' 6) Avisar
    MsgBox "El asiento Nº " & NumAsiento & " se anuló mediante" & vbCrLf & "Contrasiento Nº " & UltimoAsiento
Salida:
    ' 7) Cerrar el diario
    Call Proteger("Diario")
End Sub


