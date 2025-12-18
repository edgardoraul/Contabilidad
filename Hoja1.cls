VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If bloqueo Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub
    
    Dim wsPDC As Worksheet
    Dim titulo As String
    Dim filaTitulo As Long: filaTitulo = 4
    Dim valorBuscado As Variant
    Dim celdaEncontrada As Range
    Dim colTitulo As Range
    
    Set wsPDC = ThisWorkbook.Worksheets("PDC")
    
    If Target.Row <= filaTitulo Then Exit Sub
    titulo = Trim(UCase(Me.Cells(filaTitulo, Target.Column).Value))
    If titulo <> "CUENTA" And titulo <> "CODIGO" Then Exit Sub
    
    valorBuscado = Target.Value
    If Trim(valorBuscado & "") = "" Then Exit Sub
    
    bloqueo = True
    
    If titulo = "CUENTA" Then
        ' Buscamos en la columna donde están las CUENTAS en PDC
        ' Ajustá "B" por la letra/columna real si fuera distinta
        Set celdaEncontrada = BuscarEnColumnaRobusto(wsPDC, "B", valorBuscado)
    Else
        ' Buscamos en la columna donde están los CODIGOS en PDC
        ' Ajustá "A" por la letra/columna real si fuera distinta
        Set celdaEncontrada = BuscarEnColumnaRobusto(wsPDC, "A", valorBuscado)
    End If
    
    If Not celdaEncontrada Is Nothing Then
        '--- Buscar columna IMPUTABLE en PDC
        Dim colImputable As Long
        colImputable = BuscarColumnaPorTitulo(wsPDC, "IMPUTABLE")
        
        If colImputable = 0 Then
            MsgBox "No se encuentra la columna 'Imputable' en la hoja PDC.", vbCritical, "Error estructural"
            BorrarPar Target
            bloqueo = False
            Exit Sub
        End If
        
        '--- Verificar si la cuenta/código es imputable
        Dim valorImputable As String
        valorImputable = Trim(UCase(wsPDC.Cells(celdaEncontrada.Row, colImputable).Value & ""))
        
        If valorImputable = "NO" Then
            MsgBox "La cuenta/código elegido NO es imputable. Use uno válido.", vbExclamation, "No imputable"
            BorrarPar Target
            bloqueo = False
            Exit Sub
        End If

        
        ' Autocompletar en la otra columna (bidireccional)
        If titulo = "CUENTA" Then
            ' completar CODIGO (columna izquierda)
            Target.Offset(0, -1).Value = celdaEncontrada.Offset(0, -1).Value
        Else
            ' completar CUENTA (columna derecha)
            Target.Offset(0, 1).Value = celdaEncontrada.Offset(0, 1).Value
        End If
    Else
        MsgBox "No existe en el Plan de Cuentas.", vbCritical, "No encontrado"
        BorrarPar Target
    End If
    
    bloqueo = False
End Sub

