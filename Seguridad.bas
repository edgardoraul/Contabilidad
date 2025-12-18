Attribute VB_Name = "Seguridad"
Option Explicit
Public Const clave As String = "WebModerna"
Public bloqueo As Boolean 'Evita recursiones

Sub Proteger(Planilla)
    Call Seguro(Planilla)
End Sub
Sub Desproteger(Planilla)
    Call Inseguro(Planilla)
End Sub
Function Seguro(Planilla)
Debug.Print Planilla
    ThisWorkbook.Worksheets(Planilla).Protect (clave)
End Function
Function Inseguro(Planilla)
    ThisWorkbook.Worksheets(Planilla).Unprotect (clave)
End Function
Function ActualizarFormulas()
    ' Actualiza las referencias de fórmulas para validación de datos
End Function

' Busca en la fila 1 (o la fila que uses para títulos) la columna cuyo título coincida
Function BuscarColumnaPorTitulo(ByVal ws As Worksheet, ByVal tituloBuscado As String) As Long
    Dim c As Range
    tituloBuscado = Trim(UCase(tituloBuscado))
    
    For Each c In ws.Rows(1).Cells   ' MODIFICAR si los títulos no están en la fila 1
        If Trim(UCase(c.Value)) = tituloBuscado Then
            BuscarColumnaPorTitulo = c.Column
            Exit Function
        End If
    Next c

    BuscarColumnaPorTitulo = 0 ' No encontrado
End Function

' Borra el par (codigo / cuenta) según dónde se haya modificado
Sub BorrarPar(ByVal celda As Range)
    Dim titulo As String
    titulo = Trim(UCase(Me.Cells(4, celda.Column).Value))
    If titulo = "CUENTA" Then
        celda.ClearContents
        celda.Offset(0, -1).ClearContents
    ElseIf titulo = "CODIGO" Then
        celda.ClearContents
        celda.Offset(0, 1).ClearContents
    End If
End Sub

' ==========================
' BÚSQUEDA ROBUSTA EN COLUMNA
' ==========================
' ws       -> hoja PDC
' colLetra -> e.g. "A" o "B"
' valor    -> valor a buscar (texto o número)
' Devuelve la celda encontrada o Nothing
Function BuscarEnColumnaRobusto(ByVal ws As Worksheet, ByVal colLetra As String, ByVal valor As Variant) As Range
    Dim rngCol As Range
    Dim colNum As Long
    Dim vText As String
    Dim r As Variant
    Dim FirstAddress As String
    Dim f As Range
    
    colNum = ws.Columns(colLetra).Column
    Set rngCol = ws.Columns(colNum)
    vText = Trim(CStr(valor & ""))
    
    ' 1) Intento Find con cadena (trim)
    On Error Resume Next
    Set f = rngCol.Find(What:=vText, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, After:=rngCol.Cells(rngCol.Rows.Count))
    On Error GoTo 0
    If Not f Is Nothing Then
        Set BuscarEnColumnaRobusto = f
        Exit Function
    End If
    
    ' 2) Intento Match como texto en el rango usado (más fiable)
    On Error Resume Next
    r = Application.Match(vText, rngCol.Resize(ws.UsedRange.Rows.Count).Value, 0)
    On Error GoTo 0
    If Not IsError(r) Then
        ' r es índice dentro del arreglo; convertir a fila real
        Dim filaInicio As Long
        filaInicio = rngCol.Cells(1).Row
        Set BuscarEnColumnaRobusto = ws.Cells(filaInicio + r - 1, colNum)
        Exit Function
    End If
    
    ' 3) Si no, si el valor puede ser numérico, intentar buscar número
    If IsNumeric(valor) Then
        On Error Resume Next
        r = Application.Match(CDbl(valor), rngCol.Resize(ws.UsedRange.Rows.Count).Value, 0)
        On Error GoTo 0
        If Not IsError(r) Then
            Dim filaI As Long
            filaI = rngCol.Cells(1).Row
            Set BuscarEnColumnaRobusto = ws.Cells(filaI + r - 1, colNum)
            Exit Function
        End If
    End If
    
    ' 4) Por último, intento Find buscando dentro de fórmulas (por si estuviera ahí)
    On Error Resume Next
    Set f = rngCol.Find(What:=vText, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
    On Error GoTo 0
    If Not f Is Nothing Then
        Set BuscarEnColumnaRobusto = f
        Exit Function
    End If
    
    ' Si todo falla, devuelve Nothing
    Set BuscarEnColumnaRobusto = Nothing
End Function

