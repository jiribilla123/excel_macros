Attribute VB_Name = "Module1"
Sub Imprimir()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim numChars As Integer
    Dim lastRow As Long
    Dim lookupRange As String
    Dim keyColumn As String
    Dim formulaColumn As String

    Set ws = ThisWorkbook.Sheets("Sheet1")
    
        ' Reducir texto para imprimir
    Set rng = ws.Range("D1:D10000")

    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            cell.Value = Left(cell.Value, 40)
        End If
    Next cell
    
        ' El número de guía aparece en la misma columna de cantidades
    Set rng = ws.Range("B1:B10000")

    For Each cell In rng
        If IsNumeric(cell.Value) And cell.Value > 10000 Then
            cell.ClearContents
        End If
    Next cell

        ' Borrar info tabla
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp

        ' Limpiar columna de codigos
    Set rng = ws.Range("B1:B1000")
    
    For Each cell In rng
        If Not cell.Value Like "####-####" Then
            cell.ClearContents
        End If
    Next cell
    
    
    ' Agregar titulos
    Range("A2").Value = "Cantidad"
    Range("B2").Value = "Codigo"
    Range("C2").Value = "Descirpcion"
    Range("D2").Value = "Ubicacion"
    
    'Agregar ubicacion
    
    keyColumn = "B"
    formulaColumn = "D"
    lastRow = ws.Cells(ws.Rows.Count, keyColumn).End(xlUp).Row
    errorText = "Pendiente"
    lookupRange = "'Sheet2'!C:H"

    ' Buscar ubicación en reporte de inventario (hoja2) saltando celdas sin contenido
    For i = 3 To lastRow
        If ws.Cells(i, keyColumn).Value <> "" Then
        ws.Cells(i, formulaColumn).Formula = "=IFERROR(VLOOKUP(" & keyColumn & i & "," & lookupRange & ", 6, FALSE), """ & errorText & """)"
        End If
    Next i
    
    ' Ultimos ajustes para imprimir
    
    ws.Columns("A:D").AutoFit
    Columns("A:B").Select
    Selection.HorizontalAlignment = xlCenter
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

    ' Imprimir. Lo dejo comentado porque en MAC se va directamente a imprimir sin preview
    ' ActiveSheet.PrintPreview
    
End Sub
