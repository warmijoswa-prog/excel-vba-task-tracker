Attribute VB_Name = "Nueva_tarea"
Public tema As String
Public responsable As String
Public fecha_compromiso As String
Public notas As String
Public salir As Integer


Sub Nueva_tarea()
    Dim ws As Worksheet
    Dim numero As Integer
    Dim filaNueva As Long
    Dim marca As Range

    Set ws = ActiveSheet
    fecha = Now()

    'mostrar todas las filas sin seleccionar la fila 2
    If ws.AutoFilterMode Then
        If ws.FilterMode Then ws.ShowAllData
    Else
        ws.Rows(2).AutoFilter
    End If

    ws.Columns("K").ClearContents

    fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    filaNueva = fila + 1
    salir = 0
    numero = Application.WorksheetFunction.Max(ws.Columns("A")) + 1

    'una sola celda seleccionada y pantalla repintada antes de abrir la forma
    ws.Range("A" & filaNueva).Select
    Application.ScreenUpdating = True
    DoEvents

    Load Nueva
    Nueva.Show

    If salir = 10 Then GoTo FIN

    On Error GoTo ERROR_MACRO
    Application.ScreenUpdating = False

    With ws
        .Cells(filaNueva, "A").Value = numero
        .Cells(filaNueva, "B").Value = tema
        .Cells(filaNueva, "C").Value = responsable
        .Cells(filaNueva, "D").Value = Format(Date, "mm/dd/yyyy")
        .Cells(filaNueva, "E").Value = fecha_compromiso
        .Cells(filaNueva, "G").FormulaR1C1 = "=IF(R[0]C[-1]="""",""Abierto"",""Cerrado"")"
        .Cells(filaNueva, "H").FormulaR1C1 = "=IF(R[0]C[-2]<>"""","""",IF(R[0]C[-3]-INT(NOW())<0,"""",R[0]C[-3]-INT(NOW())))"
        .Cells(filaNueva, "I").FormulaR1C1 = "=IF(R[0]C[-3]<>"""","""",IF(R[0]C[-4]-INT(NOW())>0,"""",ABS(R[0]C[-4]-INT(NOW()))))"
        .Cells(filaNueva, "J").Value = (Format(Month(fecha), "00") & "." & Format(Day(fecha), "00") & "." & Format(Year(fecha), "00") & " " & notas)
        .Cells(filaNueva, "K").Value = "x"

        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.Range("G3:G10000"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=.Range("A3:A10000"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange ws.Range("A2:K10000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        'ubicar la tarea ingresada por la marca "x" y quedar en la celda de la nota
        Set marca = .Cells(.Rows.Count, "K").End(xlUp)
        marca.ClearContents
        marca.Offset(0, -1).Select
    End With

FIN:
    Application.ScreenUpdating = True
    Exit Sub

ERROR_MACRO:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

