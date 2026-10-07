Attribute VB_Name = "Ordenar"
Sub Ordenar()
    Dim ws As Worksheet
    Dim abiertas As Long

    Set ws = ActiveSheet

    On Error GoTo ERROR_MACRO
    Application.ScreenUpdating = False

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("G3:G10000"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range("A3:A10000"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A2:J10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'las tareas abiertas quedan primero; la última está en la fila 2 + cantidad de abiertas
    abiertas = Application.WorksheetFunction.CountIf(ws.Range("G3:G10000"), "Abierto")
    If abiertas > 0 Then ws.Range("J" & abiertas + 2).Select

    Application.ScreenUpdating = True
    Exit Sub

ERROR_MACRO:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

