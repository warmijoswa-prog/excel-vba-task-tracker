Attribute VB_Name = "Cerrar_tarea"
Public notas2 As String
Public salir2 As Integer


Sub Cerrar_tarea()
    Dim ws As Worksheet
    Dim fecha As Date
    Dim fila As Long
    Dim comentario As String
    Dim numero As Variant
    Dim pos As Variant

    Set ws = ActiveSheet
    fecha = Now()
    salir2 = 0
    fila = ActiveCell.Row

    If fila < 3 Then GoTo FIN1
    If ws.Range("A" & fila).Value = "" Then GoTo FIN1

    'pantalla repintada antes de abrir la forma
    Application.ScreenUpdating = True
    DoEvents

    Load Cierre
    Cierre.Show

    If salir2 = 10 Then GoTo FIN2

    On Error GoTo ERROR_MACRO
    Application.ScreenUpdating = False

    With ws
        numero = .Range("A" & fila).Value


        .Range("F" & fila).NumberFormat = "@"
        .Range("F" & fila).Value = Format(Date, "mm/dd/yyyy")

        comentario = .Range("J" & fila).Value
        .Range("J" & fila).Value = Format(Month(fecha), "00") & "." & Format(Day(fecha), "00") & "." & Format(Year(fecha), "00") & " " & notas2 & Chr(10) & comentario

        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.Range("G3:G10000"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=.Range("A3:A10000"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange ws.Range("A2:J10000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        'regresar a la nota de la tarea cerrada, que el orden mueve al grupo "Cerrado"
        pos = Application.Match(numero, .Range("A3:A10000"), 0)
        If IsError(pos) Then
            .Range("J" & fila).Select
        Else
            .Range("J" & pos + 2).Select
        End If
    End With

    GoTo FIN2

FIN1:
    MsgBox "Seleccione una tarea"

FIN2:
    Application.ScreenUpdating = True
    Exit Sub

ERROR_MACRO:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

