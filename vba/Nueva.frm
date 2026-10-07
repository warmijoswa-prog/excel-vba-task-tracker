VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Nueva 
   Caption         =   "Nueva Tarea"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "Nueva.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Nueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
            tema = Nueva.TextBox1.Value
     responsable = Nueva.TextBox2.Value
fecha_compromiso = Nueva.TextBox3.Value
           notas = Nueva.TextBox4.Value
Unload Nueva
End Sub

Private Sub TextBox1_Change()
End Sub

Private Sub TextBox2_Change()
TextBox2 = UCase(TextBox2)
End Sub

Private Sub TextBox3_Enter()
Dim fecha_compromiso As Date
fecha_compromiso = Calendario.GetDate(FirstDayOfWeek:=Monday, SaturdayFontColor:=RGB(250, 0, 0), SundayFontColor:=RGB(250, 0, 0))
TextBox3.Value = Format(fecha_compromiso, "mm/dd/yyyy")
End Sub

Private Sub TextBox4_Change()
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then salir = 10
End Sub
