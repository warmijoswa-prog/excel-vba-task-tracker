VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Actualizacion 
   Caption         =   "Actualizacion"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   OleObjectBlob   =   "Actualizacion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Actualizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
notas1 = Actualizacion.TextBox1.Value
Unload Actualizacion
End Sub

Private Sub TextBox1_Change()
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then salir1 = 10
End Sub
