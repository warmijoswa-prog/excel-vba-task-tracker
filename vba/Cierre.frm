VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cierre 
   Caption         =   "Cierre"
   ClientHeight    =   1410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Cierre.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
notas2 = Cierre.TextBox1.Value
Unload Cierre
End Sub

Private Sub TextBox1_Change()
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then salir2 = 10
End Sub
