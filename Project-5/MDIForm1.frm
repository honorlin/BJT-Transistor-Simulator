VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "BJT©Ò§jæπº“¿¿"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
Form1.Height = 1200
Form1.Width = 5000
Form1.Top = 0
Form1.Left = 0
End Sub

