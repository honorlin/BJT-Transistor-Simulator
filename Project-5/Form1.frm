VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "選擇"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   8310
   Begin VB.CommandButton Command2 
      Caption         =   "射集隨偶器"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "共射集放大器"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1920
      Y1              =   240
      Y2              =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Enabled = True
Form2.Top = 5
Form2.Left = 0
Form2.Height = 7700
Form2.Width = 9000

Form2.Show
End Sub

Private Sub Command2_Click()
Form4.Enabled = True
Form4.Top = 5
Form4.Left = 1000
Form4.Height = 7700
Form4.Width = 9000
Form4.Show
End Sub

