VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "Form2"
   ClientHeight    =   7245
   ClientLeft      =   540
   ClientTop       =   930
   ClientWidth     =   10755
   DrawMode        =   7  'Invert
   LinkTopic       =   "Form2"
   ScaleHeight     =   362.25
   ScaleMode       =   2  '點
   ScaleWidth      =   537.75
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command10 
      Height          =   375
      Left            =   7560
      TabIndex        =   32
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "啟動"
      Height          =   375
      Left            =   7440
      TabIndex        =   29
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4560
      TabIndex        =   28
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "減少"
      Height          =   375
      Left            =   9720
      TabIndex        =   27
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "增加"
      Height          =   375
      Left            =   8160
      TabIndex        =   26
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "減少"
      Height          =   375
      Left            =   9240
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "增加"
      Height          =   375
      Left            =   8160
      TabIndex        =   24
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8640
      TabIndex        =   23
      Text            =   "60"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Text            =   "5"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "減少"
      Height          =   375
      Left            =   9720
      TabIndex        =   19
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "增加"
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "減少"
      Height          =   375
      Left            =   9240
      TabIndex        =   17
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "增加"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Text            =   "60"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Text            =   "5"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Text            =   "3"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "60"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line32 
      BorderColor     =   &H00FF0000&
      X1              =   348
      X2              =   348
      Y1              =   282
      Y2              =   294
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00FF0000&
      X1              =   318
      X2              =   318
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00FF0000&
      X1              =   288
      X2              =   288
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line29 
      BorderColor     =   &H00FF0000&
      X1              =   258
      X2              =   258
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FF0000&
      X1              =   228
      X2              =   228
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00FF0000&
      X1              =   198
      X2              =   198
      Y1              =   282
      Y2              =   294
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FF0000&
      X1              =   168
      X2              =   168
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FF0000&
      X1              =   138
      X2              =   138
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FF0000&
      X1              =   108
      X2              =   108
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FF0000&
      X1              =   78
      X2              =   78
      Y1              =   294
      Y2              =   282
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      X1              =   54
      X2              =   42
      Y1              =   348
      Y2              =   348
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      X1              =   54
      X2              =   42
      Y1              =   318
      Y2              =   318
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      X1              =   42
      X2              =   54
      Y1              =   228
      Y2              =   228
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      X1              =   54
      X2              =   42
      Y1              =   258
      Y2              =   258
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   48
      X2              =   48
      Y1              =   228
      Y2              =   348
   End
   Begin VB.Line Line2 
      X1              =   42
      X2              =   348
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FF0000&
      X1              =   348
      X2              =   348
      Y1              =   138
      Y2              =   150
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FF0000&
      X1              =   318
      X2              =   318
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF0000&
      X1              =   288
      X2              =   288
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF0000&
      X1              =   258
      X2              =   258
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF0000&
      X1              =   228
      X2              =   228
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF0000&
      X1              =   198
      X2              =   198
      Y1              =   138
      Y2              =   150
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FF0000&
      X1              =   168
      X2              =   168
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FF0000&
      X1              =   138
      X2              =   138
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FF0000&
      X1              =   108
      X2              =   108
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF0000&
      X1              =   78
      X2              =   78
      Y1              =   150
      Y2              =   138
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      X1              =   54
      X2              =   42
      Y1              =   204
      Y2              =   204
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      X1              =   54
      X2              =   42
      Y1              =   174
      Y2              =   174
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      X1              =   42
      X2              =   54
      Y1              =   84
      Y2              =   84
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      X1              =   54
      X2              =   42
      Y1              =   114
      Y2              =   114
   End
   Begin VB.Line Line3 
      X1              =   42
      X2              =   348
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   48
      X2              =   48
      Y1              =   84
      Y2              =   204
   End
   Begin VB.Label Label8 
      Caption         =   "ms"
      Height          =   255
      Left            =   9360
      TabIndex        =   31
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "ms"
      Height          =   255
      Left            =   9360
      TabIndex        =   30
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Time/\DIV"
      Height          =   255
      Left            =   7200
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "V/DIV"
      Height          =   255
      Left            =   7440
      TabIndex        =   21
      Tag             =   "5"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Time/\DIV"
      Height          =   255
      Left            =   7200
      TabIndex        =   14
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "V/DIV"
      Height          =   255
      Left            =   7440
      TabIndex        =   13
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "輸出"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "輸入"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label29 
      Caption         =   "正弦輸入"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "振幅"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label31 
      Caption         =   "V"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label32 
      Caption         =   "頻率"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label33 
      Caption         =   "Hz"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label34 
      Caption         =   "輸出峰值"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label35 
      Caption         =   "輸出峰谷"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Amp As Integer
Dim Hz As Integer
Dim t As Single
Dim i As Single
Dim V1, V2 As Single
Dim T1, T2 As Single
Dim ff As Long
Dim VDiv1, VDiv2, TDiv1, TDiv2 As Integer
Dim x, y As Integer
Dim mm As Integer
Dim n As Double
Dim temp As Single

Private Sub Command10_Click()

ff = 960

For x = 0 To 50

  ff = ff + x * 120
  
  For y = 0 To 20
      
   Line (ff, 1680 + y * 120)-(ff + 120, 1680 + y * 120), QBColor(3)
  
  Next y

Next x

End Sub

Private Sub Command9_Click()
Call Draw
End Sub

Sub keyin()

Amp = CInt(Text8.Text)
Hz = CInt(Text9.Text)
VDiv1 = CInt(Text3.Text)
VDiv2 = CInt(Text1.Text)
TDiv1 = CInt(Text4.Text)
TDiv2 = CInt(Text2.Text)
t = 1 / Hz
End Sub


Sub Draw()

Call keyin

mm = 960

temp = (t / TDiv1) * 600

For i = 0 To 360
 

 
 V1 = 120 * Amp * Cos(i * n / 180)
 
 mm = mm + temp / 360
 
 Line (mm, 1700 - V1)-(mm + 2, 1700 - V1), QBColor(3)


Next i

End Sub

Private Sub Form_Load()

n = 4 * Atn(1#)

End Sub
