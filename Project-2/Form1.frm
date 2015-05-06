VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   7830
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "分析"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Text            =   "40"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Text            =   "233"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Text            =   "100000"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Text            =   "100000"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Text            =   "680"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   3960
      TabIndex        =   32
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Ro="
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   6840
      Width           =   375
   End
   Begin VB.Line Line20 
      X1              =   2520
      X2              =   7320
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label19 
      Caption         =   "VE="
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label21 
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "IB="
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "IC="
      Height          =   255
      Left            =   3360
      TabIndex        =   27
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "rt="
      Height          =   255
      Left            =   3360
      TabIndex        =   26
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "gm="
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   3960
      TabIndex        =   23
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "VB="
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label20 
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label24 
      Caption         =   "Ri="
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label25 
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label26 
      Caption         =   "Av="
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label27 
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label22 
      Caption         =   "+VCC"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "B"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "R1"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "R2"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "R3"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "+VCC"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Line Line19 
      X1              =   960
      X2              =   840
      Y1              =   720
      Y2              =   600
   End
   Begin VB.Line Line18 
      X1              =   720
      X2              =   840
      Y1              =   720
      Y2              =   600
   End
   Begin VB.Line Line17 
      X1              =   840
      X2              =   1680
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line7 
      X1              =   1680
      X2              =   1680
      Y1              =   2160
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   1320
      Y1              =   2160
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   1320
      Y1              =   2160
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   1680
      X2              =   1320
      Y1              =   2760
      Y2              =   2520
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   1680
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line5 
      X1              =   1560
      X2              =   1680
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   1680
      X2              =   1680
      Y1              =   2760
      Y2              =   3000
   End
   Begin VB.Line Line8 
      X1              =   840
      X2              =   1320
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line9 
      X1              =   840
      X2              =   840
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Line Line10 
      X1              =   960
      X2              =   840
      Y1              =   1200
      Y2              =   1080
   End
   Begin VB.Line Line11 
      X1              =   840
      X2              =   960
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line12 
      X1              =   960
      X2              =   840
      Y1              =   1440
      Y2              =   1320
   End
   Begin VB.Line Line13 
      X1              =   840
      X2              =   960
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line14 
      X1              =   960
      X2              =   840
      Y1              =   1680
      Y2              =   1560
   End
   Begin VB.Line Line15 
      X1              =   840
      X2              =   960
      Y1              =   1800
      Y2              =   1680
   End
   Begin VB.Line Line16 
      X1              =   840
      X2              =   840
      Y1              =   3000
      Y2              =   1800
   End
   Begin VB.Line Line25 
      X1              =   1680
      X2              =   1680
      Y1              =   2880
      Y2              =   3120
   End
   Begin VB.Line Line26 
      X1              =   1800
      X2              =   1680
      Y1              =   3240
      Y2              =   3120
   End
   Begin VB.Line Line27 
      X1              =   1680
      X2              =   1800
      Y1              =   3360
      Y2              =   3240
   End
   Begin VB.Line Line28 
      X1              =   1800
      X2              =   1680
      Y1              =   3480
      Y2              =   3360
   End
   Begin VB.Line Line29 
      X1              =   1680
      X2              =   1800
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line30 
      X1              =   1800
      X2              =   1680
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line31 
      X1              =   1680
      X2              =   1800
      Y1              =   3840
      Y2              =   3720
   End
   Begin VB.Line Line32 
      X1              =   1680
      X2              =   1680
      Y1              =   4320
      Y2              =   3840
   End
   Begin VB.Line Line33 
      X1              =   840
      X2              =   840
      Y1              =   2880
      Y2              =   3120
   End
   Begin VB.Line Line34 
      X1              =   960
      X2              =   840
      Y1              =   3240
      Y2              =   3120
   End
   Begin VB.Line Line35 
      X1              =   840
      X2              =   960
      Y1              =   3360
      Y2              =   3240
   End
   Begin VB.Line Line36 
      X1              =   960
      X2              =   840
      Y1              =   3480
      Y2              =   3360
   End
   Begin VB.Line Line37 
      X1              =   840
      X2              =   960
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line38 
      X1              =   960
      X2              =   840
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line39 
      X1              =   840
      X2              =   960
      Y1              =   3840
      Y2              =   3720
   End
   Begin VB.Line Line40 
      X1              =   840
      X2              =   840
      Y1              =   4320
      Y2              =   3840
   End
   Begin VB.Line Line45 
      X1              =   600
      X2              =   1080
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line46 
      X1              =   720
      X2              =   960
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line47 
      X1              =   720
      X2              =   960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line48 
      X1              =   1440
      X2              =   1920
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line49 
      X1              =   1560
      X2              =   1800
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line50 
      X1              =   1560
      X2              =   1800
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label2 
      Caption         =   "R1"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "R2"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "R3"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R1, R2, R3 As Long
Dim VCC As Single
Dim B As Integer
Dim IB, IC As Double
Dim gm As Double
Dim rt As Double
Dim VB, VE As Double
Dim TempR, TempV As Long
Dim Tempx As Long
Dim Ro As Double
Dim Av As Double
Dim Ri As Double

Private Sub Command1_Click()

Call keyin

TempV = (R2 / (R1 + R2)) * VCC

TempR = (R1 * R2) / (R1 + R2)

IB = (TempV - 0.7) / (TempR + R3 * (B + 1))

IC = IB * B

gm = IC / (25 * 10 ^ -3)

rt = (25 * 10 ^ -3) / IB

VB = TempV - (IB * TempR)

VE = (IB + IC) * R3

Tempx = R3 * (B + 1) + rt

Ri = (TempR * Tempx) / (TempR + Tempx)

Av = ((B + 1) * R3) / (rt + (B + 1) * R3)

Ro = (680 * (rt / (B + 1))) / (680 + (rt / (B + 1)))

Label14.Caption = IB

Label15.Caption = IC

Label16.Caption = rt

Label17.Caption = gm

Label20.Caption = VB

Label21.Caption = VE

Label25.Caption = Ri

Label27.Caption = Av

Label9.Caption = Ro

End Sub

Sub keyin()
R1 = CLng(Text1.Text)
R2 = CLng(Text2.Text)
R3 = CLng(Text3.Text)
VCC = CSng(Text5.Text)
B = CInt(Text6.Text)


End Sub

