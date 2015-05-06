VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   300
   ClientTop       =   450
   ClientWidth     =   8325
   DrawMode        =   7  'Invert
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   8325
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command2 
      Caption         =   "圖形顯示"
      Height          =   375
      Left            =   6360
      TabIndex        =   36
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6240
      TabIndex        =   35
      Text            =   "0.7"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6240
      TabIndex        =   29
      Text            =   "233"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Text            =   "40"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "分析"
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Text            =   "2000"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Text            =   "5100"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Text            =   "100000"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Text            =   "100000"
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line59 
      X1              =   3360
      X2              =   2880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line58 
      X1              =   2880
      X2              =   2880
      Y1              =   2160
      Y2              =   2640
   End
   Begin VB.Line Line57 
      X1              =   2760
      X2              =   2760
      Y1              =   2160
      Y2              =   2640
   End
   Begin VB.Line Line56 
      X1              =   2280
      X2              =   2760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line55 
      X1              =   1440
      X2              =   1080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line54 
      X1              =   1080
      X2              =   1080
      Y1              =   2520
      Y2              =   3000
   End
   Begin VB.Line Line53 
      X1              =   960
      X2              =   960
      Y1              =   2520
      Y2              =   3000
   End
   Begin VB.Line Line52 
      X1              =   480
      X2              =   960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label28 
      Caption         =   "VBE"
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label27 
      Height          =   255
      Left            =   4080
      TabIndex        =   33
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label26 
      Caption         =   "Av="
      Height          =   255
      Left            =   3480
      TabIndex        =   32
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label25 
      Height          =   255
      Left            =   4080
      TabIndex        =   31
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label24 
      Caption         =   "Ri="
      Height          =   255
      Left            =   3480
      TabIndex        =   30
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label23 
      Caption         =   "B"
      Height          =   255
      Left            =   5640
      TabIndex        =   28
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label22 
      Caption         =   "+VCC"
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label21 
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label20 
      Height          =   255
      Left            =   4080
      TabIndex        =   24
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "VC="
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "VB="
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "gm="
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "rt="
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "IC="
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "IB="
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   3600
      Width           =   375
   End
   Begin VB.Line Line51 
      X1              =   3360
      X2              =   7560
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label9 
      Caption         =   "R4"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "R3"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "R2"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "R1"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "R4"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "R3"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "R2"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "R1"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line50 
      X1              =   2160
      X2              =   2400
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line49 
      X1              =   2160
      X2              =   2400
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line48 
      X1              =   2040
      X2              =   2520
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line47 
      X1              =   1320
      X2              =   1560
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line46 
      X1              =   1320
      X2              =   1560
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line45 
      X1              =   1200
      X2              =   1680
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label1 
      Caption         =   "+VCC"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.Line Line44 
      X1              =   2400
      X2              =   2280
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line43 
      X1              =   2160
      X2              =   2280
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line42 
      X1              =   1560
      X2              =   1440
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line41 
      X1              =   1320
      X2              =   1440
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line40 
      X1              =   1440
      X2              =   1440
      Y1              =   4680
      Y2              =   4200
   End
   Begin VB.Line Line39 
      X1              =   1440
      X2              =   1560
      Y1              =   4200
      Y2              =   4080
   End
   Begin VB.Line Line38 
      X1              =   1560
      X2              =   1440
      Y1              =   4080
      Y2              =   3960
   End
   Begin VB.Line Line37 
      X1              =   1440
      X2              =   1560
      Y1              =   3960
      Y2              =   3840
   End
   Begin VB.Line Line36 
      X1              =   1560
      X2              =   1440
      Y1              =   3840
      Y2              =   3720
   End
   Begin VB.Line Line35 
      X1              =   1440
      X2              =   1560
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line34 
      X1              =   1560
      X2              =   1440
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line33 
      X1              =   1440
      X2              =   1440
      Y1              =   3240
      Y2              =   3480
   End
   Begin VB.Line Line32 
      X1              =   2280
      X2              =   2280
      Y1              =   4680
      Y2              =   4200
   End
   Begin VB.Line Line31 
      X1              =   2280
      X2              =   2400
      Y1              =   4200
      Y2              =   4080
   End
   Begin VB.Line Line30 
      X1              =   2400
      X2              =   2280
      Y1              =   4080
      Y2              =   3960
   End
   Begin VB.Line Line29 
      X1              =   2280
      X2              =   2400
      Y1              =   3960
      Y2              =   3840
   End
   Begin VB.Line Line28 
      X1              =   2400
      X2              =   2280
      Y1              =   3840
      Y2              =   3720
   End
   Begin VB.Line Line27 
      X1              =   2280
      X2              =   2400
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line26 
      X1              =   2400
      X2              =   2280
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line25 
      X1              =   2280
      X2              =   2280
      Y1              =   3240
      Y2              =   3480
   End
   Begin VB.Line Line24 
      X1              =   2280
      X2              =   2280
      Y1              =   2520
      Y2              =   2160
   End
   Begin VB.Line Line23 
      X1              =   2280
      X2              =   2400
      Y1              =   2160
      Y2              =   2040
   End
   Begin VB.Line Line22 
      X1              =   2400
      X2              =   2280
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line21 
      X1              =   2280
      X2              =   2400
      Y1              =   1920
      Y2              =   1800
   End
   Begin VB.Line Line20 
      X1              =   2400
      X2              =   2280
      Y1              =   1800
      Y2              =   1680
   End
   Begin VB.Line Line19 
      X1              =   2280
      X2              =   2400
      Y1              =   1680
      Y2              =   1560
   End
   Begin VB.Line Line18 
      X1              =   2400
      X2              =   2280
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line17 
      X1              =   2280
      X2              =   2280
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line Line16 
      X1              =   1440
      X2              =   1440
      Y1              =   3360
      Y2              =   2160
   End
   Begin VB.Line Line15 
      X1              =   1440
      X2              =   1560
      Y1              =   2160
      Y2              =   2040
   End
   Begin VB.Line Line14 
      X1              =   1560
      X2              =   1440
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line13 
      X1              =   1440
      X2              =   1560
      Y1              =   1920
      Y2              =   1800
   End
   Begin VB.Line Line12 
      X1              =   1560
      X2              =   1440
      Y1              =   1800
      Y2              =   1680
   End
   Begin VB.Line Line11 
      X1              =   1440
      X2              =   1560
      Y1              =   1680
      Y2              =   1560
   End
   Begin VB.Line Line10 
      X1              =   1560
      X2              =   1440
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line9 
      X1              =   1440
      X2              =   1440
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line Line8 
      X1              =   1440
      X2              =   1920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line7 
      X1              =   2280
      X2              =   2280
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line Line6 
      X1              =   2280
      X2              =   2280
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line Line5 
      X1              =   2160
      X2              =   2280
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line4 
      X1              =   2280
      X2              =   2280
      Y1              =   3000
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   2280
      X2              =   1920
      Y1              =   3120
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   1920
      Y1              =   2520
      Y2              =   2760
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   2520
      Y2              =   3120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public R1, R2, R3, R4 As Long
Public VCC As Integer
Public B As Integer
Public IB, IC As Double
Public gm As Double
Public rt As Double
Public VB, VC As Double
Public TempR, TempV As Long
Public Tempx As Long
Public Av As Double
Public Ri As Long
Public VBE As Single

Private Sub Command1_Click()

Call keyin

TempV = (R2 / (R1 + R2)) * VCC

TempR = (R1 * R2) / (R1 + R2)

IB = (TempV - VBE) / (TempR + (B + 1) * R4)

IC = IB * B

gm = IC / (25 * 10 ^ -3)

rt = (25 * 10 ^ -3) / IB

VC = VCC - IC * R3

VB = TempV - (IB * TempR)


Tempx = R4 * (B + 1) + rt

Ri = (Tempx * TempR) / (Tempx + TempR)

Av = gm * R3 * (rt / (R4 * (B + 1) + rt))


Label14.Caption = IB

Label15.Caption = IC

Label16.Caption = rt

Label17.Caption = gm

Label20.Caption = VB

Label21.Caption = VC

Label25.Caption = Ri

Label27.Caption = Av

End Sub

Sub keyin()
R1 = CLng(Text1.Text)
R2 = CLng(Text2.Text)
R3 = CLng(Text3.Text)
R4 = CLng(Text4.Text)
VCC = CInt(Text5.Text)
B = CInt(Text6.Text)
VBE = CSng(Text7.Text)
End Sub

Private Sub Command2_Click()
Form2.Visible = True

End Sub



Private Sub Form_Load()

End Sub
