VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "共射集放大器"
   ClientHeight    =   7185
   ClientLeft      =   1530
   ClientTop       =   660
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   7935
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6000
      TabIndex        =   38
      Text            =   "0.5"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Text            =   "1500000"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Text            =   "160000"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Text            =   "12000"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Text            =   "2000"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "分析"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Text            =   "60"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Text            =   "133"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Text            =   "0.7"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "圖形顯示"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label29 
      Caption         =   "VBC"
      Height          =   255
      Left            =   5400
      TabIndex        =   37
      Top             =   1800
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   1680
      Y1              =   2040
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   1680
      Y1              =   2040
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   2040
      X2              =   1680
      Y1              =   2640
      Y2              =   2400
   End
   Begin VB.Line Line4 
      X1              =   2040
      X2              =   2040
      Y1              =   2520
      Y2              =   2640
   End
   Begin VB.Line Line5 
      X1              =   1920
      X2              =   2040
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line6 
      X1              =   2040
      X2              =   2040
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Line Line7 
      X1              =   2040
      X2              =   2040
      Y1              =   2040
      Y2              =   1800
   End
   Begin VB.Line Line8 
      X1              =   1200
      X2              =   1680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line9 
      X1              =   1200
      X2              =   1200
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line10 
      X1              =   1320
      X2              =   1200
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line11 
      X1              =   1200
      X2              =   1320
      Y1              =   1200
      Y2              =   1080
   End
   Begin VB.Line Line12 
      X1              =   1320
      X2              =   1200
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line13 
      X1              =   1200
      X2              =   1320
      Y1              =   1440
      Y2              =   1320
   End
   Begin VB.Line Line14 
      X1              =   1320
      X2              =   1200
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line15 
      X1              =   1200
      X2              =   1320
      Y1              =   1680
      Y2              =   1560
   End
   Begin VB.Line Line16 
      X1              =   1200
      X2              =   1200
      Y1              =   2880
      Y2              =   1680
   End
   Begin VB.Line Line17 
      X1              =   2040
      X2              =   2040
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line18 
      X1              =   2160
      X2              =   2040
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line19 
      X1              =   2040
      X2              =   2160
      Y1              =   1200
      Y2              =   1080
   End
   Begin VB.Line Line20 
      X1              =   2160
      X2              =   2040
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line21 
      X1              =   2040
      X2              =   2160
      Y1              =   1440
      Y2              =   1320
   End
   Begin VB.Line Line22 
      X1              =   2160
      X2              =   2040
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line23 
      X1              =   2040
      X2              =   2160
      Y1              =   1680
      Y2              =   1560
   End
   Begin VB.Line Line24 
      X1              =   2040
      X2              =   2040
      Y1              =   2040
      Y2              =   1680
   End
   Begin VB.Line Line25 
      X1              =   2040
      X2              =   2040
      Y1              =   2760
      Y2              =   3000
   End
   Begin VB.Line Line26 
      X1              =   2160
      X2              =   2040
      Y1              =   3120
      Y2              =   3000
   End
   Begin VB.Line Line27 
      X1              =   2040
      X2              =   2160
      Y1              =   3240
      Y2              =   3120
   End
   Begin VB.Line Line28 
      X1              =   2160
      X2              =   2040
      Y1              =   3360
      Y2              =   3240
   End
   Begin VB.Line Line29 
      X1              =   2040
      X2              =   2160
      Y1              =   3480
      Y2              =   3360
   End
   Begin VB.Line Line30 
      X1              =   2160
      X2              =   2040
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line31 
      X1              =   2040
      X2              =   2160
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line32 
      X1              =   2040
      X2              =   2040
      Y1              =   4200
      Y2              =   3720
   End
   Begin VB.Line Line33 
      X1              =   1200
      X2              =   1200
      Y1              =   2760
      Y2              =   3000
   End
   Begin VB.Line Line34 
      X1              =   1320
      X2              =   1200
      Y1              =   3120
      Y2              =   3000
   End
   Begin VB.Line Line35 
      X1              =   1200
      X2              =   1320
      Y1              =   3240
      Y2              =   3120
   End
   Begin VB.Line Line36 
      X1              =   1320
      X2              =   1200
      Y1              =   3360
      Y2              =   3240
   End
   Begin VB.Line Line37 
      X1              =   1200
      X2              =   1320
      Y1              =   3480
      Y2              =   3360
   End
   Begin VB.Line Line38 
      X1              =   1320
      X2              =   1200
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line39 
      X1              =   1200
      X2              =   1320
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line40 
      X1              =   1200
      X2              =   1200
      Y1              =   4200
      Y2              =   3720
   End
   Begin VB.Line Line41 
      X1              =   1080
      X2              =   1200
      Y1              =   840
      Y2              =   720
   End
   Begin VB.Line Line42 
      X1              =   1320
      X2              =   1200
      Y1              =   840
      Y2              =   720
   End
   Begin VB.Line Line43 
      X1              =   1920
      X2              =   2040
      Y1              =   840
      Y2              =   720
   End
   Begin VB.Line Line44 
      X1              =   2160
      X2              =   2040
      Y1              =   840
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "+VCC"
      Height          =   255
      Left            =   1440
      TabIndex        =   36
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line45 
      X1              =   960
      X2              =   1440
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line46 
      X1              =   1080
      X2              =   1320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line47 
      X1              =   1080
      X2              =   1320
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line48 
      X1              =   1800
      X2              =   2280
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line49 
      X1              =   1920
      X2              =   2160
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line50 
      X1              =   1920
      X2              =   2160
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label2 
      Caption         =   "R1"
      Height          =   255
      Left            =   720
      TabIndex        =   35
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "R2"
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "R3"
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "R4"
      Height          =   255
      Left            =   2280
      TabIndex        =   32
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "R1"
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "R2"
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "R3"
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "R4"
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   1800
      Width           =   255
   End
   Begin VB.Line Line51 
      X1              =   3120
      X2              =   7320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label10 
      Caption         =   "IB="
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "IC="
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "rt="
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "gm="
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "VB="
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "VC="
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label20 
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label21 
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label22 
      Caption         =   "+VCC"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "B"
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label24 
      Caption         =   "Ri="
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label25 
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label26 
      Caption         =   "Av="
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label27 
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label28 
      Caption         =   "VBE"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Line Line52 
      X1              =   240
      X2              =   720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line53 
      X1              =   720
      X2              =   720
      Y1              =   2040
      Y2              =   2520
   End
   Begin VB.Line Line54 
      X1              =   840
      X2              =   840
      Y1              =   2040
      Y2              =   2520
   End
   Begin VB.Line Line55 
      X1              =   1200
      X2              =   840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line56 
      X1              =   2040
      X2              =   2520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line57 
      X1              =   2520
      X2              =   2520
      Y1              =   1680
      Y2              =   2160
   End
   Begin VB.Line Line58 
      X1              =   2640
      X2              =   2640
      Y1              =   1680
      Y2              =   2160
   End
   Begin VB.Line Line59 
      X1              =   3120
      X2              =   2640
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "form1"
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
Public VBE, VBC As Single

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

Command2.Enabled = True

End Sub

Sub keyin()
R1 = CLng(Text1.Text)
R2 = CLng(Text2.Text)
R3 = CLng(Text3.Text)
R4 = CLng(Text4.Text)
VCC = CInt(Text5.Text)
B = CInt(Text6.Text)
VBE = CSng(Text7.Text)
VBC = CSng(Text8.Text)
End Sub

Private Sub Command2_Click()
Form2.Visible = True
Form2.Show
End Sub



