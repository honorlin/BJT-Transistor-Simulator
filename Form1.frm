VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   11880
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   10440
      TabIndex        =   57
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   10200
      TabIndex        =   54
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   10200
      TabIndex        =   53
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "秨﹍"
      Height          =   375
      Left            =   10200
      TabIndex        =   52
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "代刚"
      Height          =   375
      Left            =   7920
      TabIndex        =   41
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "秨﹍"
      Height          =   375
      Left            =   7920
      TabIndex        =   40
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   3720
      TabIndex        =   39
      Text            =   "-100"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   3720
      TabIndex        =   38
      Text            =   "100"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   840
      TabIndex        =   37
      Text            =   "10"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   2160
      TabIndex        =   36
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   4200
      TabIndex        =   34
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   7200
      TabIndex        =   32
      Text            =   "-6.022"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   4680
      TabIndex        =   30
      Text            =   "-0.000108"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   2160
      TabIndex        =   28
      Text            =   "0.000018"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Text            =   "0.00368"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Text            =   "-13720"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Text            =   "-5018"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Text            =   "0.5"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Text            =   "0.000909"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Text            =   "10"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Text            =   "100"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Text            =   "-100"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "秨﹍"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "代刚"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label31 
      Caption         =   "舼盿Ω计"
      Height          =   255
      Left            =   10440
      TabIndex        =   58
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label30 
      Caption         =   "R3="
      Height          =   255
      Left            =   9600
      TabIndex        =   56
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label29 
      Caption         =   "R1="
      Height          =   255
      Left            =   9600
      TabIndex        =   55
      Top             =   2040
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   9720
      X2              =   9720
      Y1              =   480
      Y2              =   7320
   End
   Begin VB.Label Label28 
      Caption         =   "代刚R3=(-)"
      Height          =   375
      Left            =   2640
      TabIndex        =   51
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label27 
      Caption         =   "代刚R3=(+)"
      Height          =   255
      Left            =   2640
      TabIndex        =   50
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label26 
      Caption         =   "f(a)="
      Height          =   375
      Left            =   5520
      TabIndex        =   49
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label25 
      Caption         =   "f(b)="
      Height          =   375
      Left            =   5520
      TabIndex        =   48
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label24 
      Height          =   375
      Left            =   6120
      TabIndex        =   47
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label23 
      Height          =   375
      Left            =   6120
      TabIndex        =   46
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label22 
      Caption         =   "R1="
      Height          =   375
      Left            =   360
      TabIndex        =   45
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "舼盿Ω计"
      Height          =   255
      Left            =   1200
      TabIndex        =   44
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   "R3="
      Height          =   255
      Left            =   3600
      TabIndex        =   43
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Ans="
      Height          =   375
      Left            =   3600
      TabIndex        =   42
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "R3 = 0"
      Height          =   255
      Left            =   8640
      TabIndex        =   33
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "R1+"
      Height          =   255
      Left            =   6720
      TabIndex        =   31
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "R1R3+"
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "R3^2+ "
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Ans="
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "R1="
      Height          =   255
      Left            =   3360
      TabIndex        =   22
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "舼盿Ω计"
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "R3  =   0"
      Height          =   255
      Left            =   8760
      TabIndex        =   19
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "R1+"
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "R1R3+"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "R1^2+"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "R3="
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "f(b)="
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "f(a)="
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "代刚R1=(+)"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "代刚R1=(-)"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1, x2, x3, x4 As Double
Dim Ans, a, b, tempa, tempqq, qq As Double
Dim R1, R3 As Double

Dim y1, y2, y3, y4 As Double
Dim Ans2, a2, b2, tempa2, tempqq2, qq2 As Double
Dim R1a, R3a As Double


Private Sub Command1_Click()

R1 = CDbl(Text1.Text)
Call abc
Label5.Caption = Ans
R1 = CDbl(Text2.Text)
Call abc
Label6.Caption = Ans

End Sub


Private Sub Command2_Click()
Call FindR1
End Sub

Sub abc()

x1 = CDbl(Text4.Text)
x2 = CDbl(Text5.Text)
x3 = CDbl(Text6.Text)
x4 = CDbl(Text7.Text)
R3 = CDbl(Text3.Text)

Ans = x1 * R1 ^ 2 + x2 * R1 * R3 + x3 * R1 + x4 * R3




End Sub

Private Sub Command3_Click()
Call FindR3
End Sub

Private Sub Command4_Click()
R3a = CDbl(Text20.Text)
Call def
Label24.Caption = Ans2
R3a = CDbl(Text19.Text)
Call def
Label23.Caption = Ans2

End Sub

Sub def()
y1 = CDbl(Text11.Text)
y2 = CDbl(Text12.Text)
y3 = CDbl(Text13.Text)
y4 = CDbl(Text14.Text)
R1a = CDbl(Text18.Text)

Ans2 = y1 * R3a ^ 2 + y2 * R1a * R3a + y3 * R1a + y4 * R3a



End Sub

Sub FindR1()
a = CDbl(Text1.Text)
b = CDbl(Text2.Text)

For i = 1 To 300

Text8.Text = i
qq = (a + b) / 2

R1 = a
Call abc
tempa = Ans
R1 = qq
Call abc
tempqq = Ans

If tempqq = 0 Then

 Text9.Text = qq
 Text10.Text = Ans
 Exit For
 
End If

 If (tempa * tempqq) > 0 Then
  
  a = qq

 Else

  b = qq

 End If


Next i
End Sub
Sub FindR3()
a2 = CDbl(Text20.Text)
b2 = CDbl(Text19.Text)

For i = 1 To 100

Text17.Text = i
qq2 = (a2 + b2) / 2

R3a = a2
Call def
tempa2 = Ans2
R3a = qq2
Call def
tempqq2 = Ans2

If tempqq2 = 0 Then

 Text16.Text = qq2
 Text15.Text = Ans2
 Exit For
 
End If

 If (tempa2 * tempqq2) > 0 Then
    a2 = qq2

 Else

    b2 = qq2

 End If


Next i
End Sub

Private Sub Command5_Click()

For i = 1 To 100

Text23.Text = i

Call FindR1
Text18.Text = CStr(R1)
Call FindR3
If R3a - CDbl(Text3.Text) = 0 Then
   
   Text21.Text = CStr(R1)
   Text22.Text = CStr(R3a)
   Text3.Text = CStr(R3a)
   Exit For
    
ElseIf Abs(R3a - CDbl(Text3.Text)) <= 1 Then
   
     Text21.Text = CStr(R1)
     Text22.Text = CStr(R3a)
     Text3.Text = CStr(R3a)
     Exit For
   
    Else
    
    Text3.Text = CStr(R3a)
    
End If

Next i


End Sub

Private Sub Form_Load()

End Sub
