VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "共射集放大器-圖形模擬"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   10770
   Begin VB.CommandButton Command9 
      Caption         =   "啟動"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Text            =   "10"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Text            =   "5"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Text            =   "10"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Text            =   "10"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "3"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "60"
      Top             =   600
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1560
   End
   Begin VB.CommandButton Command10 
      Caption         =   "清除"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重新輸入"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Line Line32 
      BorderColor     =   &H00FF0000&
      X1              =   6960
      X2              =   6960
      Y1              =   5520
      Y2              =   5760
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00FF0000&
      X1              =   6360
      X2              =   6360
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00FF0000&
      X1              =   5760
      X2              =   5760
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line29 
      BorderColor     =   &H00FF0000&
      X1              =   5160
      X2              =   5160
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FF0000&
      X1              =   4560
      X2              =   4560
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00FF0000&
      X1              =   3960
      X2              =   3960
      Y1              =   5520
      Y2              =   5760
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FF0000&
      X1              =   3360
      X2              =   3360
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FF0000&
      X1              =   2760
      X2              =   2760
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FF0000&
      X1              =   2160
      X2              =   2160
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FF0000&
      X1              =   1560
      X2              =   1560
      Y1              =   5760
      Y2              =   5520
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF0000&
      X1              =   1080
      X2              =   840
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF0000&
      X1              =   1080
      X2              =   840
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      X1              =   840
      X2              =   1080
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      X1              =   1080
      X2              =   840
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   960
      X2              =   960
      Y1              =   4440
      Y2              =   6840
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   6960
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FF0000&
      X1              =   6960
      X2              =   6960
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FF0000&
      X1              =   6360
      X2              =   6360
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF0000&
      X1              =   5760
      X2              =   5760
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF0000&
      X1              =   5160
      X2              =   5160
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF0000&
      X1              =   4560
      X2              =   4560
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF0000&
      X1              =   3960
      X2              =   3960
      Y1              =   2640
      Y2              =   2880
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FF0000&
      X1              =   3360
      X2              =   3360
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FF0000&
      X1              =   2760
      X2              =   2760
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FF0000&
      X1              =   2160
      X2              =   2160
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF0000&
      X1              =   1560
      X2              =   1560
      Y1              =   2880
      Y2              =   2640
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      X1              =   1080
      X2              =   840
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      X1              =   1080
      X2              =   840
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      X1              =   840
      X2              =   1080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      X1              =   1080
      X2              =   840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      X1              =   840
      X2              =   6960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   960
      X2              =   960
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Label Label8 
      Caption         =   "ms"
      Height          =   255
      Left            =   9480
      TabIndex        =   34
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "ms"
      Height          =   255
      Left            =   9480
      TabIndex        =   33
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Time/\DIV"
      Height          =   255
      Left            =   8520
      TabIndex        =   32
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "V/DIV"
      Height          =   255
      Left            =   8520
      TabIndex        =   31
      Tag             =   "5"
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Time/\DIV"
      Height          =   255
      Left            =   8520
      TabIndex        =   30
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "V/DIV"
      Height          =   255
      Left            =   8520
      TabIndex        =   29
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "輸出"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "輸入"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label29 
      Caption         =   "正弦輸入"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "振幅"
      Height          =   255
      Left            =   1320
      TabIndex        =   25
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label31 
      Caption         =   "V"
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label32 
      Caption         =   "頻率"
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label33 
      Caption         =   "Hz"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label34 
      Caption         =   "輸出峰值"
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label35 
      Caption         =   "輸出峰谷"
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "暫態Ib"
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "暫態Ic"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "暫態輸出"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label14 
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "電壓增益"
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label18 
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label19 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Amp As Single
Public Hz As Integer
Dim t As Single
Dim i, j As Double
Dim v1, v2 As Double
Dim ff As Long
Dim VDiv1, VDiv2, TDiv1, TDiv2 As Single
Dim x, y As Integer
Dim mm, nn As Single
Dim n As Double
Dim temp1, temp2 As Single
Dim IIb, IIc As Double
Dim vv1, vv2 As Single
Dim VVC, VC2 As Double
Public max, min As Double
Dim AAv As Double
Dim Tempv2 As Double
Dim flag, flag2 As Integer



Private Sub Command1_Click()
Cls
Form2.Show
vv1 = 0
mm = 960
nn = 960
Tempv2 = 5
i = 0
j = 0
Call D1
Call D2
Command9.Caption = "啟動"
Timer1.Enabled = False
max = 0
min = 0
AAv = 0
flag = 0
flag2 = 0
End Sub

Private Sub Command10_Click()
Cls
vv1 = 0
mm = 960
nn = 960
Tempv2 = 5
i = 0
j = 0
Call D1
Call D2
Command9.Caption = "啟動"
Timer1.Enabled = False
max = 0
min = 0
AAv = 0
flag = 0
flag2 = 0
End Sub

Private Sub Command9_Click()

If Command9.Caption = "停止" Then
   Command9.Caption = "啟動"
   Timer1.Enabled = False
Else
   Command9.Caption = "停止"
   Timer1.Enabled = True
End If

 Call D1
 Call D2



End Sub

Sub keyin()

Amp = CSng(Text8.Text)
Hz = CInt(Text9.Text)
VDiv1 = CSng(Text3.Text)
VDiv2 = CSng(Text1.Text)
TDiv1 = CSng(Text4.Text)
TDiv2 = CSng(Text2.Text)
t = 1 / Hz
End Sub


Sub Draw()


End Sub

Private Sub Form_Load()
i = 0
j = 0
n = 4 * Atn(1#)
mm = 960
nn = 960
max = 0
min = 0
Tempv2 = 5
flag = 0
flag2 = 0
End Sub

Sub D1()
For x = 0 To 50

  ff = 960 + x * 120
  
  For y = 0 To 20
      
   Line (ff, 1560 + y * 120)-(ff + 20, 1560 + y * 120), QBColor(3)
  
  Next y

Next x
End Sub
Sub D2()

For x = 0 To 50

  ff = 960 + x * 120
  
  For y = 0 To 20
      
   Line (ff, 4440 + y * 120)-(ff + 20, 4440 + y * 120), QBColor(3)
  
  Next y

Next x

End Sub

Private Sub Timer1_Timer()
Call keyin

Call aa1

Call aa2

End Sub

Sub aa1()

temp1 = ((t * 1000) / TDiv1) * 600

If mm < 6960 Then

  If i = 360 Then
   
   flag2 = 1
   
   i = 0
    
  Else
       i = i + 0.5
   
       vv1 = (600 / VDiv1) * Amp * Sin(i * n / 180)
 
       mm = mm + CInt(temp1 / 360) * 0.5
 
       Line (mm, 2760 - vv1)-(mm + 2, 2760 - vv1), QBColor(5)
  
  End If
 
Else
 
 Timer1.Enabled = False
 
End If

End Sub
Sub aa2()



v1 = Amp * Sin(i * n / 180)

temp2 = ((t * 1000) / TDiv2) * 600

If (v1 + Form2.VB) > Form2.VBE Then

 IIb = (v1 + Form2.VB - Form2.VBE) / ((Form2.B + 1) * Form2.R4)
 IIc = Form2.B * IIb
 VC2 = Form2.VCC - IIc * Form2.R3
 
    If (v1 + Form2.VB) - VC2 < Form2.VBC And VC2 < Form2.VCC Then
         
       v2 = VC2 - Form2.VC
 
    End If
  
  If v2 > max Then

   max = v2

  End If
  
  If v2 < min Then

   min = v2

  End If
  
  
End If

If flag2 = 0 Then
 
 If Tempv2 <> v2 Then
  
  Tempv2 = v2

 Else

  flag = flag + 1

 End If

End If
If flag > 1 Then

 Label19.Caption = "失真!!!"

Else

 Label19.Caption = "無失真"

End If



If nn < 6960 Then

  If j = 360 Then

   j = 0

  Else

   j = j + 0.5
   
   vv2 = (600 / VDiv2) * v2
   
   nn = nn + CInt(temp2 / 360) * 0.5
   
   Line (nn, 5640 - vv2)-(nn + 2, 5640 - vv2), QBColor(5)
 
  End If
 
End If

AAv = (max - min) / (Amp * 2)

Label11.Caption = IIb
Label12.Caption = IIc
Label14.Caption = v2
Label15.Caption = max
Label16.Caption = min
Label18.Caption = AAv
End Sub

