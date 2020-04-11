VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form1 
   Caption         =   "实验一 学号：2017301000143 姓名：项一飞"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6480
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text3 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   720
      Width           =   6735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "之间的所有素数"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "和"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "计算"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As Integer, n As Integer, r As Integer, i As Integer
Dim tmp As Integer
Dim count1 As Integer
Dim cnt As Integer
Dim str1 As String



Private Sub Form_Load()
Timer1.Interval = 50
Command1.Caption = "开始"
cnt = 0
str1 = "开始"
End Sub

Private Sub Command1_Click()

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "输入为空,重新输入！"
Exit Sub
End If

If cnt = 0 Then
m = 0
n = 0
count1 = 1
Text3.Text = ""
n = Val(Text1.Text)
m = Val(Text2.Text)
If n > m Then
tmp = n
n = m
m = tmp
End If

ProgressBar1.Max = m - n + 1

End If

If cnt > 0 Then
str1 = "继续"
End If





Timer1.Enabled = Command1.Caption = str1

Command1.Caption = IIf(Command1.Caption = "继续", "暂停", "继续")


If cnt = 0 Then
Command1.Caption = "暂停"
End If
cnt = cnt + 1

End Sub









Private Sub Timer1_Timer()

If n <= 2 Then
Text3.Text = Text3.Text + Str$(2)
n = 3
count1 = 2
ProgressBar1.Value = count1
End If

For i = 2 To n - 1
    If n Mod i = 0 Then Exit For
Next i

If i >= n Then
 Text3.Text = Text3.Text + " " + Str$(n)
End If
count1 = count1 + 1
ProgressBar1.Value = count1
n = n + 1

If n = m Then
ProgressBar1.Value = ProgressBar1.Max
MsgBox "完成计算"
Unload form1
Timer1.Enabled = False
End If
End Sub
Private Sub text1_keypress(keyascii As Integer)
    If keyascii = 8 Then Exit Sub
    If keyascii < 48 Or keyascii > 57 Then
    keyascii = 0
    MsgBox "不是数字", vbExclamation, "输入错误提示框"
    
    End If
    
End Sub


Private Sub text2_keypress(keyascii As Integer)
    If keyascii = 8 Then Exit Sub
    If keyascii < 48 Or keyascii > 57 Then
    keyascii = 0
    MsgBox "不是数字", vbExclamation, "输入错误提示框"
    
    End If
    
End Sub
