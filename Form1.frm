VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "A Mind Of Its Own (A SendKeys Tutorial)"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   840
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Text            =   "What you saw are samples of what you can do with the SendKeys Statement. Cool!!!"
      Top             =   120
      Width           =   3135
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   615
      ForeColor       =   8421376
      VariousPropertyBits=   268435483
      PicturePosition =   262148
      Size            =   "1085;1085"
      Picture         =   "Form1.frx":0000
      TakeFocusOnClick=   0   'False
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton3 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   615
      ForeColor       =   8388608
      VariousPropertyBits=   268435483
      PicturePosition =   262148
      Size            =   "1085;1085"
      Picture         =   "Form1.frx":0425
      TakeFocusOnClick=   0   'False
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
      ForeColor       =   12632064
      VariousPropertyBits=   268435483
      PicturePosition =   262148
      Size            =   "1085;1085"
      Picture         =   "Form1.frx":084D
      TakeFocusOnClick=   0   'False
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton4 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   615
      ForeColor       =   8421376
      VariousPropertyBits=   268435483
      PicturePosition =   262148
      Size            =   "1085;1085"
      Picture         =   "Form1.frx":0D4B
      TakeFocusOnClick=   0   'False
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, z

Private Sub CommandButton1_Click()
CommandButton1.Visible = False
n = Shell("notepad.exe", vbNormalFocus) 'Runs Notepad
Timer1 = True
End Sub

Private Sub CommandButton2_Click()
CommandButton2.Visible = False
c = Shell("calc.exe", vbNormalFocus) 'Runs Calculator
Timer3 = True
End Sub

Private Sub CommandButton3_Click()
CommandButton3.Visible = False
s = Shell("sol.exe", vbNormalFocus) 'Runs Solitaire
Timer2 = True
End Sub

Private Sub CommandButton4_Click()
End 'Exit
End Sub

Private Sub Form_Load()
m = MsgBox("When any of the animation is on going, don't close, move or lose focus on the applications.", vbInformation, "A Mind Of It's Own")
Form1.Height = 615
Form1.Width = 615
End Sub

Private Sub Timer1_Timer()
x = x + 1
SendKeys Mid(Text1, x, 1)

If x = Len(Text1) Then
    Timer1 = False
End If
End Sub

Private Sub Timer2_Timer()
y = y + 1
SendKeys " " 'Space
If y = 24 Then
    Timer2 = False
End If
End Sub

Private Sub Timer3_Timer()
z = z + 1
SendKeys z
If z = 9 Then
    Timer3 = False
End If
End Sub
