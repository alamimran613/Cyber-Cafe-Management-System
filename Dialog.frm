VERSION 5.00
Begin VB.Form changepassword 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2130
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4965
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpassword 
      DataField       =   "pass"
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtserial 
      DataField       =   "sr"
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Enter Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   " Cyber Cafe Management System"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "changepassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Insert Into Fields", vbCritical
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Else
If Text1.Text = Text2.Text Then
DataEnvironment1.rsCommand3.MoveLast
temp = Val(txtserial.Text)
DataEnvironment1.rsCommand3.AddNew
txtserial.Text = ""
txtpassword.Text = ""
txtserial.Text = Val(temp + 1)
txtpassword.Text = Text2.Text
DataEnvironment1.rsCommand3.MoveFirst
DataEnvironment1.rsCommand3.MoveLast
MsgBox "Password Changed!!", vbInformation
DataEnvironment1.rsCommand3.MoveFirst
DataEnvironment1.rsCommand3.MoveLast
Unload Me
Else
MsgBox "Password Not Matched, Re-Enter Password!!", vbCritical
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If
End If

End Sub

Private Sub Form_Load()
DataEnvironment1.rsCommand3.MoveLast
Text1.Text = ""
Text2.Text = ""
txtserial.Enabled = False
txtpassword.Enabled = False
txtserial.Visible = False
txtpassword.Visible = False
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
