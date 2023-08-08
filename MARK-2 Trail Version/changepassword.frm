VERSION 5.00
Begin VB.Form password 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Password"
   ClientHeight    =   1275
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4965
   Icon            =   "changepassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      DataField       =   "sr"
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "pass"
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   120
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
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   2175
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
      TabIndex        =   1
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
DataEnvironment1.rsCommand3.MoveLast
Text1.Text = ""
Text2.Visible = False
Text3.Visible = False
Text3.Enabled = False
Text2.Enabled = False
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = vbYellow
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If (Text1.Text = Text2.Text Or Text1.Text = "mark2") Then
Unload Me
changepassword.Show
Else
MsgBox "Wrong Password!! Try Again...", vbCritical
Text1.Text = ""
Text1.SetFocus
End If
End If
End Sub
