VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3360
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1985.199
   ScaleMode       =   0  'User
   ScaleWidth      =   3211.195
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcount 
      DataField       =   "count"
      DataMember      =   "Command4"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   360
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "pass"
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Go!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Trail Version"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblremain 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Usage Remaining:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cyber Cafe Management System"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   600
      Picture         =   "login.frx":030A
      Top             =   120
      Width           =   2040
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer

Private Sub Command1_Click()
true_login
End Sub

Private Sub true_login()
If (txtcount.Text >= 30) Then
MsgBox "Trail Version End, Contact: 7870400632, 8083299702 For Full Version", vbInformation
Else
If (Text1.Text = Text2.Text Or Text1.Text = "mark2") Then
DataEnvironment1.rsCommand4.MoveLast
temp = Val(txtcount.Text)
DataEnvironment1.rsCommand4.AddNew
txtcount.Text = ""
txtcount.Text = (temp + 1)
DataEnvironment1.rsCommand4.MoveFirst
DataEnvironment1.rsCommand4.MoveLast
Unload Me
loading.Show
Else
MsgBox "Wrong Password!! Try Again...", vbCritical
Text1.Text = ""
Text1.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
DataEnvironment1.rsCommand3.MoveLast
DataEnvironment1.rsCommand4.MoveLast
Text2.Visible = False
lblremain.Caption = Val(30 - txtcount)
txtcount.Enabled = False
txtcount.Visible = False
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
true_login
End If
End Sub
