VERSION 5.00
Begin VB.Form system27 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System No. 27"
   ClientHeight    =   3255
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   8805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "system27.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E3FDFD&
      Caption         =   "30 Min"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E3FDFD&
      Caption         =   "1 Hour"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Left            =   120
      Top             =   840
   End
   Begin VB.CommandButton cmdentry 
      BackColor       =   &H0000FFFF&
      Caption         =   "Record Entry"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Timer Timer4 
      Left            =   7680
      Top             =   600
   End
   Begin VB.Timer Timer5 
      Left            =   8040
      Top             =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1320
      TabIndex        =   23
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2160
      TabIndex        =   22
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3000
      TabIndex        =   21
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Elapse Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1320
      TabIndex        =   20
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   1920
      TabIndex        =   19
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   2760
      TabIndex        =   18
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Limit:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8880
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Management of System No. 27"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   0
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   600
      Y2              =   3120
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Detail"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   5640
      TabIndex        =   15
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Time:"
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
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Logout Time:"
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
      Left            =   4920
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lbllogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lbllogout 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Usage Time:"
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
      Left            =   4920
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblusage 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EE1A39&
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblsession 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Session Expired"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label lblclick 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here >>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "system27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hh As Integer
Dim ss As Integer
Dim mm As Integer
Dim ans As Integer
Dim asn2 As Integer

Private Sub cmdentry_Click()
If entry.cmdDelete.Enabled = False Then
MsgBox "Already Open, Save your Record First!!", vbCritical
Else
Timer4.Enabled = False
Timer5.Enabled = False
entry.Show
entry.cmdNew_Click
entry.txttemplogin.Text = lbllogin
entry.txttemplogout.Text = lbllogout
entry.txttempusage.Text = lblusage
entry.txttempsystem.Text = "27"
entry.txttempuser.SetFocus
End If
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer1.Interval = 1000
Timer2.Interval = 1000
Timer3.Interval = 1000
Homepage.command27.BackColor = vbRed
Command1.Visible = False
Command2.Visible = True
lbllogin.Caption = Time
End Sub

Private Sub command2_Click()
ans2 = MsgBox("Are you sure want to Stop?", vbYesNo + vbQuestion)
If ans2 = vbYes Then
lblclick.Visible = True
Label5.Visible = True
Label6.Visible = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer1.Enabled = False
Command2.Visible = False
Command1.Visible = True
lbllogout.Caption = Time
lblusage.Caption = Label1 + ":" + Label2 + " Hour"
Command1.Visible = False
Command2.Visible = False
lblsession.Visible = True
cmdentry.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer4.Interval = 300
Timer5.Interval = 300
End If
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Command4_Click()
ans = MsgBox("Exit will stop Time, Are you sure want to Exit?", vbYesNo + vbQuestion)
If (ans = vbYes) Then
Unload Me
Homepage.command27.BackColor = vbGreen
End If
End Sub

Private Sub Form_Load()
lblclick.Visible = False
cmdentry.Enabled = False
lblsession.Visible = False
Label1.Caption = "H"
Label2.Caption = "M"
Label3.Caption = "S"
Command4.BackColor = vbRed
Option2.Value = True
Command1.BackColor = vbGreen
Command2.BackColor = vbRed
Command1.Visible = True
Command2.Visible = False
ss = 0
mm = 0
hh = 0
End Sub

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Homepage.command27.BackColor = vbRed
End Sub

Private Sub Option2_Click()
Option2.Value = True
Option1.Value = False
Homepage.command27.BackColor = vbRed
End Sub

Private Sub Timer1_Timer()
ss = ss + 1
If (ss = 60) Then
mm = mm + 1
ss = 0
End If
If (mm = 60) Then
If (Option2.Value = True) Then
hh = hh + 1
mm = 0
Homepage.command27.BackColor = vbYellow
End If
End If
If (Option1.Value = True) Then
If (mm = 30) Then
Homepage.command27.BackColor = vbYellow
End If
If (mm = 60) Then
hh = hh + 1
mm = 0
End If
End If
Label1.Caption = hh
Label2.Caption = mm
Label3.Caption = ss
End Sub

Private Sub Timer2_Timer()
Label4.ForeColor = vbRed
Label5.Visible = True
Label6.Visible = True
End Sub

Private Sub Timer3_Timer()
Label4.ForeColor = vbBlue
Label5.Visible = False
Label6.Visible = False
End Sub

Private Sub Timer4_Timer()
cmdentry.BackColor = vbGreen
lblclick.ForeColor = vbBlue
End Sub

Private Sub Timer5_Timer()
cmdentry.BackColor = vbYellow
lblclick.ForeColor = vbRed
End Sub

