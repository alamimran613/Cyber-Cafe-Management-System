VERSION 5.00
Begin VB.Form system49 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System No. 49"
   ClientHeight    =   3150
   ClientLeft      =   255
   ClientTop       =   1740
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "system49.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Left            =   360
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   2520
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
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   375
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   2640
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Management of System No. 49"
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
      TabIndex        =   13
      Top             =   0
      Width           =   4695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6240
      Y1              =   360
      Y2              =   360
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
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
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
      Left            =   3120
      TabIndex        =   11
      Top             =   1080
      Width           =   135
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
      Left            =   2280
      TabIndex        =   10
      Top             =   1080
      Width           =   135
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
      Left            =   1680
      TabIndex        =   9
      Top             =   360
      Width           =   2535
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
      Left            =   3360
      TabIndex        =   8
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
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   735
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
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "system49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hh As Integer
Dim ss As Integer
Dim mm As Integer
Dim ans As Integer

Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer1.Interval = 1000
Timer2.Interval = 1000
Timer3.Interval = 1000
Homepage.Command49.BackColor = vbRed
Command1.Visible = False
Command2.Visible = True
End Sub

Private Sub command2_Click()
Label5.Visible = True
Label6.Visible = True
Timer2.Enabled = False
Timer3.Enabled = False
Timer1.Enabled = False
Command2.Visible = False
Command1.Visible = True
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Private Sub Command4_Click()
ans = MsgBox("Exit will stop Time, Are you sure want to Exit?", vbYesNo + vbQuestion)
If (ans = vbYes) Then
Unload Me
Homepage.Command49.BackColor = vbGreen
End If
End Sub

Private Sub Form_Load()
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
Homepage.Command49.BackColor = vbRed
End Sub

Private Sub Option2_Click()
Option2.Value = True
Option1.Value = False
Homepage.Command49.BackColor = vbRed
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
Homepage.Command49.BackColor = vbYellow
End If
End If
If (Option1.Value = True) Then
If (mm = 30) Then
Homepage.Command49.BackColor = vbYellow
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

