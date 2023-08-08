VERSION 5.00
Begin VB.Form Homepage 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cyber Cafe Management System"
   ClientHeight    =   9150
   ClientLeft      =   2760
   ClientTop       =   4050
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Homepage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   5640
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   0
   End
   Begin VB.CommandButton Command50 
      BackColor       =   &H0000FF00&
      Caption         =   "50"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H0000FF00&
      Caption         =   "49"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H0000FF00&
      Caption         =   "48"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command47 
      BackColor       =   &H0000FF00&
      Caption         =   "47"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H0000FF00&
      Caption         =   "46"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H0000FF00&
      Caption         =   "45"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H0000FF00&
      Caption         =   "44"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command43 
      BackColor       =   &H0000FF00&
      Caption         =   "43"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H0000FF00&
      Caption         =   "42"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H0000FF00&
      Caption         =   "41"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H0000FF00&
      Caption         =   "40"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H0000FF00&
      Caption         =   "39"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H0000FF00&
      Caption         =   "38"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H0000FF00&
      Caption         =   "37"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H0000FF00&
      Caption         =   "36"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H0000FF00&
      Caption         =   "35"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H0000FF00&
      Caption         =   "34"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H0000FF00&
      Caption         =   "33"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H0000FF00&
      Caption         =   "32"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H0000FF00&
      Caption         =   "31"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H0000FF00&
      Caption         =   "30"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H0000FF00&
      Caption         =   "29"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H0000FF00&
      Caption         =   "28"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H0000FF00&
      Caption         =   "27"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H0000FF00&
      Caption         =   "26"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H0000FF00&
      Caption         =   "25"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H0000FF00&
      Caption         =   "24"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H0000FF00&
      Caption         =   "23"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H0000FF00&
      Caption         =   "22"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H0000FF00&
      Caption         =   "21"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H0000FF00&
      Caption         =   "20"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H0000FF00&
      Caption         =   "19"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H0000FF00&
      Caption         =   "18"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0000FF00&
      Caption         =   "17"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0000FF00&
      Caption         =   "16"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000FF00&
      Caption         =   "15"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FF00&
      Caption         =   "14"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000FF00&
      Caption         =   "13"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0000FF00&
      Caption         =   "12"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0000FF00&
      Caption         =   "11"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0000FF00&
      Caption         =   "10"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0000FF00&
      Caption         =   "9"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "8"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
      Caption         =   "7"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "6"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "5"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "4"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "3"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "2"
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   " Cyber Cafe Management System"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   6015
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu menuentry 
         Caption         =   "&Records"
         Shortcut        =   ^R
      End
      Begin VB.Menu menuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menusetting 
      Caption         =   "&Setting"
      Begin VB.Menu menuchangepassword 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "&About"
      Begin VB.Menu menuabout 
         Caption         =   "&About!!"
      End
      Begin VB.Menu menucredits 
         Caption         =   "&Credits!!"
      End
      Begin VB.Menu menulicense 
         Caption         =   "&License!!"
      End
   End
End
Attribute VB_Name = "Homepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnCloseEnabled As Boolean
Dim ask As Integer

Private Sub Command1_Click()
system1.Show
End Sub

Private Sub Command10_Click()
system10.Show
End Sub

Private Sub Command11_Click()
system11.Show
End Sub

Private Sub Command12_Click()
system12.Show
End Sub

Private Sub Command13_Click()
system13.Show
End Sub

Private Sub Command14_Click()
system14.Show
End Sub

Private Sub Command15_Click()
system15.Show
End Sub

Private Sub Command16_Click()
system16.Show
End Sub

Private Sub Command17_Click()
system17.Show
End Sub

Private Sub Command18_Click()
system18.Show
End Sub

Private Sub Command19_Click()
system19.Show
End Sub

Private Sub command2_Click()
system2.Show
End Sub

Private Sub Command20_Click()
system20.Show
End Sub

Private Sub Command21_Click()
system21.Show
End Sub

Private Sub Command22_Click()
system22.Show
End Sub

Private Sub Command23_Click()
system23.Show
End Sub

Private Sub Command24_Click()
system24.Show
End Sub

Private Sub Command25_Click()
system25.Show
End Sub

Private Sub Command26_Click()
system26.Show
End Sub

Private Sub Command27_Click()
system27.Show
End Sub

Private Sub Command28_Click()
system28.Show
End Sub

Private Sub Command29_Click()
system29.Show
End Sub

Private Sub Command3_Click()
system3.Show
End Sub

Private Sub Command30_Click()
system30.Show
End Sub

Private Sub Command31_Click()
system31.Show
End Sub

Private Sub Command32_Click()
system32.Show
End Sub

Private Sub Command33_Click()
system33.Show
End Sub

Private Sub Command34_Click()
system34.Show
End Sub

Private Sub Command35_Click()
system35.Show
End Sub

Private Sub Command36_Click()
system36.Show
End Sub

Private Sub Command37_Click()
system37.Show
End Sub

Private Sub Command38_Click()
system38.Show
End Sub

Private Sub Command39_Click()
system39.Show
End Sub

Private Sub Command4_Click()
system4.Show
End Sub

Private Sub Command40_Click()
system40.Show
End Sub

Private Sub Command41_Click()
system41.Show
End Sub

Private Sub Command42_Click()
system42.Show
End Sub

Private Sub Command43_Click()
system43.Show
End Sub

Private Sub Command44_Click()
system44.Show
End Sub

Private Sub Command45_Click()
system45.Show
End Sub

Private Sub Command46_Click()
system46.Show
End Sub

Private Sub Command47_Click()
system47.Show
End Sub

Private Sub Command48_Click()
system48.Show
End Sub

Private Sub Command49_Click()
system49.Show
End Sub

Private Sub Command5_Click()
system5.Show
End Sub

Private Sub Command50_Click()
system50.Show
End Sub

Private Sub Command6_Click()
system6.Show
End Sub

Private Sub Command7_Click()
system7.Show
End Sub

Private Sub Command8_Click()
system8.Show
End Sub

Private Sub Command9_Click()
system9.Show
End Sub

Private Sub Form_Load()
Timer1.Interval = 1000
Timer2.Interval = 1000
EnableCloseButton Me.hWnd, False
End Sub

Private Sub menuabout_Click()
aboutsoftware.Show
End Sub

Private Sub menuchangepassword_Click()
password.Show
End Sub

Private Sub menucredits_Click()
credits.Show
End Sub

Private Sub menuentry_Click()
entry.Show
End Sub

Private Sub menuexit_Click()
ask = MsgBox("This will Close all Activated System Timer and Other information, Are you sure want to Exit?", vbOKCancel + vbQuestion)
If ask = vbOK Then
Unload system1
Unload system2
Unload system3
Unload system4
Unload system5
Unload system6
Unload system7
Unload system8
Unload system9
Unload system10
Unload system11
Unload system12
Unload system13
Unload system14
Unload system15
Unload system16
Unload system17
Unload system18
Unload system19
Unload system20
Unload system21
Unload system22
Unload system23
Unload system24
Unload system25
Unload system26
Unload system27
Unload system28
Unload system29
Unload system30
Unload system31
Unload system32
Unload system33
Unload system34
Unload system35
Unload system36
Unload system37
Unload system38
Unload system39
Unload system40
Unload system41
Unload system42
Unload system43
Unload system44
Unload system45
Unload system46
Unload system47
Unload system48
Unload system49
Unload system50
Unload aboutsoftware
Unload credits
Unload entry
Unload license
Unload changepassword
Unload password
Unload login
Unload loading
Unload Me
End If
End Sub

Private Sub menulicense_Click()
MsgBox "This is Trail Version", vbInformation
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = vbBlack
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = vbWhite
End Sub
