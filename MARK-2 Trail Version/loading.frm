VERSION 5.00
Begin VB.Form loading 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2370
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "loading.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   960
      Picture         =   "loading.frx":000C
      Top             =   120
      Width           =   2040
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Interval = 3000
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = True
Timer1.Enabled = False
Homepage.Show
Unload Me
End Sub
