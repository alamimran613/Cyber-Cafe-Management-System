VERSION 5.00
Begin VB.Form aboutsoftware 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cyber Cafe Management System"
   ClientHeight    =   4800
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5625
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "aboutsoftware.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3313.046
   ScaleMode       =   0  'User
   ScaleWidth      =   5282.166
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   """Cyber Cafe Management System"" lets you manage Time management of systems and save user details for future use..."
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"aboutsoftware.frx":030A
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   112.686
      X2              =   5183.565
      Y1              =   1490.871
      Y2              =   1490.871
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "    Cyber Cafe Management System"
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
Attribute VB_Name = "aboutsoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
