VERSION 5.00
Begin VB.Form credits 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credits"
   ClientHeight    =   3930
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4785
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
   Icon            =   "credits.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2712.556
   ScaleMode       =   0  'User
   ScaleWidth      =   4493.362
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile: +917870400632                              +918083299702               Email: alamimran613@gmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   112.686
      X2              =   4394.761
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Development: Imran Alam          Graphics: Imran Alam         "
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5055
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
      Width           =   5775
   End
End
Attribute VB_Name = "credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
