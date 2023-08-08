VERSION 5.00
Begin VB.Form license 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "License"
   ClientHeight    =   3585
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "license.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4680
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lifetime"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Software usage limit:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "IMRAN ALAM"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is licensed to:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Version 1.0.0"
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
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
Attribute VB_Name = "license"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
