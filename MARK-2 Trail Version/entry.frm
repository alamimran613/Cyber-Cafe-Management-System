VERSION 5.00
Begin VB.Form entry 
   BackColor       =   &H00E3FDFD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Entry, Searching & Cost Management"
   ClientHeight    =   9120
   ClientLeft      =   2760
   ClientTop       =   4050
   ClientWidth     =   14220
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "entry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "entry.frx":030A
   ScaleHeight     =   9120
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttempusage 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   56
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox txttemplogout 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   55
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox txttemplogin 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   54
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txttempmobile 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   53
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txttempaddress 
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1800
      TabIndex        =   52
      Top             =   4080
      Width           =   5295
   End
   Begin VB.TextBox txttempuser 
      Height          =   420
      Left            =   1800
      MaxLength       =   48
      TabIndex        =   51
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox txttempsystem 
      Height          =   420
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   50
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txttempdate 
      Height          =   420
      Left            =   1800
      TabIndex        =   49
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txttempserial 
      Height          =   420
      Left            =   1800
      TabIndex        =   48
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Upda&te"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtaccount 
      Height          =   420
      Left            =   9960
      TabIndex        =   46
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000FF00&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FF00&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E3FDFD&
      Caption         =   "Controls"
      Height          =   1455
      Left            =   4680
      TabIndex        =   34
      Top             =   7560
      Width           =   5295
      Begin VB.CommandButton cmdMoveLast 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Last>>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdMovePrevious 
         BackColor       =   &H00FF80FF&
         Caption         =   "<&Prev"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdMoveNext 
         BackColor       =   &H00FF80FF&
         Caption         =   "Ne&xt>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdMoveFirst 
         BackColor       =   &H00FFFF00&
         Caption         =   "<<Fir&st"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000000FF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H000080FF&
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Search"
      Height          =   375
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Search"
      Height          =   375
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtsearchuser 
      Height          =   420
      Left            =   9600
      MaxLength       =   22
      TabIndex        =   31
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtsearchno 
      Height          =   420
      Left            =   9600
      TabIndex        =   30
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      Height          =   375
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      Height          =   375
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txttotal 
      DataField       =   "total_collection"
      DataMember      =   "Command2"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   9960
      TabIndex        =   24
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txttoday 
      DataField       =   "today_collection"
      DataMember      =   "Command2"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   9960
      TabIndex        =   23
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txtusage 
      DataField       =   "usage_time"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      TabIndex        =   19
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox txtlogout 
      DataField       =   "logout_time"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox txtlogin 
      DataField       =   "login_time"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      TabIndex        =   17
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtmobile 
      BackColor       =   &H00FFFFFF&
      DataField       =   "mobile_no"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   16
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtaddress 
      BackColor       =   &H00FFFFFF&
      DataField       =   "address"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      TabIndex        =   15
      Top             =   4080
      Width           =   5295
   End
   Begin VB.TextBox txtuser 
      DataField       =   "user_name"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      MaxLength       =   48
      TabIndex        =   14
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox txtsystem 
      DataField       =   "system_no"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtdate 
      DataField       =   "date"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      TabIndex        =   12
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtserial 
      DataField       =   "serial_no"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   420
      Left            =   1800
      TabIndex        =   11
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   11880
      Picture         =   "entry.frx":2286
      Top             =   4800
      Width           =   225
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Cost:"
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
      Left            =   7680
      TabIndex        =   45
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   11880
      Picture         =   "entry.frx":4202
      Top             =   6240
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   11880
      Picture         =   "entry.frx":617E
      Top             =   5520
      Width           =   225
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   29
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   28
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   9840
      TabIndex        =   27
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Collection:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   22
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Today Collection:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   21
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   9840
      TabIndex        =   20
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   7680
      X2              =   13560
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   7440
      X2              =   7440
      Y1              =   1080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   480
      X2              =   13560
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Usage Time:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Logout Time:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Time:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "System No."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Details Entry"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "              Cyber Cafe Management System"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14295
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu menuprint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ansdel As Integer
Dim ansclear As Integer
Dim ansclear2 As Integer
Dim ansexit As Integer
Dim tempserial As Integer
Dim temptoday As Integer
Dim temptotal As Integer
Dim temp As Integer



Private Sub cmdDelete_Click()
If txtserial.Text = "" Then
MsgBox "No Records found on the screen for delete!!", vbCritical
DataEnvironment1.rsCommand1.MoveLast
Else
If (txtserial.Text = 1) Then
MsgBox "You don't have Right to Delete this Record, This is system reserved Record!!", vbCritical
DataEnvironment1.rsCommand1.MoveLast
Else

ansdel = MsgBox("Do you want to Delete?", vbYesNo + vbQuestion)
If ansdel = vbYes Then
DataEnvironment1.rsCommand1.Delete
DataEnvironment1.rsCommand1.MoveFirst

While Not DataEnvironment1.rsCommand1.EOF

temp = Val(txtserial.Text)
DataEnvironment1.rsCommand1.MoveNext
txtserial.Text = Val(temp + 1)
Wend
DataEnvironment1.rsCommand1.MoveFirst
DataEnvironment1.rsCommand1.MoveLast
MsgBox "Record Deleted!!", vbInformation
End If
End If
End If
End Sub

Private Sub cmdExit_Click()
ansexit = MsgBox("Do you want to Exit?", vbYesNo + vbQuestion)
If (ansexit = vbYes) Then
Unload Me
End If
End Sub

Private Sub cmdMoveFirst_Click()
DataEnvironment1.rsCommand1.MoveFirst
DataEnvironment1.rsCommand1.MoveNext
End Sub

Private Sub cmdMoveLast_Click()
DataEnvironment1.rsCommand1.MoveLast
End Sub

Private Sub cmdMoveNext_Click()
If DataEnvironment1.rsCommand1.EOF Then
DataEnvironment1.rsCommand1.MoveFirst
Else
DataEnvironment1.rsCommand1.MoveNext
End If
End Sub

Private Sub cmdMovePrevious_Click()
If DataEnvironment1.rsCommand1.BOF Then
DataEnvironment1.rsCommand1.MoveLast
Else
DataEnvironment1.rsCommand1.MovePrevious
End If
End Sub

Public Sub cmdNew_Click()
menuprint.Enabled = False
txtsearchno.Enabled = False
txtsearchuser.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command13.Enabled = False
txttempserial.Visible = True
txttempdate.Visible = True
txttempsystem.Visible = True
txttempuser.Visible = True
txttempaddress.Visible = True
txttempmobile.Visible = True
txttemplogin.Visible = True
txttemplogout.Visible = True
txttempusage.Visible = True

txtserial.Visible = False
txtdate.Visible = False
txtsystem.Visible = False
txtuser.Visible = False
txtaddress.Visible = False
txtmobile.Visible = False
txtlogin.Visible = False
txtlogout.Visible = False
txtusage.Visible = False

cmdUpdate.Enabled = False
cmdDelete.Enabled = False
cmdExit.Enabled = False
cmdMoveFirst.Enabled = False
cmdMoveNext.Enabled = False
cmdMovePrevious.Enabled = False
cmdMoveLast.Enabled = False

DataEnvironment1.rsCommand1.MoveLast
tempserial = txtserial.Text
txttempserial.Text = tempserial + 1
txttempdate.Text = Date
txttempsystem.SetFocus
cmdNew.Visible = False
cmdSave.Visible = True

End Sub

Private Sub cmdSave_Click()
If txttempsystem.Text = "" Then
MsgBox "Enter System No.", vbCritical
txttempsystem.SetFocus
Else
If txttempuser.Text = "" Then
MsgBox "Enter User Name", vbCritical
txttempuser.SetFocus
Else
If txttempaddress.Text = "" Then
MsgBox "Enter Address", vbCritical
txttempaddress.SetFocus
Else
If txttempmobile.Text = "" Then
MsgBox "Enter Mobile No.", vbCritical
txttempmobile.SetFocus
Else
If txttemplogin.Text = "" Then
MsgBox "Enter Login Time", vbCritical
txttemplogin.SetFocus
Else
If txttemplogout.Text = "" Then
MsgBox "Enter Logout Time", vbCritical
txttemplogout.SetFocus
Else
If txttempusage.Text = "" Then
MsgBox "Enter Usage Time", vbCritical
txttempusage.SetFocus
Else
txttempuser.Text = StrConv(txttempuser.Text, vbUpperCase)
txttempaddress.Text = StrConv(txttempaddress.Text, vbUpperCase)

DataEnvironment1.rsCommand1.AddNew
txtserial.Text = txttempserial.Text
txtdate.Text = txttempdate.Text
txtsystem.Text = txttempsystem.Text
txtuser.Text = txttempuser.Text
txtaddress.Text = txttempaddress.Text
txtmobile.Text = txttempmobile.Text
txtlogin.Text = txttemplogin.Text
txtlogout.Text = txttemplogout.Text
txtusage.Text = txttempusage.Text

txttempserial.Visible = False
txttempdate.Visible = False
txttempsystem.Visible = False
txttempuser.Visible = False
txttempaddress.Visible = False
txttempmobile.Visible = False
txttemplogin.Visible = False
txttemplogout.Visible = False
txttempusage.Visible = False

txtserial.Visible = True
txtdate.Visible = True
txtsystem.Visible = True
txtuser.Visible = True
txtaddress.Visible = True
txtmobile.Visible = True
txtlogin.Visible = True
txtlogout.Visible = True
txtusage.Visible = True

cmdUpdate.Enabled = True
cmdDelete.Enabled = True
cmdExit.Enabled = True
cmdMoveFirst.Enabled = True
cmdMoveNext.Enabled = True
cmdMovePrevious.Enabled = True
cmdMoveLast.Enabled = True

DataEnvironment1.rsCommand1.MoveFirst
DataEnvironment1.rsCommand1.MoveLast
MsgBox "Record Saved!!", vbInformation
cmdSave.Visible = False
cmdNew.Visible = True
txtaccount.SetFocus
txtsearchno.Enabled = True
txtsearchuser.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command13.Enabled = True
menuprint.Enabled = True
End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub cmdUpdate_Click()
txtuser.Text = StrConv(txtuser.Text, vbUpperCase)
txtaddress.Text = StrConv(txtaddress.Text, vbUpperCase)
MsgBox "Record Updated!!", vbInformation
End Sub

Private Sub Command1_Click()
If txttoday.Text = "0" Then
MsgBox "Today Collection already Cleared!!", vbCritical
Else
ansclear = MsgBox("Do you want to Clear?", vbYesNo + vbQuestion)
If (ansclear = vbYes) Then
txttoday.Text = 0
DataEnvironment1.rsCommand2.MoveFirst
DataEnvironment1.rsCommand2.MoveLast
MsgBox "Today Collection Cleared!!", vbInformation
End If
End If
End Sub

Private Sub Command13_Click()
If txtsearchuser.Text = "" Then
MsgBox "Please Enter User Name!!", vbCritical
txtsearchuser.SetFocus
Else
DataEnvironment1.rsCommand1.MoveNext
While Not txtuser.Text = txtsearchuser.Text
If (DataEnvironment1.rsCommand1.EOF) Then
searchresult2 = MsgBox("No Record Found!!", vbOKOnly + vbInformation)
DataEnvironment1.rsCommand1.MoveLast
Exit Sub
Else
DataEnvironment1.rsCommand1.MoveNext
End If
Wend
MsgBox "Record Found!!", vbInformation
End If
End Sub

Private Sub command2_Click()
If txttotal.Text = "0" Then
MsgBox "Total Collection already Cleared!!", vbCritical
Else
ansclear2 = MsgBox("Do you want to Clear?", vbYesNo + vbQuestion)
If (ansclear2 = vbYes) Then
txttotal.Text = 0
DataEnvironment1.rsCommand2.MoveFirst
DataEnvironment1.rsCommand2.MoveLast
MsgBox "Total Collection Cleared!!", vbInformation
End If
End If
End Sub

Private Sub Command3_Click()
If txtsearchno.Text = "" Then
MsgBox "Please Enter Serial No.", vbCritical
txtsearchno.SetFocus
Else
DataEnvironment1.rsCommand1.MoveFirst
While Not txtserial.Text = txtsearchno.Text
If (DataEnvironment1.rsCommand1.EOF) Then
searchresult = MsgBox("No Record Found!!", vbOKOnly + vbInformation)
DataEnvironment1.rsCommand1.MoveLast
Exit Sub
Else
DataEnvironment1.rsCommand1.MoveNext
End If
Wend
MsgBox "Record Found!!", vbInformation
End If
End Sub

Private Sub Command4_Click()
If txtsearchuser.Text = "" Then
MsgBox "Please Enter User Name!!", vbCritical
txtsearchuser.SetFocus
Else
DataEnvironment1.rsCommand1.MoveFirst
txtsearchuser.Text = StrConv(txtsearchuser.Text, vbUpperCase)
While Not txtuser.Text = txtsearchuser.Text
If (DataEnvironment1.rsCommand1.EOF) Then
searchresult1 = MsgBox("No Record Found!!", vbOKOnly + vbInformation)
DataEnvironment1.rsCommand1.MoveLast
Exit Sub
Else
DataEnvironment1.rsCommand1.MoveNext
End If
Wend
MsgBox "Record Found!!", vbInformation
End If
End Sub

Private Sub Command5_Click()
If (txtaccount.Text = "") Then
MsgBox "Please Enter Cost", vbCritical
txtaccount.SetFocus
Else
temptoday = Val(txttoday.Text)
temptotal = Val(txttotal.Text)
DataEnvironment1.rsCommand2.AddNew
txttoday.Text = ""
txttotal.Text = ""
txttoday.Text = Val(temptoday) + Val(txtaccount)
txttotal.Text = Val(temptotal) + Val(txtaccount)
DataEnvironment1.rsCommand2.MoveFirst
DataEnvironment1.rsCommand2.MoveLast
txtaccount.Text = ""
MsgBox "Accounts Updated!!", vbInformation
End If
End Sub

Private Sub Form_Load()
DataEnvironment1.rsCommand1.MoveLast
DataEnvironment1.rsCommand2.MoveLast
cmdSave.Visible = False
cmdNew.Visible = True
txttempserial.Visible = False
txttempdate.Visible = False
txttempsystem.Visible = False
txttempuser.Visible = False
txttempaddress.Visible = False
txttempmobile.Visible = False
txttemplogin.Visible = False
txttemplogout.Visible = False
txttempusage.Visible = False

txttoday.Enabled = False
txttotal.Enabled = False
txtserial.Enabled = False
txttempserial.Enabled = False
txtdate.Enabled = False
txttempdate.Enabled = False


End Sub

Private Sub menuprint_Click()
DataReport1.Show
Unload Me
End Sub

Private Sub txtaccount_GotFocus()
txtaccount.BackColor = vbYellow
End Sub

Private Sub txtaccount_LostFocus()
txtaccount.BackColor = vbWhite
End Sub


Private Sub txtaddress_GotFocus()
txtaddress.BackColor = vbYellow
End Sub

Private Sub txtaddress_LostFocus()
txtaddress.BackColor = vbWhite
End Sub



Private Sub txtdate_GotFocus()
txtdate.BackColor = vbYellow
End Sub

Private Sub txtdate_LostFocus()
txtdate.BackColor = vbWhite
End Sub



Private Sub txtlogin_GotFocus()
txtlogin.BackColor = vbYellow
End Sub

Private Sub txtlogin_LostFocus()
txtlogin.BackColor = vbWhite
End Sub



Private Sub txtlogout_GotFocus()
txtlogout.BackColor = vbYellow
End Sub

Private Sub txtlogout_LostFocus()
txtlogout.BackColor = vbWhite
End Sub



Private Sub txtmobile_GotFocus()
txtmobile.BackColor = vbYellow
End Sub

Private Sub txtmobile_LostFocus()
txtmobile.BackColor = vbWhite
End Sub


Private Sub txtsearchno_GotFocus()
txtsearchno.BackColor = vbYellow
End Sub

Private Sub txtsearchno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3_Click
End If
End Sub

Private Sub txtsearchno_LostFocus()
txtsearchno.BackColor = vbWhite
End Sub



Private Sub txtsearchuser_GotFocus()
txtsearchuser.BackColor = vbYellow
End Sub

Private Sub txtsearchuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command4_Click
End If
End Sub

Private Sub txtsearchuser_LostFocus()
txtsearchuser.BackColor = vbWhite
End Sub



Private Sub txtserial_GotFocus()
txtserial.BackColor = vbYellow
End Sub

Private Sub txtserial_LostFocus()
txtserial.BackColor = vbWhite
End Sub



Private Sub txtsystem_GotFocus()
txtsystem.BackColor = vbYellow
End Sub

Private Sub txtsystem_LostFocus()
txtsystem.BackColor = vbWhite
End Sub



Private Sub txttempaddress_GotFocus()
txttempaddress.BackColor = vbYellow
End Sub

Private Sub txttempaddress_LostFocus()
txttempaddress.BackColor = vbWhite
End Sub



Private Sub txttempdate_GotFocus()
txttempdate.BackColor = vbYellow
End Sub

Private Sub txttempdate_LostFocus()
txttempdate.BackColor = vbWhite
End Sub



Private Sub txttemplogin_GotFocus()
txttemplogin.BackColor = vbYellow
End Sub

Private Sub txttemplogin_LostFocus()
txttemplogin.BackColor = vbWhite
End Sub



Private Sub txttemplogout_GotFocus()
txttemplogout.BackColor = vbYellow
End Sub

Private Sub txttemplogout_LostFocus()
txttemplogout.BackColor = vbWhite
End Sub



Private Sub txttempmobile_GotFocus()
txttempmobile.BackColor = vbYellow
End Sub

Private Sub txttempmobile_LostFocus()
txttempmobile.BackColor = vbWhite
End Sub



Private Sub txttempserial_GotFocus()
txttempserial.BackColor = vbYellow
End Sub

Private Sub txttempserial_LostFocus()
txttempserial.BackColor = vbWhite
End Sub



Private Sub txttempsystem_GotFocus()
txttempsystem.BackColor = vbYellow
End Sub

Private Sub txttempsystem_LostFocus()
txttempsystem.BackColor = vbWhite
End Sub



Private Sub txttempusage_GotFocus()
txttempusage.BackColor = vbYellow
End Sub

Private Sub txttempusage_LostFocus()
txttempusage.BackColor = vbWhite
End Sub



Private Sub txttempuser_GotFocus()
txttempuser.BackColor = vbYellow
End Sub

Private Sub txttempuser_LostFocus()
txttempuser.BackColor = vbWhite
End Sub



Private Sub txttoday_GotFocus()
txttoday.BackColor = vbYellow
End Sub

Private Sub txttoday_LostFocus()
txttoday.BackColor = vbWhite
End Sub



Private Sub txttotal_GotFocus()
txttotal.BackColor = vbYellow
End Sub

Private Sub txttotal_LostFocus()
txttotal.BackColor = vbWhite
End Sub



Private Sub txtusage_GotFocus()
txtusage.BackColor = vbYellow
End Sub

Private Sub txtusage_LostFocus()
txtusage.BackColor = vbWhite
End Sub



Private Sub txtuser_GotFocus()
txtuser.BackColor = vbYellow
End Sub

Private Sub txtuser_LostFocus()
txtuser.BackColor = vbWhite
End Sub
