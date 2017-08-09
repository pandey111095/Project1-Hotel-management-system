VERSION 5.00
Begin VB.Form FRMSPOR 
   Caption         =   "SPORTS FORM"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "FRMSPOR.frx":0000
   ScaleHeight     =   9555
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text12 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   27
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   25
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   5880
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   6480
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   7080
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   7680
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10680
      TabIndex        =   4
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9480
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9480
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "&UPDATE"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9480
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "&MAIN MENU"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "RECEPT_NO"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7800
      TabIndex        =   28
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL_AMNT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7800
      TabIndex        =   26
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "A_DATE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   24
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "F_NAME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   23
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "S_NAME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   22
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "REGNO"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   21
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NATIONALIT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   20
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SPORT_ACTVTY"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   19
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "FACILITIES"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "DURATION"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "CHARGES"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   7680
      Width           =   2895
   End
End
Attribute VB_Name = "FRMSPOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
