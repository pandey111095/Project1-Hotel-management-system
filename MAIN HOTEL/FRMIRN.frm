VERSION 5.00
Begin VB.Form FRMIRN 
   Caption         =   "IRONIC FORM"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
      BackColor       =   &H0080C0FF&
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
      Left            =   9480
      TabIndex        =   30
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   16
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   15
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   14
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   13
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   12
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   11
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   10
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   9
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   8
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   7
      Top             =   5880
      Width           =   2895
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H0080C0FF&
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
      Left            =   3120
      TabIndex        =   6
      Top             =   6480
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080C0FF&
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
      Left            =   12960
      TabIndex        =   5
      Top             =   360
      Width           =   2655
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   2535
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H0080C0FF&
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
      Left            =   9480
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "RCPT_NO"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   6360
      TabIndex        =   29
      Top             =   960
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NATIONALITY"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "LINEN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF CLOTH"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT CHAR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOT_AMOUNT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   6240
      TabIndex        =   17
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "FRMIRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
