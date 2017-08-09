VERSION 5.00
Begin VB.Form FRMACOMO 
   Caption         =   "ACCOMMODATION FORM"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   Picture         =   "FRMACOMO.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text12 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   27
      Top             =   7560
      Width           =   2895
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8880
      Width           =   2535
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8880
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8880
      Width           =   2415
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8880
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
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
      Left            =   8160
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   10
      Top             =   6960
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   9
      Top             =   6360
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   8
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   7
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   6
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
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
      Left            =   3600
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   360
      TabIndex        =   28
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NATIONLATY"
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
      Left            =   360
      TabIndex        =   26
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "AMNT_CHARGE"
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
      Left            =   360
      TabIndex        =   25
      Top             =   6360
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
      Left            =   360
      TabIndex        =   24
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "R_CHARGE"
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
      Left            =   360
      TabIndex        =   22
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM_CATGR"
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
      Left            =   360
      TabIndex        =   21
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ROOM_NO"
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
      Left            =   360
      TabIndex        =   20
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label4 
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
      Left            =   360
      TabIndex        =   19
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "IDCARD_NO"
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
      Left            =   360
      TabIndex        =   18
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   17
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   16
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "FRMACOMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
