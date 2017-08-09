VERSION 5.00
Begin VB.Form FRMTRANS 
   Caption         =   "TRANSPORT BOOKING"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   Picture         =   "FRMTRANS.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
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
      Left            =   9840
      TabIndex        =   28
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
      Left            =   3360
      TabIndex        =   15
      Top             =   840
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
      Left            =   3360
      TabIndex        =   14
      Top             =   1440
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
      Left            =   3360
      TabIndex        =   13
      Top             =   2040
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
      Left            =   3360
      TabIndex        =   12
      Top             =   2640
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
      Left            =   3360
      TabIndex        =   11
      Top             =   3240
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
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
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
      Left            =   3360
      TabIndex        =   9
      Top             =   4440
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
      Left            =   3360
      TabIndex        =   8
      Top             =   5040
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
      Left            =   3360
      TabIndex        =   7
      Top             =   5640
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
      Left            =   3360
      TabIndex        =   6
      Top             =   6240
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
      Left            =   3360
      TabIndex        =   5
      Top             =   6840
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
      Left            =   11040
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
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
      TabIndex        =   3
      Top             =   7920
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
      TabIndex        =   2
      Top             =   7920
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   2535
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6720
      TabIndex        =   27
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   960
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   25
      Top             =   1560
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   2160
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ID_CARD"
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
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "VEHCILE"
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
      Left            =   240
      TabIndex        =   21
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label8 
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
      Left            =   240
      TabIndex        =   19
      Top             =   5160
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "CHARGED_NO"
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
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "COMISION"
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
      Left            =   240
      TabIndex        =   17
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   6840
      Width           =   2895
   End
End
Attribute VB_Name = "FRMTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
