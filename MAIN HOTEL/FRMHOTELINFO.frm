VERSION 5.00
Begin VB.Form FRMHOTELINFO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HOTEL DETAIL"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMHOTELINFO.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text10 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1695
      Left            =   11280
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "FRMHOTELINFO.frx":2691F
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   17
      Text            =   "24 HOURS"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   11760
      TabIndex        =   15
      Text            =   "WWW.HOTELINN.COM"
      Top             =   6360
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   12480
      TabIndex        =   13
      Text            =   "08804934946"
      Top             =   5880
      Width           =   2655
   End
   Begin VB.TextBox TOTROOM_TXT 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   12480
      TabIndex        =   11
      Text            =   "200"
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox TOTEMP_TXT 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   12480
      TabIndex        =   9
      Text            =   "150"
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   7
      Text            =   "BORING ROAD, PATNA-20"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   5
      Text            =   "RAHUL KUMAR PANDEY"
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   3
      Text            =   "IND2901290129"
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   1
      Text            =   "MAHARAJA   INN   HOTEL "
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "FACILITIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   18
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "WORKTIME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "WEBSITE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTCT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL ROOM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL EMPLOYEE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PALCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OWNER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "REG_NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "FRMHOTELINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
S = "SELECT COUNT(EMP_ID) FROM EMP_RECORD"
Set R = C.Execute(S)
TOTEMP_TXT.Text = R.Fields(0)
S = "SELECT COUNT(R_NO) FROM ROOM"
Set R = C.Execute(S)
TOTROOM_TXT.Text = R.Fields(0)
End Sub
