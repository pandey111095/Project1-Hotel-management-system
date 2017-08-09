VERSION 5.00
Begin VB.Form FRMSIDEBAR 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SIDEBAR"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "    Today"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   8400
      Width           =   2775
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "   User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   2775
      Begin VB.Image Image1 
         Height          =   240
         Left            =   0
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   0
      Picture         =   "FRMSIDEBAR.frx":0000
      Top             =   480
      Width           =   750
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   0
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2280
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "PICK A TASK"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   360
      TabIndex        =   22
      Top             =   0
      Width           =   1620
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   0
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   0
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   0
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   0
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Check In"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1080
      TabIndex        =   21
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   20
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Search Guest"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   19
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Status Of Hotel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   18
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Search Employee"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   17
      Top             =   2520
      Width           =   1710
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   0
      Top             =   5280
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   16
      Top             =   5280
      Width           =   1005
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   0
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image13 
      Height          =   5130
      Left            =   0
      Top             =   6240
      Width           =   8250
   End
   Begin VB.Image Image14 
      Height          =   480
      Left            =   0
      Top             =   6720
      Width           =   480
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   15
      Top             =   5880
      Width           =   1770
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   14
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   13
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Daily Evaluation Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   12
      Top             =   4920
      Width           =   2325
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   11
      Top             =   4440
      Width           =   345
   End
   Begin VB.Image Image15 
      Height          =   480
      Left            =   0
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   0
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Make Reservation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   10
      Top             =   3960
      Width           =   1770
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Process Payroll"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   9
      Top             =   3480
      Width           =   1545
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   0
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image Image18 
      Height          =   480
      Left            =   0
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Make Payment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   0
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "FRMSIDEBAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label19_Click()

End Sub

Private Sub Image4_Click()

End Sub
