VERSION 5.00
Begin VB.Form FRMSEARCHGUESTOREMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "&REPORT"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "FRMSEARCHGUESTOREMP.frx":0000
   ScaleHeight     =   6375
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton SEARCHDELEMPREPORT_CMD 
      BackColor       =   &H8000000D&
      Caption         =   "&REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox EMPID_CMB 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   48
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox GID_CMB 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton SEARCHGUESTREPORT_CMD 
      BackColor       =   &H8000000D&
      Caption         =   "&REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame EMPDETAIL_FRAME 
      BackColor       =   &H8000000A&
      Caption         =   "EMPLOYEE DETAIL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   1680
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   9855
      Begin VB.TextBox EEMAIL_TXT 
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
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox ENAME_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   34
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox EIDCARDNO_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   33
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox ECONTACT_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox EADDRESS_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   1455
         Left            =   6840
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox EDEPARTMENT_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox EDOA_TXT 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   29
         Top             =   2280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox EEXPERIENCE_TXT 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   28
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox ESALARY_TXT 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   27
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label DEL_LBL 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   5160
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "EXPERIENCE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ID CARD NO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   480
         TabIndex        =   40
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label DEPARTMENT_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label DOA_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF ASSIGN"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.CommandButton EXIT_CMD 
      BackColor       =   &H80000002&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8640
      Top             =   840
   End
   Begin VB.Frame GUESTDETAIL_FRAME 
      BackColor       =   &H8000000A&
      Caption         =   "GUEST DETAIL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   9855
      Begin VB.TextBox GCHECKINTIME_TXT 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox GCHECKOUTDATE_TXT 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   22
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox GPROFESSION_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox GADDRESS_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   1455
         Left            =   6840
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox GCONTACT_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox GIDCARDNO_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox GNAME_TXT 
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
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox GEMAIL_TXT 
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
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox GCHECKINDATE_TXT 
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
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label CHECKINTIME_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK IN TIME"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label CHECKOUTDATE_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK OUT DATE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label LBLRESERVPROFESSION 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PROFESSION"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label ADDRESS_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label LBLRESERVCONTACT 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label LBLRESERVIDCARDNO 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ID CARD NO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label NAME_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LBLRESERVEMAIL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label CHEKINDATE_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK IN DATE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   2280
         Width           =   2415
      End
   End
   Begin VB.CommandButton SEARCHEMPREPORT_CMD 
      BackColor       =   &H8000000D&
      Caption         =   "&REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label SEARCHTIME_LBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   7800
      TabIndex        =   20
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME:-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   6840
      TabIndex        =   19
      Top             =   0
      Width           =   855
   End
   Begin VB.Label SEARCHDATE_LBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   1440
      TabIndex        =   18
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE:-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   0
      Width           =   975
   End
   Begin VB.Label IDSARCH_LBL 
      BackColor       =   &H8000000A&
      Caption         =   "ID:-"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "FRMSEARCHGUESTOREMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub IDSEARCH_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
GOCHECKIN_CMD.SetFocus
End If
End Sub





Private Sub SAVECHECKIN_CMD_Click()
Unload Me
Me.Hide
End Sub


Private Sub Command1_Click()

End Sub

Private Sub EMPID_CMB_Click()
   DEL = 0
    S = "SELECT *FROM EMP_RECORD WHERE EMP_ID='" & EMPID_CMB.Text & "'"
    Set R = C.Execute(S)
    If R.EOF = True Then
        S = "SELECT *FROM DEL_EMP_RECORD WHERE EMP_ID='" & EMPID_CMB.Text & "'"
        Set R = C.Execute(S)
        If R.EOF = True Then
            MsgBox "WRONG ID PLEASE SELECT ONTHER ID"
            EMPID_CMB.SetFocus
            ENAME_TXT.Text = ""
            EIDCARDNO_TXT.Text = ""
            ECONTACT_TXT.Text = ""
            EEMAIL_TXT.Text = ""
            EDOA_TXT.Text = ""
            EADDRESS_TXT.Text = ""
            EEXPERIENCE_TXT.Text = ""
            ESALARY_TXT.Text = ""
            DEL_LBL.Visible = False
        Else
            ENAME_TXT.Text = R.Fields("EMP_NAME")
            EDOA_TXT.Text = R.Fields("DOA")
            EADDRESS_TXT.Text = R.Fields("ADDRESS")
            EIDCARDNO_TXT.Text = R.Fields("IDCARDNO")
            ECONTACT_TXT.Text = R.Fields("CONTACT")
            DEL_LBL.Visible = True
            DEL_LBL.Caption = "HE IS NOW NOT EMPLOYEE OF HOTEL FROM " & R.Fields("DO_RESINE")
            DEL = 1
        End If
    Else
        DEL_LBL.Visible = False
        ENAME_TXT.Text = R.Fields("EMP_NAME")
        EDOA_TXT.Text = R.Fields("DOA")
        EADDRESS_TXT.Text = R.Fields("ADDRESS")
        EIDCARDNO_TXT.Text = R.Fields("IDCARDNO")
        ECONTACT_TXT.Text = R.Fields("CONTACT")
        EEMAIL_TXT.Text = R.Fields("EMAIL")
        EEXPERIENCE_TXT.Text = R.Fields("EXPERIENCE")
        ESALARY_TXT.Text = R.Fields("SALARY")
        EDEPARTMENT_TXT.Text = R.Fields("DEPARTMENT")
    End If
If DEL = 0 Then
SEARCHEMPREPORT_CMD.Visible = True
SEARCHDELEMPREPORT_CMD.Visible = False
Else
SEARCHEMPREPORT_CMD.Visible = False
SEARCHDELEMPREPORT_CMD.Visible = True
End If
End Sub

Private Sub EXIT_CMD_Click()
Unload Me
Me.Hide
End Sub




Private Sub GID_CMB_Click()
    S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & GID_CMB.Text & "'"
    Set R = C.Execute(S)
    If R.EOF = True Then
            MsgBox "WRONG ID PLEASE SELECT ONTHER ID"
            GID_CMB.SetFocus
            GNAME_TXT.Text = ""
            GIDCARDNO_TXT.Text = ""
            GCONTACT_TXT.Text = ""
            GEMAIL_TXT.Text = ""
            GPROFESSION_TXT.Text = ""
            GADDRESS_TXT.Text = ""
            GCHECKINDATE_TXT.Text = ""
            GCHECKOUTDATE_TXT.Text = ""
            GCHECKINTIME_TXT.Text = ""
        Else
            GNAME_TXT.Text = R.Fields("NAME")
            GPROFESSION_TXT.Text = R.Fields("PROFESSION")
            GADDRESS_TXT.Text = R.Fields("ADDRESS")
            GIDCARDNO_TXT.Text = R.Fields("ID_CARD_NO")
            GCONTACT_TXT.Text = R.Fields("CONTACT")
            GEMAIL_TXT.Text = R.Fields("EMAIL")
            GCHECKINDATE_TXT.Text = R.Fields("CHECK_IN_DATE")
            GCHECKOUTDATE_TXT.Text = R.Fields("CHECK_OUT_DATE")
            GCHECKINTIME_TXT.Text = R.Fields("CHECK_IN_TIME")
        End If
End Sub

Private Sub SEARCHDELEMPREPORT_CMD_Click()
If EMPID_CMB.Text = "" Then
                    MsgBox "PLESE SELECT ID FIRST"
                    EMPID_CMB.SetFocus
                Else
                    If DataEnvironment1.rsCommand3.State = 1 Then DataEnvironment1.rsCommand3.Close
                    DataEnvironment1.Command3 EMPID_CMB.Text
                    DataReport4.Show
                End If
End Sub

Private Sub SEARCHEMPREPORT_CMD_Click()
If EMPID_CMB.Text = "" Then
                    MsgBox "PLESE SELECT ID FIRST"
                    EMPID_CMB.SetFocus
                Else
                    If DataEnvironment1.rsCommand2.State = 1 Then DataEnvironment1.rsCommand2.Close
                    DataEnvironment1.Command2 EMPID_CMB.Text
                    DataReport3.Show
                End If

End Sub

Private Sub SEARCHGUESTREPORT_CMD_Click()
                If GID_CMB.Text = "" Then
                    MsgBox "PLESE SELECT ID FIRST"
                    GID_CMB.SetFocus
                Else
                    If DataEnvironment1.rsCommand1.State = 1 Then DataEnvironment1.rsCommand1.Close
                        DataEnvironment1.Command1 GID_CMB.Text
                        DataReport5.Show
                End If
End Sub

Private Sub Timer1_Timer()
SEARCHTIME_LBL.Caption = Time
SEARCHDATE_LBL.Caption = Date
End Sub
