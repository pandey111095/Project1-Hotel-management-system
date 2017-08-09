VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMRESERVATION 
   BackColor       =   &H80000003&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESERVATION FORM"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   14355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "FRMBOK.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDRESERVCLEAR 
      BackColor       =   &H80000002&
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12720
      Top             =   600
   End
   Begin VB.Frame FRAMERESERVGUEST 
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
      Left            =   5280
      TabIndex        =   21
      Top             =   2400
      Width           =   9015
      Begin VB.ComboBox CMBRESERVIDCARDTYPE 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FRMBOK.frx":6795
         Left            =   4920
         List            =   "FRMBOK.frx":67AB
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TXTRESERVPROFESSION 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Frame FRAMERESERVGENDER 
         BackColor       =   &H8000000A&
         Caption         =   "GENDER"
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
         Height          =   1215
         Left            =   5280
         TabIndex        =   33
         Top             =   2040
         Width           =   3615
         Begin VB.OptionButton OPTFEMALE 
            BackColor       =   &H00C0C0C0&
            Caption         =   "FEMALE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   13
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton OPTMALE 
            BackColor       =   &H00C0C0C0&
            Caption         =   "MALE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TXTRESERVGENDER 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.TextBox TXTRESERVAGE 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   2
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TXTRESERVADDRESS 
         BackColor       =   &H00C0C0C0&
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
         Left            =   6360
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TXTRESERVCONTACT 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   11
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox TXTRESERVIDCARDNO 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox TXTRESERVNAME 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   20
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox TXTRESERVEMAIL 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label LBLRESERVIDCARDTYPE 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ID CARD TYPE"
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
         Left            =   3000
         TabIndex        =   41
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label LBLRESERVPROFESSION 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   480
         TabIndex        =   36
         Top             =   1560
         Width           =   1620
      End
      Begin VB.Label LBLRESERVAGE 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "AGE"
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
         Left            =   480
         TabIndex        =   31
         Top             =   840
         Width           =   555
      End
      Begin VB.Label LBLRESEVADDRESS 
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
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label LBLRESERVCONTACT 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   480
         TabIndex        =   25
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label LBLRESERVIDCARDNO 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   480
         TabIndex        =   24
         Top             =   1200
         Width           =   1560
      End
      Begin VB.Label LBLRESERVNAME 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   480
         TabIndex        =   23
         Top             =   480
         Width           =   795
      End
      Begin VB.Label LBLRESERVEMAIL 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   480
         TabIndex        =   22
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.CommandButton CMDRESERVADD 
      BackColor       =   &H80000002&
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame FRAMERESERVROOM 
      BackColor       =   &H8000000A&
      Caption         =   "ROOM"
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
      Left            =   0
      TabIndex        =   17
      Top             =   2400
      Width           =   5295
      Begin VB.TextBox TXTRESERVROOMTYPE 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Height          =   360
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TXTRESERVBILLSTATUS 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPRESERV 
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   92864515
         CurrentDate     =   41970
      End
      Begin VB.TextBox TXTRESERVROOMNO 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox CMBRESERVROOMTYPE 
         BackColor       =   &H8000000A&
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
         ItemData        =   "FRMBOK.frx":67F7
         Left            =   1800
         List            =   "FRMBOK.frx":67F9
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TXTRESERVROOMCHARGE 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox CMBRESERVROOMNO 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   360
         TabIndex        =   40
         Top             =   2760
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label LBLRESERVBILL_STATUS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
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
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label LBLRSERVDTAE 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Reserv date"
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
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label LBLRSERVROOM_NO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room no"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label LBLRESERVROOMTYPE 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Room type"
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
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label LBLRSERVCHARGE 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Room charge"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton CMDRESERVSAVE 
      BackColor       =   &H80000002&
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton CMDRESERVBACK 
      BackColor       =   &H80000002&
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   15
      Height          =   855
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label GIDRESERV_LBL 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6480
      TabIndex        =   39
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Guest Id:-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4680
      TabIndex        =   38
      Top             =   1560
      Width           =   1710
   End
   Begin VB.Label LBLRESERVTIME 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13560
      TabIndex        =   29
      Top             =   120
      Width           =   90
   End
   Begin VB.Label LBLRESERVDATE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   90
   End
End
Attribute VB_Name = "FRMRESERVATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S1 As String
Dim G As String
Dim B As String

Private Sub CMBRESERVIDCARDTYPE_Click()
TXTRESERVIDCARDNO.SetFocus
End Sub

Private Sub CMBRESERVROOMTYPE_Click()
TXTRESERVROOMNO.Text = ""
TXTRESERVROOMTYPE.Text = CMBRESERVROOMTYPE.Text
CON
S = "SELECT COST FROM ROOM_TYPE WHERE TYPE='" & TXTRESERVROOMTYPE.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
MsgBox "PLEASE RE-ENTER ROOM TYPE"
Else
TXTRESERVROOMCHARGE.Text = R.Fields("COST")
CMBRESERVROOMNO.Clear
B = "NOTBOOKED"
CON
S1 = "SELECT *FROM ROOM  WHERE (TYPE='" & TXTRESERVROOMTYPE.Text & "' AND STATUS='" & B & "')"
Set R = C.Execute(S1)
Do Until R.EOF
CMBRESERVROOMNO.AddItem R.Fields("R_NO")
R.MoveNext
Loop
CMBRESERVROOMNO.SetFocus
End If
End Sub
Private Sub CMBRESERVROOMNO_Click()
TXTRESERVROOMNO.Text = CMBRESERVROOMNO.Text
TXTRESERVBILLSTATUS.SetFocus
End Sub

Private Sub CMDRESERVBACK_Click()
Unload Me
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub CMDRESERVCLEAR_Click()
'CMBRESERVIDCARDTYPE.Text = ""
TXTRESERVROOMTYPE.Text = ""
TXTRESERVROOMCHARGE.Text = ""
TXTRESERVROOMNO.Text = ""
TXTRESERVBILLSTATUS.Text = ""
TXTRESERVADDRESS.Text = ""
TXTRESERVCONTACT.Text = ""
TXTRESERVAGE.Text = ""
TXTRESERVNAME.Text = ""
TXTRESERVEMAIL.Text = ""
TXTRESERVGENDER.Text = ""
TXTRESERVPROFESSION.Text = ""
TXTRESERVIDCARDNO.Text = ""
CMBRESERVROOMTYPE.Refresh
CMBRESERVROOMNO.Refresh
OPTMALE.Refresh
OPTFEMALE.Refresh
End Sub

Private Sub CMDRESERVSAVE_Click()
If TXTRESERVROOMTYPE.Text = "" Then
MsgBox "PLEASE ENTER TYPE OF ROOM"
CMBRESERVROOMTYPE.SetFocus
ElseIf TXTRESERVROOMNO.Text = "" Then
MsgBox "PLEASE ALLOWED THE ROOM NO"
TXTRESERVROOMNO.SetFocus
ElseIf TXTRESERVNAME.Text = "" Then
MsgBox "PLEASE INPUT NAME OF GUEIST"
TXTRESERVNAME.SetFocus
ElseIf TXTRESERVAGE.Text = "" Then
MsgBox "PLEASE INPUT AGE OF GUEIST"
TXTRESERVAGE.SetFocus
ElseIf CMBRESERVIDCARDTYPE.Text = "" Then
MsgBox "PLEASE INPUT ID CARD TYPE OF GUEST"
CMBRESERVIDCARDTYPE.SetFocus
ElseIf TXTRESERVIDCARDNO.Text = "" Then
MsgBox "PLEASE INPUT ID NO"
TXTRESERVIDCARDNO.SetFocus
ElseIf TXTRESERVPROFESSION.Text = "" Then
MsgBox "PLEASE INPUT THE PROFESSION OF GUEIST"
TXTRESERVPROFESSION.SetFocus
ElseIf TXTRESERVCONTACT.Text = "" Then
MsgBox "PLEASE INPUT CONTACT NO"
TXTRESERVCONTACT.SetFocus
ElseIf TXTRESERVADDRESS.Text = "" Then
MsgBox "PLEASE INPUT ADDRESS"
TXTRESERVADDRESS.SetFocus
ElseIf TXTRESERVGENDER.Text = "" Then
MsgBox "PLEASE INPUT GENDER"
TXTRESERVGENDER.SetFocus
Else
If MsgBox("ARE YOU  SURE TO RESERVE THE ROOM", vbYesNoCancel) = vbYes Then
CON
S = "INSERT INTO CLIENT_MASTER VALUES('" & TXTRESERVNAME.Text & "','" & GIDRESERV_LBL.Caption & "'," & TXTRESERVROOMNO.Text & ",'" & TXTRESERVROOMTYPE.Text & "','" & Format(DTPRESERV.Value, "DD-MMM-YYYY") & "'," & TXTRESERVAGE.Text & ",'" & TXTRESERVGENDER.Text & "','" & TXTRESERVPROFESSION.Text & "','" & TXTRESERVADDRESS.Text & "'," & TXTRESERVBILLSTATUS.Text & ",'" & Format(DTPRESERV.Value, "DD-MMM-YYYY") & "','" & TXTRESERVIDCARDNO.Text & "'," & TXTRESERVCONTACT.Text & ",'" & TXTRESERVEMAIL.Text & "','" & Format(LBLRESERVDATE.Caption, "DD-MMM-YYYY") & "','" & LBLRESERVTIME.Caption & "','" & "NOTEXIST" & "'," & TXTRESERVROOMCHARGE.Text & "," & TXTRESERVROOMCHARGE.Text & "," & TXTRESERVROOMCHARGE.Text - TXTRESERVBILLSTATUS.Text & ",'IN',1,'" & CMBRESERVIDCARDTYPE.Text & "',0)"
'S = "INSERT INTO CLIENT_MASTER VALUES('" & TXTRESERVNAME.Text & "','" & GIDRESERV_LBL.Caption & "'," & TXTRESERVROOMNO.Text & ",'" & TXTRESERVROOMTYPE.Text & "','" & Format(DTPRESERV.Value, "DD-MMM-YYYY") & "'," & TXTRESERVAGE.Text & ",'" & TXTRESERVGENDER.Text & "','" & TXTRESERVPROFESSION.Text & "','" & TXTRESERVADDRESS.Text & "'," & TXTRESERVBILLSTATUS.Text & ",'& NOT EXIST &','" & TXTRESERVIDCARDNO.Text & "'," & TXTRESERVCONTACT.Text & ",'" & TXTRESERVEMAIL.Text & "','" & Format(LBLRESERVDATE.Caption, "DD-MMM-YYYY") & "','" & LBLRESERVTIME.Caption & "','" & "NOTEXIST" & "'," & TXTRESERVROOMCHARGE.Text & "," & TXTRESERVROOMCHARGE.Text & "," & Val(TXTRESERVROOMCHARGE.Text) - Val(TXTRESERVBILLSTATUS.Text) & ",'IN'," & 1 & ",'" & CMBRESERVIDCARDTYPE.Text & "'," & 0 & ")"
MsgBox S
Set R = C.Execute(S)
S = "UPDATE GCODE SET G=" & NOOFGUEST & " "
Set R = C.Execute(S)
BB = "BOOKED"
S = "UPDATE ROOM SET STATUS='" & BB & "' WHERE R_NO='" & TXTRESERVROOMNO.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "ROOM IS RESERVED"
CMBRESERVROOMTYPE.Clear
FRAMERESERVROOM.Enabled = False
FRAMERESERVGENDER.Enabled = False
FRAMERESERVGUEST.Enabled = False
GIDRESERV_LBL.Caption = ""
TXTRESERVROOMTYPE.Text = ""
TXTRESERVROOMCHARGE.Text = ""
TXTRESERVROOMNO.Text = ""
TXTRESERVBILLSTATUS.Text = ""
TXTRESERVADDRESS.Text = ""
TXTRESERVCONTACT.Text = ""
TXTRESERVAGE.Text = ""
TXTRESERVNAME.Text = ""
TXTRESERVEMAIL.Text = ""
TXTRESERVGENDER.Text = ""
TXTRESERVPROFESSION.Text = ""
TXTRESERVIDCARDNO.Text = ""
'CMBRESERVIDCARDTYPE.Text = ""
OPTMALE.Refresh
OPTFEMALE.Refresh
CMDRESERVSAVE.Enabled = False
CMDRESERVADD.Enabled = True
Else
MsgBox "ROOM NOT RESRVED"
'CMBRESERVROOMTYPE.Clear
'FRAMERESERVROOM.Enabled = False
'FRAMERESERVGENDER.Enabled = False
'FRAMERESERVGUEST.Enabled = False
'CMDRESERVSAVE.Enabled = False
'CMDRESERVADD.Enabled = True
'GIDRESERV_LBL.Caption = ""
'TXTRESERVROOMTYPE.Text = ""
'TXTRESERVROOMCHARGE.Text = ""
'TXTRESERVROOMNO.Text = ""
'TXTRESERVBILLSTATUS.Text = ""
'TXTRESERVADDRESS.Text = ""
'TXTRESERVCONTACT.Text = ""
'TXTRESERVAGE.Text = ""
'TXTRESERVNAME.Text = ""
'TXTRESERVEMAIL.Text = ""
'TXTRESERVGENDER.Text = ""
'TXTRESERVPROFESSION.Text = ""
'TXTRESERVIDCARDNO.Text = ""
End If
End If
End Sub
Private Sub CMDRESERVADD_Click()
FRAMERESERVROOM.Enabled = True
FRAMERESERVGUEST.Enabled = True
FRAMERESERVGENDER.Enabled = True
CMDRESERVSAVE.Enabled = True
CMDRESERVCLEAR.Enabled = True
TXTRESERVROOMTYPE.Text = ""
TXTRESERVROOMCHARGE.Text = ""
TXTRESERVROOMNO.Text = ""
TXTRESERVBILLSTATUS.Text = ""
TXTRESERVADDRESS.Text = ""
TXTRESERVCONTACT.Text = ""
TXTRESERVAGE.Text = ""
TXTRESERVNAME.Text = ""
TXTRESERVEMAIL.Text = ""
TXTRESERVGENDER.Text = ""
TXTRESERVPROFESSION.Text = ""
TXTRESERVIDCARDNO.Text = ""
'CMBRESERVIDCARDTYPE.Text = ""
CON
S = "SELECT distinct(type)FROM ROOM_TYPE "
Set R = C.Execute(S)
Do Until R.EOF
CMBRESERVROOMTYPE.AddItem R.Fields("TYPE")   '
R.MoveNext
Loop
CON
S = "SELECT (G) FROM GCODE"
Set R = C.Execute(S)
    NOOFGUEST = R.Fields("G")
NOOFGUEST = NOOFGUEST + 1
GIDRESERV_LBL.Caption = "G" & NOOFGUEST
DTPRESERV.SetFocus
CMDRESERVADD.Enabled = False
End Sub





Private Sub DTPRESERV_Click()
CMBRESERVROOMTYPE.SetFocus
End Sub

Private Sub Form_Load()
DTPRESERV.MinDate = Date
FRMRESERVATION.Width = 14500
FRMRESERVATION.Height = 7330
FRMRESERVATION.Left = 2000
FRMRESERVATION.Top = 1000
FRAMERESERVROOM.Enabled = False
FRAMERESERVGUEST.Enabled = False
FRAMERESERVGENDER.Enabled = False
CMDRESERVSAVE.Enabled = False
CMDRESERVCLEAR.Enabled = False
CON
S = "SELECT COUNT (*)FROM GCODE"
Set R = C.Execute(S)
D = R.Fields(0)
If D = 0 Then
S = "INSERT INTO GCODE VALUES(" & 0 & ")"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
Else
End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
TXTRESERVDATE.Text = MonthView1.Value
MonthView1.Visible = False
End Sub

Private Sub OPTMALE_Click()
TXTRESERVGENDER.Text = OPTMALE.Caption
End Sub

Private Sub OPTFEMALE_Click()
TXTRESERVGENDER.Text = OPTFEMALE.Caption
End Sub


Private Sub TXTRESERVAGE_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    CMBRESERVIDCARDTYPE.SetFocus
    End If
End Sub
Private Sub TXTRESERVBILLSTATUS_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
TXTRESERVNAME.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
LBLRESERVTIME.Caption = Time
LBLRESERVDATE.Caption = Date
End Sub

Private Sub TXTRESERVBILLSTATUS_LostFocus()
If Val(TXTRESERVBILLSTATUS.Text) > Val(TXTRESERVROOMCHARGE.Text) Then
    E = Val(TXTRESERVBILLSTATUS.Text) - Val(TXTRESERVROOMCHARGE.Text)
    Label2.Visible = True
    Label2.Caption = "RETURN HIM/HER " & E & " RUPEES "
Else
    Label2.Visible = False
    Label2.Caption = ""
End If
End Sub

Private Sub TXTRESERVCONTACT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    TXTRESERVEMAIL.SetFocus
End If

End Sub

Private Sub TXTRESERVEMAIL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TXTRESERVADDRESS.SetFocus
End If
End Sub
Private Sub TXTRESERVIDCARDNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TXTRESERVPROFESSION.SetFocus
End If
End Sub


Private Sub TXTRESERVNAME_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    TXTRESERVAGE.SetFocus
End If
End Sub

Private Sub TXTRESERVPROFESSION_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    TXTRESERVCONTACT.SetFocus
End If
End Sub
