VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMCHECKIN 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHECK IN FORM"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   15570
   Begin VB.CommandButton CLEARCHECKIN_CMD 
      BackColor       =   &H80000002&
      Caption         =   "CLEAR"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5400
      Width           =   2055
   End
   Begin VB.ComboBox GUESTIDCHEKIN_CMB 
      BackColor       =   &H80000000&
      Height          =   315
      Left            =   3000
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton UPDATECHECKIN_CMD 
      BackColor       =   &H80000002&
      Caption         =   "&UPDATE"
      Enabled         =   0   'False
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton SEARCHCHECKIN_CMD 
      BackColor       =   &H80000002&
      Caption         =   "&SEARCH RESERV"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   3615
   End
   Begin VB.TextBox GIDCHECKIN_TXT 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
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
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton BACKTOMDIFROMCHECKIN_CMD 
      BackColor       =   &H80000002&
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
      Height          =   435
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton SAVECHECKIN_CMD 
      BackColor       =   &H80000002&
      Caption         =   "&SAVE"
      Enabled         =   0   'False
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Frame ROOMCHECKIN_FRAME 
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
      Left            =   120
      TabIndex        =   31
      Top             =   1800
      Width           =   5415
      Begin VB.TextBox CHECKOUTDATE_TXT 
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
         Height          =   345
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox ROOMNOCHECKIN_CMB 
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
         Left            =   3720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox ROOMCHARGECHECKIN_TXT 
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
         Height          =   270
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox ROOMTYPECHECKIN_CMB 
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
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox ROOMNOCHECKIN_TXT 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox BILLSTATUSCHECKIN_TXT 
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
         Height          =   375
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox ROOMTYPECHECKIN_TXT 
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker CHECKINDATE_DTP 
         Height          =   495
         Left            =   2160
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3600
         TabIndex        =   47
         Top             =   2400
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label CHECKOUTDATE_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKOUT DATE"
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
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   2880
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label ROOMCHARGECHECKIN_LBL 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHARGE"
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
         TabIndex        =   39
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label ROOMTYPECHECKIN_LBL 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM TYPE"
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
         TabIndex        =   38
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label ROOMNOCHECKIN_LBL 
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM_NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label CHECKINDTAE_LBL 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKIN DATE"
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
         TabIndex        =   36
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label BILLSTATUSCHECKIN_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADVANCE"
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
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.CommandButton ADDNEWCHECKIN_CMD 
      BackColor       =   &H80000002&
      Caption         =   "&ADDNEW"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   2055
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
      Left            =   5520
      TabIndex        =   22
      Top             =   1800
      Width           =   9855
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
         ItemData        =   "FRMADM.frx":0000
         Left            =   4800
         List            =   "FRMADM.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox GIDCARDTYPECHECKIN_TXT 
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
         Height          =   285
         Left            =   4800
         MaxLength       =   11
         TabIndex        =   49
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox CHECKINTIME_TXT 
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
         MaxLength       =   12
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox GEMAILCHECKIN_TXT 
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
         TabIndex        =   11
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox GNAMECHECKIN_TXT 
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
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox GIDCARDNOCHECKIN_TXT 
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
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox GCONTACTCHECKIN_TXT 
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
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox GADDRESSCHECKIN_TXT 
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
         Left            =   6840
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox GAGECHECKIN_TXT 
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
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Frame GENDERCHECKIN_FRAME 
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
         TabIndex        =   23
         Top             =   2040
         Width           =   3615
         Begin VB.TextBox GGENDERCHECKIN_TXT 
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
            MaxLength       =   6
            TabIndex        =   16
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton MALE_OPT 
            BackColor       =   &H8000000A&
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
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton FEMALE_OPT 
            BackColor       =   &H8000000A&
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
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.TextBox GPROFESSIONCHECKIN_TXT 
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
         TabIndex        =   9
         Top             =   1560
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
         Left            =   2880
         TabIndex        =   48
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label CHEKINTIME_LBL 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKIN TIME"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   480
         TabIndex        =   43
         Top             =   2640
         Visible         =   0   'False
         Width           =   1635
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
         TabIndex        =   30
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label LBLRESERVNAME 
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
         TabIndex        =   29
         Top             =   480
         Width           =   1095
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
         TabIndex        =   28
         Top             =   1200
         Width           =   1695
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
         TabIndex        =   27
         Top             =   1920
         Width           =   1335
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
      Begin VB.Label LBLRESERVAGE 
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
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   840
         Width           =   975
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
         TabIndex        =   24
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8760
      Top             =   840
   End
   Begin VB.Label GIDCHECKIN_LBL 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "GUEST ID:-"
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
      Height          =   360
      Left            =   1440
      TabIndex        =   42
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label RESERVDATE_LBL 
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
      Left            =   120
      TabIndex        =   41
      Top             =   0
      Width           =   975
   End
   Begin VB.Label CHECKINTIME_LBL 
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
      Left            =   13800
      TabIndex        =   40
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "FRMCHECKIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BB As String
Dim B As String
Private Sub ADDNEWCHECKIN_CMD_Click()
CHECKOUTDATE_LBL.Visible = False
CHECKOUTDATE_TXT.Visible = False
CHECKINTIME_LBL.Visible = False
CHECKINTIME_TXT.Visible = False
GUESTIDCHEKIN_CMB.Visible = False
GIDCHECKIN_TXT.Visible = True
GIDCARDTYPECHECKIN_TXT.Visible = False
CMBRESERVIDCARDTYPE.Visible = True
GIDCARDTYPECHECKIN_TXT.Text = ""
ROOMNOCHECKIN_CMB.Refresh
MALE_OPT.Refresh
FEMALE_OPT.Refresh
SEARCHCHECKIN_CMD.Enabled = False
ROOMCHECKIN_FRAME.Enabled = True
GUESTDETAIL_FRAME.Enabled = True
GENDERCHECKIN_FRAME.Enabled = True
GENDERCHECKIN_FRAME.Enabled = True
SAVECHECKIN_CMD.Enabled = True
ROOMTYPECHECKIN_TXT.Text = ""
ROOMCHARGECHECKIN_TXT.Text = ""
ROOMNOCHECKIN_TXT.Text = ""
BILLSTATUSCHECKIN_TXT.Text = ""
GNAMECHECKIN_TXT.Text = ""
GAGECHECKIN_TXT.Text = ""
GIDCARDNOCHECKIN_TXT.Text = ""
GPROFESSIONCHECKIN_TXT.Text = ""
GCONTACTCHECKIN_TXT.Text = ""
GEMAILCHECKIN_TXT.Text = ""
GADDRESSCHECKIN_TXT.Text = ""
GGENDERCHECKIN_TXT.Text = ""
CON
S = "SELECT distinct(type)FROM ROOM_TYPE "
Set R = C.Execute(S)
Do Until R.EOF
ROOMTYPECHECKIN_CMB.AddItem R.Fields("TYPE")   '
R.MoveNext
Loop
CON
S = "SELECT G FROM GCODE"
Set R = C.Execute(S)
NOOFGUEST = R.Fields("G")
NOOFGUEST = NOOFGUEST + 1
GIDCHECKIN_TXT.Text = "G" & NOOFGUEST
CHECKINDATE_DTP.SetFocus
ADDNEWCHECKIN_CMD.Enabled = False
End Sub

Private Sub BACKTOMDIFROMCHECKIN_CMD_Click()
Unload Me
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub BILLSTATUSCHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
GNAMECHECKIN_TXT.SetFocus
End If
End Sub

Private Sub FEMALEOPT_Click()
TXTRESERVGENDER.Text = FEMALE_OPTCaption
End Sub

'Private Sub CLEARCHECKIN_CMD_Click()
'Unload FRMCHECKIN
'Load FRMCHECKIN
'FRMCHECKIN.Show
'End Sub

Private Sub BILLSTATUSCHECKIN_TXT_LostFocus()
If Val(ROOMCHARGECHECKIN_TXT.Text) < Val(BILLSTATUSCHECKIN_TXT.Text) Then
    Label1.Visible = True
    E = Val(BILLSTATUSCHECKIN_TXT.Text) - Val(ROOMCHARGECHECKIN_TXT.Text)
    Label1.Caption = "RETURN " & E & "RUPPES"
Else
    Label1.Caption = ""
    Label1.Visible = False
End If
End Sub



Private Sub CLEARCHECKIN_CMD_Click()
Label1.Caption = ""
Label1.Visible = False
GIDCHECKIN_TXT.Visible = False
GUESTIDCHEKIN_CMB.Visible = False
ROOMCHECKIN_FRAME.Enabled = False
GUESTDETAIL_FRAME.Enabled = False
SAVECHECKIN_CMD.Enabled = False
ADDNEWCHECKIN_CMD.Enabled = True
SEARCHCHECKIN_CMD.Enabled = True
GIDCARDTYPECHECKIN_TXT.Visible = False
CMBRESERVIDCARDTYPE.Visible = True
'CMBRESERVIDCARDTYPE.Text = ""
GIDCARDTYPECHECKIN_TXT.Text = ""
GIDCHECKIN_TXT.Text = ""
ROOMTYPECHECKIN_TXT.Text = ""
ROOMCHARGECHECKIN_TXT.Text = ""
ROOMNOCHECKIN_TXT.Text = ""
BILLSTATUSCHECKIN_TXT.Text = ""
GNAMECHECKIN_TXT.Text = ""
GAGECHECKIN_TXT.Text = ""
'GCOUNTRYCHECKIN_TXT.Text = ""
GIDCARDNOCHECKIN_TXT.Text = ""
GPROFESSIONCHECKIN_TXT.Text = ""
GCONTACTCHECKIN_TXT.Text = ""
GEMAILCHECKIN_TXT.Text = ""
GADDRESSCHECKIN_TXT.Text = ""
GGENDERCHECKIN_TXT.Text = ""
ROOMTYPECHECKIN_CMB.Clear
End Sub

Private Sub FEMALE_OPT_Click()
GGENDERCHECKIN_TXT.Text = FEMALE_OPT.Caption
End Sub

Private Sub Form_Load()
Me.Top = 1000
Me.Left = 3000
'CHECKINDATE_DTP.MinDate = Date
'Label2.Caption = ""
'Label2.Enabled = False
CHECKINDATE_DTP.Value = Now
CON
S = "SELECT COUNT (*)FROM GCODE"
Set R = C.Execute(S)
D = R.Fields(0)
If D = 0 Then
S = "INSERT INTO GCODE VALUES(" + 0 + ")"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
Else
End If
End Sub

Private Sub GADDRESSCHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
End Sub

Private Sub GAGECHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    GIDCARDNOCHECKIN_TXT.SetFocus
End If
End Sub



Private Sub GCONTACTCHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    GEMAILCHECKIN_TXT.SetFocus
End If

End Sub

'Private Sub GIDCHECKIN_TXT_Change()
'If Len(GIDCHECKIN_TXT.Text) > 0 Then
'GOCHECKIN_CMD.Enabled = True
'Else
'GOCHECKIN_CMD.Enabled = False
'End If
'End Sub

Private Sub GIDCHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
GIDCHECKIN_TXT.Text = UCase(GIDCHECKIN_TXT.Text)
GOCHECKIN_CMD.SetFocus
End If
End Sub

Private Sub GNAMECHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
GAGECHECKIN_TXT.SetFocus
End If
End Sub

Private Sub GIDCARDNOCHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
GPROFESSIONCHECKIN_TXT.SetFocus
End If
End Sub


Private Sub GPROFESSIONCHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
GCONTACTCHECKIN_TXT.SetFocus
End If
End Sub

Private Sub GEMAILCHECKIN_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
GADDRESSCHECKIN_TXT.SetFocus
End If
End Sub

Private Sub GUESTIDCHEKIN_CMB_Click()
CON
S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID= '" & GUESTIDCHEKIN_CMB.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
    MsgBox "WRONG GUEST ID", vbCritical
    Else
    'GIDCHECKIN_TXT.Enabled = False
    GNAMECHECKIN_TXT.Text = R.Fields("NAME")
    GAGECHECKIN_TXT.Text = R.Fields("AGE")
    GADDRESSCHECKIN_TXT.Text = R.Fields("ADDRESS")
    GEMAILCHECKIN_TXT.Text = IIf(IsNull(R.Fields("EMAIL")), "", R.Fields("EMAIL"))
    GCONTACTCHECKIN_TXT.Text = R.Fields("CONTACT")
    GIDCARDNOCHECKIN_TXT.Text = R.Fields("ID_CARD_NO")
    ROOMCHARGECHECKIN_TXT.Text = R.Fields("ROOM_CHARGE")
    ROOMNOCHECKIN_TXT.Text = R.Fields("SUIT_NO")
    CHECKINDATE_DTP.Value = R.Fields("CHECK_IN_DATE")
    ROOMTYPECHECKIN_TXT.Text = R.Fields("SUIT_PROFILE")
    BILLSTATUSCHECKIN_TXT.Text = R.Fields("BILL_STATUS")
    GGENDERCHECKIN_TXT.Text = R.Fields("SEX")
    GPROFESSIONCHECKIN_TXT.Text = R.Fields("PROFESSION")
    GIDCARDTYPECHECKIN_TXT.Text = R.Fields("ID_CARD_TYPE")
    CHEKINTIME_LBL.Visible = True
    CHECKOUTDATE_LBL.Visible = True
    CHECKOUTDATE_TXT.Visible = True
    CHECKINTIME_TXT.Visible = True
    CHECKOUTDATE_TXT.Text = R.Fields("CHECK_OUT_DATE")
    CHECKINTIME_TXT.Text = R.Fields("CHECK_IN_TIME")
    If R.Fields("CHECK_IN_TIME") = "NOTEXIST" Then
        If MsgBox("CHECK IN NOW", vbYesNo) = vbYes Then
            CON
            S = "UPDATE  CLIENT_MASTER SET CHECK_IN_DATE='" & Format(CHECKINDATE_DTP.Value, "DD-MMM-YYYY") & "',CHECK_IN_TIME='" & Time & "' WHERE CLIENT_ID='" & GUESTIDCHEKIN_CMB.Text & "' "
            Set R = C.Execute(S)
            'MsgBox "CHEACKED IN"
            MsgBox "CHECK IN COMPLETE"
            ROOMCHECKIN_FRAME.Enabled = False
            GUESTDETAIL_FRAME.Enabled = False
            ADDNEWCHECKIN_CMD.Enabled = True
            SEARCHCHECKIN_CMD.Enabled = True
            ROOMTYPECHECKIN_CMB.Clear
            CHECKINTIME_TXT.Text = Time
        Else
            UPDATECHECKIN_CMD.Enabled = True
            ROOMCHECKIN_FRAME.Enabled = False
            GUESTDETAIL_FRAME.Enabled = True
            GENDERCHECKIN_FRAME.Enabled = True
            CON
            S = "SELECT distinct(type)FROM ROOM_TYPE "
            Set R = C.Execute(S)
            Do Until R.EOF
            ROOMTYPECHECKIN_CMB.AddItem R.Fields("TYPE")   '
            R.MoveNext
            Loop
            MsgBox " GNAMECHECKIN_TXT.Text &  IS NOT CHECKIN"
        End If
    Else
        MsgBox "CHECKED IN"
        UPDATECHECKIN_CMD.Enabled = True
        ROOMCHECKIN_FRAME.Enabled = False
        GUESTDETAIL_FRAME.Enabled = True
        GENDERCHECKIN_FRAME.Enabled = True

    End If
End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub MALE_OPT_Click()
GGENDERCHECKIN_TXT.Text = MALE_OPT.Caption
End Sub

Private Sub ROOMNOCHECKIN_CMB_Click()
ROOMNOCHECKIN_TXT.Text = ROOMNOCHECKIN_CMB.Text
BILLSTATUSCHECKIN_TXT.SetFocus
End Sub

Private Sub ROOMTYPECHECKIN_CMB_Click()
ROOMTYPECHECKIN_TXT.Text = ROOMTYPECHECKIN_CMB.Text
CON
S = "SELECT COST FROM ROOM_TYPE WHERE TYPE='" & ROOMTYPECHECKIN_TXT.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
MsgBox "PLEASE RE-ENTER ROOM TYPE"
Else
ROOMCHARGECHECKIN_TXT.Text = R.Fields("COST")
ROOMNOCHECKIN_CMB.Clear
B = "NOTBOOKED"
CON
S1 = "SELECT *FROM ROOM  WHERE (TYPE='" & ROOMTYPECHECKIN_TXT.Text & "' AND STATUS='" & B & "')"
Set R = C.Execute(S1)
Do Until R.EOF
ROOMNOCHECKIN_CMB.AddItem R.Fields("R_NO")
R.MoveNext
Loop
ROOMNOCHECKIN_CMB.SetFocus
End If
End Sub
Private Sub SAVECHECKIN_CMD_Click()
If ROOMTYPECHECKIN_TXT.Text = "" Then
MsgBox "PLEASE ENTER TYPE OF ROOM"
ROOMTYPECHECKIN_CMB.SetFocus
ElseIf ROOMNOCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE ALLOWED THE ROOM NO"
ROOMNOCHECKIN_CMB.SetFocus
ElseIf GNAMECHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT NAME OF GUEST"
GNAMECHECKIN_TXT.SetFocus
ElseIf GAGECHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT AGE OF GUEST"
GAGECHECKIN_TXT.SetFocus
ElseIf CMBRESERVIDCARDTYPE.Text = "" Then
MsgBox "PLEASE INPUT ID CARD TYPE OF GUEST"
CMBRESERVIDCARDTYPE.SetFocus
ElseIf GIDCARDNOCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT ID NO"
GIDCARDNOCHECKIN_TXT.SetFocus
ElseIf GPROFESSIONCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT THE PROFESSION OF GUEIST"
GPROFESSIONCHECKIN_TXT.SetFocus
ElseIf GCONTACTCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT CONTACT NO"
GCONTACTCHECKIN_TXT.SetFocus
ElseIf GADDRESSCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT ADDRESS"
GADDRESSCHECKIN_TXT.SetFocus
ElseIf GGENDERCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT GENDER"
GGENDERCHECKIN_TXT.SetFocus
Else
If MsgBox("ARE YOU SURE TO RESERVE THE ROOM", vbYesNoCancel) = vbYes Then
    CON
    S = "INSERT INTO CLIENT_MASTER VALUES('" & GNAMECHECKIN_TXT.Text & "','" & GIDCHECKIN_TXT.Text & "'," & ROOMNOCHECKIN_TXT.Text & ",'" & ROOMTYPECHECKIN_TXT.Text & "','" & Format(CHECKINDATE_DTP.Value, "DD-MMM-YYYY") & "'," & GAGECHECKIN_TXT.Text & ",'" & GGENDERCHECKIN_TXT.Text & "','" & GPROFESSIONCHECKIN_TXT.Text & "','" & GADDRESSCHECKIN_TXT.Text & "'," & BILLSTATUSCHECKIN_TXT.Text & ",'" & Format(CHECKINDATE_DTP.Value, "DD-MMM-YYYY") & "','" & GIDCARDNOCHECKIN_TXT.Text & "'," & GCONTACTCHECKIN_TXT.Text & ",'" & GEMAILCHECKIN_TXT.Text & "','" & Format(RESERVDATE_LBL.Caption, "DD-MMM-YYYY") & "','" & CHECKINTIME_LBL.Caption & "','" & CHECKINTIME_LBL.Caption & "'," & ROOMCHARGECHECKIN_TXT.Text & "," & ROOMCHARGECHECKIN_TXT.Text & "," & Val(ROOMCHARGECHECKIN_TXT.Text) - Val(BILLSTATUSCHECKIN_TXT.Text) & ",'IN',1,'" & CMBRESERVIDCARDTYPE.Text & "',0)"
   ' S = "INSERT INTO CLIENT_MASTER VALUES('" & TXTRESERVNAME.Text & "','" & GIDRESERV_LBL.Caption & "'," & TXTRESERVROOMNO.Text & ",'" & TXTRESERVROOMTYPE.Text & "','" & Format(DTPRESERV.Value, "DD-MMM-YYYY") & "'," & TXTRESERVAGE.Text & ",'" & TXTRESERVGENDER.Text & "','" & TXTRESERVPROFESSION.Text & "','" & TXTRESERVADDRESS.Text & "'," & TXTRESERVBILLSTATUS.Text & ",'" & Format(DTPRESERV.Value, "DD-MMM-YYYY") & "','" & TXTRESERVIDCARDNO.Text & "'," & TXTRESERVCONTACT.Text & ",'" & TXTRESERVEMAIL.Text & "','" & Format(LBLRESERVDATE.Caption, "DD-MMM-YYYY") & "','" & LBLRESERVTIME.Caption & "','" & "NOTEXIST" & "'," & TXTRESERVROOMCHARGE.Text & "," & TXTRESERVROOMCHARGE.Text & "," & TXTRESERVROOMCHARGE.Text - TXTRESERVBILLSTATUS.Text & ",'IN',1,'" & CMBRESERVIDCARDTYPE.Text & "')"
    'S = "INSERT INTO CLIENT_MASTER VALUES('" & GNAMECHECKIN_TXT.Text & "','" & GIDCHECKIN_TXT.Text & "'," & ROOMNOCHECKIN_TXT.Text & ",'" & ROOMTYPECHECKIN_TXT.Text & "','" & Format(CHECKINDATE_DTP.Value, "DD-MMM-YYYY") & "'," & GAGECHECKIN_TXT.Text & ",'" & GGENDERCHECKIN_TXT.Text & "','" & GPROFESSIONCHECKIN_TXT.Text & "','" & GADDRESSCHECKIN_TXT.Text & "'," & BILLSTATUSCHECKIN_TXT.Text & ",'NOTEXIST','" & GIDCARDNOCHECKIN_TXT.Text & "'," & GCONTACTCHECKIN_TXT.Text & ",'" & GEMAILCHECKIN_TXT.Text & "','" & Format(RESERVDATE_LBL.Caption, "DD-MMM-YYYY") & "','" & CHECKINTIME_LBL.Caption & "','" & CHECKINTIME_LBL.Caption & "'," & ROOMCHARGECHECKIN_TXT.Text & "," & ROOMCHARGECHECKIN_TXT.Text & "," & Val(ROOMCHARGECHECKIN_TXT.Text) - Val(BILLSTATUSCHECKIN_TXT.Text) & ",'IN',1,'" & CMBRESERVIDCARDTYPE.Text & "'," & 0 & ")"
    Set R = C.Execute(S)
    S = "UPDATE GCODE SET G=" & NOOFGUEST & " "
    Set R = C.Execute(S)
    BB = "BOOKED"
    S = "UPDATE ROOM SET STATUS='" & BB & "' WHERE R_NO=" & ROOMNOCHECKIN_TXT.Text & ""
    Set R = C.Execute(S)
    S = "COMMIT"
    Set R = C.Execute(S)
    MsgBox "ROOM IS RESERVED"
    ROOMCHECKIN_FRAME.Enabled = False
    GUESTDETAIL_FRAME.Enabled = False
    GIDCHECKIN_TXT.Text = ""
    ROOMTYPECHECKIN_TXT.Text = ""
    ROOMCHARGECHECKIN_TXT.Text = ""
    ROOMNOCHECKIN_TXT.Text = ""
    BILLSTATUSCHECKIN_TXT.Text = ""
    GNAMECHECKIN_TXT.Text = ""
    GAGECHECKIN_TXT.Text = ""
    GIDCARDNOCHECKIN_TXT.Text = ""
    GPROFESSIONCHECKIN_TXT.Text = ""
    GCONTACTCHECKIN_TXT.Text = ""
    GEMAILCHECKIN_TXT.Text = ""
    GADDRESSCHECKIN_TXT.Text = ""
    GGENDERCHECKIN_TXT.Text = ""
'    CMBRESERVIDCARDTYPE.Text = ""
    SAVECHECKIN_CMD.Enabled = False
    ADDNEWCHECKIN_CMD.Enabled = True
    SEARCHCHECKIN_CMD.Enabled = True
    ROOMTYPECHECKIN_CMB.Clear
Else
    End If
End If
End Sub

Private Sub SEARCHCHECKIN_CMD_Click()
'Label2.Visible = True
'Label2.Caption = ""
GUESTIDCHEKIN_CMB.Clear
GIDCHECKIN_TXT.Visible = False
GUESTIDCHEKIN_CMB.Visible = True
GIDCARDTYPECHECKIN_TXT.Visible = True
CMBRESERVIDCARDTYPE.Visible = False
'CMBRESERVIDCARDTYPE.Text = ""
GIDCARDTYPECHECKIN_TXT.Text = ""
S = "SELECT *FROM CLIENT_MASTER"
Set R = C.Execute(S)
Do Until R.EOF = True
    GUESTIDCHEKIN_CMB.AddItem R.Fields("CLIENT_ID")
    R.MoveNext
Loop
ROOMNOCHECKIN_CMB.Refresh
MALE_OPT.Refresh
FEMALE_OPT.Refresh
ADDNEWCHECKIN_CMD.Enabled = False
SEARCHCHECKIN_CMD.Enabled = False
'GIDCHECKIN_TXT.Enabled = True
'GOCHECKIN_CMD.Visible = True
'GIDCHECKIN_TXT.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
CHECKINTIME_LBL.Caption = Time
RESERVDATE_LBL.Caption = Date
End Sub

Private Sub UPDATECHECKIN_CMD_Click()
If ROOMTYPECHECKIN_TXT.Text = "" Then
MsgBox "PLEASE ENTER TYPE OF ROOM"
ROOMTYPECHECKIN_TXT.SetFocus
ElseIf ROOMNOCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE ALLOWED THE ROOM NO"
ROOMNOCHECKIN_TXT.SetFocus
ElseIf GNAMECHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT NAME OF GUEST"
GNAMECHECKIN_TXT.SetFocus
ElseIf GAGECHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT AGE OF GUEST"
GAGECHECKIN_TXT.SetFocus
ElseIf GIDCARDNOCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT ID NO"
GIDCARDNOCHECKIN_TXT.SetFocus
ElseIf GPROFESSIONCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT THE PROFESSION OF GUEST"
GPROFESSIONCHECKIN_TXT.SetFocus
ElseIf GCONTACTCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT CONTACT NO"
GCONTACTCHECKIN_TXT.SetFocus
ElseIf GADDRESSCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT ADDRESS"
GADDRESSCHECKIN_TXT.SetFocus
ElseIf GGENDERCHECKIN_TXT.Text = "" Then
MsgBox "PLEASE INPUT GENDER"
GGENDERCHECKIN_TXT.SetFocus
Else
If MsgBox("ARE YOU SURE TO UPDATE THE ROOM CHECKIN", vbYesNo) = vbYes Then
CON
S = "UPDATE  CLIENT_MASTER SET NAME='" & GNAMECHECKIN_TXT.Text & "',AGE=" & GAGECHECKIN_TXT.Text & ",SEX='" & GGENDERCHECKIN_TXT.Text & "',PROFESSION='" & GPROFESSIONCHECKIN_TXT.Text & "',ADDRESS='" & GADDRESSCHECKIN_TXT.Text & "',ID_CARD_NO='" & GIDCARDNOCHECKIN_TXT.Text & "',CONTACT=" & GCONTACTCHECKIN_TXT.Text & ",EMAIL='" & GEMAILCHECKIN_TXT.Text & "',CHECK_IN_TIME='" & CHECKINTIME_TXT.Text & "'WHERE CLIENT_ID='" & GUESTIDCHEKIN_CMB.Text & "' "
MsgBox S
Set R = C.Execute(S)
'S = "UPDATE GCODE SET G=" & NOOFGUEST & " "
'Set R = C.Execute(S)
'BB = "BOOKED"
'S = "UPDATE ROOM SET STATUS='" & BB & "' WHERE R_NO=" & ROOMNOCHECKIN_TXT.Text & ""
'Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "THE UPDATION IS FINISHED"
ROOMCHECKIN_FRAME.Enabled = False
GUESTDETAIL_FRAME.Enabled = False
'GOCHECKIN_CMD.Visible = False
GIDCHECKIN_TXT.Enabled = False
GIDCHECKIN_TXT.Text = ""
ROOMTYPECHECKIN_TXT.Text = ""
ROOMCHARGECHECKIN_TXT.Text = ""
ROOMNOCHECKIN_TXT.Text = ""
BILLSTATUSCHECKIN_TXT.Text = ""
GNAMECHECKIN_TXT.Text = ""
GAGECHECKIN_TXT.Text = ""
GIDCARDNOCHECKIN_TXT.Text = ""
GPROFESSIONCHECKIN_TXT.Text = ""
GCONTACTCHECKIN_TXT.Text = ""
GEMAILCHECKIN_TXT.Text = ""
GADDRESSCHECKIN_TXT.Text = ""
GGENDERCHECKIN_TXT.Text = ""
CHECKINTIME_TXT.Text = ""
'CMBRESERVIDCARDTYPE.Text = ""
GIDCARDTYPECHECKIN_TXT.Text = ""
CHECKINTIME_TXT.Visible = False
CHEKINTIME_LBL.Visible = False
CHECKOUTDATE_LBL.Visible = False
CHECKOUTDATE_TXT.Visible = False
UPDATECHECKIN_CMD.Enabled = False
ADDNEWCHECKIN_CMD.Enabled = True
SEARCHCHECKIN_CMD.Enabled = True
ROOMTYPECHECKIN_CMB.Clear
Else
ROOMCHECKIN_FRAME.Enabled = False
GUESTDETAIL_FRAME.Enabled = False
GENDERCHECKIN_FRAME.Enabled = False
UPDATECHECKIN_CMD.Enabled = False
ADDNEWCHECKIN_CMD.Enabled = True
SEARCHCHECKIN_CMD.Enabled = True
'GOCHECKIN_CMD.Visible = False
GIDCHECKIN_TXT.Enabled = False
GIDCARDTYPECHECKIN_TXT.Visible = False
CMBRESERVIDCARDTYPE.Visible = True
'CMBRESERVIDCARDTYPE.Text = ""
GIDCARDTYPECHECKIN_TXT.Text = ""
GIDCHECKIN_TXT.Text = ""
ROOMTYPECHECKIN_TXT.Text = ""
ROOMCHARGECHECKIN_TXT.Text = ""
ROOMNOCHECKIN_TXT.Text = ""
BILLSTATUSCHECKIN_TXT.Text = ""
GNAMECHECKIN_TXT.Text = ""
GAGECHECKIN_TXT.Text = ""
GIDCARDNOCHECKIN_TXT.Text = ""
GPROFESSIONCHECKIN_TXT.Text = ""
GCONTACTCHECKIN_TXT.Text = ""
GEMAILCHECKIN_TXT.Text = ""
GADDRESSCHECKIN_TXT.Text = ""
GGENDERCHECKIN_TXT.Text = ""
CHECKINTIME_TXT.Visible = False
CHEKINTIME_LBL.Visible = False
CHECKOUTDATE_LBL.Visible = False
CHECKOUTDATE_TXT.Visible = False
ROOMTYPECHECKIN_CMB.Clear
End If
End If
End Sub
