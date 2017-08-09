VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMEMPENTRY 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMPLOYEE ENTRY FORM"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   17415
   ForeColor       =   &H00FFFF80&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMEMP.frx":0000
   ScaleHeight     =   7005
   ScaleWidth      =   17415
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox EMPID_CMB 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox EMPID_TXT 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
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
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   7680
      MaxLength       =   3
      TabIndex        =   32
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton DELETEEMP_CMD 
      BackColor       =   &H0080FF80&
      Caption         =   "DELETE"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton UPDATEEMP_CMD 
      BackColor       =   &H0080FF80&
      Caption         =   "UPDATE"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton BACKEMPTOMDIFORM_CMD 
      BackColor       =   &H000000FF&
      Caption         =   "BACK"
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame EMPINHOTEL_FRAME 
      BackColor       =   &H8000000A&
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
      Height          =   2055
      Left            =   6960
      TabIndex        =   25
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox EMPSALARY_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox EMPDEPART_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   18
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox EMPDESIG_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   18
         TabIndex        =   8
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   28
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   27
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   26
         Top             =   480
         Width           =   1665
      End
   End
   Begin VB.Frame EMPDETAIL_FRAME 
      BackColor       =   &H8000000A&
      Caption         =   "EMP DETAIL"
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
      Height          =   5775
      Left            =   720
      TabIndex        =   16
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox EMPIDCARDNO_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   16
         TabIndex        =   1
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox EMPCONTACT_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   11
         TabIndex        =   2
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox EMPEMAIL_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   5
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox EMPNAME_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox EMPADDRESS_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   2760
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3720
         Width           =   3255
      End
      Begin VB.TextBox EMPEXPERIENCE_TXT 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   6
         Top             =   3240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker EMPDOB_DTP 
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96534529
         CurrentDate     =   41977
      End
      Begin MSComCtl2.DTPicker EMPDOA_DTP 
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96534529
         CurrentDate     =   41977
      End
      Begin VB.Label IDCARDNO_LBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id card no"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label EMPCONTACT_LBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   23
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   22
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label EMPNAME_LBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   21
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   20
         Top             =   1800
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of assign"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   19
         Top             =   2280
         Width           =   1980
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   18
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experience"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   17
         Top             =   3240
         Width           =   1635
      End
   End
   Begin VB.CommandButton SAVEEMP_CMD 
      BackColor       =   &H0080FF80&
      Caption         =   "SAVE"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   8040
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=SANDEEP/KOHLI;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=SANDEEP/KOHLI;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT  *FROM  EMP_RECORD"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FRMEMP.frx":1B4F3
      Height          =   9855
      Left            =   14040
      TabIndex        =   14
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   17383
      _Version        =   393216
      BackColor       =   -2147483646
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton BACKTOMDIFROMEMPENTRY_CMD 
      BackColor       =   &H000000FF&
      Caption         =   "BACK"
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton ADDNEWEMP_CMD 
      BackColor       =   &H0080FF80&
      Caption         =   "ADD NEW"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   15
      Height          =   375
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   15
      Height          =   375
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label EMPLOID_LBL 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "EMP ID :-  "
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
      Left            =   6000
      TabIndex        =   15
      Top             =   480
      Width           =   1875
   End
End
Attribute VB_Name = "FRMEMPENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AA
Private Sub ADDNEWEMP_CMD_Click()
EMPDETAIL_FRAME.Enabled = True
EMPINHOTEL_FRAME.Enabled = True
CON
S = "SELECT (E) FROM ECODE"
Set R = C.Execute(S)
NOOFEMP = R.Fields("E")
NOOFEMP = NOOFEMP + 1
EMPID_TXT.Text = "E" & NOOFEMP
EMPNAME_TXT.SetFocus
SAVEEMP_CMD.Enabled = True
ADDNEWEMP_CMD.Enabled = False
End Sub

Private Sub BACKEMPTOMDIFORM_CMD_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub BACKTOMDIFROMEMPENTRY_CMD_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub DELETEEMP_CMD_Click()
If Len(EMPNAME_TXT.Text) > 0 Then
If MsgBox("ARE  YOU  SURE TO DELETE RECORD ?", vbYesNo) = vbYes Then
CON
S = "INSERT INTO DEL_EMP_RECORD VALUES('" & EMPID_CMB.Text & "','" & EMPNAME_TXT.Text & "','" & Format(EMPDOA_DTP.Value, "DD-MMM-YYYY") & "','" & Format(Date, "DD-MMM-YYYY") & "','" & EMPADDRESS_TXT.Text & "'," & EMPCONTACT_TXT.Text & ",'" & EMPIDCARDNO_TXT.Text & "')"
Set R = C.Execute(S)
S = "DELETE FROM EMP_RECORD WHERE EMP_ID='" & EMPID_CMB.Text & "'"
Set R = C.Execute(S)
S = "DELETE FROM EMP_ATTENDENCE_DETAIL WHERE EMP_ID='" & EMPID_CMB.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
Adodc1.Refresh
EMPID_CMB.Clear
S = "SELECT EMP_ID FROM EMP_RECORD"
Set R = C.Execute(S)
Do Until R.EOF = True
EMPID_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
MsgBox "RECORD IS DELETED"
EMPNAME_TXT.Text = ""
EMPDESIG_TXT.Text = ""
EMPDEPART_TXT.Text = ""
'EMPID_CMB.Text = ""
EMPADDRESS_TXT.Text = ""
EMPEMAIL_TXT.Text = ""
EMPCONTACT_TXT.Text = ""
EMPIDCARDNO_TXT.Text = ""
EMPSALARY_TXT.Text = ""
EMPEXPERIENCE_TXT.Text = ""
'EMPID_TXT.Enabled = True
'DELETEEMP_CMD.Enabled = False
BACKEMPTOMDIFORM_CMD.SetFocus
Else
EMPNAME_TXT.Text = ""
EMPDESIG_TXT.Text = ""
EMPDEPART_TXT.Text = ""
'EMPID_CMB.Text = ""
EMPADDRESS_TXT.Text = ""
EMPEMAIL_TXT.Text = ""
EMPCONTACT_TXT.Text = ""
EMPIDCARDNO_TXT.Text = ""
EMPSALARY_TXT.Text = ""
EMPEXPERIENCE_TXT.Text = ""

'EMPID_TXT.Enabled = True
'DELETEEMP_CMD.Enabled = False
BACKEMPTOMDIFORM_CMD.SetFocus
End If
Else
MsgBox "PLEASE SELECT EMPLOYEE ID"
EMPID_CMB.SetFocus
End If
End Sub

Private Sub EMPADDRESS_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
End Sub

Private Sub EMPCONTACT_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    EMPDOB_DTP.SetFocus
End If
End Sub

Private Sub EMPDESIG_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    EMPDEPART_TXT.SetFocus
End If
End Sub
Private Sub EMPDEPART_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
EMPSALARY_TXT.SetFocus
End If
End Sub





Private Sub EMPDOB_DTP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    EMPDOA_DTP.SetFocus
End If
End Sub

Private Sub EMPDOB_DTP_LostFocus()
DATE1 = EMPDOB_DTP.Value
DATE2 = Date
DATE3 = DateDiff("YYYY", DATE1, DATE2)
If DATE3 >= 18 Then
    EMPDOA_DTP.SetFocus
Else
    MsgBox "PLEASE RE-ENTER THE DOB, BECOUSE YOUR AGE IS LESS THAN 18"
    EMPDOB_DTP.SetFocus
End If
End Sub

Private Sub EMPEMAIL_TXT_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    EMPEXPERIENCE_TXT.SetFocus
End If
End Sub



Private Sub EMPID_CMB_Click()
CON
S = "SELECT *FROM EMP_RECORD WHERE EMP_ID='" & EMPID_CMB.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
    MsgBox "WRONG EMPLOYEE ID", vbCritical
    If MsgBox("ARE YOU WANT TO RE-ENTER THE EMPLOYEE ID ?'", vbYesNo) = vbYes Then
'        EMPID_CMB.Text = ""
        EMPNAME_TXT.Text = ""
        EMPDESIG_TXT.Text = ""
        EMPDEPART_TXT.Text = ""
        EMPID_TXT.Text = ""
        EMPADDRESS_TXT.Text = ""
        EMPEMAIL_TXT.Text = ""
        EMPCONTACT_TXT.Text = ""
        EMPIDCARDNO_TXT.Text = ""
        EMPSALARY_TXT.Text = ""
        EMPEXPERIENCE_TXT.Text = ""
        EMPID_CMB.SetFocus
    Else
'        EMPID_CMB.Text = ""
        EMPNAME_TXT.Text = ""
        EMPDESIG_TXT.Text = ""
        EMPDEPART_TXT.Text = ""
        EMPID_TXT.Text = ""
        EMPADDRESS_TXT.Text = ""
        EMPEMAIL_TXT.Text = ""
        EMPCONTACT_TXT.Text = ""
        EMPIDCARDNO_TXT.Text = ""
        EMPSALARY_TXT.Text = ""
        EMPEXPERIENCE_TXT.Text = ""
        BACKEMPTOMDIFORM_CMD.SetFocus
    End If
Else
    EMPNAME_TXT.Text = R.Fields("EMP_NAME")
    EMPDESIG_TXT.Text = R.Fields("DESIGNATION")
    EMPDEPART_TXT.Text = R.Fields("DEPARTMENT")
    EMPDOB_DTP.Value = R.Fields("DOB")
    EMPDOA_DTP.Value = R.Fields("DOA")
    EMPSALARY_TXT.Text = R.Fields("SALARY")
    EMPEXPERIENCE_TXT.Text = R.Fields("EXPERIENCE")
    EMPCONTACT_TXT.Text = R.Fields("CONTACT")
    'GEMAILCHECKIN_TXT.Text = IIf(IsNull(R.Fields("EMAIL")), "", R.Fields("EMAIL"))
    EMPEMAIL_TXT.Text = IIf(IsNull(R.Fields("EMAIL")), "", R.Fields("EMAIL"))
    EMPIDCARDNO_TXT.Text = R.Fields("IDCARDNO")
    EMPADDRESS_TXT.Text = R.Fields("ADDRESS")
    EMPDETAIL_FRAME.Enabled = True
    EMPINHOTEL_FRAME.Enabled = True
    UPDATEEMP_CMD.Enabled = True
    BACKEMPTOMDIFORM_CMD.SetFocus
    'MODIFYEMP_CMD.Enabled = True
    'SEARCHOK_CMD.Enabled = False
End If

End Sub

'Private Sub EMPID_TXT_Change()
'If Len(EMPID_TXT.Text) > 0 Then
'SEARCHOK_CMD.Enabled = True
'Else
'SEARCHOK_CMD.Enabled = False
'End If
'End Sub
Private Sub EMPID_TXT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    EMPID_TXT.Text = UCase(EMPID_TXT.Text)
    SEARCHOK_CMD.SetFocus
End If
End Sub

Private Sub EMPID_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    EMPID_TXT.Text = UCase(EMPID_TXT.Text)
    SEARCHOK_CMD.SetFocus
End If
End Sub

Private Sub EMPID_TXT_LostFocus()
     EMPID_TXT.Text = UCase(EMPID_TXT.Text)
End Sub

Private Sub EMPNAME_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    EMPIDCARDNO_TXT.SetFocus
End If
End Sub
Private Sub EMPIDCARDNO_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    EMPCONTACT_TXT.SetFocus
End If
End Sub

Private Sub EMPCONTACT_TXT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    EMPDOB_DTP.SetFocus
End If
End Sub
Private Sub EMPSALARY_TXT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
If UPDATEEMP_CMD.Visible = True Then
    UPDATEEMP_CMD.SetFocus
Else
    SAVEEMP_CMD.SetFocus
End If
End If
End Sub

Private Sub EMPSALARY_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
If UPDATEEMP_CMD.Visible = True Then
    UPDATEEMP_CMD.SetFocus
Else
    SAVEEMP_CMD.SetFocus
    End If
End If
End Sub
Private Sub EMPEXPERIENCE_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    EMPADDRESS_TXT.SetFocus
End If
End Sub

'Private Sub EMPSALARY_TXT_Validate(Cancel As Boolean)
'If Not IsNumeric(EMPSALARY_TXT.Text) Then
'Cancel = True
'MsgBox "PLEASE INPUT SALARY OF EMPLOYEE AS NUMERIC", vbCritical
'EMPSALARY_TXT.Text = ""
'EMPSALARY_TXT.SetFocus
'Else
'MsgBox "IT IS NUMERIC"
'Cancel = False
'EMPEXPERIENCE_TXT.SetFocus
'End If
Private Sub Form_Load()
EMPDOB_DTP.Value = Now
EMPDOA_DTP.Value = Now
CON
S = "SELECT COUNT (*)FROM ECODE"
Set R = C.Execute(S)
D = R.Fields(0)
If D = 0 Then
'S = "INSERT INTO ECODE VALUES(" + 0 + ")"
S = "INSERT INTO ECODE VALUES (" & 0 & ")"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
Else
End If
SAVEEMP_CMD.Enabled = False
'Me.Height = MDIForm1.Height - 20
'Me.Width = MDIForm1.Width - 50
End Sub

'Private Sub MODIFYEMP_CMD_Click()
'AA = MsgBox("WHAT DO YOU WANT IN THIS RECORD ?   'DELETE?' THEN CLICK ON 'YES'  OTHERWISE  ANY MODIFY CLICK ON   'NO'", vbYesNoCancel)
'If AA = vbYes Then
 '   EMPID_TXT.Enabled = False
  '  MODIFYEMP_CMD.Enabled = False
   ' DELETEEMP_CMD.Enabled = True
    'DELETEEMP_CMD.SetFocus
'ElseIf AA = vbNo Then
 '   EMPID_TXT.Enabled = False
  '  EMPDETAIL_FRAME.Enabled = True
   ' EMPINHOTEL_FRAME.Enabled = True
    'MODIFYEMP_CMD.Enabled = False
    'UPDATEEMP_CMD.Enabled = True
'Else
 '   MODIFYEMP_CMD.Enabled = False
  '  EMPID_TXT.Text = ""
   ' EMPNAME_TXT.Text = ""
    'EMPDESIG_TXT.Text = ""
    'EMPDEPART_TXT.Text = ""
    'EMPADDRESS_TXT.Text = ""
    'EMPEMAIL_TXT.Text = ""
    'EMPCONTACT_TXT.Text = ""
    'EMPIDCARDNO_TXT.Text = ""
    'EMPSALARY_TXT.Text = ""
    'EMPEXPERIENCE_TXT.Text = ""
    'EMPID_TXT.SetFocus
'End If
'End Sub

Private Sub SAVEEMP_CMD_Click()
If EMPNAME_TXT.Text = "" Then
MsgBox "PLESE INPUT NAME OF EMPLOYEE"
EMPNAME_TXT.SetFocus
ElseIf EMPIDCARDNO_TXT.Text = "" Then
MsgBox "PLESE INPUT IDCARDNO OF EMPLOYEE"
EMPIDCARDNO_TXT.SetFocus
ElseIf EMPCONTACT_TXT.Text = "" Then
MsgBox "PLESE INPUT CONTACT NO OF EMPLOYEE"
EMPCONTACT_TXT.SetFocus
ElseIf EMPDOB_DTP.Value = "" Then
MsgBox "PLESE INPUT DATE OF BIRTH OF EMPLOYEE"
EMPDOB_DTP.SetFocus
ElseIf EMPDOA_DTP.Value = "" Then
MsgBox "PLESE INPUT DATE OF ASSIGN JOB OF EMPLOYEE"
EMPDOA_DTP.SetFocus
ElseIf EMPEXPERIENCE_TXT.Text = "" Then
MsgBox "PLESE INPUT EXPERIENCE OF EMPLOYEE"
EMPEXPERIENCE_TXT.SetFocus
ElseIf EMPADDRESS_TXT.Text = "" Then
MsgBox "PLESE INPUT ADDRESS OF EMPLOYEE"
EMPADDRESS_TXT.SetFocus
ElseIf EMPDESIG_TXT.Text = "" Then
MsgBox "PLESE INPUT DESIGNATION OF EMPLOYEE"
EMPDESIG_TXT.SetFocus
ElseIf EMPDEPART_TXT.Text = "" Then
MsgBox "PLESE INPUT DEPARTMENT OF EMPLOYEE"
EMPDEPART_TXT.SetFocus
ElseIf EMPSALARY_TXT.Text = "" Then
MsgBox "PLESE INPUT SALARY OF EMPLOYEE"
EMPSALARY_TXT.SetFocus
ElseIf MsgBox("ARE SURE TO ADD THE EMPLOYEE IN YOUR HOTEL", vbYesNo) = vbYes Then
CON
S = "INSERT INTO EMP_RECORD VALUES('" & EMPID_TXT.Text & "','" & EMPNAME_TXT.Text & "','" & EMPDESIG_TXT.Text & "','" & EMPDEPART_TXT.Text & "','" & Format(EMPDOB_DTP.Value, "DD-MMM-YYYY") & "','" & Format(EMPDOA_DTP.Value, "DD-MMM-YYYY") & "','" & EMPADDRESS_TXT.Text & "'," & EMPSALARY_TXT.Text & "," & EMPEXPERIENCE_TXT.Text & "," & EMPCONTACT_TXT.Text & ",'" & EMPEMAIL_TXT.Text & "','" & EMPIDCARDNO_TXT.Text & "')"
Set R = C.Execute(S)
S = "INSERT INTO EMP_ATTENDENCE_DETAIL VALUES('" & EMPID_TXT.Text & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ")"
MsgBox S
Set R = C.Execute(S)
S = "UPDATE ECODE SET E=" & NOOFEMP & " "
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "THE EMPLOYEE IS NOW ADDED "
Adodc1.Refresh
EMPNAME_TXT.Text = ""
EMPDESIG_TXT.Text = ""
EMPDEPART_TXT.Text = ""
EMPID_TXT.Text = ""
EMPADDRESS_TXT.Text = ""
EMPEMAIL_TXT.Text = ""
EMPCONTACT_TXT.Text = ""
EMPIDCARDNO_TXT.Text = ""
EMPSALARY_TXT.Text = ""
EMPEXPERIENCE_TXT.Text = ""
BACKTOMDIFROMEMPENTRY_CMD.SetFocus
SAVEEMP_CMD.Enabled = False
ADDNEWEMP_CMD.Enabled = True
Else
EMPNAME_TXT.Text = ""
EMPDESIG_TXT.Text = ""
EMPDEPART_TXT.Text = ""
EMPEMAIL_TXT.Text = ""
EMPCONTACT_TXT.Text = ""
EMPIDCARDNO_TXT.Text = ""
EMPID_TXT.Text = ""
EMPADDRESS_TXT.Text = ""
EMPSALARY_TXT.Text = ""
EMPEXPERIENCE_TXT.Text = ""
BACKTOMDIFROMEMPENTRY_CMD.SetFocus
SAVEEMP_CMD.Enabled = False
ADDNEWEMP_CMD.Enabled = True
End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Print KeyAscii
End Sub

Private Sub SEARCHOK_CMD_Click()
CON
S = "SELECT *FROM EMP_RECORD WHERE EMP_ID='" & EMPID_TXT.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
    MsgBox "WRONG EMPLOYEE ID", vbCritical
    If MsgBox("ARE YOU WANT TO RE-ENTER THE EMPLOYEE ID THEN CLICK  'YES' ELSE  'NO'", vbYesNo) = vbYes Then
        EMPID_TXT.Text = ""
        EMPNAME_TXT.Text = ""
        EMPDESIG_TXT.Text = ""
        EMPDEPART_TXT.Text = ""
        EMPID_TXT.Text = ""
        EMPADDRESS_TXT.Text = ""
        EMPEMAIL_TXT.Text = ""
        EMPCONTACT_TXT.Text = ""
        EMPIDCARDNO_TXT.Text = ""
        EMPSALARY_TXT.Text = ""
        EMPEXPERIENCE_TXT.Text = ""
        EMPID_TXT.SetFocus
    Else
        EMPID_TXT.Text = ""
        EMPNAME_TXT.Text = ""
        EMPDESIG_TXT.Text = ""
        EMPDEPART_TXT.Text = ""
        EMPID_TXT.Text = ""
        EMPADDRESS_TXT.Text = ""
        EMPEMAIL_TXT.Text = ""
        EMPCONTACT_TXT.Text = ""
        EMPIDCARDNO_TXT.Text = ""
        EMPSALARY_TXT.Text = ""
        EMPEXPERIENCE_TXT.Text = ""

        BACKEMPTOMDIFORM_CMD.SetFocus
    End If
Else
    EMPNAME_TXT.Text = R.Fields("EMP_NAME")
    EMPDESIG_TXT.Text = R.Fields("DESIGNATION")
    EMPDEPART_TXT.Text = R.Fields("DEPARTMENT")
    EMPDOB_DTP.Value = R.Fields("DOB")
    EMPDOA_DTP.Value = R.Fields("DOA")
    EMPSALARY_TXT.Text = R.Fields("SALARY")
    EMPEXPERIENCE_TXT.Text = R.Fields("EXPERIENCE")
    EMPCONTACT_TXT.Text = R.Fields("CONTACT")
    'GEMAILCHECKIN_TXT.Text = IIf(IsNull(R.Fields("EMAIL")), "", R.Fields("EMAIL"))
    EMPEMAIL_TXT.Text = IIf(IsNull(R.Fields("EMAIL")), "", R.Fields("EMAIL"))
    EMPIDCARDNO_TXT.Text = R.Fields("IDCARDNO")
    EMPADDRESS_TXT.Text = R.Fields("ADDRESS")
    BACKEMPTOMDIFORM_CMD.SetFocus
    MODIFYEMP_CMD.Enabled = True
    SEARCHOK_CMD.Enabled = False
End If
End Sub

Private Sub UPDATEEMP_CMD_Click()
If EMPNAME_TXT.Text = "" Then
MsgBox "PLESE INPUT NAME OF EMPLOYEE"
EMPNAME_TXT.SetFocus
ElseIf EMPIDCARDNO_TXT.Text = "" Then
MsgBox "PLESE INPUT IDCARDNO OF EMPLOYEE"
EMPIDCARDNO_TXT.SetFocus
ElseIf EMPCONTACT_TXT.Text = "" Then
MsgBox "PLESE INPUT CONTACT NO OF EMPLOYEE"
EMPCONTACT_TXT.SetFocus
ElseIf EMPDOB_DTP.Value = "" Then
MsgBox "PLESE INPUT DATE OF BIRTH OF EMPLOYEE"
EMPDOB_DTP.SetFocus
ElseIf EMPDOA_DTP.Value = "" Then
MsgBox "PLESE INPUT DATE OF ASSIGN JOB OF EMPLOYEE"
EMPDOA_DTP.SetFocus
ElseIf EMPEXPERIENCE_TXT.Text = "" Then
MsgBox "PLESE INPUT EXPERIENCE OF EMPLOYEE"
EMPEXPERIENCE_TXT.SetFocus
ElseIf EMPADDRESS_TXT.Text = "" Then
MsgBox "PLESE INPUT ADDRESS OF EMPLOYEE"
EMPADDRESS_TXT.SetFocus
ElseIf EMPDESIG_TXT.Text = "" Then
MsgBox "PLESE INPUT DESIGNATION OF EMPLOYEE"
EMPDESIG_TXT.SetFocus
ElseIf EMPDEPART_TXT.Text = "" Then
MsgBox "PLESE INPUT DEPARTMENT OF EMPLOYEE"
EMPDEPART_TXT.SetFocus
ElseIf EMPSALARY_TXT.Text = "" Then
MsgBox "PLESE INPUT SALARY OF EMPLOYEE"
EMSALARY_TXT.SetFocus
ElseIf MsgBox("ARE SURE TO UPDATE THE EMPLOYEE INFORMATION ?", vbYesNo) = vbYes Then
CON
S = "UPDATE EMP_RECORD SET EMP_NAME='" & EMPNAME_TXT.Text & "', DESIGNATION='" & EMPDESIG_TXT.Text & "',DEPARTMENT='" & EMPDEPART_TXT.Text & "',DOB='" & Format(EMPDOB_DTP.Value, "DD-MMM-YYYY") & "',DOA='" & Format(EMPDOA_DTP.Value, "DD-MMM-YYYY") & "',ADDRESS='" & EMPADDRESS_TXT.Text & "' ,SALARY= " & EMPSALARY_TXT.Text & ",EXPERIENCE=" & EMPEXPERIENCE_TXT.Text & ",CONTACT=" & EMPCONTACT_TXT.Text & ",EMAIL='" & EMPEMAIL_TXT.Text & "',IDCARDNO='" & EMPIDCARDNO_TXT.Text & "' WHERE EMP_ID= '" & EMPID_CMB.Text & "'"
'MsgBox "UPDATE EMP_RECORD SET EMP_NAME='" & EMPNAME_TXT.Text & "', DESIGNATION='" & EMPDESIG_TXT.Text & "',DEPARTMENT='" & EMPDEPART_TXT.Text & "',DOB='" & Format(EMPDOB_DTP.Value, "DD-MMM-YYYY") & "',DOA='" & Format(EMPDOA_DTP.Value, "DD-MMM-YYYY") & "',ADDRESS='" & EMPADDRESS_TXT.Text & "' ,SALARY= " & EMPSALARY_TXT.Text & ",EXPERIENCE=" & EMPEXPERIENCE_TXT.Text & ",CONTACT=" & EMPCONTACT_TXT.Text & ",EMAIL='" & EMPEMAIL_TXT.Text & "',IDCARDNO='" & EMPIDCARDNO_TXT.Text & "' WHERE EMP_ID= '" & EMPID_TXT.Text & "'"
Set R = C.Execute(S)
MsgBox "EMPLOYEE INFORMATION IS UPDATED"
S = "COMMIT"
Set R = C.Execute(S)
Adodc1.Refresh
UPDATEEMP_CMD.Enabled = False
'EMPID_CMB.Text = ""
'EMPID_TXT.Enabled = True
EMPID_CMB.SetFocus
EMPNAME_TXT.Text = ""
EMPDESIG_TXT.Text = ""
EMPDEPART_TXT.Text = ""
EMPADDRESS_TXT.Text = ""
EMPEMAIL_TXT.Text = ""
EMPCONTACT_TXT.Text = ""
EMPIDCARDNO_TXT.Text = ""
EMPSALARY_TXT.Text = ""
EMPEXPERIENCE_TXT.Text = ""
End If
End Sub
