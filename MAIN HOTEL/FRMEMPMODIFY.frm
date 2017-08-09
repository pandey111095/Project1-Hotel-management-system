VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMEMPMODIFY 
   Caption         =   "MODIFY OF RMPLOYEE DETAIL"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   Picture         =   "FRMEMPMODIFY.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SEARCHOK_CMD 
      BackColor       =   &H00FFFF80&
      Caption         =   "OK"
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
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton BACKEMPTOMDIFORM_CMD 
      BackColor       =   &H0080FF80&
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
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8640
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
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8640
      Width           =   1575
   End
   Begin VB.TextBox EMPNAME_TXT 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   11
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox EMPDESIG_TXT 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox EMPDEPARTMENT_TXT 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox EMPDOB_TXT 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   8
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox EMPDOA_TXT 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   7
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox EMPADDRESS_TXT 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1935
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox EMPSALARY_TXT 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   5
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox EMPEXPERIENCE_TXT 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2760
      TabIndex        =   4
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox EMPID_TXT 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton DELETEEMP_CMD 
      BackColor       =   &H0080FF80&
      Caption         =   "DELETE"
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
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton MODIFYEMP_CMD 
      BackColor       =   &H0080FF80&
      Caption         =   "MODIFY"
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
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7680
      Top             =   7680
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
      Bindings        =   "FRMEMPMODIFY.frx":1B4F3
      Height          =   5655
      Left            =   8640
      TabIndex        =   2
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9975
      _Version        =   393216
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
   Begin VB.Label EMPNAME_LBL 
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
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DESIGNATION"
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
      Left            =   480
      TabIndex        =   19
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENTT"
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
      Left            =   480
      TabIndex        =   18
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
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
      Left            =   480
      TabIndex        =   17
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DOA"
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
      Left            =   480
      TabIndex        =   16
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
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
      Left            =   480
      TabIndex        =   15
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY"
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
      Left            =   480
      TabIndex        =   14
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "EXPERIENCE"
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
      Left            =   480
      TabIndex        =   13
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EMP_ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "FRMEMPMODIFY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EMPID_TXT_Change()
If Len(EMPID_TXT.Text) > 0 Then
SEARCHOK_CMD.Enabled = True
End If
End Sub

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

Private Sub MODIFYEMP_CMD_Click()
If MsgBox("WHAT DO YOU WANT IN THIS RECORD ?   'DELETE?' THEN CLICK ON 'YES'  OTHERWISE  ANY MODIFY CLICK ON   'NO'", vbYesNoCancel) = vbYes Then
MODIFYEMP_CMD.Enabled = False
DELETEEMP_CMD.Enabled = True
Else
UPDATEEMP_CMD.Enabled = True
EMPNAME_TXT.Enabled = True
EMPDESIG_TXT.Enabled = True
EMPDEPARTMENT_TXT.Enabled = True
EMPDOB_TXT.Enabled = True
EMPDOA_TXT.Enabled = True
EMPSALARY_TXT.Enabled = True
EMPEXPERIENCE_TXT.Enabled = True
EMPADDRESS_TXT.Enabled = True
EMPNAME_TXT.SetFocus
MODIFYEMP_CMD.Enabled = False
End If
End Sub

Private Sub SEARCHOK_CMD_Click()
CON
S = "SELECT *FROM EMP_RECORD WHERE EMP_ID='" & EMPID_TXT.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
    MsgBox "WRONG EMPLOYEE ID", vbCritical
    If MsgBox("ARE YOU WANT TO RE-ENTER THE EMPLOYEE ID THEN CLICK  'YES' ELSE  'NO'", vbYesNo) = vbYes Then
        EMPID_TXT.SetFocus
    Else
        BACKEMPTOMDIFORM_CMD.SetFocus
    End If
Else
    EMPNAME_TXT.Text = R.Fields("EMP_NAME")
    EMPDESIG_TXT.Text = R.Fields("DESIGNATION")
    EMPDEPARTMENT_TXT.Text = R.Fields("DEPARTMENT")
    EMPDOB_TXT.Text = R.Fields("DOB")
    EMPDOA_TXT.Text = R.Fields("DOA")
    EMPSALARY_TXT.Text = R.Fields("SALARY")
    EMPEXPERIENCE_TXT.Text = R.Fields("EXPERIENCE")
    EMPADDRESS_TXT.Text = R.Fields("ADDRESS")
    BACKEMPTOMDIFORM_CMD.SetFocus
    MODIFYEMP_CMD.Enabled = True
End If
End Sub

