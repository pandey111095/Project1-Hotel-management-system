VERSION 5.00
Begin VB.Form FRMATTENDENCE 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ATTENDANCE"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11010
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMATTENDENCE.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton EMPATTENGO_OPT 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GO"
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton ATTENMAINCOME_OPT 
      BackColor       =   &H00FFC0C0&
      Caption         =   "COME"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton ATAINMAINGO_OPT 
      BackColor       =   &H00FFC0C0&
      Caption         =   "GO"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox EMPIDATTENGO_CMB 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton CMDBACKFROMATTEN 
      BackColor       =   &H000000FF&
      Caption         =   "BACK"
      Height          =   255
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame ATTEN_FRAME 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   5520
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton LEAVE_OPT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LEAVE"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton COME_OPT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COME"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label ATTEN_LBL 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ATTENDENCE"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10680
      Top             =   4080
   End
   Begin VB.ComboBox EMPIDATTEN_CMB 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   3240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   2880
      TabIndex        =   14
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   2880
      TabIndex        =   13
      Top             =   3360
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      BorderWidth     =   5
      Height          =   3135
      Left            =   2040
      Top             =   1680
      Width           =   6015
   End
   Begin VB.Label EMPJOB_LBL 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4200
      TabIndex        =   6
      Top             =   3720
      Width           =   45
   End
   Begin VB.Label EMPNAME_LBL 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   3360
      Width           =   45
   End
   Begin VB.Label ATTENTIME_LBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   10440
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.Label ATTENDATE_LBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FRMATTENDENCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ATTEN As Integer

Private Sub Combo1_Change()

End Sub

Private Sub ATAINMAINGO_OPT_Click()
EMPATTENGO_OPT.Visible = True
EMPIDATTENGO_CMB.Visible = True
EMPIDATTEN_CMB.Visible = False
EMPIDATTENGO_CMB.Clear
CON
S = "SELECT EMP_ID FROM EMP_ATTENDENCE WHERE ATTENDENCE='PERSENT' AND GOINGTIME='NOTEXIST' AND ADATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
Do Until R.EOF
EMPIDATTENGO_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
ATTEN_FRAME.Visible = False
LEAVE_OPT.Visible = False
End Sub

Private Sub ATTENMAINCOME_OPT_Click()
EMPATTENGO_OPT.Visible = False
LEAVE_OPT.Visible = True
EMPIDATTENGO_CMB.Visible = False
EMPIDATTEN_CMB.Visible = True
EMPIDATTEN_CMB.Clear
CON
S = "SELECT EMP_ID FROM EMP_ATTENDENCE WHERE ATTENDENCE='ABSENT' AND ADATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
Do Until R.EOF
EMPIDATTEN_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
EMPNAME_LBL.Visible = True
EMPJOB_LBL.Visible = True
ATTEN_FRAME.Visible = True
End Sub

Private Sub CMDBACKFROMATTEN_Click()
Unload Me
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub COME_OPT_Click()
ATTEN_LBL.Caption = "PERSENT"
'S = "UPDATE EMP_ATTENDENCE SET ATTENDENCE='PERSENT',COMMINGTIME='" & ATTENTIME_LBL.Caption & "' WHERE ADATE='" & Format(ATTENDATE_LBL.Caption, "DD-MMM-YYYY") & "' AND EMP_ID='" & EMPIDATTEN_CMB.Text & "'"
S = "UPDATE EMP_ATTENDENCE SET ATTENDENCE='PERSENT',COMMINGTIME='" & ATTENTIME_LBL.Caption & "'WHERE EMP_ID='" & EMPIDATTEN_CMB.Text & "' AND ADATE='" & Format(ATTENDATE_LBL.Caption, "DD-MMM-YYYY") & "'"
MsgBox S
Set R = C.Execute(S)
S = "UPDATE EMP_ATTENDENCE_DETAIL SET PERSENTING=PERSENTING+1 WHERE EMP_ID='" & EMPIDATTEN_CMB.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "YOU ARE PERSENT"
LEAVE_OPT.Enabled = False
Unload FRMATTENDENCE
End Sub

Private Sub EMPATTENGO_OPT_Click()
'S = "UPDATE EMP_ATTENDENCE SET ATTENDENCE='PERSENT',COMMINGTIME='" & ATTENTIME_LBL.Caption & "' WHERE ADATE='" & Format(ATTENDATE_LBL.Caption, "DD-MMM-YYYY") & "' AND EMP_ID='" & EMPIDATTEN_CMB.Text & "'"
S = "UPDATE EMP_ATTENDENCE SET GOINGTIME='" & ATTENTIME_LBL.Caption & "'WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "' AND ADATE='" & Format(ATTENDATE_LBL.Caption, "DD-MMM-YYYY") & "'"
MsgBox S
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "GOOD BY"
S = "SELECT COMMINGTIME FROM EMP_ATTENDENCE WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "' AND ADATE='" & Format(ATTENDATE_LBL.Caption, "DD-MMM-YYYY") & "'"
MsgBox S
Set R = C.Execute(S)
h1 = Hour(Format(R.Fields("COMMINGTIME"), "hh:mm:ss"))
m1 = Minute(Format(R.Fields("COMMINGTIME"), "hh:mm:ss"))
H2 = Hour(Format(ATTENTIME_LBL.Caption, "hh:mm:ss"))
m2 = Minute(Format(ATTENTIME_LBL.Caption, "hh:mm:ss"))
H = H2 - h1
M = m2 - m1
If H < 0 Then
S = "UPDATE EMP_ATTENDENCE_DETAIL SET HOURS_WORK=HOURS_WORK-" & H & " WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "' "
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
Else
S = "UPDATE EMP_ATTENDENCE_DETAIL SET HOURS_WORK=HOURS_WORK+" & H & " WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
End If
If M < 0 Then
S = "UPDATE EMP_ATTENDENCE_DETAIL SET MINUTE_WORK=MINUTE_WORK-" & M & " WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
Else
S = "UPDATE EMP_ATTENDENCE_DETAIL SET MINUTE_WORK=MINUTE_WORK+" & M & " WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
End If
S = "SELECT MINUTE_WORK FROM EMP_ATTENDENCE_DETAIL WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "'"
Set R = C.Execute(S)
If R.Fields("MINUTE_WORK") >= 60 Then
H = R.Fields("MINUTE_WORK") / 60
S = "UPDATE EMP_ATTENDENCE_DETAIL SET HOURS_WORK=HOURS_WORK+" & H & ""
Set R = C.Execute(S)
M = R.Fields("MINUTE_WORK") Mod 60
S = "UPDATE EMP_ATTENDENCE_DETAIL SET MINUTE_WORK=MINUTE_WORK+" & M & ""
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
End If
'FRMATTENDENCE.Hide
End Sub

Private Sub EMPIDATTEN_CMB_Click()
LEAVE_OPT.Enabled = True
ATTEN_FRAME.Enabled = True
CON
S = "SELECT EMP_NAME,DESIGNATION FROM EMP_RECORD WHERE EMP_ID='" & EMPIDATTEN_CMB.Text & "' "
Set R = C.Execute(S)
If R.EOF = True Then
MsgBox "NOT AN EMPLOYEE CODE"
Else
EMPNAME_LBL.Caption = R.Fields("EMP_NAME")
EMPJOB_LBL.Caption = R.Fields("DESIGNATION")
End If
ATTEN_FRAME.Enabled = True
End Sub

Private Sub EMPIDATTENGO_CMB_Click()
EMPATTENGO_OPT.Enabled = True
CON
S = "SELECT EMP_NAME,DESIGNATION FROM EMP_RECORD WHERE EMP_ID='" & EMPIDATTENGO_CMB.Text & "' "
Set R = C.Execute(S)
If R.EOF = True Then
MsgBox "NOT AN EMPLOYEE CODE"
Else
EMPNAME_LBL.Caption = R.Fields("EMP_NAME")
EMPJOB_LBL.Caption = R.Fields("DESIGNATION")
End If
End Sub

Private Sub Form_Load()
Me.Top = 1000
Me.Left = 5000
EMPATTENGO_OPT.Visible = False
LEAVE_OPT.Visible = True
EMPIDATTENGO_CMB.Visible = False
EMPIDATTEN_CMB.Visible = False
EMPNAME_LBL.Visible = True
EMPJOB_LBL.Visible = True
ATTEN_FRAME.Visible = False
CON
EMPIDATTEN_CMB.Clear
S = "SELECT DISTINCT(EMP_ID) FROM EMP_ATTENDENCE_DETAIL "
Set R = C.Execute(S)
Do Until R.EOF
EMPIDATTENGO_CMB.AddItem R.Fields("EMP_ID")
EMPIDATTEN_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
S = "SELECT COUNT(DISTINCT(EMP_ID)) FROM EMP_ATTENDENCE WHERE ADATE='" & Format(Date, "dd-mmm-yyyy") & "'"
Set R = C.Execute(S)
If R.Fields(0) > 0 Then
Else
If ATTEN = 0 Then
S = "SELECT DISTINCT(EMP_ID) FROM EMP_ATTENDENCE_DETAIL "
Set R = C.Execute(S)
Do Until R.EOF
SS = "INSERT INTO EMP_ATTENDENCE VALUES('" & R.Fields("EMP_ID") & "','" & Format(Date, "DD-MMM-YYYY") & "','" & "ABSENT" & "','" & "NOTEXIST" & "','" & "NOTEXIST" & "')"
Set RR = C.Execute(SS)
R.MoveNext
Loop
ATTEN = 1
End If
End If
End Sub

Private Sub LEAVE_OPT_Click()
ATTEN_LBL.Caption = "NOW HE IS ON LEAVE"
S = "UPDATE EMP_ATTENDENCE SET ATTENDENCE='LEAVE' WHERE ADATE='" & Format(ATTENDATE_LBL.Caption, "DD-MMM-YYYY") & "' AND EMP_ID='" & EMPIDATTEN_CMB.Text & "' "
Set R = C.Execute(S)
S = "UPDATE EMP_ATTENDENCE_DETAIL SET LEAVES=LEAVES+1 WHERE EMP_ID='" & EMPIDATTEN_CMB.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
LEAVE_OPT.Enabled = False
Unload FRMATTENDENCE
End Sub


Private Sub Timer1_Timer()
ATTENTIME_LBL.Caption = Time
ATTENDATE_LBL.Caption = Date
End Sub

