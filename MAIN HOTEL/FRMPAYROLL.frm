VERSION 5.00
Begin VB.Form FRMPAYROLL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton PAID_CMD 
      Caption         =   "PAID"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   35
      ToolTipText     =   "Save New Payroll Record"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton BACK_CMD 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   34
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   600
      TabIndex        =   18
      Top             =   1200
      Width           =   6855
      Begin VB.ComboBox EMPID_CMB 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "EMPLOYEE ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "EMPLOYEE NAME"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label EMPNAME_LBL 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2160
         TabIndex        =   23
         Top             =   720
         Width           =   60
      End
      Begin VB.Label EMPDEPART_LBL 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2160
         TabIndex        =   22
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label6 
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label EMPDESIG_LBL 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2160
         TabIndex        =   20
         Top             =   1440
         Width           =   60
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   5175
      Left            =   7440
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
      Begin VB.TextBox PAYBAL_TXT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2280
         TabIndex        =   37
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox BALANCE_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox OTHER_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox GROSS_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox HOURSWORKED_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox HOURLYRATE_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox BASICSALARY_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox TA_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox MA_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox NET_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox PERSENT_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox ABSENT_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox LEAVE_TXT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label PAYBAL_LBL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "BALANCE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1320
         TabIndex        =   38
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PERSENT DAYS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DUES BALANCE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label Label4 
         Caption         =   "OTHER PAY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label GROSS_LBL 
         Caption         =   "GROSS PAY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hours worked"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hourly Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Basic Salary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Transport Allowance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1920
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Medical allowance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label NETPAY_TXT 
         AutoSize        =   -1  'True
         Caption         =   "NET PAY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "ABSENT DAYS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label Label18 
         Caption         =   "LEAVES DAYS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   1215
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   15
      FillColor       =   &H00FF0000&
      Height          =   615
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "PROCESS STAFF PAYROLL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3720
      TabIndex        =   27
      Top             =   720
      Width           =   3480
   End
End
Attribute VB_Name = "FRMPAYROLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command8_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub BACK_CMD_Click()
Unload Me
Me.Hide
End Sub




Private Sub EMPID_CMB_Click()
'PAYPROCESS_CMD.Enabled = True
BALANCE_TXT.Text = 0
CON
S = "SELECT *FROM EMP_RECORD WHERE EMP_ID='" & EMPID_CMB.Text & "'"
Set R = C.Execute(S)
EMPNAME_LBL.Caption = R.Fields("EMP_NAME")
EMPDESIG_LBL.Caption = R.Fields("DESIGNATION")
EMPDEPART_LBL.Caption = R.Fields("DEPARTMENT")
HOURLYRATE_TXT.Text = R.Fields("SALARY")
S = "SELECT DUES FROM PAYMENT WHERE EMP_ID='" & EMPID_CMB.Text & "' AND STATUS='DUES'"
Set R = C.Execute(S)
Do Until R.EOF = True
BALANCE_TXT.Text = (BALANCE_TXT.Text) + R.Fields("DUES")
R.MoveNext
Loop
S = "SELECT *FROM EMP_ATTENDENCE_DETAIL WHERE EMP_ID='" & EMPID_CMB.Text & "'"
Set R = C.Execute(S)
HOURSWORKED_TXT.Text = R.Fields("HOURS_WORK")
PERSENT_TXT.Text = R.Fields("PERSENTING")
LEAVE_TXT.Text = R.Fields("LEAVES")
ABSENT_TXT.Text = R.Fields("ABSENTLY")
'BALANCE_TXT.Text = R.Fields("TOT_SAL")
days = Val(PERSENT_TXT.Text) + Val(ABSENT_TXT.Text) + Val(LEAVE_TXT.Text)
If days = 0 Then
MsgBox "NOT PERSENTING SINGLE DAY"
Exit Sub
End If
BASIC = Val(HOURSWORKED_TXT.Text) * Val(HOURLYRATE_TXT.Text)
ABSENT = Val(ABSENT_TXT.Text) * (BASIC / days) / 2
LEAVE = Val(LEAVE_TXT.Text) * (BASIC / days) / 2
OTHER_TXT.Text = LEAVE - ABSENT
BASIC = BASIC + OTHER
BASICSALARY_TXT.Text = BASIC
If BASIC < 5000 Then
TA = 0
MA = 0
ElseIf BASIC >= 5000 And BASIC < 15000 Then
TA = BASIC * (5 / 100)
MA = BASIC * (5 / 100)
TA_TXT.Text = TA
MA_TXT.Text = MA
ElseIf BASIC >= 15000 And BASIC < 25000 Then
TA = BASIC * (7 / 100)
MA = BASIC * (7 / 100)
TA_TXT.Text = TA
MA_TXT.Text = MA
ElseIf BASIC >= 25000 And BASIC < 35000 Then
TA = BASIC * (10 / 100)
MA = BASIC * (10 / 100)
TA_TXT.Text = TA
MA_TXT.Text = MA
Else
TA = BASIC * (12 / 100)
MA = BASIC * (12 / 100)
TA_TXT.Text = TA
MA_TXT.Text = MA
End If
GROSS = TA + MA + OTHER
NET = BASIC + GROSS + Val(BALANCE_TXT.Text)
GROSS_TXT.Text = GROSS
NET_TXT.Text = NET
End Sub

Private Sub Form_Load()
CON
S = "SELECT *FROM EMP_RECORD "
Set R = C.Execute(S)
Do Until R.EOF = True
EMPID_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
End Sub

Private Sub PAYBAL_TXT_Change()
If Len(PAYBAL_TXT.Text) > 0 Then
PAID_CMD.Enabled = True
Else
PAID_CMD.Enabled = False
End If
End Sub

Private Sub PAYBAL_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
PAID_CMD.SetFocus
End If
End Sub

Private Sub PAYPROCESS_CMD_Click()

End Sub

Private Sub PAID_CMD_Click()
CON
If MsgBox("ARE YOU SURE TO PAY AMOUNT ?", vbYesNo) = vbYes Then
If Val(PAYBAL_TXT.Text) > Val(NET_TXT.Text) Then
MsgBox "THE PAYMENT IS MORE THAN NET PAY"
PAYBAL_TXT.Text = ""
PAYBAL_TXT.SetFocus
Exit Sub
ElseIf Val(PAYBAL_TXT.Text) < Val(NET_TXT.Text) Then
S = "UPDATE PAYMENT SET STATUS='PAID',DUES=0 WHERE EMP_ID='" & EMPID_CMB.Text & "'"
Set R = C.Execute(S)
S = "INSERT INTO PAYMENT VALUES('" & EMPID_CMB.Text & "','" & Month(Format(Date, "DD-MMM-YYYY")) & "','DUES'," & Val(NET_TXT.Text) & ",'" & Format(Date, "DD-MMM-YYYY") & "'," & PAYBAL_TXT.Text & "," & Val(NET_TXT.Text) - Val(PAYBAL_TXT.Text) & ")"
MsgBox S
Set R = C.Execute(S)
Else
S = "INSERT INTO PAYMENT VALUES('" & EMPID_CMB.Text & "','" & Month(Format(Date, "DD-MMM-YYYY")) & "','PAID'," & Val(NET_TXT.Text) & ",'" & Format(Date, "DD-MMM-YYYY") & "'," & PAYBAL_TXT.Text & "," & Val(NET_TXT.Text) - Val(PAYBAL_TXT.Text) & ")"
MsgBox S
Set R = C.Execute(S)
End If
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "PAYMENT DONE"
PAYBAL_TXT.Text = ""
'S = "SELECT COUNT(EMP_ID) FROM PAYMENT"
'Set R = C.Execute(S)
'J = R.Fields(0)
'J = J - 1
'Do Until J > 0
'S = "UPDATE PAYMENT SET STATUS='PAID' , DUES=0 WHERE EMP_ID"
S = "UPDATE EMP_ATTENDENCE_DETAIL SET HOURS_WORK=0 WHERE EMP_ID='" & EMPID_CMB.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
'PAYBAL_LBL.Visible = False
'PAYBAL_TXT.Visible = False
EMPID_CMB.SetFocus
End If
End Sub
