VERSION 5.00
Begin VB.Form FRMGUESTPAYMENT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GUEST ACCOMMODATION"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton GUESTACOMREPORT_CMD 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton CANCELFROMCHECKOUT_CMD 
      Caption         =   "CANCEL"
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton PROCESSFROMCHEKOUT_CMD 
      Caption         =   "PROCESS"
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox GUESTIDACOM_CMB 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox GUESTNAMEACOM_TXT 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox GUESTADVANCEACOM_TXT 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox GUESTPAYMENTACOM_TXT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox GUESTTOTCHARGEACOM_TXT 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox GUESTBALANCESTATUSACOM_TXT 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton GUESTACOMSAVE_CMD 
      Caption         =   "Save"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton GUESTACOMCANCLE_CMD 
      Caption         =   "Cancel"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox GUESTCHECKINDATEACOM_TXT 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label9 
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GUEST ACCOMMODATION PAYMENT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   360
      Width           =   4470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Guest ID"
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
      Left            =   1560
      TabIndex        =   15
      Top             =   960
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name"
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
      Left            =   1560
      TabIndex        =   14
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Advance Payment"
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
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Payment"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   3120
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Bill"
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
      Left            =   1560
      TabIndex        =   11
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Dues Payment"
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
      Left            =   1560
      TabIndex        =   10
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Date of Checkin"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   1485
   End
End
Attribute VB_Name = "FRMGUESTPAYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private T As Integer
Private T2 As Integer

Private Sub Command1_Click()
CON
S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & Label9.Caption & "'"
Set R = C.Execute(S)

End Sub

Private Sub CANCELFROMCHECKOUT_CMD_Click()
Unload Me
Me.Hide
Unload FRMCHECKOUT
FRMCHECKOUT.Hide
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub Form_Load()
CON
S = " SELECT CLIENT_ID FROM CLIENT_MASTER WHERE DUES_BALANCE>0 "
Set R = C.Execute(S)
Do Until R.EOF = True
    GUESTIDACOM_CMB.AddItem R.Fields("CLIENT_ID")
    R.MoveNext
Loop
End Sub

Private Sub GUESTACOMCANCLE_CMD_Click()
Unload Me
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub GUESTACOMREPORT_CMD_Click()
If GUESTIDACOM_CMB.Text = "" Then
    MsgBox "PLEASE SELECT ID FIRST"
    GUESTIDACOM_CMB.SetFocus
Else
    If DataEnvironment1.rsCommand6.State = 1 Then DataEnvironment1.rsCommand6.Close
    DataEnvironment1.Command6 GUESTIDACOM_CMB.Text
    ACOMODATION_BILL.Show
End If
End Sub

Private Sub GUESTACOMSAVE_CMD_Click()
If GUESTNAMEACOM_TXT.Text = "" Then
MsgBox "PLEASE ENTER GUEST ID"
GUESTIDACOM_CMB.SetFocus
ElseIf GUESTPAYMENTACOM_TXT.Text = "" Then
MsgBox "PLEASE INPUT AMOUNT"
GUESTPAYMENTACOM_TXT.SetFocus
Else
If MsgBox("ARE YOU SURE FOR PAYMENT", vbYesNo) = vbYes Then
CON
'S = "UPDATE CLIENT_MASTER SET BILL_STATUS=BILL_STATUS+" & Val(GUESTPAYMENTACOM_TXT.Text) & ",DUES_BALANCE=DUES_BALANCE-" & Val(GUESTPAYMENTACOM_TXT.Text) & " WHERE CLIENT_ID='" & GUESTIDACOM_CMB.Text & "' "
S = " UPDATE CLIENT_MASTER SET BILL_STATUS=BILL_STATUS+" & Val(GUESTPAYMENTACOM_TXT.Text) & ",DUES_BALANCE=DUES_BALANCE- " & Val(GUESTPAYMENTACOM_TXT.Text) & " WHERE CLIENT_ID='" & GUESTIDACOM_CMB.Text & "'"
MsgBox S
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "YOUR PAYMENT IS DONE"
GUESTNAMEACOM_TXT.Text = ""
GUESTCHECKINDATEACOM_TXT.Text = ""
GUESTTOTCHARGEACOM_TXT.Text = ""
GUESTADVANCEACOM_TXT.Text = ""
GUESTBALANCESTATUSACOM_TXT.Text = ""
GUESTPAYMENTACOM_TXT.Text = ""
Else
MsgBox "YOUR PAYMENT IS NOT COMPLETE"
GUESTNAMEACOM_TXT.Text = ""
GUESTCHECKINDATEACOM_TXT.Text = ""
GUESTTOTCHARGEACOM_TXT.Text = ""
GUESTADVANCEACOM_TXT.Text = ""
GUESTBALANCESTATUSACOM_TXT.Text = ""
GUESTPAYMENTACOM_TXT.Text = ""
End If
End If
End Sub



Private Sub GUESTIDACOM_CMB_Click()
CON
S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & GUESTIDACOM_CMB.Text & "'"
Set R = C.Execute(S)
DIN1 = R.Fields("CHECK_IN_DATE")
DIN = DateDiff("D", DIN1, Now)
TOTALTIME = R.Fields("TOT_TIME")
If DIN > TOTALTIME Then
    ROOMCHARGE = R.Fields("ROOM_CHARGE")
    DIN = DIN - TOTALTIME
    T2 = DIN + TOTALTIME
    T = (DIN * 500)
    CON
    S = "UPDATE CLIENT_MASTER SET TOT_CHARGE=TOT_CHARGE + " & T & ", TOT_TIME=" & T2 & ",DUES_BALANCE=DUES_BALANCE+" & T & " WHERE CLIENT_ID='" & GUESTIDACOM_CMB.Text & "'"
    MsgBox S
    Set R = C.Execute(S)
    S = "COMMIT"
    Set R = C.Execute(S)
    CON
    S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & GUESTIDACOM_CMB.Text & "'"
    Set R = C.Execute(S)
    GUESTNAMEACOM_TXT.Text = R.Fields("NAME")
    GUESTCHECKINDATEACOM_TXT.Text = R.Fields("CHECK_IN_DATE")
    GUESTTOTCHARGEACOM_TXT.Text = R.Fields("TOT_CHARGE")
    GUESTADVANCEACOM_TXT.Text = R.Fields("BILL_STATUS")
    GUESTBALANCESTATUSACOM_TXT.Text = R.Fields("DUES_BALANCE")
Else
    CON
    S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & GUESTIDACOM_CMB.Text & "'"
    Set R = C.Execute(S)
    GUESTNAMEACOM_TXT.Text = R.Fields("NAME")
    GUESTCHECKINDATEACOM_TXT.Text = R.Fields("CHECK_IN_DATE")
    GUESTTOTCHARGEACOM_TXT.Text = R.Fields("TOT_CHARGE")
    GUESTADVANCEACOM_TXT.Text = R.Fields("BILL_STATUS")
    GUESTBALANCESTATUSACOM_TXT.Text = R.Fields("DUES_BALANCE")
End If
End Sub

Private Sub GUESTPAYMENTACOM_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
GUESTACOMSAVE_CMD.SetFocus
End If
End Sub

Private Sub PROCESSFROMCHEKOUT_CMD_Click()
If GUESTPAYMENTACOM_TXT.Text = "" Then
    MsgBox "PLEASE INPUT AMOUNT"
    GUESTPAYMENTACOM_TXT.SetFocus
ElseIf Val(GUESTPAYMENTACOM_TXT.Text) < Val(GUESTBALANCESTATUSACOM_TXT.Text) Then
    MsgBox "PLESE CHECK THE DUES BALANCE PAID LESS THAN DUES BALANCE WHICH IS NOT VALLID."
    If MsgBox("ARE YOU WANT TO MAKE PAYMENT EQUAL TO DUES BALANCE THEN CLICK 'YES' OTHERWISE 'NO' .", vbYesNo) = vbYes Then
        GUESTPAYMENTACOM_TXT.SetFocus
    Else
        MsgBox "YOUR PAYMENT IS NOT COMPLETE"
        Unload Me
        Me.Hide
    End If
ElseIf Val(GUESTPAYMENTACOM_TXT.Text) > Val(GUESTBALANCESTATUSACOM_TXT.Text) Then
    MsgBox "PLESE CHECK THE DUES BALANCE PAID GREATER THAN DUES BALANCE WHICH IS NOT VALLID."
    If MsgBox("ARE YOU WANT TO MAKE PAYMENT EQUAL TO DUES BALANCE THEN CLICK 'YES' OTHERWISE 'NO' .", vbYesNo) = vbYes Then
        GUESTPAYMENTACOM_TXT.SetFocus
    Else
        MsgBox "YOUR PAYMENT IS NOT COMPLETE"
        Unload Me
        Me.Hide
    End If
Else
If MsgBox("ARE YOU SURE FOR PAYMENT", vbYesNo) = vbYes Then
CON
'S = "UPDATE CLIENT_MASTER SET BILL_STATUS=BILL_STATUS+" & Val(GUESTPAYMENTACOM_TXT.Text) & ",DUES_BALANCE=DUES_BALANCE-" & Val(GUESTPAYMENTACOM_TXT.Text) & " WHERE CLIENT_ID='" & GUESTIDACOM_CMB.Text & "' "
S = " UPDATE CLIENT_MASTER SET BILL_STATUS=BILL_STATUS+" & Val(GUESTPAYMENTACOM_TXT.Text) & ",DUES_BALANCE=DUES_BALANCE- " & Val(GUESTPAYMENTACOM_TXT.Text) & " WHERE CLIENT_ID='" & Label9.Caption & "'"
MsgBox S
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "YOUR PAYMENT IS DONE"
FRMCHECKOUT.DUESAMOUNTCHECKOUT_TXT.Text = 0
Unload Me
Me.Hide
Else
MsgBox "YOUR PAYMENT IS NOT COMPLETE"
Unload Me
Me.Hide
End If
End If
End Sub
