VERSION 5.00
Begin VB.Form FRMCHECKOUT 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHECK OUT DATE"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox GUESTID_LIST 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5100
      ItemData        =   "FRMCHECKOUT.frx":0000
      Left            =   8160
      List            =   "FRMCHECKOUT.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox CHECKOUTDATE_TXT 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8640
      Top             =   840
   End
   Begin VB.CommandButton CHECKOUT_CMD 
      BackColor       =   &H8000000A&
      Caption         =   "CHECK OUT"
      Enabled         =   0   'False
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
      TabIndex        =   10
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Frame CHECKOUT_FRAME 
      BackColor       =   &H8000000A&
      Caption         =   "CHECKOUT"
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
      Height          =   3495
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   5415
      Begin VB.TextBox DUESAMOUNTCHECKOUT_TXT 
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
         Left            =   2400
         TabIndex        =   21
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox TOTCHARGECHECKOUT_TXT 
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
         Left            =   2400
         TabIndex        =   19
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox GNAMECHECKOUT_TXT 
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
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   2895
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
         Left            =   2400
         TabIndex        =   15
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox CHECKINDATE_TXT 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox ROOMTYPE_TXT 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox ROOMNO_TXT 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label DUESAMOUNTCHECKOUT_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DUES AMOUNT"
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
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label TOTCHARGECHECKOUT_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL CHARGE"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   2055
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
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label CHEKINTIME_LBL 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "CHECKIN TIME"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label CHECKINDTAE_LBL 
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
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label ROOMNOCHECKOUT_LBL 
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
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label ROOMTYPECHECKOUT_LBL 
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1530
      End
   End
   Begin VB.CommandButton REPORTCHECKIN_CMD 
      BackColor       =   &H8000000A&
      Caption         =   "&REPORT"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton BACKTOMDIFROMCHECKIN_CMD 
      BackColor       =   &H8000000A&
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox GIDCHECKOUT_TXT 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
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
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton GOCHECKOUT_CMD 
      BackColor       =   &H8000000A&
      Caption         =   "&GO"
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
      Height          =   435
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label CHECKOUTDATE_LBL 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
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
      Height          =   285
      Left            =   2400
      TabIndex        =   24
      Top             =   1440
      Width           =   2220
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
      Left            =   9600
      TabIndex        =   13
      Top             =   0
      Width           =   975
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
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.Label GIDCHECKOUT_LBL 
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
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "FRMCHECKOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ADDNEWCHECKIN_CMD_Click()

End Sub

Private Sub BACKTOMDIFROMCHECKIN_CMD_Click()
Unload Me
Me.Hide
End Sub

Private Sub CHECKOUT_CMD_Click()
If GIDCHECKOUT_TXT.Text = "" Then
MsgBox "PLESE INPUT GUEST INFORMATION FIRST"
ElseIf Val(DUESAMOUNTCHECKOUT_TXT.Text) > 0 Then
MsgBox "PLESE PAY DUES FIRST"
Load FRMGUESTPAYMENT
FRMGUESTPAYMENT.Show
FRMGUESTPAYMENT.GUESTIDACOM_CMB.Visible = False
FRMGUESTPAYMENT.GUESTACOMSAVE_CMD.Visible = False
FRMGUESTPAYMENT.GUESTACOMCANCLE_CMD.Visible = False
FRMGUESTPAYMENT.PROCESSFROMCHEKOUT_CMD.Visible = True
FRMGUESTPAYMENT.Label9.Visible = True
FRMGUESTPAYMENT.CANCELFROMCHECKOUT_CMD.Visible = True
FRMGUESTPAYMENT.Label9.Caption = GIDCHECKOUT_TXT.Text
FRMGUESTPAYMENT.GUESTNAMEACOM_TXT.Text = GNAMECHECKOUT_TXT.Text
FRMGUESTPAYMENT.GUESTCHECKINDATEACOM_TXT.Text = CHECKINDATE_TXT.Text
FRMGUESTPAYMENT.GUESTTOTCHARGEACOM_TXT.Text = TOTCHARGECHECKOUT_TXT.Text
FRMGUESTPAYMENT.GUESTADVANCEACOM_TXT.Text = D
FRMGUESTPAYMENT.GUESTBALANCESTATUSACOM_TXT.Text = DUESAMOUNTCHECKOUT_TXT.Text
Else
CON
S = "UPDATE CLIENT_MASTER SET CHECK_OUT_DATE='" & Format(Date, "DD-MMM-YYYY") & "',OUT='OUT' WHERE CLIENT_ID='" & GIDCHECKOUT_TXT.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox " CHECKOUT COMPLITE"
REPORTCHECKIN_CMD.Enabled = True
End If
End Sub

Private Sub DUESAMOUNTCHECKOUT_TXT_Change()
If Val(DUESAMOUNTCHECKOUT_TXT.Text) = 0 Then
CHECKOUT_CMD.Enabled = True
Else
CHECKOUT_CMD.Enabled = False
End If
End Sub

Private Sub Form_Load()
CON
S = "SELECT CLIENT_ID FROM CLIENT_MASTER WHERE DUES_BALANCE>0"
Set R = C.Execute(S)
Do Until R.EOF = True
    GUESTID_LIST.AddItem R.Fields("CLIENT_ID")
    R.MoveNext
Loop
End Sub

Private Sub GIDCHECKOUT_TXT_Change()
If Len(GIDCHECKOUT_TXT.Text) > 0 Then
GOCHECKOUT_CMD.Enabled = True
Else
GOCHECKOUT_CMD.Enabled = False
End If
End Sub

Private Sub GIDCHECKOUT_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
    GIDCHECKOUT_TXT.Text = UCase(GIDCHECKOUT_TXT.Text)
    GOCHECKOUT_CMD.SetFocus
End If
End Sub

Private Sub GNAMECHECKOUT_TXT_Change()
If Len(GNAMECHECKOUT_TXT.Text) > 0 Then
REPORTCHECKIN_CMD.Enabled = True
Else
REPORTCHECKIN_CMD.Enabled = False
End If
End Sub

Private Sub GOCHECKOUT_CMD_Click()
CON
S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & GIDCHECKOUT_TXT.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
    MsgBox "WRONG GUEST ID", vbCritical
    GIDCHECKOUT_TXT.Text = ""
    GIDCHECKOUT_TXT.SetFocus
    GNAMECHECKOUT_TXT.Text = ""
    CHECKINDATE_TXT.Text = ""
    ROOMTYPE_TXT.Text = ""
    ROOMNO_TXT.Text = ""
    CHECKINTIME_TXT.Text = ""
    TOTCHARGECHECKOUT_TXT.Text = ""
    DUESAMOUNTCHECKOUT_TXT.Text = ""
Else
    GIDCHECKOUT_TXT.Text = R.Fields("CLIENT_ID")
    GNAMECHECKOUT_TXT.Text = R.Fields("NAME")
    CHECKINDATE_TXT.Text = R.Fields("CHECK_IN_DATE")
    ROOMTYPE_TXT.Text = R.Fields("SUIT_PROFILE")
    ROOMNO_TXT.Text = R.Fields("SUIT_NO")
    CHECKINTIME_TXT.Text = R.Fields("CHECK_IN_TIME")
    TOTCHARGECHECKOUT_TXT.Text = R.Fields("TOT_CHARGE")
    DUESAMOUNTCHECKOUT_TXT.Text = R.Fields("DUES_BALANCE")
    D = R.Fields("BILL_STATUS")
    If DUESAMOUNTCHECKOUT_TXT.Text > 0 Then
        If MsgBox("YOU HAVE DUES . ARE YOU WANT TO PAID YOUR DUES? ,IF YES THEN CLICK ON 'YES' OTHERWISE CLICK ON 'NO' ", vbYesNo) = vbYes Then
            Load FRMGUESTPAYMENT
            FRMGUESTPAYMENT.Show
            FRMGUESTPAYMENT.GUESTIDACOM_CMB.Visible = False
            FRMGUESTPAYMENT.GUESTACOMSAVE_CMD.Visible = False
            FRMGUESTPAYMENT.GUESTACOMCANCLE_CMD.Visible = False
            FRMGUESTPAYMENT.PROCESSFROMCHEKOUT_CMD.Visible = True
            FRMGUESTPAYMENT.Label9.Visible = True
            FRMGUESTPAYMENT.CANCELFROMCHECKOUT_CMD.Visible = True
            FRMGUESTPAYMENT.Label9.Caption = GIDCHECKOUT_TXT.Text
            FRMGUESTPAYMENT.GUESTNAMEACOM_TXT.Text = GNAMECHECKOUT_TXT.Text
            FRMGUESTPAYMENT.GUESTCHECKINDATEACOM_TXT.Text = CHECKINDATE_TXT.Text
            FRMGUESTPAYMENT.GUESTTOTCHARGEACOM_TXT.Text = TOTCHARGECHECKOUT_TXT.Text
            FRMGUESTPAYMENT.GUESTADVANCEACOM_TXT.Text = D
            FRMGUESTPAYMENT.GUESTBALANCESTATUSACOM_TXT.Text = DUESAMOUNTCHECKOUT_TXT.Text
        Else
            CHECKOUT_CMD.Enabled = False
            REPORTCHECKIN_CMD.Enabled = True
        End If
    Else
        CHECKOUT_CMD.SetFocus
    End If
End If
End Sub

Private Sub GUESTID_LIST_Click()
CON
S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & GUESTID_LIST.Text & "'"
Set R = C.Execute(S)
GIDCHECKOUT_TXT.Text = R.Fields("CLIENT_ID")
GNAMECHECKOUT_TXT.Text = R.Fields("NAME")
CHECKINDATE_TXT.Text = R.Fields("CHECK_IN_DATE")
ROOMTYPE_TXT.Text = R.Fields("SUIT_PROFILE")
ROOMNO_TXT.Text = R.Fields("SUIT_NO")
CHECKINTIME_TXT.Text = R.Fields("CHECK_IN_TIME")
TOTCHARGECHECKOUT_TXT.Text = R.Fields("TOT_CHARGE")
DUESAMOUNTCHECKOUT_TXT.Text = R.Fields("DUES_BALANCE")
D = R.Fields("BILL_STATUS")
If DUESAMOUNTCHECKOUT_TXT.Text > 0 Then
If MsgBox("YOU HAVE DUES . ARE YOU WANT TO PAID YOUR DUES? ,IF YES THEN CLICK ON 'YES' OTHERWISE CLICK ON 'NO' ", vbYesNo) = vbYes Then
Load FRMGUESTPAYMENT
FRMGUESTPAYMENT.Show
FRMGUESTPAYMENT.GUESTIDACOM_CMB.Visible = False
FRMGUESTPAYMENT.GUESTACOMSAVE_CMD.Visible = False
FRMGUESTPAYMENT.GUESTACOMCANCLE_CMD.Visible = False
FRMGUESTPAYMENT.PROCESSFROMCHEKOUT_CMD.Visible = True
FRMGUESTPAYMENT.Label9.Visible = True
FRMGUESTPAYMENT.CANCELFROMCHECKOUT_CMD.Visible = True
FRMGUESTPAYMENT.Label9.Caption = GIDCHECKOUT_TXT.Text
FRMGUESTPAYMENT.GUESTNAMEACOM_TXT.Text = GNAMECHECKOUT_TXT.Text
FRMGUESTPAYMENT.GUESTCHECKINDATEACOM_TXT.Text = CHECKINDATE_TXT.Text
FRMGUESTPAYMENT.GUESTTOTCHARGEACOM_TXT.Text = TOTCHARGECHECKOUT_TXT.Text
FRMGUESTPAYMENT.GUESTADVANCEACOM_TXT.Text = D
FRMGUESTPAYMENT.GUESTBALANCESTATUSACOM_TXT.Text = DUESAMOUNTCHECKOUT_TXT.Text
Else
GIDCHECKOUT_TXT.Text = ""
GNAMECHECKOUT_TXT.Text = ""
CHECKINDATE_TXT.Text = ""
ROOMTYPE_TXT.Text = ""
ROOMNO_TXT.Text = ""
CHECKINTIME_TXT.Text = ""
TOTCHARGECHECKOUT_TXT.Text = ""
DUESAMOUNTCHECKOUT_TXT.Text = ""
End If
Else
CHECKOUT_CMD.SetFocus
End If
End Sub

Private Sub REPORTCHECKIN_CMD_Click()
If GIDCHECKOUT_TXT.Text = "" Then
MsgBox "PLESE INPUT GUEST ID FIRST"
GIDCHECKOUT_TXT.SetFocus
REPORTCHECKIN_CMD.Enabled = False
Else
DataEnvironment1.Command1 GIDCHECKOUT_TXT
DataReport2.Show
End If
End Sub

Private Sub Timer1_Timer()
CHECKINTIME_LBL.Caption = Time
RESERVDATE_LBL.Caption = Date
End Sub
