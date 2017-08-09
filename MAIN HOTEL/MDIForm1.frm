VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "HOTEL MANAGEMENT SYSTEM"
   ClientHeight    =   9675
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   20250
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4800
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   20190
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      Begin VB.PictureBox Picture3 
         Height          =   1575
         Left            =   18480
         Picture         =   "MDIForm1.frx":34B79
         ScaleHeight     =   1515
         ScaleWidth      =   1995
         TabIndex        =   3
         Top             =   0
         Width           =   2055
      End
      Begin VB.PictureBox Picture2 
         Height          =   1575
         Left            =   0
         Picture         =   "MDIForm1.frx":36F8E
         ScaleHeight     =   1515
         ScaleWidth      =   1995
         TabIndex        =   2
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "MAHARAJA INN HOTEL"
         BeginProperty Font 
            Name            =   "Myriad Pro Light"
            Size            =   48
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   1575
         Left            =   4200
         TabIndex        =   1
         Top             =   0
         Width           =   9015
      End
   End
   Begin VB.Menu MNUGUSEST 
      Caption         =   "&GUEST"
      Begin VB.Menu MNURESEV 
         Caption         =   "RESERVATION"
         Shortcut        =   ^R
      End
      Begin VB.Menu MNUCHECKIN 
         Caption         =   "CHECK IN"
         Shortcut        =   ^N
      End
      Begin VB.Menu MNUPAYMENT 
         Caption         =   "MAKE PAYMENT"
         Shortcut        =   ^M
      End
      Begin VB.Menu MNUCHECKOUT 
         Caption         =   "CHECK OUT"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu MNUEMPDETAIL 
      Caption         =   "&EMPLOYEE "
      Begin VB.Menu MNUNEWEMP 
         Caption         =   "ADD NEW EMPLOYEE"
      End
      Begin VB.Menu MNUMODIFYEMPDETAIL 
         Caption         =   "MODIFY EMPLOYEE DETAIL"
      End
      Begin VB.Menu EMP_ATTEN 
         Caption         =   "EMPLOYEE ATTENDENCE"
      End
      Begin VB.Menu MNUPAYROLLEMP 
         Caption         =   "EMPLOYEE PAYROLL"
      End
      Begin VB.Menu MNUDELEMP 
         Caption         =   "DELETE EMPLOYEE"
      End
   End
   Begin VB.Menu MNUREPORT 
      Caption         =   "&REPORT"
      Begin VB.Menu MNUGUIEST 
         Caption         =   "GUEST"
      End
      Begin VB.Menu MNUEMP 
         Caption         =   "EMPLOYEE"
      End
   End
   Begin VB.Menu MNUVIEW 
      Caption         =   "&VIEW"
      Begin VB.Menu MNUHOTELSTATUS 
         Caption         =   "STATUS OF HOTEL"
      End
   End
   Begin VB.Menu MNUADMIN 
      Caption         =   "&ADMINSTRATOR"
      Begin VB.Menu FRMLGDETAIL 
         Caption         =   "USER DETAIL"
      End
      Begin VB.Menu MNUCHANGECHARGE 
         Caption         =   "CHANGE CHARGES"
         Begin VB.Menu MNURESTO 
            Caption         =   "RESTORENT"
            Begin VB.Menu MNURESTONEW 
               Caption         =   "NEW ENTRY"
            End
            Begin VB.Menu MNURESTOUPDATE 
               Caption         =   "UPDATE DATA"
            End
         End
         Begin VB.Menu MNUROOM 
            Caption         =   "ROOM"
         End
      End
   End
   Begin VB.Menu MNURESTORENT 
      Caption         =   "RESTAURENT"
      Begin VB.Menu MNUGUESTRESTO 
         Caption         =   "GUEST RESTAURENT"
      End
      Begin VB.Menu MNUCUSTRESTO 
         Caption         =   "CUSTOMER RESTAURENT"
      End
   End
   Begin VB.Menu LOG 
      Caption         =   "&SETING"
      Begin VB.Menu CHANGEPASS 
         Caption         =   "CHANGE PASSWORD"
      End
      Begin VB.Menu FRMLGOUT 
         Caption         =   "LOG OUT"
      End
      Begin VB.Menu OUT 
         Caption         =   "EXIT"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CHANGEPASS_Click()
Load FRMCHANGEPASSWORD
FRMCHANGEPASSWORD.Show
End Sub

Private Sub EMP_ATTEN_Click()
Load FRMATTENDENCE
FRMATTENDENCE.Show
End Sub

Private Sub FRMLGDETAIL_Click()
Load FRMLGDETAL
FRMLGDETAL.Show
End Sub
Private Sub FRMLGOUT_Click()
Unload Me
Load FRMLOG
FRMLOG.Show
End Sub

Private Sub MDIForm_Load()
CON
S = "SELECT *FROM DAILY WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
If R.EOF = True Then
S = "INSERT INTO DAILY VALUES(" & 0 & "," & 0 & "," & 0 & "," & 0 & ",'" & Format(Date, "DD-MMM-YYYY") & "')"
MsgBox S
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "YOU ARE WELCOME"
Else
MsgBox "YOU ARE WELCOME AGAIN"
End If
End Sub



Private Sub MNUCHECKIN_Click()
Load FRMCHECKIN
FRMCHECKIN.Show
End Sub

Private Sub MNUCHECKOUT_Click()
Load FRMCHECKOUT
FRMCHECKOUT.Show
End Sub

Private Sub MNUCOMPNYINF_Click()

End Sub

Private Sub MNUDAILYEVOLUTION_Click()

End Sub

Private Sub MNUCUSTRESTO_Click()
Load FRMRESTORENT
FRMRESTORENT.Show
End Sub

Private Sub MNUDELEMP_Click()
Unload FRMEMPENTRY
FRMEMPENTRY.Refresh
Load FRMEMPENTRY
FRMEMPENTRY.Show
FRMEMPENTRY.EMPID_TXT.Visible = False
FRMEMPENTRY.ADDNEWEMP_CMD.Visible = False
FRMEMPENTRY.SAVEEMP_CMD.Visible = False
FRMEMPENTRY.BACKTOMDIFROMEMPENTRY_CMD.Visible = False
FRMEMPENTRY.Shape1.Visible = False
FRMEMPENTRY.Shape2.Visible = False
FRMEMPENTRY.UPDATEEMP_CMD.Visible = False
FRMEMPENTRY.EMPID_CMB.Clear
S = "SELECT  EMP_ID FROM EMP_RECORD"
Set R = C.Execute(S)
Do Until R.EOF = True
FRMEMPENTRY.EMPID_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
End Sub

Private Sub MNUEMP_Click()
Unload FRMSEARCHGUESTOREMP
Load FRMSEARCHGUESTOREMP
FRMSEARCHGUESTOREMP.Show
FRMSEARCHGUESTOREMP.GID_CMB.Visible = False
'FRMSEARCHGUESTOREMP.EMPID_TXT.Visible = True
FRMSEARCHGUESTOREMP.GUESTDETAIL_FRAME.Visible = False
FRMSEARCHGUESTOREMP.EMPDETAIL_FRAME.Visible = True
FRMSEARCHGUESTOREMP.EMPID_CMB.Clear
FRMSEARCHGUESTOREMP.SEARCHGUESTREPORT_CMD.Visible = False
CON
S = "SELECT EMP_ID FROM EMP_RECORD"
Set R = C.Execute(S)
Do Until R.EOF = True
FRMSEARCHGUESTOREMP.EMPID_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
S = "SELECT EMP_ID FROM DEL_EMP_RECORD"
Set R = C.Execute(S)
Do Until R.EOF = True
FRMSEARCHGUESTOREMP.EMPID_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
'FRMSEARCHGUESTOREMP.SEARCHEMP_CMD.Visible = True
'FRMSEARCHGUESTOREMP.SEARCHGUEST_CMD.Visible = False
End Sub

Private Sub MNUGUESTRESTO_Click()
Load FRMRESTORENTGUEST
FRMRESTORENTGUEST.Show
End Sub

Private Sub MNUGUIEST_Click()
Unload FRMSEARCHGUESTOREMP
Load FRMSEARCHGUESTOREMP
FRMSEARCHGUESTOREMP.Show
FRMSEARCHGUESTOREMP.GUESTDETAIL_FRAME.Visible = True
FRMSEARCHGUESTOREMP.EMPDETAIL_FRAME.Visible = False
FRMSEARCHGUESTOREMP.SEARCHGUESTREPORT_CMD.Visible = True
FRMSEARCHGUESTOREMP.GID_CMB.Visible = True
FRMSEARCHGUESTOREMP.EMPID_CMB.Visible = False
FRMSEARCHGUESTOREMP.GID_CMB.Clear
CON
S = "SELECT CLIENT_ID FROM CLIENT_MASTER"
Set R = C.Execute(S)
Do Until R.EOF = True
FRMSEARCHGUESTOREMP.GID_CMB.AddItem R.Fields("CLIENT_ID")
R.MoveNext
Loop
S = "SELECT CLIENT_ID FROM DEL_CLIENT_MASTER"
Set R = C.Execute(S)
Do Until R.EOF = True
FRMSEARCHGUESTOREMP.GID_CMB.AddItem R.Fields("CLIENT_ID")
R.MoveNext
Loop
'FRMSEARCHGUESTOREMP.SEARCHGUEST_CMD.Visible = True
End Sub

Private Sub MNUHOTELSTATUS_Click()
Load FRMHOTELINFO
FRMHOTELINFO.Show
End Sub

Private Sub MNUMODIFYEMPDETAIL_Click()
Unload FRMEMPENTRY
FRMEMPENTRY.Refresh
Load FRMEMPENTRY
FRMEMPENTRY.Show
FRMEMPENTRY.EMPID_TXT.Visible = False
FRMEMPENTRY.ADDNEWEMP_CMD.Visible = False
FRMEMPENTRY.SAVEEMP_CMD.Visible = False
FRMEMPENTRY.BACKTOMDIFROMEMPENTRY_CMD.Visible = False
FRMEMPENTRY.Shape1.Visible = False
FRMEMPENTRY.DELETEEMP_CMD.Visible = False
FRMEMPENTRY.EMPID_CMB.Clear
S = "SELECT  EMP_ID FROM EMP_RECORD"
Set R = C.Execute(S)
Do Until R.EOF = True
FRMEMPENTRY.EMPID_CMB.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
End Sub

Private Sub MNUNEWEMP_Click()
Unload FRMEMPENTRY
FRMEMPENTRY.Refresh
Load FRMEMPENTRY
FRMEMPENTRY.Show
'FRMEMPENTRY.SEARCHOK_CMD.Visible = False
FRMEMPENTRY.DELETEEMP_CMD.Visible = False
FRMEMPENTRY.UPDATEEMP_CMD.Visible = False
FRMEMPENTRY.BACKEMPTOMDIFORM_CMD.Visible = False
'FRMEMPENTRY.MODIFYEMP_CMD.Visible = False
FRMEMPENTRY.Shape2.Visible = False
FRMEMPENTRY.EMPID_CMB.Visible = False
End Sub

Private Sub MNUPAYMENT_Click()
Load FRMGUESTPAYMENT
FRMGUESTPAYMENT.Show
End Sub

Private Sub MNUPAYROLLEMP_Click()
Load FRMPAYROLL
FRMPAYROLL.Show
End Sub

Private Sub MNURESEV_Click()
Load FRMRESERVATION
FRMRESERVATION.Show
End Sub




Private Sub MNURESTONEW_Click()
Unload FRMDISHANDDRINKENTRY
Load FRMDISHANDDRINKENTRY
FRMDISHANDDRINKENTRY.Height = 5490
FRMDISHANDDRINKENTRY.Width = 8460
FRMDISHANDDRINKENTRY.Show
FRMDISHANDDRINKENTRY.ENTRY_FRAME.Visible = True
FRMDISHANDDRINKENTRY.UPDATE_FRAME.Visible = False
End Sub

Private Sub MNURESTOUPDATE_Click()
Unload FRMDISHANDDRINKENTRY
Load FRMDISHANDDRINKENTRY
FRMDISHANDDRINKENTRY.Show
FRMDISHANDDRINKENTRY.ENTRY_FRAME.Visible = False
FRMDISHANDDRINKENTRY.UPDATE_FRAME.Visible = True
End Sub

Private Sub MNUROOM_Click()
Load FRMROOMDETAIL
FRMROOMDETAIL.Show
End Sub

Private Sub OUT_Click()
If MsgBox("ARE YOU SURE TO EXIT", vbYesNo) = vbYes Then
    End
End If
End Sub

Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 100
If Label1.Left + Label1.Width <= 0 Then
Label1.Left = Picture1.Width
End If
End Sub
