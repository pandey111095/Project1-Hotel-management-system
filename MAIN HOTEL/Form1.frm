VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRMLOG 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN FORM"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C000C0&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   8805
      TabIndex        =   11
      Top             =   0
      Width           =   8865
      Begin VB.Image Image3 
         Height          =   1260
         Left            =   7560
         Picture         =   "Form1.frx":240042
         Top             =   0
         Width           =   1260
      End
      Begin VB.Image Image2 
         Height          =   1260
         Left            =   11280
         Picture         =   "Form1.frx":240947
         Top             =   0
         Width           =   1260
      End
      Begin VB.Image Image1 
         Height          =   1260
         Left            =   0
         Picture         =   "Form1.frx":24124C
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "MAHRAJA INN "
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
         Height          =   1335
         Left            =   1440
         TabIndex        =   12
         Top             =   0
         Width           =   6015
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   720
      Top             =   1680
   End
   Begin TabDlg.SSTab LGSSTab1 
      Height          =   1935
      Left            =   960
      TabIndex        =   10
      Top             =   2040
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777088
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ADMINISTRATOR"
      TabPicture(0)   =   "Form1.frx":241B51
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LBLADMINPASS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LBLADMINID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CHEKADMINPASS"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CMDEXITADMINLOGIN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CMDADMINLOGIN"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TXTADMINPASS"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXTADMINID"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "USER"
      TabPicture(1)   =   "Form1.frx":241B6D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TXTUSERPASS"
      Tab(1).Control(1)=   "CMDUSERLOGIN"
      Tab(1).Control(2)=   "CMDEXITUSERLOGIN"
      Tab(1).Control(3)=   "CMBUSERID"
      Tab(1).Control(4)=   "CHEKUSERPASS"
      Tab(1).Control(5)=   "LBLPASSWORD"
      Tab(1).Control(6)=   "LBLUSER"
      Tab(1).Control(7)=   "Label2"
      Tab(1).ControlCount=   8
      Begin VB.TextBox TXTADMINID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   840
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   "INPUT ADMINISTRATOR ID"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox TXTADMINPASS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   4200
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "INPUT PASSWORD"
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton CMDADMINLOGIN 
         BackColor       =   &H00FFFF80&
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
         Left            =   1440
         Picture         =   "Form1.frx":241B89
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton CMDEXITADMINLOGIN 
         BackColor       =   &H000000C0&
         Caption         =   "EXIT"
         Height          =   300
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox CHEKADMINPASS 
         BackColor       =   &H00404080&
         Caption         =   "SHOW PASSWORD"
         ForeColor       =   &H8000000B&
         Height          =   300
         Left            =   3120
         TabIndex        =   4
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox TXTUSERPASS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Prestige Elite Std"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   -71640
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         ToolTipText     =   "INPUT PASSORD"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton CMDUSERLOGIN 
         BackColor       =   &H00FFFF80&
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
         Left            =   -74400
         Picture         =   "Form1.frx":24489F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton CMDEXITUSERLOGIN 
         BackColor       =   &H000000C0&
         Caption         =   "EXIT"
         Height          =   300
         Left            =   -69600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox CMBUSERID 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Prestige Elite Std"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox CHEKUSERPASS 
         BackColor       =   &H00004080&
         Caption         =   "SHOW PASSWORD"
         ForeColor       =   &H8000000B&
         Height          =   300
         Left            =   -72720
         TabIndex        =   9
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label LBLADMINID 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ADMINISTRATOR"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label LBLADMINPASS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Height          =   2535
         Left            =   -1680
         TabIndex        =   16
         Top             =   360
         Width           =   10695
      End
      Begin VB.Label LBLPASSWORD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PASSWORD"
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
         Left            =   -71640
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label LBLUSER 
         BackColor       =   &H00FFC0C0&
         Caption         =   "USER"
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
         Left            =   -74520
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   2535
         Left            =   -75000
         TabIndex        =   13
         Top             =   360
         Width           =   9975
      End
   End
   Begin VB.Shape LGShape1 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00FFC0C0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   20
      Height          =   2295
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   7455
   End
End
Attribute VB_Name = "FRMLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CMDADMINLOGIN_Click()
 If T = "ADMIN" Then
    TXTADMINPASS.Enabled = True
    TXTADMINPASS.SetFocus
Else
    MsgBox "YOU ARE NOT ALLOWED HERE PLEASE INPUT ADMIN  ID", vbCritical
    TXTADMINID.Text = ""
    TXTADMINID.SetFocus
End If
If TXTADMINID.Text = UCase(TXTADMINID.Text) And TXTADMINPASS.Text = P Then
USER = UCase(TXTADMINID.Text)
Unload Me
Load MDIForm1
MDIForm1.Show
Load FRMSIDEBAR
FRMSIDEBAR.Show
Else
MsgBox "WRONG PASSWORD", vbCritical
TXTADMINPASS.Text = ""
TXTADMINPASS.Enabled = True
TXTADMINPASS.SetFocus
CMDADMINLOGIN.Enabled = False
End If
End Sub
Private Sub CMDEXITADMINLOGIN_Click()
If MsgBox("ARE YOU SURE TO  EXIT ?", vbYesNo) = vbYes Then
End
Else
CMDEXITADMINLOGIN.SetFocus
End If
End Sub




Private Sub TXTADMINID_Change()
 If Len(TXTADMINID.Text) > 0 Then
CMDADMINLOGIN.Enabled = True
Else
CMDADMINLOGIN.Enabled = False
End If
End Sub

Private Sub TXTADMINID_LostFocus()
If TXTADMINID.Text = "" Then
    If MsgBox(" ENTER THE ADMIN ID ?", vbYesNo) = vbNo Then
        CMDEXITADMINLOGIN.SetFocus
    Else
        TXTADMINID.SetFocus
    End If
Else
    TXTADMINID.Text = UCase(TXTADMINID.Text)
    CON
    S = "SELECT *FROM ACOUNT WHERE USERID='" & TXTADMINID.Text & "'"
    Set R = C.Execute(S)
    If R.EOF = True Then
        If MsgBox("WRONG ID,ARE YOU RE-ENTER THE ADMIN ID ?", vbYesNo) = vbNo Then
            CMDEXITADMINLOGIN.SetFocus
        Else
            TXTADMINID.Text = ""
            TXTADMINID.SetFocus
        End If
    Else
        P = R.Fields("PASSWORD")
        T = R.Fields("TYPE")
        End If
    End If
End Sub

Private Sub TXTADMINPASS_Change()
If Len(TXTADMINPASS.Text) > 0 Then
CMDADMINLOGIN.Enabled = True
Else
CMDADMINLOGIN.Enabled = False
End If
End Sub

Private Sub TXTADMINPASS_LostFocus()
'TXTADMINPASS.Enabled = False
TXTADMINPASS.Text = UCase(TXTADMINPASS.Text)
End Sub



Private Sub TXTUSERPASS_Change()
If Len(TXTUSERPASS.Text) > 0 Then
CMDUSERLOGIN.Enabled = True
Else
CMDUSERLOGIN.Enabled = False
End If
End Sub

Private Sub TXTUSERPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If CMBUSERID.Text = "" Then
    If MsgBox("PLEASE ENTER USER ID ?", vbYesNo) = vbYes Then
        CMBUSERID.SetFocus
    Else
        CMDEXITUSERLOGIN.SetFocus
    End If
Else
    TXTUSERPASS.Text = UCase(TXTUSERPASS.Text)
    CMDUSERLOGIN.Enabled = True
    CMDUSERLOGIN.SetFocus
End If
End If
End Sub

Private Sub TXTUSERPASS_LostFocus()
If CMBUSERID.Text = "" Then
    If MsgBox("PLEASE ENTER THER USER ID ?", vbYesNo) = vbYes Then
        CMBUSERID.SetFocus
    Else
        CMDEXITUSERLOGIN.SetFocus
    End If
Else
    TXTUSERPASS.Text = UCase(TXTUSERPASS.Text)
    CMDUSERLOGIN.Enabled = True
    CMDUSERLOGIN.SetFocus
End If
End Sub

Private Sub CHEKADMINPASS_Click()
If CHEKADMINPASS.Value = 1 Then
TXTADMINPASS.PasswordChar = ""
Else
TXTADMINPASS.PasswordChar = "*"
End If
End Sub

Private Sub CMDUSERLOGIN_Click()
If CMBUSERID.Text = UCase(CMBUSERID.Text) And TXTUSERPASS.Text = P Then
USER = UCase(CMBUSERID.Text)
Unload Me
Load MDIForm1
MDIForm1.Show
MDIForm1.MNUEMPDETAIL.Enabled = False
MDIForm1.FRMLGDETAIL.Enabled = False
Else
MsgBox "WRONG PASSWORD", vbCritical
TXTUSERPASS.Text = ""
TXTUSERPASS.SetFocus
CMDUSERLOGIN.Enabled = False
End If
End Sub

Private Sub CMDEXITUSERLOGIN_Click()
If MsgBox("ARE YOU SURE TO  EXIT ?", vbYesNo) = vbYes Then
End
Else
CMDEXITUSERLOGIN.SetFocus
End If
End Sub
Private Sub CHEKUSERPASS_Click()
If CHEKUSERPASS.Value = 1 Then
TXTUSERPASS.PasswordChar = ""
Else
TXTUSERPASS.PasswordChar = "*"
End If
End Sub
Private Sub Form_Load()
CMDUSERLOGIN.Enabled = False
CMDADMINLOGIN.Enabled = False
TXTUSERPASS.Text = ""
TXTADMINID.Text = ""
TXTADMINPASS.Text = ""
CMBUSERID.Clear
CON
S = "SELECT *FROM ACOUNT "
Set R = C.Execute(S)
Do Until R.EOF
If R.Fields("TYPE") = "USER" Then
CMBUSERID.AddItem R("USERID")
End If
R.MoveNext
Loop
End Sub
'****** TO MOVE THE LABEL IN WHICH THE NAME OF HOTEL *******
Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 100
If Label1.Left + Label1.Width <= 0 Then
Label1.Left = Picture1.Width
End If
End Sub
Private Sub CMBUSERID_Click()
CON
S = "SELECT PASSWORD FROM ACOUNT WHERE USERID='" & CMBUSERID.Text & "'"
Set R = C.Execute(S)
P = R.Fields("PASSWORD")
TXTUSERPASS.Enabled = True
TXTUSERPASS.SetFocus
End Sub
Private Sub TXTADMINID_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("z") Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13 Then
    If KeyAscii = 91 Or KeyAscii = 92 Or KeyAscii = 93 Or KeyAscii = 94 Or KeyAscii = 95 Or KeyAscii = 96 Then
        KeyAscii = 0
    End If
Else
KeyAscii = 0
End If
If KeyAscii = 13 Or KeyAscii = 9 Then
    If TXTADMINID.Text = "" Then
        If MsgBox("WRONG ID, ARE YOU RE-ENTER THE ADMIN ID ?", vbYesNo) = vbNo Then
            CMDEXITADMINLOGIN.SetFocus
        Else
            TXTADMINID.Text = ""
            TXTADMINID.SetFocus
        End If
    Else
        TXTADMINID.Text = UCase(TXTADMINID.Text)
        CON
        S = "SELECT *FROM ACOUNT WHERE USERID='" & TXTADMINID.Text & "'"
        Set R = C.Execute(S)
        If R.EOF = True Then
            MsgBox "PLEASE INPUT VALID USER ID", vbCritical
            TXTADMINID.Text = ""
            TXTADMINID.SetFocus
        Else
            P = R.Fields("PASSWORD")
            T = R.Fields("TYPE")
            If T = "ADMIN" Then
                TXTADMINPASS.Enabled = True
                TXTADMINPASS.SetFocus
            Else
                MsgBox "YOU ARE NOT ALLOWED HERE PLEASE INPUT ADMIN  ID", vbCritical
                TXTADMINID.Text = ""
                TXTADMINID.SetFocus
            End If
        End If
    End If
End If
End Sub

Private Sub TXTADMINPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
    TXTADMINPASS.Text = UCase(TXTADMINPASS.Text)
    If TXTADMINPASS.Text = "" Then
        If MsgBox("YOU HAVE DO NOT ENTER THE PASSWORD, IF YOU WANT TO EXIT THEN CLICK ON 'YES' OTHERWISE IF YOU WANT TO ENTER PASSWORD THEN CLICK 'NO'", vbYesNo) = vbYes Then
            CMDEXITADMINLOGIN.SetFocus
        Else
            TXTADMINPASS.SetFocus
        End If
    Else
        CMDADMINLOGIN.Enabled = True
        CMDADMINLOGIN.SetFocus
    End If
End If
End Sub
