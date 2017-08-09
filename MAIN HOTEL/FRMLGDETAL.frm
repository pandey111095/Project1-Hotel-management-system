VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRMLGDETAL 
   BorderStyle     =   0  'None
   Caption         =   "LOGON DETAIL  FORM"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMLGDETAL.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7011
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ADD USER"
      TabPicture(0)   =   "FRMLGDETAL.frx":240042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LBLCREUSERID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LBLCREUSERPASS"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LBLCRECONFIRMPASS"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LBLCREUSERNAME"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TXTCREUSERPASS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CMDCREATEUSER"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXTCREUSERCONFIRMPASS"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CMDCREBACK"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CHEKCREPASS"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CMBCREUSERID"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "CHANGE PASSWORD"
      TabPicture(1)   =   "FRMLGDETAL.frx":24005E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CHEKCHANGEPASS"
      Tab(1).Control(1)=   "CMDCHANGEPASS"
      Tab(1).Control(2)=   "CMBCHANGEPASS"
      Tab(1).Control(3)=   "TXTCHANGENEWPASS"
      Tab(1).Control(4)=   "TXTCHANGEOLDPASS"
      Tab(1).Control(5)=   "TXTCHANGEUSERID"
      Tab(1).Control(6)=   "CMDCHANGEBACK"
      Tab(1).Control(7)=   "LBLCHANGENEWPASS"
      Tab(1).Control(8)=   "LBLCHANGEOLDPASS"
      Tab(1).Control(9)=   "LBLCHANGEUSERID"
      Tab(1).Control(10)=   "Label2"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "DELETE USER"
      TabPicture(2)   =   "FRMLGDETAL.frx":24007A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CHEKDELUSER"
      Tab(2).Control(1)=   "CMDDELBACK"
      Tab(2).Control(2)=   "CMDDELETEUSER"
      Tab(2).Control(3)=   "CMBDELUSER"
      Tab(2).Control(4)=   "TXTDELPASS"
      Tab(2).Control(5)=   "TXTDELUSERID"
      Tab(2).Control(6)=   "LBLDELPASS"
      Tab(2).Control(7)=   "LBLDELUSEID"
      Tab(2).Control(8)=   "Label3"
      Tab(2).ControlCount=   9
      Begin VB.ComboBox CMBCREUSERID 
         Height          =   315
         Left            =   2640
         TabIndex        =   30
         Text            =   "Combo1"
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox CHEKDELUSER 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SHOW PASSWORD"
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
         Left            =   -70320
         TabIndex        =   27
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CommandButton CMDDELBACK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EXIT"
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
         Left            =   -72240
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton CMDDELETEUSER 
         BackColor       =   &H00FFC0C0&
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
         Height          =   375
         Left            =   -74520
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox CMBDELUSER 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -67200
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TXTDELPASS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   -69960
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox TXTDELUSERID 
         Appearance      =   0  'Flat
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
         Height          =   480
         Left            =   -74640
         TabIndex        =   22
         Top             =   1680
         Width           =   4695
      End
      Begin VB.CheckBox CHEKCHANGEPASS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SHOW PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69600
         TabIndex        =   17
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CommandButton CMDCHANGEPASS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "UPDATE"
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
         Left            =   -74160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox CMBCHANGEPASS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -66600
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TXTCHANGENEWPASS 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   -71400
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   2400
         Width           =   4695
      End
      Begin VB.TextBox TXTCHANGEOLDPASS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   -71400
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox TXTCHANGEUSERID 
         Appearance      =   0  'Flat
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
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   4695
      End
      Begin VB.CommandButton CMDCHANGEBACK 
         BackColor       =   &H00FFC0C0&
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
         Height          =   495
         Left            =   -71760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CheckBox CHEKCREPASS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SHOW PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   6
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton CMDCREBACK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EXIT"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox TXTCREUSERCONFIRMPASS 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2280
         Width           =   4695
      End
      Begin VB.CommandButton CMDCREATEUSER 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CREATE"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox TXTCREUSERPASS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label LBLCREUSERNAME 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   6000
         TabIndex        =   31
         Top             =   840
         Width           =   105
      End
      Begin VB.Label LBLDELPASS 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   -69960
         TabIndex        =   29
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label LBLDELUSEID 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " USER ID"
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
         Left            =   -74640
         TabIndex        =   28
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Enabled         =   0   'False
         Height          =   3735
         Left            =   -75000
         TabIndex        =   21
         Top             =   240
         Width           =   9975
      End
      Begin VB.Label LBLCHANGENEWPASS 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW PASSWORD"
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
         Left            =   -74160
         TabIndex        =   20
         Top             =   2400
         Width           =   2565
      End
      Begin VB.Label LBLCHANGEOLDPASS 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "OLD PASSWORD"
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
         Left            =   -74160
         TabIndex        =   19
         Top             =   1680
         Width           =   2460
      End
      Begin VB.Label LBLCHANGEUSERID 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " USER ID"
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
         Left            =   -74160
         TabIndex        =   18
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Enabled         =   0   'False
         Height          =   3735
         Left            =   -75000
         TabIndex        =   10
         Top             =   240
         Width           =   9975
      End
      Begin VB.Label LBLCRECONFIRMPASS 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRM PASSWORD"
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
         Left            =   720
         TabIndex        =   9
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label LBLCREUSERPASS 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   720
         TabIndex        =   8
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label LBLCREUSERID 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " USER ID"
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
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Enabled         =   0   'False
         Height          =   3735
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   9975
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   20
      Height          =   4695
      Left            =   1800
      Top             =   840
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   20
      Height          =   4215
      Left            =   2040
      Top             =   1080
      Width           =   10215
   End
End
Attribute VB_Name = "FRMLGDETAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim U As String
Private Sub CHEKCREPASS_Click()
If CHEKCREPASS.Value = 1 Then
TXTCREUSERPASS.PasswordChar = ""
TXTCREUSERCONFIRMPASS.PasswordChar = ""
Else
TXTCREUSERPASS.PasswordChar = "*"
TXTCREUSERCONFIRMPASS.PasswordChar = "*"
End If
End Sub
Private Sub CHEKCHANGEPASS_Click()
If CHEKCHANGEPASS.Value = 1 Then
TXTCHANGEOLDPASS.PasswordChar = ""
TXTCHANGENEWPASS.PasswordChar = ""
Else
TXTCHANGEOLDPASS.PasswordChar = "*"
TXTCHANGENEWPASS.PasswordChar = "*"
End If
End Sub
Private Sub CHEKDELUSER_Click()
If CHEKDELUSER.Value = 1 Then
TXTDELPASS.PasswordChar = ""
Else
TXTDELPASS.PasswordChar = "*"
End If
End Sub
Private Sub CMBCHANGEPASS_Click()
TXTCHANGEOLDPASS.SetFocus
TXTCHANGEUSERID.Text = CMBCHANGEPASS.Text
End Sub

Private Sub CMBCREUSERID_Click()
S = "SELECT EMP_NAME FROM EMP_RECORD WHERE EMP_ID='" & CMBCREUSERID.Text & "'"
Set R = C.Execute(S)
LBLCREUSERNAME.Caption = R.Fields("EMP_NAME")
End Sub

Private Sub CMBDELUSER_Click()
TXTDELPASS.SetFocus
TXTDELUSERID.Text = CMBDELUSER.Text
End Sub
Private Sub CMDCREATEUSER_Click()
If CMBCREUSERID.Text = "" Then
    MsgBox "PLEASE SELECT USERID"
    CMBCREUSERID.SetFocus
ElseIf TXTCREUSERPASS.Text = "" Then
    MsgBox "PLEASE INPUT PASSWORD"
    TXTCREUSERPASS.SetFocus
ElseIf TXTCREUSERCONFIRMPASS.Text = "" Then
    MsgBox "PLEASE INPUT CONFIRM PASSWORD"
    TXTCREUSERCONFIRMPASS.SetFocus
Else
    If MsgBox("ARE YOU  SURE TO ADD NEW USER", vbYesNo) = vbYes Then
        CON
        S = "INSERT INTO ACOUNT VALUES('" & CMBCREUSERID.Text & "','" & TXTCREUSERPASS.Text & "','USER')"
        Set R = C.Execute(S)
        S = "UPDATE EMP_RECORD SET DEPARTMENT='OPERATOR' WHERE EMP_ID='" & CMBCREUSERID.Text & "'"
        Set R = C.Execute(S)
        S = "COMMIT"
        Set R = C.Execute(S)
        MsgBox "NEW USER CREATED"
        'Combo1.AddItem  Text1.Text
        'Combo2.AddItem = Text1.Text
        'TXTCREUSERID.Text = ""
        TXTCREUSERPASS.Text = ""
        TXTCREUSERCONFIRMPASS.Text = ""
        CMDCREATEUSER.Enabled = False
        CMDCREBACK.SetFocus
        CMBCHANGEPASS.Clear
        CMBDELUSER.Clear
        Call Form_Load
        Else
        TXTCREUSERID.Text = ""
        TXTCREUSERPASS.Text = ""
        TXTCREUSERCONFIRMPASS.Text = ""
        CMDCREATEUSER.Enabled = False
        CMDCREBACK.SetFocus
    End If
End If
'CON
'S = "COMMIT"
'Set R = C.Execute(S)
End Sub

Private Sub CMDCREBACK_Click()
Unload FRMLGDETAL
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub CMDCHANGEPASS_Click()
If MsgBox("ARE SURE TO CHANGE THE PASSWORD", vbYesNo) = vbYes Then
CON
S = "UPDATE ACOUNT  SET PASSWORD='" & TXTCHANGENEWPASS.Text & "' WHERE USERID='" & TXTCHANGEUSERID.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
Call Form_Load
MsgBox "PASSWORD CHANGED"
CMDCHANGEPASS.Enabled = False
CMDCHANGEBACK.SetFocus
TXTCHANGEUSERID.Text = ""
TXTCHANGEOLDPASS.Text = ""
TXTCHANGENEWPASS.Text = ""
Else
MsgBox "PASSWORD NOT CHANGED"
CMDCHANGEPASS.Enabled = False
CMDCHANGEBACK.SetFocus
TXTCHANGEUSERID.Text = ""
TXTCHANGEOLDPASS.Text = ""
TXTCHANGENEWPASS.Text = ""
End If
End Sub

Private Sub CMDCHANGEBACK_Click()
Unload FRMLGDETAL
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub CMDDELETEUSER_Click()
If MsgBox("ARE YOU SURE TO DELETE USER", vbYesNo) = vbYes Then
CON
S = "DELETE FROM ACOUNT WHERE USERID='" & TXTDELUSERID.Text & "'"
Set R = C.Execute(S)
MsgBox "USER DELETED"
CMDDELETEUSER.Enabled = False
TXTDELUSERID.Text = ""
TXTDELPASS.Text = ""
CMDDELBACK.SetFocus
CMBCHANGEPASS.Clear
CMBDELUSER.Clear
Call Form_Load
Else
TXTDELUSERID.Text = ""
TXTDELPASS.Text = ""
CMDDELETEUSER.Enabled = False
CMDDELBACK.SetFocus
End If
End Sub

Private Sub CMDDELBACK_Click()
Unload FRMLGDETAL
Load MDIForm1
MDIForm1.Show
End Sub
Private Sub Form_Load()
CMBCREUSERID.Clear
CMDCREATEUSER.Enabled = False
CMDCHANGEPASS.Enabled = False
CMDDELETEUSER.Enabled = False
CMBCHANGEPASS.Clear
CMBDELUSER.Clear
CON
S = "SELECT *FROM ACOUNT WHERE TYPE='" & "USER" & "'"
Set R = C.Execute(S)
Do Until R.EOF
CMBCHANGEPASS.AddItem R("USERID")
CMBDELUSER.AddItem R("USERID")
R.MoveNext
Loop
S = "SELECT EMP_ID FROM EMP_RECORD WHERE DEPARTMENT<>'OPERATOR'"
Set R = C.Execute(S)
Do Until R.EOF = True
CMBCREUSERID.AddItem R.Fields("EMP_ID")
R.MoveNext
Loop
End Sub







Private Sub TXTCREUSERPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TXTCREUSERPASS.Text = UCase(TXTCREUSERPASS.Text)
TXTCREUSERCONFIRMPASS.SetFocus
End If
End Sub

Private Sub TXTCREUSERCONFIRMPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TXTCREUSERCONFIRMPASS.Text = UCase(TXTCREUSERCONFIRMPASS.Text)
    If TXTCREUSERPASS.Text = TXTCREUSERCONFIRMPASS.Text Then
        CMDCREATEUSER.Enabled = True
        CMDCREATEUSER.SetFocus
    Else
        MsgBox "PLESE RE-ENTER YOUR PASSWORD", vbCritical
        TXTCREUSERCONFIRMPASS.Text = ""
        TXTCREUSERCONFIRMPASS.SetFocus
    End If
End If
End Sub
Private Sub TXTCHANGEOLDPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TXTCHANGEOLDPASS.Text = UCase(TXTCHANGEOLDPASS.Text)
CON
S = "SELECT PASSWORD FROM ACOUNT WHERE USERID='" & TXTCHANGEUSERID.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
MsgBox "PLEASE INPUT VALID USER ID"
TXTCHANGEUSERID.Text = ""
TXTCHANGEOLDPASS.Text = ""
CMDEXITCHANGE.SetFocus
Else
If TXTCHANGEOLDPASS.Text = R.Fields("PASSWORD") Then
TXTCHANGENEWPASS.SetFocus
Else
MsgBox "WRONG PASSWORD", vbCritical
TXTCHANGEOLDPASS.Text = ""
TXTCHANGEOLDPASS.SetFocus
End If
End If
End If
End Sub
Private Sub TXTCHANGENEWPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If TXTCHANGEUSERID.Text = "" Or TXTCHANGEOLDPASS.Text = "" Or TXTCHANGENEWPASS.Text = "" Then
MsgBox " PLESE INPUT USERID AND PASSWORD"
Else
TXTCHANGENEWPASS.Text = UCase(TXTCHANGENEWPASS.Text)
CMDCHANGEPASS.Enabled = True
CMDCHANGEPASS.SetFocus
End If
End If
End Sub
Private Sub TXTDELPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TXTDELPASS.Text = UCase(TXTDELPASS.Text)
CON
S = "SELECT PASSWORD FROM ACOUNT WHERE USERID='" & TXTDELUSERID.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
MsgBox "PLEASE INPUT VALID USER ID"
'Text4.Text = ""
'Text4.SetFocus
Else
If TXTDELPASS.Text = R.Fields("PASSWORD") Then
CMDDELETEUSER.Enabled = True
CMDDELETEUSER.SetFocus
Else
MsgBox "WRONG PASSWORD", vbCritical
TXTDELPASS.Text = ""
TXTDELPASS.SetFocus
End If
End If
End If
End Sub
