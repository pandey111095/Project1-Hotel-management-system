VERSION 5.00
Begin VB.Form FRMCHANGEPASSWORD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMCHANGEPASSWORD.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
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
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2775
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
      Height          =   345
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
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
      Height          =   360
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
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
      Left            =   3000
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
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
      Left            =   3000
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton CMDCHANGEBACK 
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
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label LBLCHANGENEWPASS 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   2565
   End
   Begin VB.Label LBLCHANGEOLDPASS 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2460
   End
   Begin VB.Label LBLCHANGEUSERID 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "FRMCHANGEPASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CHEKCHANGEPASS_Click()
If CHEKCHANGEPASS.Value = 1 Then
TXTCHANGEOLDPASS.PasswordChar = ""
TXTCHANGENEWPASS.PasswordChar = ""
Else
TXTCHANGEOLDPASS.PasswordChar = "*"
TXTCHANGENEWPASS.PasswordChar = "*"
End If
End Sub



Private Sub CMDCHANGEBACK_Click()
Unload Me
Me.Hide
End Sub


Private Sub CMDCHANGEPASS_Click()
If MsgBox("ARE SURE TO CHANGE THE PASSWORD", vbYesNo) = vbYes Then
CON
S = "UPDATE ACOUNT  SET PASSWORD='" & TXTCHANGENEWPASS.Text & "' WHERE USERID='" & TXTCHANGEUSERID.Text & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
P = TXTCHANGENEWPASS.Text
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


Private Sub Form_Load()
Me.Top = 1000
Me.Left = 5000
CMDCHANGEPASS.Enabled = False
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


Private Sub TXTCHANGEOLDPASS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TXTCHANGEOLDPASS.Text = UCase(TXTCHANGEOLDPASS.Text)
If P = TXTCHANGEOLDPASS.Text Then
TXTCHANGENEWPASS.SetFocus
Else
MsgBox "WRONG PASSWORD", vbCritical
TXTCHANGEOLDPASS.Text = ""
TXTCHANGEOLDPASS.SetFocus
End If
End If
End Sub

Private Sub TXTCHANGEUSERID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TXTCHANGEUSERID.Text = UCase(TXTCHANGEUSERID.Text)
If TXTCHANGEUSERID.Text = USER Then
TXTCHANGEOLDPASS.SetFocus
Else
MsgBox "YOU CAN NOT  CHANGE  PASSWORD OF THIS USER"
TXTCHANGEUSERID.Text = ""
End If
End If
End Sub
