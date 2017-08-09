VERSION 5.00
Begin VB.Form FRMROOMDETAIL 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ROOM DETAIL"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BACKTOMDIFROMROOM_CMD 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton GOROOM_CMD 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox GORNO_TXT 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton MODIFYROOM_CMD 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MODIFY"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton ADDNEWROOM_CMD 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&ADDNEW"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox NOTBOOKEDROOM_TXT 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
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
      Height          =   360
      Left            =   10200
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox BOOKEDROOM_TXT 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
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
      Height          =   360
      Left            =   6000
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox TOTNOROOM_TXT 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
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
      Height          =   330
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame ROOM_FRAME 
      BackColor       =   &H8000000A&
      Caption         =   "ROOM"
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
      Height          =   2415
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   6375
      Begin VB.TextBox RNO_TXT 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
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
         Height          =   240
         Left            =   4800
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox RSTATUS_CMB 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMLAUND.frx":0000
         Left            =   2640
         List            =   "FRMLAUND.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox RSTATUS_TXT 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
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
         Height          =   225
         Left            =   2640
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox RTYPE_CMB 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMLAUND.frx":0021
         Left            =   2160
         List            =   "FRMLAUND.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox RCOST_TXT 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox RTYPE_TXT 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
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
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label RNO_LBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM NO:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3480
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label RSTATUS_LBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROOM STATUS:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   480
         TabIndex        =   22
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label RCOST_LBL 
         BackStyle       =   0  'Transparent
         Caption         =   "COST:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label RTYPE_LBL 
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton SAVEROOM_CMD 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&SAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton UPDATEROOM_CMD 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&UPDATE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404080&
      BorderWidth     =   15
      Height          =   1935
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label GORNO_LBL 
      BackColor       =   &H8000000A&
      Caption         =   "ROOM NO:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL NOT BOOKED ROOM:-"
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " BOOKED ROOM:-"
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL NO OF ROOM:-"
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "FRMROOMDETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As String

Private Sub ADDNEWROOM_CMD_Click()
RSTATUS_CMB.Visible = True
RSTATUS_TXT.Visible = False
RSTATUS_CMB.AddItem "NOTBOOKED"
RSTATUS_CMB.AddItem "BOOKED"
MODIFYROOM_CMD.Enabled = False
ROOM_FRAME.Enabled = True
RNO_TXT.Text = Val(TOTNOROOM_TXT.Text) + 1
'RTYPE_TXT.SetFocus
ADDNEWROOM_CMD.Enabled = False
SAVEROOM_CMD.Enabled = True
End Sub

Private Sub BACKTOMDIFROMROOM_CMD_Click()
If MsgBox("ARE YOU SURE TO EXIT", vbYesNo) = vbYes Then
    Unload Me
    FRMROOMDETAIL.Hide
    MDIForm1.Show
Else
    Call Form_Load
End If
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Click()

End Sub

Private Sub Form_Load()
Me.Top = 1000
Me.Left = 5000
B = "BOOKED"
CON
S = "SELECT COUNT(*) FROM ROOM"
Set R = C.Execute(S)
TOTNOROOM_TXT.Text = R.Fields(0)
S = "SELECT COUNT(*) FROM ROOM WHERE STATUS='" & B & "'"
Set R = C.Execute(S)
BOOKEDROOM_TXT.Text = R.Fields(0)
NOTBOOKEDROOM_TXT.Text = Val(TOTNOROOM_TXT.Text) - Val(BOOKEDROOM_TXT.Text)
End Sub

Private Sub GORNO_TXT_Change()
If Len(GORNO_TXT.Text) > 0 Then
    GOROOM_CMD.Enabled = True
Else
    GOROOM_CMD.Enabled = False
End If
End Sub

Private Sub GORNO_TXT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
 GOROOM_CMD.SetFocus
End If
End Sub

Private Sub GORNO_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    GOROOM_CMD.SetFocus
End If
End Sub

Private Sub GOROOM_CMD_Click()
CON
S = "SELECT *FROM ROOM WHERE R_NO='" & GORNO_TXT.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
    MsgBox "IT IS NOT EXIST ", vbCritical
    If MsgBox("ARE YOU WANT TO RE-ENTER THE ROOM NO:-", vbYesNo) = vbYes Then
        GORNO_TXT.Text = ""
        RTYPE_TXT.Text = ""
        RCOST_TXT.Text = ""
        RNO_TXT.Text = ""
        RSTATUS_TXT.Text = ""
        GORNO_TXT.SetFocus
    Else
        GORNO_TXT.Text = ""
        RTYPE_TXT.Text = ""
        RCOST_TXT.Text = ""
        RNO_TXT.Text = ""
        RSTATUS_TXT.Text = ""
        GORNO_TXT.Visible = False
        GORNO_LBL.Visible = False
        GOROOM_CMD.Visible = False
        RSTATUS_TXT.Visible = False
        ROOM_FRAME.Enabled = False
        SAVEROOM_CMD.Enabled = False
        RSTATUS_CMB.Visible = False
        ADDNEWROOM_CMD.Enabled = True
        MODIFYROOM_CMD.Enabled = True
        Call Form_Load
        BACKTOMDIFROMROOM_CMD.SetFocus
    End If
Else
    ROOM_FRAME.Enabled = True
    RNO_TXT.Text = R.Fields("R_NO")
    RSTATUS_TXT.Text = R.Fields("STATUS")
    RTYPE_TXT.Text = R.Fields("TYPE")
    S = "SELECT *FROM ROOM_TYPE WHERE TYPE='" & RTYPE_TXT.Text & "'"
    Set R = C.Execute(S)
    If R.EOF = True Then
    MsgBox "TYPE  DOES NOT EXIST"
    Else
    RCOST_TXT.Text = R.Fields("COST")
    End If
    RNO_TXT.Enabled = False
'    RTYPE_TXT.SetFocus
    UPDATEROOM_CMD.Enabled = True
    GOROOM_CMD.Enabled = False
End If
End Sub

Private Sub MODIFYROOM_CMD_Click()
GORNO_TXT.Visible = True
GORNO_LBL.Visible = True
GORNO_TXT.SetFocus
GOROOM_CMD.Visible = True
RSTATUS_CMB.Visible = False
RSTATUS_TXT.Visible = True
MODIFYROOM_CMD.Enabled = False
ADDNEWROOM_CMD.Enabled = False
SAVEROOM_CMD.Enabled = False
End Sub

Private Sub RCOST_TXT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
 R_STATUS_CMB.SetFocus
End If
End Sub

Private Sub RCOST_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
If KeyAscii = 13 Then
    RSTATUS_CMB.SetFocus
End If
End Sub

Private Sub RSTATUS_CMB_Click()
RSTATUS_TXT.Text = RSTATUS_CMB.Text
End Sub

Private Sub RSTATUS_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SAVEROOM_CMD.SetFocus
End If
End Sub

Private Sub RTYPE_CMB_Click()
RTYPE_TXT.Text = RTYPE_CMB.Text
RTYPE_CMB.Refresh
RCOST_TXT.SetFocus
End Sub

Private Sub RTYPE_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    RCOST_TXT.SetFocus
End If
End Sub


Private Sub SAVEROOM_CMD_Click()
If RTYPE_TXT.Text = "" Then
MsgBox "PLEASE SELECT ROOM TYPE"
RTYPE_CMB.SetFocus
ElseIf RCOST_TXT.Text = "" Then
MsgBox "PLEASE INPUT ROOM CHARGE"
RCOST_TXT.SetFocus
ElseIf RSTATUS_TXT.Text = "" Then
MsgBox "PLEASE SELECT ROOM STATUS "
RSTATUS_CMB.SetFocus
Else
If MsgBox("ARE YOU SURE TO SAVE THE ROOM DETAIL ?", vbYesNo) = vbYes Then
    CON
    S = "INSERT INTO ROOM VALUES(" & RNO_TXT.Text & ",'" & RTYPE_TXT.Text & "','" & RSTATUS_CMB.Text & "')"
    Set R = C.Execute(S)
    S = "INSERT INTO ROOM_TYPE VALUES('" & RTYPE_TXT.Text & "'," & RCOST_TXT.Text & "," & RNO_TXT.Text & ")"
    Set R = C.Execute(S)
    S = "COMMIT"
    Set R = C.Execute(S)
    MsgBox "RECORD SAVED"
    ROOM_FRAME.Enabled = False
    SAVEROOM_CMD.Enabled = False
    RSTATUS_CMB.Visible = False
    RNO_TXT.Text = ""
    RTYPE_TXT.Text = ""
    RCOST_TXT.Text = ""
    RSTATUS_TXT.Text = ""
    ADDNEWROOM_CMD.Enabled = True
    MODIFYROOM_CMD.Enabled = True
    Call Form_Load
Else
    MsgBox "RECORD NOT SAVED"
    SAVEROOM_CMD.Enabled = False
    RSTATUS_CMB.Visible = False
    ADDNEWROOM_CMD.Enabled = True
    MODIFYROOM_CMD.Enabled = True
    RNO_TXT.Text = ""
    RTYPE_TXT.Text = ""
    RCOST_TXT.Text = ""
    RSTATUS_TXT.Text = ""
    Call Form_Load
End If
End If
End Sub

Private Sub UPDATEROOM_CMD_Click()
If RTYPE_TXT.Text = "" Then
    MsgBox "PLESE INPUT TYPE OF ROOM"
    RTYPE_TXT.SetFocus
ElseIf RCOST_TXT.Text = "" Then
    MsgBox "PLESE INPUT CHARGE OF THIS ROOM"
    RCOST_TXT.SetFocus
ElseIf RSTATUS_TXT.Text = "" Then
    MsgBox "PLESE INPUT ROOM STATUS"
    RSTATUS_TXT.SetFocus
Else
    If MsgBox("ARE SURE TO UPDATE THIS ROOM", vbYesNo) = vbYes Then
    CON
    S = "UPDATE ROOM SET TYPE='" & RTYPE_TXT.Text & "', STATUS='" & RSTATUS_TXT.Text & "' WHERE R_NO=" & RNO_TXT.Text & ""
    Set R = C.Execute(S)
    S = "UPDATE ROOM_TYPE SET COST=" & RCOST_TXT.Text & ", TYPE='" & RTYPE_TXT.Text & "' WHERE R_NO='" & RNO_TXT.Text & "'"
    Set R = C.Execute(S)
    S = "COMMIT"
    Set R = C.Execute(S)
    MsgBox "NOW THIS ROOM IS UPDATED"
    GORNO_TXT.Text = ""
    RTYPE_TXT.Text = ""
    RCOST_TXT.Text = ""
    RNO_TXT.Text = ""
    RSTATUS_TXT.Text = ""
    GORNO_LBL.Visible = False
    GORNO_TXT.Visible = False
    GOROOM_CMD.Visible = False
    UPDATEROOM_CMD.Enabled = False
    ADDNEWROOM_CMD.Enabled = True
    MODIFYROOM_CMD.Enabled = True
    ROOM_FRAME.Enabled = False
    Call Form_Load
    Else
    MsgBox "NOW THIS ROOM IS  NOT UPDATE"
    GORNO_TXT.Text = ""
    RTYPE_TXT.Text = ""
    RCOST_TXT.Text = ""
    RNO_TXT.Text = ""
    RSTATUS_TXT.Text = ""
    GORNO_LBL.Visible = False
    GORNO_TXT.Visible = False
    GOROOM_CMD.Visible = False
    UPDATEROOM_CMD.Enabled = False
    ADDNEWROOM_CMD.Enabled = True
    MODIFYROOM_CMD.Enabled = True
    ROOM_FRAME.Enabled = False
    Call Form_Load
End If
End If
End Sub
