VERSION 5.00
Begin VB.Form FRMRESTORENTGUEST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GUEST RESTAURANT "
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMRESTOREGUEST.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   14835
   Begin VB.CommandButton GUESTADDONDUES_CMD 
      Caption         =   "ADD ON DUES"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8520
      TabIndex        =   50
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton GUESTNEWORDER_CMD 
      Caption         =   "NEW OEDER"
      Height          =   495
      Left            =   2400
      TabIndex        =   49
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton GUESTADDORDER_CMD 
      Caption         =   "ADD ORDER"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   48
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton GUESTPAYMENT_CMD 
      Caption         =   "PAYMENT NOW"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   47
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton GUESTEXIT_CMD 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   9840
      TabIndex        =   46
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton GUESTREMOVEORDER_CMD 
      Caption         =   "REMOVE ORDER"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5160
      TabIndex        =   45
      Top             =   7440
      Width           =   1815
   End
   Begin VB.ListBox GUESTDISH_LST 
      Height          =   2595
      Left            =   480
      TabIndex        =   38
      Top             =   3240
      Width           =   1815
   End
   Begin VB.ListBox GUESTNOOFDISH_LST 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   2760
      TabIndex        =   37
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox GUESTDISHCHARGE_LST 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   4560
      TabIndex        =   36
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ListBox GUESTDRINK_LST 
      Height          =   2400
      Left            =   8040
      TabIndex        =   35
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ListBox GUESTNOOFDRINK_LST 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   10320
      TabIndex        =   34
      Top             =   3360
      Width           =   855
   End
   Begin VB.ListBox GUESTDRINKCHARGE_LST 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   11760
      TabIndex        =   33
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox GUESTDISHPRICEPERPLATE_TXT 
      Alignment       =   2  'Center
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
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   22
      Text            =   "0"
      Top             =   2400
      Width           =   735
   End
   Begin VB.ComboBox GUESTNONVEGDISH_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   480
      TabIndex        =   21
      Text            =   "NONVEG DISH"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox GUESTVEGDISH_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   20
      Text            =   "VEG DISH"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox GUESTSOFTDRINK_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8040
      TabIndex        =   19
      Text            =   "SOFT DRINK"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox GUESTJUICE_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9840
      TabIndex        =   18
      Text            =   "JUICE"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox GUESTBEER_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   11400
      TabIndex        =   17
      Text            =   "BEER"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox GUESTWINE_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   13200
      TabIndex        =   16
      Text            =   "WINE"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox GUESTDRINKPRICEPERGLASS_TXT 
      Alignment       =   2  'Center
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
      Left            =   9720
      MaxLength       =   2
      TabIndex        =   15
      Text            =   "0"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox GUESTNOOFDISH_TXT 
      Alignment       =   2  'Center
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
      Left            =   5880
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "0"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox GUESTNOOFDRINK_TXT 
      Alignment       =   2  'Center
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
      Left            =   12720
      MaxLength       =   2
      TabIndex        =   13
      Text            =   "0"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox GUESTNAME_TXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9000
      MaxLength       =   20
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox GUESTID_TXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7680
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox GUESTID_CMB 
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Text            =   "GUEST ID"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DISH CHARGE:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1200
      TabIndex        =   44
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label GUESTTOTALDISHCHARGE_LBL 
      Height          =   255
      Left            =   3600
      TabIndex        =   43
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DRINK CHARGE:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9360
      TabIndex        =   42
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label GUESTTOTALDRINKCHARGE_LBL 
      Height          =   255
      Left            =   11760
      TabIndex        =   41
      Top             =   6120
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderWidth     =   10
      X1              =   0
      X2              =   15000
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL CHARGE:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5520
      TabIndex        =   40
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label GUESTTOTALCHARGE_LBL 
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
      Left            =   8160
      TabIndex        =   39
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE PER  PLATE/PIECE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2400
      TabIndex        =   31
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CHARGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4800
      TabIndex        =   30
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE PER GLASS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   8040
      TabIndex        =   29
      Top             =   2520
      Width           =   1665
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   9960
      TabIndex        =   28
      Top             =   2880
      Width           =   1590
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHARGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   11880
      TabIndex        =   27
      Top             =   2880
      Width           =   780
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DISH NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DRINK NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   8040
      TabIndex        =   25
      Top             =   2880
      Width           =   1185
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   11040
      TabIndex        =   23
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Label CUSTDRINK_LBL 
      BackStyle       =   0  'Transparent
      Caption         =   "DRINKS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11040
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DISH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   0
      X2              =   15000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   10
      X1              =   7200
      X2              =   7200
      Y1              =   1440
      Y2              =   6480
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "NONVEG DISH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "VEG DISH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "SOFT DRINK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "JUICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9840
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "BEER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   11400
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "WINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   13200
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9000
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FRMRESTORENTGUEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GUEST As String
Dim DISHNO As Integer
Dim DRINKNO As Integer
Dim DISHNAME As String
Dim DRINKNAME As String
Dim L As Integer
Dim T As Integer



Private Sub GUESTADDONDUES_CMD_Click()
If Val(GUESTTOTALCHARGE_LBL.Caption) > 0 Then
If MsgBox("ARE YOU SURE TO ADD ON DUES ?", vbYesNo) = vbYes Then
'S = "UPDATE DAILY SET TOT_INCOME=TOT_INCOME+" & GUESTRESTOTOTCHARGE_TXT.Text & " WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
'MsgBox S
 '       Set R = C.Execute(S)
        S = "UPDATE CLIENT_MASTER SET DUES_BALANCE=DUES_BALANCE+" & Val(GUESTTOTALCHARGE_LBL.Caption) & ",TOT_CHARGE=TOT_CHARGE+" & Val(GUESTTOTALCHARGE_LBL.Caption) & ",RESTO_CHARGE=RESTO_CHARGE+" & Val(GUESTTOTALCHARGE_LBL.Caption) & " WHERE CLIENT_ID='" & GUESTID_CMB.Text & "'"
        Set R = C.Execute(S)
        S = "COMMIT"
        Set R = C.Execute(S)
    '    If MsgBox("ARE YOU TAKE REPORT THEN CLICK 'YES'", vbYesNo) = vbYes Then
     '       If DataEnvironment1.rsCommand5.State = 1 Then DataEnvironment1.rsCommand5.Close
      '          DataEnvironment1.Command5 GUESTID_CMB.Text
       '         DataReport7.Show
        '        GUESTBARPAYMENT.Enabled = False
        'End If
    'End If
'Else
 '   If MsgBox("ARE YOU TAKE REPORT THEN CLICK 'YES'", vbYesNo) = vbYes Then
  '          If DataEnvironment1.rsCommand5.State = 1 Then DataEnvironment1.rsCommand5.Close
   '             DataEnvironment1.Command5 GUESTID_CMB.Text
    '            DataReport7.Show
     '           GUESTBARPAYMENT.Enabled = False
      '      End If
'End If
Do Until DRINKNO = 0
S = "INSERT INTO CUST VALUES('" & GUESTID_TXT.Text & "','" & GUESTNAME_TXT.Text & "'," & Val(GUESTTOTALCHARGE_LBL.Caption) & ",'" & Format(Date, "DD-MMM-YYYY") & "','" & GUESTDRINK_LST.List(DRINKNO - 1) & "',''," & Val(GUESTDRINKCHARGE_LST.List(DRINKNO - 1)) & "," & 0 & ")"
MsgBox S
Set R = C.Execute(S)
DRINKNO = DRINKNO - 1
Loop
Do Until DISHNO = 0
S = "INSERT INTO CUST VALUES('" & GUESTID_TXT.Text & "','" & GUESTNAME_TXT.Text & "'," & Val(GUESTTOTALCHARGE_LBL.Caption) & ",'" & Format(Date, "DD-MMM-YYYY") & "','','" & GUESTDISH_LST.List(DISHNO - 1) & "'," & 0 & "," & GUESTDISHCHARGE_LST.List(DISHNO - 1) & ")"
MsgBox S
Set R = C.Execute(S)
DISHNO = DISHNO - 1
Loop
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "DUES ADDED !"
Unload Me
Load FRMRESTORENTGUEST
FRMRESTORENTGUEST.Show
End If
Else
MsgBox "PLEASE ADD ORDER FIRST"
End If
End Sub

Private Sub GUESTADDORDER_CMD_Click()
RES = 0
If Val(GUESTNOOFDISH_TXT.Text) > 0 Then
GUESTDISH_LST.AddItem DISHNAME
GUESTNOOFDISH_LST.AddItem Val(GUESTNOOFDISH_TXT.Text)
GUESTDISHCHARGE_LST.AddItem Val(GUESTDISHPRICEPERPLATE_TXT.Text) * Val(GUESTNOOFDISH_TXT.Text)
GUESTNONVEGDISH_CMB.Text = ""
GUESTVEGDISH_CMB.Text = ""
GUESTDISHPRICEPERPLATE_TXT.Text = ""
GUESTNOOFDISH_TXT.Text = ""
GUESTNOOFDISH_TXT.Enabled = False
RES = 1
DISHNO = DISHNO + 1
T = 0
L = DISHNO
Do Until L = 0
T = T + Val(GUESTDISHCHARGE_LST.List(L - 1))
L = L - 1
Loop
GUESTTOTALDISHCHARGE_LBL.Caption = T
End If
If Val(GUESTNOOFDRINK_TXT.Text) > 0 Then
GUESTDRINK_LST.AddItem DRINKNAME
GUESTNOOFDRINK_LST.AddItem Val(GUESTNOOFDRINK_TXT.Text)
GUESTDRINKCHARGE_LST.AddItem Val(GUESTDRINKPRICEPERGLASS_TXT.Text) * Val(GUESTNOOFDRINK_TXT.Text)
GUESTSOFTDRINK_CMB.Text = ""
GUESTJUICE_CMB.Text = ""
GUESTBEER_CMB.Text = ""
GUESTWINE_CMB.Text = ""
GUESTDRINKPRICEPERGLASS_TXT.Text = ""
GUESTNOOFDRINK_TXT.Text = ""
GUESTNOOFDRINK_TXT.Enabled = False
RES = 1
DRINKNO = DRINKNO + 1
T = 0
L = DRINKNO
Do Until L = 0
T = T + Val(GUESTDRINKCHARGE_LST.List(L - 1))
L = L - 1
Loop
GUESTTOTALDRINKCHARGE_LBL.Caption = T
End If
If RES = 0 Then
MsgBox " PLEASE INPUT NO OF QUANTITY ? "
End If
GUESTTOTALCHARGE_LBL.Caption = Val(GUESTTOTALDISHCHARGE_LBL.Caption) + Val(GUESTTOTALDRINKCHARGE_LBL.Caption)
End Sub

Private Sub GUESTBEER_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & GUESTBEER_CMB.Text & "'"
Set R = C.Execute(S)
GUESTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = GUESTBEER_CMB.Text
GUESTNOOFDRINK_TXT.Enabled = True
GUESTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub GUESTDISH_LST_Click()
GUESTREMOVEORDER_CMD.Enabled = True
End Sub

Private Sub GUESTDRINK_LST_Click()
GUESTREMOVEORDER_CMD.Enabled = True
End Sub

Private Sub GUESTEXIT_CMD_Click()
Me.Hide
Unload Me
End Sub

Private Sub GUESTID_CMB_Click()
S = "SELECT *FROM CLIENT_MASTER WHERE CLIENT_ID='" & GUESTID_CMB.Text & "'"
Set R = C.Execute(S)
GUESTNAME_TXT.Text = R.Fields("NAME")
GUESTID_TXT.Text = R.Fields("CLIENT_ID")
GUESTNONVEGDISH_CMB.Enabled = True
GUESTVEGDISH_CMB.Enabled = True
GUESTSOFTDRINK_CMB.Enabled = True
GUESTJUICE_CMB.Enabled = True
GUESTBEER_CMB.Enabled = True
GUESTWINE_CMB.Enabled = True
GUESTNOOFDISH_TXT.Enabled = True
GUESTNOOFDRINK_TXT.Enabled = True
GUESTDISH_LST.Clear
GUESTNOOFDISH_LST.Clear
GUESTDISHCHARGE_LST.Clear
GUESTDRINK_LST.Clear
GUESTNOOFDRINK_LST.Clear
GUESTDRINKCHARGE_LST.Clear
End Sub

Private Sub GUESTJUICE_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & GUESTJUICE_CMB.Text & "'"
Set R = C.Execute(S)
GUESTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = GUESTJUICE_CMB.Text
GUESTNOOFDRINK_TXT.Enabled = True
GUESTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub GUESTNEWORDER_CMD_Click()
Unload Me
Load FRMRESTORENTGUEST
FRMRESTORENTGUEST.Show
End Sub
Private Sub GUESTNONVEGDISH_CMB_Click()
S = "SELECT CHARGE FROM DISH WHERE DNAME='" & GUESTNONVEGDISH_CMB.Text & "'"
Set R = C.Execute(S)
GUESTDISHPRICEPERPLATE_TXT.Text = R.Fields("CHARGE")
DISHNAME = GUESTNONVEGDISH_CMB.Text
GUESTNOOFDISH_TXT.Enabled = True
GUESTNOOFDISH_TXT.SetFocus
End Sub

Private Sub GUESTNOOFDISH_TXT_Change()
If Val(GUESTNOOFDISH_TXT.Text) > 0 Then
GUESTADDORDER_CMD.Enabled = True
Else
GUESTPAYMENT_CMD.Enabled = True
GUESTADDONDUES_CMD.Enabled = True
End If
End Sub

Private Sub GUESTNOOFDISH_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
End Sub

Private Sub GUESTNOOFDRINK_TXT_Change()
If Val(GUESTNOOFDISH_TXT.Text) > 0 Then
GUESTADDORDER_CMD.Enabled = True
GUESTPAYMENT_CMD.Enabled = True
End If
End Sub

Private Sub GUESTNOOFDRINK_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
End Sub

Private Sub GUESTPAYMENT_CMD_Click()
If Val(GUESTTOTALCHARGE_LBL.Caption) > 0 Then
If MsgBox("ARE YOU SURE TO PAYMENT ?", vbYesNo) = vbYes Then
'S = "UPDATE DAILY SET TOT_INCOME=TOT_INCOME+" & GUESTRESTOTOTCHARGE_TXT.Text & " WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
'MsgBox S
 '       Set R = C.Execute(S)
        S = "UPDATE CLIENT_MASTER SET BILL_STATUS=BILL_STATUS+" & Val(GUESTTOTALCHARGE_LBL.Caption) & ",TOT_CHARGE=TOT_CHARGE+" & Val(GUESTTOTALCHARGE_LBL.Caption) & ",RESTO_CHARGE=RESTO_CHARGE+" & Val(GUESTTOTALCHARGE_LBL.Caption) & " WHERE CLIENT_ID='" & GUESTID_CMB.Text & "'"
        Set R = C.Execute(S)
        S = "COMMIT"
        Set R = C.Execute(S)
    '    If MsgBox("ARE YOU TAKE REPORT THEN CLICK 'YES'", vbYesNo) = vbYes Then
     '       If DataEnvironment1.rsCommand5.State = 1 Then DataEnvironment1.rsCommand5.Close
      '          DataEnvironment1.Command5 GUESTID_CMB.Text
       '         DataReport7.Show
        '        GUESTBARPAYMENT.Enabled = False
        'End If
    'End If
'Else
 '   If MsgBox("ARE YOU TAKE REPORT THEN CLICK 'YES'", vbYesNo) = vbYes Then
  '          If DataEnvironment1.rsCommand5.State = 1 Then DataEnvironment1.rsCommand5.Close
   '             DataEnvironment1.Command5 GUESTID_CMB.Text
    '            DataReport7.Show
     '           GUESTBARPAYMENT.Enabled = False
      '      End If
'End If
Do Until DRINKNO = 0
S = "INSERT INTO CUST VALUES('" & GUESTID_TXT.Text & "','" & GUESTNAME_TXT.Text & "'," & Val(GUESTTOTALCHARGE_LBL.Caption) & ",'" & Format(Date, "DD-MMM-YYYY") & "','" & GUESTDRINK_LST.List(DRINKNO - 1) & "',''," & Val(GUESTDRINKCHARGE_LST.List(DRINKNO - 1)) & "," & 0 & ")"
Set R = C.Execute(S)
DRINKNO = DRINKNO - 1
Loop
Do Until DISHNO = 0
S = "INSERT INTO CUST VALUES('" & GUESTID_TXT.Text & "','" & GUESTNAME_TXT.Text & "'," & Val(GUESTTOTALCHARGE_LBL.Caption) & ",'" & Format(Date, "DD-MMM-YYYY") & "','','" & GUESTDISH_LST.List(DISHNO - 1) & "'," & 0 & "," & GUESTDISHCHARGE_LST.List(DISHNO - 1) & ")"
Set R = C.Execute(S)
DISHNO = DISHNO - 1
Loop
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "PAYMENT COMPLETE !"
Unload Me
Load FRMRESTORENTGUEST
FRMRESTORENTGUEST.Show
End If
Else
MsgBox "PLEASE ADD ORDER FIRST"
End If
End Sub

Private Sub GUESTREMOVEORDER_CMD_Click()
If GUESTDISH_LST.ListIndex >= 0 Then
GUESTTOTALDISHCHARGE_LBL.Caption = Val(GUESTTOTALDISHCHARGE_LBL.Caption) - Val(GUESTDISHCHARGE_LST.List(GUESTDISH_LST.ListIndex))
GUESTTOTALCHARGE_LBL.Caption = Val(GUESTTOTALCHARGE_LBL.Caption) - Val(GUESTDISHCHARGE_LST.List(GUESTDISH_LST.ListIndex))
GUESTNOOFDISH_LST.RemoveItem (GUESTDISH_LST.ListIndex)
GUESTDISHCHARGE_LST.RemoveItem (GUESTDISH_LST.ListIndex)
GUESTDISH_LST.RemoveItem (GUESTDISH_LST.ListIndex)
DISHNO = DISHNO - 1
End If
If GUESTDRINK_LST.ListIndex >= 0 Then
GUESTTOTALDRINKCHARGE_LBL.Caption = Val(GUESTTOTALDRINKCHARGE_LBL.Caption) - Val(GUESTDRINKCHARGE_LST.List(GUESTDRINK_LST.ListIndex))
GUESTTOTALCHARGE_LBL.Caption = Val(GUESTTOTALCHARGE_LBL.Caption) - Val(GUESTDRINKCHARGE_LST.List(GUESTDRINK_LST.ListIndex))
GUESTNOOFDRINK_LST.RemoveItem (GUESTDRINK_LST.ListIndex)
GUESTDRINKCHARGE_LST.RemoveItem (GUESTDRINK_LST.ListIndex)
GUESTDRINK_LST.RemoveItem (GUESTDRINK_LST.ListIndex)
DRINKNO = DRINKNO - 1
End If
End Sub

Private Sub GUESTSOFTDRINK_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & GUESTSOFTDRINK_CMB.Text & "'"
Set R = C.Execute(S)
GUESTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = GUESTSOFTDRINK_CMB.Text
GUESTNOOFDRINK_TXT.Enabled = True
GUESTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub GUESTSTPAYMENT_CMD_Click()

End Sub

Private Sub GUESTVEGDISH_CMB_Click()
S = "SELECT CHARGE FROM DISH WHERE DNAME='" & GUESTVEGDISH_CMB.Text & "'"
Set R = C.Execute(S)
GUESTDISHPRICEPERPLATE_TXT.Text = R.Fields("CHARGE")
DISHNAME = GUESTVEGDISH_CMB.Text
GUESTNOOFDISH_TXT.Enabled = True
GUESTNOOFDISH_TXT.SetFocus
End Sub

Private Sub GUESTWINE_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & GUESTWINE_CMB.Text & "'"
Set R = C.Execute(S)
GUESTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = GUESTWINE_CMB.Text
GUESTNOOFDRINK_TXT.Enabled = True
GUESTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub Form_Load()
CON
Me.Top = 200
Me.Left = 4000
DISHNO = 0
DRINKNO = 0
GUESTID_CMB.Clear
GUESTNONVEGDISH_CMB.Clear
GUESTVEGDISH_CMB.Clear
GUESTSOFTDRINK_CMB.Clear
GUESTBEER_CMB.Clear
GUESTJUICE_CMB.Clear
GUESTWINE_CMB.Clear
S = "SELECT *FROM CLIENT_MASTER WHERE OUT='IN'"
Set R = C.Execute(S)
Do Until R.EOF = True
    GUESTID_CMB.AddItem R.Fields("CLIENT_ID")
    R.MoveNext
Loop
S = "SELECT *FROM DISH WHERE TYPE='NONVEGDISH'"
Set R = C.Execute(S)
Do Until R.EOF = True
GUESTNONVEGDISH_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DISH WHERE TYPE='VEGDISH'"
Set R = C.Execute(S)
Do Until R.EOF = True
GUESTVEGDISH_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='SOFTDRINK'"
Set R = C.Execute(S)
Do Until R.EOF = True
GUESTSOFTDRINK_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='JUICE'"
Set R = C.Execute(S)
Do Until R.EOF = True
GUESTJUICE_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='BEER'"
Set R = C.Execute(S)
Do Until R.EOF = True
GUESTBEER_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='WINE'"
Set R = C.Execute(S)
Do Until R.EOF = True
GUESTWINE_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
End Sub

