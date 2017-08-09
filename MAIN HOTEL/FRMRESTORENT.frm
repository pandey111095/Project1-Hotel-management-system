VERSION 5.00
Begin VB.Form FRMRESTORENT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMER RESTAURENT"
   ClientHeight    =   8565
   ClientLeft      =   3870
   ClientTop       =   1095
   ClientWidth     =   14895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMRESTORENT.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   14895
   Begin VB.CommandButton NEWCUSTADD_CMD 
      Caption         =   "NEW CUSTOMER"
      Height          =   495
      Left            =   3000
      TabIndex        =   50
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton CUSTREMOVEORDER_CMD 
      Caption         =   "REMOVE ORDER"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7440
      TabIndex        =   43
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton CUSTEXIT_CMD 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   10560
      TabIndex        =   42
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton CUSTPAYMENT_CMD 
      Caption         =   "PAYMENT"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9240
      TabIndex        =   41
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton CUSTADDORDER_CMD 
      Caption         =   "ADD ORDER"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      TabIndex        =   40
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton CUSTNEWORDER_CMD 
      Caption         =   "NEW OEDER"
      Height          =   495
      Left            =   4680
      TabIndex        =   39
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox CUSTNOOFDRINK_TXT 
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
      TabIndex        =   30
      Text            =   "0"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox CUSTNOOFDISH_TXT 
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
      TabIndex        =   28
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox CUSTDRINKPRICEPERGLASS_TXT 
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
      TabIndex        =   24
      Text            =   "0"
      Top             =   3120
      Width           =   855
   End
   Begin VB.ListBox CUSTDRINKCHARGE_LST 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   11760
      TabIndex        =   23
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ListBox CUSTNOOFDRINK_LST 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   10320
      TabIndex        =   22
      Top             =   3840
      Width           =   855
   End
   Begin VB.ListBox CUSTDRINK_LST 
      Height          =   2400
      Left            =   8040
      TabIndex        =   21
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ComboBox CUSTWINE_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   13200
      TabIndex        =   20
      Text            =   "WINE"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox CUSTBEER_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   11400
      TabIndex        =   19
      Text            =   "BEER"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ComboBox CUSTJUICE_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9840
      TabIndex        =   18
      Text            =   "JUICE"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox CUSTSOFTDRINK_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8040
      TabIndex        =   17
      Text            =   "SOFT DRINK"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ListBox CUSTDISHCHARGE_LST 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   4560
      TabIndex        =   16
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ListBox CUSTNOOFDISH_LST 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   2760
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ListBox CUSTDISH_LST 
      Height          =   2595
      Left            =   480
      TabIndex        =   14
      Top             =   3720
      Width           =   1815
   End
   Begin VB.ComboBox CUSTVEGDISH_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Text            =   "VEG DISH"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox CUSTNONVEGDISH_CMB 
      Enabled         =   0   'False
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Text            =   "NONVEG DISH"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox CUSTDISHPRICEPERPLATE_TXT 
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
      TabIndex        =   11
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox CUSTID_CMB 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Text            =   "CUSTOMER ID"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox CUSTID_TXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7200
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox CUSTNAME_TXT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8520
      MaxLength       =   20
      TabIndex        =   0
      Top             =   960
      Width           =   1575
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
      TabIndex        =   49
      Top             =   2160
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
      TabIndex        =   48
      Top             =   2160
      Width           =   615
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
      TabIndex        =   47
      Top             =   2160
      Width           =   615
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
      TabIndex        =   46
      Top             =   2160
      Width           =   1455
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
      TabIndex        =   45
      Top             =   2040
      Width           =   975
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
      TabIndex        =   44
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label CUSTTOTALCHARGE_LBL 
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
      TabIndex        =   38
      Top             =   7320
      Width           =   1335
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
      TabIndex        =   37
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Line Line3 
      BorderWidth     =   10
      X1              =   0
      X2              =   15000
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label CUSTTOTALDRINKCHARGE_LBL 
      Height          =   255
      Left            =   11760
      TabIndex        =   36
      Top             =   6600
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
      TabIndex        =   35
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label CUSTTOTALDISHCHARGE_LBL 
      Height          =   255
      Left            =   3600
      TabIndex        =   34
      Top             =   6600
      Width           =   975
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
      TabIndex        =   33
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   10
      X1              =   7200
      X2              =   7200
      Y1              =   1920
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   0
      X2              =   15000
      Y1              =   1800
      Y2              =   1800
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
      Left            =   8520
      TabIndex        =   32
      Top             =   600
      Width           =   1815
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
      Left            =   6960
      TabIndex        =   31
      Top             =   600
      Width           =   1335
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
      TabIndex        =   29
      Top             =   3120
      Width           =   1590
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
      TabIndex        =   27
      Top             =   3000
      Width           =   1695
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
      TabIndex        =   26
      Top             =   3480
      Width           =   1185
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
      TabIndex        =   25
      Top             =   3360
      Width           =   1575
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
      TabIndex        =   10
      Top             =   3480
      Width           =   780
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
      TabIndex        =   9
      Top             =   3480
      Width           =   1590
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
      TabIndex        =   8
      Top             =   3120
      Width           =   1665
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
      TabIndex        =   7
      Top             =   3360
      Width           =   855
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
      TabIndex        =   6
      Top             =   3360
      Width           =   1935
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
      TabIndex        =   5
      Top             =   3000
      Width           =   2415
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
      TabIndex        =   4
      Top             =   1440
      Width           =   975
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
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "FRMRESTORENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NO_CUST As Integer
Dim DISHNO As Integer
Dim DRINKNO As Integer
Dim DISHNAME As String
Dim DRINKNAME As String
Dim L As Integer
Dim T As Integer







Private Sub CUSTADDORDER_CMD_Click()
RES = 0
If Val(CUSTNOOFDISH_TXT.Text) > 0 Then
CUSTDISH_LST.AddItem DISHNAME
CUSTNOOFDISH_LST.AddItem Val(CUSTNOOFDISH_TXT.Text)
CUSTDISHCHARGE_LST.AddItem Val(CUSTDISHPRICEPERPLATE_TXT.Text) * Val(CUSTNOOFDISH_TXT.Text)
CUSTNONVEGDISH_CMB.Text = ""
CUSTVEGDISH_CMB.Text = ""
CUSTDISHPRICEPERPLATE_TXT.Text = ""
CUSTNOOFDISH_TXT.Text = ""
CUSTNOOFDISH_TXT.Enabled = False
RES = 1
DISHNO = DISHNO + 1
T = 0
L = DISHNO
Do Until L = 0
T = T + Val(CUSTDISHCHARGE_LST.List(L - 1))
L = L - 1
Loop
CUSTTOTALDISHCHARGE_LBL.Caption = T
End If
If Val(CUSTNOOFDRINK_TXT.Text) > 0 Then
CUSTDRINK_LST.AddItem DRINKNAME
CUSTNOOFDRINK_LST.AddItem Val(CUSTNOOFDRINK_TXT.Text)
CUSTDRINKCHARGE_LST.AddItem Val(CUSTDRINKPRICEPERGLASS_TXT.Text) * Val(CUSTNOOFDRINK_TXT.Text)
CUSTSOFTDRINK_CMB.Text = ""
CUSTJUICE_CMB.Text = ""
CUSTBEER_CMB.Text = ""
CUSTWINE_CMB.Text = ""
CUSTDRINKPRICEPERGLASS_TXT.Text = ""
CUSTNOOFDRINK_TXT.Text = ""
CUSTNOOFDRINK_TXT.Enabled = False
RES = 1
DRINKNO = DRINKNO + 1
T = 0
L = DRINKNO
Do Until L = 0
T = T + Val(CUSTDRINKCHARGE_LST.List(L - 1))
L = L - 1
Loop
CUSTTOTALDRINKCHARGE_LBL.Caption = T
End If
If RES = 0 Then
MsgBox " PLEASE INPUT NO OF QUANTITY ? "
End If
CUSTTOTALCHARGE_LBL.Caption = Val(CUSTTOTALDISHCHARGE_LBL.Caption) + Val(CUSTTOTALDRINKCHARGE_LBL.Caption)
End Sub

Private Sub CUSTBEER_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & CUSTBEER_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = CUSTBEER_CMB.Text
CUSTNOOFDRINK_TXT.Enabled = True
CUSTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub CUSTDISH_LST_Click()
CUSTREMOVEORDER_CMD.Enabled = True
End Sub

Private Sub CUSTDRINK_LST_Click()
CUSTREMOVEORDER_CMD.Enabled = True
End Sub

Private Sub CUSTEXIT_CMD_Click()
Me.Hide
Unload Me
End Sub

Private Sub CUSTID_CMB_Click()
S = "SELECT *FROM CUST WHERE CUST_ID='" & CUSTID_CMB.Text & "' AND TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
CUSTNAME_TXT.Text = R.Fields("CUST_NAME")
CUSTID_TXT.Text = R.Fields("CUST_ID")
CUSTNONVEGDISH_CMB.Enabled = True
CUSTVEGDISH_CMB.Enabled = True
CUSTSOFTDRINK_CMB.Enabled = True
CUSTJUICE_CMB.Enabled = True
CUSTBEER_CMB.Enabled = True
CUSTWINE_CMB.Enabled = True
CUSTNOOFDISH_TXT.Enabled = True
CUSTNOOFDRINK_TXT.Enabled = True
CUSTDISH_LST.Clear
CUSTNOOFDISH_LST.Clear
CUSTDISHCHARGE_LST.Clear
CUSTDRINK_LST.Clear
CUSTNOOFDRINK_LST.Clear
CUSTDRINKCHARGE_LST.Clear

End Sub

Private Sub CUSTJUICE_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & CUSTJUICE_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = CUSTJUICE_CMB.Text
CUSTNOOFDRINK_TXT.Enabled = True
CUSTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub CUSTNEWORDER_CMD_Click()
Unload Me
Load FRMRESTORENT
FRMRESTORENT.Show
End Sub

Private Sub CUSTNONVEGDISH_CMB_Click()
S = "SELECT CHARGE FROM DISH WHERE DNAME='" & CUSTNONVEGDISH_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDISHPRICEPERPLATE_TXT.Text = R.Fields("CHARGE")
DISHNAME = CUSTNONVEGDISH_CMB.Text
CUSTNOOFDISH_TXT.Enabled = True
CUSTNOOFDISH_TXT.SetFocus
End Sub

Private Sub CUSTNOOFDISH_TXT_Change()
If Val(CUSTNOOFDISH_TXT.Text) > 0 Then
CUSTADDORDER_CMD.Enabled = True
Else
CUSTPAYMENT_CMD.Enabled = True
End If
End Sub

Private Sub CUSTNOOFDISH_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
End Sub

Private Sub CUSTNOOFDRINK_TXT_Change()
If Val(CUSTNOOFDISH_TXT.Text) > 0 Then
CUSTADDORDER_CMD.Enabled = True
CUSTPAYMENT_CMD.Enabled = True
End If
End Sub

Private Sub CUSTNOOFDRINK_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
End Sub

Private Sub CUSTPAYMENT_CMD_Click()
'If Val(CUSTNOOFDISH_TXT.Text) > 0 Then
'CUSTDISH_LST.AddItem DISHNAME
'CUSTNOOFDISH_LST.AddItem Val(CUSTNOOFDISH_TXT.Text)
'CUSTDISHCHARGE_LST.AddItem Val(CUSTDISHPRICEPERPLATE_TXT.Text) * Val(CUSTNOOFDISH_TXT.Text)
'CUSTNONVEGDISH_CMB.Refresh
'CUSTVEGDISH_CMB.Refresh
'CUSTDISHPRICEPERPLATE_TXT.Text = ""
'CUSTNOOFDISH_TXT.Text = ""
'CUSTNOOFDISH_TXT.Enabled = False
'DISHNO = DISHNO + 1
'End If
'If Val(CUSTNOOFDRINK_TXT.Text) > 0 Then
'CUSTDRINK_LST.AddItem DRINKNAME
'CUSTNOOFDRINK_LST.AddItem Val(CUSTNOOFDRINK_TXT.Text)
'CUSTDRINKCHARGE_LST.AddItem Val(CUSTDRINKPRICEPERGLASS_TXT.Text) * Val(CUSTNOOFDRINK_TXT.Text)
'CUSTSOFTDRINK_CMB.Refresh
'CUSTJUICE_CMB.Refresh
'CUSTBEER_CMB.Refresh
'CUSTWINE_CMB.Refresh
'CUSTDRINKPRICEPERGLASS_TXT.Text = ""
'CUSTNOOFDRINK_TXT.Text = ""
'CUSTNOOFDRINK_TXT.Enabled = False
'DRINKNO = DRINKNO + 1
'End If
If Val(CUSTTOTALCHARGE_LBL.Caption) > 0 Then
If MsgBox("ARE YOU SURE TO PAYMENT ?", vbYesNo) = vbYes Then
Do Until DRINKNO = 0
S = "INSERT INTO CUST VALUES('" & CUSTID_TXT.Text & "','" & CUSTNAME_TXT.Text & "'," & Val(CUSTTOTALCHARGE_LBL.Caption) & ",'" & Format(Date, "DD-MMM-YYYY") & "','" & CUSTDRINK_LST.List(DRINKNO - 1) & "',''," & Val(CUSTDRINKCHARGE_LST.List(DRINKNO - 1)) & "," & 0 & ")"
Set R = C.Execute(S)
DRINKNO = DRINKNO - 1
Loop
Do Until DISHNO = 0
S = "INSERT INTO CUST VALUES('" & CUSTID_TXT.Text & "','" & CUSTNAME_TXT.Text & "'," & Val(CUSTTOTALCHARGE_LBL.Caption) & ",'" & Format(Date, "DD-MMM-YYYY") & "','','" & CUSTDISH_LST.List(DISHNO - 1) & "'," & 0 & "," & CUSTDISHCHARGE_LST.List(DISHNO - 1) & ")"
Set R = C.Execute(S)
DISHNO = DISHNO - 1
Loop
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "PAYMENT COMPLETE !"
Unload Me
Load FRMRESTORENT
FRMRESTORENT.Show
End If
Else
MsgBox "PLEASE ADD ORDER FIRST"
End If
End Sub

Private Sub CUSTREMOVEORDER_CMD_Click()
If CUSTDISH_LST.ListIndex >= 0 Then
CUSTTOTALDISHCHARGE_LBL.Caption = Val(CUSTTOTALDISHCHARGE_LBL.Caption) - Val(CUSTDISHCHARGE_LST.List(CUSTDISH_LST.ListIndex))
CUSTTOTALCHARGE_LBL.Caption = Val(CUSTTOTALCHARGE_LBL.Caption) - Val(CUSTDISHCHARGE_LST.List(CUSTDISH_LST.ListIndex))
CUSTNOOFDISH_LST.RemoveItem (CUSTDISH_LST.ListIndex)
CUSTDISHCHARGE_LST.RemoveItem (CUSTDISH_LST.ListIndex)
CUSTDISH_LST.RemoveItem (CUSTDISH_LST.ListIndex)
DISHNO = DISHNO - 1
End If
If CUSTDRINK_LST.ListIndex >= 0 Then
CUSTTOTALDRINKCHARGE_LBL.Caption = Val(CUSTTOTALDRINKCHARGE_LBL.Caption) - Val(CUSTDRINKCHARGE_LST.List(CUSTDRINK_LST.ListIndex))
CUSTTOTALCHARGE_LBL.Caption = Val(CUSTTOTALCHARGE_LBL.Caption) - Val(CUSTDRINKCHARGE_LST.List(CUSTDRINK_LST.ListIndex))
CUSTNOOFDRINK_LST.RemoveItem (CUSTDRINK_LST.ListIndex)
CUSTDRINKCHARGE_LST.RemoveItem (CUSTDRINK_LST.ListIndex)
CUSTDRINK_LST.RemoveItem (CUSTDRINK_LST.ListIndex)
DRINKNO = DRINKNO - 1
End If
End Sub

Private Sub CUSTSOFTDRINK_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & CUSTSOFTDRINK_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = CUSTSOFTDRINK_CMB.Text
CUSTNOOFDRINK_TXT.Enabled = True
CUSTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub CUSTVEGDISH_CMB_Click()
S = "SELECT CHARGE FROM DISH WHERE DNAME='" & CUSTVEGDISH_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDISHPRICEPERPLATE_TXT.Text = R.Fields("CHARGE")
DISHNAME = CUSTVEGDISH_CMB.Text
CUSTNOOFDISH_TXT.Enabled = True
CUSTNOOFDISH_TXT.SetFocus
End Sub

Private Sub CUSTWINE_CMB_Click()
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & CUSTWINE_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDRINKPRICEPERGLASS_TXT.Text = R.Fields("CHARGE")
DRINKNAME = CUSTWINE_CMB.Text
CUSTNOOFDRINK_TXT.Enabled = True
CUSTNOOFDRINK_TXT.SetFocus
End Sub

Private Sub Form_Load()
CON
Me.Top = 200
Me.Left = 4000
DISHNO = 0
DRINKNO = 0
CUSTID_CMB.Clear
CUSTNONVEGDISH_CMB.Clear
CUSTVEGDISH_CMB.Clear
CUSTSOFTDRINK_CMB.Clear
CUSTBEER_CMB.Clear
CUSTJUICE_CMB.Clear
CUSTWINE_CMB.Clear
S = "SELECT COUNT(CUST_NAME) FROM CUST WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
NO_CUST = R.Fields(0)
If NO_CUST = 0 Then
    CUSTID_TXT.Text = "C1"
    CUSTID_CMB.Enabled = False
    CUSTNAME_TXT.Enabled = True
'    CUSTNAME_TXT.SetFocus
     CUSTNONVEGDISH_CMB.Enabled = True
    CUSTVEGDISH_CMB.Enabled = True
    CUSTSOFTDRINK_CMB.Enabled = True
    CUSTJUICE_CMB.Enabled = True
    CUSTBEER_CMB.Enabled = True
    CUSTWINE_CMB.Enabled = True
Else
    CUSTID_CMB.Enabled = True
    S = "SELECT DISTINCT(CUST_ID) FROM CUST WHERE CUST_ID LIKE'C%' AND TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
    Set R = C.Execute(S)
    Do Until R.EOF = True
    CUSTID_CMB.AddItem R.Fields("CUST_ID")
    R.MoveNext
    Loop
    CUSTID_CMB.Visible = True
    CUSTID_CMB.Enabled = True
'    CUSTID_CMB.SetFocus
End If
S = "SELECT *FROM DISH WHERE TYPE='NONVEGDISH'"
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTNONVEGDISH_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DISH WHERE TYPE='VEGDISH'"
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTVEGDISH_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='SOFTDRINK'"
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTSOFTDRINK_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='JUICE'"
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTJUICE_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='BEER'"
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTBEER_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='WINE'"
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTWINE_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
End Sub

Private Sub List1_Click()

End Sub

Private Sub NEWCUSTADD_CMD_Click()
Unload Me
Load FRMRESTORENT
FRMRESTORENT.Show
S = "SELECT COUNT(DISTINCT(CUST_ID))FROM CUST WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
NO_CUST = R.Fields(0)
If NO_CUST = 0 Then
    CUSTID_TXT.Text = "C1"
Else
    CUSTID_TXT.Text = "C" & (NO_CUST + 1)
    CUSTNAME_TXT.Enabled = True
    CUSTNAME_TXT.SetFocus
    CUSTNONVEGDISH_CMB.Enabled = True
    CUSTVEGDISH_CMB.Enabled = True
    CUSTSOFTDRINK_CMB.Enabled = True
    CUSTJUICE_CMB.Enabled = True
    CUSTBEER_CMB.Enabled = True
    CUSTWINE_CMB.Enabled = True
    CUSTNOOFDISH_TXT.Enabled = True
    CUSTNOOFDRINK_TXT.Enabled = True
    CUSTDISH_LST.Clear
    CUSTNOOFDISH_LST.Clear
    CUSTDISHCHARGE_LST.Clear
    CUSTDRINK_LST.Clear
    CUSTNOOFDRINK_LST.Clear
    CUSTDRINKCHARGE_LST.Clear
End If
End Sub
