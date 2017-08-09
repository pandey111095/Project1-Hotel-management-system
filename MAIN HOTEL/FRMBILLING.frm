VERSION 5.00
Begin VB.Form FRMBARORDER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESTAURENT AND BAR "
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   255
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame CUST_FRAME 
      Caption         =   "Customer"
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
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   1320
      TabIndex        =   17
      Top             =   2760
      Width           =   12135
      Begin VB.CommandButton ADDORDER_CMD 
         Caption         =   "ADD ORDER"
         Height          =   375
         Left            =   3000
         TabIndex        =   48
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Payment"
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
         Height          =   375
         Left            =   5280
         TabIndex        =   44
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox CUSTRESTOTOTCHARGE_TXT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   40
         Text            =   "0"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox CUSTDISHQUANTITY_TXT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   8640
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton CUSTTAKEORDER_CMD 
         Caption         =   "Take Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton CUSTNEWORDER_CMD 
         Caption         =   "New Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label CUSTTOTCHARGE_LBL 
         Caption         =   "TOTAL CHARGE"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label CUSTDISHTYPE_LBL 
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
         Left            =   8640
         TabIndex        =   32
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "DISH_TYPE"
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
         Left            =   7080
         TabIndex        =   31
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label CUSTDISHRATE_LBL 
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
         Left            =   8640
         TabIndex        =   30
         Top             =   1200
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
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
         Left            =   7080
         TabIndex        =   29
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "Rate"
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
         Left            =   7080
         TabIndex        =   28
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Rate"
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
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
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
         TabIndex        =   23
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label CUSTDRINKRATE_LBL 
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
         Left            =   1680
         TabIndex        =   22
         Top             =   1200
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Drink Type"
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
         TabIndex        =   21
         Top             =   840
         Width           =   990
      End
      Begin VB.Label CUSTDRINKTYPE_LBL 
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
         Left            =   1680
         TabIndex        =   20
         Top             =   840
         Width           =   60
      End
   End
   Begin VB.Frame GUEST_FRAME 
      Caption         =   "Hotel Guest"
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
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   1320
      TabIndex        =   3
      Top             =   5160
      Width           =   12135
      Begin VB.TextBox GUESTRESTOTOTCHARGE_TXT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   41
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox GUESTDISHQUANTITY_TXT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   8640
         MaxLength       =   2
         TabIndex        =   33
         Text            =   "0"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox GUESTDRINKQUANTITY_TXT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox GUESTID_CMB 
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
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton GUESTTAKEBARORDER_CMD 
         Caption         =   "Take Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton GUESTNEWBARORDER_CMD 
         Caption         =   "New Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton GUESTBARPAYMENT 
         Caption         =   "Payment"
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
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label NAME_LBL 
         Height          =   255
         Left            =   1680
         TabIndex        =   47
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   15
         Left            =   1680
         TabIndex        =   46
         Top             =   720
         Width           =   495
      End
      Begin VB.Label GUESTTOTCHARGE_LBL 
         Caption         =   "TOTAL CHARGE"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4800
         TabIndex        =   42
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label GUESTDISHTYPE_LBL 
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
         Left            =   8640
         TabIndex        =   39
         Top             =   720
         Width           =   60
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "DISH TYPE"
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
         Left            =   7080
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label GUESTDISHRATE_LBL 
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
         Left            =   8640
         TabIndex        =   37
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
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
         Left            =   7080
         TabIndex        =   36
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label21 
         Caption         =   "Rate"
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
         Left            =   7080
         TabIndex        =   35
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "DISH"
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
         Left            =   7080
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Drinks"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Drink Type"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label GUESTDRINKTYPE_LBL 
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
         Left            =   1680
         TabIndex        =   14
         Top             =   1560
         Width           =   60
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label GUESTDRINKRATE_LBL 
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
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   60
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Guest Name"
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
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1155
      End
   End
   Begin VB.CommandButton EXITTOMDIFORM 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   7440
      Width           =   975
   End
   Begin VB.OptionButton CUSTBAR_OPT 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton GUESTBAR_OPT 
      Caption         =   "Guest"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Select Order Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1320
      TabIndex        =   26
      Top             =   1440
      Width           =   2250
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Process Bar Order"
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
      Left            =   2520
      TabIndex        =   25
      Top             =   960
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   5640
      Picture         =   "FRMBILLING.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "FRMBARORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ORDER(0 To 10) As Integer
Private Sub ADDORDER_CMD_Click()
CUSTID_CMB.Clear
S = "SELECT CUST_ID FROM CUST"
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTID_CMB.AddItem R.Fields("CUST_ID")
R.MoveNext
Loop
CUSTID_CMB.Visible = True
If Val(CUSTDRINKRATE_LBL.Caption) < 0 Then
    If CUSTDRINKQUANTITY_TXT.Text = "" Then
        MsgBox "PLEASE INPUT QUANTITY OF DRINK"
        CUSTDRINKQUANTITY_TXT.SetFocus
    End If
ElseIf Val(CUSTDISHRATE_LBL.Caption) < 0 Then
    MsgBox "PLEASE INPUT QUANTITY OF DISH"
    CUSTDISHQUANTITY_TXT.SetFocus
Else
    If MsgBox("ARE YOU SURE TO ADD ORDER ?", vbYesNo) = vbYes Then
        CUSTRESTOTOTCHARGE_TXT.Text = (Val(CUSTDRINKQUANTITY_TXT.Text) * Val(CUSTDRINKRATE_LBL.Caption) + Val(CUSTDISHQUANTITY_TXT.Text) * Val(CUSTDISHRATE_LBL.Caption))
        S = "INSERT INTO CUST VALUES('" & ID_TXT.Text & "','" & NAME_TXT.Text & "','" & CUSTRESTOTOTCHARGE_TXT.Text & "','" & Format(Date, "DD-MMM-YYYY") & "','" & CUSTDRINK_CMB.Text & "','" & CUSTDISH_CMB.Text & "'," & Val(CUSTDRINKQUANTITY_TXT.Text) * Val(CUSTDRINKRATE_LBL.Caption) & "," & Val(CUSTDISHQUANTITY_TXT.Text) * Val(CUSTDISHRATE_LBL.Caption) & ")"
        Set R = C.Execute(S)
        'CUSTTAKEORDER_CMD.Enabled = False
        CUSTID_CMB.SetFocus
    Else
        CUSTRESTOTOTCHARGE_TXT.Text = "0"
        CUSTDISHTYPE_LBL.Caption = ""
        CUSTDISHRATE_LBL.Caption = ""
        CUSTDRINKTYPE_LBL.Caption = ""
        CUSTDRINKRATE_LBL.Caption = ""
        GUESTDRINKTYPE_LBL.Caption = ""
        GUESTDRINKRATE_LBL.Caption = ""
        GUESTDISHTYPE_LBL.Caption = ""
        GUESTDISHRATE_LBL.Caption = ""
        GUESTDRINKQUANTITY_TXT.Text = ""
        GUESTDISHQUANTITY_TXT.Text = ""
        GUESTRESTOTOTCHARGE_TXT.Text = ""
        CUSTDRINKQUANTITY_TXT.Text = ""
        CUSTDISHQUANTITY_TXT.Text = ""
        CUSTRESTOTOTCHARGE_TXT.Text = ""
        CUST_FRAME.Refresh
        GUEST_FRAME.Refresh
        S = "SELECT COUNT(DISTINCT CUST_ID)FROM CUST WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
        Set R = C.Execute(S)
        NO_CUST = R.Fields(0)
        If NO_CUST = 0 Then
            ID_TXT.Text = "C1"
        Else
            ID_TXT.Text = "C" & (NO_CUST + 1)
        End If
        NAME_TXT.Enabled = True
        NAME_TXT.SetFocus
        NAME_TXT.Text = ""
        GUEST_FRAME.Enabled = False
        CUSTDISHQUANTITY_TXT.Text = "0"
        CUSTDRINKQUANTITY_TXT.Text = "0"
        CUST_FRAME.Enabled = True
        CUSTTAKEORDER_CMD.Enabled = False
    End If
End If
End Sub

Private Sub Command1_Click()
If MsgBox("ARE SURE FOR PAYMENT?", vbYesNo) = vbYes Then
    S = "UPDATE DAILY SET TOT_INCOME=TOT_INCOME+" & CUSTRESTOTOTCHARGE_TXT.Text & " WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
    Set R = C.Execute(S)
    If MsgBox("ARE YOU TAKE REPORT THEN CLICK 'YES'", vbYesNo) = vbYes Then
        If DataEnvironment1.rsCommand4.State = 1 Then DataEnvironment1.rsCommand4.Close
        DataEnvironment1.Command4 ID_TXT.Text, Format(Date, "DD-MMM-YYYY")
        DataReport6.Sections("section1").Controls("label5").Caption = CUSTDRINKTYPE_LBL.Caption
        DataReport6.Show
    End If
End If
CUSTRESTOTOTCHARGE_TXT.Text = "0"
CUSTDISHTYPE_LBL.Caption = ""
CUSTDISHRATE_LBL.Caption = ""
CUSTDRINKTYPE_LBL.Caption = ""
CUSTDRINKRATE_LBL.Caption = ""
GUESTDRINKTYPE_LBL.Caption = ""
GUESTDRINKRATE_LBL.Caption = ""
GUESTDISHTYPE_LBL.Caption = ""
GUESTDISHRATE_LBL.Caption = ""
GUESTDRINKQUANTITY_TXT.Text = ""
GUESTDISHQUANTITY_TXT.Text = ""
GUESTRESTOTOTCHARGE_TXT.Text = ""
CUSTDRINKQUANTITY_TXT.Text = ""
CUSTDISHQUANTITY_TXT.Text = ""
CUSTRESTOTOTCHARGE_TXT.Text = ""
Call Form_Load
CUST_FRAME.Refresh
GUEST_FRAME.Refresh
S = "SELECT COUNT(CUST_NAME)FROM CUST WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
NO_CUST = R.Fields(0)
If NO_CUST = 0 Then
    ID_TXT.Text = "C1"
Else
    ID_TXT.Text = "C" & (NO_CUST + 1)
End If
NAME_TXT.Enabled = True
NAME_TXT.SetFocus
NAME_TXT.Text = ""
GUEST_FRAME.Enabled = False
CUSTDISHQUANTITY_TXT.Text = "0"
CUSTDRINKQUANTITY_TXT.Text = "0"
CUST_FRAME.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me
Me.Hide
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub CUSTBAR_OPT_Click()
GUESTDRINKTYPE_LBL.Caption = ""
GUESTDRINKRATE_LBL.Caption = ""
GUESTDISHTYPE_LBL.Caption = ""
GUESTDISHRATE_LBL.Caption = ""
CUSTDISHTYPE_LBL.Caption = ""
CUSTDISHRATE_LBL.Caption = ""
CUSTDRINKTYPE_LBL.Caption = ""
CUSTDRINKRATE_LBL.Caption = ""
GUESTDRINKQUANTITY_TXT.Text = ""
GUESTDISHQUANTITY_TXT.Text = ""
GUESTRESTOTOTCHARGE_TXT.Text = ""
CUSTDRINKQUANTITY_TXT.Text = ""
CUSTDISHQUANTITY_TXT.Text = ""
CUSTRESTOTOTCHARGE_TXT.Text = ""
Call Form_Load
CUST_FRAME.Refresh
GUEST_FRAME.Refresh
S = "SELECT COUNT(CUST_NAME)FROM CUST WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
NO_CUST = R.Fields(0)
If NO_CUST = 0 Then
    ID_TXT.Text = "C1"
Else
    ID_TXT.Text = "C" & (NO_CUST + 1)
End If
NAME_TXT.Enabled = True
NAME_TXT.SetFocus
NAME_TXT.Text = ""
GUEST_FRAME.Enabled = False
CUSTDISHQUANTITY_TXT.Text = "0"
CUSTDRINKQUANTITY_TXT.Text = "0"
CUST_FRAME.Enabled = True
End Sub



Private Sub CUSTDISH_CMB_Click()
CON
S = "SELECT *FROM DISH WHERE DNAME='" & CUSTDISH_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDISHTYPE_LBL.Caption = R.Fields("TYPE")
CUSTDISHRATE_LBL.Caption = R.Fields("CHARGE")
End Sub

Private Sub CUSTDISHQUANTITY_TXT_Change()
If Val(CUSTDISHQUANTITY_TXT.Text) > 0 Then
    CUSTTAKEORDER_CMD.Enabled = True
Else
    CUSTTAKEORDER_CMD.Enabled = False
End If
End Sub

Private Sub CUSTDRINK_CMB_Click()
CON
S = "SELECT *FROM DRINK WHERE DNAME='" & CUSTDRINK_CMB.Text & "'"
Set R = C.Execute(S)
CUSTDRINKTYPE_LBL.Caption = R.Fields("TYPE")
CUSTDRINKRATE_LBL.Caption = R.Fields("CHARGE")
End Sub

Private Sub CUSTDRINKQUANTITY_TXT_Change()
If Val(CUSTDRINKQUANTITY_TXT.Text) > 0 Then
    CUSTTAKEORDER_CMD.Enabled = True
Else
    CUSTTAKEORDER_CMD.Enabled = False
End If
End Sub

Private Sub CUSTNEWORDER_CMD_Click()
ORDER = 1
CUSTID_CMB.Visible = False
CUSTRESTOTOTCHARGE_TXT.Text = "0"
CUSTDISHTYPE_LBL.Caption = ""
CUSTDISHRATE_LBL.Caption = ""
CUSTDRINKTYPE_LBL.Caption = ""
CUSTDRINKRATE_LBL.Caption = ""
GUESTDRINKTYPE_LBL.Caption = ""
GUESTDRINKRATE_LBL.Caption = ""
GUESTDISHTYPE_LBL.Caption = ""
GUESTDISHRATE_LBL.Caption = ""
GUESTDRINKQUANTITY_TXT.Text = ""
GUESTDISHQUANTITY_TXT.Text = ""
GUESTRESTOTOTCHARGE_TXT.Text = ""
CUSTDRINKQUANTITY_TXT.Text = ""
CUSTDISHQUANTITY_TXT.Text = ""
CUSTRESTOTOTCHARGE_TXT.Text = ""
Call Form_Load
CUST_FRAME.Refresh
GUEST_FRAME.Refresh
S = "SELECT COUNT(CUST_NAME)FROM CUST WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
Set R = C.Execute(S)
NO_CUST = R.Fields(0)
If NO_CUST = 0 Then
    ID_TXT.Text = "C1"
Else
    ID_TXT.Text = "C" & (NO_CUST + 1)
End If
NAME_TXT.Enabled = True
NAME_TXT.SetFocus
NAME_TXT.Text = ""
GUEST_FRAME.Enabled = False
CUSTDISHQUANTITY_TXT.Text = "0"
CUSTDRINKQUANTITY_TXT.Text = "0"
CUST_FRAME.Enabled = True
End Sub

Private Sub CUSTRESTOTOTCHARGE_TXT_Change()
If Val(CUSTRESTOTOTCHARGE_TXT.Text) > 0 Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub

Private Sub CUSTTAKEORDER_CMD_Click()
If Val(CUSTDRINKRATE_LBL.Caption) < 0 Then
    If CUSTDRINKQUANTITY_TXT.Text = "" Then
        MsgBox "PLEASE INPUT QUANTITY OF DRINK"
        CUSTDRINKQUANTITY_TXT.SetFocus
    End If
ElseIf Val(CUSTDISHRATE_LBL.Caption) < 0 Then
    MsgBox "PLEASE INPUT QUANTITY OF DISH"
    CUSTDISHQUANTITY_TXT.SetFocus
Else
    If MsgBox("ARE YOU SURE TO TAKE ORDER ?", vbYesNo) = vbYes Then
        CUSTRESTOTOTCHARGE_TXT.Text = (Val(CUSTDRINKQUANTITY_TXT.Text) * Val(CUSTDRINKRATE_LBL.Caption) + Val(CUSTDISHQUANTITY_TXT.Text) * Val(CUSTDISHRATE_LBL.Caption))
        S = "INSERT INTO CUST VALUES('" & ID_TXT.Text & "','" & NAME_TXT.Text & "','" & CUSTRESTOTOTCHARGE_TXT.Text & "','" & Format(Date, "DD-MMM-YYYY") & "','" & CUSTDRINK_CMB.Text & "','" & CUSTDISH_CMB.Text & "'," & Val(CUSTDRINKQUANTITY_TXT.Text) * Val(CUSTDRINKRATE_LBL.Caption) & "," & Val(CUSTDISHQUANTITY_TXT.Text) * Val(CUSTDISHRATE_LBL.Caption) & ")"
        Set R = C.Execute(S)
        CUSTTAKEORDER_CMD.Enabled = False
        Command1.SetFocus
    Else
        CUSTRESTOTOTCHARGE_TXT.Text = "0"
        CUSTDISHTYPE_LBL.Caption = ""
        CUSTDISHRATE_LBL.Caption = ""
        CUSTDRINKTYPE_LBL.Caption = ""
        CUSTDRINKRATE_LBL.Caption = ""
        GUESTDRINKTYPE_LBL.Caption = ""
        GUESTDRINKRATE_LBL.Caption = ""
        GUESTDISHTYPE_LBL.Caption = ""
        GUESTDISHRATE_LBL.Caption = ""
        GUESTDRINKQUANTITY_TXT.Text = ""
        GUESTDISHQUANTITY_TXT.Text = ""
        GUESTRESTOTOTCHARGE_TXT.Text = ""
        CUSTDRINKQUANTITY_TXT.Text = ""
        CUSTDISHQUANTITY_TXT.Text = ""
        CUSTRESTOTOTCHARGE_TXT.Text = ""
        Call Form_Load
        CUST_FRAME.Refresh
        GUEST_FRAME.Refresh
        S = "SELECT COUNT( DISTINCT CUST_NAME)FROM CUST WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
        Set R = C.Execute(S)
        NO_CUST = R.Fields(0)
        If NO_CUST = 0 Then
            ID_TXT.Text = "C1"
        Else
            ID_TXT.Text = "C" & (NO_CUST + 1)
        End If
        NAME_TXT.Enabled = True
        NAME_TXT.SetFocus
        NAME_TXT.Text = ""
        GUEST_FRAME.Enabled = False
        CUSTDISHQUANTITY_TXT.Text = "0"
        CUSTDRINKQUANTITY_TXT.Text = "0"
        CUST_FRAME.Enabled = True
        CUSTTAKEORDER_CMD.Enabled = False
        Call Form_Load
    End If
End If
End Sub

Private Sub Form_Load()
Me.Top = 1000
Me.Left = 5000
CUSTDRINK_CMB.Clear
GUESTDRINK_CMB.Clear
CUSTDISH_CMB.Clear
GUESTDISH_CMB.Clear
GUESTID_CMB.Clear
CON
S = "SELECT CLIENT_ID FROM CLIENT_MASTER"
Set R = C.Execute(S)
Do Until R.EOF = True
    GUESTID_CMB.AddItem R.Fields("CLIENT_ID")
    R.MoveNext
Loop
S = "SELECT DNAME FROM DRINK "
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTDRINK_CMB.AddItem R.Fields("DNAME")
GUESTDRINK_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT DNAME FROM DISH "
Set R = C.Execute(S)
Do Until R.EOF = True
CUSTDISH_CMB.AddItem R.Fields("DNAME")
GUESTDISH_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
End Sub

Private Sub GUESTBAR_OPT_Click()
GUESTDRINKTYPE_LBL.Caption = ""
GUESTDRINKRATE_LBL.Caption = ""
GUESTDISHTYPE_LBL.Caption = ""
GUESTDISHRATE_LBL.Caption = ""
CUSTDISHTYPE_LBL.Caption = ""
CUSTDISHRATE_LBL.Caption = ""
CUSTDRINKTYPE_LBL.Caption = ""
CUSTDRINKRATE_LBL.Caption = ""
GUESTDRINKQUANTITY_TXT.Text = ""
GUESTDISHQUANTITY_TXT.Text = ""
GUESTRESTOTOTCHARGE_TXT.Text = ""
CUSTDRINKQUANTITY_TXT.Text = ""
CUSTDISHQUANTITY_TXT.Text = ""
CUSTRESTOTOTCHARGE_TXT.Text = ""
Call Form_Load
CUST_FRAME.Refresh
GUEST_FRAME.Refresh
ID_TXT.Enabled = False
NAME_TXT.Enabled = False
ID_TXT.Text = "ID"
NAME_TXT.Text = "CUSTOMER NAME"
GUEST_FRAME.Enabled = True
CUST_FRAME.Enabled = False
GUESTDISHQUANTITY_TXT.Text = "0"
GUESTDRINKQUANTITY_TXT.Text = "0"
End Sub

Private Sub GUESTBARPAYMENT_Click()
If MsgBox("PAYMENT NOW  THEN CLICK ON  'YES'  OTHERWISE AT THE TIME OF CHECKOUT THEN CLICK ON  'NO'", vbYesNo) = vbYes Then
    If MsgBox("ARE SURE FOR PAYMENT " & Val(GUESTRESTOTOTCHARGE_TXT.Text) & " RUPEES ", vbYesNo) = vbYes Then
        S = "UPDATE DAILY SET TOT_INCOME=TOT_INCOME+" & GUESTRESTOTOTCHARGE_TXT.Text & " WHERE TODAYDATE='" & Format(Date, "DD-MMM-YYYY") & "'"
MsgBox S
        Set R = C.Execute(S)
        S = "UPDATE CLIENT_MASTER SET DUES_BALANCE=DUES_BALANCE-" & Val(GUESTRESTOTOTCHARGE_TXT.Text) & " WHERE CLIENT_ID='" & GUESTID_CMB.Text & "'"
        Set R = C.Execute(S)
        S = "COMMIT"
        Set R = C.Execute(S)
        If MsgBox("ARE YOU TAKE REPORT THEN CLICK 'YES'", vbYesNo) = vbYes Then
            If DataEnvironment1.rsCommand5.State = 1 Then DataEnvironment1.rsCommand5.Close
                DataEnvironment1.Command5 GUESTID_CMB.Text
                DataReport7.Show
                GUESTBARPAYMENT.Enabled = False
        End If
    End If
Else
    If MsgBox("ARE YOU TAKE REPORT THEN CLICK 'YES'", vbYesNo) = vbYes Then
            If DataEnvironment1.rsCommand5.State = 1 Then DataEnvironment1.rsCommand5.Close
                DataEnvironment1.Command5 GUESTID_CMB.Text
                DataReport7.Show
                GUESTBARPAYMENT.Enabled = False
            End If
End If
GUESTRESTOTOTCHARGE_TXT.Text = "0"
NAME_LBL.Caption = ""
GUESTDRINKTYPE_LBL.Caption = ""
GUESTDRINKRATE_LBL.Caption = ""
GUESTDISHTYPE_LBL.Caption = ""
GUESTDISHRATE_LBL.Caption = ""
CUSTDISHTYPE_LBL.Caption = ""
CUSTDISHRATE_LBL.Caption = ""
CUSTDRINKTYPE_LBL.Caption = ""
CUSTDRINKRATE_LBL.Caption = ""
GUESTDRINKQUANTITY_TXT.Text = ""
GUESTDISHQUANTITY_TXT.Text = ""
GUESTRESTOTOTCHARGE_TXT.Text = ""
CUSTDRINKQUANTITY_TXT.Text = ""
CUSTDISHQUANTITY_TXT.Text = ""
CUSTRESTOTOTCHARGE_TXT.Text = ""
Call Form_Load
End Sub

Private Sub GUESTDISH_CMB_Click()
CON
S = "SELECT *FROM DISH WHERE DNAME='" & GUESTDISH_CMB.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
Else
GUESTDISHTYPE_LBL.Caption = R.Fields("TYPE")
GUESTDISHRATE_LBL.Caption = R.Fields("CHARGE")
End If
End Sub



Private Sub GUESTDISHQUANTITY_TXT_Change()
If Val(GUESTDISHQUANTITY_TXT.Text) > 0 Then
    GUESTTAKEBARORDER_CMD.Enabled = True
Else
    GUESTTAKEBARORDER_CMD.Enabled = False
End If
End Sub

Private Sub GUESTDRINK_CMB_Click()
CON
S = "SELECT *FROM DRINK WHERE DNAME='" & GUESTDRINK_CMB.Text & "'"
Set R = C.Execute(S)
GUESTDRINKTYPE_LBL.Caption = R.Fields("TYPE")
GUESTDRINKRATE_LBL.Caption = R.Fields("CHARGE")
End Sub

Private Sub GUESTDRINKQUANTITY_TXT_Change()
If Val(GUESTDRINKQUANTITY_TXT.Text) > 0 Then
    GUESTTAKEBARORDER_CMD.Enabled = True
Else
    GUESTTAKEBARORDER_CMD.Enabled = False
End If
End Sub

Private Sub GUESTID_CMB_Click()
CON
S = "SELECT NAME FROM CLIENT_MASTER WHERE CLIENT_ID='" & GUESTID_CMB.Text & "'"
Set R = C.Execute(S)
If R.EOF = True Then
Else
NAME_LBL.Caption = R.Fields("NAME")
End If
End Sub
Private Sub GUESTNEWBARORDER_CMD_Click()
GUESTRESTOTOTCHARGE_TXT.Text = "0"
NAME_LBL.Caption = ""
GUESTDRINKTYPE_LBL.Caption = ""
GUESTDRINKRATE_LBL.Caption = ""
GUESTDISHTYPE_LBL.Caption = ""
GUESTDISHRATE_LBL.Caption = ""
CUSTDISHTYPE_LBL.Caption = ""
CUSTDISHRATE_LBL.Caption = ""
CUSTDRINKTYPE_LBL.Caption = ""
CUSTDRINKRATE_LBL.Caption = ""
GUESTDRINKQUANTITY_TXT.Text = ""
GUESTDISHQUANTITY_TXT.Text = ""
GUESTRESTOTOTCHARGE_TXT.Text = ""
CUSTDRINKQUANTITY_TXT.Text = ""
CUSTDISHQUANTITY_TXT.Text = ""
CUSTRESTOTOTCHARGE_TXT.Text = ""
Call Form_Load
End Sub

Private Sub GUESTRESTOTOTCHARGE_TXT_Change()
If Val(GUESTRESTOTOTCHARGE_TXT.Text) > 0 Then
    GUESTBARPAYMENT.Enabled = True
Else
    GUESTBARPAYMENT.Enabled = False
End If
End Sub

Private Sub GUESTTAKEBARORDER_CMD_Click()
If Val(GUESTDRINKRATE_LBL.Caption) < 0 Then
    If GUESTDRINKQUANTITY_TXT.Text = "" Then
        MsgBox "PLEASE INPUT QUANTITY OF DRINK"
        GUESTDRINKQUANTITY_TXT.SetFocus
    End If
ElseIf Val(GUESTDISHRATE_LBL.Caption) < 0 Then
    If GUESTDISHQUANTITY_TXT.Text = "" Then
        MsgBox "PLEASE INPUT QUANTITY OF DISH"
        GUESTDISHQUANTITY_TXT.SetFocus
    End If
Else
    If MsgBox("ARE YOU SURE TO TAKE ORDER ?", vbYesNo) = vbYes Then
        GUESTRESTOTOTCHARGE_TXT.Text = (Val(GUESTDRINKQUANTITY_TXT.Text) * Val(GUESTDRINKRATE_LBL.Caption) + Val(GUESTDISHQUANTITY_TXT.Text) * Val(GUESTDISHRATE_LBL.Caption))
        CON
        S = "UPDATE CLIENT_MASTER SET TOT_CHARGE=TOT_CHARGE+" & Val(GUESTRESTOTOTCHARGE_TXT.Text) & ",DUES_BALANCE=DUES_BALANCE+" & Val(GUESTRESTOTOTCHARGE_TXT.Text) & " ,RESTO_CHARGE=RESTO_CHARGE+" & Val(GUESTRESTOTOTCHARGE_TXT.Text) & " WHERE CLIENT_ID='" & GUESTID_CMB.Text & "'"
        MsgBox S
        Set R = C.Execute(S)
        S = "COMMIT"
        Set R = C.Execute(S)
        MsgBox "HIS DUES IS UPDATED"
        GUESTTAKEBARORDER_CMD.Enabled = False
        GUESTBARPAYMENT.SetFocus
    Else
        GUESTRESTOTOTCHARGE_TXT.Text = "0"
        NAME_LBL.Caption = ""
        GUESTDRINKTYPE_LBL.Caption = ""
        GUESTDRINKRATE_LBL.Caption = ""
        GUESTDISHTYPE_LBL.Caption = ""
        GUESTDISHRATE_LBL.Caption = ""
        CUSTDISHTYPE_LBL.Caption = ""
        CUSTDISHRATE_LBL.Caption = ""
        CUSTDRINKTYPE_LBL.Caption = ""
        CUSTDRINKRATE_LBL.Caption = ""
        GUESTDRINKQUANTITY_TXT.Text = ""
        GUESTDISHQUANTITY_TXT.Text = ""
        GUESTRESTOTOTCHARGE_TXT.Text = ""
        CUSTDRINKQUANTITY_TXT.Text = ""
        CUSTDISHQUANTITY_TXT.Text = ""
        CUSTRESTOTOTCHARGE_TXT.Text = ""
        Call Form_Load
        GUESTTAKEBARORDER_CMD.Enabled = False
    End If
End If
End Sub
