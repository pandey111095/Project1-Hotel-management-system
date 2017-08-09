VERSION 5.00
Begin VB.Form FRMDISHANDDRINKENTRY 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DISH AND DRINK ENTRY"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FRMDISHANDDRINKENTRY.frx":0000
   ScaleHeight     =   4245
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin VB.Frame ENTRY_FRAME 
      BackColor       =   &H8000000A&
      Caption         =   "NEW ENTRY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   2040
      TabIndex        =   13
      Top             =   480
      Width           =   4695
      Begin VB.CommandButton DRINKADD_CMD 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ADD NEW DRINK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton DRINKSAVE_CMD 
         BackColor       =   &H00808000&
         Caption         =   "SAVE DRINK"
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
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2280
         Width           =   1215
      End
      Begin VB.OptionButton DISHVEG_OPT 
         BackColor       =   &H8000000A&
         Caption         =   "VEG DISH"
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
         Left            =   1560
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton DISHNONVEG_OPT 
         BackColor       =   &H8000000A&
         Caption         =   "NON VEG DISH"
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
         Left            =   2760
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton DRINKWINE_OPT 
         BackColor       =   &H8000000A&
         Caption         =   "Wine"
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
         Left            =   1560
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton DRINKJUICE_OPT 
         BackColor       =   &H8000000A&
         Caption         =   "Juice"
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
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton DRINKBEER_OPT 
         BackColor       =   &H8000000A&
         Caption         =   "Beer"
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
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton DRINKSOFT_OPT 
         BackColor       =   &H8000000A&
         Caption         =   "Soft Drinks"
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
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox INPUDISHNAME_TXT 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
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
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox INPUDISHPRICE_TXT 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
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
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton DISHCANCEL_CMD 
         BackColor       =   &H000000FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton DISHADD_CMD 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ADD NEW DISH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton DISHSAVE_CMD 
         BackColor       =   &H00808000&
         Caption         =   "SAVE DISH"
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
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox DISHPRICE_TXT 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
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
         Left            =   5640
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox DISHNAME_TXT 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
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
         Left            =   5640
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "PRICE"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "NAME"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame UPDATE_FRAME 
      BackColor       =   &H8000000C&
      Caption         =   "UPDATE DATA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   7695
      Begin VB.CommandButton DRINKDELETE_CMD 
         BackColor       =   &H00FFC0FF&
         Caption         =   "DELETE DRINK"
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
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton DISHDELETE_CMD 
         BackColor       =   &H00FFC0FF&
         Caption         =   "DELETE DISH"
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
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton DISHUPDATE_CMD 
         BackColor       =   &H00404000&
         Caption         =   "UPDATE DISH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox VEGDISH_CMB 
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
         Height          =   330
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   29
         Text            =   "VEGDISH_CMB"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox NONVEGDISH_CMB 
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
         Height          =   330
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   28
         Text            =   "NONVEGDISH_CMB"
         Top             =   2640
         Width           =   2175
      End
      Begin VB.OptionButton VEGDISH_OPT 
         BackColor       =   &H8000000C&
         Caption         =   "VEG DISH"
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
         TabIndex        =   27
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton NONVEGDISH_OPT 
         BackColor       =   &H8000000C&
         Caption         =   "NON VEG DISH"
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
         TabIndex        =   26
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox PRICE_TXT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
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
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton DRINKUPDATE_CMD 
         BackColor       =   &H00404000&
         Caption         =   "UPDATE DRINK"
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
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox SOFTDRINK_CMB 
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
         Height          =   330
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   10
         Text            =   "SOFTDRINK_CMB"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox BEERDRINK_CMB 
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
         Height          =   330
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "BEERDRINK_CMB"
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox JUICEDRINK_CMB 
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
         Height          =   330
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "JUICEDRINK_CMB"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox WINEDRINK_CMB 
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
         Height          =   330
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "WINEDRINK_CMB"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton DRINKCANCLE_CMD 
         BackColor       =   &H000000FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton SOFTDRINK_OPT 
         BackColor       =   &H8000000C&
         Caption         =   "Soft Drinks"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton BEERDRINK_OPT 
         BackColor       =   &H8000000C&
         Caption         =   "Beer"
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
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton JUICEDRINK_OPT 
         BackColor       =   &H8000000C&
         Caption         =   "Juice"
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton WINEDRINK_OPT 
         BackColor       =   &H8000000C&
         Caption         =   "Wine"
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
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " NAME"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " PRICE"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "NEW  PRICE"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6480
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label NAME_LBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4320
         TabIndex        =   12
         Top             =   720
         Width           =   60
      End
      Begin VB.Label PRICE_LBL 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4320
         TabIndex        =   11
         Top             =   1440
         Width           =   60
      End
   End
End
Attribute VB_Name = "FRMDISHANDDRINKENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DRITYP As String
Dim DISHTYP As String
Private Sub BEERDRINK_OPT_Click()
SOFTDRINK_CMB.Text = ""
BEERDRINK_CMB.Text = ""
JUICEDRINK_CMB.Text = ""
WINEDRINK_CMB.Text = ""
VEGDISH_CMB.Text = ""
NONVEGDISH_CMB.Text = ""
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
DISHDELETE_CMD.Enabled = False
DRINKDELETE_CMD.Enabled = False
'DRINKSAVE_CMD.Enabled = False
DRINKUPDATE_CMD.Enabled = False
DRITYP = "BEER"
DISHUPDATE_CMD.Enabled = False
SOFTDRINK_CMB.Enabled = False
WINEDRINK_CMB.Enabled = False
JUICEDRINK_CMB.Enabled = False
BEERDRINK_CMB.Enabled = True
'DRINKADD_CMD.Enabled = True
VEGDISH_CMB.Enabled = False
NONVEGDISH_CMB.Enabled = False
End Sub

Private Sub DISH_FRAME_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub DISHADD_CMD_Click()
DISHVEG_OPT.Visible = True
DISHNONVEG_OPT.Visible = True
DRINKSOFT_OPT.Visible = False
DRINKJUICE_OPT.Visible = False
DRINKBEER_OPT.Visible = False
DRINKWINE_OPT.Visible = False
'DISHADD_CMD.Enabled = False
'DRINKADD_CMD.Enabled = False
INPUDISHNAME_TXT.Text = ""
INPUDISHPRICE_TXT.Text = ""
'INPUDISHNAME_TXT.SetFocus
DISHSAVE_CMD.Enabled = True
DRINKSAVE_CMD.Enabled = False
End Sub

Private Sub DISHCANCEL_CMD_Click()
Load MDIForm1
MDIForm1.Show
Unload Me
Me.Hide
End Sub

Private Sub DISHDELETE_CMD_Click()
If NAME_LBL.Caption = "" Then
MsgBox "PLEASE SELECT A DISH ?"
Else
If MsgBox("ARE YOU SURE TO DELETE? ", vbYesNo) = vbYes Then
S = "DELETE FROM DISH WHERE DNAME='" & NAME_LBL.Caption & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "DATA DELETED !!"
Call Form_Load
Else
Call Form_Load
End If
End If
End Sub

Private Sub DISHNONVEG_OPT_Click()
DISHTYP = "NONVEGDISH"
DISHSAVE_CMD.Enabled = True
DRINKSAVE_CMD.Enabled = False
INPUDISHNAME_TXT.Enabled = True
INPUDISHPRICE_TXT.Enabled = True
INPUDISHNAME_TXT.SetFocus
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Text = ""
End Sub

Private Sub DISHSAVE_CMD_Click()
If INPUDISHNAME_TXT.Text = "" Then
MsgBox "PLESE INPUT DISH NAME"
INPUDISHNAME_TXT.SetFocus
ElseIf INPUDISHPRICE_TXT.Text = "" Then
MsgBox "PLESE INPUT DISH PRICE"
INPUDISHPRICE_TXT.SetFocus
Else
If MsgBox("ARE YOU TO SAVE THIS ITEM ?", vbYesNo) = vbYes Then
CON
S = "INSERT INTO DISH VALUES('" & INPUDISHNAME_TXT.Text & "','" & DISHTYP & "','" & INPUDISHPRICE_TXT.Text & "')"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "ITEM IS SAVED"
INPUDISHNAME_TXT.Text = ""
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Enabled = False
INPUDISHPRICE_TXT.Enabled = False
'DISHADD_CMD.Enabled = False
DISHSAVE_CMD.Enabled = False
Call Form_Load
Else
MsgBox "ITEM IS NOT SAVED"
INPUDISHNAME_TXT.Text = ""
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Enabled = False
INPUDISHPRICE_TXT.Enabled = False
'DISHADD_CMD.Enabled = False
DISHSAVE_CMD.Enabled = False
Call Form_Load
End If
End If
End Sub

Private Sub DISHUPDATE_CMD_Click()
If NAME_LBL.Caption = "" Then
MsgBox "PLEASE SELECT A DISH "
ElseIf PRICE_TXT.Text = "" Then
MsgBox "PLESE INPUT PRICE"
PRICE_TXT.SetFocus
Else
If MsgBox("ARE YOU SURE TO UPDATE THIS ITEM ?", vbYesNo) = vbYes Then
CON
S = "UPDATE DISH SET CHARGE='" & PRICE_TXT.Text & "' WHERE TYPE='" & DRITYP & "' AND DNAME='" & NAME_LBL.Caption & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "ITEM IS UPDATED"
Call Form_Load
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
VEGDISH_CMB.Enabled = False
NONVEGDISH_CMB.Enabled = False
PRICE_TXT.Enabled = False
'DISHADD_CMD.Enabled = False
'DISHSAVE_CMD.Enabled = False
DISHUPDATE_CMD.Enabled = False
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
Else
MsgBox "ITEM IS NOT UPDATE"
PRICE_TXT.Text = ""
VEGDISH_CMB.Enabled = False
NONVEGDISH_CMB.Enabled = False
PRICE_TXT.Enabled = False
'DISHADD_CMD.Enabled = False
'DISHSAVE_CMD.Enabled = False
DISHUPDATE_CMD.Enabled = False
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
Call Form_Load
End If
End If
End Sub

Private Sub DISHVEG_OPT_Click()
DISHTYP = "VEGDISH"
DISHSAVE_CMD.Enabled = True
DRINKSAVE_CMD.Enabled = False
INPUDISHNAME_TXT.Enabled = True
INPUDISHPRICE_TXT.Enabled = True
INPUDISHNAME_TXT.SetFocus
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Text = ""
End Sub

Private Sub DRINKADD_CMD_Click()
DISHVEG_OPT.Visible = False
DISHNONVEG_OPT.Visible = False
DRINKSOFT_OPT.Visible = True
DRINKJUICE_OPT.Visible = True
DRINKBEER_OPT.Visible = True
DRINKWINE_OPT.Visible = True
INPUDISHNAME_TXT.Text = ""
INPUDISHPRICE_TXT.Text = ""
DISHSAVE_CMD.Enabled = False
DRINKSAVE_CMD.Enabled = True
End Sub

Private Sub DRINKBEER_OPT_Click()
DISHSAVE_CMD.Enabled = False
DRITYP = "BEER"
DRINKSAVE_CMD.Enabled = True
INPUDISHNAME_TXT.Enabled = True
INPUDISHPRICE_TXT.Enabled = True
INPUDISHNAME_TXT.SetFocus
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Text = ""
End Sub

Private Sub DRINKCANCLE_CMD_Click()
'Load MDIForm
'MDIForm1.Show
Unload Me
Me.Hide
End Sub

Private Sub DRINKDELETE_CMD_Click()
If NAME_LBL.Caption = "" Then
MsgBox "PLEASE SELECT A DRINK ?"
Else
If MsgBox("ARE YOU SURE TO DELETE? ", vbYesNo) = vbYes Then
S = "DELETE FROM DRINK WHERE DNAME='" & NAME_LBL.Caption & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "DATA DELETED !!"
Call Form_Load
Else
Call Form_Load
End If
End If
End Sub

Private Sub DRINKJUICE_OPT_Click()
DISHSAVE_CMD.Enabled = False
DRITYP = "JUICE"
DRINKSAVE_CMD.Enabled = True
INPUDISHNAME_TXT.Enabled = True
INPUDISHPRICE_TXT.Enabled = True
INPUDISHNAME_TXT.SetFocus
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Text = ""
End Sub

Private Sub DRINKSAVE_CMD_Click()
If INPUDISHNAME_TXT.Text = "" Then
MsgBox "PLESE INPUT DRINK NAME"
INPUDISHNAME_TXT.SetFocus
ElseIf INPUDISHPRICE_TXT.Text = "" Then
MsgBox "PLESE INPUT DRINK PRICE"
INPUDDISHPRICE_TXT.SetFocus
Else
If MsgBox("ARE YOU TO SAVE THIS ITEM ?", vbYesNo) = vbYes Then
CON
S = "INSERT INTO DRINK VALUES('" & INPUDISHNAME_TXT.Text & "','" & DRITYP & "','" & INPUDISHPRICE_TXT.Text & "')"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "ITEM IS SAVED"
INPUDISHNAME_TXT.Text = ""
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Enabled = False
INPUDISHPRICE_TXT.Enabled = False
'DRINKADD_CMD.Enabled = False
DRINKSAVE_CMD.Enabled = False
Call Form_Load
Else
MsgBox "ITEM IS NOT SAVED"
INPUDISHNAME_TXT.Text = ""
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Enabled = False
INPUDISHPRICE_TXT.Enabled = False
'DRINKADD_CMD.Enabled = False
DRINKSAVE_CMD.Enabled = False
Call Form_Load
End If
End If
End Sub


Private Sub BEERDRINK_CMB_Click()
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
'DRINKADD_CMD.Enabled = False
NAME_LBL.Caption = BEERDRINK_CMB.Text
DRINKDELETE_CMD.Enabled = True
DRINKUPDATE_CMD.Enabled = True
CON
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & BEERDRINK_CMB.Text & "'"
Set R = C.Execute(S)
PRICE_LBL.Caption = R.Fields("CHARGE")
PRICE_TXT.Enabled = True
End Sub

Private Sub DRINKSOFT_OPT_Click()
DISHSAVE_CMD.Enabled = False
DRITYP = "SOFTDRINK"
DRINKSAVE_CMD.Enabled = True
INPUDISHNAME_TXT.Enabled = True
INPUDISHPRICE_TXT.Enabled = True
INPUDISHNAME_TXT.SetFocus
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Text = ""
End Sub

Private Sub DRINKUPDATE_CMD_Click()
If NAME_LBL.Caption = "" Then
MsgBox "PLEASE SELECT A DRINK "
ElseIf PRICE_TXT.Text = "" Then
MsgBox "PLESE INPUT PRICE"
PRICE_TXT.SetFocus
Else
If MsgBox("ARE YOU SURE TO UPDATE THIS ITEM ?", vbYesNo) = vbYes Then
CON
S = "UPDATE DRINK SET CHARGE='" & PRICE_TXT.Text & "' WHERE TYPE='" & DRITYP & "' AND DNAME='" & NAME_LBL.Caption & "'"
Set R = C.Execute(S)
S = "COMMIT"
Set R = C.Execute(S)
MsgBox "ITEM IS UPDATED"
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
SOFTDRINK_CMB.Enabled = False
BEERDRINK_CMB.Enabled = False
JUICEDRINK_CMB.Enabled = False
WINEDRINK_CMB.Enabled = False
PRICE_TXT.Enabled = False
'DRINKADD_CMD.Enabled = False
'DRINKSAVE_CMD.Enabled = False
DRINKUPDATE_CMD.Enabled = False
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
Call Form_Load
Else
MsgBox "ITEM IS NOT UPDATE"
PRICE_TXT.Text = ""
SOFTDRINK_CMB.Enabled = False
BEERDRINK_CMB.Enabled = False
JUICEDRINK_CMB.Enabled = False
WINEDRINK_CMB.Enabled = False
PRICE_TXT.Enabled = False
'DRINKADD_CMD.Enabled = False
'DRINKSAVE_CMD.Enabled = False
DRINKUPDATE_CMD.Enabled = False
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
Call Form_Load
End If
End If
End Sub

Private Sub DRINKWINE_OPT_Click()
DRITYP = "WINE"
DISHSAVE_CMD.Enabled = False
DRINKSAVE_CMD.Enabled = True
INPUDISHNAME_TXT.Enabled = True
INPUDISHPRICE_TXT.Enabled = True
INPUDISHNAME_TXT.SetFocus
INPUDISHPRICE_TXT.Text = ""
INPUDISHNAME_TXT.Text = ""
End Sub

Private Sub Form_Load()
Me.Top = 1000
Me.Left = 5000
SOFTDRINK_OPT.Refresh
WINEDRINK_OPT.Refresh
BEERDRINK_OPT.Refresh
JUICEDRINK_OPT.Refresh
VEGDISH_OPT.Refresh
NONVEGDISH_OPT.Refresh
SOFTDRINK_CMB.Clear
WINEDRINK_CMB.Clear
BEERDRINK_CMB.Clear
JUICEDRINK_CMB.Clear
VEGDISH_CMB.Clear
NONVEGDISH_CMB.Clear
CON
S = "SELECT *FROM DISH WHERE TYPE='VEGDISH'"
Set R = C.Execute(S)
Do Until R.EOF
VEGDISH_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DISH WHERE TYPE='NONVEGDISH'"
Set R = C.Execute(S)
Do Until R.EOF
NONVEGDISH_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='SOFTDRINK'"
Set R = C.Execute(S)
Do Until R.EOF
SOFTDRINK_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='BEER'"
Set R = C.Execute(S)
Do Until R.EOF
BEERDRINK_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='JUICE'"
Set R = C.Execute(S)
Do Until R.EOF
JUICEDRINK_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
S = "SELECT *FROM DRINK WHERE TYPE='WINE'"
Set R = C.Execute(S)
Do Until R.EOF
WINEDRINK_CMB.AddItem R.Fields("DNAME")
R.MoveNext
Loop
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
End Sub


Private Sub INPUDISHNAME_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
INPUDISHPRICE_TXT.SetFocus
End If
End Sub

Private Sub INPUDISHPRICE_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
End Sub

Private Sub JUICEDRINK_CMB_Click()
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
'DRINKADD_CMD.Enabled = False
NAME_LBL.Caption = JUICEDRINK_CMB.Text
DRINKDELETE_CMD.Enabled = True
DRINKUPDATE_CMD.Enabled = True
CON
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & JUICEDRINK_CMB.Text & "'"
Set R = C.Execute(S)
PRICE_LBL.Caption = R.Fields("CHARGE")
PRICE_TXT.Enabled = True
End Sub

Private Sub JUICEDRINK_OPT_Click()
SOFTDRINK_CMB.Text = ""
BEERDRINK_CMB.Text = ""
JUICEDRINK_CMB.Text = ""
WINEDRINK_CMB.Text = ""
VEGDISH_CMB.Text = ""
NONVEGDISH_CMB.Text = ""
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
DISHDELETE_CMD.Enabled = False
DRINKDELETE_CMD.Enabled = False
'DRINKSAVE_CMD.Enabled = False
DRINKUPDATE_CMD.Enabled = False
DRITYP = "JUICE"
DISHUPDATE_CMD.Enabled = False
SOFTDRINK_CMB.Enabled = False
WINEDRINK_CMB.Enabled = False
JUICEDRINK_CMB.Enabled = True
BEERDRINK_CMB.Enabled = False
'DRINKADD_CMD.Enabled = True
VEGDISH_CMB.Enabled = False
NONVEGDISH_CMB.Enabled = False
End Sub

Private Sub NONVEGDISH_OPT_Click()
SOFTDRINK_CMB.Text = ""
BEERDRINK_CMB.Text = ""
JUICEDRINK_CMB.Text = ""
WINEDRINK_CMB.Text = ""
VEGDISH_CMB.Text = ""
NONVEGDISH_CMB.Text = ""
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
DISHDELETE_CMD.Enabled = False
DRINKDELETE_CMD.Enabled = False
'DISHSAVE_CMD.Enabled = False
DISHUPDATE_CMD.Enabled = False
VEGDISH_CMB.Enabled = False
DRITYP = "NONVEGDISH"
NONVEGDISH_CMB.Enabled = True
'DISHADD_CMD.Enabled = True
DRINKUPDATE_CMD.Enabled = False
SOFTDRINK_CMB.Enabled = False
WINEDRINK_CMB.Enabled = False
JUICEDRINK_CMB.Enabled = False
BEERDRINK_CMB.Enabled = False
End Sub

Private Sub Option4_Click()

End Sub

Private Sub Option6_Click()

End Sub

Private Sub PRICE_TXT_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = Asc(".") Then
Else
KeyAscii = 0
End If
End Sub

Private Sub SOFTDRINK_CMB_Click()
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
'DRINKADD_CMD.Enabled = False
NAME_LBL.Caption = SOFTDRINK_CMB.Text
DRINKDELETE_CMD.Enabled = True
DRINKUPDATE_CMD.Enabled = True
CON
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & SOFTDRINK_CMB.Text & "'"
Set R = C.Execute(S)
PRICE_LBL.Caption = R.Fields("CHARGE")
PRICE_TXT.Enabled = True
End Sub

Private Sub SOFTDRINK_OPT_Click()
SOFTDRINK_CMB.Text = ""
BEERDRINK_CMB.Text = ""
JUICEDRINK_CMB.Text = ""
WINEDRINK_CMB.Text = ""
VEGDISH_CMB.Text = ""
NONVEGDISH_CMB.Text = ""
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
DISHDELETE_CMD.Enabled = False
DRINKDELETE_CMD.Enabled = False
'DRINKSAVE_CMD.Enabled = False
DRINKUPDATE_CMD.Enabled = False
DISHUPDATE_CMD.Enabled = False
WINEDRINK_CMB.Enabled = False
JUICEDRINK_CMB.Enabled = False
BEERDRINK_CMB.Enabled = False
DRITYP = "SOFTDRINK"
SOFTDRINK_CMB.Enabled = True
VEGDISH_CMB.Enabled = False
NONVEGDISH_CMB.Enabled = False
'DRINKADD_CMD.Enabled = True
End Sub


Private Sub VEGDISH_CMB_Click()
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
'DISHADD_CMD.Enabled = False
NAME_LBL.Caption = VEGDISH_CMB.Text
DISHDELETE_CMD.Enabled = True
DISHUPDATE_CMD.Enabled = True
CON
S = "SELECT CHARGE FROM DISH WHERE DNAME='" & VEGDISH_CMB.Text & "'"
Set R = C.Execute(S)
PRICE_LBL.Caption = R.Fields("CHARGE")
PRICE_TXT.Enabled = True
End Sub

Private Sub NONVEGDISH_CMB_Click()
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
'DISHADD_CMD.Enabled = False
NAME_LBL.Caption = NONVEGDISH_CMB.Text
DISHUPDATE_CMD.Enabled = True
DISHDELETE_CMD.Enabled = True
CON
S = "SELECT CHARGE FROM DISH WHERE DNAME='" & NONVEGDISH_CMB.Text & "'"
Set R = C.Execute(S)
PRICE_LBL.Caption = R.Fields("CHARGE")
PRICE_TXT.Enabled = True
End Sub
Private Sub VEGDISH_OPT_Click()
SOFTDRINK_CMB.Text = ""
BEERDRINK_CMB.Text = ""
JUICEDRINK_CMB.Text = ""
WINEDRINK_CMB.Text = ""
VEGDISH_CMB.Text = ""
NONVEGDISH_CMB.Text = ""
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
DISHDELETE_CMD.Enabled = False
DRINKDELETE_CMD.Enabled = False
'DISHSAVE_CMD.Enabled = False
DISHUPDATE_CMD.Enabled = False
NONVEGDISH_CMB.Enabled = False
DRITYP = "VEGDISH"
VEGDISH_CMB.Enabled = True
'DISHADD_CMD.Enabled = True
DRINKUPDATE_CMD.Enabled = False
SOFTDRINK_CMB.Enabled = False
WINEDRINK_CMB.Enabled = False
JUICEDRINK_CMB.Enabled = False
BEERDRINK_CMB.Enabled = False
End Sub

Private Sub WINEDRINK_CMB_Click()
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
'DRINKADD_CMD.Enabled = False
NAME_LBL.Caption = WINEDRINK_CMB.Text
DRINKDELETE_CMD.Enabled = True
DRINKUPDATE_CMD.Enabled = True
CON
S = "SELECT CHARGE FROM DRINK WHERE DNAME='" & WINEDRINK_CMB.Text & "'"
Set R = C.Execute(S)
PRICE_LBL.Caption = R.Fields("CHARGE")
PRICE_TXT.Enabled = True
End Sub

Private Sub WINEDRINK_OPT_Click()
SOFTDRINK_CMB.Text = ""
BEERDRINK_CMB.Text = ""
JUICEDRINK_CMB.Text = ""
WINEDRINK_CMB.Text = ""
VEGDISH_CMB.Text = ""
NONVEGDISH_CMB.Text = ""
NAME_LBL.Caption = ""
PRICE_LBL.Caption = ""
PRICE_TXT.Text = ""
DISHDELETE_CMD.Enabled = False
DRINKDELETE_CMD.Enabled = False
'DRINKSAVE_CMD.Enabled = False
DRINKUPDATE_CMD.Enabled = False
DISHUPDATE_CMD.Enabled = False
DRITYP = "WINE"
SOFTDRINK_CMB.Enabled = False
WINEDRINK_CMB.Enabled = True
JUICEDRINK_CMB.Enabled = False
BEERDRINK_CMB.Enabled = False
'DRINKADD_CMD.Enabled = True
VEGDISH_CMB.Enabled = False
NONVEGDISH_CMB.Enabled = False
End Sub
