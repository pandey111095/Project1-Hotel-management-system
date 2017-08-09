VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRMUSEPAS 
   Caption         =   "USER AND PASSWORD FORM"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17130
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   17130
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   12726
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "CREATE"
      TabPicture(0)   =   "FRMUSEPAS.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "UPDATE"
      TabPicture(1)   =   "FRMUSEPAS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "DELETE"
      TabPicture(2)   =   "FRMUSEPAS.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
   End
End
Attribute VB_Name = "FRMUSEPAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
