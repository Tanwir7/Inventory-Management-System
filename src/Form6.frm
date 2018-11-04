VERSION 5.00
Begin VB.Form TariffSearch 
   Caption         =   "Tariff Search"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      ToolTipText     =   "Closes the form"
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox cmbspeed 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "Select Speed"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cmbplan 
      Height          =   315
      ItemData        =   "Form6.frx":0000
      Left            =   1800
      List            =   "Form6.frx":000A
      TabIndex        =   0
      Text            =   "Select plan"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Tk."
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Rent"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblrent 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "TariffSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim tariffrs As Recordset
Dim con As Connection

Private Sub cmbplan_Click()
Set tariffrs = New Recordset
'**** Opens Tariff_Master and Searches for selected plan
tariffrs.Open "Select*From Tariff_Master where Plan='" & cmbplan.Text & "'", con, adOpenDynamic, adLockOptimistic
cmbspeed.Clear
cmbspeed.Text = "Select Speed"
'**** Displays speed according to selected plan
While tariffrs.EOF = False
   If cmbplan.Text = "Unlimited (24 hrs)" Then
      cmbspeed.AddItem (tariffrs.Fields("Speed"))
      tariffrs.MoveNext
   ElseIf cmbplan.Text = "Night (10 p.m - 10 a.m)" Then
      cmbspeed.AddItem (tariffrs.Fields("Speed"))
      tariffrs.MoveNext
   End If
Wend
End Sub

Private Sub cmbspeed_Click()
Set tariffrs = New Recordset
'**** Opens Tariff_Master and Searches for Rent according to selected plan and speed
tariffrs.Open "Select*from Tariff_Master where Plan='" & cmbplan.Text & "' and Speed='" & cmbspeed.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Displays rent according to selected plan and speed
While tariffrs.EOF = False
   lblrent.Caption = tariffrs.Fields("Amount")
   tariffrs.MoveNext
Wend
End Sub


Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
End Sub

