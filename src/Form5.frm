VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form HolderMonthlyRent 
   Caption         =   "Holder's Monthly Rent"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5850
   LinkTopic       =   "Form5"
   ScaleHeight     =   2880
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport cr 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdstatus 
      Caption         =   "Check Status"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Checks whether the Holder has paid or not"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   "Refreshes the entire form"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      ToolTipText     =   "Closes the form"
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdpaid 
      Caption         =   "Paid"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "Click to pay the rent of the user"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtmonth 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtrent 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cmbid 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Text            =   "Select ID"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lbldate 
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Month"
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
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Holder ID"
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
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "HolderMonthlyRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim rentrs, accountrs As Recordset
Dim con As Connection

Private Sub cmbid_Click()
Set rentrs = New Recordset
'**** Opens Holder_Details table and searches for rent according to holder id
rentrs.Open "Select*from Holder_Details where ID='" & cmbid.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Displays rent registered holder ids
While rentrs.EOF = False
   txtrent.Text = rentrs.Fields("Amount")
   rentrs.MoveNext
Wend
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdpaid_Click()
'**** Checks for empty boxes
If cmbid.Text = "Select ID" Or txtrent.Text = Empty Then
'**** Generates error msgbox due to incomplete details
   MsgBox ("Select Holder ID")
   Exit Sub
Else
   Set rentrs = New Recordset
   Set accountrs = New Recordset
'**** Opens Holders_Rent table
   rentrs.Open "Select*from Holders_Rent", con, adOpenDynamic, adLockOptimistic
'**** Adds payment details in the table
   rentrs.AddNew
   rentrs.Fields("ID") = cmbid.Text
   rentrs.Fields("Month") = txtmonth.Text
   rentrs.Fields("Rent") = txtrent.Text
'**** Saves the payment information in the database
   rentrs.Update
'**** Opens Accounts table
   accountrs.Open "Select*from Accounts", con, adOpenDynamic, adLockOptimistic
'**** Adds payment details in the table
   accountrs.AddNew
   accountrs.Fields("Date") = lbldate.Caption
   accountrs.Fields("Service") = "Monthly Rent"
   accountrs.Fields("Amount") = txtrent.Text
'**** Saves the payment information in the database
   accountrs.Update
'**** Generates msgbox for confirmation
   MsgBox ("Holder's Rent is paid successfully")
End If
End Sub

Private Sub cmdprint_Click()
'**** Creates acknowledgement slip
cr.ReportFileName = "F:\IT Project\A2 level\Programming\Report\MonthlyRentSlip.rpt"
cr.SelectionFormula = "{Holder_Details.ID}='" & cmbid.Text & "'"
cr.Action = 2
End Sub

Private Sub cmdrefresh_Click()
'**** Clears all the boxes for next input
cmbid.Text = "Select ID"
txtmonth.Text = Format(Now, "MMMM-yyyy")
txtrent.Text = Empty
End Sub

Private Sub cmdstatus_Click()
'**** Checks for empty boxes
If cmbid.Text = Empty Then
   MsgBox ("Select Holder ID")
   Exit Sub
Else
'**** Checks payment status
   Set rentrs = New Recordset
   rentrs.Open "Select*from Holders_Rent where ID='" & cmbid.Text & "' and Month='" & txtmonth.Text & "'", con, adOpenDynamic, adLockOptimistic
   If rentrs.EOF = False Then
      MsgBox "Rent paid"
   Else
      MsgBox "Rent not paid"
   End If
End If
End Sub

Private Sub Form_Load()
Set rentrs = New Recordset
Set con = New Connection
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
rentrs.Open "Select*from Holder_Details", con, adOpenDynamic, adLockOptimistic
rentrs.MoveFirst
'**** Displays id of all the holders
While rentrs.EOF = False
   cmbid.AddItem (rentrs.Fields("ID"))
   rentrs.MoveNext
Wend
'**** Displays current sysytem date and month
txtmonth.Text = Format(Now, "MMMM-yyyy")
lbldate.Caption = Format(Now, "dd/mm/yyyy")
End Sub
