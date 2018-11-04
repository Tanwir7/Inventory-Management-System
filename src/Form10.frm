VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CustomerService 
   Caption         =   "Customer Service"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   5250
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
      Left            =   1800
      TabIndex        =   20
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   960
      TabIndex        =   19
      ToolTipText     =   "Saves a record of the services used"
      Top             =   6000
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid mfg 
      Height          =   1575
      Left            =   600
      TabIndex        =   17
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   4095
      Begin VB.ComboBox cmbservice 
         Height          =   315
         ItemData        =   "Form10.frx":0000
         Left            =   2040
         List            =   "Form10.frx":0010
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtnop 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtamount 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdatl 
         Caption         =   "Add to list"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         ToolTipText     =   "Click to add service to the list"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Service"
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
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "No. of Pages"
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
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Amount"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Tk."
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   1320
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      ToolTipText     =   "Closes the form"
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "New"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Refreshes the entire form"
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label lblnum 
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Transaction No."
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
      Left            =   600
      TabIndex        =   16
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbldate 
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Grand Total"
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
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label5 
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
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "CustomerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim con As Connection
Dim servicedrs, servicemrs, accountrs As Recordset
Dim str, s As String
Dim r As Integer

Private Sub cmdatl_Click()
'**** Checks whether required boxes are left empty
If cmbservice.Text = Empty Then
'**** Generates error msgbox if no service is selected
   MsgBox ("No Service Selected")
   Exit Sub
ElseIf txtnop.Text = Empty Then
'**** Generates error msgbox if no. of pages are not entered
   MsgBox ("Insert number of pages used")
   Exit Sub
End If
'**** Adds item details information into the mfg table
r = r + 1
mfg.TextMatrix(r, 0) = r
mfg.TextMatrix(r, 1) = cmbservice.Text
mfg.TextMatrix(r, 2) = txtnop.Text
mfg.TextMatrix(r, 3) = txtamount.Text
'**** Increases number of rows as items are added
mfg.Rows = mfg.Rows + 1
'**** Calculates and displays the grand total
lbltotal.Caption = Val(txtamount.Text) + Val(lbltotal.Caption)
'**** Clears the boxes for next service input
cmbservice.Text = Empty
txtnop.Text = Empty
txtamount.Text = Empty
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdprint_Click()
'**** Creates invoice
cr.ReportFileName = "F:\IT Project\A2 level\Programming\Report\CustomerServiceBilling.rpt"
cr.SelectionFormula = "{Service_Master.Transaction}='" & lblnum.Caption & "'"
cr.Action = 2
End Sub

Private Sub cmdrefresh_Click()
Set servicedrs = New Recordset
'**** Auto-generates a new Transaction number for the next customer to enter
servicedrs.Open "select*from Service_Details", con, adOpenDynamic, adLockOptimistic
servicedrs.MoveLast
str = servicedrs.Fields("Transaction")
s = Val(str)
s = Val(s) + 1
lblnum.Caption = s
'**** Clears all the boxes in the form
cmbservice.Text = Empty
txtnop.Text = Empty
txtamount.Text = Empty
'**** Sets the grand total to Zero
lbltotal.Caption = 0
'**** Resets and labels the mfgbox
mfg.Clear
mfg.TextMatrix(0, 0) = "Sl.no"
mfg.TextMatrix(0, 1) = "Service"
mfg.TextMatrix(0, 2) = "No.of Pages"
mfg.TextMatrix(0, 3) = "Amount"
mfg.Rows = 2
r = 0
End Sub

Private Sub cmdsave_Click()
'**** Checks for empty mfg table
If mfg.TextMatrix(0, 0) = "Sl.no" And mfg.TextMatrix(0, 1) = "Service" And mfg.TextMatrix(0, 2) = "No.of Pages" And mfg.TextMatrix(0, 3) = "Amount" And mfg.Rows = 2 Then
'**** Generates msgbox if empty mfg table found
   MsgBox ("No service(s) were added to list")
   Exit Sub
End If
Set servicedrs = New Recordset
Set servicemrs = New Recordset
Set accountrs = New Recordset
'**** Opens Service_Details table
servicedrs.Open "Select*from Service_Details", con, adOpenDynamic, adLockOptimistic
'**** Opens Service_Master table
servicemrs.Open "Select*from Service_Master", con, adOpenDynamic, adLockOptimistic
'**** Adds information to the Service_Master table according to their fields
servicemrs.AddNew
servicemrs.Fields("Transaction") = lblnum.Caption
servicemrs.Fields("Total") = lbltotal.Caption
servicemrs.Fields("Date") = lbldate.Caption
'**** Saves the added information Service_Master table
servicemrs.Update
'**** Adds information to the Service_Details table according to their fields
For i = 1 To r
   servicedrs.AddNew
   servicedrs.Fields("Transaction") = lblnum.Caption
   servicedrs.Fields("Service") = mfg.TextMatrix(i, 1)
   servicedrs.Fields("NumofPages") = mfg.TextMatrix(i, 2)
   servicedrs.Fields("Amount") = mfg.TextMatrix(i, 3)
   servicedrs.Fields("Date") = lbldate.Caption
'**** Saves the added information in the Service_Master table
   servicedrs.Update
Next i
'**** Opens the Accounts table
accountrs.Open "Select*from Accounts", con, adOpenDynamic, adLockOptimistic
'**** Adds information in the Acccounts table according to the field
accountrs.AddNew
accountrs.Fields("Date") = lbldate.Caption
accountrs.Fields("Service") = "Print/Scan/Photocopy"
accountrs.Fields("Amount") = lbltotal.Caption
'**** Saves the added information in the Accounts table
accountrs.Update
'**** Generates a msgbox for confirmation
MsgBox "Your data is saved"
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Establishes connectiom of the form with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Displays current system date
lbldate.Caption = Format(Now, "dd/mm/yyyy")
'**** Intializes grand total to zero
lbltotal.Caption = 0
Set servicedrs = New Recordset
'**** Opens Service _Details table
servicedrs.Open "select*from Service_Details", con, adOpenDynamic, adLockOptimistic
'**** Generates the next code by searching previous code from database
servicedrs.MoveLast
str = servicedrs.Fields("Transaction")
s = Val(str)
s = Val(s) + 1
lblnum.Caption = s
'**** Labels the mfg table
mfg.TextMatrix(0, 0) = "Sl.no"
mfg.TextMatrix(0, 1) = "Service"
mfg.TextMatrix(0, 2) = "No.of Pages"
mfg.TextMatrix(0, 3) = "Amount"
End Sub

Private Sub txtnop_Change()
Set servicers = New Recordset
'**** Calculates and displays the cost for each service according to the no. of pages used by customer
servicers.Open "Select*from ServiceRate_Master", con, adOpenDynamic, adLockOptimistic
If cmbservice.Text = "Photocopy" Then
   txtamount.Text = servicers.Fields("Photocopy") * Val(txtnop.Text)
ElseIf cmbservice.Text = "Scanning" Then
   txtamount.Text = servicers.Fields("Scanning") * Val(txtnop.Text)
ElseIf cmbservice.Text = "Colour Printing" Then
   txtamount.Text = servicers.Fields("ColourPrint") * Val(txtnop.Text)
ElseIf cmbservice.Text = "B & W Printing" Then
   txtamount.Text = servicers.Fields("BWPrint") * Val(txtnop.Text)
End If
End Sub


Private Sub txtnop_KeyPress(KeyAscii As Integer)
'**** Validation for invalid characters being entered
If KeyAscii = 8 Then Exit Sub
   If Not IsNumeric(Chr(KeyAscii)) Then
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub
