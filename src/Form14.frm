VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SalesBilling 
   Caption         =   "Sales Billing"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   5640
      Width           =   855
   End
   Begin Crystal.CrystalReport cr 
      Left            =   5280
      Top             =   240
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
      Left            =   1920
      TabIndex        =   21
      Top             =   5640
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Details"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
      Begin VB.ComboBox cmbiname 
         Height          =   315
         ItemData        =   "Form14.frx":0000
         Left            =   1440
         List            =   "Form14.frx":0002
         TabIndex        =   10
         Text            =   "Select Item"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtqty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtprice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdatl 
         Caption         =   "Add to cart"
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         ToolTipText     =   "Adds item to the cart"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Tk."
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Item Name"
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
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Unit Price"
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
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox txtnet 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdcalc 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      ToolTipText     =   "Calculates total amount added to cart"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdpaid 
      Caption         =   "Paid"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      ToolTipText     =   "Stores the Customer sales information"
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Closes the form"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Empty cart"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid mfg 
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.Label lbldate 
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Invoice No."
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
      TabIndex        =   17
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lbl 
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
      Left            =   1200
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   3240
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblinvoice 
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "SalesBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all variables
Dim r As Integer
Dim itemrs, accountrs As Recordset
Dim con As Connection
Dim str, s As String

Private Sub cmbiname_Click()
Set itemrs = New Recordset
'**** Opens the Item_Master table and Searches for the price of the selected item
itemrs.Open "Select*from Item_Master where ItemName='" & cmbiname.Text & "'", con, adOpenDynamic, adLockOptimistic
txtprice.Text = itemrs.Fields("Price")
End Sub

Private Sub cmdatl_Click()
'**** Checks for empty boxes and other errors
Set itemrs = New Recordset
If cmbiname.Text = "Select Item" And txtprice.Text = Empty Then
'**** Error msgbox generated for no selection
   MsgBox ("No Item(s) Selected")
   Exit Sub
ElseIf txtqty.Text = Empty Then
'**** Error msgbox generated for no details
   MsgBox ("Quantity should be present")
   Exit Sub
End If
itemrs.Open "Select*from Item_Master where ItemName='" & cmbiname.Text & "'", con, adOpenDynamic, adLockOptimistic
If itemrs.Fields("Quantity") = 0 Then
'**** Error msgbox generated
   MsgBox ("Out of stock")
   Exit Sub
ElseIf itemrs.Fields("Quantity") <= itemrs.Fields("ReorderLevel") Then
'**** Warning msgbox generated
   MsgBox ("Items are running low. Please Reorder")
ElseIf itemrs.Fields("Quantity") <= Val(txtqty.Text) Then
'**** Error msgbox generated
   MsgBox "Inputted quantity is exceeding stock level"
   Exit Sub
End If
'**** Adds purchased item details to the mfg table
r = r + 1
mfg.TextMatrix(r, 0) = r
mfg.TextMatrix(r, 1) = cmbiname.Text
mfg.TextMatrix(r, 2) = txtqty.Text
mfg.TextMatrix(r, 3) = txtprice.Text
mfg.TextMatrix(r, 4) = mfg.TextMatrix(r, 2) * mfg.TextMatrix(r, 3)
'**** Adds another row to the mfg table
mfg.Rows = mfg.Rows + 1
'**** Clears boxes for next item input
cmbiname.Text = "Select Item"
txtqty.Text = Empty
txtprice.Text = Empty
txtnet.Text = Empty
End Sub

Private Sub cmdcalc_Click()
'**** Checks for empty boxes
If mfg.TextMatrix(0, 1) = "Item Name" And mfg.TextMatrix(0, 2) = "Quantity" And mfg.TextMatrix(0, 3) = "Price" And mfg.TextMatrix(0, 4) = "Total" And mfg.TextMatrix(0, 0) = "Sl.no" And mfg.Rows = 2 Then
'**** Error msgbox generated for no details
   MsgBox ("No item(s) added to cart")
   Exit Sub
End If
'**** Calculates net total from items purchased
For i = 1 To r
   txtnet.Text = Val(txtnet.Text) + Val(mfg.TextMatrix(i, 4))
Next i
End Sub

Private Sub cmdclear_Click()
'**** Clears and resets mfg table
mfg.Clear
mfg.TextMatrix(0, 1) = "Item Name"
mfg.TextMatrix(0, 2) = "Quantity"
mfg.TextMatrix(0, 3) = "Price"
mfg.TextMatrix(0, 4) = "Total"
mfg.TextMatrix(0, 0) = "Sl.no"
mfg.Rows = 2
r = 0
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdnew_Click()
Set itemrs = New Recordset
'**** Opens Sales_Master table
itemrs.Open "select *from Sales_Master", con, adOpenDynamic, adLockOptimistic
'**** Auto-generates a new invoice no. for next sales input
itemrs.MoveLast
str = itemrs.Fields("Invoice")
str = Mid(str, 7, 4)
s = Val(str)
s = Val(s) + 1
lblinvoice.Caption = "Sales-" & s
'**** Clears the boxes in the form
cmbiname.Text = "Select Item"
txtqty.Text = Empty
txtprice.Text = Empty
txtnet.Text = Empty
mfg.Clear
mfg.TextMatrix(0, 1) = "Item Name"
mfg.TextMatrix(0, 2) = "Quantity"
mfg.TextMatrix(0, 3) = "Price"
mfg.TextMatrix(0, 4) = "Total"
mfg.TextMatrix(0, 0) = "Sl.no"
mfg.Rows = 2
r = 0
End Sub

Private Sub cmdpaid_Click()
'**** Checks for empty boxes and mfg table
If cmbiname.Text = "Select Item" And txtqty.Text = Empty And txtprice.Text = Empty And mfg.TextMatrix(0, 1) = "Item Name" And mfg.TextMatrix(0, 2) = "Quantity" And mfg.TextMatrix(0, 3) = "Price" And mfg.TextMatrix(0, 4) = "Total" And mfg.TextMatrix(0, 0) = "Sl.no" And mfg.Rows = 2 Then
'**** Error msgbox generated for no details
   MsgBox ("No item(s) have been purchased")
   Exit Sub
ElseIf txtnet.Text = Empty Then
'**** Error msgbox generated for no details
   MsgBox "Please check whether you have calculated or not"
   Exit Sub
End If
Set itemrs = New Recordset
For i = 1 To r
   itemrs.Open "select *from Item_Master where ItemName='" & mfg.TextMatrix(i, 1) & "'", con, adOpenDynamic, adLockOptimistic
   itemrs.Fields("Quantity") = itemrs.Fields("Quantity") - mfg.TextMatrix(i, 2)
'**** Updates stock quantity level
   itemrs.Update
   itemrs.Close
Next i
Set itemrs = New Recordset
'**** Opens Sales_Details table
itemrs.Open "select *from Sales_Details", con, adOpenDynamic, adLockOptimistic
For i = 1 To r
'**** Adds sales information per item in the table
   itemrs.AddNew
   itemrs.Fields("Invoice") = lblinvoice.Caption
   itemrs.Fields("Item") = mfg.TextMatrix(i, 1)
   itemrs.Fields("Quantity") = mfg.TextMatrix(i, 2)
   itemrs.Fields("Price") = mfg.TextMatrix(i, 3)
   itemrs.Fields("Total") = mfg.TextMatrix(i, 4)
'**** Saves sales information per item in the table
   itemrs.Update
Next i
'**** Saves Sales information in the database
Set itemrs = New Recordset
'**** Opens Sales_Master table
itemrs.Open "select *from Sales_Master", con, adOpenDynamic, adLockOptimistic
'**** Adds summarized Sales information in the table
itemrs.AddNew
itemrs.Fields("Invoice") = lblinvoice.Caption
itemrs.Fields("Date") = lbldate.Caption
itemrs.Fields("NetTotal") = txtnet.Text
'**** Saves summarized Sales information in the table
itemrs.Update
itemrs.MoveNext
Set accountrs = New Recordset
'**** Opens Accounts table
accountrs.Open "Select*from Accounts", con, adOpenDynamic, adLockOptimistic
'**** Adds total amount information of the sales in the table
accountrs.AddNew
accountrs.Fields("Date") = lbldate.Caption
accountrs.Fields("Service") = "Sales"
accountrs.Fields("Amount") = txtnet.Text
'**** Saves total amount information of the sales in the table
accountrs.Update
'**** Generates msgbox for confirmation
MsgBox "Item(s) have been Purchased"
End Sub

Private Sub cmdprint_Click()
'**** Creates invoice
cr.ReportFileName = "F:\IT Project\A2 level\Programming\Report\CustomerSalesBilling.rpt"
cr.SelectionFormula = "{Sales_Master.Invoice}='" & lblinvoice.Caption & "'"
cr.Action = 2
End Sub

Private Sub Form_Load()
Set con = New Connection
Set itemrs = New Recordset
'**** Establishes Connection between form and database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Opens Item_Master table
itemrs.Open "select *from Item_Master", con, adOpenDynamic, adLockOptimistic
'**** Displays available items from the table to the combo box
itemrs.MoveFirst
While itemrs.EOF = False
   cmbiname.AddItem (itemrs.Fields("ItemName"))
   itemrs.MoveNext
Wend
itemrs.Close
'**** Opens Sales_Master table
itemrs.Open "select *from Sales_Master", con, adOpenDynamic, adLockOptimistic
'**** Auto-generates a new invoice no. by searching previous invoice from the table
itemrs.MoveLast
str = itemrs.Fields("Invoice")
str = Mid(str, 7, 4)
s = Val(str)
s = Val(s) + 1
lblinvoice.Caption = "Sales-" & s
'**** Resets the mfg table
r = 0
lbldate.Caption = Format(Now, "dd/mm/yyyy")
mfg.TextMatrix(0, 0) = "Sl.no"
mfg.TextMatrix(0, 1) = "Item Name"
mfg.TextMatrix(0, 2) = "Quantity"
mfg.TextMatrix(0, 3) = "Price"
mfg.TextMatrix(0, 4) = "Total"
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
'**** Checks for Invalid Characters
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

