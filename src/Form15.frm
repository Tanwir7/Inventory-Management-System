VERSION 5.00
Begin VB.Form Itemupdate 
   Caption         =   "Item Re-order Update"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpqty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox cmbiname 
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Text            =   "Select Item"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Click to update the item details"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtprice 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtsqty 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Purchased Quantity"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Reorder Level"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Price"
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
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Stock Quantity"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2280
      Width           =   255
   End
End
Attribute VB_Name = "Itemupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim con As Connection
Dim itemrs As Recordset

Private Sub cmbiname_Click()
Set itemrs = New Recordset
'**** Opens Item_Master table and searches item from the table
itemrs.Open "Select*from Item_Master where ItemName='" & cmbiname.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Loads information of the item from the table
txtsqty.Text = itemrs.Fields("Quantity")
txtprice.Text = itemrs.Fields("Price")
txtorder.Text = itemrs.Fields("ReorderLevel")
End Sub

Private Sub cmdupdate_Click()
Set itemrs = New Recordset
'**** Checks for empty boxes
If cmbiname.Text = "Select Item" Or txtprice.Text = Empty Or txtpqty.Text = Empty Or txtorder.Text = Empty Then
'**** Displys error msg for incomplete information
   MsgBox ("All informations need to be provided for update!")
Else
   '**** Opens Item_Master table
   itemrs.Open "Select*from Item_Master where ItemName='" & cmbiname.Text & "'", con, adOpenDynamic, adLockOptimistic
   '**** Overwrites information according to their respective field
   itemrs.Fields("Quantity") = itemrs.Fields("Quantity") + Val(txtpqty.Text)
   itemrs.Fields("Price") = txtprice.Text
   itemrs.Fields("ReorderLevel") = txtorder.Text
   '**** Updates information of the item in the table
   itemrs.Update
   '**** Generates msgbox for confirmation
   MsgBox ("Your Item Stock has been updated")
   '**** Clears the boxes for next item update
   cmbiname.Text = "Select Item"
   txtsqty.Text = Empty
   txtpqty.Text = Empty
   txtprice.Text = Empty
   txtorder.Text = Empty
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
Set itemrs = New Recordset
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Opens Item_Master table
itemrs.Open "select *from Item_Master", con, adOpenDynamic, adLockOptimistic
'**** Displays available items from database
While itemrs.EOF = False
   cmbiname.AddItem (itemrs.Fields("ItemName"))
   itemrs.MoveNext
Wend
End Sub

Private Sub txtorder_KeyPress(KeyAscii As Integer)
'**** Checks for invalid characters
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtpqty_KeyPress(KeyAscii As Integer)
'**** Checks for invalid characters
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
'**** Checks for invalid characters
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub
