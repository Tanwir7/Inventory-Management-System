VERSION 5.00
Begin VB.Form ItemEntry 
   Caption         =   "Item Entry"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtitem 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtprice 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Click to add the item to the stock"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      ToolTipText     =   "Click to remove the selected item from the stock"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox lstitem 
      Height          =   2205
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Tk."
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   1920
      Width           =   255
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
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
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
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
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
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lbl5 
      Caption         =   "Item List"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "ItemEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim itemrs As Recordset
Dim con As Connection

Private Sub cmddel_Click()
'**** Checks for any item selected
If lstitem.Text = Empty Then
'**** Generates msgbox
   MsgBox "No item Selected"
   Exit Sub
Else
'**** Deletes item information from the database
   Set itemrs = New Recordset
'**** Searches for item in Item_Master table
   itemrs.Open "delete from Item_Master where ItemName='" & lstitem.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Generates msgbox for confirmation
   MsgBox "Record is deleted"
'**** Opens Item_Master table
   itemrs.Open "Select*From Item_Master", con, adOpenDynamic, adLockOptimistic
   lstitem.Clear
'**** Removes the deleted item from the list
   While itemrs.EOF = False
      lstitem.AddItem (itemrs.Fields("ItemName"))
      itemrs.MoveNext
   Wend
End If
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdsave_Click()
'**** Checks for empty boxes in the form
If txtitem.Text = Empty Or txtprice.Text = Empty Or txtqty.Text = Empty Or txtorder.Text = Empty Then
'**** Generates a msgbox
   MsgBox "All informations are required"
   Exit Sub
Else
'**** Saves information of the item in the database
   Set itemrs = New Recordset
'**** Opens Item_Master table
   itemrs.Open "Select*From Item_Master", con, adOpenDynamic, adLockOptimistic
   '**** Adds new item to the table according to the field
   itemrs.AddNew
   itemrs.Fields("ItemName") = txtitem.Text
   itemrs.Fields("Price") = txtprice.Text
   itemrs.Fields("Quantity") = txtqty.Text
   itemrs.Fields("ReorderLevel") = txtorder.Text
   '**** Saves the information in the table
   itemrs.Update
   '**** Generates a msgbox
   MsgBox "Item has been added to the stock"
   itemrs.MoveFirst
   lstitem.Clear
'**** Displays the added item in the list
   While itemrs.EOF = False
      lstitem.AddItem (itemrs.Fields("ItemName"))
      itemrs.MoveNext
   Wend
'**** Clears the text boxes for next item entry
   txtitem.Text = Empty
   txtqty.Text = Empty
   txtprice.Text = Empty
   txtorder.Text = Empty
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
Set itemrs = New Recordset
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Opens Item_Master table
itemrs.Open "select *from Item_Master", con, adOpenDynamic, adLockOptimistic
'**** Loads and displays available item names from database
While itemrs.EOF = False
   lstitem.AddItem (itemrs.Fields("ItemName"))
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

Private Sub txtprice_KeyPress(KeyAscii As Integer)
'**** Checks for invalid characters
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
'**** Checks for invalid characters
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub
