VERSION 5.00
Begin VB.Form HolderRegistration 
   Caption         =   "Home Internet Connection-Registration"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10215
   LinkTopic       =   "Form4"
   ScaleHeight     =   3825
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   7920
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtdate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5640
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8880
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      ToolTipText     =   "Refreshes the entire form"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      ToolTipText     =   "Stores the Holder's Information"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtmail 
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox cmbspeed 
      Height          =   315
      Left            =   4440
      TabIndex        =   10
      Text            =   "Select Speed"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox cmbplan 
      Height          =   315
      ItemData        =   "Form4.frx":0FF6
      Left            =   2760
      List            =   "Form4.frx":1000
      TabIndex        =   9
      Text            =   "Select Plan"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtaddress 
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox txtnum 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   1290
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label8 
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
      Left            =   4920
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblid 
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label7 
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
      Left            =   960
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "E-mail"
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
      Left            =   4920
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
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
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Package"
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
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
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
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Contact Number"
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
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Holder Name"
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
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "HolderRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim tariffrs As Recordset
Dim regrs As Recordset
Dim con As Connection
Dim str, s As String

Private Sub cmbplan_Click()
Set tariffrs = New Recordset
'**** Opens Tariff_Master table and searches for selected plan
tariffrs.Open "Select*From Tariff_Master where Plan='" & cmbplan.Text & "'", con, adOpenDynamic, adLockOptimistic
cmbspeed.Clear
txtamount.Text = Empty
cmbspeed.Text = "Select Speed"
'**** Loads speeds according to selected plan
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
'**** Opens Tariff_Master table and searches for selected plan and speed
tariffrs.Open "Select*from Tariff_Master where Plan='" & cmbplan.Text & "' and Speed='" & cmbspeed.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Displays amount according to selected speed and plan
While tariffrs.EOF = False
   txtamount.Text = tariffrs.Fields("Amount")
   tariffrs.MoveNext
Wend
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdrefresh_Click()
'**** Clears the boxes of the entire form
txtname.Text = Empty
txtnum.Text = Empty
txtaddress.Text = Empty
txtmail.Text = Empty
cmbplan.Text = "Select Plan"
cmbspeed.Clear
cmbspeed = "Select Speed"
txtamount.Text = Empty
End Sub

Private Sub cmdsave_Click()
Set regrs = New Recordset
'**** Checks for empty boxes
If txtname.Text = Empty Or txtnum.Text = Empty Or cmbplan.Text = "Select Plan" Or cmbspeed.Text = "Select Speed" Then
'**** Generates error message for incomplete information
   MsgBox ("Please fill up all the necessary details")
Else
'**** Opens Holder_Details table
   regrs.Open "Select*From Holder_Details", con, adOpenDynamic, adLockOptimistic
'**** Adds Holder details according to respective fields
   regrs.AddNew
   regrs.Fields("ID") = lblid.Caption
   regrs.Fields("Name") = txtname.Text
   regrs.Fields("ContactNumber") = txtnum.Text
   regrs.Fields("E-mail") = txtmail.Text
   regrs.Fields("Address") = txtaddress.Text
   regrs.Fields("Plan") = cmbplan.Text
   regrs.Fields("Speed") = cmbspeed.Text
   regrs.Fields("Amount") = txtamount.Text
   regrs.Fields("Date") = txtdate.Text
'**** Saves information of the holder in the table
   regrs.Update
'**** Displays msgbox for confirmation
   MsgBox ("New Internet holder has been registered successfully")
'**** Auto-generates a new holder id for registration of next holder
   regrs.MoveLast
   str = regrs.Fields("ID")
   str = Mid(str, 4, 3)
   s = Val(str)
   s = s + 1
   lblid.Caption = "FN-" & s
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
Set regrs = New Recordset
'**** Connects the form with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Loads and displays system date
txtdate.Text = Format(Now, "dd/mm/yyyy")
'**** Opens Holder_Details table
regrs.Open "select*from Holder_Details", con, adOpenDynamic, adLockOptimistic
'**** Auto-generates a new holder id by searching previous holder ID
regrs.MoveLast
str = regrs.Fields("ID")
str = Mid(str, 4, 3)
s = Val(str)
s = Val(s) + 1
lblid.Caption = "FN-" & s

End Sub


Private Sub txtnum_KeyPress(KeyAscii As Integer)
'**** Checks for invalid characters in the box
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub
