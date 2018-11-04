VERSION 5.00
Begin VB.Form TariffEntry 
   Caption         =   "Tariff Entry"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   LinkTopic       =   "Form3"
   ScaleHeight     =   4245
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      ToolTipText     =   "Closes the form"
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ammend/Delete Plan"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4200
      TabIndex        =   9
      Top             =   960
      Width           =   6135
      Begin VB.TextBox txtspeed2 
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmddel 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         ToolTipText     =   "Deletes the selected plan"
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdrefresh2 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         ToolTipText     =   "Refershes the Amend plan module"
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   600
         TabIndex        =   16
         ToolTipText     =   "Updates current rate to a new rate"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ListBox lstspeed 
         Height          =   1425
         ItemData        =   "Form3.frx":0000
         Left            =   4200
         List            =   "Form3.frx":0007
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtamount2 
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cmbplan2 
         Height          =   315
         ItemData        =   "Form3.frx":0015
         Left            =   2040
         List            =   "Form3.frx":001F
         TabIndex        =   13
         Text            =   "Select plan"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Tk."
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Speed"
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Set new Amount"
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
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Select Speed"
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
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Select Plan"
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
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create Plan"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   3375
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "Refreshes the creation plan module"
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         ToolTipText     =   "Creates a new browsing plan"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtamount 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtspeed 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cmbplan 
         Height          =   315
         ItemData        =   "Form3.frx":0050
         Left            =   1200
         List            =   "Form3.frx":005A
         TabIndex        =   2
         Text            =   "Select Plan"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Tk."
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   1560
         Width           =   255
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
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "TariffEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Connection
Dim tariffrs As Recordset
Dim tariffrs1 As Recordset
Dim holderrs As Recordset

Private Sub cmbplan2_Click()
Set tariffrs = New Recordset
'**** Searches and displays speed according to selected plan
tariffrs.Open "Select*From Tariff_Master where Plan='" & cmbplan2.Text & "'", con, adOpenDynamic, adLockOptimistic
lstspeed.Clear
While tariffrs.EOF = False
   If cmbplan2.Text = "Unlimited (24 hrs)" Then
      lstspeed.AddItem (tariffrs.Fields("Speed"))
   ElseIf cmbplan2.Text = "Night (10 p.m - 10 a.m)" Then
      lstspeed.AddItem (tariffrs.Fields("Speed"))
   End If
   tariffrs.MoveNext
Wend
End Sub

Private Sub cmddel_Click()
Set tariffrs = New Recordset
Set tariffrs1 = New Recordset
'**** Checks for empty boxes
If cmbplan2.Text = "Select Plan" Or txtspeed2.Text = Empty Then
'**** Generates Msgbox
   MsgBox ("Select a Plan and Speed!")
Else
'**** Opens the table and Deletes tariff according to selected Plan and Speed
   tariffrs.Open "delete from Tariff_Master where Plan='" & cmbplan2.Text & "' and Speed='" & lstspeed.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Generates msgbox for confirmation
   MsgBox "Data plan deleted successfully"
'**** Refreshes the list after the deletion
   tariffrs1.Open "Select*From Tariff_Master where Plan='" & cmbplan2.Text & "'", con, adOpenDynamic, adLockOptimistic
   lstspeed.Clear
   While tariffrs1.EOF = False
      lstspeed.AddItem (tariffrs1.Fields("Speed"))
      tariffrs1.MoveNext
   Wend
End If
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdrefresh_Click()
'**** Clears the boxes
cmbplan.Text = "Select Plan"
txtspeed.Text = Empty
txtamount.Text = Empty
End Sub

Private Sub cmdrefresh2_Click()
'**** Clears the boxes
cmbplan2.Text = "Select Plan"
txtamount2.Text = Empty
lstspeed.Clear
txtspeed2.Text = Empty
End Sub

Private Sub cmdsave_Click()
Set tariffrs = New Recordset
'**** Checks for empty boxes
If cmbplan.Text = "Select plan" Or txtspeed.Text = Empty Or txtamount.Text = Empty Then
   MsgBox ("Please provide all the information!")
Else
'**** Opens the Tariff_Master table
   tariffrs.Open "Select*From Tariff_Master", con, adOpenDynamic, adLockOptimistic
   '**** Adds information acording to their respective fields
   tariffrs.AddNew
   tariffrs.Fields("Plan") = cmbplan.Text
   tariffrs.Fields("Speed") = txtspeed.Text
   tariffrs.Fields("Amount") = txtamount.Text
   '**** Saves tariff information in the table
   tariffrs.Update
   '**** Generates msgbox for confirmation
   MsgBox "New Data Plan Created"
   tariffrs.MoveFirst
End If
End Sub
Private Sub cmdupdate_Click()
Set tariffrs = New Recordset
Set holderrs = New Recordset
'**** Checks for empty boxes
If cmbplan2.Text = "Select plan" Or txtspeed2.Text = Empty Or txtamount2.Text = Empty Then
'**** Generates error msgbox for incomplete information
   MsgBox ("All informations need to be provided for update!")
Else
'**** Opens the Tariff_Master table and searches for amount according to selected speed and plan
   tariffrs.Open "Select*from Tariff_Master where Plan='" & cmbplan2.Text & "' and Speed='" & txtspeed2.Text & "'", con, adOpenDynamic, adLockOptimistic
   '**** Overwrites the amount in the respective field in the table
   tariffrs.Fields("Amount") = txtamount2.Text
   '**** Updates tariff in the table
   tariffrs.Update
   '**** Opens Holder_Details and searches for the Tariff plan according to selected plan and speed
   holderrs.Open "Select*from Holder_Details where Plan='" & cmbplan2.Text & "' and Speed='" & txtspeed2.Text & "'", con, adOpenDynamic, adLockOptimistic
   While holderrs.EOF = False
      '**** Overwrites the existing amount with the new amount
      holderrs.Fields("Amount") = txtamount2.Text
      '**** Saves the information in the table
      holderrs.Update
      holderrs.MoveNext
   Wend
   '**** Displays msgbox for confirmation
   MsgBox ("Your data plan has been updated")
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\IT Project\A2 level\Databases\FUTURENET.mdb"
lstspeed.Clear
End Sub

Private Sub lstspeed_Click()
'**** Displays selected speed from listbox to textbox
txtspeed2.Text = lstspeed.Text
End Sub

Private Sub txtamount_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtamount2_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub
