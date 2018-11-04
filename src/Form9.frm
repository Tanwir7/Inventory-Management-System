VERSION 5.00
Begin VB.Form ServiceRateEntry 
   Caption         =   "Service Rate Entry"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   720
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   32
      Top             =   960
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   6120
      Picture         =   "Form9.frx":0A03
      ScaleHeight     =   675
      ScaleWidth      =   1395
      TabIndex        =   31
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Height          =   1575
      Left            =   4200
      TabIndex        =   17
      Top             =   2040
      Width           =   3615
      Begin VB.TextBox txtscan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdscan 
         Caption         =   "Update"
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         ToolTipText     =   "Updates current rate to new rate"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Tk."
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "/page"
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Scanning"
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
         Left            =   1320
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      ToolTipText     =   "Closes the form"
      Top             =   5520
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Height          =   1575
      Left            =   4200
      TabIndex        =   11
      Top             =   3720
      Width           =   3615
      Begin VB.CommandButton cmdbwprint 
         Caption         =   "Update"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         ToolTipText     =   "Updates current rate to new rate"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtbwprint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Tk."
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "/page"
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Black and White Printing"
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
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   480
      TabIndex        =   8
      Top             =   3720
      Width           =   3615
      Begin VB.CommandButton cmdcprint 
         Caption         =   "Update"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         ToolTipText     =   "Updates current rate to new rate"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtcprint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Tk."
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "/page"
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Colour Printing"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
      Begin VB.CommandButton cmdcopy 
         Caption         =   "Update"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         ToolTipText     =   "Updates current rate to new rate"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtcopy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Tk."
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "/page"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Photocopy"
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
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.CommandButton cmdweb 
         Caption         =   "Update"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Updates current rate to new rate"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtweb 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Tk."
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "/min"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Web Browsing"
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
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "ServiceRateEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim raters As Recordset
Dim con As Connection

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdweb_Click()
'**** Checks for empty boxes
If txtweb.Text = Empty Then
'**** Generates a msgbox
   MsgBox ("Please provide the rate")
   Exit Sub
Else
'**** Updates the rate to the database
   Set raters = New Recordset
   '**** Opens the ServiceRate_Master table
   raters.Open "select*from ServiceRate_Master", con, adOpenDynamic, adLockOptimistic
   '**** Overwrites the data in the respective field
   raters.Fields("WebBrowse") = txtweb.Text
   '**** Saves the information in the table
   raters.Update
   '**** Generates a msgbox for confirmation
   MsgBox ("Rate has been updated successfully")
   '**** Clears the text box
   txtweb.Text = Empty
End If
End Sub

Private Sub cmdcopy_Click()
'**** Checks for empty boxes
If txtcopy.Text = Empty Then
'**** Generates a msgbox
   MsgBox ("Please provide the rate")
   Exit Sub
Else
'**** Updates the rate to the database
   Set raters = New Recordset
   '**** Opens the ServiceRate_Master table
   raters.Open "select*from ServiceRate_Master", con, adOpenDynamic, adLockOptimistic
   '**** Overwrites the data in the respective field
   raters.Fields("Photocopy") = txtcopy.Text
   '**** Saves the information in the table
   raters.Update
   '**** Generates a msgbox for confirmation
   MsgBox ("Rate has been updated successfully")
   '**** Clears the text box
   txtcopy.Text = Empty
End If
End Sub

Private Sub cmdscan_Click()
'**** Checks for empty boxes
If txtscan.Text = Empty Then
'**** Generates a msgbox
   MsgBox ("Please provide the rate")
   Exit Sub
Else
'**** Updates the rate to the database
   Set raters = New Recordset
   '**** Opens the ServiceRate_Master table
   raters.Open "select*from ServiceRate_Master", con, adOpenDynamic, adLockOptimistic
   '**** Overwrites the data in the respective field
   raters.Fields("Scanning") = txtscan.Text
   '**** Saves the information in the table
   raters.Update
   '**** Generates a msgbox for confirmation
   MsgBox ("Rate has been updated successfully")
   '**** Clears the text box
   txtscan.Text = Empty
End If
End Sub

Private Sub cmdcprint_Click()
'**** Checks for empty boxes
If txtcprint.Text = Empty Then
'**** Generates a msgbox
   MsgBox ("Please provide the rate")
   Exit Sub
Else
'**** Updates the rate to the database
   Set raters = New Recordset
   '**** Opens the ServiceRate_Master table
   raters.Open "select*from ServiceRate_Master", con, adOpenDynamic, adLockOptimistic
   '**** Overwrites the data in the respective field
   raters.Fields("ColourPrint") = txtcprint.Text
   '**** Saves the information in the table
   raters.Update
   '**** Generates a msgbox for confirmation
   MsgBox ("Rate has been updated successfully")
   '**** Clears the text box
   txtcprint.Text = Empty
End If
End Sub

Private Sub cmdbwprint_Click()
'**** Checks for empty boxes
If txtbwprint.Text = Empty Then
'**** Generates a msgbox
   MsgBox ("Please provide the rate")
   Exit Sub
Else
'**** Updates the rate to the database
   Set raters = New Recordset
   '**** Opens the ServiceRate_Master table
   raters.Open "select*from ServiceRate_Master", con, adOpenDynamic, adLockOptimistic
   '**** Overwrites the data in the respective field
   raters.Fields("BWPrint") = txtbwprint.Text
   '**** Saves the information in the table
   raters.Update
   '**** Generates a msgbox for confirmation
   MsgBox ("Rate has been updated successfully")
   '**** Clears the text box
   txtbwprint.Text = Empty
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\IT Project\A2 level\Databases\FUTURENET.mdb"
End Sub


Private Sub txtbwprint_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtcopy_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtcprint_KeyPress(KeyAscii As Integer)
'**** Checks for invalid caharacter
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtscan_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

