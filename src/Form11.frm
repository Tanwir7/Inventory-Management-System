VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ServicingEntry 
   Caption         =   "Computer Servicing"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      ToolTipText     =   "Prints a memo"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      ToolTipText     =   "Closes the form"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "New"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      ToolTipText     =   "Refreshes the entire form"
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      ToolTipText     =   "Stores servicing information of the customer"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtdue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox txtadvance 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTP 
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   109117443
      CurrentDate     =   41371
   End
   Begin VB.TextBox txtinfo 
      Height          =   1695
      Left            =   2520
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtnum 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin Crystal.CrystalReport cr 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label12 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label lbldate 
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      Left            =   4320
      TabIndex        =   16
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Due"
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
      TabIndex        =   12
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Advance"
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
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      Left            =   840
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Delivery Date"
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
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Description"
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
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Name"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblservice 
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Service No."
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
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "ServicingEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim servicingrs As Recordset
Dim con As Connection
Dim str, s As String

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdprint_Click()
'**** Creates acknowledgement slip
cr.ReportFileName = "F:\IT Project\A2 level\Programming\Report\ServicingSlip.rpt"
cr.SelectionFormula = "{Servicing_Master.ServiceNum}='" & lblservice.Caption & "'"
cr.Action = 2
End Sub

Private Sub cmdrefresh_Click()
Set servicingrs = New Recordset
'**** Opens Servicing_Master table
servicingrs.Open "select*from Servicing_Master", con, adOpenDynamic, adLockOptimistic
'**** Auto-generates a new service number by using previous service Num
servicingrs.MoveLast
str = servicingrs.Fields("ServiceNum")
str = Mid(str, 9, 4)
s = Val(str)
s = Val(s) + 1
lblservice.Caption = "Service-" & s
'**** Refreshes the entire form
txtname.Text = Empty
txtinfo.Text = Empty
txtnum.Text = Empty
DTP.Value = Format(Now, "dd/mm/yyyy")
txtamount.Text = Empty
txtadvance.Text = Empty
txtdue.Text = Empty
End Sub

Private Sub cmdsave_Click()
'**** Checks for empty boxes
If txtname.Text = Empty Or txtnum.Text = Empty Or txtinfo.Text = Empty Or txtamount.Text = Empty Or txtdue.Text = Empty Then
'**** Generates error msgbox for empty information
   MsgBox "Fill up the entire form"
   Exit Sub
Else
   Set servicingrs = New Recordset
'**** Opens Servicing_Master table
   servicingrs.Open "select*from Servicing_Master", con, adOpenDynamic, adLockOptimistic
'**** Sets a date format for DTP
   Dim d As Date
   d = Format(DTP.Value, "dd/mm/yyyy")
'**** Adds new Servicing details in the table according to respective field
   servicingrs.AddNew
   servicingrs.Fields("ServiceNum") = lblservice.Caption
   servicingrs.Fields("Name") = txtname.Text
   servicingrs.Fields("ContactNum") = txtnum.Text
   servicingrs.Fields("Description") = txtinfo.Text
   servicingrs.Fields("Amount") = txtamount.Text
   servicingrs.Fields("Advance") = txtadvance.Text
   servicingrs.Fields("Due") = txtdue.Text
   servicingrs.Fields("Delivery") = d
   servicingrs.Fields("Date") = lbldate.Caption
   servicingrs.Fields("Status") = "Not delivered"
   servicingrs.MoveNext
'**** Saves added servicing details in the database
   servicingrs.Update
'**** Generates msgbox for confirmation
   MsgBox ("Record Saved")
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
Set servicingrs = New Recordset
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Opens Servicing_Master table
servicingrs.Open "select*from Servicing_Master", con, adOpenDynamic, adLockOptimistic
'**** Auto-generates a new service number by searching previous Service Num
servicingrs.MoveLast
str = servicingrs.Fields("ServiceNum")
str = Mid(str, 9, 4)
s = Val(str)
s = Val(s) + 1
lblservice.Caption = "Service-" & s
'**** Displays current system date
lbldate.Caption = Format(Now, "dd/mm/yyyy")
'**** Sets DTP to current date format
DTP.Value = Format(Now, "dd/mm/yyyy")
'**** Initializes due to zero
txtdue.Text = 0
End Sub

Private Sub txtadvance_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtadvance_Change()
'**** Calculates the due
txtdue.Text = Val(txtamount.Text) - Val(txtadvance.Text)
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

Private Sub txtdue_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
'**** Checks for invalid character
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr(KeyAscii)) Then
'**** Generates error msgbox for entering invalid character
   MsgBox "Invalid character"
   KeyAscii = 0
End If
End Sub
