VERSION 5.00
Begin VB.Form ServicingDelivery 
   Caption         =   "Computer Delivery"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      ToolTipText     =   "Searches for details of the Service No."
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      ToolTipText     =   "Closes the form"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      ToolTipText     =   "Refreshes the entire form"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmddeliver 
      Caption         =   "Delivered"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Changes the  status of delivery"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtdue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtdate 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtinfo 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtname 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtnum 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Tk."
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lbldate 
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label5 
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
      Left            =   600
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   600
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
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
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "ServicingDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim servicingrs, accountrs As Recordset
Dim con As Connection

Private Sub cmddeliver_Click()
Set servicingrs = New Recordset
Set accountrs = New Recordset
'**** Checks for empty boxes
If txtnum.Text = Empty Or txtname.Text = Empty Or txtinfo.Text = Empty Or txtdate.Text = Empty Or txtdue.Text = Empty Then
'**** Generates error message for incorrect information
   MsgBox "Delivery not possible. Make sure you have searched for Service No."
   Exit Sub
Else
'**** Opens Servicing_Master table and searches for service num
   servicingrs.Open "Select*from Servicing_Master where ServiceNum='" & txtnum.Text & "'", con, adOpenDynamic, adLockOptimistic
   servicingrs.Fields("Status") = "Delivered"
   '**** Saves delivery information in the database
   servicingrs.Update
   '**** Open Accounts table
   accountrs.Open "Select*from Accounts", con, adOpenDynamic, adLockOptimistic
   '**** Adds Total Amount details according to field in the table
   accountrs.AddNew
   accountrs.Fields("Date") = lbldate.Caption
   accountrs.Fields("Service") = "Servicing"
   accountrs.Fields("Amount") = servicingrs.Fields("Amount")
   '****Saves Amount of service in the table
   accountrs.Update
   '**** Delivers msgbox for confirmation
   MsgBox "Delivered"
End If
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdrefresh_Click()
'**** Clears the form
txtnum.Text = Empty
txtname.Text = Empty
txtdate.Text = Empty
txtinfo.Text = Empty
txtdue.Text = Empty
End Sub

Private Sub cmdsearch_Click()
Set servicingrs = New Recordset
'**** Opens Servicing_Master table and searches for details according to service num
servicingrs.Open "Select*from Servicing_Master where ServiceNum='" & txtnum.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Checks for empty boxes
If txtnum.Text = Empty Or servicingrs.EOF = True Then
'**** Generates error message for incorrect details
   MsgBox ("Input a valid Service No.")
   Exit Sub
Else
'**** Searches and displays the servicing details from the table
   While servicingrs.EOF = False
      txtname.Text = servicingrs.Fields("Name")
      txtinfo.Text = servicingrs.Fields("Description")
      txtdate.Text = servicingrs.Fields("Delivery")
      txtdue.Text = servicingrs.Fields("Due")
      servicingrs.MoveNext
   Wend
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Displays current system date
lbldate.Caption = Format(Now, "dd/mm/yyyy")
End Sub
