VERSION 5.00
Begin VB.Form BrowsingSignIn 
   Caption         =   "Sign In"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   3240
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      ToolTipText     =   "Refreshes the entire form"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "Closes the form"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      ToolTipText     =   "Click to start web browsing"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtsignin 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   480
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
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lbluser 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Sign In"
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
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User No."
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
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "BrowsingSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim con As Connection
Dim browsers As Recordset
Dim str, s As String

Private Sub cmdclear_Click()
'**** Refreshes all the box from the form
txtname.Text = Empty
txtsignin.Text = Empty
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdstart_Click()
If txtname.Text = Empty Then
   MsgBox "Enter name of user"
   Exit Sub
End If
'**** Displays the current time from the system
txtsignin.Text = Format(Now, "hh:nn AM/PM")
Set browsers = New Recordset
'**** Opens table and adds data according to the respective fields
browsers.Open "Select*from Browser_Details", con, adOpenDynamic, adLockOptimistic
browsers.AddNew
   browsers.Fields("UserNum") = lbluser.Caption
   browsers.Fields("Name") = txtname.Text
   browsers.Fields("SignIn") = txtsignin.Text
   browsers.Fields("Date") = lbldate.Caption
'**** Saves data that were added in the table
browsers.Update
'**** Displays a Msgbox for confirmation
MsgBox ("Internet browsing has started")
'**** Auto-generates a new user num for the next browser to enter
browsers.MoveLast
str = browsers.Fields("UserNum")
str = Mid(str, 6, 4)
s = Val(str)
s = s + 1
lbluser.Caption = "Cafe-" & s
End Sub

Private Sub Form_Load()
Set browsers = New Recordset
Set con = New Connection
'**** Connects the form with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Loads the current date from the system
lbldate.Caption = Format(Now, "dd/mm/yyyy")
'**** Opens the Browser_Details table
browsers.Open "select*from Browser_Details", con, adOpenDynamic, adLockOptimistic
'**** Generates the next user num by searching previous user num from table
browsers.MoveLast
   str = browsers.Fields("UserNum")
   str = Mid(str, 6, 4)
   s = Val(str)
   s = Val(s) + 1
lbluser.Caption = "Cafe-" & s
End Sub
