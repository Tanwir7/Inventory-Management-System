VERSION 5.00
Begin VB.Form BrowsingSignOut 
   Caption         =   "Sign Out"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4740
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   315
      Left            =   480
      TabIndex        =   18
      ToolTipText     =   "Stores browsing details"
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   1920
      TabIndex        =   16
      ToolTipText     =   "Refreshes the entire form"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtname 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      ToolTipText     =   "Closes the form"
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      ToolTipText     =   "Click to end web browsing"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtduration 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtsignout 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtsignin 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Searches the User No."
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtuser 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Tk."
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "min"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lbldate 
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label7 
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
      Left            =   720
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Left            =   720
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Duration"
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
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Sign Out"
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
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
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
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "BrowsingSignOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim browsers, raters, accountrs As Recordset
Dim con As Connection
Dim Time1, Time2, Dur As Date
Dim min, str, hh, mm, h, m As Integer

Private Sub cmdclear_Click()
'**** Refreshes all the boxes from the form
txtuser.Text = Empty
txtname.Text = Empty
txtsignin.Text = Empty
txtsignout.Text = Empty
txtduration.Text = Empty
txtamount.Text = Empty
End Sub

Private Sub cmdend_Click()
Set browsers = New Recordset

'**** Opens the table and searches for User num
browsers.Open "Select*from Browser_Details where UserNum='" & txtuser.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Checks for the availabiity information in the box
If txtuser.Text = Empty Or browsers.EOF = True Or txtsignin.Text = Empty Then
   MsgBox "User did not sign in.Please search for User No."
   Exit Sub
Else
'**** Displays the current sign out time from the system and duration spent by the user
   txtsignout.Text = Format(Now, "hh:nn:00 AM/PM")
   Time2 = CDate(txtsignout.Text)
   Dur = Time2 - Time1
   str = Format(Dur, "hh:mm")
   h = Mid(str, 1, 2)
   m = Mid(str, 4, 5)
   hh = Val(h)
   mm = Val(m)
'**** Converts the duration in minutes
   min = Val(hh * 60) + Val(mm * 1)
   txtduration.Text = min

'**** Checks whether the duration is more than 10 minutes
If txtduration.Text < 10 Then
'**** Displays Msgbox alert
   MsgBox ("Minimum browsing time is 10 minutes")
'**** Clears the boxes
   txtsignout.Text = Empty
   txtduration.Text = Empty
   Exit Sub
Else
   Set raters = New Recordset
'**** Opens ServiceRate_Master table to search for the rate
   raters.Open "select*from ServiceRate_Master", con, adOpenDynamic, adLockOptimistic
'**** Calculates the amount
   txtamount.Text = min * Val(raters.Fields("WebBrowse"))
End If
End If
End Sub

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdsave_Click()
'**** Checks for any empty field
If txtuser.Text = Empty Or txtsignin.Text = Empty Or txtsignout.Text = Empty Or txtduration.Text = Empty Or txtamount.Text = Empty Then
   MsgBox "Unable to save. Make sure you have searched for User No."
   Exit Sub
Else
   Set browsers = New Recordset
   Set accountrs = New Recordset
   '**** Opens the Browser_Details table
   browsers.Open "Select*from Browser_Details where UserNum='" & txtuser.Text & "'", con, adOpenDynamic, adLockOptimistic
   '**** Adds rest of the data according to the particular field
   browsers.Fields("SignOut") = txtsignout.Text
   browsers.Fields("Duration") = txtduration.Text
   browsers.Fields("Amount") = txtamount.Text
   '**** Saves the rest of the informaton of that particular user in the Browser_Details table
   browsers.Update
   '**** Opens the accounts table
   accountrs.Open "Select*from Accounts", con, adOpenDynamic, adLockOptimistic
   '**** Adds data according to the field
   accountrs.AddNew
   accountrs.Fields("Date") = lbldate.Caption
   accountrs.Fields("Service") = "Web browsing"
   accountrs.Fields("Amount") = txtamount.Text
   '**** Saves the information in the Accounts table
   accountrs.Update
   '**** Displays a Msgbox as confirmation
   MsgBox "Data has been saved"
End If
End Sub

Private Sub cmdsearch_Click()
Set browsers = New Recordset
'**** Opens table and Searches for User num
browsers.Open "Select*from Browser_Details where UserNum='" & txtuser.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Checks for valid or existing user no.
If txtuser.Text = Empty Or browsers.EOF = True Then
    MsgBox ("Input a valid User No.")
    Exit Sub
Else
'**** Searches for cafe user information and displays them if found
   While browsers.EOF = False
      txtname.Text = browsers.Fields("Name")
      txtsignin.Text = browsers.Fields("SignIn")
      browsers.MoveNext
   Wend
   Time1 = CDate(txtsignin.Text)
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Establishes connection of the form with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
'**** Displays current date from the system
lbldate.Caption = Format(Now, "dd/mm/yyyy")
End Sub
