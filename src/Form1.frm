VERSION 5.00
Begin VB.Form UserAccount 
   Caption         =   "UserAccount"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdrefresh2 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      ToolTipText     =   "Refreshes the entire form"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdsave2 
      Caption         =   "Save"
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      ToolTipText     =   "Creates a new user account"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtverify2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6840
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtnpass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6840
      PasswordChar    =   "*"
      TabIndex        =   20
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtopass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6840
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtcuser 
      Height          =   285
      Left            =   6840
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox cmbcuser 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   6840
      List            =   "Form1.frx":000A
      TabIndex        =   17
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      ToolTipText     =   "Refreshes the entire form"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      ToolTipText     =   "Creates a new user account"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtverify 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtpass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtuser 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox cmbuser 
      Height          =   315
      ItemData        =   "Form1.frx":001E
      Left            =   2160
      List            =   "Form1.frx":0028
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Current User ID"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Current User Type"
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
      Left            =   4800
      TabIndex        =   15
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Verify Password"
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
      Left            =   4800
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "New Password"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Old Password"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   9240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   4560
      X2              =   4560
      Y1              =   3840
      Y2              =   0
   End
   Begin VB.Label Label5 
      Caption         =   "User Creation"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Verify Password"
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
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
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
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User Type"
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
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "UserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim loginrs, cpassrs As Recordset
Dim con As Connection
Dim str As String

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdrefresh_Click()
'**** Clears the boxes of the form
cmbuser.Text = Empty
txtuser.Text = Empty
txtpass.Text = Empty
txtverify.Text = Empty
End Sub

Private Sub cmdrefresh2_Click()
'**** Clears the boxes of the form
cmbcuser.Text = Empty
txtuser.Text = Empty
txtopass.Text = Empty
txtnpass.Text = Empty
txtverify2.Text = Empty
End Sub

Private Sub cmdsave_Click()
Set loginrs = New Recordset
'**** Opens User_Master table
loginrs.Open "Select*from User_Master where Username='" & txtuser.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Checks for existing user name
While loginrs.EOF = False
   If txtuser.Text = loginrs.Fields("Username") Then
'**** Generates error msgbox
   MsgBox ("User Name has already been taken. Please select another.")
   Exit Sub
 End If
Wend
loginrs.Close
'**** Opens User_Master table
loginrs.Open "Select*from User_Master", con, adOpenDynamic, adLockOptimistic
str = txtpass.Text
'**** Checks for empty boxes and generates error msgbox for incomplete data
If cmbuser.Text = Empty Then
   MsgBox ("User type required")
ElseIf txtuser.Text = Empty Then
   MsgBox ("User Name required")
ElseIf Len(str) < 5 Then
   MsgBox ("Password length should be more than 5 characters")
ElseIf txtverify.Text = Empty Then
   MsgBox ("Password Veification Required")
'**** Verifies passwords
ElseIf txtpass.Text <> txtverify.Text Then
'**** Generates error msgbox for unmatched password
   MsgBox ("Passwords did not match")
ElseIf txtpass.Text = txtverify.Text Then
'**** Adds new user information in the table
   loginrs.AddNew
   loginrs.Fields("Usertype") = cmbuser.Text
   loginrs.Fields("Username") = txtuser.Text
   loginrs.Fields("Password") = txtpass.Text
   loginrs.Update
'**** Saves new user information in the table
   MsgBox ("New Login ID is created")
End If
End Sub

Private Sub cmdsave2_Click()
Set cpassrs = New Recordset
'**** Checks for empty boxes and incorrect information and generates error msgbox for incomplete data
If txtcuser.Text = Empty Then
   MsgBox ("User Name required")
   Exit Sub
End If
If cmbcuser.Text = Empty Then
   MsgBox ("User type required")
   Exit Sub
End If
If Len(txtnpass.Text) < 5 Then
   MsgBox ("Password length should be more than 5 characters")
   Exit Sub
End If
'**** Opens User_Master table and checks for the UserName
cpassrs.Open "Select*from User_Master where Username='" & txtcuser.Text & "'", con, adOpenDynamic, adLockOptimistic
If cpassrs.EOF = False Then
'**** Verifies the old password with the table
      If txtopass.Text <> cpassrs.Fields("Password") Then
'**** Generates error msgbox if unmatched
      MsgBox ("Invalid Old Password")
      Exit Sub
   End If
Else
   MsgBox ("Invalid user name")
   Exit Sub
End If
If txtverify2.Text = Empty Then
   MsgBox ("Password Veification Required")
   Exit Sub
End If
'**** Verifies the new password entered
If txtnpass.Text <> txtverify2.Text Then
'**** Error msgbox gererated if unmatched
   MsgBox ("Passwords did not match")
Else
'**** Overwrites old password with new password in the table
   cpassrs.Fields("Password") = txtnpass.Text
'**** Saves the password in the table
   cpassrs.Update
'**** Displays msgbox for confirmation
   MsgBox ("Your password is Changed")
End If
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Creates connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
End Sub
