VERSION 5.00
Begin VB.Form UserLogin 
   Caption         =   "User Login"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4410
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      ToolTipText     =   "Closes the form"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox cmbuser 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   2160
      List            =   "Form2.frx":000A
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtuser 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtpass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      TabIndex        =   6
      Top             =   840
      Width           =   1095
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
      TabIndex        =   5
      Top             =   1320
      Width           =   735
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
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "LOGIN"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "UserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**** Declaration of all the variables
Dim loginrs As Recordset
Dim con As Connection

Private Sub cmdexit_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub cmdlogin_Click()
'**** Checks for empty boxes and incorrect information
If txtuser.Text = Empty Then
'**** Generates error msgbox to check is user name is left empty
   MsgBox "User Name Required"
   Exit Sub
ElseIf txtpass.Text = Empty Then
'**** Generates error msgbox to check if password is left empty
   MsgBox "Password Required"
   Exit Sub
End If
Set loginrs = New Recordset
'**** Opens User_Master table and Searches for the user name
loginrs.Open "Select*from User_Master where Username='" & txtuser.Text & "'", con, adOpenDynamic, adLockOptimistic
'**** Checks for invalid user type or password
While loginrs.EOF = False
   If txtpass.Text <> loginrs.Fields("Password") Or cmbuser.Text <> loginrs.Fields("Usertype") Then
'**** Generates an error msgbox
      MsgBox ("Invalid User name or Password")
      Exit Sub
'**** Proceeds to next form if correct data is entered
   Else
'**** Opens the MDI form
      MDI.Show
      If loginrs.Fields("Usertype") = "General" Then
'**** Access rights given for general users
      MDI.mnsr.Enabled = False
      MDI.mi.Enabled = False
      MDI.mnua.Enabled = False
      MDI.mnte.Enabled = False
      End If
'**** Hides the login form
      UserLogin.Hide
      loginrs.MoveNext
'**** Clears all the boxes in the login form
      cmbuser.Text = Empty
      txtuser.Text = Empty
      txtpass.Text = Empty
   End If
Wend
End Sub

Private Sub Form_Load()
Set con = New Connection
'**** Establishes connection with the database
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\IT Project\A2 level\Databases\FUTURENET.mdb"
End Sub
