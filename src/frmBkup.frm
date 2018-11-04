VERSION 5.00
Begin VB.Form frmBkup 
   BackColor       =   &H8000000A&
   Caption         =   "Daily Database Backup"
   ClientHeight    =   4485
   ClientLeft      =   3225
   ClientTop       =   4290
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   Picture         =   "frmBkup.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   9645
   Begin VB.TextBox textfile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   3000
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtpath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdbackup 
      Caption         =   "Database Download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "FILE NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "SELECTED PATH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER PATH:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmBkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileSystemObject As Object
Dim filename As String
Dim d, fs As Object

Private Sub Command1_Click()
'**** Exits the form
Unload Me
End Sub

Private Sub Form_Load()
Me.Drive1.Refresh
Me.Dir1.Refresh
Me.Drive1.Refresh
Me.textfile.Text = "UDS" + Format$(Now, "d-mm-YYYY")
End Sub

Private Sub Dir1_Change()
Me.txtpath.Text = "" & Dir1.Path
End Sub

Private Sub Drive1_Change()
Set fs = CreateObject("Scripting.FileSystemObject")
Set d = fs.getdrive(fs.getdrivename(Drive1.Drive))
If d.isready Then
    Dir1.Path = Drive1.Drive
    Dir1.SetFocus
Else
'**** Generates error msgbox
    MsgBox "DRIVE IS NOT READY!!", "Warning"
End If
End Sub

Private Sub cmdbackup_Click()
filename = "" + Me.txtpath.Text + "\" + Me.textfile.Text + ".mdb"
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.copyfile "F:\IT Project\A2 level\Databases\FUTURENET.mdb", filename
'**** Displays msgbox for confirmation
MsgBox "DATA IS SAVED "
Me.Drive1.SetFocus
End Sub
