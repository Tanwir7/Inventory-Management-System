VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form DateSelect5 
   Caption         =   "Servicing Details Search"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "Go!"
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97517569
         CurrentDate     =   41407
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97517569
         CurrentDate     =   41407
      End
      Begin VB.Label Label2 
         Caption         =   "To"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "From"
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
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
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
End
Attribute VB_Name = "DateSelect5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgo_Click()
Dim sd As Date
Dim td As Date
sd = Format(DTP1.Value, "mm/dd/yyyy")
td = Format(DTP2.Value, "mm/dd/yyyy")
cr.ReportFileName = "F:\IT Project\A2 level\Programming\Report\ServicingDetails.rpt"
cr.SelectionFormula = "{Servicing_Master.Date}>=#" & sd & "# And {Servicing_Master.Date}<=#" & td & "#"
cr.Action = 2
End Sub

Private Sub Form_Load()
DTP1.Value = Format(Now, "dd/mm/yyyy")
DTP2.Value = Format(Now, "dd/mm/yyyy")
End Sub

