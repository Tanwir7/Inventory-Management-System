VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form Help 
   Caption         =   "Help"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      _cx             =   5080
      _cy             =   5080
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc As String



Private Sub Form_Load()
loc = "F:\IT Project\A2 level\Documentation\Future Net User Guide.pdf"
Me.AcroPDF1.LoadFile (loc)
End Sub
