VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H00000000&
   Caption         =   "FUTURE NET"
   ClientHeight    =   5010
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   7545
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm.frx":0000
   WindowState     =   2  'Maximized
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
   Begin VB.Menu mnentry 
      Caption         =   "Entry"
      Begin VB.Menu mnsr 
         Caption         =   "Services Rate"
      End
      Begin VB.Menu mnte 
         Caption         =   "Tarrif Entry"
      End
      Begin VB.Menu mniteme 
         Caption         =   "Item Entry"
      End
      Begin VB.Menu mniru 
         Caption         =   "Item Reorder Update"
      End
      Begin VB.Menu mnse 
         Caption         =   "Servicing Entry"
      End
      Begin VB.Menu mnhreg 
         Caption         =   "Holder's Registration"
      End
   End
   Begin VB.Menu mntrans 
      Caption         =   "Transaction"
      Begin VB.Menu mnb 
         Caption         =   "Browsing"
         Begin VB.Menu mns 
            Caption         =   "Sign in"
         End
         Begin VB.Menu mnso 
            Caption         =   "Sign Out"
         End
      End
      Begin VB.Menu mncs 
         Caption         =   "Customer Service"
      End
      Begin VB.Menu mnsb 
         Caption         =   "Sales Billing"
      End
      Begin VB.Menu mnsd 
         Caption         =   "Servicing Delivery"
      End
      Begin VB.Menu mnhmr 
         Caption         =   "Holder's Monthly Rent"
      End
   End
   Begin VB.Menu mnsearch 
      Caption         =   "Search"
      Begin VB.Menu mnt 
         Caption         =   "Tariff"
      End
   End
   Begin VB.Menu mr 
      Caption         =   "Report"
      Begin VB.Menu mi 
         Caption         =   "Accounts"
      End
      Begin VB.Menu mcbd 
         Caption         =   "Cafe Browse Details"
      End
      Begin VB.Menu msd 
         Caption         =   "Service Details"
      End
      Begin VB.Menu mhd 
         Caption         =   "Holder Details"
      End
      Begin VB.Menu msale 
         Caption         =   "Sales Details"
      End
      Begin VB.Menu mpc 
         Caption         =   "Servicing Details"
      End
   End
   Begin VB.Menu mntools 
      Caption         =   "Tools"
      Begin VB.Menu mncb 
         Caption         =   "Create Backup"
      End
      Begin VB.Menu mnua 
         Caption         =   "User Account"
      End
      Begin VB.Menu mnlg 
         Caption         =   "Log out"
      End
   End
   Begin VB.Menu mnhelp 
      Caption         =   "Help"
      Begin VB.Menu mntrouble 
         Caption         =   "Troubleshooting"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mcbd_Click()
'**** Shows Date Selection form for Cafe Browse Details
DateSelect2.Show
End Sub

Private Sub mhd_Click()
'**** Shows record of the holders as a report
cr.ReportFileName = "F:\IT Project\A2 level\Programming\Report\Holderdetail.rpt"
cr.Action = 2
End Sub

Private Sub mi_Click()
'**** Shows date selection form for accounts
DateSelect.Show
End Sub

Private Sub mncb_Click()
'**** Shows back up form
frmBkup.Show
End Sub

Private Sub mncs_Click()
'**** Shows Customer Service form
CustomerService.Show
End Sub

Private Sub mnhmr_Click()
'**** Shows Holder's Monthly Rent form
HolderMonthlyRent.Show
End Sub

Private Sub mnhreg_Click()
'**** Shows Holder Registration form
HolderRegistration.Show
End Sub

Private Sub mniru_Click()
'**** Shows Item Reorder Update form
Itemupdate.Show
End Sub

Private Sub mniteme_Click()
'**** Shows Item Entry form
ItemEntry.Show
End Sub

Private Sub mnlg_Click()
'**** Closes current form
Unload Me
'**** Opens User Login form
UserLogin.Show
End Sub

Private Sub mns_Click()
'**** Shows Sign In Form for Cafe Browsing
BrowsingSignIn.Show
End Sub

Private Sub mnsb_Click()
'**** Show Sales Billing form
SalesBilling.Show
End Sub

Private Sub mnsd_Click()
'**** Shows Servicing Delivery form
ServicingDelivery.Show
End Sub

Private Sub mnse_Click()
'**** Shows Servicing Entry form
ServicingEntry.Show
End Sub

Private Sub mnso_Click()
'**** Shows Sign Out form for Cafe Browsing
BrowsingSignOut.Show
End Sub

Private Sub mnsr_Click()
'**** Shows Service Rate Entry form
ServiceRateEntry.Show
End Sub

Private Sub mnt_Click()
'**** Shows Tariff Search form
TariffSearch.Show
End Sub

Private Sub mnte_Click()
'**** Shows the Tariff Entry form
TariffEntry.Show
End Sub

Private Sub mntrouble_Click()
'**** Shows the Help form
Help.Show
End Sub

Private Sub mnua_Click()
'**** Shows the User Account form
UserAccount.Show
End Sub

Private Sub mpc_Click()
'**** Shows Date Selection form for Servicing Details
DateSelect5.Show
End Sub

Private Sub msale_Click()
'**** Shows Date Selection form for Sales Details
DateSelect4.Show
End Sub

Private Sub msd_Click()
'**** Shows Date Selection form for Service Details
DateSelect3.Show
End Sub
