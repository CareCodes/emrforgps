VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Electronic Medical Assistant 3.0"
   ClientHeight    =   7920
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10290
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIMain.frx":058A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Patients"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Queue"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Patient Details"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Pharmacy Queue"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2520
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   1800
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":103FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":1042DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":1045F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":10490E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":1053A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":10BC42
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":1124DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIMain.frx":118D76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBackup 
         Caption         =   "&Backup"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditClinilcalDetails 
         Caption         =   "&Clinilcal Deatails"
         Begin VB.Menu mnuEditSystems 
            Caption         =   "Systems"
         End
         Begin VB.Menu mnuEditSystemOrder 
            Caption         =   "System Order"
         End
         Begin VB.Menu mnuEditSystemFindings 
            Caption         =   "System Findings"
         End
         Begin VB.Menu mnuEditSystemFindingOrder 
            Caption         =   "System Finding Order"
         End
         Begin VB.Menu mnuEditSystemFindingValues 
            Caption         =   "System Finding Values"
         End
         Begin VB.Menu mnuEditDiagnosis 
            Caption         =   "Diagnosis"
         End
         Begin VB.Menu mnuEditDiagnosisInformation 
            Caption         =   "Diagnosis Information"
         End
         Begin VB.Menu mnuEditICD9 
            Caption         =   "ICD9"
         End
         Begin VB.Menu mnuEditICD10 
            Caption         =   "ICD10"
         End
         Begin VB.Menu mnuEditProcedures 
            Caption         =   "Procedures"
         End
      End
      Begin VB.Menu mnuEditMedicines 
         Caption         =   "&Medicines"
         Begin VB.Menu mnuEditGroups 
            Caption         =   "Groups"
         End
         Begin VB.Menu mnuEditSubgroups 
            Caption         =   "Subgroups"
         End
         Begin VB.Menu mnuEditGenericNames 
            Caption         =   "Generic Names"
         End
         Begin VB.Menu mnuEditMedicinesTradeNames 
            Caption         =   "Trade Names"
         End
         Begin VB.Menu mnuEditItemCategories 
            Caption         =   "Item Categories"
         End
         Begin VB.Menu mnuEditItems 
            Caption         =   "Items"
         End
         Begin VB.Menu mnuEditItemSuppliers 
            Caption         =   "Item Suppliers"
         End
         Begin VB.Menu mnuEditBatches 
            Caption         =   "Batches"
         End
         Begin VB.Menu mnuEditMedicineExport 
            Caption         =   "Export"
            Begin VB.Menu mnuEditMedicineExportItems 
               Caption         =   "Items"
            End
         End
         Begin VB.Menu mnuEditMedicineImport 
            Caption         =   "Import"
            Begin VB.Menu mnuEditMedicineImportItems 
               Caption         =   "Item"
            End
         End
      End
      Begin VB.Menu mnuEditInvestigations 
         Caption         =   "Investigations"
         Begin VB.Menu mnuEditInvestigationsInvestigations 
            Caption         =   "Investigations"
         End
         Begin VB.Menu mnuEditInvestigationCategories 
            Caption         =   "Investigation Categories"
         End
         Begin VB.Menu mnuEditIxIxFormat 
            Caption         =   "Investigation Format"
         End
         Begin VB.Menu mnuEditInvestigationValueCategories 
            Caption         =   "Investigation Value Categories"
         End
         Begin VB.Menu mnuEditIxIxValues 
            Caption         =   "Investigation Values"
         End
         Begin VB.Menu mnuEditIxIxValuesForCategories 
            Caption         =   "Investigation Values for Categories"
         End
         Begin VB.Menu mnuIxEditAssignValuesForFields 
            Caption         =   "Assign values for fields"
         End
         Begin VB.Menu mnuEditInvestigationsSpecimans 
            Caption         =   "Specimans"
         End
         Begin VB.Menu mnuEditInvestigationsTubes 
            Caption         =   "Tubes"
         End
      End
      Begin VB.Menu mnuEditUnits 
         Caption         =   "&Units"
         Begin VB.Menu mnuEditItemUnit 
            Caption         =   "Item Unit"
         End
         Begin VB.Menu mnuEditDurationUnits 
            Caption         =   "Duration Units"
         End
         Begin VB.Menu mnuEditFrequencyUnits 
            Caption         =   "Frequency Units"
         End
      End
      Begin VB.Menu mnuEditCategories 
         Caption         =   "&Categories"
         Begin VB.Menu mnuEditPharmacySaleCategories 
            Caption         =   "Pharmacy Sale Categories"
         End
         Begin VB.Menu mnuEditIncomeCategories 
            Caption         =   "Income Categories"
         End
         Begin VB.Menu mnuEditExpenceCategories 
            Caption         =   "Expence Categories"
         End
         Begin VB.Menu mnuEditFoodItems 
            Caption         =   "Food Item"
         End
         Begin VB.Menu mnuEditDepositCategories 
            Caption         =   "Deposit Categories"
         End
         Begin VB.Menu mnuEditWithdrawCategories 
            Caption         =   "Withdraw Categories"
         End
      End
      Begin VB.Menu mnuEditStaff 
         Caption         =   "&Staff"
         Begin VB.Menu mnuEditDoctorSpecialities 
            Caption         =   "Doctor Specialities"
         End
         Begin VB.Menu mnuEditDoctors 
            Caption         =   "Doctors"
         End
         Begin VB.Menu mnuEditOtherStaff 
            Caption         =   "Users"
         End
         Begin VB.Menu mnuEditStaffAllStaff 
            Caption         =   "All Staff Details"
         End
         Begin VB.Menu mnuEditStaffAuthority 
            Caption         =   "Authority && Previlages"
         End
      End
      Begin VB.Menu mnuEditHospital 
         Caption         =   "&Hospital"
         Begin VB.Menu mnuEditHospitalDetails 
            Caption         =   "Hospital Details"
         End
         Begin VB.Menu mnuEditHospitalDepartments 
            Caption         =   "Departments"
         End
         Begin VB.Menu mnuEditHospitalDrawer 
            Caption         =   "Drawers"
         End
         Begin VB.Menu mnuEditPharmacyLocations 
            Caption         =   "Pharmacy Locations"
         End
         Begin VB.Menu mnuEditRooms 
            Caption         =   "Rooms"
         End
      End
      Begin VB.Menu mnuEditInstitutions 
         Caption         =   "&Institutions"
         Begin VB.Menu mnuEditHospitals 
            Caption         =   "Hospitals"
         End
         Begin VB.Menu mnuEditInsurers 
            Caption         =   "Insurers"
         End
         Begin VB.Menu mnuEditImporters 
            Caption         =   "Importers"
         End
         Begin VB.Menu mnuEditManufacturers 
            Caption         =   "Manufacturers"
         End
         Begin VB.Menu mnuEditSuppliers 
            Caption         =   "Suppliers"
         End
         Begin VB.Menu mnuEditCustomers 
            Caption         =   "Customers"
         End
         Begin VB.Menu mnuEditInstitutionsAll 
            Caption         =   "All"
         End
      End
      Begin VB.Menu mnuEditDocuments 
         Caption         =   "Documents"
         Begin VB.Menu mnuEditReferralActions 
            Caption         =   "Referral Actions"
         End
      End
      Begin VB.Menu mnuEditPrinting 
         Caption         =   "Printing"
         Begin VB.Menu mnuEditPrintFormats 
            Caption         =   "Print Formats"
         End
      End
   End
   Begin VB.Menu mnuClinical 
      Caption         =   "&Practice"
      Begin VB.Menu mnuClinicalPatients 
         Caption         =   "Patients"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuClinicalPatientDetails 
         Caption         =   "Patient Details"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuClinicalQueue 
         Caption         =   "queue"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuClinicalTodaysQueue 
         Caption         =   "Queue"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuClinicalFavourites 
         Caption         =   "Favourites"
         Begin VB.Menu mnuClinicalFavouriteItemsByAge 
            Caption         =   "Items by Age"
         End
         Begin VB.Menu mnuClinicalFavouriteItemsByWeight 
            Caption         =   "Items by Weight"
         End
         Begin VB.Menu mnuClinicalFavouriteDiagnosis 
            Caption         =   "Diagnosis"
         End
      End
      Begin VB.Menu mnuClinicalObstetricCalculater 
         Caption         =   "Obstrtic Calculater"
      End
      Begin VB.Menu mnuClinicalManageVisits 
         Caption         =   "Manage Visits"
         Begin VB.Menu mnuClinicalManageVisitsDeleteTodaysVisits 
            Caption         =   "Delete Todays Visits"
         End
      End
   End
   Begin VB.Menu mnuPharmacy 
      Caption         =   "&Pharmacy"
      Begin VB.Menu mnuPharmacySale 
         Caption         =   "Sale"
         Begin VB.Menu mnuPharmacySaleSale 
            Caption         =   "Sale"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuPharmacyQueue 
            Caption         =   "queue"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuPharmacyTodaysQueue 
            Caption         =   "Queue"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnuPharmacyList 
            Caption         =   "list"
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnuPharmacyTodaysList 
            Caption         =   "List"
            Shortcut        =   +^{F4}
         End
      End
      Begin VB.Menu mnuPharmacyPurchase 
         Caption         =   "Purchase"
         Begin VB.Menu mnuPharmacyPurchaseGoodReceive 
            Caption         =   "Good Receive"
         End
         Begin VB.Menu mnuPharmacyPurchaseCancel 
            Caption         =   "Cancel Purchase"
         End
      End
      Begin VB.Menu mnuPharmacyAdjust 
         Caption         =   "Adjust"
         Begin VB.Menu mnuPharmacyAdjustmentSalePrice 
            Caption         =   "Sale Price Adjustment"
         End
         Begin VB.Menu mnuPharmacyAdjustmentPurchasePrice 
            Caption         =   "Purchase Price Adjustment"
         End
         Begin VB.Menu mnuPharmacyAdjustmentStockPrice 
            Caption         =   "Stock Adjustment"
         End
      End
      Begin VB.Menu mnuOpenDrawer 
         Caption         =   "Open Drawer"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuInstitution 
      Caption         =   "&Institution"
      Begin VB.Menu mnuInstitutionIncome 
         Caption         =   "income"
      End
      Begin VB.Menu mnuInstitutionTodaysIncome 
         Caption         =   "Income"
      End
      Begin VB.Menu mnuInstitutionExpence 
         Caption         =   "Expence"
      End
      Begin VB.Menu mnuInstitutionDeposit 
         Caption         =   "Deposit"
      End
      Begin VB.Menu mnuInstitutionWithdraw 
         Caption         =   "Withdraw"
      End
      Begin VB.Menu mnuInstitutionDrawerBalance 
         Caption         =   "Drawer Balance"
      End
      Begin VB.Menu mnuInstitutionResetDrawer 
         Caption         =   "Reset Drawer"
      End
   End
   Begin VB.Menu mnuBackOffice 
      Caption         =   "&Back Office"
      Begin VB.Menu mnuBackOfficePracticeReports 
         Caption         =   "Practice Reports"
         Begin VB.Menu mnuBackOfficePracticeReportsIncome 
            Caption         =   "Income Reports"
            Begin VB.Menu mnuBackOfficePractiseReportsQueueIncome 
               Caption         =   "Queue Income"
            End
            Begin VB.Menu mnuBackOfficePractiseReportsDailyIncome 
               Caption         =   "Daily Income"
            End
            Begin VB.Menu mnuBackOffficePracticeReportsTimeViceIncome 
               Caption         =   "Time-vice Income"
            End
            Begin VB.Menu mnuBackOfficePractiseReportsPeriodIncome 
               Caption         =   "Period Income"
            End
         End
         Begin VB.Menu mnuBackOfficePracticeReportsClinicalReports 
            Caption         =   "Clinical Reports"
            Begin VB.Menu mnuBackOfficeClinicalReportsDiagnosis 
               Caption         =   "Visits by Diagnosis"
            End
         End
      End
      Begin VB.Menu mnuBackOfficePharmacyReports 
         Caption         =   "Pharmacy Reports"
         Begin VB.Menu mnuBackOfficePharmacyReportsSaleReports 
            Caption         =   "Sale Reports"
            Begin VB.Menu mnuBackOfficePharmacyReportsShiftEndSaleReports 
               Caption         =   "Shift End Sale Reports"
            End
            Begin VB.Menu mnuBackOfficeDayEndSaleReports 
               Caption         =   "Day End Sale Reports"
            End
            Begin VB.Menu mnuBackOfficePharmacyReportsSaleCategoryReport 
               Caption         =   "Sale Category Reports"
            End
            Begin VB.Menu mnuBackOfficePharmacyQueueIncome 
               Caption         =   "Queue Income"
            End
         End
         Begin VB.Menu mnuBackOfficePharmacyReportsPurchaseReports 
            Caption         =   "Purchase Reports"
            Begin VB.Menu mnuBackOfficePharmacyReportsPurchaseBills 
               Caption         =   "Purchase Bills"
            End
         End
         Begin VB.Menu mnuBackOfficePharmacyReportsStockReports 
            Caption         =   "Stock Reports"
            Begin VB.Menu mnuBackOfficePharmacyReportsItemStock 
               Caption         =   "Item Stock"
            End
            Begin VB.Menu mnuBackOfficePharmacyReportsBatchStock 
               Caption         =   "Batch Stocks"
            End
            Begin VB.Menu mnuBackOfficePharmacyReportsStocksBelowROL 
               Caption         =   "Stocks Below ROL"
            End
            Begin VB.Menu mnuBackOfficePharmacyReportsExpiringStocks 
               Caption         =   "Expiring Stocks"
            End
            Begin VB.Menu mnuBackOfficePharmacyExpiredStock 
               Caption         =   "Expired Stock"
            End
         End
         Begin VB.Menu mnuBackOfficePharmacyReportsItemReports 
            Caption         =   "Item Reports"
            Begin VB.Menu mnuBackOfficePharmacyReportsItemSale 
               Caption         =   "Item Sale"
            End
            Begin VB.Menu mnuBackOfficePharmacyReportsItemPurchase 
               Caption         =   "Item Purchase"
            End
         End
      End
      Begin VB.Menu mnuBackOfficeInstitutionReports 
         Caption         =   "Institution Reports"
         Begin VB.Menu mnuBackOfficeInstitutionReportsDeposits 
            Caption         =   "Deposits"
         End
         Begin VB.Menu mnuBackOfficeInstitutionReportsWithdraws 
            Caption         =   "Withdraws"
         End
      End
      Begin VB.Menu mnuBackOfficeMessage 
         Caption         =   "Message"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPracticePreferances 
         Caption         =   "Practice Preferances"
      End
      Begin VB.Menu mnuProgramPPreferances 
         Caption         =   "Program Preferrances"
      End
      Begin VB.Menu mnuOptionsDatabase 
         Caption         =   "Database"
      End
      Begin VB.Menu mnuOptionsPrinting 
         Caption         =   "Printing"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On Web"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About EMA 3"
      End
      Begin VB.Menu Menu1 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   ""
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

    Dim FSys As New Scripting.FileSystemObject

Private Sub mnuWindowArrangeIcons_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
  MsgBox "New WIndow Code goes here!"
End Sub

Private Sub mnuWindowTileHorizontal_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  Me.Arrange vbTileVertical
End Sub

Private Sub mnuFileClose_Click()
  MsgBox "Close Code goes here!"
End Sub

Private Sub mnuBackOfficeClinicalReportsDiagnosis_Click()
    frmPeriodDiagnosis.Show
    frmPeriodDiagnosis.ZOrder 0
End Sub

Private Sub mnuBackOfficeInstitutionReportsDeposits_Click()
    frmDepositReport.Show
    frmDepositReport.ZOrder 0
End Sub

Private Sub mnuBackOfficeInstitutionReportsWithdraws_Click()
    frmWithdrawReport.Show
    frmWithdrawReport.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyExpiredStock_Click()
    frmItemBatchStockExpired.Show
    frmItemBatchStockExpired.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsBatchStock_Click()
    frmItemBatchStock.Show
    frmItemBatchStock.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsExpiringStocks_Click()
    frmItemBatchStockExpiring.Show
    frmItemBatchStockExpiring.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsItemPurchase_Click()
    frmPharmacyDailyItemPurchase.Show
    frmPharmacyDailyItemPurchase.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsItemSale_Click()
    frmPharmacyDailyItemIssue.Show
    frmPharmacyDailyItemIssue.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsPurchaseBills_Click()
    frmPurchaseBills.Show
    frmPurchaseBills.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsStocksBelowROL_Click()
    frmItemStockBelowROL.Show
    frmItemStockBelowROL.ZOrder 0
End Sub

Private Sub mnuClinicalManageVisitsDeleteTodaysVisits_Click()
    frmDeleteTodaysVisits.Show
    frmDeleteTodaysVisits.ZOrder 0
End Sub

Private Sub mnuClinicalTodaysQueue_Click()
    frmTodaysQueue.Show
    frmTodaysQueue.ZOrder 0
End Sub

Private Sub mnuEditDepositCategories_Click()
    frmDepositCategory.Show
    frmDepositCategory.ZOrder 0
End Sub

Private Sub mnuEditExpenceCategories_Click()
    frmExpenceCategory.Show
    frmExpenceCategory.ZOrder 0
End Sub

Private Sub mnuEditHealthSchemeProviders_Click()
    
End Sub

Private Sub mnuEditHospitalDrawer_Click()
    frmDrawer.Show
    frmDrawer.ZOrder 0
End Sub

Private Sub mnuEditHospitals_Click()
    frmHospitals.Show
    frmHospitals.ZOrder 0
End Sub

Private Sub mnuEditInstitutionsAll_Click()
    frmInstitutions.Show
    frmInstitutions.ZOrder 0
End Sub

Private Sub mnuEditInsurers_Click()
    frmInsurers.Show
    frmInsurers.ZOrder 0
End Sub

Private Sub mnuEditWithdrawCategories_Click()
    frmWithdrawCategory.Show
    frmWithdrawCategory.ZOrder 0
End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub mnuFileNew_Click()
  MsgBox "New File Code goes here!"
End Sub

Private Sub mnuFileOpen_Click()
  MsgBox "Open Code goes here!"
End Sub

Private Sub mnuFilePrint_Click()
  MsgBox "Print Code goes here!"
End Sub

Private Sub mnuFilePrintPreview_Click()
  MsgBox "Print Preview Code goes here!"
End Sub

Private Sub mnuFilePrintSetup_Click()
  MsgBox "Print Setup Code goes here!"
End Sub

Private Sub mnuFileProperties_Click()
  MsgBox "Properties Code goes here!"
End Sub

Private Sub mnuFileSave_Click()
  MsgBox "Save File Code goes here!"
End Sub

Private Sub mnuFileSaveAll_Click()
  MsgBox "Save All Code goes here!"
End Sub

Private Sub mnuFileSaveAs_Click()
  MsgBox "Save As Code goes here!"
End Sub

Private Sub mnuFileSend_Click()
  MsgBox "Send Code goes here!"
End Sub

Private Sub mnuEditCopy_Click()
  MsgBox "Place Copy Code here!"
End Sub

Private Sub mnuEditCut_Click()
  MsgBox "Place Cut Code here!"
End Sub

Private Sub mnuEditDSelectAll_Click()
  MsgBox "Place Select All Code here!"
End Sub

Private Sub mnuEditInvertSelection_Click()
  MsgBox "Place Invert Selection Code here!"
End Sub

Private Sub mnuEditPaste_Click()
  MsgBox "Place Paste Code here!"
End Sub

Private Sub mnuEditPasteSpecial_Click()
  MsgBox "Place Paste Special Code here!"
End Sub

Private Sub mnuEditUndo_Click()
  MsgBox "Place Undo Code here!"
End Sub


Private Sub mnuFileRestore_Click()
    frmRestore.Show
    frmRestore.ZOrder 0
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuHelpAdministrator_Click()
    frmTem1.Show
    
End Sub

Private Sub mnuHelpContents_Click()
    frmTemHelp.Show
    DoEvents
    SendKeys "{F1}"
    DoEvents
    Unload frmTemHelp
End Sub

Private Sub MDIForm_Load()
''    EnableControls Me
   VisibleControls Me
    TotalMessageCount = GetSetting(App.EXEName, UserID, "TotalMessageCount", 0)
    DoEvents
    Screen.MousePointer = vbDefault
    dlgCommonDialog.HelpFile = App.HelpFile
    
    
    
    If FSys.FileExists(MDIImageFile) = True Then
        On Error Resume Next
        Set Me.Picture = LoadPicture(MDIImageFile, vbLPCustom, , Me.Width, Me.height)
    End If

End Sub

Private Sub MenuEnable()
    On Error Resume Next
    Dim MenuNameID As Long
    Dim rsTem As New ADODB.Recordset
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If Left(MyCtrl.Name, 3) = "mnu" Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblMenuName where MenuName = '" & MyCtrl.Name & "'"
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    MenuNameID = !MenuNameID
                    If .State = 1 Then .Close
                    temSql = "Select * from tblstaffauthority where StaffID = " & UserID & " AND MenuNameID = " & MenuNameID
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    If .RecordCount > 0 Then
                        MyCtrl.Enabled = !Enabled
                    End If
                    .Close
                Else
                    MyCtrl.Enabled = False
                End If
            End With
        End If
    Next

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    i = MsgBox("Are you sure you want to exit?", vbYesNo, "EXIT?")
    If i = vbNo Then Cancel = True
    SaveSetting App.EXEName, UserID, "TotalMessageCount", TotalMessageCount
End Sub

Private Sub mnuBackOffficePracticeReportsTimeViceIncome_Click()
    frmTimeviceIncome.Show
    frmTimeviceIncome.ZOrder 0
End Sub

Private Sub mnuBackOfficeDayEndSaleReports_Click()
    frmPharmacyDailyIncome.Show
    frmPharmacyDailyIncome.ZOrder 0
End Sub

Private Sub mnuBackOfficeIncomeReports_Click()
    frmTem1.Show
    frmTem1.ZOrder 0
End Sub

Private Sub mnuBackOfficeMessage_Click()
    frmMessages.Show
    frmMessages.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsItemBatchStock_Click()
    frmItemBatchStock.Show
    frmItemBatchStock.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyQueueIncome_Click()
    frmPharmacyQueueIncome.Show
    frmPharmacyQueueIncome.ZOrder 0
End Sub

Private Sub mnuBackOfficePharmacyReportsItemStock_Click()
    frmItemStock.Show
    frmItemStock.ZOrder 0
End Sub

Private Sub mnuBackOfficePractiseReportsDailyIncome_Click()
    frmDailyIncome.Show
    frmDailyIncome.ZOrder 0
End Sub

Private Sub mnuBackOfficePractiseReportsPeriodIncome_Click()
    frmPeriodIncome.Show
    frmPeriodIncome.ZOrder 0
End Sub

Private Sub mnuBackOfficePractiseReportsQueueIncome_Click()
    frmQueueIncome.Show
    frmQueueIncome.ZOrder 0
End Sub

Private Sub mnuClinicalFavouriteDiagnosis_Click()
    frmFavouriteDiagnosis.Show
    frmFavouriteDiagnosis.ZOrder 0
End Sub

Private Sub mnuClinicalFavouriteItemsByAge_Click()
    frmFavouriteItemByAge.Show
    frmFavouriteItemByAge.ZOrder 0
End Sub

Private Sub mnuClinicalFavouriteItemsByWeight_Click()
    frmFavouriteItemByWeight.Show
    frmFavouriteItemByWeight.ZOrder 0
End Sub

Private Sub mnuClinicalObstetricCalculater_Click()
    frmObstetricCalculator.Show
    frmObstetricCalculator.ZOrder 0
End Sub

Private Sub mnuClinicalPatientDetails_Click()
    frmPatientDetails.Show
    frmPatientDetails.ZOrder 0
End Sub

Private Sub mnuClinicalPatients_Click()
    Dim MyForm As Form
    Set MyForm = frmPatient
    MyForm.Show
    MyForm.ZOrder 0

'    frmPatient.Show
'    frmPatient.ZOrder 0
End Sub

Private Sub mnuClinicalQueue_Click()
    frmQueue.Show
    frmQueue.ZOrder 0
End Sub

Private Sub mnuClinicalSearchPatients_Click()
    frmSearchPatient.Show
    frmSearchPatient.ZOrder 0
End Sub

Private Sub mnuEditCustomers_Click()
    frmCustomer.Show
    frmCustomer.ZOrder 0
End Sub

Private Sub mnuEditDiagnosis_Click()
    frmDiagnosis.Show
    frmDiagnosis.ZOrder 0
End Sub

Private Sub mnuEditDiagnosisInformation_Click()
    frmDiseaseInformation.Show
    frmDiseaseInformation.ZOrder 0
End Sub

Private Sub mnuEditDoctors_Click()
    frmDoctors.Show
    frmDoctors.ZOrder 0
End Sub

Private Sub mnuEditDoctorSpecialities_Click()
    frmDoctorSpeciality.Show
    frmDoctorSpeciality.ZOrder 0
End Sub

Private Sub mnuEditDurationUnits_Click()
    frmDurationUnit.Show
    frmDurationUnit.ZOrder 0
End Sub

Private Sub mnuEditFoodItems_Click()
    frmFoodItem.Show
    frmFoodItem.ZOrder 0
End Sub

Private Sub mnuEditFrequencyUnits_Click()
    frmFrequencyUnit.Show
    frmFrequencyUnit.ZOrder 0
End Sub

Private Sub mnuEditGenericNames_Click()
    frmGenericName.Show
    frmGenericName.ZOrder 0
End Sub

Private Sub mnuEditGroups_Click()
    frmMedicineGroup.Show
    frmMedicineGroup.ZOrder 0
End Sub

Private Sub mnuEditHospitalDepartments_Click()
    frmDepartment.Show
    frmDepartment.ZOrder 0
End Sub

Private Sub mnuEditICD10_Click()
    frmICD10.Show
    frmICD10.ZOrder 0
End Sub

Private Sub mnuEditICD9_Click()
    frmICD9.Show
    frmICD9.ZOrder 0
End Sub

Private Sub mnuEditImporters_Click()
    frmImporters.Show
    frmImporters.ZOrder 0
End Sub

Private Sub mnuEditIncomeCategories_Click()
    frmIncomeCategory.Show
    frmIncomeCategory.ZOrder 0
End Sub

Private Sub mnuEditInvestigationCategories_Click()
    frmIxCategory.Show
    frmIxCategory.ZOrder 0
End Sub

Private Sub mnuEditInvestigationsInvestigations_Click()
    frmIx.Show
    frmIx.ZOrder 0
End Sub

Private Sub mnuEditInvestigationsSpecimans_Click()
    frmSpecimanCategory.Show
    frmSpecimanCategory.ZOrder 0
End Sub

Private Sub mnuEditInvestigationsTubes_Click()
    frmTubeCategory.Show
    frmTubeCategory.ZOrder 0
End Sub

Private Sub mnuEditInvestigationValueCategories_Click()
    frmIxValueCategory.Show
    frmIxValueCategory.ZOrder 0
End Sub

Private Sub mnuEditIssueUnits_Click()
    frmIssueUnit.Show
    frmIssueUnit.ZOrder 0
End Sub

Private Sub mnuEditItemCategories_Click()
    frmItemCategory.Show
    frmItemCategory.ZOrder 0
End Sub

Private Sub mnuEditItems_Click()
    frmItem.Show
    frmItem.ZOrder 0
End Sub

Private Sub mnuEditItemUnit_Click()
    frmItemUnit.Show 0
    frmItemUnit.ZOrder 0
End Sub

Private Sub mnuEditIxIxFormat_Click()
'    frmNewIxFormat.Show
'    frmNewIxFormat.ZOrder 0
    frmIxFormat.Show
    frmIxFormat.ZOrder 0
End Sub

Private Sub mnuEditIxIxValues_Click()
    frmIxValues.Show
    frmIxValues.ZOrder 0
End Sub

Private Sub mnuEditIxIxValuesForCategories_Click()
    frmAssignValuesForCategories.Show
    frmAssignValuesForCategories.ZOrder 0
End Sub

Private Sub mnuEditManufacturers_Click()
    frmManufacturer.Show
    frmManufacturer.ZOrder 0
End Sub

Private Sub mnuEditMedicineExportItems_Click()
    frmExportItems.Show
    frmExportItems.ZOrder 0
End Sub

Private Sub mnuEditMedicineImportItems_Click()
    frmImportItems.Show
    frmImportItems.ZOrder 0
End Sub

Private Sub mnuEditMedicinesTradeNames_Click()
    frmTradeName.Show
    frmTradeName.ZOrder 0
End Sub

Private Sub mnuEditObsFindings_Click()
    frmObsFindings.Show
    frmObsFindings.ZOrder 0
End Sub

Private Sub mnuEditObstetricFindingOrder_Click()
'    frmObsFindingOrder.Show
'    frmObsFindingOrder.ZOrder 0
End Sub

Private Sub mnuEditObstetricFindingValues_Click()
    frmObsFindingValues.Show
    frmObsFindingValues.ZOrder 0
End Sub

Private Sub mnuEditObstetricHistory_Click()
    frmObsHistory.Show
    frmObsHistory.ZOrder 0
End Sub

Private Sub mnuEditObstetricHistoryValues_Click()
    frmObsHistoryValues.Show
    frmObsHistoryValues.ZOrder 0
End Sub

Private Sub mnuEditObstetricOrder_Click()
'    frmObsHistoryOrder.Show
'    frmObsHistoryOrder.ZOrder 0
End Sub

Private Sub mnuEditOtherStaff_Click()
    frmStaff.Show
    frmStaff.ZOrder
End Sub

Private Sub mnuEditPackUnits_Click()
    frmPackUnit.Show
    frmPackUnit.ZOrder 0
End Sub

Private Sub mnuEditPharmacySaleCategories_Click()
    frmPharmacySaleCatogeries.Show
    frmPharmacySaleCatogeries.ZOrder 0
End Sub

Private Sub mnuEditPrintFormats_Click()
    frmNewPrintFormat.Show
    frmNewPrintFormat.ZOrder 0
End Sub

Private Sub mnuEditProcedures_Click()
    frmProcedure.Show
    frmProcedure.ZOrder 0
End Sub

Private Sub mnuEditReferralActions_Click()
    frmActionExpected.Show
    frmActionExpected.ZOrder 0
End Sub

Private Sub mnuEditStaff_Click()
'    frmStaffDetails.Show
 '   frmStaffDetails.ZOrder 0
End Sub

Private Sub mnuEditStaffAllStaff_Click()
    frmAllStaff.Show
    frmAllStaff.ZOrder 0
End Sub

Private Sub mnuEditStrengthUnits_Click()
    frmStrengthUnit.Show
    frmStrengthUnit.ZOrder 0
End Sub

Private Sub mnuEditStaffAuthority_Click()
    frmAuthorityPrevilagesMenuVisible.Show
    frmAuthorityPrevilagesMenuVisible.ZOrder 0
End Sub

Private Sub mnuEditSubgroups_Click()
    frmMedicineSubgroup.Show
    frmMedicineSubgroup.ZOrder 0
End Sub

Private Sub mnuEditSuppliers_Click()
    frmSupplier.Show
    frmSupplier.ZOrder 0
End Sub

Private Sub mnuEditSystemFindingOrder_Click()
    frmSystemFindingOrder.Show
    frmSystemFindingOrder.ZOrder 0
End Sub

Private Sub mnuEditSystemFindings_Click()
    frmSystemFindings.Show
    frmSystemFindings.ZOrder 0
End Sub

Private Sub mnuEditSystemFindingValues_Click()
    frmSystemFindingValues.Show
    frmSystemFindingValues.ZOrder 0
End Sub

Private Sub mnuEditSystemOrder_Click()
    frmSystemOrder.Show
    frmSystemOrder.ZOrder 0
End Sub

Private Sub mnuEditSystems_Click()
    frmSystem.Show
    frmSystem.ZOrder 0
End Sub

Private Sub mnuFileBackup_Click()
    frmBackUp.Show
    frmBackUp.ZOrder 0
End Sub

Private Sub mnuHelpSearch_Click()
    Dim r As Long
    r = ShellExecute(0, "open", "www.lakmedi.com", 0, 0, 1)

End Sub

Private Sub mnuInstitutionDeposit_Click()
    frmDeposit.Show
    frmDeposit.ZOrder 0
End Sub

Private Sub mnuInstitutionDrawerBalance_Click()
    frmDrawerBalance.Show
    frmDrawerBalance.ZOrder 0
End Sub

Private Sub mnuInstitutionExpence_Click()
    frmExpence.Show
    frmExpence.ZOrder 0
End Sub

Private Sub mnuInstitutionIncome_Click()
    frmIncome.Show
    frmIncome.ZOrder 0
End Sub

Private Sub mnuInstitutionResetDrawer_Click()
    frmResetDrawer.Show
    frmResetDrawer.ZOrder 0
End Sub

Private Sub mnuInstitutionTodaysIncome_Click()
    frmTodaysIncome.Show
    frmTodaysIncome.ZOrder 0
End Sub

Private Sub mnuInstitutionWithdraw_Click()
    frmWithdraw.Show
    frmWithdraw.ZOrder 0
End Sub

Private Sub mnuIxEditAssignValuesForFields_Click()
    frmAssignIxValuesToFields.Show
    frmAssignIxValuesToFields.ZOrder 0
End Sub

Private Sub mnuLab_Click()
'    Dim rsTem As New ADODB.Recordset
'    Dim MyForm As Control
'    With rsTem
'        For Each MyForm In Controls
'            If Left(MyForm.Name, 3) = "mnu" Then
'                If .State = 1 Then .Close
'                temSql = "Select * from tblMenuName where MenuName = '" & MyForm.Name & "'"
'                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                If .RecordCount > 0 Then
'
'                Else
'                    .AddNew
'                    !MenuName = MyForm.Name
'                End If
'                !MenuCaption = MyForm.Caption
'                .Update
'            End If
'        Next
'        .Close
'    End With
End Sub


Private Sub mnuOpenDrawer_Click()
      Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim CsetPrinter As New cSetDfltPrinter
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    Printer.Font.Name = "Control"
    Printer.Print "a"
    Printer.EndDoc
End Sub

Private Sub mnuOptionsDatabase_Click()
    frmSelectDatabase.Show
    frmSelectDatabase.ZOrder 0
End Sub

Private Sub mnuOptionsPrinting_Click()
    frmPrintingPreferances.Show
    frmPrintingPreferances.ZOrder 0
End Sub

Private Sub mnuPharmacyAdjustmentPurchasePrice_Click()
    frmPPAdjustment.Show
    frmPPAdjustment.ZOrder 0
End Sub

Private Sub mnuPharmacyAdjustmentSalePrice_Click()
    frmSPAdjustment.Show
    frmSPAdjustment.ZOrder
End Sub

Private Sub mnuPharmacyAdjustmentStockPrice_Click()
    frmStockAdjustment.Show
    frmStockAdjustment.ZOrder 0
End Sub

Private Sub mnuPharmacyList_Click()
    frmPharmacyList.Show
    frmPharmacyList.ZOrder 0
End Sub

Private Sub mnuPharmacyPurchaseCancel_Click()
    frmSearchToCancelPurchaseBills.Show
    frmSearchToCancelPurchaseBills.ZOrder 0
End Sub

Private Sub mnuPharmacyPurchaseGoodReceive_Click()
    frmItemPurchase.Show
    frmItemPurchase.ZOrder 0
End Sub

Private Sub mnuPharmacyQueue_Click()
    frmPharmacyQueue.Show
    frmPharmacyQueue.ZOrder 0
End Sub

Private Sub mnuPharmacySaleSale_Click()
    frmPharmacySale.Show
    frmPharmacySale.ZOrder
End Sub

Private Sub mnuPharmacyTodaysList_Click()
    frmTodaysPharmacyList.Show
    frmTodaysPharmacyList.ZOrder 0
End Sub

Private Sub mnuPharmacyTodaysQueue_Click()
    frmTodaysPharmacyQueue.Show
    frmTodaysPharmacyQueue.ZOrder 0
End Sub

Private Sub mnuPracticePreferances_Click()
    frmPracticePreferances.Show
    frmPracticePreferances.ZOrder 0
End Sub

Private Sub mnuProgramPPreferances_Click()
    frmProgramPreferances.Show
    frmProgramPreferances.ZOrder 0
End Sub

'Private Sub mnuRxPopUpItem_Click(Index As Integer)
'    If SingleVisit = True Then
'        frmNewVisit.RxPopAction (Index)
'    Else
'        Dim MyfrmNewVisit As frmNewVisit
'        For Each MyfrmNewVisit In frmNewVisit
'            If MyfrmNewVisit.Name = "frmNewVisit" Then
'                If MDIMain.ActiveForm.hwnd = MDIMain.ActiveForm.hwnd Then
'                    MyfrmNewVisit.RxPopAction (Index)
'                End If
'            End If
'        Next
'    End If
'End Sub

Private Sub mnuTem_Click()
    frmTem.Show
End Sub

Private Sub Timer1_Timer()
    Dim rsMessage As New ADODB.Recordset
    With rsMessage
        temSql = "SELECT tblFromStaff.Staff, tblToStaff.Staff, tblMessage.* FROM tblStaff AS tblFromStaff RIGHT JOIN (tblStaff AS tblToStaff RIGHT JOIN tblMessage ON tblToStaff.StaffID = tblMessage.ToID) ON tblFromStaff.StaffID = tblMessage.FromID where Cleared = False and (ToID = " & UserID & " OR FromID = " & UserID & ") Order by MessageID"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If TotalMessageCount <> .RecordCount And TotalMessageCount <> 0 Then
            TotalMessageCount = .RecordCount
            frmMessages.Show
            frmMessages.ZOrder 0
        Else
        
        End If
        .Close
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1: If mnuClinicalPatients.Visible = True Then mnuClinicalPatients_Click
        Case 2: If mnuClinicalTodaysQueue.Visible = True Then mnuClinicalTodaysQueue_Click
        Case 3: If mnuClinicalPatientDetails.Visible = True Then mnuClinicalPatientDetails_Click
        Case 4: If mnuPharmacyTodaysQueue.Visible = True Then mnuPharmacyTodaysQueue_Click
    End Select
End Sub
