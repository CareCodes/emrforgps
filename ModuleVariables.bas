Attribute VB_Name = "ModuleVariables"
Option Explicit

' API Variables
    Public Declare Function GetActiveWindow Lib "user32" () As Long

'Clinical Variables
    Public CurrentPatient As New clsPatient
    Public CurrentVisit As New clsVisit

'Programing Variables
Public Type ReportFont
    Name As String
    Size As Long
    Bold As Boolean
    Italic As Boolean
    Strikethrough As Boolean
    Underline As Boolean
End Type

Public Enum ControlType
    Textbox = 1
    ComboBox = 2
    ListBox = 3
    DataCombo = 4
    DataList = 5
    Grid = 6
    Button = 7
    MenuItem = 8
    CheckBox = 9
    OptionButton = 10
    DateTimePicker = 11
    Label = 12
    SSTab = 13
    Unknown = 100
End Enum

' Pharmacy Variables
    Public Type ItemStock
        Stock As Double
        MaxDOE As Date
        MinDOE As Date
    End Type
    
    Public Type ItemBatch
        ItemID As Long
        BatchID As Long
        DOM As Date
        DOE As Date
        Exists As Boolean
        ItemBatch As String
        Stock As Double
    End Type
    
    Public Enum ExpenceMethod
        PharmacyPurchase = 1
        OtherExpence = 2
    End Enum
    
    Public Enum IncomeMethod
        PharmacyIncome = 1
        PracticeIncome = 2
        LaboratoryIncome = 3
        InwardIncome = 4
        OtherIncome = 5
    End Enum
    
    Public Enum UserAction
        StockReset = 1
        DrawerReset = 2
        SalePriceAdjustment = 3
        PurchasePriceAdjustment = 4
        
    End Enum
    
    
    Public DefaultFont As ReportFont
    
    
'Practice Preferances
    Public UsualConsultationFee As Double
    Public UsualConsultationCost As Double
    
    Public AutoIssueIndoor As Boolean
    
    Public AddMedicineChargesToFee As Boolean
    Public AddIndoorMedicineChargesToFee As Boolean


'Program Preferances
    Public AutoSecession As Boolean
    Public SingleVisit As Boolean
    
    Public MDIImageFile As String
    
    Public Type ColourScheme
        LabelForeColour As Long
        LabelBackColour As Long
        ButtonForeColour As Long
        ButtonBackColour As Long
        ButtonBorderColour As Long
        GridLightBackColour As Long
        GridDarkBackColour As Long
        ComboBackColour As Long
        ComboForeColour As Long
        TextBackColour As Long
        TextForeColour As Long
    End Type
    
    Public Energy As ColourScheme
    Public Aqua As ColourScheme
    Public Sunny As ColourScheme

    Public DefaultColourScheme As ColourScheme

' database Variables
    Public Database As String
    Public cnnStores As New ADODB.Connection
    Public cnnImport As New ADODB.Recordset
    Public ImportDatabase As String

' User Variables
    Public currentDoctorId As Long
    Public UserName As String
    Public UserID As Long
    Public UserAuthority As Long
    Public UserDepartmentID As Long
    Public UserDepartment As String
    Public NewPatientMessageUserID As Long
    Public OutdoorDepartmentID As Long
    Public IndoorDepartmentID As Long
    
    Public InstitutionName As String
    Public InstitutionAddress As String
    Public InstitutionTelephone As String
    
    Public TotalMessageCount As Long

' Printing Preferances
    Public PrintingOnBlankPaper As Boolean
    Public PrintingOnPrintedPaper As Boolean
    
    Public PrescreptionPrinterName As String
    Public PrescreptionPaperName As String
    Public PrescreptionPaperHeight As Long
    Public PrescreptionPaperWidth As Long
    
    Public ReportPrinterName As String
    Public ReportPaperName As String
    Public ReportPaperWidth As Long
    Public ReportPaperHeight As Long
    
    Public DocumentPrinterName As String
    Public DocumentPaperName As String
    Public DocumentPaperWidth As Long
    Public DocumentPaperHeight As Long
    
    
    Public LongDateFormat As String
    Public ShortDateFormat As String
    
    Public PrintOK As Boolean
    
    Public printPOS As Boolean
    
' Program Variable
    Public DemoCopy As Boolean
    Public ExpiaryDate As Date
    Public DemoCount As Long
    Public ExpiaryCount As Long
    Public ExpiaryDays As Long
    Public WillExpire As Boolean
    Public HiddenDemo As Boolean
    Public PCActivated As Boolean
    Public DatabaseActivated As Boolean
    Public UsedDays As Long
    Public TodayStartedDate As Date
    
' Data transfer variables
    Public StaffCommentIDTx As Long
    Public GraphPathTx As String


'Authorizing Variables
   
    Public Enum Authority
        Viewer = 1
        Analyzer = 2
        OrdinaryUser = 3
        PowerUser = 4
        SuperUser = 5
        Administrator = 6
        NotIdentified = 7
    End Enum
    
    Public UserAuthorityLevel As Authority

' Graphic Variables
    Public Type DrawingPoint
        X As Long
        Y As Long
    End Type
    Public Type DrawingRectangle
        X1 As Long
        Y1 As Long
        X2 As Long
        Y2 As Long
        Width As Long
        height As Long
        CantrePoint As DrawingPoint
    End Type
    


' Family History Variables

    Public Enum FamilyRelationShip
        Aunt
        Brother
        Brother_in_law
        Child
        Daughter
        Daughter_in_law
        Father
        Father_in_law
        Grandaunt
        Grandfather
        Grandmother
        Grandnephew
        Grandniece
        Grandparent
        Grandson
        Granddaughter
        Grandchild
        Granduncle
        Great_grandfather
        Great_grandmother
        Half_Brother
        Half_Sister
        Husband
        Mother
        Mother_in_law
        MySelf
        Nephew
        Niece
        Parent
        Sibling
        Sister
        Sister_in_law
        Son
        Son_in_law
        Spouse
        Step_Brother
        Step_Daughter
        Step_Father
        Step_Mother
        Step_Relative
        Step_Sister
        Step_Son
        Uncle
        Wife
        Other
    End Enum
    
    ' Arithmatic Variables
    
    Public mmtoinches As Double
