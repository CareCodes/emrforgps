VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim rsItem As New ADODB.Recordset
    Dim temSql As String
        
    Private IsGroupValue As Boolean
    Private IsSubGroupValue As Boolean
    Private IsGenericNameValue As Boolean
    Private IsTradeNameValue As Boolean
    Private IsItemNameValue As Boolean
    Private IDValue As Long
    Private ItemValue As String
    Private CodeValue As String
    Private CategoryValue As String
    Private CategoryIDValue As Long
    Private GroupValue As String
    Private GroupIDValue As Long
    Private SubGroupValue As String
    Private SubGroupIDValue As Long
    Private GenericNameValue As String
    Private GenericNameIDValue As Long
    Private TradeNameValue As String
    Private TradeNameIDValue As Long
    Private ManufacturerValue As String
    Private ManufacturerIDValue As Long
    Private ImporterValue As String
    Private ImporterIDValue As Long
    Private SupplierValue As String
    Private SupplierIDValue As Long
    Private StrengthUnitValue As String
    Private StrengthUnitIDValue As Long
    Private IssueUnitValue As String
    Private IssueUnitIDValue As Long
    Private PackUnitValue As String
    Private PackUnitIDValue As Long
    Private DoseUnitIDValue As Long
    Private DoseUnitValue As String
    Private StrengthUnitsPerIssueUnitValue As Double
    Private IssueUnitsPerPackValue As Double
    Private ROLValue As Double
    Private ROQValue As Double
    Private MinQtyValue As Double
    Private MinIQtyValue As Double
    Private MaxIQtyValue As Double
    Private DivQtyValue As Double
    Private ProfitMargine As Double
    Private PurchasePriceValue As Double
    Private SalePriceValue As Double
    Private LastPurchasePriceValue As Double
    Private LastSalePriceValue As Double
    Private CommentsValue As String

Public Property Let ID(ID As Long)
    IDValue = ID
    
    
     IsGroupValue = False
     IsSubGroupValue = False
     IsGenericNameValue = False
     IsTradeNameValue = False
     IsItemNameValue = False
     ItemValue = ""
     CodeValue = ""
     CategoryValue = ""
     CategoryIDValue = 0
     GroupValue = ""
     GroupIDValue = 0
     SubGroupValue = ""
     SubGroupIDValue = 0
     GenericNameValue = ""
     GenericNameIDValue = 0
     TradeNameValue = ""
     TradeNameIDValue = 0
     ManufacturerValue = ""
     ManufacturerIDValue = 0
     ImporterValue = ""
     ImporterIDValue = 0
     SupplierValue = ""
     SupplierIDValue = 0
     StrengthUnitValue = ""
     StrengthUnitIDValue = 0
     IssueUnitValue = ""
     IssueUnitIDValue = 0
     PackUnitValue = ""
     PackUnitIDValue = 0
     DoseUnitIDValue = 0
     DoseUnitValue = ""
     StrengthUnitsPerIssueUnitValue = 0
     IssueUnitsPerPackValue = 0
     ROLValue = 0
     ROQValue = 0
     MinQtyValue = 0
     MinIQtyValue = 0
     MaxIQtyValue = 0
     DivQtyValue = 0
     ProfitMargine = 0
     PurchasePriceValue = 0
     SalePriceValue = 0
    CommentsValue = Empty
    
    
    With rsItem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.*, tblGroup.ItemID, tblSubGroup.ItemID, tblGenericName.ItemID, tblTradeName.ItemID, tblGroup.Item, tblSubGroup.Item, tblGenericName.Item, tblTradeName.Item, tblCategory.Category, tblCategory.SalesMargin, tblCategory.CategoryID, tblStrengthUnit.ItemUnit, tblPackUnit.ItemUnit, tblIssueUnit.ItemUnit, tblDoseUnit.ItemUnit, tblManufacturer.Institution, tblImporter.Institution, tblSupplier.Institution,  tblStrengthUnit.ItemUnitID, tblPackUnit.ItemUnitID, tblIssueUnit.ItemUnitID, tblDoseUnit.ItemUnitID, tblManufacturer.InstitutionID, tblImporter.InstitutionID, tblSupplier.InstitutionID " & _
                    "FROM (((((((((((tblItem LEFT JOIN tblItem AS tblGroup ON tblItem.GroupID = tblGroup.ItemID) LEFT JOIN tblItem AS tblTradeName ON tblItem.TradeNameID = tblTradeName.ItemID) LEFT JOIN tblItem AS tblSubGroup ON tblItem.SubGroupID = tblSubGroup.ItemID) LEFT JOIN tblItem AS tblGenericName ON tblItem.GenericNameID = tblGenericName.ItemID) LEFT JOIN tblCategory ON tblItem.ItemCategoryID = tblCategory.CategoryID) LEFT JOIN tblItemUnit AS tblStrengthUnit ON tblItem.StrengthUnitID = tblStrengthUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblPackUnit ON " & _
                    "tblItem.PackUnitID = tblPackUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDoseUnit ON tblItem.DoseUnitID = tblDoseUnit.ItemUnitID) LEFT JOIN tblInstitution AS tblManufacturer ON tblItem.ManufacturerID = tblManufacturer.InstitutionID) LEFT JOIN tblInstitution AS tblImporter ON tblItem.ImporterID = tblImporter.InstitutionID) LEFT JOIN tblInstitution AS tblSupplier ON tblItem.SupplierID = tblSupplier.InstitutionID " & _
                    " WHERE (((tblItem.ItemID)=" & IDValue & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNumeric(![tblItem.Item]) Then ItemValue = ![tblItem.Item]
            If Not IsNull(!ItemCode) Then CodeValue = .Fields("Itemcode").Value
            If Not IsNull(!Category) Then CategoryValue = .Fields("Category").Value
            
            If IsNull(![Category]) = False Then CategoryValue = ![Category]
            If IsNull(![tblGroup.Item]) = False Then GroupValue = ![tblGroup.Item]
            If IsNull(![tblSubGroup.Item]) = False Then SubGroupValue = ![tblSubGroup.Item]
            If IsNull(![tblGenericName.Item]) = False Then GenericNameValue = ![tblGenericName.Item]
            If IsNull(![tblTradeName.Item]) = False Then TradeNameValue = ![tblTradeName.Item]
            If IsNull(![tblStrengthUnit.ItemUnit]) = False Then StrengthUnitValue = ![tblStrengthUnit.ItemUnit]
            If IsNull(![tblPackUnit.ItemUnit]) = False Then PackUnitValue = ![tblPackUnit.ItemUnit]
            If IsNull(![tblIssueUnit.ItemUnit]) = False Then IssueUnitValue = ![tblIssueUnit.ItemUnit]
            If IsNull(![tblDoseUnit.ItemUnit]) = False Then DoseUnitValue = ![tblDoseUnit.ItemUnit]
            If IsNull(![tblManufacturer.Institution]) = False Then ManufacturerValue = ![tblManufacturer.Institution]
            If IsNull(![tblImporter.Institution]) = False Then ImporterValue = ![tblImporter.Institution]
            If IsNull(![tblSupplier.Institution]) = False Then SupplierValue = ![tblSupplier.Institution]
            
            If IsNull(![CategoryID]) = False Then CategoryIDValue = ![CategoryID]
            If IsNull(![tblGroup.ItemID]) = False Then GroupIDValue = ![tblGroup.ItemID]
            If IsNull(![tblSubGroup.ItemID]) = False Then SubGroupIDValue = ![tblSubGroup.ItemID]
            If IsNull(![tblGenericName.ItemID]) = False Then GenericNameIDValue = ![tblGenericName.ItemID]
            If IsNull(![tblTradeName.ItemID]) = False Then TradeNameIDValue = ![tblTradeName.ItemID]
            If IsNull(![tblStrengthUnit.ItemUnitID]) = False Then StrengthUnitIDValue = ![tblStrengthUnit.ItemUnitID]
            If IsNull(![tblPackUnit.ItemUnitID]) = False Then PackUnitIDValue = ![tblPackUnit.ItemUnitID]
            If IsNull(![tblIssueUnit.ItemUnitID]) = False Then IssueUnitIDValue = ![tblIssueUnit.ItemUnitID]
            If IsNull(![tblDoseUnit.ItemUnitID]) = False Then DoseUnitIDValue = ![tblDoseUnit.ItemUnitID]
            If IsNull(![tblManufacturer.InstitutionID]) = False Then ManufacturerIDValue = ![tblManufacturer.InstitutionID]
            If IsNull(![tblImporter.InstitutionID]) = False Then ImporterIDValue = ![tblImporter.InstitutionID]
            If IsNull(![tblSupplier.InstitutionID]) = False Then SupplierIDValue = ![tblSupplier.InstitutionID]
                        
            If Not IsNull(!MinIQty) Then MinIQtyValue = .Fields("MinIQty").Value
            If Not IsNull(!MaxIQty) Then MaxIQtyValue = !MaxIQty
            If Not IsNull(!DivQty) Then DivQtyValue = !DivQty
                        
                        
            If Not IsNull(!ROL) Then ROLValue = .Fields("ROL").Value
            If Not IsNull(!ROQ) Then ROQValue = .Fields("ROQ").Value
            If Not IsNull(!MinROQ) Then MinQtyValue = .Fields("MinROQ").Value
            If Not IsNull(!StrengthUnitsPerIssueUnit) Then StrengthUnitsPerIssueUnitValue = .Fields("StrengthUnitsPerIssueUnit").Value
            If Not IsNull(!IssueUnitsPerPack) Then IssueUnitsPerPackValue = .Fields("IssueUnitsPerPack").Value
            If Not IsNull(!SalesMargin) Then ProfitMargine = .Fields("SalesMargin").Value
        
        
            IsGenericNameValue = ![IsGenericName]
            IsTradeNameValue = ![IsTradeName]
            IsGroupValue = ![IsGroup]
            IsSubGroupValue = ![IsSubGroup]
            IsItemNameValue = ![IsItemName]
        
            SalePriceValue = ![SalePrice]
            PurchasePriceValue = ![PurchasePrice]
            
            LastSalePriceValue = ![LastSalePrice]
            LastPurchasePriceValue = ![LastPurchasePrice]
            
            CommentsValue = Format(![Comments], "")
        End If
        If .State = 1 Then .Close
    End With
End Property

Public Property Get PurchasePrice() As Double
    PurchasePrice = PurchasePriceValue
End Property

Public Property Get SalePrice() As Double
    SalePrice = SalePriceValue
End Property

Public Property Get LastPurchasePrice() As Double
    LastPurchasePrice = LastPurchasePriceValue
End Property

Public Property Get LastSalePrice() As Double
    LastSalePrice = LastSalePriceValue
End Property

Public Property Get ID() As Long
    ID = IDValue
End Property

Public Property Get Item() As String
    Item = ItemValue
End Property

Public Property Get Code() As String
    Code = CodeValue
End Property

Public Property Get Category() As String
    Category = CategoryValue
End Property

Public Property Get CategoryID() As Long
    CategoryID = CategoryIDValue
End Property

Public Property Get GenericName() As String
    GenericName = GenericNameValue
End Property

Public Property Get GenericNameID() As Long
    GenericNameID = GenericNameIDValue
End Property

Public Property Get TradeName() As String
    TradeName = TradeNameValue
End Property

Public Property Get TradeNameID() As Long
    TradeNameID = TradeNameIDValue
End Property

Public Property Get Manufacturer() As String
    Manufacturer = ManufacturerValue
End Property

Public Property Get ManufacturerID() As String
    ManufacturerID = ManufacturerIDValue
End Property

Public Property Get Importer() As String
    Importer = ImporterValue
End Property

Public Property Get ImporterID() As String
    ImporterID = ImporterIDValue
End Property

Public Property Get Supplier() As String
    Supplier = SupplierValue
End Property

Public Property Get SupplierID() As Long
    SupplierID = SupplierIDValue
End Property

Public Property Get ROL() As Double
    ROL = ROLValue
End Property

Public Property Get ROQ() As Double
    ROQ = ROQValue
End Property

Public Property Get MinQty() As Double
    MinQty = MinQtyValue
End Property

Public Property Get MinIQty() As Double
    MinIQty = MinIQtyValue
End Property

Public Property Get MaxIQty() As Double
    MaxIQty = MaxIQtyValue
End Property

Public Property Get DivQty() As Double
    DivQty = DivQtyValue
End Property

Public Property Get StrengthUnit() As String
    StrengthUnit = StrengthUnitValue
End Property

Public Property Get StrengthUnitID() As Double
    StrengthUnitID = StrengthUnitIDValue
End Property

Public Property Get IssueUnit() As String
    IssueUnit = IssueUnitValue
End Property

Public Property Get IssueUnitID() As Long
    IssueUnitID = IssueUnitIDValue
End Property

Public Property Get PackUnit() As String
    PackUnit = PackUnitValue
End Property

Public Property Get PackUnitID() As String
    PackUnitID = PackUnitIDValue
End Property

Public Property Get StrengthUnitsPerIssueUnit()
    StrengthUnitsPerIssueUnit = StrengthUnitsPerIssueUnitValue
End Property

Public Property Get IssueUnitsPerPack()
    IssueUnitsPerPack = IssueUnitsPerPackValue
End Property

Public Property Get SalesMargin() As Double
    SalesMargin = ProfitMargine
End Property

Public Property Get IsGenericName() As Boolean
    IsGenericName = IsGenericNameValue
End Property

Public Property Get IsTradeName() As Boolean
    IsTradeName = IsTradeNameValue
End Property

Public Property Get IsGroup() As Boolean
    IsGroup = IsGroupValue
End Property

Public Property Get IsSubGroup() As Boolean
    IsSubGroup = IsSubGroupValue
End Property

Public Property Get IsItem() As Boolean
    IsItem = IsItemNameValue
End Property

Public Property Get GroupID() As Long
    GroupID = GroupIDValue
End Property

Public Property Get SubGroupID() As Long
    SubGroupID = SubGroupIDValue
End Property

Public Property Get Group() As String
    Group = GroupValue
End Property

Public Property Get SubGroup() As String
    SubGroup = SubGroupValue
End Property

Public Property Get Comments() As String
    Comments = CommentsValue
End Property
