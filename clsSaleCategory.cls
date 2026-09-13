VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaleCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim SaleDiscountPercentValue As Double
    Dim ProfitMarginValue As Double
    Dim PaymentMethodIDValue As Long
    Dim PaymentMethodValue As String
    Dim SaleCategoryIDValue As Long
    Dim SaleCategoryValue As String
    Dim ForOutDoorValue As Boolean
    Dim ForBHTValue As Boolean
    Dim ForStaffValue As Boolean
    Dim ForUnitValue As Boolean
    Dim ForInstitutionValue As Boolean
    Dim temSql As String
    Dim rsTemSale As New ADODB.Recordset

Public Property Let SaleCategoryID(IDValue As Long)
    SaleCategoryIDValue = IDValue
    With rsTemSale
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleCategory.*, tblPaymentMethod.PaymentMethod " & _
                    " FROM tblPaymentMethod RIGHT JOIN tblSaleCategory ON tblPaymentMethod.PaymentMethodID = tblSaleCategory.PaymentMethodID " & _
                    " WHERE (((tblSaleCategory.SaleCategoryID)=" & SaleCategoryIDValue & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        If .RecordCount > 0 Then
            SaleCategoryValue = .Fields("SaleCategory").Value
            SaleDiscountPercentValue = .Fields("SaleDiscountPercent").Value
            PaymentMethodIDValue = !PaymentMethodID
            PaymentMethodValue = !PaymentMethod
            ForOutDoorValue = .Fields("ForOutDoor").Value
            ForBHTValue = !ForBHT
            If IsNull(!ProfitMargin) = False Then
                ProfitMarginValue = !ProfitMargin
            Else
                ProfitMarginValue = 0
            End If
            ForStaffValue = .Fields("ForStaff").Value
            ForUnitValue = .Fields("ForUnit").Value
            ForInstitutionValue = !ForInstitution
        Else
            SaleCategoryValue = Empty
            SaleDiscountPercentValue = Empty
            PaymentMethodIDValue = Empty
            PaymentMethodValue = Empty
            ForOutDoorValue = Empty
            ForBHTValue = Empty
            ProfitMarginValue = Empty
            ForStaffValue = Empty
            ForUnitValue = Empty
            ForInstitutionValue = Empty
        End If
        If .State = 1 Then .Close
    End With
End Property



Public Property Get SaleCategoryID() As Long
    SaleCategoryID = SaleCategoryIDValue
End Property

Public Property Get SaleCategory() As String
    SaleCategory = SaleCategoryValue
End Property

Public Property Get ProfitMargin() As Double
    ProfitMargin = ProfitMarginValue
End Property

Public Property Get SaleDiscountPercent() As Double
    SaleDiscountPercent = SaleDiscountPercentValue
End Property

Public Property Get PaymentMethod() As String
    PaymentMethod = Trim(UCase(PaymentMethodValue))
End Property

Public Property Get PaymentMethodID() As Long
    PaymentMethodID = PaymentMethodIDValue
End Property

Public Property Get ForBHT() As Boolean
    ForBHT = ForBHTValue
End Property

Public Property Get ForOutDoor() As Boolean
    ForOutDoor = ForOutDoorValue
End Property

Public Property Get ForStaff() As Boolean
    ForStaff = ForStaffValue
End Property

Public Property Get ForUnit() As Boolean
    ForUnit = ForUnitValue
End Property

Public Property Get ForInstitution() As Boolean
    ForInstitution = ForInstitutionValue
End Property
