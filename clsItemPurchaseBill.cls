VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPurchaseBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim temSql As String
    Dim BillDateValue As Date
    Dim BillTimeValue As Date
    Dim BillUserIDValue As Long
    Dim BillUserValue As String
    Dim DepartmentIDValue As Long
    Dim DepartmentValue As String
    Dim PaymentMethodIDValue As Long
    Dim PaymentMethodValue As String
    Dim SupplierIDValue As Long
    Dim SupplierValue As String
    Dim InvoiceNumberValue As String
    Dim InvoiceDateValue As Date
    Dim GRNNumberValue As String
    Dim TotalValue As Double
    Dim DiscountValue As Double
    Dim DiscountPercentValue As Double
    
    
    
Public Function NewPurchaseBillID() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select tblPurchaseBill.* from tblpurchasebill where cancelled = false AND Completed = False AND departmentID = " & UserDepartmentID & " AND Deleted = False order by PurchaseBillID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            NewPurchaseBillID = !PurchaseBillID
        Else
            .AddNew
            !BillDateTime = Date
            !BillUserID = UserID
            !DepartmentID = UserDepartmentID
            .Update
            NewPurchaseBillID = !PurchaseBillID
        End If
        .Close
    End With
End Function

