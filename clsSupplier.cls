VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim NameValue As String
    Dim IDValue As String
    Dim CodeValue As String
    Dim IsAFarmerValue As Boolean
    Dim IsACollectorValue As Boolean
    Dim IsDeletedValue As Boolean
    Dim CollectingCenterIDValue As Long
    Dim AddressValue As String
    Dim ExistsValue As Boolean
    Dim HasCommisionValue As Boolean
    Dim AccountNoValue As String
    Dim AccountHolderValue As String
    
    Dim rsSupplier As New ADODB.Recordset
    Dim temSql As String
    
    
Public Property Let ID(ID As Long)
    IDValue = ID
    With rsSupplier
        temSql = "Select * from tblSupplier where SupplierID = " & IDValue
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            ExistsValue = True
            NameValue = !Supplier
            CodeValue = !SupplierCode
            IsAFarmerValue = !Farmer
            HasCommisionValue = !Commision
            IsACollectorValue = !Collector
            If IsNull(!Address) = False Then AddressValue = Format(!Address, "")
           AccountNoValue = !AccountNo
            If IsNull(!AccountHolder) = False Then
                AccountHolderValue = !AccountHolder
            Else
                AccountHolderValue = Empty
            End If
        Else
            ExistsValue = False
        End If
    End With
End Property

Public Property Get ID() As Long
    ID = IDValue
End Property

Public Property Get Name() As String
    Name = NameValue
End Property

Public Property Get Code() As String
    Code = CodeValue
End Property

Public Property Get Address() As String
    Address = AddressValue
End Property

Public Property Get HasCommision() As Boolean
    HasCommision = HasCommisionValue
End Property

Public Property Get AccountNo() As String
    AccountNo = AccountNoValue
End Property

Public Property Get AccountHolder() As String
    AccountHolder = AccountHolderValue
End Property
