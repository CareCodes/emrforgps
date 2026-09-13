VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFillCombos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim rsFill As New ADODB.Recordset
    Dim temSql As String
    
    
    
Public Sub FillAnyCombo(ComboToFill As DataCombo, table As String, Optional DoNotIncludeDeleted As Boolean)
    temSql = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then temSql = temSql & " Where Deleted = False "
    temSql = temSql & " Order by " & table
    With rsFill
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = table
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillBoolCombo(ComboToFill As DataCombo, table As String, ListField As String, BoolField As String, Optional DoNotIncludeDeleted As Boolean)
    temSql = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSql = temSql & " Where Deleted = False AND " & BoolField & " = True "
    Else
        temSql = temSql & " Where " & BoolField & " = True "
    End If
    temSql = temSql & " Order by " & ListField
    With rsFill
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = ListField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillSpecificField(ComboToFill As DataCombo, table As String, DisplayField As String, Optional DoNotIncludeDeleted As Boolean)
    temSql = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then temSql = temSql & " Where Deleted = False "
    temSql = temSql & " Order by " & DisplayField
    With rsFill
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = DisplayField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillSpecificFieldBoolCombo(ComboToFill As DataCombo, table As String, ListField As String, DisplayField As String, BoolField As String, Optional DoNotIncludeDeleted As Boolean)
    temSql = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSql = temSql & " Where Deleted = False AND " & BoolField & " = True "
    Else
        temSql = temSql & " Where " & BoolField & " = True "
    End If
    temSql = temSql & " Order by " & DisplayField
    With rsFill
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = DisplayField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillLongCombo(ComboToFill As DataCombo, table As String, ListField As String, LongField As String, LongValue As Long, Optional DoNotIncludeDeleted As Boolean)
    temSql = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSql = temSql & " Where Deleted = False AND " & LongField & " = " & LongValue
    Else
        temSql = temSql & " Where " & LongField & " =  " & LongValue
    End If
    temSql = temSql & " Order by " & ListField
    With rsFill
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = ListField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillLongBoolCombo(ComboToFill As DataCombo, table As String, ListField As String, LongField As String, LongValue As Long, BoolField As String, Optional DoNotIncludeDeleted As Boolean)
    temSql = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSql = temSql & " Where Deleted = False AND " & LongField & " = " & LongValue & " AND " & BoolField & " = True "
    Else
        temSql = temSql & " Where " & LongField & " =  " & LongValue & " AND " & BoolField & " = True "
    End If
    temSql = temSql & " Order by " & ListField
    With rsFill
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = ListField
        .BoundColumn = table & "ID"
    End With
End Sub
