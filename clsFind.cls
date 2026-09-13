VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim temSql As String
    Dim rsCombo As New ADODB.Recordset
    
    
Public Sub FillCombo(FillCombo As DataCombo, FillTable As String, ListField As String, BoundField As String, NoDeletedValues As Boolean, Optional FilterField As String, Optional FilterID As Long)
    With rsCombo
        If .State = 1 Then .Close
        If NoDeletedValues = True Then
            If FilterField <> "" And FilterID <> 0 Then
                temSql = "SELECT " & ListField & " , " & BoundField & " FROM " & FillTable & " WHERE Deleted = False And " & FilterField & " = " & FilterID & " ORDER BY " & ListField
            Else
                temSql = "SELECT " & ListField & " , " & BoundField & " FROM " & FillTable & " WHERE Deleted = False ORDER BY " & ListField
            End If
        Else
            If FilterField <> "" And FilterID <> 0 Then
                temSql = "SELECT " & ListField & " , " & BoundField & " FROM " & FillTable & " WHERE " & FilterField & " = " & FilterID & " ORDER BY " & ListField
            Else
                temSql = "SELECT " & ListField & " , " & BoundField & " FROM " & FillTable & " ORDER BY " & ListField
            End If
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With FillCombo
        Set .RowSource = rsCombo
        .ListField = ListField
        .BoundColumn = BoundField
    End With
End Sub

