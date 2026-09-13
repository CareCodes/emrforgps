VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim temSql As String
    
    Dim IDValue As Long
    Dim IxValue As String
    Dim IxValueValue As Double
    Dim IxCostValue As Double
    Dim IxTubeIDValue As Long
    Dim IxTubeValue As String
    Dim IxSpecimanIDValue As Long
    Dim IxSpecimanValue As String
    Dim IxDepartmentIDValue As Long
    Dim IxDepartmentValue As String
    Dim IxCommentValue As String
    Dim IxVolumeValue As String
    
Public Property Let ID(ByVal IxID As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIx.*, tblIxCategory.Category, tblTube.Category, tblSpeciman.Category, tblDepartment.Department " & _
                    "FROM (((tblIx LEFT JOIN tblDepartment ON tblIx.DepartmentID = tblDepartment.DepartmentID) LEFT JOIN tblCategory AS tblIxCategory ON tblIx.IxCategoryID = tblIxCategory.CategoryID) LEFT JOIN tblCategory AS tblTube ON tblIx.TubeUnitID = tblTube.CategoryID) LEFT JOIN tblCategory AS tblSpeciman ON tblIx.SpecimanID = tblSpeciman.CategoryID " & _
                    "WHERE (((tblIx.IxID)=" & IxID & "))"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            IDValue = IxID
            IxValue = !Ix
            If IsNull(!TubeID) = False Then IxTubeIDValue = !TubeID
            If IsNull(![tblTube.Category]) = False Then IxTubeValue = ![tblTube.Category]
            If IsNull(!SpecimanID) = False Then IxSpecimanIDValue = !SpecimanID
            If IsNull(![tblSpeciman.Category]) = False Then IxSpecimanValue = ![tblSpeciman.Category]
            If IsNull(!DepartmentID) = False Then IxDepartmentIDValue = !DepartmentID
            If IsNull(!Department) = False Then IxDepartmentValue = !Department
            If IsNull(!Comments) = False Then IxCommentValue = !Comments
            If IsNull(!TubeVolume) = False Then IxVolumeValue = !TubeVolume
            If IsNull(!IxValue) = False Then IxValueValue = !IxValue
            If IsNull(!IxCost) = False Then IxCostValue = !IxCost
        Else
            IDValue = IxID
            IxValue = Empty
            IxValueValue = Empty
            IxCostValue = Empty
            IxTubeIDValue = Empty
            IxTubeValue = Empty
            IxSpecimanIDValue = Empty
            IxSpecimanValue = Empty
            IxDepartmentIDValue = Empty
            IxDepartmentValue = Empty
            IxCommentValue = Empty
            IxVolumeValue = Empty
        End If
        .Close
    End With
End Property

Public Property Get IxID() As Long
    ID = IDValue
End Property

Public Property Get Ix() As String
    Ix = IxValue
End Property

Public Property Get Tube() As String
    Tube = IxTubeValue
End Property

Public Property Get TubeID() As Long
    TubeID = IxTubeIDValue
End Property

Public Property Get Speciman() As String
    Speciman = IxSpecimanValue
End Property

Public Property Get SpecimanID() As Long
    SpecimanID = IxSpecimanIDValue
End Property

Public Property Get Department() As String
    Department = IxDepartmentValue
End Property

Public Property Get DepartmentID() As Long
    DepartmentID = IxDepartmentIDValue
End Property

Public Property Get Comments() As String
    Comments = IxCommentValue
End Property

Public Property Get Volume() As Double
    Volume = IxVolumeValue
End Property

Public Property Get Value() As Double
    Value = IxValueValue
End Property

Public Property Get Cost() As Double
    Cost = IxCostValue
End Property
