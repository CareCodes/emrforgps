VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamilyRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarRelationID As Integer 'local copy
Private mvarRelation As String 'local copy
Public Property Let Relation(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Relation = 5
    mvarRelation = vData
End Property


Public Property Get Relation() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Relation
    Relation = mvarRelation
End Property



Public Property Let RelationID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RelationID = 5
    mvarRelationID = vData
End Property


Public Property Get RelationID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RelationID
    RelationID = mvarRelationID
End Property



