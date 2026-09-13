VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamilyUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    'local variable(s) to hold property value(s)
    Private mvarPatientID As Long 'local copy
    Private mvarAffected As Boolean 'local copy
    'local variable(s) to hold property value(s)
    Private mvarCentreUnit As Boolean 'local copy
    'local variable(s) to hold property value(s)
    Private mvarX1 As Long 'local copy
    Private mvarX2 As Long 'local copy
    Private mvarY1 As Long 'local copy
    Private mvarY2 As Long 'local copy
    'local variable(s) to hold property value(s)
    Private mvarSexID As Integer 'local copy

Public Property Let SexID(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SexID = 5
    mvarSexID = vData
End Property


Public Property Get SexID() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SexID
    SexID = mvarSexID
End Property



Public Property Let Y2(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y2 = 5
    mvarY2 = vData
End Property


Public Property Get Y2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y2
    Y2 = mvarY2
End Property



Public Property Let Y1(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y1 = 5
    mvarY1 = vData
End Property


Public Property Get Y1() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y1
    Y1 = mvarY1
End Property



Public Property Let X2(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X2 = 5
    mvarX2 = vData
End Property


Public Property Get X2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X2
    X2 = mvarX2
End Property



Public Property Let X1(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X1 = 5
    mvarX1 = vData
End Property


Public Property Get X1() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X1
'    Set X1 = mvarX1
End Property



Public Property Let MySelf(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CentreUnit = 5
    mvarCentreUnit = vData
End Property


Public Property Get MySelf() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CentreUnit
    MySelf = mvarCentreUnit
End Property




Public Property Let Affected(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Affected = 5
    mvarAffected = vData
End Property


Public Property Get Affected() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Affected
    Affected = mvarAffected
End Property



Public Property Let PatientID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PatientID = 5
    mvarPatientID = vData
End Property


Public Property Get PatientID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PatientID
    PatientID = mvarPatientID
End Property



