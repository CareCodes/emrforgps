VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintFormatItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Private mvarX1 As Long 'local copy
    Private mvarY1 As Long 'local copy
    Private mvarX2 As Long 'local copy
    Private mvarY2 As Long 'local copy
    Private mvarScaleX1 As Double 'local copy
    Private mvarScaleX2 As Double 'local copy
    Private mvarScaleY1 As Double 'local copy
    Private mvarScaleY2 As Double 'local copy
    Private mvarRadious As Long 'local copy
    Private mvarScaleRadious As Double 'local copy
    Private mvarIsLabel As Boolean 'local copy
    Private mvarIsValue As Boolean 'local copy
    Private mvarIsText As Boolean 'local copy
    Private mvarIsLine As Boolean 'local copy
    Private mvarIsCircle As Boolean 'local copy
    Private mvarIsRectangle As Boolean 'local copy
    Private mvarItemName As String 'local copy
    Private mvarItemText As String 'local copy
    Private mvarZOrder As Integer 'local copy
    Private mvarIsFocus As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarFontName As String 'local copy
Private mvarFontSize As Integer 'local copy
Private mvarFontBold As Boolean 'local copy
Private mvarFontItalic As Boolean 'local copy
Private mvarFontUnderline As Boolean 'local copy
Private mvarFontStrikethough As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarTextAlignmane As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarForeColour As Long 'local copy
Private mvarBackColour As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarSerialNo As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarIsCalc As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarIsFlag As Boolean 'local copy
Public Property Let IsFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsFlag = 5
    mvarIsFlag = vData
End Property


Public Property Get IsFlag() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsFlag
    IsFlag = mvarIsFlag
End Property



Public Property Let IsCalc(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsCalc = 5
    mvarIsCalc = vData
End Property


Public Property Get IsCalc() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsCalc
    IsCalc = mvarIsCalc
End Property



Public Property Let SerialNo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SerialNo = 5
    mvarSerialNo = vData
End Property


Public Property Get SerialNo() As Long
Attribute SerialNo.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SerialNo
    SerialNo = mvarSerialNo
End Property



Public Property Let ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



Public Property Let BackColour(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BackColour = 5
    mvarBackColour = vData
End Property


Public Property Get BackColour() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BackColour
    BackColour = mvarBackColour
End Property



Public Property Let ForeColour(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeColour = 5
    mvarForeColour = vData
End Property


Public Property Get ForeColour() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeColour
    ForeColour = mvarForeColour
End Property



Public Property Let TextAlignment(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TextAlignmane = 5
    mvarTextAlignmane = vData
End Property


Public Property Get TextAlignment() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TextAlignmane
    TextAlignment = mvarTextAlignmane
End Property



Public Property Let FontStrikethough(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontStrikethough = 5
    mvarFontStrikethough = vData
End Property


Public Property Get FontStrikethough() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontStrikethough
    FontStrikethough = mvarFontStrikethough
End Property



Public Property Let FontUnderline(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontUnderline = 5
    mvarFontUnderline = vData
End Property


Public Property Get FontUnderline() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontUnderline
    FontUnderline = mvarFontUnderline
End Property



Public Property Let FontItalic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontItalic = 5
    mvarFontItalic = vData
End Property


Public Property Get FontItalic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontItalic
    FontItalic = mvarFontItalic
End Property



Public Property Let FontBold(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontBold = 5
    mvarFontBold = vData
End Property


Public Property Get FontBold() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontBold
    FontBold = mvarFontBold
End Property



Public Property Let FontSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontSize = 5
    mvarFontSize = vData
End Property


Public Property Get FontSize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontSize
    FontSize = mvarFontSize
End Property



Public Property Let FontName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FontName = 5
    mvarFontName = vData
End Property


Public Property Get FontName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FontName
    FontName = mvarFontName
End Property




Public Property Let IsFocus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsFocus = 5
    mvarIsFocus = vData
End Property


Public Property Get IsFocus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsFocus
    IsFocus = mvarIsFocus
End Property



Public Property Let ZOrder(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ZOrder = 5
    mvarZOrder = vData
End Property


Public Property Get ZOrder() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ZOrder
    ZOrder = mvarZOrder
End Property



Public Property Let ItemText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemText = 5
    mvarItemText = vData
End Property


Public Property Get ItemText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemText
    ItemText = mvarItemText
End Property



Public Property Let ItemName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemName = 5
    mvarItemName = vData
End Property


Public Property Get ItemName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemName
    ItemName = mvarItemName
End Property



Public Property Let IsRectangle(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsRectangle = 5
    mvarIsRectangle = vData
End Property


Public Property Get IsRectangle() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsRectangle
    IsRectangle = mvarIsRectangle
End Property



Public Property Let IsCircle(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsCircle = 5
    mvarIsCircle = vData
End Property


Public Property Get IsCircle() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsCircle
    IsCircle = mvarIsCircle
End Property



Public Property Let IsLine(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsLine = 5
    mvarIsLine = vData
End Property


Public Property Get IsLine() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsLine
    IsLine = mvarIsLine
End Property



Public Property Let IsText(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsText = 5
    mvarIsText = vData
End Property


Public Property Get IsText() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsText
    IsText = mvarIsText
End Property



Public Property Let IsValue(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsValue = 5
    mvarIsValue = vData
End Property


Public Property Get IsValue() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsValue
    IsValue = mvarIsValue
End Property



Public Property Let IsLabel(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsLabel = 5
    mvarIsLabel = vData
End Property


Public Property Get IsLabel() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsLabel
    IsLabel = mvarIsLabel
End Property

Public Property Let ScaleRadious(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScaleRadious = 5
    mvarScaleRadious = vData
End Property


Public Property Get ScaleRadious() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScaleRadious
    ScaleRadious = mvarScaleRadious
End Property



Public Property Let Radious(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Radious = 5
    mvarRadious = vData
End Property


Public Property Get Radious() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Radious
    Radious = mvarRadious
End Property



Public Property Let ScaleY2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScaleY2 = 5
    mvarScaleY2 = vData
End Property


Public Property Get ScaleY2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScaleY2
    ScaleY2 = mvarScaleY2
End Property



Public Property Let ScaleY1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScaleY1 = 5
    mvarScaleY1 = vData
End Property


Public Property Get ScaleY1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScaleY1
    ScaleY1 = mvarScaleY1
End Property



Public Property Let ScaleX2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScaleX2 = 5
    mvarScaleX2 = vData
End Property


Public Property Get ScaleX2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScaleX2
    ScaleX2 = mvarScaleX2
End Property



Public Property Let ScaleX1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ScaleX1 = 5
    mvarScaleX1 = vData
End Property


Public Property Get ScaleX1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ScaleX1
    ScaleX1 = mvarScaleX1
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



Public Property Let X1(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X1 = 5
    mvarX1 = vData
End Property


Public Property Get X1() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X1
    X1 = mvarX1
End Property



