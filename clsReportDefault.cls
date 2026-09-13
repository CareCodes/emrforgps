VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String
    Dim LabelFontNameValue As String
    Dim LabelFontSizeValue As Long
    Dim LabelFontBoldValue As Boolean
    Dim LabelFontItalicValue As Boolean
    Dim LabelFontUnderlineValue As Boolean
    Dim LabelFontStrikeThroughValue As Boolean
    
    Dim TextFontNameValue As String
    Dim TextFontSizeValue As Long
    Dim TextFontBoldValue As Boolean
    Dim TextFontItalicValue As Boolean
    Dim TextFontUnderlineValue As Boolean
    Dim TextFontStrikeThroughValue As Boolean

    Dim LabelBackColourValue As Long
    Dim LabelForeColourValue As Long

    Dim TextBackColourValue As Long
    Dim TextForeColourValue As Long

Public Property Get LabelBackColour() As Long
    Call GetDefaults
    LabelBackColour = LabelBackColourValue
End Property

Public Property Get LabelForeColour() As Long
    Call GetDefaults
    LabelForeColour = LabelForeColourValue
End Property

Public Property Get TextBackColour() As Long
    Call GetDefaults
    TextBackColour = TextBackColourValue
End Property

Public Property Get TextForeColour() As Long
    Call GetDefaults
    TextForeColour = TextForeColourValue
End Property


Public Property Get TextFontName() As String
    If TextFontNameValue = "" Then GetDefaults
    TextFontName = TextFontNameValue
End Property

Public Property Get TextFontSize() As Long
    If TextFontSizeValue = 0 Then GetDefaults
    TextFontSize = TextFontSizeValue
End Property

Public Property Get TextFontBold() As Boolean
    GetDefaults
    TextFontBold = TextFontBoldValue
End Property
    
Public Property Get TextFontItalic() As Boolean
    GetDefaults
    TextFontItalic = TextFontItalicValue
End Property

Public Property Get TextFontUnderline() As Boolean
    GetDefaults
    TextFontUnderline = TextFontUnderlineValue
End Property

Public Property Get TextFontStrikeThrough() As Boolean
    GetDefaults
    TextFontStrikeThrough = TextFontStrikeThroughValue
End Property

Public Property Get LabelFontName() As String
    If LabelFontNameValue = "" Then GetDefaults
    LabelFontName = LabelFontNameValue
End Property

Public Property Get LabelFontSize() As Long
    If LabelFontSizeValue = 0 Then GetDefaults
    LabelFontSize = LabelFontSizeValue
End Property

Public Property Get LabelFontBold() As Boolean
    GetDefaults
    LabelFontBold = LabelFontBoldValue
End Property
    
Public Property Get LabelFontItalic() As Boolean
    GetDefaults
    LabelFontItalic = LabelFontItalicValue
End Property

Public Property Get LabelFontUnderline() As Boolean
    GetDefaults
    LabelFontUnderline = LabelFontUnderlineValue
End Property

Public Property Get LabelFontStrikeThrough() As Boolean
    GetDefaults
    LabelFontStrikeThrough = LabelFontStrikeThroughValue
End Property

Private Sub GetDefaults()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPrintDefault"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        
            LabelFontNameValue = !LabelFontName
            LabelFontSizeValue = !LabelFontSize
            LabelFontBoldValue = !LabelFontBold
            LabelFontItalicValue = !LabelFontItalic
            LabelFontStrikeThroughValue = !LabelFontStrikeThrough
            LabelFontUnderlineValue = !LabelFontUnderline
            
            TextFontNameValue = !TextFontName
            TextFontSizeValue = !TextFontSize
            TextFontBoldValue = !TextFontBold
            TextFontItalicValue = !TextFontItalic
            TextFontStrikeThroughValue = !TextFontStrikeThrough
            TextFontUnderlineValue = !TextFontUnderline
        
            LabelBackColourValue = vbYellow
            LabelForeColourValue = vbBlack
            
            TextBackColourValue = vbYellow
            TextForeColourValue = vbBlack
            
            
        End If
        .Close
    End With
End Sub
