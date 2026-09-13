VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddSpacesInbetween"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    
Public Function AddSPace(InputWord As String, Optional NumberOfSpaces As Long) As String
    Dim temWord As String
    Dim WordLength As Long
    Dim i As Long
    If NumberOfSpaces = 0 Then NumberOfSpaces = 1
    For i = 1 To Len(InputWord)
        temWord = temWord & Mid(InputWord, i, 1) & Space(NumberOfSpaces)
    Next
    AddSPace = temWord
End Function
