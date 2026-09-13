VERSION 5.00
Begin VB.UserControl UCPopup 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   0  'User
   ScaleWidth      =   202.105
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopupItem 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "UCPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private m_ReturnIndex As Integer
    Private m_ReturnCaption As String

Public Property Get ReturnIndex() As Integer
    ReturnIndex = m_ReturnIndex
End Property

Public Property Get ReturnCaption() As String
    ReturnCaption = m_ReturnCaption
End Property

Private Sub mnuPopupItem_Click(Index As Integer)
    m_ReturnIndex = Index
    m_ReturnCaption = mnuPopupItem(Index).Caption
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 180, 210
End Sub

Public Function Popup(Items As Variant) As Integer
    On Error Resume Next
    Dim i As Long
    Dim n As Long
    
    n = UBound(Items)
    For i = 0 To n
    If i > mnuPopupItem.Count - 1 Then Load mnuPopupItem(i)
        mnuPopupItem(i).Caption = Items(i)
    Next
    Do Until mnuPopupItem.Count - 1 <= n
        Unload mnuPopupItem(mnuPopupItem.Count - 1)
    Loop
    '//Reset
    m_ReturnIndex = -1
    m_ReturnCaption = vbNullString
    '//Wait user response
    UserControl.PopupMenu mnuPopup
    Popup = m_ReturnIndex
End Function
