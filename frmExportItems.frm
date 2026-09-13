VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExportItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Items"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   5640
      TabIndex        =   7
      Top             =   360
      Width           =   3255
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   83886083
         CurrentDate     =   39993
      End
      Begin VB.OptionButton optBetween 
         Caption         =   "&Added Between"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optAll 
         Caption         =   "&All"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   83886083
         CurrentDate     =   39993
      End
   End
   Begin MSComDlg.CommonDialog cdlExport 
      Left            =   6000
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnNone 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&None"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnAll 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&All"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstItem 
      Height          =   6540
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   480
      Width           =   5415
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnExport 
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Export"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstID 
      Height          =   6540
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmExportItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnAll_Click()
    With lstItem
        Dim i As Integer
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    
    optAll.Value = GetSetting(App.EXEName, Me.Name, optAll.Name, True)
    optBetween.Value = GetSetting(App.EXEName, Me.Name, optBetween.Name, False)
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, Me.Name, optAll.Name, optAll.Value
    SaveSetting App.EXEName, Me.Name, optBetween.Name, optBetween.Value
End Sub

Private Sub SetColours()
    Me.ForeColor = DefaultColourScheme.LabelForeColour
    Me.BackColor = DefaultColourScheme.LabelBackColour
    On Error Resume Next
    Dim MyControl As Control
    For Each MyControl In Controls
        If InStr(UCase(MyControl.Name), "BTN") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
        ElseIf InStr(UCase(MyControl.Name), "LST") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXTID") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "CMB") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ComboForeColour
            MyControl.BackColor = DefaultColourScheme.ComboBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXT") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        ElseIf InStr(UCase(MyControl.Name), "DTP") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        Else
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
            MyControl.BackStyle = 0
        End If
    Next
End Sub

Private Sub btnExport_Click()
    Dim FSys As New FileSystemObject
    Dim i As Integer
    Dim FC As Integer
    Dim temFile As String
    Dim rsTem As New ADODB.Recordset
    Dim TemLineText As String
    
    cdlExport.Flags = cdlOFNNoChangeDir
    cdlExport.DefaultExt = "EEI"
    cdlExport.Filter = "EMA 3 Item File|*.EEI"
    cdlExport.ShowSave
    
    temFile = cdlExport.filename
    Open temFile For Append As #1
    
    For i = 0 To lstItem.ListCount - 1
        If lstItem.Selected(i) = True Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "SELECT tblItem.*, tblGeneric.Item, tblTrade.Item FROM (tblItem LEFT JOIN tblItem AS tblTrade ON tblItem.TradeNameID = tblTrade.ItemID) LEFT JOIN tblItem AS tblGeneric ON tblItem.GenericNameID = tblGeneric.ItemID where tblItem.Itemid = " & Val(lstID.List(i))
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    TemLineText = lstItem.List(i) & vbTab & ![tblGeneric.Item] & vbTab & ![tblTrade.Item]
                    For FC = 0 To .Fields.Count - 1
                        If IsNull(.Fields(FC).Value) = False Then
                            TemLineText = TemLineText & vbTab & .Fields(FC).Value
                        Else
                            TemLineText = TemLineText & vbTab & ""
                        End If
                    Next
                End If
                .Close
            End With
            Print #1, TemLineText
        End If
    Next i
    
    Close #1
    
    MsgBox "Successfully Exported"

End Sub

Private Sub btnNone_Click()
    With lstItem
        Dim i As Integer
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
    End With
End Sub


Private Sub dtpTo_Change()
    Call FillList
End Sub

Private Sub dtpFrom_Change()
    Call FillList
End Sub

Private Sub Form_Load()
    Call SetColours
    Call GetSettings
    Call FillList
End Sub

Private Sub FillList()
    lstID.Clear
    lstItem.Clear
    Dim rsItem As New ADODB.Recordset
    With rsItem
        If .State = 1 Then .Close
        If optAll.Value = True Then
            temSql = "Select * from tblItem where IsItemName = true and Deleted = False order by Item"
        Else
            temSql = "Select * from tblItem where AddedDate between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# AND #" & Format(dtpTo.Value, "dd MMMM yyyy") & "# AND IsItemName = true and Deleted = False order by Item"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstItem.AddItem !Item
            lstID.AddItem !ItemID
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub optAll_Click()
    Call FillList
End Sub

Private Sub optBetween_Click()
    Call FillList
End Sub
