VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Items"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
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
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      Height          =   6270
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   600
      Width           =   6975
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6000
      TabIndex        =   0
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
   Begin btButtonEx.ButtonEx btnImport 
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Import"
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
   Begin btButtonEx.ButtonEx btnSelectFile 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Select File"
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
End
Attribute VB_Name = "frmImportItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim temFile As String

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
    
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    
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


Private Sub btnImport_Click()
    Dim nFileNum As Integer
    Dim sText As String
    Dim sNextLine As String
    Dim lLineCount As Long
    Dim FirstTabLocation As Integer
    
    Dim i As Integer
    Dim n As Integer
    
    Dim ItemName As String
    Dim GenericName As String
    Dim TradeName As String
    Dim FieldValue()
    Dim FieldNo As Long
    Dim RemainingText As String
    
    Dim GenericNameID As Long
    Dim TradeNameID As Long

    Dim rsTem As New ADODB.Recordset

    nFileNum = FreeFile

    Open temFile For Input As nFileNum
    lLineCount = 1


    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReDim FieldValue(.Fields.Count)
    End With


    Do While Not EOF(nFileNum)
       Line Input #nFileNum, sNextLine
            sNextLine = sNextLine & vbCrLf
            FirstTabLocation = InStr(sNextLine, vbTab)
            RemainingText = Right(sNextLine, Len(sNextLine) - FirstTabLocation)
            
            FirstTabLocation = InStr(RemainingText, vbTab)
            GenericName = Left(RemainingText, FirstTabLocation - 1)
            GenericNameID = FindGenericNameID(GenericName)
            RemainingText = Right(sNextLine, Len(RemainingText) - FirstTabLocation)
    
            FirstTabLocation = InStr(RemainingText, vbTab)
            TradeName = Left(RemainingText, FirstTabLocation - 1)
            TradeNameID = FindTradeNameID(TradeName)
            RemainingText = Right(sNextLine, Len(RemainingText) - FirstTabLocation)
            
            i = 0
            
            While InStr(RemainingText, vbTab) > 0
                FirstTabLocation = InStr(RemainingText, vbTab)
                FieldValue(i) = Left(RemainingText, FirstTabLocation - 1)
                RemainingText = Right(sNextLine, Len(RemainingText) - FirstTabLocation)
                i = i + 1
            Wend
            
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblItem"
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                .AddNew
                For n = 1 To .Fields.Count - 1
                    If IsNull(FieldValue(n)) = False And FieldValue(n) <> "" Then
                        .Fields(n).Value = FieldValue(n)
                    End If
                Next
                .Fields("TradeNameID").Value = TradeNameID
                .Fields("GenericNameID").Value = GenericNameID
                .Update
            End With
            
    Loop
    

    Close nFileNum
    
End Sub

Private Sub btnNone_Click()
    With lstItem
        Dim i As Integer
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
    End With
End Sub

Private Sub btnSelectFile_Click()
    Dim nFileNum As Integer
    Dim sText As String
    Dim sNextLine As String
    Dim lLineCount As Long
    
    Dim FirstTabLocation As Integer
    
    cdlExport.Flags = cdlOFNNoChangeDir
    cdlExport.DefaultExt = "EEI"
    cdlExport.Filter = "EMA 3 Item File|*.EEI"
    cdlExport.ShowOpen
    temFile = cdlExport.filename
    
    ' Get a free file number
    nFileNum = FreeFile

    lstItem.Clear
    
    ' Open a text file for input. inputbox returns the path to read the file
    Open temFile For Input As nFileNum
    
    lLineCount = 1
    ' Read the contents of the file
    
    Do While Not EOF(nFileNum)
       Line Input #nFileNum, sNextLine
            sNextLine = sNextLine & vbCrLf
            FirstTabLocation = InStr(sNextLine, vbTab)
            lstItem.AddItem Left(sNextLine, FirstTabLocation)
       'sText = sText & sNextLine
    Loop
    
    ' Close the file
    Close nFileNum
 
End Sub

Private Sub Form_Load()
    Call SetColours
    Call GetSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

