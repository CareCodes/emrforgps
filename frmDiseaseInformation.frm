VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDiseaseInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disease Information"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
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
   ScaleHeight     =   8670
   ScaleWidth      =   11265
   Begin VB.TextBox txtImagePath 
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtLink 
      Height          =   360
      Left            =   8280
      TabIndex        =   11
      Top             =   7320
      Width           =   2775
   End
   Begin VB.TextBox txtImage 
      Height          =   360
      Left            =   5400
      TabIndex        =   5
      Top             =   7320
      Width           =   2775
   End
   Begin MSDataListLib.DataList lstImage 
      Height          =   1980
      Left            =   5400
      TabIndex        =   3
      ToolTipText     =   "Click to display the picture, Double-click to get a large view."
      Top             =   4920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3493
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbDiagnosis 
      Height          =   7620
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   13441
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList lstLinks 
      Height          =   1980
      Left            =   8280
      TabIndex        =   9
      ToolTipText     =   "Double-click to visit the weblink"
      Top             =   4920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3493
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnAddImage 
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnDeletedImage 
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnAddLink 
      Height          =   375
      Left            =   9000
      TabIndex        =   12
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnDeletedLink 
      Height          =   375
      Left            =   10080
      TabIndex        =   13
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Weblink"
      Height          =   255
      Left            =   8280
      TabIndex        =   10
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Image Details"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "&Weblink"
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "&Images"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Image imgDx 
      Height          =   4500
      Left            =   5400
      Stretch         =   -1  'True
      ToolTipText     =   "Double click to get a large view"
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "&Diagnosis"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDiseaseInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
    Dim rsImage As New ADODB.Recordset
    Dim rsLink As New ADODB.Recordset
    Dim temSql As String
    Dim FSys As New FileSystemObject
    

Private Sub btnAddImage_Click()
    If IsNumeric(cmbDiagnosis.BoundText) = False Then
        MsgBox "Please select a diagnosis"
        cmbDiagnosis.SetFocus
        Exit Sub
    End If
    
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "Bitmap|*.BMP|JPEG|*.JPG;*.JPE;*.JPEG|GIF|*.GIF|All Supported Images|*.BMP;*.JPG;*.JPE;*.JPEG;*.GIF"
    On Error GoTo eh
    CommonDialog1.ShowOpen
    imgDx.Stretch = True
    If CommonDialog1.CancelError = False Then
        txtImagePath.Text = CommonDialog1.filename
        Set imgDx.Picture = LoadPicture(txtImagePath.Text)
    Else
        MsgBox "You have not selected valid image.", vbCritical
        Set imgDx.Picture = LoadPicture()
        txtImagePath.Text = Empty
    End If
    Dim rsTem As New ADODB.Recordset
    If FSys.FileExists(CommonDialog1.filename) = False Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblImage"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        If Trim(txtImage.Text) = "" Then
            !Image = FSys.GetFileName(CommonDialog1.filename)
        Else
            !Image = txtImage.Text
        End If
        !DiagnosisID = Val(cmbDiagnosis.BoundText)
        
        GetPhoto txtImagePath.Text, rsTem, "ImageData", "ImageSize"
        .Update
        .Close
    End With
    Call FillImages
    txtImage.Text = Empty
    Exit Sub
eh:
    MsgBox "Error loading the image"
End Sub

Private Sub btnAddLink_Click()
    If IsNumeric(cmbDiagnosis.BoundText) = False Then
        MsgBox "Please select a diagnosis"
        cmbDiagnosis.SetFocus
        Exit Sub
    End If
    If Trim(txtLink.Text) = Empty Then
        MsgBox "Please enter a link"
        txtLink.SetFocus
        Exit Sub
    End If
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblLink"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DiagnosisID = Val(cmbDiagnosis.BoundText)
        !Link = Trim(txtLink.Text)
        .Update
        .Close
    End With
    Call FillLinks
    txtLink.Text = Empty
    txtLink.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDeletedImage_Click()
    If IsNumeric(lstImage.BoundText) = False Then
        MsgBox "Noting to Delete"
        lstImage.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete the Photo?", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Delete * from tblImage where ImageID = " & Val(lstImage.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        If .RecordCount > 0 Then
'            !Deleted = True
'            !DeletedUserID = UserID
'            !DeletedTime = Now
'            .Update
'        End If
        If .State = 1 Then .Close
    End With
    Call FillImages
End Sub

Private Sub btnDeletedLink_Click()
    If IsNumeric(lstLinks.BoundText) = False Then
        MsgBox "Noting to Delete"
        lstLinks.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete the Link?", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblLink where LinkID = " & Val(lstLinks.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedTime = Now
            .Update
        End If
        .Close
    End With
    Call FillLinks
End Sub

Private Sub cmbDiagnosis_Change()
    Call FillImages
    Call FillLinks
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

Private Sub Form_Load()
    Call SetColours
    GetCommonSettings Me
    
    
    Call FillCombos
End Sub

Private Sub FillCombos()
    Dim Pt As New clsFillCombos
    Pt.FillBoolCombo cmbDiagnosis, "Diagnosis", "Diagnosis", "IsDiagnosis", True
End Sub

Private Sub FillImages()
    Dim tsmID As Long
    Dim temID As Long
    With rsImage
        If .State = 1 Then .Close
        temSql = "Select * from tblImage where Deleted = False and DiagnosisID = " & Val(cmbDiagnosis.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temID = !ImageID
        Else
            temID = 0
        End If
    End With
    With lstImage
        Set .RowSource = rsImage
        .ListField = "Image"
        .BoundColumn = "ImageID"
        .BoundText = temID
    End With
    lstImage_Click
End Sub

Private Sub FillLinks()
    With rsLink
        If .State = 1 Then .Close
        temSql = "Select * from tblLink where Deleted = False and DiagnosisID = " & Val(cmbDiagnosis.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstLinks
        Set .RowSource = rsLink
        .ListField = "Link"
        .BoundColumn = "LinkID"
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub imgDx_Click()
    If IsNumeric(lstImage.BoundText) = False Then Exit Sub
    lstImage_DblClick
End Sub

Private Sub lstImage_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblImage where ImageID = " & Val(lstImage.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            FillImage rsTem, "ImageData", "ImageSize", imgDx
        Else
            imgDx.Picture = LoadPicture()
        End If
        .Close
    End With
End Sub

Private Sub lstImage_DblClick()
    lstImage_Click
    Dim MyForm As Form
    Set MyForm = New frmDiagnosisImage
    MyForm.Show
    MyForm.ZOrder 0
End Sub

Private Sub lstLinks_DblClick()
    If IsNumeric(lstLinks.BoundText) = False Then Exit Sub
    Dim r As Long
    r = ShellExecute(0, "open", lstLinks.Text, 0, 0, 1)
End Sub
