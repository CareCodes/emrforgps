VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVisitImages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visit Images"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
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
   ScaleHeight     =   8760
   ScaleWidth      =   11310
   Begin VB.PictureBox pctTem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1875
      Left            =   3360
      ScaleHeight     =   1815
      ScaleWidth      =   1860
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox pctImage 
      AutoRedraw      =   -1  'True
      Height          =   5835
      Left            =   3360
      ScaleHeight     =   5775
      ScaleWidth      =   7740
      TabIndex        =   18
      Top             =   120
      Width           =   7800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkLoaded 
      Caption         =   "Check3"
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtImagePath 
      Height          =   360
      Left            =   3360
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSDataListLib.DataCombo cmbImage 
      Height          =   7620
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   13441
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin VB.TextBox txtComments 
      Height          =   885
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6960
      Width           =   5535
   End
   Begin VB.CheckBox chkVisitDefault 
      Caption         =   "Visit's Default Image"
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CheckBox chkPatientDefault 
      Caption         =   "Patient's Default Image"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   7920
      Width           =   2655
   End
   Begin VB.TextBox txtImage 
      Height          =   405
      Left            =   5640
      TabIndex        =   0
      Top             =   6480
      Width           =   5535
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   8280
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
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   8280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   8280
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   8280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9960
      TabIndex        =   10
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
   Begin btButtonEx.ButtonEx btnRemove 
      Height          =   375
      Left            =   9960
      TabIndex        =   14
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Remove"
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
   Begin btButtonEx.ButtonEx btnLoad 
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Load"
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
      Caption         =   "Comments"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Image"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Images"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmVisitImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsVisit As New ADODB.Recordset
    Dim rsImage As New ADODB.Recordset
    Dim FSys As New FileSystemObject
    
    Dim FormPatient As New clsPatient
    Dim FormVisit As New clsVisit

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

Private Sub Form_Activate()
    If IsNumeric(cmbImage.BoundText) = False Then
        cmbImage.SetFocus
        SendKeys "{down}"
    End If
End Sub

Private Sub btnAdd_Click()
    Call ClearValues
    cmbImage.Text = Empty
    Call EditMode
    btnLoad.Enabled = True
    btnRemove.Enabled = True
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbImage.SetFocus
    cmbImage.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If IsNumeric(cmbImage.BoundText) = False Then
        MsgBox "Noting to Delete"
        cmbImage.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete the Photo?", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblImage where ImageID = " & Val(cmbImage.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedTime = Now
            .Update
        End If
        .Close
    End With
    Call ClearValues
    Call FillCombos
    Call SelectMode
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbImage.BoundText) = False Then
        MsgBox "Nothing to Edit"
        cmbImage.SetFocus
        Exit Sub
    End If
    Call EditMode
    btnLoad.Enabled = False
    btnRemove.Enabled = False
    txtImage.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnLoad_Click()
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "Bitmap|*.BMP|JPEG|*.JPG;*.JPE;*.JPEG|GIF|*.GIF|All Supported Images|*.BMP;*.JPG;*.JPE;*.JPEG;*.GIF"
    On Error GoTo eh
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtImagePath.Text = CommonDialog1.filename
        Set pctImage.Picture = LoadPicture(txtImagePath.Text)
        Set pctTem.Picture = LoadPicture(txtImagePath.Text)
        chkLoaded.Value = 1
    Else
        MsgBox "You have not selected valid image.", vbCritical
        Set pctImage.Picture = LoadPicture()
        Set pctTem.Picture = LoadPicture()
        txtImagePath.Text = Empty
        chkLoaded.Value = 0
    End If
    Exit Sub
eh:
    MsgBox "Error loading the image"

End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbImage.Enabled = True
    
    txtImage.Enabled = False
    txtComments.Enabled = False
    chkPatientDefault.Enabled = False
    chkVisitDefault.Enabled = False
    
    btnLoad.Enabled = False
    btnRemove.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
    
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbImage.Enabled = False
    
    txtImage.Enabled = True
    txtComments.Enabled = True
    chkPatientDefault.Enabled = True
    chkVisitDefault.Enabled = True
    
    btnLoad.Enabled = False
    btnRemove.Enabled = False
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
End Sub

Private Sub DisplayDetails()
    Call ClearValues
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblImage where ImageID = " & Val(cmbImage.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtImage.Text = !Image
            txtComments.Text = Format(!Comments, "")
            If !PatientDefault = True Then chkPatientDefault.Value = 1
            If !VisitDefault = True Then chkVisitDefault.Value = 1
            pctTem.AutoRedraw = True
            pctTem.AutoSize = True
            Call FillPhoto(rsTem, "ImageData", "ImageSize", Me.pctTem)
            pctImage.AutoRedraw = True
            pctImage.AutoSize = False
            pctImage.PaintPicture pctTem.Image, 0, 0, pctImage.Width, pctImage.Height
        End If
    End With
    
End Sub

Private Sub ClearValues()
            txtImage.Text = Empty
            txtComments.Text = Empty
            chkPatientDefault.Value = 0
            chkVisitDefault.Value = 0
            Set pctImage.Picture = LoadPicture()
End Sub

Private Sub btnRemove_Click()
    Set pctImage.Picture = LoadPicture()
    txtImagePath.Text = Empty
    chkLoaded.Value = 0
End Sub

Private Sub btnSave_Click()
    If IsNumeric(cmbImage.BoundText) = False Then
        If FSys.FileExists(txtImagePath.Text) = False Then
            MsgBox "Plase select a photo to save"
            btnLoad_Click
            Exit Sub
        End If
    End If
    If Trim(txtImage.Text) = Empty Then
        txtImage.Text = "A photo of " & FormPatient.NameWithTitle & " of Visit on " & Format(FormVisit.VisitDate, LongDateFormat)
    End If
    If IsNumeric(cmbImage.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call ClearValues
    Call SelectMode
    Call FillCombos
    cmbImage.SetFocus
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
    
        If chkPatientDefault.Value = 1 Then
            If .State = 1 Then .Close
            temSql = "Select * from tblImage where PatientID = " & FormPatient.ID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            While .EOF = False
                !PatientDefault = False
                .Update
                .MoveNext
            Wend
            .Close
        End If
        If chkVisitDefault.Value = 1 Then
            If .State = 1 Then .Close
            temSql = "Select * from tblImage where VisitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            While .EOF = False
                !VisitDefault = False
                .Update
                .MoveNext
            Wend
            .Close
        End If
        If .State = 1 Then .Close
        temSql = "Select * from tblImage"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Image = txtImage.Text
        !Comments = txtComments.Text
        !PatientID = FormPatient.ID
        !VisitID = FormVisit.VisitID
        If chkPatientDefault.Value = 1 Then
            !PatientDefault = True
        Else
            !PatientDefault = False
        End If
        If chkVisitDefault.Value = 1 Then
            !VisitDefault = True
        Else
            !VisitDefault = False
        End If
        GetPhoto txtImagePath.Text, rsTem, "ImageData", "ImageSize"
        .Update
        .Close
    End With
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If chkPatientDefault.Value = 1 Then
            If .State = 1 Then .Close
            temSql = "Select * from tblImage where PatientID = " & FormPatient.ID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            While .EOF = False
                !PatientDefault = False
                .Update
                .MoveNext
            Wend
            .Close
        End If
        If chkVisitDefault.Value = 1 Then
            If .State = 1 Then .Close
            temSql = "Select * from tblImage where VisitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            While .EOF = False
                !VisitDefault = False
                .Update
                .MoveNext
            Wend
            .Close
        End If
        If .State = 1 Then .Close
        temSql = "Select * from tblImage Where ImageID = " & Val(cmbImage.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Image = txtImage.Text
            !Comments = txtComments.Text
            If chkPatientDefault.Value = 1 Then
                !PatientDefault = True
            Else
                !PatientDefault = False
            End If
            If chkVisitDefault.Value = 1 Then
                !VisitDefault = True
            Else
                !VisitDefault = False
            End If
            .Update
        End If
        .Close
    End With
End Sub

Private Sub cmbImage_Change()
    Call DisplayDetails
End Sub

Private Sub cmbImage_Click(Area As Integer)
    Call cmbImage_Change
End Sub

Private Sub Form_Load()
    FormPatient.ID = CurrentPatient.ID
    FormVisit.VisitID = CurrentVisit.VisitID
    Call FillCombos
    Call SelectMode
    Call ClearValues
    Call SetColours
End Sub

Private Sub FillCombos()
    Dim Img As New clsFillCombos
    Img.FillLongCombo cmbImage, "Image", "Image", "VisitID", FormVisit.VisitID
End Sub

Private Sub pctImage_DblClick()
    Dim MyFrm As Form
    Set MyFrm = New frmFullVisitImage
    MyFrm.Show
    MyFrm.Caption = cmbImage.Text
    MyFrm.ZOrder 0
End Sub
