VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCertificates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificates"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
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
   ScaleHeight     =   9450
   ScaleWidth      =   10500
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   102957059
      CurrentDate     =   39890
   End
   Begin VB.TextBox txtScaleWidth 
      Height          =   360
      Left            =   1200
      TabIndex        =   24
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox txtScaleHeight 
      Height          =   360
      Left            =   1200
      TabIndex        =   21
      Top             =   7680
      Width           =   1575
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   7200
      Width           =   3735
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   6480
      Width           =   3735
   End
   Begin MSDataListLib.DataList lstReferrals 
      Height          =   3180
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5609
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnCreate 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&reate"
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
   Begin MSDataListLib.DataCombo cmbReferring 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin RichTextLib.RichTextBox rtbReferrals 
      Height          =   8295
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   14631
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmCertificates.frx":0000
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   9000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Print"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9000
      TabIndex        =   13
      Top             =   9000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   9000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin btButtonEx.ButtonEx btnSpellCheck 
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Spe&ll Check"
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
   Begin btButtonEx.ButtonEx btnThesarus 
      Height          =   375
      Left            =   9000
      TabIndex        =   10
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Thesaurus"
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
   Begin btButtonEx.ButtonEx btnClear 
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   8520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Clear"
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   87228419
      CurrentDate     =   39890
   End
   Begin VB.Label Label9 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "inches"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "inches"
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Height"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Width"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Referring to :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmCertificates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NumForms As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    Dim FSys As New Scripting.FileSystemObject
    Private CsetPrinter As New cSetDfltPrinter
    Dim temSql As String
    
    Dim rsRfList As New ADODB.Recordset
    
    Dim i As Integer

    Dim FormPatient As New clsPatient
    Dim FormVisit As New clsVisit

Private Function Address(DoctorID As Long) As String
    Address = Empty
    Dim rsDOc As New ADODB.Recordset
    With rsDOc
        If .State = 1 Then .Close
        temSql = "SELECT tblStaff.HowToAddress FROM tblStaff WHERE (((tblStaff.StaffID)= " & DoctorID & " ))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Address = Address & Format(!HowToAddress, "")
        End If
    End With
End Function

Private Function DocDetails(DoctorID As Long) As String
    DocDetails = Empty
    Dim rsDOc As New ADODB.Recordset
    With rsDOc
        If .State = 1 Then .Close
        temSql = "SELECT tblTitle.Title, tblStaff.Staff, tblStaff.Designation, tblStaff.Address " & _
                    "FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID " & _
                    "WHERE (((tblStaff.StaffID)= " & DoctorID & " ))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Format(!Title, "") = "" Then
                DocDetails = Format(!Staff, "")
            Else
                DocDetails = Format(!Title, "") & " " & Format(!Staff, "")
            End If
            'DocDetails = DocDetails & vbNewLine
            If Format(!Designation, "") <> "" Then
                DocDetails = DocDetails & vbNewLine & Format(!Designation, "")
            End If
            'DocDetails = DocDetails & vbNewLine
            If Format(!Address, "") <> "" Then
                DocDetails = DocDetails & vbNewLine & Format(!Address, "")
            End If
        End If
    End With
    
End Function

Private Sub btnClear_Click()
    rtbReferrals.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnCreate_Click()
    With rtbReferrals
        .Text = Empty
        .Text = .Text & Format(Date, LongDateFormat) & vbNewLine & vbNewLine
        .Text = .Text & "To whom it may concern:" & vbNewLine & vbNewLine
        .Text = .Text & vbNewLine & "Re: " & cmbReferring.Text & " - " & FormPatient.NameWithTitle & vbNewLine & vbNewLine
        If Trim(PatientDetails) <> "" Then
            .Text = .Text & PatientDetails & vbNewLine & vbNewLine
        End If
        If Trim(PastHistory) <> "" Then
            .Text = .Text & PastHistory & vbNewLine & vbNewLine
        End If
        If Trim(DrugHistory) <> "" Then
            .Text = .Text & DrugHistory & vbNewLine & vbNewLine
        End If
        If Trim(PresentHistory) <> "" Then
            .Text = .Text & PresentHistory & vbNewLine & vbNewLine
        End If
        If Trim(NewIx) <> "" Then
            .Text = .Text & NewIx & vbNewLine & vbNewLine
        End If
        
        If Trim(ExpectedAction) <> "" Then
            .Text = .Text & ExpectedAction & vbNewLine & vbNewLine
        End If
        .Text = .Text & "Thank you," & vbNewLine & "Yours truly" & vbNewLine & vbNewLine & vbNewLine
        .Text = .Text & "..................................." & vbNewLine
        .Text = .Text & DocDetails(currentDoctorId)
    End With
End Sub

Private Function ExpectedAction() As String
    ExpectedAction = Empty
    If IsNumeric(cmbReferring.BoundText) = False Then
        MsgBox "Please select a type of document"
        cmbReferring.SetFocus
        Exit Function
    End If
    If cmbReferring.BoundText = 4 Then
        ExpectedAction = "I recommended medical leave for "
        If UCase(FormPatient.Sex) = UCase("male") Then
            ExpectedAction = ExpectedAction & "him "
        ElseIf UCase(FormPatient.Sex) = UCase("female") Then
            ExpectedAction = ExpectedAction & "her "
        Else
            ExpectedAction = ExpectedAction & "the patient "
        End If
        If DateDiff("d", dtpFrom.Value, dtpTo.Value) < 2 Then
            ExpectedAction = ExpectedAction & " for one day on " & Format(dtpFrom.Value, LongDateFormat) & "."
        Else
            ExpectedAction = ExpectedAction & " for " & DateDiff("d", dtpFrom.Value, dtpTo.Value) + 1 & " days from " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat) & "."
        End If
    ElseIf cmbReferring.BoundText = 5 Then
        ExpectedAction = "I charged "
        If UCase(FormPatient.Sex) = UCase("male") Then
            ExpectedAction = ExpectedAction & "him "
        ElseIf UCase(FormPatient.Sex) = UCase("female") Then
            ExpectedAction = ExpectedAction & "her "
        Else
            ExpectedAction = ExpectedAction & "the patient "
        End If
        ExpectedAction = ExpectedAction & " Rs. " & Format(FormVisit.TotalFee, "0.00") & "."
        ExpectedAction = ExpectedAction & vbNewLine & vbTab & "Consultation Fee : " & Format(FormVisit.ConsultationFee, "0.00")
        
        ExpectedAction = ExpectedAction & vbNewLine & vbTab & "Medicine cost : " & Format(FormVisit.MedicinesFee, "0.00")
        
        
    ElseIf cmbReferring.BoundText = 6 Then
        ExpectedAction = "I certify that "
        If UCase(FormPatient.Sex) = UCase("male") Then
            ExpectedAction = ExpectedAction & "he "
        ElseIf UCase(FormPatient.Sex) = UCase("female") Then
            ExpectedAction = ExpectedAction & "she "
        Else
            ExpectedAction = ExpectedAction & "this person "
        End If
        ExpectedAction = ExpectedAction & " is physically and mentally fit enough to carry out any ordinary task assigned to him."
    ElseIf cmbReferring.BoundText = 7 Then
        ExpectedAction = "Please be kind enough to look in to "
        If UCase(FormPatient.Sex) = UCase("male") Then
            ExpectedAction = ExpectedAction & "his "
        ElseIf UCase(FormPatient.Sex) = UCase("female") Then
            ExpectedAction = ExpectedAction & "her "
        Else
            ExpectedAction = ExpectedAction & "this person's health status "
        End If
        ExpectedAction = ExpectedAction & " and do the needful."
    End If
End Function

Private Function PresentHistory() As String
    PresentHistory = Empty
    Dim DxCount As Long
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblVisitDiagnosis.VisitDiagnosisID, tblVisitDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblVisitDiagnosis ON tblDiagnosis.DiagnosisID = tblVisitDiagnosis.DiagnosisID where tblVisitDiagnosis.Deleted = False AND tblVisitDiagnosis.VisitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            PresentHistory = "This time "
            If UCase(FormPatient.Sex) = UCase("male") Then
                PresentHistory = PresentHistory & "he "
            ElseIf UCase(FormPatient.Sex) = UCase("female") Then
                PresentHistory = PresentHistory & "she "
            Else
                PresentHistory = PresentHistory & "patient "
            End If
            PresentHistory = PresentHistory & "came to me with "
            .MoveLast
            DxCount = .RecordCount
            .MoveFirst
            For i = 0 To DxCount - 1
                If i = 0 Then
                    PresentHistory = PresentHistory & !Diagnosis
                ElseIf i = DxCount - 1 Then
                    PresentHistory = PresentHistory & " and " & !Diagnosis
                Else
                    PresentHistory = PresentHistory & ", " & !Diagnosis
                End If
                If !Comments <> "" Then
                    PresentHistory = PresentHistory & " (" & !Comments & ")"
                End If
                .MoveNext
            Next
            PresentHistory = PresentHistory & "."
        End If
    End With
    Set rsDx = Nothing
End Function

Private Function NewIx() As String
    NewIx = Empty
    Dim DxCount As Long
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientIX.PatientIXID, tblIx.Ix, tblIx.IxID FROM tblIx RIGHT JOIN tblPatientIX ON tblIx.IxID = tblPatientIX.IxID where tblPatientIX.Deleted = False AND tblPatientIX.VisitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            NewIx = "I have requested "
            .MoveLast
            DxCount = .RecordCount
            .MoveFirst
            For i = 0 To DxCount - 1
                If i = 0 Then
                    NewIx = NewIx & !Ix
                ElseIf i = DxCount - 1 Then
                    NewIx = NewIx & " and " & !Ix
                Else
                    NewIx = NewIx & ", " & !Ix
                End If
                .MoveNext
            Next
            NewIx = NewIx & "."
        End If
        If .State = 1 Then .Close
        temSql = "SELECT tblIx.Ix, tblIx.IxID, tblPatientIx.Deleted FROM tblPatientIx LEFT JOIN tblIx ON tblPatientIx.IxID = tblIx.IxID WHERE (((tblPatientIx.VisitID)= " & FormVisit.VisitID & ") AND ((tblPatientIx.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            DxCount = .RecordCount
            If DxCount = 1 Then
                NewIx = NewIx & " I have attached the results of "
            Else
                NewIx = NewIx & " I have attached the result of "
            End If
            .MoveFirst
            For i = 0 To DxCount - 1
                If i = 0 Then
                    NewIx = NewIx & !Ix
                ElseIf i = DxCount - 1 Then
                    NewIx = NewIx & " and " & !Ix
                Else
                    NewIx = NewIx & ", " & !Ix
                End If
                .MoveNext
            Next
            NewIx = NewIx & "."
        End If
        'SELECT tblIx.Ix, tblIx.IxID, tblPatientIx.Deleted FROM tblPatientIx LEFT JOIN tblIx ON tblPatientIx.IxID = tblIx.IxID WHERE (((tblPatientIx.VisitID)=137) AND ((tblPatientIx.Deleted)=False))
    End With
    Set rsDx = Nothing
End Function

Private Function PastHistory() As String
    PastHistory = Empty
    Dim DxCount As Long
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientDiagnosis.PatientDiagnosisID, tblPatientDiagnosis.Comments, tblDiagnosis.Diagnosis, tblDiagnosis.DiagnosisID FROM tblDiagnosis RIGHT JOIN tblPatientDiagnosis ON tblDiagnosis.DiagnosisID = tblPatientDiagnosis.DiagnosisID where tblPatientDiagnosis.Deleted = False AND tblPatientDiagnosis.PatientID = " & FormPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            DxCount = .RecordCount
            .MoveFirst
            PastHistory = "The patient was diagnosed to be having "
            For i = 0 To DxCount - 1
                If i = 0 Then
                    PastHistory = PastHistory & !Diagnosis
                ElseIf i = DxCount - 1 Then
                    PastHistory = PastHistory & " and " & !Diagnosis
                Else
                    PastHistory = PastHistory & ", " & !Diagnosis
                End If
                If !Comments <> "" Then
                    PastHistory = PastHistory & " (" & !Comments & ")"
                End If
                .MoveNext
            Next
            PastHistory = PastHistory & "."
        End If
    End With
    Set rsDx = Nothing
End Function

Private Function DrugHistory() As String
    DrugHistory = Empty
    Dim ItemCount As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientItem.PatientItemID, [tblItem.Item] & ' ' & [tblPatientItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] AS Prescreption " & _
                    "FROM (((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblPatientItem ON tblIssueUnit.ItemUnitID = tblPatientItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblPatientItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblPatientItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblPatientItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblPatientItem.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblPatientItem.PatientID)=" & FormPatient.ID & ") AND ((tblPatientItem.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            ItemCount = .RecordCount
            .MoveFirst
            If Trim(PastHistory) = "" Then
                DrugHistory = "This patient is on "
            Else
                If UCase(FormPatient.Sex) = UCase("male") Then
                    DrugHistory = "He is on "
                ElseIf UCase(FormPatient.Sex) = UCase("female") Then
                    DrugHistory = "She is on "
                Else
                    DrugHistory = "Patient is on "
                End If
            End If
            For i = 0 To ItemCount - 1
                If i = 0 Then
                    DrugHistory = DrugHistory & !Prescreption
                ElseIf i = ItemCount - 1 Then
                    DrugHistory = DrugHistory & " and " & !Prescreption
                Else
                    DrugHistory = DrugHistory & ", " & !Prescreption
                End If
                .MoveNext
            Next
            DrugHistory = DrugHistory & "."
        End If
    End With
End Function

Private Function PatientDetails() As String
    PatientDetails = Empty
        PatientDetails = FormPatient.NameWithTitle & " is a " & FormPatient.AgeInWords & " old " & FormPatient.Sex
        If FormPatient.Address <> "" Then
            PatientDetails = PatientDetails & " from " & FormPatient.Address & "."
        Else
            PatientDetails = PatientDetails & "."
        End If
End Function

Private Sub btnDelete_Click()
    i = MsgBox("Are you sure you want to delete the referral", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsR As New ADODB.Recordset
    With rsR
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitCertificate where VisitCertificateID = " & Val(lstReferrals.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
            MsgBox "Deleted"
            Call FillReferrals
        Else
            MsgBox "Couldn't delete"
        End If
    End With
End Sub

Private Sub btnPrint_Click()
    Dim temText As String
    Dim rstemReportItems As New ADODB.Recordset
    
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    If SelectForm(cmbPaper.Text, Me.hwnd) <> 1 Then
        MsgBox "Printer Error"
        Exit Sub
    End If
    
'    Printer.ScaleWidth = Val(txtScaleWidth.Text) * 1440
'    Printer.ScaleHeight = Val(txtScaleHeight.Text) * 1440
    
    
    Printer.FillStyle = vbFSTransparent
    
    With rstemReportItems
        If .State = 1 Then .Close
        temSql = "Select * from tblPrintItem  where Deleted = false AND PrintItemCategoryID = 3"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            TemX1 = Printer.ScaleWidth * !X1
            TemX2 = Printer.ScaleWidth * !X2
            TemY1 = Printer.ScaleHeight * !Y1
            TemY2 = Printer.ScaleHeight * !Y2
            If IsNull(!CircleRadius) = False Then
                temRadius = Printer.ScaleWidth * !CircleRadius
            End If
            If !IsLine = True Then
                Printer.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour
            ElseIf !IsRectangle = True Then
                Printer.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour, B
            ElseIf !IsCircle = True Then
                Printer.Circle (TemX1, TemY1), temRadius, !ForeColour
            ElseIf !IsLabel = True Then
                temText = Format(!ItemText, "")
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            ElseIf !IsText = True Then
                Select Case !ItemText
                    Case "Patient Name": temText = FormPatient.NameWithTitle
                    Case "Patient ID": temText = FormPatient.ID
                    Case "Patient Age in Words": temText = FormPatient.AgeInWords
                    Case "Patient Data of Birth": temText = FormPatient.DateOfBirth
                    Case "Patient Sex": temText = FormPatient.Sex
                    Case "Patient Civil Status": temText = ""
                    Case "Patient Address": temText = FormPatient.Address
                    Case "Patient Telephone": temText = FormPatient.Telephone
                    Case "Date (dd mm yy)": temText = Format(FormVisit.VisitDate, "dd mm yy")
                    Case "Date (dd mmm yy)": temText = Format(FormVisit.VisitDate, "dd mmm yy")
                    Case "Date (dd mmm yyyy)": temText = Format(FormVisit.VisitDate, "dd mmm yyyy")
                    Case "Date (dd mmmm yyyy)": temText = Format(FormVisit.VisitDate, "dd mmmm yyyy")
                    Case "Date (dd mm yyyy)": temText = Format(FormVisit.VisitDate, "dd mm yyy")
                    Case "Date (yy mm dd)": temText = Format(FormVisit.VisitDate, "yy mm dd")
                    Case "Date (yy mmm dd)": temText = Format(FormVisit.VisitDate, "yy mmm dd")
                    Case "Date (yyyy mmm dd)": temText = Format(FormVisit.VisitDate, "yyyy mmm dd")
                    Case "Date (yyyy mmmm dd)": temText = Format(FormVisit.VisitDate, "yyyy mmmm dd")
                    Case "Date (yyyy mm dd)": temText = Format(FormVisit.VisitDate, "yyyy mm dd")
                    Case "Short Time": temText = Format(FormVisit.VisitTime, "h:m AMPM")
                    Case "Long Time": temText = Format(FormVisit.VisitTime, "HH:MM AMPM")
                    Case "Document": temText = rtbReferrals.Text
                    Case Else
                        temText = ""
                End Select
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            End If
            .MoveNext
        Wend
        .Close
    End With
    Printer.EndDoc
End Sub


Private Sub btnSave_Click()
    Dim rsVR As New ADODB.Recordset
    With rsVR
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitCertificate"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !VisitCertificate = cmbReferring.Text
        !PatientID = FormPatient.ID
        !VisitID = FormVisit.VisitID
        !Comments = rtbReferrals.TextRTF
        !SavedDate = Date
        !SavedTime = Time
        .Update
    End With
    Call FillReferrals
End Sub

Private Sub FillReferrals()
    With rsRfList
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitCertificate where Deleted = False AND PatientID = " & FormPatient.ID & " order by VisitCertificateID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstReferrals
        Set .RowSource = rsRfList
        .ListField = "VisitCertificate"
        .BoundColumn = "VisitCertificateID"
    End With
End Sub

Private Sub btnSpellCheck_Click()
    Screen.MousePointer = vbHourglass
    Dim ReferralText As String
    ReferralText = rtbReferrals.Text
    rtbReferrals.Text = SpellChecker(ReferralText)
    MsgBox "Spell check Completed"
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnThesarus_Click()
    Dim objWord As Word.Application
    Dim objDoc  As Document
    Dim strText As String
    On Error GoTo NormalExit
    If rtbReferrals.SelText = "" Then
        MsgBox "Please select a word first"
        Exit Sub
    End If
    Set objWord = CreateObject("word.application")
    objWord.Visible = False
    objWord.WindowState = wdWindowStateMinimize
    Set objDoc = objWord.Documents.Add
    objDoc.Range = rtbReferrals.SelText
    objDoc.Range.CheckSynonyms
    strText = objDoc.Range
    strText = Left$(strText, Len(strText) - 1)
    rtbReferrals.SelText = strText

NormalExit:
    On Error Resume Next
    objDoc.Close SaveChanges:=wdDoNotSaveChanges
    objWord.Quit SaveChanges:=wdDoNotSaveChanges
    Set objWord = Nothing
    Set objDoc = Nothing

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
    
    
    FormPatient.ID = CurrentPatient.ID
    FormVisit.VisitID = CurrentVisit.VisitID
    
    Call FillReferrals
    Call FillCombos
    Call FillPrinters
    
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPrinter_Click
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
    txtScaleWidth.Text = Val(GetSetting(App.EXEName, Me.Name, "ScaleWidth", 4.1))
    txtScaleHeight.Text = Val(GetSetting(App.EXEName, Me.Name, "ScaleHeight", 6.3))
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub FillCombos()
    Dim AE As New clsFillCombos
    AE.FillBoolCombo cmbReferring, "FixedCategory", "FixedCategory", "IsCertificateCategory", True
End Sub


Private Sub cmbPrinter_Change()
    cmbPrinter_Click
End Sub

Private Sub cmbPrinter_Click()
    'cmbPaper.Clear
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = PrescreptionPaperHeight
            .cy = PrescreptionPaperWidth
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    
    
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
    SaveSetting App.EXEName, Me.Name, "ScaleHeight", Val(txtScaleHeight.Text)
    SaveSetting App.EXEName, Me.Name, "ScaleWidth", Val(txtScaleWidth.Text)
End Sub

Private Sub lstReferrals_Click()
    rtbReferrals.Text = Empty
    Dim rsR As New ADODB.Recordset
    With rsR
        If .State = 1 Then .Close
        temSql = "Select * from tblVisitCertificate where VisitCertificateID = " & Val(lstReferrals.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Comments) = False Then
                rtbReferrals.TextRTF = !Comments
            End If
        End If
    End With
End Sub
