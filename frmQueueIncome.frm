VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmQueueIncome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Queue Income"
   ClientHeight    =   8355
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
   ScaleHeight     =   8355
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   7320
      Width           =   3735
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   7920
      Width           =   3735
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   7800
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
   Begin VB.TextBox txtTotalIncome 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5160
      TabIndex        =   8
      Top             =   7080
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GridIncome 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20774915
      CurrentDate     =   39859
   End
   Begin MSDataListLib.DataCombo cmbSecession 
      Height          =   360
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDoctor 
      Height          =   360
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label9 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Total Income"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Secession"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmQueueIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
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
    Dim i As Integer
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbDoctor_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbDoctor.BoundText = Empty
    End If
End Sub

Private Sub cmbSecession_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSecession_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbSecession.Text = Empty
    End If
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

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call SetColours
    Call FillCombos
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub

Private Sub FillCombos()
    Dim Doc As New clsFillCombos
    Doc.FillBoolCombo cmbDoctor, "staff", "staff", "IsADoctor", True
    Dim Secession As New clsFillCombos
    Secession.FillBoolCombo cmbSecession, "Secession", "Secession", "IsVisitSecession", True
End Sub

Private Sub FormatGrid()
    With GridIncome
        .Clear
        .Cols = 6
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "No."
        
        .Col = 1
        .Text = "Visit ID"
        
        .Col = 2
        .Text = "Time"
        
        .Col = 3
        .Text = "Patient"
        
        .Col = 4
        .Text = "Fee"
        
        .Col = 5
        .Text = "PatientID"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 700
        .ColWidth(2) = 1000
        .ColWidth(3) = 3430
        .ColWidth(4) = 1000
        .ColWidth(5) = 0
        
    End With

End Sub

Private Sub FillGrid()
'    If IsNumeric(cmbDoctor.BoundText) = False Then Exit Sub
'    If IsNumeric(cmbSecession.BoundText) = False Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    Dim TotalFee As Double
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbDoctor.BoundText) = True And IsNumeric(cmbSecession.BoundText) = True Then
            temSql = "SELECT  tblPatient.PatientID, tblQueue.QueueNo, tblQueue.Removed, tblPatient.Patient, tblVisit.VisitID,  tblVisit.VisitTime, tblVisit.TotalFee " & _
                        "FROM (tblVisit RIGHT JOIN tblQueue ON tblVisit.QueueID = tblQueue.QueueID) LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID " & _
                        "WHERE (((tblVisit.Deleted)=False) AND ((tblQueue.DoctorID)=" & Val(cmbDoctor.BoundText) & ") AND ((tblVisit.Deleted)=False) AND ((tblQueue.QueueDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#) AND ((tblQueue.SecessionID)= " & Val(cmbSecession.BoundText) & "))"
        ElseIf IsNumeric(cmbDoctor.BoundText) = False And IsNumeric(cmbSecession.BoundText) = True Then
            temSql = "SELECT  tblPatient.PatientID, tblQueue.QueueNo, tblQueue.Removed, tblPatient.Patient, tblVisit.VisitID,   tblVisit.VisitTime, tblVisit.TotalFee " & _
                        "FROM (tblVisit RIGHT JOIN tblQueue ON tblVisit.QueueID = tblQueue.QueueID) LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID " & _
                        "WHERE (((tblVisit.Deleted)=False) AND ((tblQueue.QueueDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#) AND ((tblQueue.SecessionID)= " & Val(cmbSecession.BoundText) & "))"
        
        ElseIf IsNumeric(cmbDoctor.BoundText) = True And IsNumeric(cmbSecession.BoundText) = False Then
            temSql = "SELECT  tblPatient.PatientID, tblQueue.QueueNo, tblQueue.Removed, tblPatient.Patient, tblVisit.VisitID,   tblVisit.VisitTime, tblVisit.TotalFee " & _
                        "FROM (tblVisit RIGHT JOIN tblQueue ON tblVisit.QueueID = tblQueue.QueueID) LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID " & _
                        "WHERE (((tblVisit.Deleted)=False) AND ((tblQueue.DoctorID)=" & Val(cmbDoctor.BoundText) & ") AND ((tblVisit.Deleted)=False) AND ((tblQueue.QueueDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#))"
        
        ElseIf IsNumeric(cmbDoctor.BoundText) = False And IsNumeric(cmbSecession.BoundText) = False Then
            temSql = "SELECT  tblPatient.PatientID, tblQueue.QueueNo, tblQueue.Removed, tblPatient.Patient, tblVisit.VisitID,   tblVisit.VisitTime, tblVisit.TotalFee " & _
                        "FROM (tblVisit RIGHT JOIN tblQueue ON tblVisit.QueueID = tblQueue.QueueID) LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID " & _
                        "WHERE (((tblVisit.Deleted)=False) AND ((tblQueue.QueueDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#))"
        
        Else
            Exit Sub
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If IsNull(!TotalFee) = False Then
                TotalFee = TotalFee + !TotalFee
            End If
            GridIncome.Rows = GridIncome.Rows + 1
            GridIncome.Row = GridIncome.Rows - 1
            
            GridIncome.Col = 0
            GridIncome.Text = !QueueNo
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            
            GridIncome.Col = 1
            If IsNull(!VisitID) = False Then
                GridIncome.Text = ![VisitID]
            End If
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 2
            If IsNull(!VisitTime) = False Then
                GridIncome.Text = Format(![VisitTime], "HH:MM AMPM")
            End If
            
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 3
            If IsNull(!Patient) = False Then
                GridIncome.Text = !Patient
            Else
                If !Removed = True Then
                    GridIncome.Text = "Removed"
                Else
                    GridIncome.Text = "Not seen yet"
                End If
            End If
            
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 4
            GridIncome.Text = Format(!TotalFee, "0.00")
            
            
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 5
            If IsNull(!PatientID) = False Then
                GridIncome.Text = ![PatientID]
            End If
            
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            
            .MoveNext
        Wend
    End With

txtTotalIncome.Text = Format(TotalFee, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbSecession.Name, cmbSecession.BoundText
    SaveSetting App.EXEName, Me.Name, cmbDoctor.Name, cmbDoctor.BoundText
    SaveCommonSettings Me
    
End Sub

Private Sub GetSettings()
    dtpDate.Value = Date
    GetCommonSettings Me
    
    cmbSecession.BoundText = GetSetting(App.EXEName, Me.Name, cmbSecession.Name, 1)
    cmbDoctor.BoundText = GetSetting(App.EXEName, Me.Name, cmbDoctor.Name, 1)
End Sub

Private Sub GridIncome_Click()
    With GridIncome
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub GridIncome_DblClick()
    Dim PtID As Long
    Dim VtID As Long
    With GridIncome
        PtID = .TextMatrix(.Row, 5)
        VtID = .TextMatrix(.Row, 2)
                
        If SingleVisit = True Then
            Unload frmPatient
            Unload frmPatientDetails
            Unload frmNewVisit
            CurrentPatient.ID = PtID
            CurrentVisit.VisitID = VtID
            frmNewVisit.Show
        Else
            Dim MyNewVisit As New frmNewVisit
            CurrentPatient.ID = PtID
            CurrentVisit.VisitID = VtID
            MyNewVisit.Show
            MyNewVisit.ZOrder 0
        End If
        

    End With
End Sub
