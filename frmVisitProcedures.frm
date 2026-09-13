VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmVisitProcedures 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visit Procedures"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
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
   ScaleHeight     =   6495
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDeleteID 
      Height          =   360
      Left            =   5520
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtEditID 
      Height          =   360
      Left            =   5520
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTotalValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5400
      Width           =   3735
   End
   Begin VB.TextBox txtTotalCost 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4920
      Width           =   3735
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin MSFlexGridLib.MSFlexGrid gridPx 
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      _Version        =   393216
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtCost 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin MSDataListLib.DataCombo cmbPx 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
      Left            =   5520
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
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
   Begin btButtonEx.ButtonEx btnTT 
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&TT Print"
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
   Begin btButtonEx.ButtonEx btnDMPA 
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&DMPA Print"
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
   Begin VB.Label Label5 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total Cost"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Value"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "&Cost"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Procedure"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmVisitProcedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim FormPatient As New clsPatient
    Dim FormVisit As New clsVisit
    
    
Private Sub btnAdd_Click()
    If IsNumeric(cmbPx.BoundText) = False Then
        MsgBox "Please select a procedure"
        cmbPx.SetFocus
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtEditID.Text) <> 0 Then
            temSql = "Select * from tblVisitProcedure where VisitProcedureID = " & Val(txtEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !ProcedureID = Val(cmbPx.BoundText)
                !Cost = Val(txtCost.Text)
                !Value = Val(txtValue.Text)
                .Update
            Else
                .AddNew
                !VisitID = FormVisit.VisitID
                !PatientID = FormPatient.ID
                !ProcedureID = Val(cmbPx.BoundText)
                !Cost = Val(txtCost.Text)
                !Value = Val(txtValue.Text)
                .Update
            End If
        Else
            temSql = "Select * from tblVisitProcedure"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !VisitID = FormVisit.VisitID
            !PatientID = FormPatient.ID
            !ProcedureID = Val(cmbPx.BoundText)
            !Cost = Val(txtCost.Text)
            !Value = Val(txtValue.Text)
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    cmbPx.Text = Empty
    txtCost.Text = Empty
    txtValue.Text = Empty
    cmbPx.SetFocus


End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If IsNumeric(txtDeleteID.Text) = False Then
        MsgBox "Please select a procedure to delete"
        Exit Sub
    End If
    Dim rsPx As New ADODB.Recordset
    With rsPx
        If Val(txtDeleteID.Text) <> 0 Then
            temSql = "Select * from tblVisitProcedure where VisitProcedureID = " & Val(txtDeleteID.Text)
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedTime = Now
                !DeletedUserID = UserID
                .Update
            End If
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    cmbPx.Text = Empty
    txtCost.Text = Empty
    txtValue.Text = Empty
    cmbPx.SetFocus
    txtEditID.Text = Empty
    txtDeleteID.Text = Empty
    
End Sub

Private Sub btnDMPA_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    frmDMPAShedule.Show
    frmDMPAShedule.ZOrder
End Sub

Private Sub btnTT_Click()
    If CurrentPatient.ID <> FormPatient.ID Then CurrentPatient.ID = FormPatient.ID
    If CurrentVisit.VisitID <> FormVisit.VisitID Then CurrentVisit.VisitID = FormVisit.VisitID
    frmTetenusToxoidShedule.Show
    frmTetenusToxoidShedule.ZOrder 0
End Sub

Private Sub cmbPx_Change()
    Dim rsPx As New ADODB.Recordset
    With rsPx
        If .State = 1 Then .Close
        temSql = "Select * from tblProcedure where ProcedureID = " & Val(cmbPx.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtValue.Text = Format(!TotalValue, "0.00")
            txtCost.Text = Format(!TotalCost, "0.00")
        Else
            txtValue.Text = "0.00"
            txtCost.Text = "0.00"
        End If
        .Close
    End With
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
    GetCommonSettings Me
    
    FormPatient.ID = CurrentVisit.PatientID
    FormVisit.VisitID = CurrentVisit.VisitID
    
    If FormVisit.VisitDate <> Date Then
        btnDelete.Enabled = False
        btnAdd.Enabled = False
    End If
    
    
    Call SetColours
    Call FillCombos
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FillCombos()
    Dim Px As New clsFillCombos
    Px.FillAnyCombo cmbPx, "Procedure", True
End Sub

Private Sub FillGrid()
    Dim rsPx As New ADODB.Recordset
    Dim TotalCost As Double
    Dim TotalValue As Double
    With rsPx
        If .State = 1 Then .Close
        temSql = "Select * FROM tblProcedure RIGHT JOIN tblVisitProcedure ON tblProcedure.ProcedureID = tblVisitProcedure.ProcedureID where tblVisitProcedure.Deleted = False AND VisitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            gridPx.Rows = gridPx.Rows + 1
            gridPx.Row = gridPx.Rows - 1
            gridPx.Col = 0
            gridPx.Text = !Procedure
            gridPx.Col = 1
            gridPx.Text = Format(!Cost, "0.00")
            TotalCost = TotalCost + !Cost
            gridPx.Col = 2
            gridPx.Text = Format(!Value, "0.00")
            TotalValue = TotalValue + !Value
            gridPx.Col = 3
            gridPx.Text = !VisitProcedureID
            gridPx.Col = 4
            gridPx.Text = ![tblProcedure.ProcedureID]
            
            .MoveNext
        Wend
        .Close
    End With
    txtTotalCost.Text = Format(TotalCost, "0.00")
    txtTotalValue.Text = Format(TotalValue, "0.00")
End Sub

Private Sub FormatGrid()
    With gridPx
        .Clear
        .Rows = 1
        .Cols = 5
        
        .Row = 0
        
        .Col = 0
        .Text = "Procedure"
        
        .Col = 1
        .Text = "Cost"
        
        .Col = 2
        .Text = "Value"
        
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(0) = .Width - 2600
        .ColWidth(3) = 0
        .ColWidth(4) = 0
    End With
End Sub

Private Sub SaveDetails()
    Dim rsVisit As New ADODB.Recordset
    With rsVisit
        If .State = 1 Then .Close
        temSql = "Select * from tblVisit where visitID = " & FormVisit.VisitID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ProcedureFee = Val(txtTotalValue.Text)
            !ProcedureCost = Val(txtTotalCost.Text)
            !TotalFee = !ProcedureFee + !ConsultationFee + !MedicineFee
            !TotalCost = !ProcedureCost + !MedicineCost + !ConsultationCost
            .Update
        End If
        .Close
    End With
    
    Dim MyfrmNewVisit As Form
    
    For Each MyfrmNewVisit In Forms
        If MyfrmNewVisit.Name = "frmNewVisit" Then
            MyfrmNewVisit.DisplayVisitCharges
        End If
    Next
    
'    Call frmNewVisit.DisplayVisitCharges
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveDetails
    SaveCommonSettings Me
    
    
End Sub

Private Sub gridPx_Click()
    Dim temGridRow As Long
    With gridPx
        temGridRow = .Row
        .Col = 3
        txtDeleteID.Text = .Text
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub gridPx_DblClick()
    Dim temGridRow As Long
    With gridPx
        temGridRow = .Row
        .Col = 3
        txtEditID.Text = .Text
        .ColSel = .Cols - 1
        .Col = 4
        cmbPx.BoundText = Val(.Text)
        .Col = 2
        txtCost.Text = .Text
        .Col = 3
        txtValue.Text = .Text
        
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

