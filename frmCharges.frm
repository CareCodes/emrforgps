VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCharges 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Charges"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
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
   ScaleHeight     =   4455
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3840
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
   Begin VB.TextBox txtProcedureFee 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtMedicinesFee 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtConsultationFee 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtTotalCharge 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label6 
      Caption         =   "Drawer"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "For Procedures"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "For Medicines"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultation Fee"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Charge"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "frmCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FormVisit As New clsVisit
    Dim FormPatient As New clsPatient
    Dim temSql As String
    
Private Sub FillCombos()
    Dim PeyMethod As New clsFillCombos
    PeyMethod.FillAnyCombo cmbPaymentMethod, "PaymentMethod", True
    Dim Drawer As New clsFillCombos
    Drawer.FillAnyCombo cmbDrawer, "Drawer", True
End Sub
    
Private Sub btnClose_Click()
    Call txtConsultationFee_LostFocus
    
    Call cmbPaymentMethod_LostFocus
    Call cmbDrawer_LostFocus
    If SingleVisit = True Then
        frmNewVisit.DisplayVisitCharges
    Else
        Dim MyForm As Form
        For Each MyForm In Forms
            If MyForm.Name = "frmNewVisit" Then
                If MyForm.FormVisit.VisitID = FormVisit.VisitID Then
                    MyForm.DisplayVisitCharges
                End If
            End If
        Next
    End If
    Unload Me
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

Private Sub cmbPaymentMethod_LostFocus()
    If Val(cmbPaymentMethod.BoundText) <> FormVisit.PaymentMethodID Then
        Dim rsVisit As New ADODB.Recordset
        With rsVisit
            If .State = 1 Then .Close
            temSql = "Select * from tblVisit where visitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                ChangeDrawerBalances Val(Format(!DrawerID, "0")), Val(Format(!PaymentMethodID, "0")), !TotalFee, False, True, False
                !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
                ChangeDrawerBalances Val(Format(!DrawerID, "0")), Val(Format(!PaymentMethodID, "0")), !TotalFee, True, False, False
                .Update
            End If
            .Close
        End With
    End If
End Sub


Private Sub cmbDrawer_LostFocus()
    If Val(cmbDrawer.BoundText) <> FormVisit.DrawerID Then
        Dim rsVisit As New ADODB.Recordset
        With rsVisit
            If .State = 1 Then .Close
            temSql = "Select * from tblVisit where visitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                ChangeDrawerBalances Val(Format(!DrawerID, "0")), Val(Format(!PaymentMethodID, "0")), !TotalFee, False, True, False
                !DrawerID = Val(cmbDrawer.BoundText)
                ChangeDrawerBalances Val(Format(!DrawerID, "0")), Val(Format(!PaymentMethodID, "0")), !TotalFee, True, False, False
                .Update
            End If
            .Close
        End With
    End If
End Sub

Private Sub Form_Load()
    Call SetColours
    Call FillCombos
    FormVisit.VisitID = CurrentVisit.VisitID
    GetCommonSettings Me
    
    
    Call DisplayVisitCharges
End Sub

Public Sub DisplayVisitCharges()
    FormVisit.Refresh
    txtTotalCharge.Text = Format(FormVisit.TotalFee, "0.00")
    txtProcedureFee.Text = Format(FormVisit.ProcedureFee, "0.00")
    txtMedicinesFee.Text = Format(FormVisit.MedicinesFee, "0.00")
    txtConsultationFee.Text = Format(FormVisit.ConsultationFee, "0.00")
    cmbPaymentMethod.BoundText = FormVisit.PaymentMethodID
    cmbDrawer.BoundText = FormVisit.DrawerID
End Sub

Private Sub CalculateTotalFee()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
     
    SaveSetting App.EXEName, "frmNewVisit", "cmbDrawer", cmbDrawer.BoundText
End Sub

Private Sub txtConsultationFee_LostFocus()
    If Val(txtConsultationFee.Text) <> FormVisit.ConsultationFee Then
        Dim rsVisit As New ADODB.Recordset
        With rsVisit
            If .State = 1 Then .Close
            temSql = "Select * from tblVisit where visitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !ConsultationFee = Val(txtConsultationFee.Text)
                ChangeDrawerBalances !DrawerID, !PaymentMethodID, !TotalFee, False, True, False
                !TotalFee = !ProcedureFee + !ConsultationFee + !MedicineFee
                ChangeDrawerBalances !DrawerID, !PaymentMethodID, !TotalFee, True, False, False
                !TotalCost = !ProcedureCost + !MedicineCost + !ConsultationCost
                .Update
            End If
            .Close
        End With
    End If
    Call DisplayVisitCharges
End Sub



Private Sub txtMedicinesFee_Change()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub

Private Sub txtMedicinesFee_LostFocus()
    If Val(txtMedicinesFee.Text) <> FormVisit.MedicinesFee Then
        Dim rsVisit As New ADODB.Recordset
        With rsVisit
            If .State = 1 Then .Close
            temSql = "Select * from tblVisit where visitID = " & FormVisit.VisitID
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !MedicineFee = Val(txtMedicinesFee.Text)
                ChangeDrawerBalances !DrawerID, !PaymentMethodID, !TotalFee, False, True, False
                !TotalFee = !ProcedureFee + !ConsultationFee + !MedicineFee
                ChangeDrawerBalances !DrawerID, !PaymentMethodID, !TotalFee, True, False, False
                !TotalCost = !ProcedureCost + !MedicineCost + !ConsultationCost
                .Update
            End If
            .Close
        End With
    End If
    Call DisplayVisitCharges

End Sub

Private Sub txtProcedureFee_Change()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub

Private Sub txtTotalCharge_Change()
    txtTotalCharge.Text = Format(Val(txtConsultationFee.Text) + Val(txtProcedureFee.Text) + Val(txtMedicinesFee.Text), "0.00")
End Sub

