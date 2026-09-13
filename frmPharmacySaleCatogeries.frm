VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPharmacySaleCatogeries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Categories"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9780
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
   ScaleHeight     =   8160
   ScaleWidth      =   9780
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   4800
      TabIndex        =   19
      Top             =   6720
      Width           =   4815
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Cancel"
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
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   6720
      Width           =   4455
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Edit"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   240
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
   End
   Begin VB.Frame Frame4 
      Height          =   6495
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   4455
      Begin MSDataListLib.DataCombo cmbSaleCategory 
         Height          =   5940
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   10478
         _Version        =   393216
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   4800
      TabIndex        =   12
      Top             =   240
      Width           =   4815
      Begin VB.Frame Frame2 
         Caption         =   "Charges"
         Height          =   1695
         Left            =   240
         TabIndex        =   24
         Top             =   4680
         Width           =   4335
         Begin VB.OptionButton optByWP 
            Caption         =   "By Wholesale Price"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1320
            Width           =   2415
         End
         Begin VB.OptionButton optNoCharge 
            Caption         =   "No Charge"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton optByPP 
            Caption         =   "By Purchase Price"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton optByRP 
            Caption         =   "By Retail Price"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   2415
         End
      End
      Begin VB.TextBox txtProfitMargin 
         Height          =   360
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   4335
      End
      Begin MSDataListLib.DataCombo cmbPaymentMethod 
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   4200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtPercentage 
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox txtCategoryName 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Customer type"
         Height          =   1455
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   4335
         Begin VB.OptionButton optInstitution 
            Caption         =   "Institution"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton optUnit 
            Caption         =   "Hospital Unit"
            Height          =   255
            Left            =   2160
            TabIndex        =   22
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optionStaff 
            Caption         =   "Staff"
            Height          =   255
            Left            =   2160
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optionBHT 
            Caption         =   "BHT"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton optionOutPation 
            Caption         =   "Out Patient"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Additional Profit Margin"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Discount Percent"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Payment Method"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Sale Category"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
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
End
Attribute VB_Name = "frmPharmacySaleCatogeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String
    Dim A As Byte

Private Sub bttnAdd_Click()
    Dim temStr As String
    If IsNumeric(cmbSaleCategory.BoundText) = False Then
        temStr = cmbSaleCategory.Text
    Else
        temStr = Empty
    End If
    cmbSaleCategory.Text = Empty
    Call ClearValues
    Call EditMode
    txtCategoryName.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbSaleCategory.Text = Empty
    cmbSaleCategory.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Call EditMode
    txtCategoryName.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Function ClearValues()
    txtCategoryName.Text = Empty
    optionOutPation.Value = False
    optionBHT.Value = False
    optionStaff.Value = False
    optUnit.Value = False
    optInstitution.Value = False
    cmbPaymentMethod.Text = Empty
    txtPercentage.Text = Empty
    txtProfitMargin.Text = Empty
    optNoCharge.Value = False
    optByPP.Value = False
    optByRP.Value = False
    optByWP.Value = False
End Function

Private Function CheckValues() As Boolean
    CheckValues = False
    If txtCategoryName.Text = Empty Then
        A = MsgBox("Enter Category Name", vbCritical + vbOKOnly, "Error"): txtCategoryName.SetFocus
        txtCategoryName.SetFocus
        Exit Function
    End If
    If optInstitution.Value = False And optUnit.Value = False And optionOutPation.Value = False And optionBHT.Value = False And optionStaff.Value = False Then
        A = MsgBox("Select Customer Type", vbCritical + vbOKOnly, "Error")
        optionOutPation.SetFocus
        Exit Function
    End If
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        A = MsgBox("Select Pay Mode", vbCritical + vbOKOnly, "Error")
        cmbPaymentMethod.SetFocus
        Exit Function
    End If
    If IsNumeric(txtPercentage.Text) = False Then
        A = MsgBox("Enter Percentage", vbCritical + vbOKOnly, "Error")
        txtPercentage.SetFocus
        Exit Function
    End If
    If optNoCharge.Value = False And optByPP.Value = False And optByRP.Value = False And optByWP.Value = False Then
        A = MsgBox("Select Payment Option", vbCritical + vbOKOnly, "Error")
        optByRP.SetFocus
        Exit Function
    End If
    CheckValues = True
End Function

Private Sub SelectMode()
    Frame4.Enabled = True
    
    Frame1.Enabled = False

    
    bttnSave.Enabled = False
    bttnCancel.Enabled = False
    
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    btnDelete.Enabled = True
End Sub

Private Sub EditMode()
    Frame4.Enabled = False
    
    Frame1.Enabled = True
    
    bttnSave.Enabled = True
    bttnCancel.Enabled = True
    
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    btnDelete.Enabled = False
End Sub


Private Sub bttnSave_Click()
    If CheckValues = False Then Exit Sub
    If IsNumeric(cmbSaleCategory.BoundText) = False Then
        Call SaveNew
    Else
        Call SaveOld
    End If
    Call ClearValues
    Call SelectMode
    Call FillCombos
    cmbSaleCategory.Text = Empty
    cmbSaleCategory.SetFocus
End Sub

Private Sub FillCombos()
    Dim SC As New clsFillCombos
    SC.FillAnyCombo cmbSaleCategory, "SaleCategory", True
    Dim PM As New clsFillCombos
    PM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "ToReceive", True
    
End Sub


Private Sub DisplaySeleted()
    Call ClearValues
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * From tblSaleCategory Where (SaleCategoryID = " & cmbSaleCategory.BoundText & ")"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            
            txtCategoryName.Text = !SaleCategory
            txtPercentage.Text = !SaleDiscountPercent
            
            cmbPaymentMethod.BoundText = !PaymentMethodID
            
            If Not IsNull(!ProfitMargin) Then txtProfitMargin.Text = !ProfitMargin
            If Not IsNull(!PaymentMethodID) Then cmbPaymentMethod.BoundText = !PaymentMethodID
            
            If !ForOutDoor = True Then optionOutPation.Value = True
            If !ForBHT = True Then optionBHT.Value = True
            If !ForStaff = True Then optionStaff.Value = True
            If !ForUnit = True Then optUnit.Value = True
            If !ForInstitution = True Then optInstitution.Value = True
            
            If !NoCharge = True Then optNoCharge.Value = True
            If !ByrP = True Then optByRP.Value = True
            If !ByWP = True Then optByWP.Value = True
            If !Bypp = True Then optByPP.Value = True
        
        End If
        If .State = 1 Then .Close
    End With
End Sub

Private Sub SaveNew()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * From tblSaleCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !SaleCategory = txtCategoryName.Text
        !SaleDiscountPercent = Val(txtPercentage.Text)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        If optNoCharge.Value = True Then
            !NoCharge = True
        Else
            !NoCharge = False
        End If
        If optByRP.Value = True Then
            !ByrP = True
        Else
            !ByrP = False
        End If
        If optByWP.Value = True Then
            !ByWP = True
        Else
            !ByWP = False
        End If
        If optByPP.Value = True Then
            !Bypp = True
        Else
            !Bypp = False
        End If
        If optionOutPation.Value Then
            !ForOutDoor = True
        Else
            !ForOutPatient = False
        End If
        If optionBHT.Value Then
            !ForBHT = True
        Else
            !ForBHT = False
        End If
        If optionStaff.Value Then
            !ForStaff = True
        Else
            !ForStaff = False
        End If
        If optUnit.Value = True Then
            !ForUnit = True
        Else
            !ForUnit = False
        End If
        If optInstitution.Value = True Then
            !ForInstitution = True
        Else
            !ForInstitution = False
        End If
        !ProfitMargin = Val(txtProfitMargin.Text)
        .Update
        Call SelectMode
        Call ClearValues
        Call FillCombos
        If .State = 1 Then .Close
        cmbSaleCategory.SetFocus
        cmbSaleCategory.Text = Empty
    End With
End Sub

Private Sub SaveOld()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * From tblSaleCategory Where (SaleCategoryID = " & cmbSaleCategory.BoundText & ")"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !SaleCategory = txtCategoryName.Text
            !SaleDiscountPercent = Val(txtPercentage.Text)
            !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
            If optNoCharge.Value = True Then
                !NoCharge = True
            Else
                !NoCharge = False
            End If
            If optByRP.Value = True Then
                !ByrP = True
            Else
                !ByrP = False
            End If
            If optByWP.Value = True Then
                !ByWP = True
            Else
                !ByWP = False
            End If
            If optByPP.Value = True Then
                !Bypp = True
            Else
                !Bypp = False
            End If
            If optionOutPation.Value Then
                !ForOutDoor = True
            Else
                !ForOutDoor = False
            End If
            If optionBHT.Value Then
                !ForBHT = True
            Else
                !ForBHT = False
            End If
            If optionStaff.Value Then
                !ForStaff = True
            Else
                !ForStaff = False
            End If
            If optUnit.Value = True Then
                !ForUnit = True
            Else
                !ForUnit = False
            End If
            If optInstitution.Value = True Then
                !ForInstitution = True
            Else
                !ForInstitution = False
            End If
            !ProfitMargin = Val(txtProfitMargin.Text)
            .Update
            Call SelectMode
            Call ClearValues
            Call FillCombos
        End If
        If .State = 1 Then .Close
    End With
End Sub

Private Sub DtcSaleCategory_Click(Area As Integer)
    If IsNumeric(cmbSaleCategory.BoundText) = False Then Exit Sub
    Call DisplaySeleted
End Sub

Private Sub cmbSaleCategory_Click(Area As Integer)
    Call ClearValues
    Call DisplaySeleted
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
    
    Call SetColours
    Call FillCombos
    Call SelectMode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub
