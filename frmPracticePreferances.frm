VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPracticePreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Practice Preferances"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   6225
   ScaleWidth      =   8700
   Begin VB.CommandButton btnAdmin 
      Caption         =   "Administration"
      Height          =   495
      Left            =   6960
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add Indoor Medicine Charges to Consultation"
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   4935
      Begin VB.OptionButton optNoAddIndoor 
         Caption         =   "No"
         Height          =   240
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optAddIndoor 
         Caption         =   "Yes"
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox txtSupplier 
      Height          =   360
      Left            =   1320
      TabIndex        =   13
      Top             =   4560
      Width           =   4575
   End
   Begin VB.TextBox txtAddress 
      Height          =   330
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   5040
      Width           =   4575
   End
   Begin VB.TextBox txtTelephone 
      Height          =   345
      Left            =   1320
      TabIndex        =   11
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add Outdoor Medicine Charges to Consultation"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4935
      Begin VB.OptionButton optAddMedicineCharges 
         Caption         =   "Yes"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optDoNotAddDrugCharges 
         Caption         =   "No"
         Height          =   240
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Issue Indoor"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
      Begin VB.OptionButton optAutoIssueIndoorNo 
         Caption         =   "No"
         Height          =   240
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optAutoIssueIndoorYes 
         Caption         =   "Yes"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.TextBox txtConsultationCost 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin VB.TextBox txtConsultationFee 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Institution"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Telephone"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Usual Consultation Cost"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Usual Consultation Fee"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmPracticePreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String

Private Sub DisplaySettings()
    txtConsultationFee.Text = Format(UsualConsultationFee, "0.00")
    txtConsultationCost.Text = Format(UsualConsultationCost, "0.00")
    If AutoIssueIndoor = True Then
        optAutoIssueIndoorYes.Value = True
    Else
        optAutoIssueIndoorNo.Value = True
    End If
    If AddMedicineChargesToFee = True Then
        optAddMedicineCharges.Value = True
    Else
        optDoNotAddDrugCharges.Value = True
    End If
    'AddMedicineChargesToFee
    If AddIndoorMedicineChargesToFee = True Then
        optAddIndoor.Value = True
    Else
        optNoAddIndoor.Value = True
    End If
End Sub


Private Sub SaveSettings()

    'To Registry
    SaveSetting App.EXEName, "Options", "UsualConsultationFee", Val(txtConsultationFee.Text)
    SaveSetting App.EXEName, "Options", "UsualConsultationCost", Val(txtConsultationCost.Text)
    SaveSetting App.EXEName, "Options", "AutoIssueIndoor", optAutoIssueIndoorYes.Value
    SaveSetting App.EXEName, "Options", "AddMedicineChargesToFee", optAddMedicineCharges.Value
    SaveSetting App.EXEName, "Options", "AddIndoorMedicineChargesToFee", optAddIndoor.Value
    SaveCommonSettings Me
    
    'To Database
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblInstitution where IsOwner = True AND Deleted = False"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
            !IsOwner = True
        End If
        !Institution = txtSupplier.Text
        !Address = txtAddress.Text
        !Telephone = txtTelephone.Text
        .Update
        .Close
    End With
    Set rsTem = Nothing
    
    
    
    'To Memory
    UsualConsultationFee = Val(txtConsultationFee.Text)
    UsualConsultationCost = Val(txtConsultationCost.Text)
    AutoIssueIndoor = optAutoIssueIndoorYes.Value
    AddIndoorMedicineChargesToFee = optAddIndoor.Value
    AddMedicineChargesToFee = optAddMedicineCharges.Value
End Sub

Private Sub btnAdmin_Click()
    frmAdmin.Show
    frmAdmin.ZOrder 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblInstitution where IsOwner = True AND Deleted = False"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtSupplier.Text = Format(!Institution, "")
            txtAddress.Text = Format(!Address, "")
            txtTelephone.Text = Format(!Telephone, "")
        End If
        .Close
    End With
    Set rsTem = Nothing
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
    Call DisplayDetails
    
    Call DisplaySettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
