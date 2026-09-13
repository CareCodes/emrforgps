VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDrawerBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Income"
   ClientHeight    =   4365
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
   ScaleHeight     =   4365
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   3720
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
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5280
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GridBalance 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label4 
      Caption         =   "Total"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Drawer"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDrawerBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    
    cmbDrawer.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbDrawer.Name, 1))
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, Me.Name, cmbDrawer.Name, cmbDrawer.BoundText
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

Private Sub cmbDrawer_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbDrawer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbDrawer.Text = Empty
    End If
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
    Doc.FillAnyCombo cmbDrawer, "Drawer", True
End Sub

Private Sub FormatGrid()
    With GridBalance
        .Clear
        .Cols = 2
        .Rows = 1
        .Row = 0
        .Col = 0
        .Text = "Payment Method"
        .Col = 1
        .Text = "Balance"
        .ColWidth(0) = 2700
        .ColWidth(1) = .Width - .ColWidth(0) - 150
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    Dim TotalFee As Double
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbDrawer.BoundText) = True Then
            temSql = "SELECT tblPaymentMethod.PaymentMethod, tblDrawerBalance.DrawerBalance " & _
                        "FROM tblPaymentMethod LEFT JOIN tblDrawerBalance ON tblPaymentMethod.PaymentMethodID = tblDrawerBalance.PaymentMethodID " & _
                        "WHERE tblDrawerBalance.DrawerID = " & Val(cmbDrawer.BoundText) & " " & _
                        "Order by tblPaymentMethod.PaymentMethodID"
                        
        Else
            temSql = "SELECT tblPaymentMethod.PaymentMethod, tblDrawerBalance.DrawerBalance " & _
                        "FROM tblPaymentMethod LEFT JOIN tblDrawerBalance ON tblPaymentMethod.PaymentMethodID = tblDrawerBalance.PaymentMethodID " & _
                        "Order by tblPaymentMethod.PaymentMethodID"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If IsNull(!DrawerBalance) = False Then
                TotalFee = TotalFee + !DrawerBalance
            End If
            GridBalance.Rows = GridBalance.Rows + 1
            GridBalance.Row = GridBalance.Rows - 1
            GridBalance.Col = 0
            GridBalance.Text = !PaymentMethod
            If GridBalance.Row Mod 2 = 0 Then
                GridBalance.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridBalance.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            GridBalance.Col = 1
            If IsNull(!DrawerBalance) = False Then
                GridBalance.Text = Format(!DrawerBalance, "0.00")
            Else
                GridBalance.Text = "0.00"
            End If
            If GridBalance.Row Mod 2 = 0 Then
                GridBalance.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridBalance.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            .MoveNext
        Wend
    End With
    txtTotal.Text = Format(TotalFee, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

