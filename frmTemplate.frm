VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTemplate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Income"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
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
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   4800
      Value           =   1  'Checked
      Width           =   855
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
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
   Begin VB.TextBox txtEnd 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1800
      TabIndex        =   6
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox txtStart 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txtTotalIncome 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5040
      TabIndex        =   2
      Top             =   6600
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GridIncome 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   7560
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
   Begin MSDataListLib.DataCombo dtcpatientName 
      Height          =   360
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin MSDataListLib.DataCombo cmbDoc 
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin MSDataListLib.DataCombo cmbIns 
      Height          =   360
      Left            =   2040
      TabIndex        =   11
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin btButtonEx.ButtonEx bttnSettle 
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Settle"
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   7440
      TabIndex        =   17
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Close"
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
   Begin MSDataListLib.DataCombo dtcIx 
      Height          =   360
      Left            =   2040
      TabIndex        =   20
      Top             =   1560
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin VB.Label Label8 
      Caption         =   "&Investigation"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Patient &Name"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Referring Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Referring Institution"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Drawer to End"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Drawer to Start"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Total Income"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTemplate"
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
    dtpDate.Value = Date
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
    Call FillCombos
    Call GetSettings
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FillCombos()
    Dim Doc As New clsFillCombos
    Doc.FillBoolCombo cmbDoctor, "staff", "staff", "IsADoctor", True
End Sub

Private Sub FormatGrid()
    With GridIncome
        .Clear
        .Cols = 6
        .Rows = 1
        
        .row = 0
        
        .col = 0
        .Text = "No."
        
        .col = 1
        .Text = "Time"
        
        .col = 2
        .Text = "Patient"
        
        .col = 3
        .Text = "Total Fee"
        
        .col = 4
        .Text = "Visit ID"
        
        .col = 5
        .Text = "PatientID"
        
        .ColWidth(0) = 700
        .ColWidth(1) = 1100
        .ColWidth(2) = 3730
        .ColWidth(3) = 1100
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        
    End With
End Sub

Private Sub FillGrid()
    'If IsNumeric(cmbDoctor.BoundText) = False Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    Dim TotalFee As Double
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbDoctor.BoundText) = True Then
            temSql = "SELECT tblPatient.Patient, tblVisit.VisitID, tblVisit.VisitTime, tblPatient.PatientID, tblVisit.TotalFee " & _
                        "FROM tblVisit LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID " & _
                        "WHERE (((tblVisit.DoctorID)=" & Val(cmbDoctor.BoundText) & ") AND ((tblVisit.VisitDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#))" & _
                        "ORDER BY tblVisit.VisitID"
                        
        Else
            temSql = "SELECT tblPatient.Patient, tblVisit.VisitID,  tblVisit.VisitTime, tblPatient.PatientID, tblVisit.TotalFee " & _
                        "FROM tblVisit LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID " & _
                        "WHERE (((tblVisit.VisitDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#))" & _
                        "ORDER BY tblVisit.VisitID"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If IsNull(!TotalFee) = False Then
                TotalFee = TotalFee + !TotalFee
            End If
            GridIncome.Rows = GridIncome.Rows + 1
            GridIncome.row = GridIncome.Rows - 1
            
            GridIncome.col = 0
            
            'GridIncome.Text = !PatientID
            
            GridIncome.Text = !VisitID
            
            
            GridIncome.col = 1
            If IsNull(!VisitTime) = False Then
                GridIncome.Text = Format(!VisitTime, "HH:MM AMPM")
            Else
                GridIncome.Text = "Not Seen"
            End If
            
            GridIncome.col = 2
            If IsNull(!Patient) = False Then
                GridIncome.Text = !Patient
            Else
                GridIncome.Text = "Not Seen"
            End If
            
            GridIncome.col = 3
            GridIncome.Text = Format(!TotalFee, "0.00")
            
            GridIncome.col = 4
            If IsNull(!VisitID) = False Then
                GridIncome.Text = ![VisitID]
            End If
            
            GridIncome.col = 5
            If Not IsNull(!PatientID) Then GridIncome.Text = !PatientID
            
            .MoveNext
        Wend
    End With
    
    ColourGrid GridIncome
    
    txtTotalIncome.Text = Format(TotalFee, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

