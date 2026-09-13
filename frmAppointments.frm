VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAppointments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appointments"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
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
   ScaleHeight     =   8460
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66715651
      CurrentDate     =   40118
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   7920
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
   Begin MSFlexGridLib.MSFlexGrid GridApp 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11668
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbDoctor 
      Height          =   360
      Left            =   2040
      TabIndex        =   2
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
   Begin btButtonEx.ButtonEx bttnSettle 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   7800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Visit from Appointment"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel Appointment"
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
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add Appointment"
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
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "&Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAppointments"
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
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub

Private Sub FillCombos()
    Dim Doc As New clsFillCombos
    Doc.FillBoolCombo cmbDoctor, "staff", "staff", "IsADoctor", True
End Sub

Private Sub FormatGrid()
    With GridApp
        .Clear
        .Cols = 6
        .Rows = 1
        
        .row = 0
        
        .col = 0
        .Text = "ID"
        
        .col = 1
        .Text = "Time"
        
        .col = 2
        .Text = "Patient"
        
        .col = 3
        .Text = "ID"
        
        .col = 4
        .Text = "Status"
        
        .col = 5
        .Text = "Comments"
        
        .ColWidth(5) = 0
        
    End With
End Sub

Private Sub FillGrid()
    
    Dim rsTem As New ADODB.Recordset
    Dim TotalFee As Double
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbDoctor.BoundText) = True Then
            temSql = "SELECT tblAppointment.AppointmentID, tblAppointment.AppTime, tblPatient.Patient, tblAppointment.PatientID, tblAppointment.Cancelled, tblAppointment.NewVisit, tblAppointment.Comments " & _
                        "FROM tblStaff RIGHT JOIN (tblAppointment LEFT JOIN tblPatient ON tblAppointment.PatientID = tblPatient.PatientID) ON tblStaff.StaffID = tblAppointment.DoctorID " & _
                        "WHERE (((tblAppointment.DoctorID)=" & Val(cmbDoctor.BoundText) & ") AND ((tblAppointment.AppDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#))" & _
                        "ORDER BY tblAppointment.AppTime"
                        
        Else
            temSql = "SELECT tblAppointment.AppointmentID, tblAppointment.AppTime, tblPatient.Patient, tblAppointment.PatientID, tblAppointment.Cancelled, tblAppointment.NewVisit, tblAppointment.Comments " & _
                        "FROM tblStaff RIGHT JOIN (tblAppointment LEFT JOIN tblPatient ON tblAppointment.PatientID = tblPatient.PatientID) ON tblStaff.StaffID = tblAppointment.DoctorID " & _
                        "WHERE (((tblAppointment.AppDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#))" & _
                        "ORDER BY tblAppointment.AppTime"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            
            GridApp.Rows = GridApp.Rows + 1
            GridApp.row = GridApp.Rows - 1
            
            GridApp.col = 0
            GridApp.Text = !AppointmentID
            
            GridApp.col = 1
            GridApp.Text = Format(!AppTime, "HH:MM AMPM")
            
            GridApp.col = 2
            If IsNull(!Patient) = False Then
                GridApp.Text = !Patient
            Else
                GridApp.Text = ""
            End If
            
            GridApp.col = 3
            GridApp.Text = Format(!PatientID, "0")
            
            GridApp.col = 4
            If IsNull(!Cancelled) = False Then
                GridApp.Text = "Appointment Cancelled."
            End If
            If IsNull(!NewVisit) = False Then
                GridApp.Text = "Appointment Completed."
            End If
            
            
            GridApp.col = 5
            If Not IsNull(!Comments) Then GridApp.Text = !Comments
            
            .MoveNext
        Wend
    End With
    
    ColourGrid GridApp
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

