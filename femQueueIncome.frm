VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmQueueIncome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secession Income"
   ClientHeight    =   8190
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
   ScaleHeight     =   8190
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEnd 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1800
      TabIndex        =   12
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox txtStart 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1800
      TabIndex        =   10
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txtTotalIncome 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5040
      TabIndex        =   8
      Top             =   6600
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GridIncome 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9551
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
      Format          =   20578307
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
   Begin VB.Label Label6 
      Caption         =   "Drawer to End"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Drawer to Start"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Total Income"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   6600
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
    
Private Sub cmbDoctor_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSecession_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
'    Call SaveSettings
    Call FormatGrid
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
        .Cols = 4
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "No."
        
        .Col = 1
        .Text = "Patient"
        
        .Col = 2
        .Text = "Total Fee"
        
        .Col = 3
        .Text = "Visit ID"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 3000
        .ColWidth(2) = 1600
        .ColWidth(3) = 0
        
    End With
End Sub

Private Sub FillGrid()
    If IsNumeric(cmbDoctor.BoundText) = False Then Exit Sub
    If IsNumeric(cmbSecession.BoundText) = False Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    Dim TotalFee As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblQueue.QueueNo, tblQueue.Removed, tblPatient.Patient, tblVisit.VisitID, tblVisit.TotalFee " & _
                    "FROM (tblVisit RIGHT JOIN tblQueue ON tblVisit.QueueID = tblQueue.QueueID) LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID " & _
                    "WHERE (((tblQueue.DoctorID)=" & Val(cmbDoctor.BoundText) & ") AND ((tblQueue.QueueDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#) AND ((tblQueue.SecessionID)= " & Val(cmbSecession.BoundText) & "))"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If IsNull(!TotalFee) = False Then
                TotalFee = TotalFee + !TotalFee
            End If
            GridIncome.Rows = GridIncome.Rows + 1
            GridIncome.Row = GridIncome.Rows - 1
            GridIncome.Col = 0
            GridIncome.Text = !QueueNo
            GridIncome.Col = 1
            If IsNull(!Patient) = False Then
                GridIncome.Text = !Patient
            Else
                If !Removed = True Then
                    GridIncome.Text = "Removed"
                Else
                    GridIncome.Text = "Not seen yet"
                End If
            End If
            
            GridIncome.Col = 2
            GridIncome.Text = Format(!TotalFee, "0.00")
            GridIncome.Col = 3
            If IsNull(!VisitID) = False Then
                GridIncome.Text = ![VisitID]
            End If
            .MoveNext
        Wend
    End With

txtTotalIncome.Text = Format(TotalFee, "0.00")
End Sub

Private Sub CalculateTotals()
    txtEnd.Text = Format(Val(txtTotalIncome.Text) + Val(txtStart.Text), "0.00")
End Sub

Private Sub txtEnd_Change()
    Call CalculateTotals
End Sub

Private Sub txtStart_Change()
    Call CalculateTotals
End Sub

Private Sub txtTotalIncome_Change()
    Call CalculateTotals
End Sub
