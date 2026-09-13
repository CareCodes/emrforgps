VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPharmacyQueue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pharmacy Queue"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpharmacyqueue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   7860
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   6120
      TabIndex        =   23
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ListBox lstNotAvailable 
      Height          =   1740
      Left            =   3000
      TabIndex        =   21
      Top             =   6480
      Width           =   2655
   End
   Begin VB.ListBox lstAvailable 
      Height          =   1740
      Left            =   120
      TabIndex        =   20
      Top             =   6480
      Width           =   2655
   End
   Begin VB.OptionButton optAll 
      Caption         =   "All"
      Height          =   480
      Left            =   3480
      TabIndex        =   18
      Top             =   2040
      Width           =   3255
   End
   Begin VB.OptionButton optOutSideIssue 
      Caption         =   "Outside Medicines from outside"
      Height          =   480
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   3255
   End
   Begin VB.OptionButton optOutdoorIssued 
      Caption         =   "Outdoor Medicines Issued"
      Height          =   600
      Left            =   3480
      TabIndex        =   16
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   6240
      Top             =   5520
   End
   Begin VB.OptionButton optIndoorIssued 
      Caption         =   "Indoor Medicines Issued"
      Height          =   600
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.OptionButton optToOutdoorIssue 
      Caption         =   "To Issue Outdoor Medicines"
      Height          =   480
      Left            =   3480
      TabIndex        =   1
      Top             =   1200
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.OptionButton optToIndoorIssue 
      Caption         =   "To Issue Indoor Medicines"
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
   Begin btButtonEx.ButtonEx btnSale 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Pharmacy Sale"
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
   Begin MSDataListLib.DataList lstQueue 
      Height          =   3420
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6033
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   73662467
      CurrentDate     =   39859
   End
   Begin MSDataListLib.DataCombo cmbSecession 
      Height          =   360
      Left            =   4800
      TabIndex        =   6
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
      TabIndex        =   7
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnOutSide 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Outside"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnIncome 
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Income"
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
   Begin btButtonEx.ButtonEx btnIndoor 
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   6000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Indoor Issue"
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
   Begin btButtonEx.ButtonEx btnIndoorIssue 
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Indoor Issue"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Available"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Secession"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmPharmacyQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String

Private Sub chkCompleted_Click()
    Call FillQueue
End Sub

Private Sub chkRemoved_Click()
    Call FillQueue
End Sub

Private Sub chkToComplete_Click()
    Call FillQueue
End Sub

Private Sub btnAddAgain_Click()
    Dim temLong As Integer
    temLong = MsgBox("Are you sure you want to Add this patient?", vbYesNo)
    If temLong = vbNo Then Exit Sub
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSQL = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Completed = True Then
                MsgBox "This visit is cimpleted. Therefore you can't remove it"
                .Close
                Exit Sub
            End If
            If !Removed = False Then
                MsgBox "This visit is not a removed visit"
                .Close
                Exit Sub
            End If
            !Removed = False
            !RemovedTime = Null
            !RemovedUserID = Null
            .Update
        End If
        .Close
    End With
    Call FillQueue

End Sub

Private Sub btnRemove_Click()
    Dim temLong As Integer
    temLong = MsgBox("Are you sure you want to remove this patient?", vbYesNo)
    If temLong = vbNo Then Exit Sub
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSQL = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Completed = True Then
                MsgBox "This visit is cimpleted. Therefore you can't remove it"
                .Close
                Exit Sub
            End If
            If !Removed = True Then
                MsgBox "This visit is already removed"
                .Close
                Exit Sub
            End If
            !Removed = True
            !RemovedTime = Now
            !RemovedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FillQueue
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnIncome_Click()
    frmPharmacyQueueIncome.Show
    frmPharmacyQueueIncome.ZOrder 0
End Sub

Private Sub btnIndoor_Click()
'    If IsNumeric(lstQueue.BoundText) = False Then
'        Beep
'        Exit Sub
'    End If
'
'    Dim MyVisitID As Long
'
'    frmIndoorIssue.Show
'    frmIndoorIssue.Caption = "Indoor Issue - " & lstQueue.Text
'
'
'    Dim rsTemItem As New ADODB.Recordset
'    With rsTemItem
'        If .State = 1 Then .Close
'        temSQL = "SELECT tblVisit.VisitID, tblVisitItem.PItemID, tblVisitItem.PIssueQuentity " & _
'                    "FROM (tblQueue LEFT JOIN tblVisit ON tblQueue.QueueID = tblVisit.QueueID) LEFT JOIN tblVisitItem ON tblVisit.VisitID = tblVisitItem.VisitID " & _
'                    "WHERE tblQueue.QueueID = " & Val(lstQueue.BoundText) & " AND tblVisitItem.Indoor = True AND tblVisitItem.Deleted = False"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            MyVisitID = !VisitID
'            While .EOF = False
'                If CalculateStock(!PItemID, 0, UserDepartmentID).Stock >= !PIssueQuentity Then
'                    frmIndoorIssue.cmbItem.BoundText = !PItemID
'                    frmIndoorIssue.txtQty.Text = !PIssueQuentity
'                    frmIndoorIssue.btnAdd_Click
'                End If
'                .MoveNext
'            Wend
'        End If
'        .Close
'    End With
'
'    Dim rsQueue As New ADODB.Recordset
'    With rsQueue
'        If .State = 1 Then .Close
'        temSQL = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
'        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
'        If .RecordCount > 0 Then
'            !IndoorIssued = True
'            !IndoorIssuedTime = Now
'            !IndoorIssuedUserID = Null
'            .Update
'        End If
'        .Close
'    End With
'    Call FillQueue
'
'
'    Dim temText As String
'    Dim rsTem As New ADODB.Recordset
'
'    With rsTem
'        If .State = 1 Then .Close
'        temSQL = "SELECT tblVisitItem.VisitItemID, tblVisitItem.ItemID, tblVisitItem.Indoor,  [tblItem.Item] & ' '   &  tblCategory.Category   & ' '   & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblVisitItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] AS Prescreption, 'Issue ' & tblVisitItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
'                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
'                    "WHERE (((tblVisitItem.VisitID)=" & MyVisitID & ") AND ((tblVisitItem.Deleted)=False))" & _
'                    "Order by tblVisitItem.VisitItemID"
'
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            If !Indoor = True Then
'                temText = temText & !Prescreption & vbNewLine
'            End If
'            .MoveNext
'        Wend
'    End With
'
'    Dim MyVisit As New clsVisit
'
'    MyVisit.VisitID = MyVisitID
'
'    frmIndoorIssue.cmbCustomer.BoundText = MyVisit.PatientID
'    frmIndoorIssue.txtComments.Text = temText
'    frmIndoorIssue.txtQueueID.Text = MyVisit.QueueID
'    frmIndoorIssue.ZOrder 0
'
'
End Sub

Private Sub btnIndoorIssue_Click()
    If IsNumeric(lstQueue.BoundText) = False Then
        Beep
        Exit Sub
    End If

    Dim MyVisitID As Long

    frmIndoorIssue.Show
    frmIndoorIssue.Caption = "Indoor Issue - " & lstQueue.Text


    Dim rsTemItem As New ADODB.Recordset
    With rsTemItem
        If .State = 1 Then .Close
        temSQL = "SELECT tblVisit.VisitID, tblVisitItem.PItemID, tblVisitItem.PIssueQuentity " & _
                    "FROM (tblQueue LEFT JOIN tblVisit ON tblQueue.QueueID = tblVisit.QueueID) LEFT JOIN tblVisitItem ON tblVisit.VisitID = tblVisitItem.VisitID " & _
                    "WHERE tblQueue.QueueID = " & Val(lstQueue.BoundText) & " AND tblVisitItem.Indoor = True AND tblVisitItem.Deleted = False"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyVisitID = !VisitID
            While .EOF = False
                If CalculateStock(!PItemID, 0, UserDepartmentID).Stock >= !PIssueQuentity Then
                    frmIndoorIssue.cmbItem.BoundText = !PItemID
                    frmIndoorIssue.txtQty.Text = !PIssueQuentity
                    frmIndoorIssue.btnAdd_Click
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With

    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSQL = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !IndoorIssued = True
            !IndoorIssuedTime = Now
            !IndoorIssuedUserID = Null
            .Update
        End If
        .Close
    End With
    Call FillQueue


    Dim temText As String
    Dim rsTem As New ADODB.Recordset

    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblVisitItem.VisitItemID, tblVisitItem.ItemID, tblVisitItem.Indoor,  [tblItem.Item] & ' '   &  tblCategory.Category   & ' '   & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblVisitItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] AS Prescreption, 'Issue ' & tblVisitItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
                    "WHERE (((tblVisitItem.VisitID)=" & MyVisitID & ") AND ((tblVisitItem.Deleted)=False))" & _
                    "Order by tblVisitItem.VisitItemID"

        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If !Indoor = True Then
                temText = temText & !Prescreption & vbNewLine
            End If
            .MoveNext
        Wend
    End With

    Dim MyVisit As New clsVisit

    MyVisit.VisitID = MyVisitID

    frmIndoorIssue.cmbCustomer.BoundText = MyVisit.PatientID
    frmIndoorIssue.txtComments.Text = temText
    frmIndoorIssue.txtQueueID.Text = MyVisit.QueueID
    frmIndoorIssue.ZOrder 0



End Sub

Private Sub btnOutSide_Click()
    Dim temLong As Integer
    temLong = MsgBox("Are you sure you want let patient to take outside medicines from any other pharmacy?", vbYesNo)
    If temLong = vbNo Then Exit Sub
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSQL = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !OutsideIssued = True
            !OutsideIssuedTime = Now
            !OutsideIssuedUserID = Null
            .Update
        End If
        .Close
    End With
    Call FillQueue
End Sub

Private Sub btnSale_Click()
    If IsNumeric(lstQueue.BoundText) = False Then
        Beep
        Exit Sub
    End If
    
    Dim MyVisitID As Long
    
    frmPharmacySale.Show
    frmPharmacySale.Caption = "Pharmacy Sale - " & lstQueue.Text
    
    
    Dim rsTemItem As New ADODB.Recordset
    With rsTemItem
        If .State = 1 Then .Close
        temSQL = "SELECT tblVisit.VisitID, tblVisitItem.PItemID, tblVisitItem.PIssueQuentity " & _
                    "FROM (tblQueue LEFT JOIN tblVisit ON tblQueue.QueueID = tblVisit.QueueID) LEFT JOIN tblVisitItem ON tblVisit.VisitID = tblVisitItem.VisitID " & _
                    "WHERE tblQueue.QueueID = " & Val(lstQueue.BoundText) & " AND tblVisitItem.Indoor = False AND tblVisitItem.Deleted = False"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyVisitID = !VisitID
            While .EOF = False
                If CalculateStock(!PItemID, 0, UserDepartmentID).Stock >= !PIssueQuentity Then
                    frmPharmacySale.cmbItem.BoundText = !PItemID
                    frmPharmacySale.txtQty.Text = !PIssueQuentity
                    frmPharmacySale.btnAdd_Click
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSQL = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !OutdoorIssued = True
            !OutdoorIssuedTime = Now
            !OutdoorIssuedUserID = Null
            .Update
        End If
        .Close
    End With
    Call FillQueue
    
    
    Dim temText As String
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblVisitItem.VisitItemID, tblVisitItem.ItemID, tblVisitItem.Indoor,  [tblItem.Item] & ' '   &  tblCategory.Category   & ' '   & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblVisitItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] AS Prescreption, 'Issue ' & tblVisitItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
                    "WHERE (((tblVisitItem.VisitID)=" & MyVisitID & ") AND ((tblVisitItem.Deleted)=False))" & _
                    "Order by tblVisitItem.VisitItemID"
                    
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If !Indoor = False Then
                temText = temText & !Prescreption & vbNewLine
            End If
            .MoveNext
        Wend
    End With
    
    Dim MyVisit As New clsVisit
    
    MyVisit.VisitID = MyVisitID
    
    frmPharmacySale.cmbCustomer.BoundText = MyVisit.PatientID
    frmPharmacySale.txtComments.Text = temText
    frmPharmacySale.txtQueueID.Text = MyVisit.QueueID
    frmPharmacySale.ZOrder 0
    
End Sub

Private Sub cmbDoctor_Change()
    Call FillQueue
End Sub

Private Sub cmbSecession_Change()
    Call FillQueue
End Sub

Private Sub dtpDate_Change()
    Call FillQueue
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
    Call FillCombos
    Call GetFormSettings
    dtpDate.Value = Date
    Call FillQueue
     If AutoSecession = True Then Call SetAutoSecession
End Sub

Private Sub SetAutoSecession()
    
    If Time > #2:00:00 PM# Then
        cmbSecession.Text = "Evening"
    Else
        cmbSecession.Text = "Morning"
    End If
End Sub


Private Sub FillCombos()
    Dim Doc As New clsFillCombos
    Doc.FillBoolCombo cmbDoctor, "staff", "staff", "IsADoctor", True
    Dim Secession As New clsFillCombos
    Secession.FillBoolCombo cmbSecession, "Secession", "Secession", "IsVisitSecession", True
End Sub

Private Sub GetFormSettings()
    cmbDoctor.BoundText = Val(GetSetting(App.EXEName, Me.Name, "DoctorID", 0))
    cmbSecession.BoundText = Val(GetSetting(App.EXEName, Me.Name, "SecessionID", 0))
    GetCommonSettings Me
    
    
    On Error Resume Next
    Dim MyControl  As Control
    For Each MyControl In Controls
        If Left(MyControl.Name, 3) = "opt" Then
            MyControl.Value = GetSetting(App.EXEName, Me.Name, MyControl.Name, "False")
        End If
    Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetFormSettings
End Sub

Private Sub SetFormSettings()
    SaveSetting App.EXEName, Me.Name, "DoctorID", Val(cmbDoctor.BoundText)
    SaveSetting App.EXEName, Me.Name, "SecessionID", Val(cmbSecession.BoundText)
    SaveCommonSettings Me
    
    
    Dim MyControl  As Control
    For Each MyControl In Controls
        If Left(MyControl.Name, 3) = "opt" Then
            SaveSetting App.EXEName, Me.Name, MyControl.Name, MyControl.Value
        End If
        If Left(MyControl.Name, 3) = "cmb" Then
            SaveSetting App.EXEName, Me.Name, MyControl.Name, MyControl.Text
        End If
    Next


End Sub

Private Sub FillQueue()
    Dim temQID As Long
    temQID = Val(lstQueue.BoundText)
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSQL = "SELECT tblQueue.QueueNo & ' ' & tblPatient.Patient & ' (' & tblPatient.PatientID & ') ' & ' - ' & tblQueue.Age  as QueueDisplay, tblQueue.QueueID FROM tblQueue LEFT JOIN tblPatient ON tblQueue.PatientID = tblPatient.PatientID where DoctorID = " & Val(cmbDoctor.BoundText) & " AND QueueDate = #" & Format(dtpDate.Value, "dd MMMM yyyy") & "# AND SEcessionID = " & Val(cmbSecession.BoundText) & " AND tblQueue.Deleted = False "
        
        If optToIndoorIssue.Value = True Then
            temSQL = temSQL & " and Completed = true And IndoorIssued = False  AND HasIndoorDrugs = True "
        ElseIf optToOutdoorIssue.Value = True Then
            temSQL = temSQL & " and  Completed = true AND OutdoorIssued = false AND HasOutdoorDrugs = True And OutSideIssued = False "
        ElseIf optIndoorIssued.Value = True Then
            temSQL = temSQL & " and  Completed = true  AND IndoorIssued = true  AND HasIndoorDrugs = True "
        ElseIf optOutdoorIssued.Value = True Then
            temSQL = temSQL & " and  Completed = true  AND OutdoorIssued = True And HasOutdoorDrugs = True "
        ElseIf optOutdoorIssued.Value = True Then
            temSQL = temSQL & " and  Completed = true  AND OutsideIssued = true AND HasOutdoorDrugs = True "
        ElseIf optOutSideIssue.Value = True Then
            temSQL = temSQL & " And Completed = true And OutSideIssued = true"
        ElseIf optAll.Value = True Then
            
        End If
        
        temSQL = temSQL & " Order by QueueNo"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstQueue
        Set .RowSource = rsQueue
        .ListField = "QueueDisplay"
        .BoundColumn = "QueueID"
        .BoundText = temQID
    End With
End Sub

Private Sub lstQueue_Click()
    lstAvailable.Clear
    lstNotAvailable.Clear
    Dim rsTemItem As New ADODB.Recordset
    With rsTemItem
        If .State = 1 Then .Close
        temSQL = "SELECT tblItem.Item, tblVisitItem.PItemID, tblVisitItem.PIssueQuentity FROM ((tblQueue LEFT JOIN tblVisit ON tblQueue.QueueID = tblVisit.QueueID) LEFT JOIN tblVisitItem ON tblVisit.VisitID = tblVisitItem.VisitID) LEFT JOIN tblItem ON tblVisitItem.PItemID = tblItem.ItemID " & _
                    "WHERE (((tblVisitItem.PItemID)<>0) AND ((tblQueue.QueueID)=" & Val(lstQueue.BoundText) & ") AND ((tblVisitItem.Indoor)=False) AND ((tblVisitItem.Deleted)=False))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If CalculateStock(!PItemID, 0, UserDepartmentID).Stock < !PIssueQuentity Then
                lstNotAvailable.AddItem !Item
            Else
                lstAvailable.AddItem !Item
            End If
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub optAll_Click()
    Call FillQueue

End Sub

Private Sub optIndoorIssued_Click()
    Call FillQueue

End Sub

Private Sub optOutdoorIssued_Click()
    Call FillQueue

End Sub

Private Sub optOutSideIssue_Click()
    Call FillQueue

End Sub

Private Sub optToIndoorIssue_Click()
    Call FillQueue
End Sub

Private Sub optToOutdoorIssue_Click()
    Call FillQueue

End Sub

Private Sub Timer1_Timer()
    Call FillQueue
End Sub

