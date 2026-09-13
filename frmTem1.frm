VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTem1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   885
   ClientTop       =   1890
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9720
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9340
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTem1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Address"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtAge"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAddress"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtComments"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btnSave"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtpDOB"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "btnStart"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton btnStart 
         Caption         =   "Sart"
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   6600
         TabIndex        =   22
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   133890051
         CurrentDate     =   40892
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   4680
         TabIndex        =   21
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtComments 
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   2640
         Width           =   4815
      End
      Begin VB.TextBox txtAddress 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   2040
         Width           =   4815
      End
      Begin VB.TextBox txtAge 
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label6 
         Caption         =   "Comments"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Address 
         Caption         =   "Adddress"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Age"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Dates"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IxItemCategories"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Correct Print Items"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton cmmdAddRandomNumbers 
      Caption         =   "Add Random Numbers"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton btnAddFormNames 
      Caption         =   "Add Menu Names"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton btnCapitalize 
      Caption         =   "Capitalize"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtField 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtTable 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   5040
   End
   Begin VB.TextBox txtActive 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtWHandle 
      Height          =   405
      Left            =   5640
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fiels"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Table"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmTem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddFormNames_Click()
    Dim rsForm As New ADODB.Recordset
    Dim rsControl As New ADODB.Recordset
    Dim MyForm As Form
    Dim FormID As Long
    Dim i As Integer
    Dim MyControl As Control
    Dim temText As String
    For Each MyForm In Forms
        FormID = GetFormID(MyForm.Name, MyForm.Caption)
        For Each MyControl In MyForm.Controls
            With rsForm
                If TypeOf MyControl Is SSTab Then
                    For i = 0 To MyControl.Tabs - 1
                        If .State = 1 Then .Close
                        temSql = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "' AND ControlIndex = " & i
                        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                        MyControl.Tab = i
                        If .RecordCount > 0 Then
                            !ControlText = GetControlText(MyControl)
                        Else
                            .AddNew
                            !FormID = FormID
                            !Control = MyControl.Name
                            !ControlType = GetControlType(MyControl)
                            !ControlText = GetControlText(MyControl)
                            !ControlIndex = i
                        End If
                        .Update
                        .Close
                    Next i
                Else
                    If .State = 1 Then .Close
                    temSql = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "'"
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !ControlText = GetControlText(MyControl)
                    Else
                        .AddNew
                        !FormID = FormID
                        !Control = MyControl.Name
                        !ControlType = GetControlType(MyControl)
                        !ControlText = GetControlText(MyControl)
                    End If
                    .Update
                    .Close
                End If
            End With
        Next
    Next
End Sub

Private Sub btnCapitalize_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from " & txtTable.Text
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            .Fields(txtField.Text).Value = UCase(.Fields(txtField.Text).Value)
            .Update
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub btnStart_Click()
    Dim temSql As String
                        
    Dim rsCustomer As New ADODB.Recordset
    Dim rsPatient As New ADODB.Recordset
    Dim temAgeAndAddress As String
    Dim addStart As Integer
    Dim temAge As Integer
    Dim myInt As Integer
    
    With rsCustomer
        If .State = 1 Then .Close
        temSql = "SELECT * from CUSTOMER"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            On Error Resume Next
            txtName.Text = !Name
            temAgeAndAddress = !Field2
            addStart = InStr(LCase(temAgeAndAddress), "yr,")
            If addStart > 0 Then
                temAge = Val(Left(temAgeAndAddress, 2))
                txtAge.Text = Left(temAgeAndAddress, 2)
                dtpDOB.Value = DateSerial((Year(Date) - temAge), 1, 1)
                txtAddress.Text = Right(temAgeAndAddress, Len(temAgeAndAddress) - addStart - 2)
                txtComments.Text = txtComments.Text & !Field3
                txtComments.Text = txtComments.Text & !Field4
                txtComments.Text = txtComments.Text & !Field5
                txtComments.Text = txtComments.Text & !Field6
                txtComments.Text = txtComments.Text & !Field7
                txtComments.Text = txtComments.Text & !Field8
                txtComments.Text = txtComments.Text & !Field9
            End If
'             myInt = MsgBox("Continut?", vbYesNo)
'             If myInt = vbNo Then
'                End
'            Else

                If rsPatient.State = 1 Then rsPatient.Close
                temSql = "SELECT * from tblPatient where PatientID= 1"
                rsPatient.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                rsPatient.AddNew
                rsPatient!Patient = txtName.Text
                rsPatient!GivenName = txtName.Text
                rsPatient!Comments = txtComments.Text
                rsPatient!DateOfBirth = dtpDOB.Value
                rsPatient!SexID = 1
                rsPatient!Address = txtAddress.Text
                rsPatient.Update
'             End If

                txtName.Text = ""
                txtAge.Text = ""
                txtAddress.Text = ""
                txtComments.Text = ""

            .MoveNext
        Wend
    End With
    
End Sub

Private Sub cmmdAddRandomNumbers_Click()
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    Dim d As Long
    Dim m As Long
    Dim Y As Long
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblDoctorNumber"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        For i = 1 To 10
            For Y = 0 To 9
                For m = 1 To 12
                    For d = 1 To 31
                        .AddNew
                        !DoctorID = i
                        !AddedYear = Y
                        !AddedMonth = m
                        !AddedDate = d
                        !DeletedYear = Round(Rnd * 10, 0)
                        !DeletedMonth = Round(Rnd * 12, 0)
                        !DeletedDate = Round(Rnd * 31, 0)
                        If !DeletedYear < 0 Then !DeletedYear = 0
                        If !DeletedYear > 9 Then !DeletedYear = 9
                        If !DeletedMonth < 1 Then !DeletedMonth = 1
                        If !DeletedMonth > 12 Then !DeletedMonth = 12
                        If !DeletedDate < 1 Then !DeletedDate = 1
                        If !DeletedDate > 31 Then !DeletedDate = 31
                        !DoctorNo = Rnd * 9999
                        !HospitalNo = Rnd * 9999
                        If !DoctorNo > 9999 Then !DoctorNo = 9999
                        If !HospitalNo > 9999 Then !HospitalNo = 9999
                        .Update
                    Next d
                Next m
            Next Y
        Next i
    End With
End Sub

Private Sub Command1_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl1 As Double
    Dim temDbl2 As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPrintItem where X1 > X2"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            temDbl1 = !X1
            temDbl2 = !X2
            !X1 = temDbl2
            !X2 = temDbl1
            .Update
            .MoveNext
        Wend
        If .State = 1 Then .Close
        temSql = "Select * from tblPrintItem where Y1 > Y2"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            temDbl1 = !Y1
            temDbl2 = !Y2
            !Y1 = temDbl2
            !Y2 = temDbl1
            .Update
            .MoveNext
        Wend
        .Close
        
    End With
End Sub

Private Sub Command2_Click()
    Dim rsTem As New ADODB.Recordset
    Dim rsTemA As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIxValueCategory.IxValueCategory, tblIxValue.* FROM tblIxValue LEFT JOIN tblIxValueCategory ON tblIxValue.IxValueCategoryID = tblIxValueCategory.IxValueCategoryID"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            While .EOF = False
                If rsTemA.State = 1 Then rsTemA.Close
                temSql = "Select * from tblCategory where Category = '" & !IxValueCategory & "'"
                rsTemA.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If rsTemA.RecordCount > 0 Then
                    !IxValueCategoryID = rsTemA!CategoryID
                    .Update
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub Command3_Click()
    Dim rsPurchaseBill As New ADODB.Recordset
    With rsPurchaseBill
        If .State = 1 Then .Close
        temSql = "Select * from tblPurchaseBill "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !CompletedDate = Format(!CompletedDateTime, "dd MMMM yyyy")
            !CompletedTime = Format(!CompletedDateTime, "HH:MM AMPM")
            .Update
            .MoveNext
        Wend
    End With
End Sub

Private Sub Command4_Click()
    
End Sub

Private Sub Form_Load()
    txtWHandle.Text = Me.hwnd
End Sub

Private Sub Timer1_Timer()
    txtWHandle.Text = Me.hwnd
    'txtActive.Text = GetActiveWindow()
    txtActive.Text = MDIMain.ActiveForm.hwnd
End Sub
