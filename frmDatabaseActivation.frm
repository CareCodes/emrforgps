VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDatabaseActivation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Activation"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
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
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHospitalNo 
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtInsID 
      Height          =   495
      Left            =   600
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtDocID 
      Height          =   495
      Left            =   1080
      TabIndex        =   38
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtOtherUserNames 
      Height          =   495
      Left            =   1560
      TabIndex        =   32
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFrameNo 
      Height          =   495
      Left            =   2040
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin btButtonEx.ButtonEx btnNext 
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Next"
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
   Begin btButtonEx.ButtonEx btnExit 
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Exit"
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
   Begin btButtonEx.ButtonEx btnBack 
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Back"
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
   Begin VB.Frame frameActivation 
      Caption         =   "Activation"
      Height          =   5295
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtA4 
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtA3 
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtA2 
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtA1 
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtR1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   840
         Width           =   520
      End
      Begin VB.TextBox txtR2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   840
         Width           =   520
      End
      Begin VB.TextBox txtR3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   840
         Width           =   520
      End
      Begin VB.TextBox txtR4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   840
         Width           =   520
      End
      Begin VB.Label Label10 
         Caption         =   "Registration Key"
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Enter Activation Key"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Details"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtRepeataPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   30
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtInstitutionAddress 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   2280
         Width           =   5415
      End
      Begin VB.TextBox txtInstitution 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox txtRegistration 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox txtDesignation 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo cmbTitle 
         Height          =   360
         Left            =   1920
         TabIndex        =   33
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label11 
         Caption         =   "Repeat Password"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Institution Address"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Institution Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "SLMC Reg. No"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Designation"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Doctor Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frameActivationFailed 
      Caption         =   "Activation Failed"
      Height          =   5295
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   7575
      Begin VB.Label Label12 
         Caption         =   "You have successfully activated the Electronic Medical Assistant. You will also be able to "
         Height          =   4575
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.Frame frameActivationSuccessfull 
      Caption         =   "Activation Successfull"
      Height          =   5295
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   7575
      Begin VB.Label Label6 
         Caption         =   "You have successfully activated the Electronic Medical Assistant. You will also be able to "
         Height          =   4575
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmDatabaseActivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim FSys As New FileSystemObject
    Dim MyDrive As Drive
    
Private Sub ButtonEx1_Click()
    End
End Sub

Private Sub btnBack_Click()
    txtFrameNo.Text = Val(txtFrameNo.Text) - 1
End Sub

Private Sub btnNext_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temDocID As Long
    Dim temInsID As Long
    Dim temnum As Long
    
    If Val(txtFrameNo.Text) = 1 Then
        If Trim(txtUserName.Text) = Empty Then
            MsgBox "Please enter a username"
            txtUserName.SetFocus
            Exit Sub
        End If
        If Trim(txtPassword.Text) = Empty Then
            MsgBox "Please enter a password"
            txtUserName.SetFocus
            Exit Sub
        End If
        If txtPassword.Text <> txtRepeataPassword.Text Then
            MsgBox "The password and re-entered password are different"
            txtRepeataPassword.SetFocus
            Exit Sub
        End If
        If NoSuchUserName = False Then
            MsgBox "This username already exists. Please enter something else"
            txtUserName.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
            Exit Sub
        End If
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblStaff where StaffID = " & Val(txtDocID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
            
            Else
                .AddNew
            End If
            !Staff = txtName.Text
            !TitleID = Val(cmbTitle.Text)
            !SLMCRegNo = txtRegistration.Text
            !Designation = txtDesignation.Text
            !UserName = EncreptedWord(txtUserName.Text)
            !Password = EncreptedWord(txtPassword.Text)
            !IsADoctor = True
            !IsAUser = True
            !IsOwner = True
            .Update
            txtDocID.Text = !StaffID
            .Close
            temSql = "Select * from tblInstitution where InstitutionID = " & Val(txtInsID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                
            Else
                .AddNew
            End If
            !Institution = txtInstitution.Text
            !Address = txtInstitutionAddress.Text
            !IsOwner = True
            .Update
            txtInsID.Text = !InstitutionID
            .Close
                    
            temDocID = Asc(Left(txtName.Text, 1))
            
            txtR1.Text = Format(temDocID, "0000")
            
            temSql = "Select * from tblDoctorNUmber where AddedYear = " & Year(Date) Mod 10 & " ANd AddedMonth = " & Month(Date) & " And AddedDate = " & Day(Date) & " AND DoctorID = " & (temDocID Mod 9) + 1
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                txtR2.Text = Format(!DoctorNo, "0000")
                txtHospitalNo.Text = !HospitalNo
            Else
                txtR2.Text = "0000"
                temInsID = 0
            End If
            .Close
            txtR3.Text = Format(Year(Date) + Month(Date) + Day(Date) + temDocID, "0000")
            temSql = "Select sum(DoctorNo) as SumOfDoc, sum(HospitalNo) as SumOfHos from tblDoctorNumber"
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                temnum = !SumOfDoc + !SumOfHos
                txtR4.Text = Format(Abs(temnum Mod 1000 - temDocID), "0000")
            Else
                txtR4.Text = "0000"
            End If
        End With
        
    ElseIf Val(txtFrameNo.Text) = 2 Then
        Dim temHosNum As String
        Dim temGivenNum As Long
        
        temGivenNum = Val(Left(txtA1.Text, 3))
        If temGivenNum Mod 11 <> 3 Then
            MsgBox "Wrong Activation Code"
            txtA1.Text = Empty
            txtA2.Text = Empty
            txtA3.Text = Empty
            txtA4.Text = Empty
            txtA1.SetFocus
            Exit Sub
        End If
        temGivenNum = Val(Left(txtA2.Text, 3))
        If temGivenNum Mod 13 <> 4 Then
            MsgBox "Wrong Activation Code"
            txtA1.Text = Empty
            txtA2.Text = Empty
            txtA3.Text = Empty
            txtA4.Text = Empty
            txtA1.SetFocus
            Exit Sub
        End If
        temGivenNum = Val(Left(txtA3.Text, 3))
        If temGivenNum Mod 17 <> 8 Then
            MsgBox "Wrong Activation Code"
            txtA1.Text = Empty
            txtA2.Text = Empty
            txtA3.Text = Empty
            txtA4.Text = Empty
            txtA1.SetFocus
            Exit Sub
        End If
        temGivenNum = Val(Left(txtA4.Text, 3))
        If temGivenNum Mod 19 <> 18 Then
            MsgBox "Wrong Activation Code"
            txtA1.Text = Empty
            txtA2.Text = Empty
            txtA3.Text = Empty
            txtA4.Text = Empty
            txtA1.SetFocus
            Exit Sub
        End If
        temHosNum = Right(txtA1.Text, 1) & Right(txtA2.Text, 1) & Right(txtA3.Text, 1) & Right(txtA4.Text, 1)
        If Val(txtHospitalNo.Text) <> Val(temHosNum) Then
            MsgBox "Wrong Activation Code"
            txtA1.Text = Empty
            txtA2.Text = Empty
            txtA3.Text = Empty
            txtA4.Text = Empty
            txtA1.SetFocus
            Exit Sub
        End If
        MsgBox "Activation Done"
        
    End If
    txtFrameNo.Text = Val(txtFrameNo.Text) + 1
End Sub

Private Sub txtA1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
        If Len(txtA1.Text) > 3 Then
            txtA2.SetFocus
        End If
    Else
        KeyCode = Empty
    End If

End Sub

Private Sub txtA2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
        If Len(txtA2.Text) > 3 Then
            txtA3.SetFocus
        End If
    Else
        KeyCode = Empty
    End If

End Sub

Private Sub txtA3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
        If Len(txtA3.Text) > 3 Then
            txtA4.SetFocus
        End If
    Else
        KeyCode = Empty
    End If

End Sub

Private Sub txtA4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
        If Len(txtA4.Text) > 3 Then
            btnNext.SetFocus
        End If
    Else
        KeyCode = Empty
    End If

End Sub


Private Sub txtA1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
    
    Else
        KeyCode = Empty
    End If

End Sub

Private Sub txtA2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
    
    Else
        KeyCode = Empty
    End If

End Sub

Private Sub txtA3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
    
    Else
        KeyCode = Empty
    End If

End Sub

Private Sub txtA4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    ElseIf KeyCode = 46 Then
    
    ElseIf KeyCode > 47 And KeyCode < 58 Then
    
    Else
        KeyCode = Empty
    End If

End Sub

Private Sub txtA1_Change()
    If Len(txtA1.Text) > 4 Then
        txtA2.Text = Right(txtA1.Text, (Len(txtA1.Text) - 4))
        txtA1.Text = Left(txtA1.Text, 4)
        txtA2.SetFocus
        SendKeys "{end}"
    End If
End Sub

Private Sub txtA2_Change()
    If Len(txtA2.Text) > 4 Then
        txtA3.Text = Right(txtA2.Text, (Len(txtA2.Text) - 4))
        txtA2.Text = Left(txtA2.Text, 4)
        txtA3.SetFocus
        SendKeys "{end}"
    End If
End Sub
Private Sub txtA3_Change()
    If Len(txtA3.Text) > 4 Then
        txtA4.Text = Right(txtA3.Text, (Len(txtA3.Text) - 4))
        txtA3.Text = Left(txtA3.Text, 4)
        txtA4.SetFocus
        SendKeys "{end}"
    End If
End Sub

Private Sub txtA4_Change()
    If Len(txtA4.Text) > 4 Then
        txtA4.Text = Left(txtA4.Text, 4)
        btnNext.SetFocus
    End If
End Sub

Private Sub txtFrameNo_Change()
    Select Case Val(txtFrameNo.Text)
        Case 1:     btnBack.Enabled = False
                    btnNext.Enabled = True
                    frameDetails.Visible = True
                    frameActivation.Visible = False
                    frameActivationSuccessfull.Visible = False
                    frameActivationFailed.Visible = False
        Case 2:     btnBack.Enabled = True
                    btnNext.Enabled = True
                    frameDetails.Visible = False
                    frameActivation.Visible = True
                    frameActivationSuccessfull.Visible = False
                    frameActivationFailed.Visible = False
        Case 3:     btnBack.Enabled = True
                    btnNext.Enabled = True
                    frameDetails.Visible = True
                    frameActivation.Visible = False
                    frameActivationSuccessfull.Visible = True
                    frameActivationFailed.Visible = False
        Case 4:     Unload Me
    End Select
End Sub

Private Function NoSuchUserName() As Boolean
    NoSuchUserName = True
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtDocID.Text) = False Then
            temSql = "Select * from tblStaff where Deleted = false"
        Else
            temSql = "Select * from tblStaff where Deleted = false AND StaffID <> " & Val(txtDocID.Text)
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If IsNull(!UserName) = False Then
                txtOtherUserNames.Text = DecreptedWord(!UserName)
                If UCase(txtUserName.Text) = UCase(txtOtherUserNames.Text) Then
                    NoSuchUserName = False
                    Exit Function
                End If
            End If
            .MoveNext
        Wend
        .Close
    End With
End Function

