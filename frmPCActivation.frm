VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPCActivation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Activation"
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
   Begin VB.CommandButton Command6 
      Caption         =   "Check"
      Height          =   495
      Left            =   480
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "A"
      Height          =   495
      Left            =   1800
      TabIndex        =   23
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "D"
      Height          =   495
      Left            =   3120
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "D"
      Height          =   495
      Left            =   3120
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A"
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   495
      Left            =   480
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   2520
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtHospitalNo 
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtInsID 
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtDocID 
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtOtherUserNames 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFrameNo 
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin btButtonEx.ButtonEx btnActivate 
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Activate"
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
      TabIndex        =   0
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
   Begin VB.Frame frameActivation 
      Caption         =   "Activation"
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtA4 
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtA3 
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtA2 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtA1 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   2280
         Width           =   520
      End
      Begin VB.TextBox txtR1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   520
      End
      Begin VB.TextBox txtR2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   840
         Width           =   520
      End
      Begin VB.TextBox txtR3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   840
         Width           =   520
      End
      Begin VB.TextBox txtR4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   840
         Width           =   520
      End
      Begin VB.Label Label10 
         Caption         =   "Registration Key"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Enter Activation Key"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPCActivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim FSys As New FileSystemObject
    Dim MyDrive As Drive

Private Sub btnActivate_Click()
    Dim temNo1 As Integer
    Dim temNo2 As Integer
    Dim temNo3 As Integer
    Dim temNo4 As Integer
    Dim temNo5 As Integer
    Dim temNo6 As Integer
    Dim temNo7 As Integer
    Dim temNo8 As Integer
        
    Dim HDDID As Long
    Dim temStr As String
    Dim CanActivate As Boolean
    
    CanActivate = True
    
    temStr = txtA1.Text & txtA2.Text & txtA3.Text & txtA4.Text
    
    Set MyDrive = FSys.GetDrive("C:\")
    HDDID = MyDrive.SerialNumber
    
    
    temNo1 = Val(Mid(temStr, 1, 2))
    temNo2 = Val(Mid(temStr, 3, 2))
    temNo3 = Val(Mid(temStr, 5, 2))
    temNo4 = Val(Mid(temStr, 7, 2))
    temNo5 = Val(Mid(temStr, 9, 2))
    temNo6 = Val(Mid(temStr, 11, 2))
    temNo7 = Val(Mid(temStr, 13, 2))
    temNo8 = Val(Mid(temStr, 15, 2))

    If HDDID Mod 97 <> temNo1 Then CanActivate = False
    If HDDID Mod 73 <> temNo2 Then CanActivate = False
    If HDDID Mod 89 <> temNo3 Then CanActivate = False
    If HDDID Mod 71 <> temNo4 Then CanActivate = False
    If HDDID Mod 83 <> temNo5 Then CanActivate = False
    If HDDID Mod 67 <> temNo6 Then CanActivate = False
    If HDDID Mod 79 <> temNo7 Then CanActivate = False
    If HDDID Mod 61 <> temNo8 Then CanActivate = False
    
    If CanActivate = True Then
        MsgBox "Activated"
    Else
        MsgBox "Not Activated"
    End If

End Sub

Private Sub btnExit_Click()
    End
End Sub

Private Sub Command1_Click()
    MsgBox CheckDatabaseActivated
End Sub

Private Sub Command2_Click()
    ActivateDatabase
End Sub

Private Sub Command3_Click()
    Call DeactivateDatabase
End Sub

Private Sub Command4_Click()
    Call DeActivatePC
End Sub

Private Sub Command5_Click()
    Call ActivatePC
End Sub

Private Sub Command6_Click()
    MsgBox CheckPCActivated
End Sub

Private Sub Form_Load()
    Call GenerateRegistrationKey
End Sub

Private Sub GenerateRegistrationKey()
    Dim rsTem As New ADODB.Recordset
    Dim temDocID As Long
    Dim temInsID As Long
    Dim temnum As Long
    Dim HDDID As Long
    
    Dim HexHDDID As String
    Dim HexRndNo As String
    Dim HexLen As String
    
    Dim temStr1 As String
    Dim temStr2 As String
    Dim temStr As String
    
    Dim SNLen As Integer
    
'    If CheckDatabaseActivated = False Then
'        MsgBox "This database is NOT activated."
'        End
'    End If
    
    Set MyDrive = FSys.GetDrive("C:\")
    HDDID = MyDrive.SerialNumber

    HexHDDID = CStr(Hex(HDDID))
    
    SNLen = Len(HexHDDID)
    If SNLen > 15 Then SNLen = 15
    
    HexLen = CStr(Hex(SNLen))
    
    Randomize
    
    HexRndNo = Left(CStr(Hex(Rnd * 1000000000#)) & "1234567890", 15 - SNLen)

    temStr = HexLen & HexHDDID & HexRndNo

    txtR1.Text = Mid(temStr, 1, 4)
    txtR2.Text = Mid(temStr, 5, 4)
    txtR3.Text = Mid(temStr, 9, 4)
    txtR4.Text = Mid(temStr, 13, 4)

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
            btnActivate.SetFocus
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
        btnActivate.SetFocus
    End If
End Sub

