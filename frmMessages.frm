VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Messages"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
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
   ScaleHeight     =   7350
   ScaleWidth      =   8895
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   6600
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtColour 
      Height          =   360
      Left            =   2880
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkStrikeThrough 
      Caption         =   "ST"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFontSize 
      Height          =   360
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtFontName 
      Height          =   360
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin btButtonEx.ButtonEx btnClear 
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Clear"
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
   Begin VB.TextBox txtToSend 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Width           =   8535
   End
   Begin MSDataListLib.DataCombo cmbStaff 
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   4215
      _ExtentX        =   7435
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   6720
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
   Begin btButtonEx.ButtonEx btnSend 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Send"
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
   Begin btButtonEx.ButtonEx btnFont 
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Font"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfMessages 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8493
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMessages.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "Received Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub FillMessages()
    Dim rsMessage As New ADODB.Recordset
    Dim temRTF As String
    
    Dim StartLen() As Long
    Dim ThisLen() As Long
    Dim FontName() As String
    Dim FontSize() As Long
    Dim FontBold() As Boolean
    Dim FontItalic() As Boolean
    Dim FontUnderline() As Boolean
    Dim FontStrikeThrough() As Boolean
    Dim FontColour() As Long
    Dim i As Integer

    With rsMessage
        temSql = "SELECT tblFromStaff.Staff, tblToStaff.Staff, tblMessage.* FROM tblStaff AS tblFromStaff RIGHT JOIN (tblStaff AS tblToStaff RIGHT JOIN tblMessage ON tblToStaff.StaffID = tblMessage.ToID) ON tblFromStaff.StaffID = tblMessage.FromID where Cleared = False and (ToID = " & UserID & " OR FromID = " & UserID & ") Order by MessageID"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            ReDim StartLen(.RecordCount) As Long
            ReDim ThisLen(.RecordCount) As Long
            ReDim FontName(.RecordCount) As String
            ReDim FontSize(.RecordCount) As Long
            ReDim FontBold(.RecordCount) As Boolean
            ReDim FontItalic(.RecordCount) As Boolean
            ReDim FontUnderline(.RecordCount) As Boolean
            ReDim FontStrikeThrough(.RecordCount) As Boolean
            ReDim FontColour(.RecordCount) As Long
            
            If .RecordCount <> TotalMessageCount Then
                .MoveFirst
                Me.ZOrder 0
                rtfMessages.Text = Empty
                i = 0
                While .EOF = False
                    StartLen(i) = Len(rtfMessages.Text)
                    If rtfMessages.Text = Empty Then
                        rtfMessages.Text = ![tblFromStaff.staff] & " to " & ![tblToStaff.staff] & ": " & !Message
                    Else
                        rtfMessages.Text = rtfMessages.Text & vbNewLine & ![tblFromStaff.staff] & " to " & ![tblToStaff.staff] & ": " & !Message
                    End If
                    ThisLen(i) = Len(rtfMessages.Text) - StartLen(i)
                    FontName(i) = !FontName
                    FontBold(i) = !FontBold
                    FontItalic(i) = !FontItalic
                    FontColour(i) = !FontColour
                    FontSize(i) = !FontSize
                    FontStrikeThrough(i) = !FontStrikeThrough
                    FontUnderline(i) = !FontUnderline
                    i = i + 1
                    .MoveNext
                Wend
                For i = 0 To .RecordCount - 1
                    rtfMessages.SelStart = StartLen(i)
                    rtfMessages.SelLength = ThisLen(i)
                    rtfMessages.SelFontName = FontName(i)
                    rtfMessages.SelFontSize = FontSize(i)
                    rtfMessages.SelBold = FontBold(i)
                    rtfMessages.SelItalic = FontItalic(i)
                    rtfMessages.SelUnderline = FontUnderline(i)
                    rtfMessages.SelStrikeThru = FontStrikeThrough(i)
                    rtfMessages.SelColor = FontColour(i)
                Next
                TotalMessageCount = .RecordCount
            End If
        End If
        .Close
    End With
    
End Sub

Private Sub GetSettings()
    txtFontName.Text = GetSetting(App.EXEName, UserID, "FontName", "Arial")
    chkBold.Value = GetSetting(App.EXEName, UserID, "FontBold", 0)
    chkItalic.Value = GetSetting(App.EXEName, UserID, "FontItalic", 0)
    chkStrikeThrough.Value = GetSetting(App.EXEName, UserID, "FontStrikeThrough", 0)
    chkUnderline.Value = GetSetting(App.EXEName, UserID, "FontUnderline", 0)
    txtFontSize.Text = GetSetting(App.EXEName, UserID, "FontSize", 12)
    txtColour.Text = GetSetting(App.EXEName, UserID, "FontColour", vbRed)
    cmbStaff.BoundText = GetSetting(App.EXEName, UserID, "StaffID", 1)
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, UserID, "FontName", txtFontName.Text
    SaveSetting App.EXEName, UserID, "FontBold", chkBold.Value
    SaveSetting App.EXEName, UserID, "FontItalic", chkItalic.Value
    SaveSetting App.EXEName, UserID, "FontStrikeThrough", chkStrikeThrough.Value
    SaveSetting App.EXEName, UserID, "FontUnderline", chkUnderline.Value
    SaveSetting App.EXEName, UserID, "FontSize", txtFontSize.Text
    SaveSetting App.EXEName, UserID, "FontColour", txtColour.Text
    SaveSetting App.EXEName, UserID, "StaffID", cmbStaff.BoundText
    SaveCommonSettings Me
    
    
End Sub

Private Sub btnClear_Click()
    Dim rsMessage As New ADODB.Recordset
    With rsMessage
        temSql = "SELECT tblStaff.Staff, tblMessage.* FROM tblMessage LEFT JOIN tblStaff ON tblMessage.FromID = tblStaff.StaffID where Cleared = False and (ToID = " & UserID & " OR FromID = " & UserID & ")"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            While .EOF = False
                !Cleared = True
                .Update
                .MoveNext
            Wend
        End If
        .Close
    End With
    rtfMessages.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnColour_Click()
    CommonDialog1.color = txtColour.Text
    CommonDialog1.ShowColor
    txtColour.Text = CommonDialog1.color
    Call TextFont
End Sub

Private Sub btnFont_Click()
    CommonDialog1.FontName = txtFontName.Text
    CommonDialog1.FontSize = txtFontSize.Text
    CommonDialog1.color = txtColour.Text
    If chkBold.Value = 1 Then
        CommonDialog1.FontBold = True
    Else
        CommonDialog1.FontBold = False
    End If
    If chkItalic.Value = 1 Then
        CommonDialog1.FontItalic = True
    Else
        CommonDialog1.FontItalic = False
    End If
    If chkUnderline.Value = 1 Then
        CommonDialog1.FontUnderline = True
    Else
        CommonDialog1.FontUnderline = False
    End If
    If chkStrikeThrough.Value = 1 Then
        CommonDialog1.FontStrikethru = True
    Else
        CommonDialog1.FontStrikethru = False
    End If
    CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
    CommonDialog1.ShowFont
    If CommonDialog1.FontBold = True Then
        chkBold.Value = 1
    Else
        chkBold.Value = 0
    End If
    If CommonDialog1.FontItalic = True Then
        chkItalic.Value = 1
    Else
        chkItalic.Value = 0
    End If
    If CommonDialog1.FontUnderline = True Then
        chkUnderline.Value = 1
    Else
        chkUnderline.Value = 0
    End If
    If CommonDialog1.FontStrikethru = True Then
        chkStrikeThrough.Value = 1
    Else
        chkStrikeThrough.Value = 0
    End If
    txtFontName.Text = CommonDialog1.FontName
    txtFontSize.Text = CommonDialog1.FontSize
    txtColour.Text = CommonDialog1.color
    Call TextFont
    
End Sub

Private Sub btnSend_Click()
    If IsNumeric(cmbStaff.BoundText) = False Then
        MsgBox "Please select a user to send the message"
        cmbStaff.SetFocus
        Exit Sub
    End If

    Dim rsMessage As New ADODB.Recordset
    With rsMessage
        temSql = "SELECT tblMessage.* From tblMessage "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ToID = Val(cmbStaff.BoundText)
        !FromID = UserID
        !Message = txtToSend.Text
        !FontName = txtFontName.Text
        If chkBold.Value = 1 Then
            !FontBold = True
        Else
            !FontBold = False
        End If
        If chkItalic.Value = 1 Then
            !FontItalic = True
        Else
            !FontItalic = False
        End If
        If chkUnderline.Value = 1 Then
            !FontUnderline = True
        Else
            !FontUnderline = False
        End If
        If chkStrikeThrough.Value = 1 Then
            !FontStrikeThrough = True
        Else
            !FontStrikeThrough = False
        End If
        !FontColour = txtColour.Text
        !FontSize = txtFontSize.Text
        .Update
        .Close
    End With
    txtToSend.Text = Empty
    txtToSend.SetFocus
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
    
    
    Call FillCombos
    Call GetSettings
    Call FillInitialMessages
    Call TextFont
    Call SetColours
End Sub

Private Sub FillCombos()
    Dim Staff As New clsFillCombos
    Staff.FillAnyCombo cmbStaff, "staff", True
End Sub

Private Sub TextFont()
    txtToSend.Font.Name = txtFontName.Text
    If chkBold.Value = 1 Then
        txtToSend.Font.Bold = True
    Else
        txtToSend.Font.Bold = False
    End If
    If chkItalic.Value = 1 Then
        txtToSend.Font.Italic = True
    Else
        txtToSend.Font.Italic = False
    End If
    If chkUnderline.Value = 1 Then
        txtToSend.Font.Underline = True
    Else
        txtToSend.Font.Underline = False
    End If
    If chkStrikeThrough.Value = 1 Then
        txtToSend.Font.Strikethrough = True
    Else
        txtToSend.Font.Strikethrough = False
    End If
    txtToSend.Font.Size = txtFontSize.Text
    txtToSend.ForeColor = txtColour.Text
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub Timer1_Timer()
    Call FillMessages
End Sub

Private Sub FillInitialMessages()
    Dim rsMessage As New ADODB.Recordset
    Dim temRTF As String
    
    Dim StartLen() As Long
    Dim ThisLen() As Long
    Dim FontName() As String
    Dim FontSize() As Long
    Dim FontBold() As Boolean
    Dim FontItalic() As Boolean
    Dim FontUnderline() As Boolean
    Dim FontStrikeThrough() As Boolean
    Dim FontColour() As Long
    Dim i As Integer

    With rsMessage
        temSql = "SELECT tblFromStaff.Staff, tblToStaff.Staff, tblMessage.* FROM tblStaff AS tblFromStaff RIGHT JOIN (tblStaff AS tblToStaff RIGHT JOIN tblMessage ON tblToStaff.StaffID = tblMessage.ToID) ON tblFromStaff.StaffID = tblMessage.FromID where Cleared = False and (ToID = " & UserID & " OR FromID = " & UserID & ") Order by MessageID"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            ReDim StartLen(.RecordCount) As Long
            ReDim ThisLen(.RecordCount) As Long
            ReDim FontName(.RecordCount) As String
            ReDim FontSize(.RecordCount) As Long
            ReDim FontBold(.RecordCount) As Boolean
            ReDim FontItalic(.RecordCount) As Boolean
            ReDim FontUnderline(.RecordCount) As Boolean
            ReDim FontStrikeThrough(.RecordCount) As Boolean
            ReDim FontColour(.RecordCount) As Long
            .MoveFirst
            Me.ZOrder 0
            rtfMessages.Text = Empty
            i = 0
            While .EOF = False
                StartLen(i) = Len(rtfMessages.Text)
                If rtfMessages.Text = Empty Then
                    rtfMessages.Text = ![tblFromStaff.staff] & " to " & ![tblToStaff.staff] & ": " & !Message
                Else
                    rtfMessages.Text = rtfMessages.Text & vbNewLine & ![tblFromStaff.staff] & " to " & ![tblToStaff.staff] & ": " & !Message
                End If
                ThisLen(i) = Len(rtfMessages.Text) - StartLen(i)
                FontName(i) = !FontName
                FontBold(i) = !FontBold
                FontItalic(i) = !FontItalic
                FontColour(i) = !FontColour
                FontSize(i) = !FontSize
                FontStrikeThrough(i) = !FontStrikeThrough
                FontUnderline(i) = !FontUnderline
                i = i + 1
                .MoveNext
            Wend
            For i = 0 To .RecordCount - 1
                rtfMessages.SelStart = StartLen(i)
                rtfMessages.SelLength = ThisLen(i)
                rtfMessages.SelFontName = FontName(i)
                rtfMessages.SelFontSize = FontSize(i)
                rtfMessages.SelBold = FontBold(i)
                rtfMessages.SelItalic = FontItalic(i)
                rtfMessages.SelUnderline = FontUnderline(i)
                rtfMessages.SelStrikeThru = FontStrikeThrough(i)
                rtfMessages.SelColor = FontColour(i)
            Next
            TotalMessageCount = .RecordCount
        End If
        .Close
    End With
End Sub
