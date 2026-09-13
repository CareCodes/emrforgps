VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSelectDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Database"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
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
   ScaleHeight     =   2070
   ScaleWidth      =   5340
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDatabase 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin btButtonEx.ButtonEx bttnSelectDatabasePath 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Select Database"
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
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
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
   Begin VB.Label Label1 
      Caption         =   "Database"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmSelectDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject

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
    
    Call SetColours
    Call SetPreferances
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SetPreferances()
    Dim TemResponce As Integer
    If FSys.FileExists(Database) = True Then
        txtDatabase.Text = Database
    Else
        txtDatabase.Text = "You have not selected a valid database"
        txtDatabase.ForeColor = vbYellow
        txtDatabase.BackColor = vbRed
    End If
End Sub


Private Sub SavePreferancesToFile()
    SaveSetting App.EXEName, "Options", "Database", txtDatabase.Text
End Sub

Private Sub SavePreferancesToMemory()
    Database = txtDatabase.Text
End Sub

Private Sub bttnSelectDatabasePath_Click()
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "EMA 3 Database|EMA3.mdb"
    On Error Resume Next
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDatabase.Text = CommonDialog1.filename
        SaveSetting App.EXEName, "Options", "Database", txtDatabase.Text
'        Unload Me
    Else
        MsgBox "You have not selected valid database. The program may not function", vbCritical, "No database"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    
    Dim TemResponce As Integer
    Dim constr As String
    If FSys.FileExists(txtDatabase.Text) = False Then
        MsgBox "You have not selected a valid database", vbCritical, "Database?"
        Cancel = True
        txtDatabase.SetFocus
        On Error Resume Next: SendKeys "{home}+{end}"
    Else
        
        Call SavePreferancesToFile
        Call SavePreferancesToMemory
    
        Dim MyForm As Form
        
        For Each MyForm In Forms
            If MyForm.Name <> Me.Name And MyForm.Name <> "MDIMain" Then Unload MyForm
        Next
        
        If cnnStores.State = 1 Then cnnStores.Close
        constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Database & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=REDACTED"
        cnnStores.Open constr
        DataEnvironment1.deCnnStores.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Database & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=REDACTED"    ' "data source=" & constr  'GetSetting(App.EXEName, "Options", "Database", App.Path & "\hospital.mdb")
        
        MsgBox "Database Changed"
    
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'MsgBox "The change of the database will be active after the next restart of the program"
End Sub

