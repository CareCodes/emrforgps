VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDepositReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deposit Report"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
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
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   7440
      Width           =   3735
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   7440
      Width           =   3495
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   7440
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
   Begin VB.TextBox txtTotalIncome 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   8160
      TabIndex        =   6
      Top             =   6840
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GridIncome 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9975
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62259203
      CurrentDate     =   39859
   End
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   7440
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62259203
      CurrentDate     =   39859
   End
   Begin VB.Label Label5 
      Caption         =   "Paper"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total"
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Drawer"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmDepositReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim NumForms As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    Dim FSys As New Scripting.FileSystemObject
    Private CsetPrinter As New cSetDfltPrinter
    Dim i As Integer

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    Dim temTopic As String
    Dim temSubTopic As String
    
    temTopic = InstitutionName & " - Deposit Report"
    
    If dtpFrom.Value = dtpTo.Value Then
        temSubTopic = "Date : " & Format(dtpFrom.Value, LongDateFormat)
    Else
        temSubTopic = "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
    End If
    
    GetPrintDefaults ThisReportFormat
    ThisReportFormat.ReportPrintOrientation = Landscape
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    
    Dim MyPrinter As Printer
    
    For Each MyPrinter In Printers
        If MyPrinter.DeviceName = cmbPrinter.Text Then
            Set Printer = MyPrinter
        End If
    Next
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        GridPrint GridIncome, ThisReportFormat, temTopic, temSubTopic
        Printer.EndDoc
    End If
End Sub

Private Sub cmbDrawer_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub GetSettings()
    cmbDrawer.BoundText = Val(GetSetting(App.EXEName, Me.Name, "DrawerID", 1))
    GetCommonSettings Me
    
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, "DrawerID", Val(cmbDrawer.BoundText)
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
    SaveCommonSettings Me
    
End Sub

Private Sub cmbDrawer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbDrawer.Text = Empty
    End If
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
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
    Call FormatGrid
    Call SetColours
    Call FillCombos
    Call GetSettings
    Call FillGrid
    Call PopulatePrinters
    Call cmbPrinter_Change
End Sub

Private Sub FillCombos()
    Dim Doc As New clsFillCombos
    Doc.FillAnyCombo cmbDrawer, "Drawer", True
End Sub

Private Sub FormatGrid()
    With GridIncome
        .Clear
        .Cols = 5
        .Rows = 1
        
        .Row = 0
        
        
        .Col = 0
        .Text = "Date"
        
        .Col = 1
        .Text = "Time"
        
        .Col = 2
        .Text = "Category"
        
        .Col = 3
        .Text = "Comments"
        
        .Col = 4
        .Text = "Value"
        
        .ColWidth(0) = 1100
        .ColWidth(1) = 1100
        .ColWidth(2) = 3100
        .ColWidth(3) = 3100
        .ColWidth(4) = 1100
        
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    Dim TotalFee As Double
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbDrawer.BoundText) = True Then
            temSql = "SELECT tblDeposit.DepositDate, tblDeposit.DepositTime, tblCategory.Category, tblDeposit.Comments, tblDeposit.DepositValue " & _
                        "FROM tblDeposit LEFT JOIN tblCategory ON tblDeposit.DepositCategoryID = tblCategory.CategoryID " & _
                        "WHERE (((tblDeposit.Deleted)=False) AND ((tblDeposit.DepositDate) Between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# AND #" & Format(dtpTo.Value, "dd MMMM yyyy") & "#) AND ((tblDeposit.ToDrawerID)=" & Val(cmbDrawer.BoundText) & ")) " & _
                        "ORDER BY tblDeposit.DepositID"
                        
        Else
            temSql = "SELECT tblDeposit.DepositDate, tblDeposit.DepositTime, tblCategory.Category, tblDeposit.Comments, tblDeposit.DepositValue " & _
                        "FROM tblDeposit LEFT JOIN tblCategory ON tblDeposit.DepositCategoryID = tblCategory.CategoryID " & _
                        "WHERE (((tblDeposit.Deleted)=False) AND ((tblDeposit.DepositDate) Between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# AND #" & Format(dtpTo.Value, "dd MMMM yyyy") & "#)) " & _
                        "ORDER BY tblDeposit.DepositID"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
        
            If IsNull(!DepositValue) = False Then
                TotalFee = TotalFee + !DepositValue
            End If
            
            GridIncome.Rows = GridIncome.Rows + 1
            GridIncome.Row = GridIncome.Rows - 1
            
            GridIncome.Col = 0
            GridIncome.Text = !DepositDate
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 1
            If IsNull(!DepositTime) = False Then
                GridIncome.Text = Format(!DepositTime, "HH:MM AMPM")
            End If
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 2
            If IsNull(!Category) = False Then
                GridIncome.Text = !Category
            Else
                GridIncome.Text = "Not Seen"
            End If
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 3
            GridIncome.Text = Format(!Comments, "")
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 4
            If IsNull(!DepositValue) = False Then
                GridIncome.Text = Format(![DepositValue], "0.00")
            End If
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            .MoveNext
        Wend
    End With

txtTotalIncome.Text = Format(TotalFee, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub GridIncome_Click()
    With GridIncome
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub GridIncome_DblClick()
    Dim PtID As Long
    Dim VtID As Long
    With GridIncome
        PtID = .TextMatrix(.Row, 5)
        VtID = .TextMatrix(.Row, 4)
        CurrentPatient.ID = PtID
        CurrentVisit.VisitID = VtID
        If SingleVisit = True Then
            Unload frmPatient
            Unload frmPatientDetails
            Unload frmNewVisit
            frmNewVisit.Show
        Else
            Dim MyNewVisit As New frmNewVisit
            MyNewVisit.Show
        End If
    End With
End Sub

Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers()
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = PrescreptionPaperHeight
'            .cy = PrescreptionPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub


