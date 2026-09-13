VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTetenusToxoidShedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetenus Toxoid Shedule"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
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
   ScaleHeight     =   4830
   ScaleWidth      =   6180
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2040
      Width           =   3735
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox txtScaleHeight 
      Height          =   360
      Left            =   1440
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtScaleWidth 
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Close"
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
   Begin MSComCtl2.DTPicker dtpTT1 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62521347
      CurrentDate     =   39897
   End
   Begin MSComCtl2.DTPicker dtpTT2 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62521347
      CurrentDate     =   39897
   End
   Begin MSComCtl2.DTPicker dtpTT3 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62521347
      CurrentDate     =   39897
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print"
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
   Begin VB.Label Label9 
      Caption         =   "Printer"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Paper"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Width"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Height"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "inches"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "inches"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Third Dose"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Second Dose"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "First Dose"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTetenusToxoidShedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
    Dim temSql As String
    
    Dim rsRfList As New ADODB.Recordset
    
    Dim i As Integer

Private Sub GetSettings()
    txtScaleHeight.Text = GetSetting(App.EXEName, Me.Name, "ScaleHeight", 0)
    txtScaleWidth.Text = GetSetting(App.EXEName, Me.Name, "ScaleWidth", 0)
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPrinter_Change
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub

Private Sub btnClose_Click()
    Unload Me
    
End Sub

Private Sub cmbPaper_Click()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                If cmbPaper.Text = PtrCtoVbString(.pName) Then
                    txtScaleHeight.Text = Format(.Size.cy / 1000 * mmtoinches, "0.00")
                    txtScaleWidth.Text = Format(.Size.cx / 1000 * mmtoinches, "0.00")
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If

End Sub

Private Sub dtpTT1_Change()
    dtpTT2.Value = dtpTT1.Value + (7 * 6)
    If Month(dtpTT2.Value) > 6 Then
        dtpTT3.Value = DateSerial(Year(dtpTT2.Value) + 1, Month(dtpTT2.Value) - 6, Day(dtpTT2.Value))
    Else
        dtpTT3.Value = DateSerial(Year(dtpTT2.Value), Month(dtpTT2.Value) + 6, Day(dtpTT2.Value))
    End If
End Sub

Private Sub dtpTT2_Change()
    dtpTT1.Value = dtpTT2.Value - (7 * 6)
    If Month(dtpTT2.Value) > 6 Then
        dtpTT3.Value = DateSerial(Year(dtpTT2.Value) + 1, Month(dtpTT2.Value) - 6, Day(dtpTT2.Value))
    Else
        dtpTT3.Value = DateSerial(Year(dtpTT2.Value), Month(dtpTT2.Value) + 6, Day(dtpTT2.Value))
    End If
End Sub



Private Sub cmbPrinter_Change()
    cmbPrinter_Click
End Sub

Private Sub cmbPrinter_Click()
    'cmbPaper.Clear
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = PrescreptionPaperHeight
            .cy = PrescreptionPaperWidth
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
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
    Call FillPrinters
    Call GetSettings
    dtpTT1.Value = Date
    dtpTT1_Change
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    GetCommonSettings Me
    
    
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
    SaveSetting App.EXEName, Me.Name, "ScaleHeight", Val(txtScaleHeight.Text)
    SaveSetting App.EXEName, Me.Name, "ScaleWidth", Val(txtScaleWidth.Text)
End Sub



Private Sub btnPrint_Click()
    
        
    
    Dim temText As String
    Dim rstemReportItems As New ADODB.Recordset
    
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long
    
    
    Dim w As New WshNetwork
    w.SetDefaultPrinter (cmbPrinter.Text)
    
    
    If SelectForm(cmbPaper.Text, Me.hdc) <> 1 Then
        MsgBox "Printer Error"
        Exit Sub
    End If
    Printer.PrintQuality = vbPRPQHigh
    Printer.ScaleWidth = Val(txtScaleWidth.Text) * 1440
    Printer.ScaleHeight = Val(txtScaleHeight.Text) * 1440
    
    
    Printer.FillStyle = vbFSTransparent
    
    With rstemReportItems
        If .State = 1 Then .Close
        temSql = "Select * from tblPrintItem  where Deleted = false AND PrintItemCategoryID = 9"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            TemX1 = Printer.ScaleWidth * !X1
            TemX2 = Printer.ScaleWidth * !X2
            TemY1 = Printer.ScaleHeight * !Y1
            TemY2 = Printer.ScaleHeight * !Y2
            If IsNull(!CircleRadius) = False Then
                temRadius = Printer.ScaleWidth * !CircleRadius
            End If
            If !IsLine = True Then
                Printer.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour
            ElseIf !IsRectangle = True Then
                Printer.Line (TemX1, TemY1)-(TemX2, TemY2), !ForeColour, B
            ElseIf !IsCircle = True Then
                Printer.Circle (TemX1, TemY1), temRadius, !ForeColour
            ElseIf !IsLabel = True Then
                temText = Format(!ItemText, "")
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            ElseIf !IsText = True Then
                Select Case !ItemText
                    Case "Patient Name": temText = CurrentPatient.NameWithTitle
                    Case "Patient ID": temText = CurrentPatient.ID
                    Case "Patient Age in Words": temText = CurrentPatient.AgeInWords
                    Case "Patient Data of Birth": temText = CurrentPatient.DateOfBirth
                    Case "Patient Sex": temText = CurrentPatient.Sex
                    Case "Patient Civil Status": temText = ""
                    Case "Patient Address": temText = CurrentPatient.Address
                    Case "Patient Telephone": temText = CurrentPatient.Telephone
                    Case "Date (dd mm yy)": temText = Format(CurrentVisit.VisitDate, "dd mm yy")
                    Case "Date (dd mmm yy)": temText = Format(CurrentVisit.VisitDate, "dd mmm yy")
                    Case "Date (dd mmm yyyy)": temText = Format(CurrentVisit.VisitDate, "dd mmm yyyy")
                    Case "Date (dd mmmm yyyy)": temText = Format(CurrentVisit.VisitDate, "dd mmmm yyyy")
                    Case "Date (dd mm yyyy)": temText = Format(CurrentVisit.VisitDate, "dd mm yyy")
                    Case "Date (yy mm dd)": temText = Format(CurrentVisit.VisitDate, "yy mm dd")
                    Case "Date (yy mmm dd)": temText = Format(CurrentVisit.VisitDate, "yy mmm dd")
                    Case "Date (yyyy mmm dd)": temText = Format(CurrentVisit.VisitDate, "yyyy mmm dd")
                    Case "Date (yyyy mmmm dd)": temText = Format(CurrentVisit.VisitDate, "yyyy mmmm dd")
                    Case "Date (yyyy mm dd)": temText = Format(CurrentVisit.VisitDate, "yyyy mm dd")
                    Case "Short Time": temText = Format(CurrentVisit.VisitTime, "h:m AMPM")
                    Case "Long Time": temText = Format(CurrentVisit.VisitTime, "HH:MM AMPM")
                    Case "Tetenus Toxoid Shedule":
                            temText = "1st Dose : " & Format(dtpTT1.Value, LongDateFormat) & vbNewLine
                            temText = temText & "2nd Dose : " & Format(dtpTT2.Value, LongDateFormat) & vbNewLine
                            temText = temText & "3rd Dose : " & Format(dtpTT3.Value, LongDateFormat) & vbNewLine
                    Case Else
                        temText = ""
                End Select
                
'    cmbText.AddItem "Tetenus Toxoid Shedule"
'    cmbText.AddItem "DMPA Shedule"
                
                
'                If PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, True) = False Then
'                    Printer.KillDoc
'                    Exit Sub
'                Else
'                    Printer.KillDoc
'                End If
                    
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            
            End If
            .MoveNext
        Wend
        .Close
    End With
    
    
    Printer.EndDoc
    
    
End Sub

