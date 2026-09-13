VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmIxReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
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
   ScaleHeight     =   9450
   ScaleWidth      =   11940
   Begin VB.TextBox txtSpecimanNo 
      Height          =   360
      Left            =   9480
      TabIndex        =   22
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   8880
      Width           =   2895
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   8880
      Width           =   3255
   End
   Begin MSDataListLib.DataCombo ListTitle 
      Height          =   360
      Left            =   1800
      TabIndex        =   17
      Top             =   8400
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Frame FrameDetails 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   11655
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   375
         Left            =   9720
         TabIndex        =   16
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62324738
         CurrentDate     =   39812
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   6600
         TabIndex        =   15
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   62324739
         CurrentDate     =   39812
      End
      Begin VB.Label lblInstitution 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   28
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label lblDoctor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   27
         Top             =   660
         Width           =   3615
      End
      Begin VB.Label lblPatient 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblID 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   6
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   13
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   1
         Top             =   660
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Referring Doctor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Referring Institute"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Patient"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   0
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7320
      Top             =   1680
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   8880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx bttPrint 
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   8880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   8880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin MSDataListLib.DataCombo cmbIx 
      Height          =   360
      Left            =   2160
      TabIndex        =   23
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin VB.TextBox txtScaleWidth 
      Height          =   360
      Left            =   5040
      TabIndex        =   29
      Top             =   8880
      Width           =   855
   End
   Begin VB.TextBox txtScaleHeight 
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      Top             =   8880
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Investigation Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Speciman No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Paper"
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Printer"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   8880
      Width           =   1575
   End
End
Attribute VB_Name = "frmIxReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

    Public myPATIENTiXid As Long
    
    Dim MyIx As New clsIx

    Dim NumForms As Long, i As Long
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
    
   
    Dim rsViewPatient As New ADODB.Recordset
    Dim rsIx As New ADODB.Recordset
    Dim temSql As String


Private Sub bttnChange_Click()
    Call ChangeData
    bttPrint.Enabled = True
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub
Private Sub bttPrint_Click()
    Dim rsPatietnIX As New ADODB.Recordset
    With rsPatietnIX
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientIX where PatientIxID = " & myPATIENTiXid
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Printed = True
            !PrintedDate = Date
            !PrintedTime = Time
            !PrintedUserID = UserID
            .Update
            .Close
        Else
            .Close
            MsgBox "Error"
            Exit Sub
        End If
    End With

    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        If MyPrinter.DeviceName = cmbPrinter.Text Then
            Set MyPrinter = Printer
        End If
    Next
    If SelectForm(cmbPaper.Text, Me.hdc) <> 1 Then
        MsgBox "Printer Error"
        Exit Sub
    End If
    
    'Printer.PrintQuality = vbPRPQHigh
    
    Printer.Print

    Call PrintReportFormat
    Call PrintIxFormat



    Printer.EndDoc
End Sub

Private Sub PrintIxFormat()
    Dim rstemPrintItems As New ADODB.Recordset
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long

    Dim temText As String

    Printer.FillStyle = vbFSTransparent

    With rstemPrintItems
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem  where Deleted = false  AND IxID = " & Val(cmbIx.BoundText)
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
                temText = !IxItem
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            ElseIf !IsValue = True Then
                temText = FindIxValue(!IxItemID)
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            ElseIf !IsCalc = True Then
                temText = FindCalcValue(!IxItemID)
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            End If
            .MoveNext
        Wend
        .Close
    End With
'    temText = "Software by Lanka Medical Programmers - 071 5812399"
'            TemX1 = Printer.ScaleWidth * 0.01
'            TemX2 = Printer.ScaleWidth * 0.99
'            TemY1 = Printer.ScaleHeight * 0.89
'            TemY2 = Printer.ScaleHeight * 0.99
'    PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, "Arial", 8, False, True, False, False, TextAlignment.CentreAlign, False)
End Sub

Private Function FindCalcValue(ByVal temIxItemID As Long) As String
    Dim rsTem As New ADODB.Recordset
    Dim temCalc As Double
    Dim temFunction As String
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItemCalc Where IxItemID = " & temIxItemID & " AND Deleted = false Order by ItemOrderNo"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If !ItemOrderNo = 1 Then
                If !FunctionType = "Value" Then
                    If IsNumeric(FindIxValue(!ValueIxItemID)) = True Then
                        temCalc = CDbl(FindIxValue(!ValueIxItemID))
                    Else
                        temCalc = 0
                    End If
                ElseIf !FunctionType = "Constant" Then
                    If IsNumeric(!ConstantValue) = True Then
                        temCalc = CDbl(!ConstantValue)
                    Else
                        temCalc = 0
                    End If
                End If
            Else
                Select Case !FunctionType
                    Case "Constant":
                        If IsNumeric(!ConstantValue) = True Then
                            temCalc = CalculateValue(temFunction, temCalc, !ConstantValue)
                        Else
                            temCalc = 0
                        End If

                    Case "Value":
                        If IsNumeric(FindIxValue(!ValueIxItemID)) = True Then
                            temCalc = CalculateValue(temFunction, temCalc, FindIxValue(!ValueIxItemID))
                        Else
                            temCalc = 0
                        End If
                    Case Else:
                        temFunction = !FunctionType
                End Select
            End If
            .MoveNext
        Wend
        FindCalcValue = temCalc
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem where IxItemID = " & temIxItemID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!ItemFormat) = False Then
                If !ItemFormat <> "" Then
                    FindCalcValue = Format(FindCalcValue, !ItemFormat)
                End If
            End If
            If IsNull(!ItemPrefix) = False Then
                If !ItemPrefix <> "" Then
                    FindCalcValue = !ItemPrefix & FindCalcValue
                End If
            End If
            If IsNull(!ItemSuffix) = False Then
                If !ItemSuffix <> "" Then
                    FindCalcValue = FindCalcValue & !ItemSuffix
                End If
            End If
        End If
        .Close
    End With
End Function

Private Function CalculateValue(ByVal FunctionType As String, ByVal StartingValue As Double, ByVal NewValue As Double) As Double
    Select Case FunctionType
        Case "Add": CalculateValue = StartingValue + NewValue
        Case "Subtract": CalculateValue = StartingValue - NewValue
        Case "Multiply": CalculateValue = StartingValue * NewValue
        Case "Divide": CalculateValue = StartingValue / NewValue
        Case Else:
    End Select
End Function


Private Function FindIxValue(ByVal temIxItemID As Long) As String
    With rsIx
        If .State = 1 Then .Close
        temSql = "SELECT * from tblPatientIxItem where PatientIxId = " & myPATIENTiXid & " AND IxItemID = " & temIxItemID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            FindIxValue = !Value
        Else
            FindIxValue = Empty
        End If
    End With
End Function

Private Sub PrintReportFormat()
    Dim rstemPrintItems As New ADODB.Recordset
    Dim TemX1 As Long
    Dim TemY1 As Long
    Dim TemX2 As Long
    Dim TemY2 As Long
    Dim temRadius As Long
    Dim temText As String

    Printer.FillStyle = vbFSTransparent

    With rstemPrintItems
        If .State = 1 Then .Close
        temSql = "Select * from tblPrintItem  where Deleted = false AND PrintItemCategoryID = 8"
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
                
                    Case "Investigation": temText = MyIx.Ix
                    Case "Investigation ID": temText = MyIx.IxID
                    Case "Investigation Comments": temText = MyIx.Comments
                    Case "Department": temText = MyIx.Department
                    Case "Speciman": temText = MyIx.Speciman
                    Case "Speciman No.": temText = txtSpecimanNo.Text
                    Case "Speciman Comments": temText = ""
'                    Case "Referring Doctor": temText = cmbDoctor.Text
'                    Case "Referring Institution": temText = cmbInstitution.Text
                    
                    Case Else
                        temText = ""
                End Select
                PrintOK = PrintReportText(temText, TemX1, TemX2, TemY1, TemY2, !FontName, !FontSize, !FontBold, !FontItalic, !FontUnderline, !FontStrikeThrough, !TextAlignment, False)
            End If
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub ChangeData()
'On Error GoTo ErrorHandler

    Dim MyCombo As Control
    Dim temIxItemID As Long

    For Each MyCombo In Controls
        If Left(MyCombo.Name, 5) = "Ixcmb" Then
            temIxItemID = Val(Right(MyCombo.Name, Len(MyCombo.Name) - 5))
            With rsIx
                If .State = 1 Then .Close
                temSql = "SELECT * from tblPatientIxItem where PatientIxId = " & myPATIENTiXid & " AND IxItemID = " & temIxItemID
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                     !Value = MyCombo.Text
                Else
                    .AddNew
                    !PatientID = CurrentPatient.ID
                    !IxID = Val(cmbIx.BoundText)
                    !PatientIxID = myPATIENTiXid
                    !IxItemID = temIxItemID
                    !Value = MyCombo.Text
                End If
                .Update
            End With
        End If
    Next

    With rsIx
        If .State = 1 Then .Close
        temSql = "SELECT * from tblPatientIx where PatientIxId = " & myPATIENTiXid
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
'            !PatientID = temPatientID
'            If IsNumeric(cmbDoctor.BoundText) Then !DoctorID = cmbDoctor.BoundText
'            If IsNumeric(cmbInstitution.BoundText) Then !InstitutionID = cmbInstitution.BoundText
'            !IxID = cmbIx.BoundText
            !SavedDate = dtpDate.Value
            !SavedTime = dtpTime.Value
            !SavedUserID = UserID
            !Saved = True
            !SpecimanNo = txtSpecimanNo.Text
            .Update
        End If
        .Close
    Exit Sub

ErrorHandler:
 i = MsgBox("An Unknown Error Occured, Please contact Fintec", vbCritical, "Error")
 .CancelUpdate
 .Close
End With
End Sub




Private Sub LocatePatient()
    lblSex.Caption = CurrentPatient.Sex
    lblAge.Caption = CurrentPatient.AgeInWords
    lblID.Caption = CurrentPatient.ID
    lblPatient.Caption = CurrentPatient.NameWithTitle
End Sub

Private Sub FillCombos()
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
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

Private Sub cmbPrinter_Change()
    cmbPrinter_Click
End Sub

Private Sub cmbPrinter_Click()
    cmbPaper.Clear
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
    Call FillCombos
    bttnChange.Enabled = True
    bttPrint.Enabled = False
    On Error Resume Next
    
    Call FillPrinters
    
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPrinter_Click
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub

Public Sub InvestigationDetails()
    With rsIx
        If .State = 1 Then .Close
        temSql = "SELECT * from tblPatientIx where PatientIxID = " & myPATIENTiXid
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        CurrentPatient.ID = !PatientID
        MyIx.ID = !IxID
        cmbIx.BoundText = !IxID
        dtpDate.Value = Date
        dtpTime.Value = Time
        If Not IsNull(!SpecimanNo) Then
            txtSpecimanNo.Text = !SpecimanNo
        End If
    End With
    Call LocatePatient
    Call FillPrinters
    Call AddCombos
    Call FillComboValues
End Sub


Private Sub FillComboValues()
    Dim MyCombo As Control
    Dim temIxItemID As Long
    For Each MyCombo In Controls
        If Left(MyCombo.Name, 5) = "Ixcmb" Then
            temIxItemID = Val(Right(MyCombo.Name, Len(MyCombo.Name) - 5))
            With rsIx
                If .State = 1 Then .Close
                temSql = "SELECT * from tblPatientIxItem where PatientIxId = " & myPATIENTiXid & " AND IxItemID = " & temIxItemID
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    MyCombo.Text = !Value
                Else
                    MyCombo.Text = FindDefaultValue(temIxItemID)
                End If
            End With
        End If
    Next
End Sub

Private Function FindDefaultValue(IxItemID As Long) As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIxValue.IxValue " & _
                    "FROM tblIxItem LEFT JOIN tblIxValue ON tblIxItem.IxValueDefaultID = tblIxValue.IxValueID " & _
                    "WHERE (((tblIxItem.IxItemID)=" & IxItemID & "))"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!IxValue) = False Then
                FindDefaultValue = !IxValue
            Else
                FindDefaultValue = Empty
            End If
        Else
            FindDefaultValue = Empty
        End If
    End With
End Function

Private Sub AddCombos()
    Dim MyCombo As Control
    Dim MyLabel As Label
    Dim rsTemIxItem As New ADODB.Recordset

    Dim LeftMargin As Long
    Dim TopMargin As Long
    Dim BottomMargin As Long
    Dim ComboHeight As Long
    Dim LabelWidth As Long
    Dim ComboWidth As Long
    Dim VerticalSpace As Long
    Dim HorizentalSpace As Long

    BottomMargin = 8450
    LeftMargin = 120
    TopMargin = 2160
    ComboHeight = 360
    LabelWidth = 2055
    ComboWidth = 3000
    VerticalSpace = 120
    HorizentalSpace = 400


    Dim IxItemHCount As Long
    Dim IxItemVCount As Long

    IxItemHCount = 1
    IxItemVCount = 0


    With rsTemIxItem
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem where Deleted = false AND (IsValue = true  or IsCalc = True or IsFlag = True) AND IxID = " & Val(cmbIx.BoundText) & " Order by Format(X1, '0.0'), Y1"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                IxItemVCount = IxItemVCount + 1

                If TopMargin + ((IxItemVCount - 1) * (ComboHeight + VerticalSpace)) > BottomMargin Then
                    IxItemHCount = IxItemHCount + 1
                    IxItemVCount = 1
                End If

                Set MyCombo = Me.Controls.Add("VB.ComboBox", "Ixcmb" & !IxItemID, Me)
                MyCombo.Visible = True
                MyCombo.Top = TopMargin + ((IxItemVCount - 1) * (ComboHeight + VerticalSpace))
                MyCombo.Left = LeftMargin + LabelWidth + ((IxItemHCount - 1) * (LeftMargin + LabelWidth + ComboWidth + HorizentalSpace))
                MyCombo.Width = ComboWidth
'                MyCombo.Height = ComboHeight
                Call FillNewCombos(MyCombo, !IxItemID)

                Set MyLabel = Me.Controls.Add("VB.label", "lbl" & !IxItemID, Me)
                MyLabel.Visible = True
                MyLabel.Top = TopMargin + ((IxItemVCount - 1) * (ComboHeight + VerticalSpace))
                MyLabel.Left = LeftMargin + ((IxItemHCount - 1) * (LeftMargin + LabelWidth + ComboWidth + HorizentalSpace))
                MyLabel.Width = LabelWidth
                MyLabel.Height = ComboHeight
                MyLabel.BackStyle = 0
                MyLabel.Caption = !IxItem
                .MoveNext
            Wend
        End If
        .Close

    End With
End Sub


Private Sub FillNewCombos(NewCombo As ComboBox, IxItemID As Long)
    Dim rsTem As New ADODB.Recordset
    Dim temIxValueCategoryID As Long
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem where IxItemID = " & IxItemID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temIxValueCategoryID = !IxValueCategoryID
        Else
            temIxValueCategoryID = 0
        End If
        If .State = 1 Then .Close
        temSql = "SELECT tblIxValue.* FROM (tblIxValuesForCategories LEFT JOIN tblIxValue ON tblIxValuesForCategories.ValueID = tblIxValue.IxValueID) RIGHT JOIN tblCategory ON tblIxValuesForCategories.CategoryID = tblCategory.CategoryID WHERE (((tblCategory.CategoryID)=" & temIxValueCategoryID & ")) ORDER BY tblIxValue.IxValue "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            NewCombo.AddItem Format(!IxValue, "")
            .MoveNext
        Wend
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    
    
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
End Sub

