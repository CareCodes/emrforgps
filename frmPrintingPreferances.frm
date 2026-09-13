VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPrintingPreferances 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printing Preferances"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
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
   ScaleHeight     =   3765
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7560
      TabIndex        =   24
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   4895
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Prescreptions"
      TabPicture(0)   =   "frmPrintingPreferances.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label20"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPScaleWidth"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPScaleHeight"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbPPaper"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbPPrinter"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkPOS"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Documents"
      TabPicture(1)   =   "frmPrintingPreferances.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDScaleWidth"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtDScaleHeight"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmbDPaper"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbDPrinter"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label3"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Reports"
      TabPicture(2)   =   "frmPrintingPreferances.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRScaleWidth"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtRScaleHeight"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmbRPaper"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmbRPrinter"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label14"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label13"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label12"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label11"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label10"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label9"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.CheckBox chkPOS 
         Caption         =   "POS Printer"
         Height          =   375
         Left            =   4800
         TabIndex        =   32
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtRScaleWidth 
         Height          =   360
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtRScaleHeight 
         Height          =   360
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbRPaper 
         Height          =   360
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox cmbRPrinter 
         Height          =   360
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtDScaleWidth 
         Height          =   360
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtDScaleHeight 
         Height          =   360
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cmbDPaper 
         Height          =   360
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox cmbDPrinter 
         Height          =   360
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox cmbPPrinter 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox cmbPPaper 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtPScaleHeight 
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtPScaleWidth 
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   375
         Left            =   -74760
         TabIndex        =   22
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         Height          =   255
         Left            =   -72000
         TabIndex        =   31
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         Height          =   255
         Left            =   -72000
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper"
         Height          =   375
         Left            =   -74760
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   375
         Left            =   -74760
         TabIndex        =   14
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         Height          =   255
         Left            =   -72000
         TabIndex        =   29
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         Height          =   255
         Left            =   -72000
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         Height          =   255
         Left            =   -74760
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper"
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Paper"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPrintingPreferances"
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
    Dim temSQL As String
    Dim rsIxList As New ADODB.Recordset
    Dim i As Integer


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbDPaper_Change()
    CsetPrinter.SetPrinterAsDefault (cmbDPrinter.Text)
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
                If cmbDPaper.Text = PtrCtoVbString(.pName) Then
                    txtDScaleHeight.Text = Format(.Size.cy / 1000 * mmtoinches, "0.00")
                    txtDScaleWidth.Text = Format(.Size.cx / 1000 * mmtoinches, "0.00")
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbDPaper_Click()
    CsetPrinter.SetPrinterAsDefault (cmbDPrinter.Text)
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
                If cmbDPaper.Text = PtrCtoVbString(.pName) Then
                    txtDScaleHeight.Text = Format(.Size.cy / 1000 * mmtoinches, "0.00")
                    txtDScaleWidth.Text = Format(.Size.cx / 1000 * mmtoinches, "0.00")
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub









Private Sub cmbRPaper_Change()
    CsetPrinter.SetPrinterAsDefault (cmbRPrinter.Text)
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
                If cmbRPaper.Text = PtrCtoVbString(.pName) Then
                    txtRScaleHeight.Text = Format(.Size.cy / 1000 * mmtoinches, "0.00")
                    txtRScaleWidth.Text = Format(.Size.cx / 1000 * mmtoinches, "0.00")
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbRPaper_Click()
    CsetPrinter.SetPrinterAsDefault (cmbRPrinter.Text)
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
                If cmbRPaper.Text = PtrCtoVbString(.pName) Then
                    txtRScaleHeight.Text = Format(.Size.cy / 1000 * mmtoinches, "0.00")
                    txtRScaleWidth.Text = Format(.Size.cx / 1000 * mmtoinches, "0.00")
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbPPaper_Change()
    CsetPrinter.SetPrinterAsDefault (cmbPPrinter.Text)
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
                If cmbPPaper.Text = PtrCtoVbString(.pName) Then
                    txtPScaleHeight.Text = Format(.Size.cy / 1000 * mmtoinches, "0.00")
                    txtPScaleWidth.Text = Format(.Size.cx / 1000 * mmtoinches, "0.00")
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbPPaper_Click()
    CsetPrinter.SetPrinterAsDefault (cmbPPrinter.Text)
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
                If cmbPPaper.Text = PtrCtoVbString(.pName) Then
                    txtPScaleHeight.Text = Format(.Size.cy / 1000 * mmtoinches, "0.00")
                    txtPScaleWidth.Text = Format(.Size.cx / 1000 * mmtoinches, "0.00")
                End If
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbPPrinter_Change()
    cmbPPrinter_Click
End Sub

Private Sub cmbRPrinter_Change()
    cmbRPrinter_Click
End Sub

Private Sub cmbDPrinter_Change()
    cmbDPrinter_Click
End Sub


Private Sub cmbPPrinter_Click()
    cmbPPaper.Clear
    CsetPrinter.SetPrinterAsDefault (cmbPPrinter.Text)
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
                cmbPPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbRPrinter_Click()
    cmbRPaper.Clear
    CsetPrinter.SetPrinterAsDefault (cmbRPrinter.Text)
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
                cmbRPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbDPrinter_Click()
    cmbDPaper.Clear
    CsetPrinter.SetPrinterAsDefault (cmbDPrinter.Text)
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
                cmbDPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPPrinter.AddItem MyPrinter.DeviceName
        cmbRPrinter.AddItem MyPrinter.DeviceName
        cmbDPrinter.AddItem MyPrinter.DeviceName
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
End Sub

Private Sub GetSettings()
    On Error Resume Next
    GetCommonSettings Me
    
    
    txtPScaleWidth.Text = Val(GetSetting(App.EXEName, App.EXEName, "PrescreptionPaperWidth", 4.1))
    txtPScaleHeight.Text = Val(GetSetting(App.EXEName, App.EXEName, "PrescreptionPaperHeight", 6.3))
    cmbPPrinter.Text = GetSetting(App.EXEName, App.EXEName, "PrescreptionPrinterName", "")
    cmbPPrinter_Click
    cmbPPaper.Text = GetSetting(App.EXEName, App.EXEName, "PrescreptionPaperName", "")
    
    txtRScaleWidth.Text = Val(GetSetting(App.EXEName, App.EXEName, "ReportPaperWidth", 4.1))
    txtRScaleHeight.Text = Val(GetSetting(App.EXEName, App.EXEName, "ReportPaperHeight", 6.3))
    cmbRPrinter.Text = GetSetting(App.EXEName, App.EXEName, "ReportPrinterName", "")
    cmbRPrinter_Click
    cmbRPaper.Text = GetSetting(App.EXEName, App.EXEName, "ReportPaperName", "")
    
    txtDScaleWidth.Text = Val(GetSetting(App.EXEName, App.EXEName, "DocumentPaperWidth", 4.1))
    txtDScaleHeight.Text = Val(GetSetting(App.EXEName, App.EXEName, "DocumentPaperHeight", 6.3))
    cmbDPrinter.Text = GetSetting(App.EXEName, App.EXEName, "DocumentPrinterName", "")
    cmbDPrinter_Click
    cmbDPaper.Text = GetSetting(App.EXEName, App.EXEName, "DocumentPaperName", "")
    
    If GetSetting(App.EXEName, App.EXEName, "printPOS", 0) = "1" Then
        chkPOS.Value = 1
    Else
        chkPOS.Value = 0
    End If
    
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, App.EXEName, "PrescreptionPrinterName", cmbPPrinter.Text
    SaveSetting App.EXEName, App.EXEName, "PrescreptionPaperName", cmbPPaper.Text
    SaveSetting App.EXEName, App.EXEName, "PrescreptionPaperHeight", Val(txtPScaleHeight.Text)
    SaveSetting App.EXEName, App.EXEName, "PrescreptionPaperWidth", Val(txtPScaleWidth.Text)

    SaveSetting App.EXEName, App.EXEName, "ReportPrinterName", cmbRPrinter.Text
    SaveSetting App.EXEName, App.EXEName, "ReportPaperName", cmbRPaper.Text
    SaveSetting App.EXEName, App.EXEName, "ReportPaperHeight", Val(txtRScaleHeight.Text)
    SaveSetting App.EXEName, App.EXEName, "ReportPaperWidth", Val(txtRScaleWidth.Text)

    SaveSetting App.EXEName, App.EXEName, "DocumentPrinterName", cmbDPrinter.Text
    SaveSetting App.EXEName, App.EXEName, "DocumentPaperName", cmbDPaper.Text
    SaveSetting App.EXEName, App.EXEName, "DocumentPaperHeight", Val(txtDScaleHeight.Text)
    SaveSetting App.EXEName, App.EXEName, "DocumentPaperWidth", Val(txtDScaleWidth.Text)
    
    SaveSetting App.EXEName, App.EXEName, "printPOS", chkPOS.Value
    SaveCommonSettings Me
    

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

