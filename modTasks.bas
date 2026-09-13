Attribute VB_Name = "modTasks"
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
    Dim rsIxList As New ADODB.Recordset
    Dim i As Integer


Public Sub SaveCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                SaveSetting App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i)
            Next
        End If
    Next
    SaveSetting App.EXEName, MyForm.Name, "Top", MyForm.Top
    SaveSetting App.EXEName, MyForm.Name, "Left", MyForm.Left
End Sub

Public Sub GetCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                MyCtrl.ColWidth(i) = GetSetting(App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i))
                MyCtrl.AllowUserResizing = flexResizeColumns
            Next
        End If
    Next
    MyForm.Top = GetSetting(App.EXEName, MyForm.Name, "Top", MyForm.Top)
    MyForm.Left = GetSetting(App.EXEName, MyForm.Name, "Left", MyForm.Left)
End Sub

Public Sub FillPrinters(PrinterCombo As ComboBox)
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        PrinterCombo.AddItem MyPrinter.DeviceName
    Next
End Sub

Public Sub FillPrinterPapers(PrinterCombo As ComboBox, PaperCombo As ComboBox)
    PaperCombo.Clear
    If PrinterCombo.Text = Empty Then Exit Sub
    CsetPrinter.SetPrinterAsDefault (PrinterCombo.Text)
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
                PaperCombo.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Public Sub ColourGrid(GridToColour As MSFlexGrid)
    Dim col As Integer
    Dim row As Integer
    With GridToColour
        .Visible = False
        For row = 0 To .Rows - 1
            For col = 0 To .Cols - 1
                .col = col
                .row = row
                If .row Mod 2 = 0 Then
                    .CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    .CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
            Next
        Next
        .Visible = True
    End With
End Sub
