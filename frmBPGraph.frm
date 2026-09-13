VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBPGraph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blood Pressure Graph"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
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
   ScaleHeight     =   3945
   ScaleWidth      =   5970
   Begin VB.Frame Frame11 
      Caption         =   "Ligends"
      Height          =   735
      Left            =   3480
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
      Begin VB.OptionButton optNoLegend 
         Caption         =   "Hide"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optDisplayLegend 
         Caption         =   "Display"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Values"
      Height          =   735
      Left            =   3480
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
      Begin VB.OptionButton optDisplayValues 
         Caption         =   "Display"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optDoNotDisplayValues 
         Caption         =   "Hide"
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot by"
      Height          =   735
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
      Begin VB.OptionButton optYAxis 
         Caption         =   "Rows"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optXAxis 
         Caption         =   "Columns"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Title"
      Height          =   735
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   2295
      Begin VB.OptionButton optNoTitle 
         Caption         =   "Hide"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optDisplayTitle 
         Caption         =   "Display"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbChartType 
      Height          =   360
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   15859715
      CurrentDate     =   39910
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   15859715
      CurrentDate     =   39910
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
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
   Begin btButtonEx.ButtonEx btnRefresh 
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&reate"
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
   Begin VB.Label Label3 
      Caption         =   "Chart type"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBPGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim temPath As String
    Dim FSys As New Scripting.FileSystemObject
    Dim temTopic As String
    Dim temSubTopic As String
    
    Dim rsTem As New ADODB.Recordset
    Dim temCol As Long
    Dim temSql As String
    Dim TemString As String
    
Private Sub btnRefresh_Click()
    Call CreateGraph
    Call SetGraph
    Call DisplayGraph
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub cmbChartType_Change()
    'Call SetGraph
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    
    
    On Error Resume Next
    Dim MyControl  As Control
    For Each MyControl In Controls
        If Left(MyControl.Name, 3) = "opt" Then
            MyControl.Value = GetSetting(App.EXEName, Me.Name, MyControl.Name, "False")
        End If
        If Left(MyControl.Name, 3) = "cmb" Then
            MyControl.Text = GetSetting(App.EXEName, Me.Name, MyControl.Name, "")
        End If
    Next
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    
    
    Dim MyControl  As Control
    For Each MyControl In Controls
        If Left(MyControl.Name, 3) = "opt" Then
            SaveSetting App.EXEName, Me.Name, MyControl.Name, MyControl.Value
        End If
        If Left(MyControl.Name, 3) = "cmb" Then
            SaveSetting App.EXEName, Me.Name, MyControl.Name, MyControl.Text
        End If
    Next
End Sub

Private Sub CreateGraph()
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set mychart = AppExcel.Charts.Add
    Set myworksheet = AppExcel.Worksheets(1)
    Set mychart = AppExcel.Charts(1)
    myworksheet.Activate
    
    temTopic = "BP chart of " & CurrentPatient.NameWithTitle & " From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
    myworksheet.Cells(1, 1) = temTopic
    myworksheet.Cells(2, 1) = "Date"
    myworksheet.Cells(3, 1) = "SBP"
    myworksheet.Cells(4, 1) = "DBP"
    
    temCol = 2
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select min(VisitDate) as MinOfVisitDate from tblVisit where PatientID = " & CurrentPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MinOfVisitDate) = False Then
                dtpFrom.Value = !MinOfVisitDate
            Else
                dtpFrom.Value = Date
            End If
        End If
        
        temSql = "Select * from tblVisit where PatientID = " & CurrentPatient.ID & " And VIsitDate between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# AND #" & Format(dtpTo.Value, "dd MMMM yyyy") & "# order by VisitDate"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If !SBP <> 0 And !DBP <> 0 Then
                myworksheet.Cells(2, temCol) = Format(!VisitDate, LongDateFormat)
                myworksheet.Cells(3, temCol) = !SBP
                myworksheet.Cells(4, temCol) = !DBP
                temCol = temCol + 1
            End If
            .MoveNext
        Wend
        .Close
    End With
    
    mychart.SetSourceData myworksheet.Range("A2:" & GetColumnName(temCol - 1) & "4")
    mychart.Activate
    mychart.HasTitle = True
    mychart.ChartTitle.Caption = temTopic
    
    Call SetGraph
    

    TemString = ""
    While FSys.FileExists(temPath & "\" & temTopic & " " & temSubTopic & TemString & ".xls") = True
        TemString = TemString & "1"
    Wend
    
    DoEvents
    
    mychart.Activate
    If Right(temPath, 1) <> "\" And Right(temPath, 1) <> "/" Then
        myworkbook.SaveAs (temPath & "\" & temTopic & " " & temSubTopic & TemString & ".xls")
        GraphPathTx = (temPath & "\" & temTopic & " " & temSubTopic & TemString & ".xls")
    Else
        myworkbook.SaveAs (temPath & temTopic & " " & temSubTopic & TemString & ".xls")
        GraphPathTx = (temPath & temTopic & " " & temSubTopic & TemString & ".xls")
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
    Call SetColours
    Call FillChartTypes
    
    dtpFrom.Value = Date
    dtpTo.Value = Date
        
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select min(VisitDate) as MinOfVisitDate from tblVisit where PatientID = " & CurrentPatient.ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MinOfVisitDate) = False Then
                dtpFrom.Value = !MinOfVisitDate
            Else
                dtpFrom.Value = Date
            End If
        End If
    End With
    
    temPath = FSys.GetParentFolderName(Database)
   
    Call GetSettings
    
End Sub

Private Sub DisplayGraph()
    Set mychart = Nothing
    Set myworksheet = Nothing
    Set myworkbook = Nothing
    AppExcel.ActiveWorkbook.Save
    AppExcel.ActiveWorkbook.Close
    AppExcel.Quit
    Set AppExcel = Nothing
    Dim MyForm As Form
    Set MyForm = New frmGraph
    MyForm.Show
    MyForm.ZOrder 0
    
End Sub

Private Sub FillChartTypes()
    With cmbChartType
        .AddItem "3D Area"
        .AddItem "3D Area Stacked"
        .AddItem "3D Area Stacked 100"
        .AddItem "xl3DBar"
        .AddItem "3D Bar Clustered"
        .AddItem "3DBarStacked"
        .AddItem "3DBarStacked100"
        .AddItem "3DColumn"
        .AddItem "3DColumnClustered"
        .AddItem "3DColumnStacked"
        .AddItem "3DColumnStacked100"
        .AddItem "3DLine"
        .AddItem "3DPie"
        .AddItem "3DPieExploded"
        .AddItem "Area"
        .AddItem "AreaStacked"
        .AddItem "AreaStacked100"
        .AddItem "BarClustered"
        .AddItem "BarOfPie"
        .AddItem "BarStacked"
        .AddItem "BarStacked"
        .AddItem "BarStacked100"
        .AddItem "Bubble"
        .AddItem "Bubble3DEffect"
        .AddItem "Column"
        .AddItem "ColumnClustered"
        .AddItem "ColumnStacked"
        .AddItem "ColumnStacked100"
        .AddItem "ConeBarClustered"
        .AddItem "ConeBarStacked"
        .AddItem "ConeBarStacked100"
        .AddItem "ConeCol"
        .AddItem "ConeColClustered"
        .AddItem "ConeColStacked"
        .AddItem "ConeColStacked100"
        .AddItem "Cylinder"
        .AddItem "CylinderBarClustered"
        .AddItem "CylinderBarStacked"
        .AddItem "CylinderBarStacked100"
        .AddItem "CylinderCol"
        .AddItem "CylinderColClustered"
        .AddItem "CylinderColStacked"
        .AddItem "CylinderColStacked100"
        .AddItem "Doughnut"
        .AddItem "DoughnutExploded"
        .AddItem "Line"
        .AddItem "LineMarkers"
        .AddItem "LineMarkersStacked"
        .AddItem "LineMarkersStacked100"
        .AddItem "LineStacked"
        .AddItem "LineStacked100"
        .AddItem "Pie"
        .AddItem "PieExploded"
        .AddItem "PieOfPie"
        .AddItem "PyramidBarClustered"
        .AddItem "PyramidBarStacked"
        .AddItem "PyramidBarStacked100"
        .AddItem "PyramidCol"
        .AddItem "PyramidColClustered"
        .AddItem "PyramidColStacked"
        .AddItem "PyramidColStacked100"
        .AddItem "Radar"
        .AddItem "RadarFilled"
        .AddItem "RadarMarkers"
        .AddItem "Surface"
        .AddItem "SurfaceTopView"
        .AddItem "SurfaceTopViewWireframe"
        .AddItem "SurfaceWireframe"
        .AddItem "XYScatter"
        .AddItem "XYScatterLines"
        .AddItem "XYScatterLinesNoMarkers"
        .AddItem "XYScatterSmooth"
        .AddItem "XYScatterSmoothNoMarkers"
        
        .ItemData(0) = xl3DArea
        .ItemData(1) = xl3DAreaStacked
        .ItemData(2) = xl3DAreaStacked
        .ItemData(3) = xl3DBarClustered
        .ItemData(4) = xl3DBarClustered
        .ItemData(5) = xl3DBarStacked
        .ItemData(6) = xl3DBarStacked100
        .ItemData(7) = xl3DColumn
        .ItemData(8) = xl3DColumnClustered
        .ItemData(9) = xl3DColumnStacked
        .ItemData(10) = xl3DColumnStacked100
        .ItemData(11) = xl3DLine
        .ItemData(12) = xl3DPie
        .ItemData(13) = xl3DPieExploded
        .ItemData(14) = xlArea
        .ItemData(15) = xlAreaStacked
        .ItemData(16) = xlAreaStacked100
        .ItemData(17) = xlBarClustered
        .ItemData(18) = xlBarOfPie
        .ItemData(19) = xlBarStacked
        .ItemData(20) = xlBarStacked
        .ItemData(21) = xlBarStacked100
        .ItemData(22) = xlBubble
        .ItemData(23) = xlBubble3DEffect
        .ItemData(24) = xlColumnClustered
        .ItemData(25) = xlColumnStacked
        .ItemData(26) = xlColumnStacked
        .ItemData(27) = xlColumnStacked100
        .ItemData(28) = xlConeBarClustered
        .ItemData(29) = xlConeBarStacked
        .ItemData(30) = xlConeBarStacked100
        .ItemData(31) = xlConeCol
        .ItemData(32) = xlConeColClustered
        .ItemData(33) = xlConeColStacked
        .ItemData(34) = xlConeColStacked100
        .ItemData(35) = xlCylinderBarClustered
        .ItemData(36) = xlCylinderBarClustered
        .ItemData(37) = xlCylinderBarStacked
        .ItemData(38) = xlCylinderBarStacked100
        .ItemData(39) = xlCylinderCol
        .ItemData(40) = xlCylinderColClustered
        .ItemData(40) = xlCylinderColStacked
        .ItemData(42) = xlCylinderColStacked100
        .ItemData(43) = xlDoughnut
        .ItemData(44) = xlDoughnutExploded
        .ItemData(45) = xlLine
        .ItemData(46) = xlLineMarkers
        .ItemData(47) = xlLineMarkersStacked
        .ItemData(48) = xlLineMarkersStacked100
        .ItemData(49) = xlLineStacked
        .ItemData(50) = xlLineStacked100
        .ItemData(51) = xlPie
        .ItemData(52) = xlPieExploded
        .ItemData(53) = xlPieOfPie
        .ItemData(54) = xlPyramidBarClustered
        .ItemData(55) = xlPyramidBarStacked
        .ItemData(56) = xlPyramidBarStacked100
        .ItemData(57) = xlPyramidCol
        .ItemData(58) = xlPyramidColClustered
        .ItemData(59) = xlPyramidColStacked
        .ItemData(60) = xlPyramidColStacked100
        .ItemData(61) = xlRadar
        .ItemData(62) = xlRadarFilled
        .ItemData(63) = xlRadarMarkers
        .ItemData(64) = xlSurface
        .ItemData(65) = xlSurfaceTopView
        .ItemData(66) = xlSurfaceTopViewWireframe
        .ItemData(67) = xlSurfaceWireframe
        .ItemData(68) = xlXYScatter
        .ItemData(69) = xlXYScatterLines
        .ItemData(70) = xlXYScatterLinesNoMarkers
        .ItemData(71) = xlXYScatterSmooth
        .ItemData(72) = xlXYScatterSmoothNoMarkers
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub SetGraph()
    If optDisplayTitle.Value = True Then
        mychart.HasTitle = True
    Else
        mychart.HasTitle = False
    End If
    If optDisplayLegend.Value = True Then
        mychart.HasLegend = True
    Else
        mychart.HasLegend = False
    End If
    If optDisplayValues.Value = True Then
        mychart.ApplyDataLabels xlDataLabelsShowValue
    Else
        mychart.ApplyDataLabels xlDataLabelsShowNone
    End If
    If optXAxis.Value = True Then
        mychart.PlotBy = xlColumns
    ElseIf optYAxis.Value = True Then
        mychart.PlotBy = xlRows
    End If
    On Error Resume Next
    If cmbChartType.ListIndex > 0 Then
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
End Sub
