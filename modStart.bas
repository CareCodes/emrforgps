Attribute VB_Name = "modStart"
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim constr As String
    Dim temSQL As String

Public Sub Main()
    On Error Resume Next
    If App.PrevInstance = True Then
        MsgBox "Program is already running"
        End
    End If
    
    Dim TemResponce As Byte
    Dim temPath As String
    
    WillExpire = True
    DemoCopy = False
    ExpiaryCount = 4800
    ExpiaryDays = 965
    ExpiaryDate = #3/31/2016#
    HiddenDemo = True
    
    TodayStartedDate = Date
    
    UsedDays = (Val(GetSetting(App.EXEName, "modStart", "UserID", 13)))
    
    If UsedDays = 13 Then
        UsedDays = 1
        SaveSetting App.EXEName, "modStart", "UserID", UsedDays * 97
    Else
        If UsedDays Mod 97 <> 0 Then
            MsgBox "You can't convert a Demo Copy to a Legal Copy. Please contact Fintec. Sorry"
            End
        End If
        UsedDays = UsedDays / 97
    End If
    
    App.HelpFile = App.Path & "\EMA3Help.chm"
    
    
    If Date <> CDate((GetSetting(App.EXEName, "modStart", "UserDOD", Date))) Then
        SaveSetting App.EXEName, "modStart", "UserID", ((UsedDays * 97) + 97)
        SaveSetting App.EXEName, "modStart", "UserDOD", Date
    End If
    
    If WillExpire = True Then
        DemoCount = Val(GetSetting(App.EXEName, "modStart", "SystemID", 1))
        If DemoCount > ExpiaryCount Then
            TemResponce = MsgBox("The Program has expiared. Please contact Fintec for Assistant", vbCritical, "Expired")
            End
        Else
            SaveSetting App.EXEName, "modStart", "SystemID", DemoCount + 1
        End If
        If ExpiaryDate < Date Then
            TemResponce = MsgBox("The Program has expiared. Please contact Fintec for Assistant", vbCritical, "Expired")
            End
        Else
            If HiddenDemo = False Then
                TemResponce = MsgBox("This is a Demo Copy." & vbNewLine & "User Name :" & vbTab & "b" & vbTab & vbNewLine & "Password :" & vbTab & "b" & vbTab, vbInformation, "Demo")
            ElseIf UsedDays = 1 Then
                'TemResponce = MsgBox("Until users are configured in Menu > Edit > Staff, use the following details in loggin." & vbNewLine & "User Name :" & vbTab & "b" & vbTab & vbNewLine & "Password :" & vbTab & "b" & vbTab, vbInformation, "User Loggin")
            End If
        End If
        If ExpiaryDays < UsedDays Then
            TemResponce = MsgBox("The Program has expiared. Please contact Fintec for Assistant", vbCritical, "Expired")
            End
        Else
        
        End If
    Else
        TodayStartedDate = Date
    End If
    
    
    Call LoadColourPreferances
    Call LoadPreferances
    
    If Database = "" Then
        Database = App.Path & "\EMA3.MDB"
    End If
    
    If FSys.FileExists(Database) = False Then
        TemResponce = MsgBox("The path of the database you have selected does not exist. Please select the database", vbCritical, "Wrong database path")
        frmInitialPreferances.Show 1
    End If
    
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Database & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=REDACTED"
    cnnStores.Open constr
    DataEnvironment1.deCnnStores.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Database & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=REDACTED"    ' "data source=" & constr  'GetSetting(App.EXEName, "Options", "Database", App.Path & "\hospital.mdb")
    
    Call GetInstitution
    
'    PCActivated = CheckPCActivated
'    DatabaseActivated = CheckDatabaseActivated
    
'    If PCActivated = True And DatabaseActivated = True Then
        frmLogin.Show
'    ElseIf PCActivated = False And DatabaseActivated = False Then
'        frmDatabaseActivation.Show 1
'        frmPCActivation.Show 1
'    ElseIf PCActivated = True And DatabaseActivated = False Then
'        MsgBox "Your computer is activated, but NOT the database. Please contact your reseller"
'        End
'    ElseIf PCActivated = False And DatabaseActivated = True Then
'        frmPCActivation.Show
'    End If
    
End Sub


Private Sub LoadPreferances()
On Error Resume Next
    Database = GetSetting(App.EXEName, "Options", "Database", "")
    ImportDatabase = "http://www.lakmedi.com/EMA3.mdb"

    LongDateFormat = "yyyy mmmm dd" ' GetSetting(App.EXEName, "Options", "LongDateFormat", "dddd, dd MMMM yyyy")
    ShortDateFormat = "dd mm yy"  ' GetSetting(App.EXEName, "Options", "ShortDateFormat", "dd MM yy")
    
    PrescreptionPaperWidth = Val(GetSetting(App.EXEName, App.EXEName, "PrescreptionPaperWidth", 4.1))
    PrescreptionPaperHeight = Val(GetSetting(App.EXEName, App.EXEName, "PrescreptionPaperHeight", 6.3))
    PrescreptionPrinterName = GetSetting(App.EXEName, App.EXEName, "PrescreptionPrinterName", "")
    PrescreptionPaperName = GetSetting(App.EXEName, App.EXEName, "PrescreptionPaperName", "")
    
    ReportPaperWidth = Val(GetSetting(App.EXEName, App.EXEName, "ReportPaperWidth", 4.1))
    ReportPaperHeight = Val(GetSetting(App.EXEName, App.EXEName, "ReportPaperHeight", 6.3))
    ReportPrinterName = GetSetting(App.EXEName, App.EXEName, "ReportPrinterName", "")
    ReportPaperName = GetSetting(App.EXEName, App.EXEName, "ReportPaperName", "")
    
    DocumentPaperWidth = Val(GetSetting(App.EXEName, App.EXEName, "DocumentPaperWidth", 4.1))
    DocumentPaperHeight = Val(GetSetting(App.EXEName, App.EXEName, "DocumentPaperHeight", 6.3))
    DocumentPrinterName = GetSetting(App.EXEName, App.EXEName, "DocumentPrinterName", "")
    DocumentPaperName = GetSetting(App.EXEName, App.EXEName, "DocumentPaperName", "")
    
    PrintingOnBlankPaper = GetSetting(App.EXEName, "Options", "PrintingOnBlankPaper", True)
    PrintingOnPrintedPaper = GetSetting(App.EXEName, "Options", "PrintingOnPrintedPaper", False)
    
    printPOS = GetSetting(App.EXEName, App.EXEName, "printPOS", False)
    
    DemoCount = GetSetting(App.EXEName, "Options", "DemoCount", 1)
    
    
    InstitutionName = "Boraluketiya Medical Centre"
    InstitutionAddress = "Malana, Kamburupitiya" & vbNewLine & "071 5812399, 041 2292022"

    
    
    UsualConsultationFee = Val(GetSetting(App.EXEName, "Options", "UsualConsultationFee", 150))
    UsualConsultationCost = Val(GetSetting(App.EXEName, "Options", "UsualConsultationCost", 150))
    
    AddMedicineChargesToFee = CBool(GetSetting(App.EXEName, "Options", "AddMedicineChargesToFee", False))
    AddIndoorMedicineChargesToFee = CBool(GetSetting(App.EXEName, "Options", "AddIndoorMedicineChargesToFee", False))
    
    AutoIssueIndoor = CBool(GetSetting(App.EXEName, "Options", "AutoIssueIndoor", False))
    
    AutoSecession = CBool(GetSetting(App.EXEName, "Options", "AutoSecession", True))
    SingleVisit = CBool(GetSetting(App.EXEName, "Options", "SingleVisit", True))
    
    If CBool(GetSetting(App.EXEName, "Options", "Energy", True)) = True Then
        DefaultColourScheme = Energy
    ElseIf CBool(GetSetting(App.EXEName, "Options", "Sunny", False)) = True Then
        DefaultColourScheme = Sunny
    ElseIf CBool(GetSetting(App.EXEName, "Options", "Aqua", False)) = True Then
        DefaultColourScheme = Aqua
    End If
    
    mmtoinches = 0.0393700787
    
    MDIImageFile = GetSetting(App.EXEName, "Options", "MDIImageFile", "")
    
End Sub

Private Sub GetInstitution()
On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblInstitution where IsOwner = True AND Deleted = False"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            InstitutionName = !Institution
            InstitutionAddress = !Address
            InstitutionTelephone = !Telephone
        End If
        .Close
    End With
End Sub

Private Sub LoadColourPreferances()
On Error Resume Next
    Energy.LabelForeColour = RGB(154, 14, 20)
    Energy.LabelBackColour = RGB(255, 220, 168)
    Energy.ButtonBackColour = RGB(255, 129, 81)
    Energy.ButtonForeColour = RGB(154, 14, 20)
    Energy.ButtonBorderColour = RGB(154, 14, 20)
    Energy.GridDarkBackColour = RGB(255, 220, 168)
    Energy.GridLightBackColour = RGB(255, 245, 218)
    Energy.ComboForeColour = RGB(154, 14, 20)
    Energy.ComboBackColour = RGB(255, 245, 218)
    Energy.TextForeColour = RGB(154, 14, 20)
    Energy.TextBackColour = RGB(255, 245, 218)
    

    Sunny.LabelForeColour = RGB(154, 97, 14)
    Sunny.LabelBackColour = RGB(255, 236, 168)
    Sunny.ButtonBackColour = RGB(255, 200, 0)
    Sunny.ButtonForeColour = RGB(154, 97, 14)
    Sunny.ButtonBorderColour = RGB(154, 97, 14)
    Sunny.GridDarkBackColour = RGB(255, 236, 168)
    Sunny.GridLightBackColour = RGB(255, 255, 232)
    Sunny.ComboForeColour = RGB(154, 97, 14)
    Sunny.ComboBackColour = RGB(255, 255, 232)
    Sunny.TextForeColour = RGB(154, 97, 14)
    Sunny.TextBackColour = RGB(255, 255, 232)
    
    Aqua.LabelForeColour = RGB(34, 134, 84)
    Aqua.LabelBackColour = RGB(168, 212, 255)
    Aqua.ButtonBackColour = RGB(100, 225, 225)
    Aqua.ButtonForeColour = RGB(34, 134, 84)
    Aqua.ButtonBorderColour = RGB(34, 134, 84)
    Aqua.GridDarkBackColour = RGB(168, 212, 255)
    Aqua.GridLightBackColour = RGB(232, 255, 255)
    Aqua.ComboForeColour = RGB(34, 134, 84)
    Aqua.ComboBackColour = RGB(232, 255, 255)
    Aqua.TextForeColour = RGB(34, 134, 84)
    Aqua.TextBackColour = RGB(232, 255, 255)


End Sub


Public Function CheckDatabaseActivated() As Boolean
On Error Resume Next
    CheckDatabaseActivated = True
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDoctorNumber where DoctorNUmberID = " & 20150
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If !DoctorNo = 6110 And !HospitalNo = 3988 Then
                
            Else
                CheckDatabaseActivated = False
            End If
        Else
            CheckDatabaseActivated = False
        End If
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDoctorNumber where DoctorNUmberID = " & 20148
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If !DoctorNo = 2406 And !HospitalNo = 6739 Then
                
            Else
                CheckDatabaseActivated = False
            End If
        Else
            CheckDatabaseActivated = False
        End If
        
        .Close
    End With
End Function

Public Function CheckPCActivated() As Boolean
On Error Resume Next
    Dim MyStr As String
    Dim MyLong As Long
    Dim MyDbl As Double
    Dim MyBool As Boolean
    MyStr = GetSetting(App.EXEName, "Preferances", "DatabaseLocation", "C:\Program Files\Lakmedi\EMA3.MDB")
    MyLong = GetSetting(App.EXEName, "frmPreferances", "Ht", "1975914")
    MyDbl = GetSetting(App.EXEName, "frmCharges", "BackColour", "17.12")
    MyBool = GetSetting(App.EXEName, "frmDeposit", "CanDeposit", False)
    If MyStr <> "C:\\Program Files\\Lakmedi\\EMA3.mdb" Then
        CheckPCActivated = False
        Exit Function
    End If
    If MyLong <> 19751217 Then
        CheckPCActivated = False
        Exit Function
    End If
    If MyDbl <> 14.9 Then
        CheckPCActivated = False
        Exit Function
    End If
    If MyBool = False Then
        CheckPCActivated = False
        Exit Function
    End If
    CheckPCActivated = True
End Function

Public Sub ActivateDatabase()
On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDoctorNumber where DoctorNUmberID = " & 20150
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DoctorNo = 6110
            !HospitalNo = 3988
            .Update
        End If
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDoctorNumber where DoctorNUmberID = " & 20148
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DoctorNo = 2406
            !HospitalNo = 6739
            .Update
        End If
        .Close
    End With
End Sub

Public Sub DeactivateDatabase()
On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDoctorNumber where DoctorNUmberID = " & 20148
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !HospitalNo = 2406
            !DoctorNo = 6739
            .Update
        End If
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDoctorNumber where DoctorNUmberID = " & 20150
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !HospitalNo = 6110
            !DoctorNo = 3988
            .Update
        End If
        .Close
    End With
End Sub

Public Sub DeActivatePC()
On Error Resume Next
    SaveSetting App.EXEName, "Preferances", "DatabaseLocation", "C:\Program Files\Lakmedi\EMA3.MDB"
    SaveSetting App.EXEName, "frmPreferances", "Ht", "1975914"
    SaveSetting App.EXEName, "frmCharges", "BackColour", "17.12"
    SaveSetting App.EXEName, "frmDeposit", "CanDeposit", False
End Sub

Public Sub ActivatePC()
On Error Resume Next
    SaveSetting App.EXEName, "Preferances", "DatabaseLocation", "C:\\Program Files\\Lakmedi\\EMA3.mdb"
    SaveSetting App.EXEName, "frmPreferances", "Ht", "19751217"
    SaveSetting App.EXEName, "frmCharges", "BackColour", "14.9"
    SaveSetting App.EXEName, "frmDeposit", "CanDeposit", True
End Sub

