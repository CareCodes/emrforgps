VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim CodeValue As String
    Dim IDValue As Long
    Dim AgeInWordsValue As String
    Dim AgeInDaysValue As Long
    Dim AgeInMonthsValue As Long
    Dim AgeInYearsValue As Long
    Dim SexValueValue As String
    Dim SexIDValue As Long
    Dim AddressValue As String
    Dim DateOfBirthValue As Date
    Dim TelephoneValue As String
    Dim MobileValue As String
    Dim BloodGroupValue As String
    Dim BloodGroupIDValue As Long
    Dim PatientValue As String
    Dim RaceValue As String
    Dim RaceIDValue As Long
    Dim GivenNameValue As String
    Dim FamilyNameValue As String
    Dim TitleValue As String
    Dim TitleIDValue As Long
    Dim CommentsValue As String
    Dim FullNameValue As String
    Dim eMailValue As String
    
    
    Dim NameWithTitleValue As String
    Dim temSql As String
    Dim CivilStateIDValue As Long
    Dim rsTem As New ADODB.Recordset


Public Property Let ID(PatientID As Long)
    Call ClearValues
    IDValue = PatientID
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatient.*, tblSex.Sex, tblTitle.Title, tblBloodGroup.BloodGroup, tblCivilStatus.CivilStatus, tblRace.Race " & _
                    "FROM ((((tblPatient LEFT JOIN tblSex ON tblPatient.SexID = tblSex.SexID) LEFT JOIN tblTitle ON tblPatient.TitleID = tblTitle.TitleID) LEFT JOIN tblBloodGroup ON tblPatient.BloodGroupID = tblBloodGroup.BloodGroupID) LEFT JOIN tblCivilStatus ON tblPatient.CivilStatusID = tblCivilStatus.CivilStatusID) LEFT JOIN tblRace ON tblPatient.RaceID = tblRace.RaceID " & _
                    "WHERE (((tblPatient.PatientID)=" & IDValue & "))"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Patient) Then PatientValue = !Patient
            If Not IsNull(!Address) Then AddressValue = !Address
            If Not IsNull(!DateOfBirth) Then DateOfBirthValue = !DateOfBirth
            If Not IsNull(!Telephone) Then TelephoneValue = !Telephone
            If Not IsNull(!Mobile) Then MobileValue = !Mobile
            If Not IsNull(!Sex) Then SexValueValue = !Sex
            If Not IsNull(!Title) Then TitleValue = !Title
            If Not IsNull(!GivenName) Then GivenNameValue = !GivenName
            If Not IsNull(!FamilyName) Then FamilyNameValue = !FamilyName
            If Not IsNull(!BloodGroupID) Then BloodGroupValue = !BloodGroupID
            If Not IsNull(!Race) Then RaceValue = !Race
            If Not IsNull(!PatientCode) Then CodeValue = !PatientCode
            If Not IsNull(!Comments) Then CommentsValue = Format(!Comments, "")
            If Not IsNull(!FullName) Then FullNameValue = !FullName
            If Not IsNull(!eMail) Then eMailValue = !eMail
            NameWithTitleValue = TitleValue & " " & PatientValue
            AgeInWordsValue = CalculateAgeInWords(DateOfBirthValue)
            AgeInYearsValue = DateDiff("yyyy", DateOfBirthValue, Date)
            AgeInMonthsValue = DateDiff("m", DateOfBirthValue, Date)
            AgeInDaysValue = DateDiff("d", DateOfBirthValue, Date)
            SexIDValue = !SexID
            TitleIDValue = !TitleID
            BloodGroupIDValue = !BloodGroupID
            CivilStateIDValue = !CivilStatusID
            RaceIDValue = !RaceID
        End If
        .Close
    End With
End Property

Private Sub ClearValues()
    AddressValue = Empty
    DateOfBirthValue = Empty
    TelephoneValue = Empty
    MobileValue = Empty
    RaceValue = Empty
    BloodGroupValue = Empty
    SexValueValue = Empty
    TitleValue = Empty
    GivenNameValue = Empty
    FamilyNameValue = Empty
    NameWithTitleValue = Empty
    AgeInWordsValue = Empty
    AgeInYearsValue = Empty
    AgeInMonthsValue = Empty
    AgeInDaysValue = Empty
    CodeValue = Empty
    SexIDValue = Empty
    TitleIDValue = Empty
    BloodGroupIDValue = Empty
    CivilStateIDValue = Empty
    RaceIDValue = Empty
    CommentsValue = Empty
    FullNameValue = Empty
    eMailValue = Empty
End Sub

Public Property Get ID() As Long
    ID = IDValue
End Property

Public Property Get Patient() As String
    Patient = PatientValue
End Property

Public Property Get AgeInWords() As String
    AgeInWords = AgeInWordsValue
End Property

Public Property Get AgeInDays() As Long
    AgeInDays = AgeInDaysValue
End Property

Public Property Get AgeInMonths() As Long
    AgeInMonths = AgeInMonthsValue
End Property

Public Property Get AgeInYears() As Long
    AgeInYears = AgeInYearsValue
End Property

Public Property Get DateOfBirth() As Date
    DateOfBirth = DateOfBirthValue
End Property

Public Property Get Sex() As String
    Sex = SexValueValue
End Property

Public Property Get Title() As String
    Title = TitleValue
End Property

Public Property Get NameWithTitle() As String
    NameWithTitle = NameWithTitleValue
End Property

Public Property Get Address() As String
    Address = AddressValue
End Property

Public Property Get Telephone() As String
    Telephone = TelephoneValue
End Property

Public Property Get Mobile() As String
    Mobile = MobileValue
End Property

Public Property Get BloodGroup() As String
    BloodGroup = BloodGroupValue
End Property

Public Property Get Race() As String
    Race = RaceValue
End Property

Public Property Get Code() As String
    Code = CodeValue
End Property

Public Property Get SexID() As Long
    SexID = SexIDValue
End Property

Public Property Get RaceID() As Long
    RaceID = RaceIDValue
End Property
Public Property Get BloodGroupID() As Long
    BloodGroupID = BloodGroupIDValue
End Property
Public Property Get CivilStatusID() As Long
    CivilStatusID = CivilStateIDValue
End Property
Public Property Get TitleID() As Long
    TitleID = TitleIDValue
End Property
Public Property Get Comments() As String
    Comments = CommentsValue
End Property
Public Property Get FullName() As String
    FullName = FullNameValue
End Property
Public Property Get eMail() As String
    eMail = eMailValue
End Property
