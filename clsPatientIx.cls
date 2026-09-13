VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatientIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim PatientIxIDValue As Long
    Dim PatientIxBillIDValue As Long
    Dim IxIDValue As Long
'    Dim VisitIxIDValue As Long
    Dim PatientIDValue As Long
    Dim VisitIDValue As Long
    Dim DoctorIDValue As Long
    Dim InstitutionIDValue As Long
    Dim SpecimanNoValue As String
    Dim SpecimanCommentsValue As String
    Dim ValueValue As Double
    Dim CostValue As Double
    Dim RequestedValue As Boolean
    Dim RequestedDateValue As Date
    Dim RequestedTimeValue As Date
    Dim RequestedUserIDValue As Long
    Dim BilledValue As Boolean
    Dim BilledDateValue As Date
    Dim BilledTimeValue As Date
    Dim BilledUserIDValue As Long
    Dim SavedValue As Boolean
    Dim SavedDateValue As Date
    Dim SavedTimeValue As Date
    Dim SavedUserIDValue As Long
    Dim PrintedValue As Boolean
    Dim PrintedDateValue As Date
    Dim PrintedTimeValue As Date
    Dim PrintedUserIDValue As Long
    Dim IssuedValue As Boolean
    Dim IssuedDateValue As Date
    Dim IssuedTimeValue As Date
    Dim IssuedUserIDValue As Long
    Dim SeenValue As Boolean
    Dim SeenDateValue As Date
    Dim SeenTimeValue As Date
    Dim SeenUserIDValue As Long
    Dim DeletedValue As Boolean
    Dim DeletedUserIDValue As Long
    Dim DeletedDateValue As Date
    Dim DeletedTimeValue As Date
    Dim PaymentMethodIDValue As Long
    Dim CancelledValue As Boolean
    Dim temSql As String
    

Public Property Let ID(ID As Long)
    Dim rsPatientIx As New ADODB.Recordset
    temSql = "Select * from tblPatientIx where PatientIx = " & ID
    With rsPatientIx
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            PatientIxIDValue = !PatientIxID
            PatientIxBillIDValue = !PatientIxBillID
            IxIDValue = !IxID
            PatientIDValue = !PatientID
            VisitIDValue = !VisitID
'            VisitIxIDValue = !VisitIxID
            DoctorIDValue = !DoctorID
            InstitutionIDValue = !InstitutionID
            SpecimanNoValue = !SpecimanNo
            SpecimanCommentsValue = !SpecimanComments
            ValueValue = !Value
            CostValue = !Cost
            RequestedValue = !Requested
            RequestedDateValue = !RequestedDate
            RequestedTimeValue = !RequestedTime
            RequestedUserIDValue = !RequestedUserID
            BilledValue = !Billed
            BilledDateValue = !BilledDate
            BilledTimeValue = !BilledTime
            BilledUserIDValue = !BilledUserID
            SavedValue = !Saved
            SavedDateValue = !SavedDate
            SavedTimeValue = !SavedTime
            SavedUserIDValue = !SavedUserID
            PrintedValue = !Printed
            PrintedDateValue = !PrintedDate
            PrintedTimeValue = !PrintedTime
            PrintedUserIDValue = !PrintedUserID
            IssuedValue = !Issued
            IssuedDateValue = !IssuedDate
            IssuedTimeValue = !IssuedTime
            IssuedUserIDValue = !IssuedUserID
            SeenValue = !Seen
            SeenDateValue = !SeenDate
            SeenTimeValue = !SeenTime
            SeenUserIDValue = !SeenUserID
            DeletedValue = !Deleted
            DeletedUserIDValue = !DeletedUserID
            DeletedDateValue = !DeletedDate
            DeletedTimeValue = !DeletedTime
        Else
            PatientIxIDValue = Empty
            PatientIxBillIDValue = Empty
            IxIDValue = Empty
            PatientIDValue = Empty
            VisitIDValue = Empty
'            VisitIxIDValue = Empty
            DoctorIDValue = Empty
            InstitutionIDValue = Empty
            SpecimanNoValue = Empty
            SpecimanCommentsValue = Empty
            ValueValue = Empty
            CostValue = Empty
            RequestedValue = Empty
            RequestedDateValue = Empty
            RequestedTimeValue = Empty
            RequestedUserIDValue = Empty
            BilledValue = Empty
            BilledDateValue = Empty
            BilledTimeValue = Empty
            BilledUserIDValue = Empty
            SavedValue = Empty
            SavedDateValue = Empty
            SavedTimeValue = Empty
            SavedUserIDValue = Empty
            PrintedValue = Empty
            PrintedDateValue = Empty
            PrintedTimeValue = Empty
            PrintedUserIDValue = Empty
            IssuedValue = Empty
            IssuedDateValue = Empty
            IssuedTimeValue = Empty
            IssuedUserIDValue = Empty
            SeenValue = Empty
            SeenDateValue = Empty
            SeenTimeValue = Empty
            SeenUserIDValue = Empty
            DeletedValue = Empty
            DeletedUserIDValue = Empty
            DeletedDateValue = Empty
            DeletedTimeValue = Empty
        End If
        .Close
    End With
End Property


Public Property Get PatientIxID() As Long
    PatientIxID = PatientIxIDValue
End Property

Public Property Get PatientIxBillID() As Long
    PatientIxBillID = PatientIxBillIDValue
End Property

Public Property Get IxID() As Long
    IxID = IxIDValue
End Property

Public Property Get PatientID() As Long
    PatientID = PatientIDValue
End Property

Public Property Get VisitID() As Long
    VisitID = VisitIDValue
End Property

'Public Property Get VisitIxID() As Long
'    VisitIxID = VisitIxIDValue
'End Property

Public Property Get DoctorID() As Long
    DoctorID = DoctorIDValue
End Property

Public Property Get InstitutionID() As Long
    InstitutionID = InstitutionIDValue
End Property

Public Property Get SpecimanNo() As String
    SpecimanNo = SpecimanNoValue
End Property

Public Property Get SpecimanComments() As String
    SpecimanComments = SpecimanCommentsValue
End Property

Public Property Get Value() As Double
     Value = ValueValue
End Property

Public Property Get Cost() As Double
    Cost = CostValue
End Property

Public Property Get Requested() As Boolean
    Requested = RequestedValue
End Property

Public Property Get RequestedDate() As Date
    RequestedDate = RequestedDateValue
End Property

Public Property Get RequestedTime() As Date
    RequestedTime = RequestedTimeValue
End Property

Public Property Get RequestedUserID() As Long
    RequestedUserID = RequestedUserIDValue
End Property

Public Property Get Billed() As Boolean
    Billed = BilledValue
End Property

Public Property Get BilledDate() As Date
    BilledDate = BilledDateValue
End Property

Public Property Get BilledTime() As Date
    BilledTime = BilledTimeValue
End Property

Public Property Get BilledUserID() As Long
    BilledUserID = BilledUserIDValue
End Property

Public Property Get Saved() As Boolean
    Saved = SavedValue
End Property

Public Property Get SavedDate() As Date
    SavedDate = SavedDateValue
End Property

Public Property Get SavedTime() As Date
    SavedTime = SavedTimeValue
End Property

Public Property Get SavedUserID() As Long
    SavedUserID = SavedUserIDValue
End Property

Public Property Get Printed() As Boolean
    Printed = PrintedValue
End Property

Public Property Get PrintedDate() As Date
    PrintedDate = PrintedDateValue
End Property

Public Property Get PrintedTime() As Date
    PrintedTime = PrintedTimeValue
End Property

Public Property Get PrintedUserID() As Long
    PrintedUserID = PrintedUserIDValue
End Property

Public Property Get Issued() As Boolean
    Issued = IssuedValue
End Property

Public Property Get IssuedDate() As Date
    IssuedDate = IssuedDateValue
End Property

Public Property Get IssuedTime() As Date
    IssuedTime = IssuedTimeValue
End Property

Public Property Get IssuedUserID() As Long
    IssuedUserID = IssuedUserIDValue
End Property

Public Property Get Seen() As Boolean
    Seen = SeenValue
End Property

Public Property Get SeenDate() As Date
    SeenDate = SeenDateValue
End Property

Public Property Get SeenTime() As Date
    SeenTime = SeenTimeValue
End Property

Public Property Get SeenUserID() As Long
    SeenUserID = SeenUserIDValue
End Property

Public Property Get Deleted() As Boolean
    Deleted = DeletedValue
End Property

Public Property Get DeletedUserID() As Long
    DeletedUserID = DeletedUserIDValue
End Property

Public Property Get DeletedDate() As Date
    DeletedDate = DeletedDateValue
End Property

Public Property Get DeletedTime() As Date
    DeletedTime = DeletedTimeValue
End Property


