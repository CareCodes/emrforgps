VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
    Dim temSql As String
    Dim rsVisit As New ADODB.Recordset
    Dim rsQueue As New ADODB.Recordset
    
    Dim QueueIDValue As Long
    Dim QueueNoValue As String
    Dim SecessionIDValue As Long
    Dim SecessionValue As String
    Dim VisitIDValue As Long
    Dim PatientIDValue As Long
    Dim VisitDateValue As Date
    Dim VisitTimeValue As Date
    Dim VisitBPValue As String
    Dim VisitWeightValue As Double
    Dim ConsultationFeeValue As Double
    Dim MedicinesFeeValue As Double
    Dim ProcedureFeeValue As Double
    Dim TotalFeeValue As Double
    Dim ConsultationCostValue As Double
    Dim MedicinesCostValue As Double
    Dim ProcedureCostValue As Double
    Dim TotalCostValue As Double

    
    Dim PaymentMethodIDValue As Long
    Dim DrawerIDValue As Long
    
Public Function NewVisitID(NewPatientID As Long, NewVisitDate As Date, DoctorID As Long, Optional SecessionID As Long, Optional QueueID As Long) As Long
    Call ClearValues
    With rsVisit
        If .State = 1 Then .Close
        temSql = "Select * from tblVisit"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PatientID = NewPatientID
        !VisitDate = NewVisitDate
        !VisitTime = Time
        !BookedUserID = UserID
        !BookedTIme = Now
        !ConsultationFee = UsualConsultationFee
        !TotalFee = UsualConsultationFee
        !DoctorID = DoctorID
        !PaymentMethodID = Val(GetSetting(App.EXEName, "frmNewVisit", "cmbPaymentMethod", 1))
        !DrawerID = Val(GetSetting(App.EXEName, "frmNewVisit", "cmbDrawer", 1))
        PaymentMethodIDValue = !PaymentMethodID = Val(GetSetting(App.EXEName, "frmNewVisit", "cmbPaymentMethod", 1))
        DrawerIDValue = Val(GetSetting(App.EXEName, "frmNewVisit", "cmbDrawer", 1))
        ChangeDrawerBalances !DrawerID, !PaymentMethodID, !TotalFee, True, False, False
        If SecessionID <> 0 Then
            SecessionIDValue = SecessionID
            !SecessionID = SecessionID
        Else
            SecessionIDValue = 0
            SecessionValue = Empty
        End If
        If QueueID <> 0 Then
            QueueIDValue = QueueID
            !QueueID = QueueID
        Else
            QueueIDValue = 0
            QueueNoValue = Empty
        End If
        .Update
        NewVisitID = !VisitID
        VisitIDValue = !VisitID
        PatientIDValue = NewPatientID
        VisitDateValue = NewVisitDate
        VisitTimeValue = Time
        ConsultationFeeValue = UsualConsultationFee
        MedicinesFeeValue = 0
        ProcedureFeeValue = 0
        TotalFeeValue = UsualConsultationFee
        ConsultationCostValue = UsualConsultationCost
        MedicinesCostValue = 0
        ProcedureCostValue = 0
        TotalCostValue = UsualConsultationCost
        
    End With
    With rsQueue
        If .State = 1 Then .Close
        temSql = "SELECT tblQueue.QueueNo, tblSecession.SecessionID, tblSecession.Secession FROM tblQueue LEFT JOIN tblSecession ON tblQueue.SecessionID = tblSecession.SecessionID WHERE (((tblQueue.QueueID)=" & QueueIDValue & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            QueueNoValue = Format(!QueueNo, "")
            SecessionValue = Format(!Secession, "")
        End If
        .Close
    End With
End Function

Public Property Let VisitID(ID As Long)
    Call ClearValues
    With rsVisit
        If .State = 1 Then .Close
        temSql = "Select * from tblVisit where VisitID = " & ID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            PatientIDValue = !PatientID
            VisitDateValue = !VisitDate
            VisitTimeValue = !VisitTime
            TotalFeeValue = !TotalFee
            ConsultationFeeValue = !ConsultationFee
            ProcedureFeeValue = !ProcedureFee
            MedicinesFeeValue = !MedicineFee
            TotalCostValue = !TotalCost
            ConsultationCostValue = !ConsultationCost
            ProcedureCostValue = !ProcedureCost
            MedicinesCostValue = !MedicineCost
            VisitIDValue = ID
            If Not IsNull(!DrawerID) Then
                DrawerIDValue = !DrawerID
            Else
                DrawerIDValue = 0
            End If
            If Not IsNull(!PaymentMethodID) Then
                PaymentMethodIDValue = !PaymentMethodID
            Else
                PaymentMethodIDValue = 0
            End If
            QueueIDValue = !QueueID
            VisitBPValue = !SBP & "/" & !DBP & " mmHg"
            VisitWeightValue = !VisitWeight
        Else
        
        End If
        .Close
    End With
    With rsQueue
        If .State = 1 Then .Close
        temSql = "SELECT tblQueue.QueueNo, tblSecession.SecessionID, tblSecession.Secession FROM tblQueue LEFT JOIN tblSecession ON tblQueue.SecessionID = tblSecession.SecessionID WHERE (((tblQueue.QueueID)=" & QueueIDValue & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            QueueNoValue = Format(!QueueNo, "")
            SecessionValue = Format(!Secession, "")
            SecessionIDValue = !SecessionID
        End If
        .Close
    End With
End Property

Private Sub ClearValues()
    PatientIDValue = Empty
    VisitDateValue = Empty
    VisitTimeValue = Empty
    TotalFeeValue = Empty
    ConsultationFeeValue = Empty
    ProcedureFeeValue = Empty
    MedicinesFeeValue = Empty
    TotalCostValue = Empty
    ConsultationCostValue = Empty
    ProcedureCostValue = Empty
    MedicinesCostValue = Empty
    QueueIDValue = 0
    QueueNoValue = Empty
    SecessionValue = Empty
    SecessionIDValue = 0
    VisitBPValue = Empty
    PaymentMethodIDValue = 0
    DrawerIDValue = 0
    VisitWeightValue = 0
End Sub

Public Sub Refresh()
    Call ClearValues
    With rsVisit
        If .State = 1 Then .Close
        temSql = "Select * from tblVisit where VisitID = " & VisitIDValue
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            PatientIDValue = !PatientID
            VisitDateValue = !VisitDate
            VisitTimeValue = !VisitTime
            TotalFeeValue = !TotalFee
            ConsultationFeeValue = !ConsultationFee
            ProcedureFeeValue = !ProcedureFee
            MedicinesFeeValue = !MedicineFee
            TotalCostValue = !TotalCost
            ConsultationCostValue = !ConsultationCost
            ProcedureCostValue = !ProcedureCost
            MedicinesCostValue = !MedicineCost
            QueueIDValue = !QueueID
            VisitBPValue = !SBP & "/" & !DBP & " mmHg"
            VisitWeightValue = !VisitWeight
            If Not IsNull(!DrawerID) Then
                DrawerIDValue = !DrawerID
            Else
                DrawerIDValue = 0
            End If
            If Not IsNull(!PaymentMethodID) Then
                PaymentMethodIDValue = !PaymentMethodID
            Else
                PaymentMethodIDValue = 0
            End If
            
        Else
        
        End If
        .Close
    End With
    With rsQueue
        If .State = 1 Then .Close
        temSql = "SELECT tblQueue.QueueNo, tblSecession.SecessionID, tblSecession.Secession FROM tblQueue LEFT JOIN tblSecession ON tblQueue.SecessionID = tblSecession.SecessionID WHERE (((tblQueue.QueueID)=" & QueueIDValue & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            QueueNoValue = Format(!QueueNo, "")
            SecessionValue = Format(!Secession, "")
            SecessionIDValue = !SecessionID
        End If
        .Close
    End With
End Sub

Public Property Get VisitID() As Long
    VisitID = VisitIDValue
End Property

Public Property Get VisitDate() As Long
    VisitDate = VisitDateValue
End Property

Public Property Get PatientID() As Long
    PatientID = PatientIDValue
End Property

Public Property Get VisitTime() As Date
    VisitTime = VisitTimeValue
End Property

Public Property Get ConsultationFee() As Double
    ConsultationFee = ConsultationFeeValue
End Property

Public Property Get ProcedureFee() As Double
    ProcedureFee = ProcedureFeeValue
End Property

Public Property Get MedicinesFee() As Double
    MedicinesFee = MedicinesFeeValue
End Property

Public Property Get TotalFee() As Double
    TotalFee = TotalFeeValue
End Property

Public Property Get ConsultationCost() As Double
    ConsultationCost = ConsultationCostValue
End Property

Public Property Get ProcedureCost() As Double
    ProcedureCost = ProcedureCostValue
End Property

Public Property Get MedicinesCost() As Double
    MedicinesCost = MedicinesCostValue
End Property

Public Property Get TotalCost() As Double
    TotalCost = TotalCostValue
End Property

Public Property Get QueueNo() As String
    QueueNo = QueueNoValue
End Property

Public Property Get QueueID() As Long
    QueueID = QueueIDValue
End Property

Public Property Get SecessionID() As Long
    SecessionID = SecessionIDValue
End Property

Public Property Get Secession() As String
    Secession = SecessionValue
End Property

Public Property Get VisitBP() As String
    VisitBP = VisitBPValue
End Property

Public Property Get VisitWeight() As Double
    VisitWeight = VisitWeightValue
End Property

Public Property Get PaymentMethodID() As Long
    PaymentMethodID = PaymentMethodIDValue
End Property

Public Property Get DrawerID() As Long
    DrawerID = DrawerIDValue
End Property
