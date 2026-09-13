Attribute VB_Name = "ModuleFind"
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSQL As String
    Dim SpellCheck As Object


Public Function IsAllergic(PatientID As Long, ItemID As Long) As Boolean
    Dim AllergicItem As New clsItem
    Dim ThisItem As New clsItem
    IsAllergic = True
    ThisItem.ID = ItemID
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPatientAllergy where PatientID = " & PatientID & " AND Deleted = false"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            AllergicItem.ID = !ItemID
            If AllergicItem.IsGroup = True Then
                If ThisItem.GroupID <> 0 Then
                    If ThisItem.GroupID = AllergicItem.GroupID Then Exit Function
                End If
            ElseIf AllergicItem.IsSubGroup = True Then
                If ThisItem.SubGroupID <> 0 Then
                    If ThisItem.SubGroupID = AllergicItem.SubGroupID Then Exit Function
                End If
            ElseIf AllergicItem.IsTradeName = True Then
                If ThisItem.GenericNameID <> 0 Then
                    If ThisItem.GenericNameID = AllergicItem.GenericNameID Then Exit Function
                End If
            ElseIf AllergicItem.IsGenericName = True Then
                If ThisItem.GenericNameID <> 0 Then
                    If ThisItem.GenericNameID = AllergicItem.GenericNameID Then Exit Function
                End If
            ElseIf AllergicItem.IsItem = True Then
                If ThisItem.GenericNameID <> 0 Then
                    If ThisItem.GenericNameID = AllergicItem.GenericNameID Then Exit Function
                End If
            End If
            .MoveNext
        Wend
    End With
    IsAllergic = False
End Function

Public Function GetLastQueueNo(DoctorID As Long, QueueDate As Date, SecessionID As Long) As Long
    Dim rsQueue As New ADODB.Recordset
    Dim i As Integer
    Randomize
    i = Rnd(1000)
    For i = 0 To i
        
    Next
    With rsQueue
        If .State = 1 Then .Close
        temSQL = "Select count(QueueID) as QueueCount from tblQueue where DoctorID = " & DoctorID & " AND QueueDate = #" & Format(QueueDate) & "# AND SecessionID = " & SecessionID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If IsNull(!QueueCount) = False Then
            GetLastQueueNo = !QueueCount + 1
        Else
            GetLastQueueNo = 1
        End If
    
    End With
End Function

Public Function LastPurchasePrice(ItemID As Long) As Double
    LastPurchasePrice = 0
    Dim rsTemPP As New ADODB.Recordset
    With rsTemPP
        If .State = 1 Then .Close
        temSQL = "SELECT tblPurchaseItem.PurchaseRate " & _
                    "FROM tblPurchaseBill LEFT JOIN tblPurchaseItem ON tblPurchaseBill.PurchaseBillID = tblPurchaseItem.PurchaseBillID " & _
                    "Where (((tblPurchaseBill.Completed) = True) And ((tblPurchaseItem.ItemID) = " & ItemID & ")) " & _
                    "ORDER BY tblPurchaseBill.CompletedDateTime DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!PurchaseRate) = False Then LastPurchasePrice = !PurchaseRate
        End If
        .Close
    End With
End Function

Public Function LastSalePrice(ItemID As Long) As Double
    LastSalePrice = 0
    Dim rsTemPP As New ADODB.Recordset
    With rsTemPP
        If .State = 1 Then .Close
        temSQL = "SELECT tblPurchaseItem.SaleRate " & _
                    "FROM tblPurchaseBill LEFT JOIN tblPurchaseItem ON tblPurchaseBill.PurchaseBillID = tblPurchaseItem.PurchaseBillID " & _
                    "Where (((tblPurchaseBill.Completed) = True) And ((tblPurchaseItem.ItemID) = " & ItemID & ")) " & _
                    "ORDER BY tblPurchaseBill.CompletedDateTime DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SaleRate) = False Then LastSalePrice = !SaleRate
        End If
        .Close
    End With
End Function

Public Function ItemBatchID(ItemID As Long, ItemBatch As String, DOE As Date, DOM As Date) As Long
    Dim rsBatch As New ADODB.Recordset
    With rsBatch
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemBatch where ItemID = " & ItemID & "  AND ItemBatch = '" & ItemBatch & "' AND DOE = #" & Format(DOE, "dd MMMM yyyy") & "#"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
            !ItemID = ItemID
            !ItemBatch = ItemBatch
            !DOE = DOE
            !DOM = DOM
            .Update
        End If
        ItemBatchID = !ItemBatchID
        .Close
    End With
End Function

Public Function ChangeStocks(ItemID As Long, ItemBatchID As Long, Quentity As Double, DepartmentID As Long, ToAdd As Boolean, ToDeduct As Boolean, ToReset As Boolean, Optional NoMinus As Boolean) As Boolean
    ChangeStocks = True
    Dim rsStock As New ADODB.Recordset
    With rsStock
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemBatchStock where ItemID = " & ItemID & " AND ItemBatchID = " & ItemBatchID & " AND DepartmentID = " & DepartmentID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
            !ItemID = ItemID
            !ItemBatchID = ItemBatchID
            !DepartmentID = DepartmentID
        End If
        If ToAdd = True Then
            !Stock = !Stock + Quentity
        ElseIf ToDeduct = True Then
            If !Stock < Quentity Then
                If NoMinus = True Then
                    ChangeStocks = False
                Else
                    !Stock = !Stock - Quentity
                End If
            Else
                !Stock = !Stock - Quentity
            End If
        ElseIf ToReset = True Then
            !Stock = Quentity
        End If
        .Update
        .Close
    End With
End Function

Public Function CalculateStock(ByVal ItemID As Long, Optional ByVal BatchID As Long, Optional ByVal StoreID As Long) As ItemStock
    With rsTem
        temSQL = "SELECT Sum([tblItemBatchStock].[Stock]) AS TotalStock, Min([tblItemBatch].[DOE]) AS MinDOE, Max([tblItemBatch].[DOE]) AS MaxDOE FROM tblItemBatch LEFT JOIN tblItemBatchStock ON tblItemBatch.ItemBatchID = tblItemBatchStock.ItemBatchID " & _
                    " WHERE tblItemBatch.ItemID=" & ItemID & " "
        If BatchID <> 0 Then temSQL = temSQL & " AND tblItemBatch.ItemBatchID =" & BatchID & " "
        If StoreID <> 0 Then temSQL = temSQL & "AND tblItemBatchStock.DepartmentID=" & StoreID & " "
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If Not IsNull(!TotalStock) Then
            CalculateStock.Stock = !TotalStock
        Else
            CalculateStock.Stock = 0
        End If
        If Not IsNull(!MinDOE) Then
            CalculateStock.MinDOE = !MinDOE
        Else
            CalculateStock.MinDOE = TimeSerial(0, 0, 0)
        End If
        If Not IsNull(!MaxDOE) Then
            CalculateStock.MaxDOE = !MaxDOE
        Else
            CalculateStock.MaxDOE = TimeSerial(0, 0, 0)
        End If
    End With
End Function


Public Function SpellChecker(Textbox As String) As String
    On Error GoTo ErrorHandler
     
    Set SpellCheck = CreateObject("Word.Application")      'Set Up Reference To WORD
    SpellCheck.Visible = False                            'Hide WORD From User
    SpellCheck.Documents.Add                              'Open New Document (Hidden)
    Clipboard.Clear
    Clipboard.SetText Textbox                          'Copy Text To Clipboard
    SpellCheck.Selection.Paste                            'Paste Text Into WORD
    SpellCheck.ActiveDocument.CheckSpelling               'Activate The Spell Checker
    SpellCheck.Visible = False                            'Hide WORD From User
    SpellCheck.ActiveDocument.Select                      'Select The Corrected Text
    SpellCheck.Selection.Cut                              'Cut The Text To Clipboard
    SpellChecker = Clipboard.GetText                      'Assign Text To SPELLCHECKER Function
    SpellCheck.ActiveDocument.Close SaveChanges:= _
     wdDoNotSaveChanges                                   'Close & DO NOT Save Changes
    SpellCheck.Quit                                       'Quit
    Set SpellCheck = Nothing                               'Release Reference
     
     
    
    Exit Function    'Exit To Avoid ErrorHandler.
    
ErrorHandler:
      SpellCheck.ActiveDocument.Close SaveChanges:= _
       wdDoNotSaveChanges                                   'Close & DO NOT Save Changes
      SpellCheck.Quit                                       'Quit
     Set SpellCheck = Nothing
     
     MsgBox "Error: " & Err.Number & "  " & Err.Description, vbExclamation, "School Report Card Spell Check Error"
    
End Function


Public Function AddNewPatient(Name As String, Optional DateOfBirth As Date, Optional SexID As Long, Optional TitleID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    Dim TemDate As Date
    AddNewPatient = 0
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPatient"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Patient = Name
        !TitleID = TitleID
        !SexID = SexID
        If DateOfBirth <> TemDate Then
            !DateOfBirth = Date
        End If
        .Update
        AddNewPatient = !PatientID
        .Close
    End With
    Set rsTem = Nothing
End Function



Public Function CalIssueQty(ItemID As Long, Dose As Double, FrequencyUnitID As Long, Duration As Double, DurationUnitID As Long, Optional MinIQty As Double) As Double
    Dim Fint As Double
    Dim Dint As Double
    Dim Doses As Double
    Dim CalItem As New clsItem
    
    CalIssueQty = 1
    
    CalItem.ID = ItemID
    
    Dim rsF As New ADODB.Recordset
    With rsF
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemUnit where ItemUnitID = " & FrequencyUnitID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Fint = !Interval
        Else
            Exit Function
        End If
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemUnit where ItemUnitID = " & DurationUnitID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Dint = !Interval
        Else
            Exit Function
        End If
    End With
    If Fint <> 0 Then
        Doses = (Dint * Duration) / Fint
    Else
        Doses = Dint * Duration
    End If
    CalIssueQty = Dose * Doses
    
    If CalItem.DivQty <> 0 Then
        If CalIssueQty / CalItem.DivQty <> 0 Then
            CalIssueQty = ((CalIssueQty \ CalItem.DivQty) * CalItem.DivQty) ' + 1
        End If
    End If
    
    If CalItem.MinIQty <> 0 And CalIssueQty < MinIQty Then
        CalIssueQty = MinIQty
    End If
    
    
    If CalItem.MaxIQty <> 0 And CalIssueQty > CalItem.MaxIQty Then
        CalIssueQty = CalItem.MaxIQty
    End If
    
    If CalIssueQty <= 0 Then CalIssueQty = 1

End Function

Public Function HasIndoorDrugs(VisitID As Long) As Boolean
    Dim rsV As New ADODB.Recordset
    With rsV
        If .State = 1 Then .Close
        temSQL = "Select * from tblVisitItem where VisitID = " & VisitID & " AND Deleted = false And Indoor = True"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            HasIndoorDrugs = True
        Else
            HasIndoorDrugs = False
        End If
        .Close
    End With
End Function

Public Function HasOutdoorDrugs(VisitID As Long) As Boolean
    Dim rsV As New ADODB.Recordset
    With rsV
        If .State = 1 Then .Close
        temSQL = "Select * from tblVisitItem where VisitID = " & VisitID & " AND Deleted = false And Indoor = False"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            HasOutdoorDrugs = True
        Else
            HasOutdoorDrugs = False
        End If
        .Close
    End With
End Function

Public Function GetColumnName(ColumnNo As Long) As String
    Dim temnum As Integer
    Dim temnum1 As Integer
    
    If ColumnNo < 27 Then
        GetColumnName = Chr(ColumnNo + 64)
    Else
        temnum = ColumnNo \ 26
        temnum1 = ColumnNo Mod 26
        GetColumnName = Chr(temnum + 64) & Chr(temnum1 + 64)
    End If
End Function

Public Sub SetPPrice(ItemID As Long, Price As Double)
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !PurchasePrice = Price
            .Update
        End If
        .Close
    End With
End Sub

Public Sub SetSPrice(ItemID As Long, Price As Double)
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !SalePrice = Price
            .Update
        End If
        .Close
    End With
End Sub

Public Sub SetLastPPrice(ItemID As Long, Price As Double)
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !LastPurchasePrice = Price
            .Update
        End If
        .Close
    End With
End Sub

Public Sub SetLastSPrice(ItemID As Long, Price As Double)
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !LastSalePrice = Price
            .Update
        End If
        .Close
    End With
End Sub

Public Function GetPPrice(ItemID As Long) As Double
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            GetPPrice = !PurchasePrice
        End If
        .Close
    End With
End Function

Public Function GetSPrice(ItemID As Long) As Double
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            GetSPrice = !SalePrice
        End If
        .Close
    End With
End Function

Public Function GetLastPPrice(ItemID As Long) As Double
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            GetLastPPrice = !LastPurchasePrice
        End If
        .Close
    End With
End Function

Public Function GetLastSPrice(ItemID As Long) As Double
    Dim rsItem As New ADODB.Recordset
    With rsItem
        temSQL = "Select * from tblItem where ItemID = " & ItemID
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            GetLastSPrice = !LastSalePrice
        End If
        .Close
    End With
End Function

'Public Function CalculateAgeInWordsOld(PatientDOB As Date) As String
'    Dim Age As Long
'    Age = DateDiff("yyyy", PatientDOB, Now)
'    If Age >= 12 Then
'        CalculateAgeInWords = Age & " Years"
'        Exit Function
'    Else
'        Age = DateDiff("m", PatientDOB, Now)
'
'        If Age > 144 Then CalculateAgeInWords = "12" & " Years and " & Age - 144 & " months": Exit Function
'        If Age = 144 Then CalculateAgeInWords = "12" & " Years": Exit Function
'        If Age > 132 Then CalculateAgeInWords = "11" & " Years and " & Age - 132 & " months": Exit Function
'        If Age = 132 Then CalculateAgeInWords = "11" & " Years": Exit Function
'
'        If Age > 120 Then CalculateAgeInWords = "10" & " Years and " & Age - 120 & " months": Exit Function
'        If Age = 120 Then CalculateAgeInWords = "10" & " Years": Exit Function
'        If Age > 108 Then CalculateAgeInWords = "9" & " Years and " & Age - 108 & " months": Exit Function
'        If Age = 108 Then CalculateAgeInWords = "9" & " Years": Exit Function
'        If Age > 96 Then CalculateAgeInWords = "8" & " Years and " & Age - 96 & " months": Exit Function
'        If Age = 96 Then CalculateAgeInWords = "8" & " Years": Exit Function
'
'        If Age > 84 Then CalculateAgeInWords = "7" & " Years and " & Age - 84 & " months": Exit Function
'        If Age = 84 Then CalculateAgeInWords = "7" & " Years": Exit Function
'        If Age > 72 Then CalculateAgeInWords = "6" & " Years and " & Age - 72 & " months": Exit Function
'        If Age = 72 Then CalculateAgeInWords = "6" & " Years": Exit Function
'        If Age > 60 Then CalculateAgeInWords = "5" & " Years and " & Age - 60 & " months": Exit Function
'        If Age = 60 Then CalculateAgeInWords = "5" & " Years": Exit Function
'
'
'        If Age > 48 Then CalculateAgeInWords = "4" & " Years and " & Age - 48 & " months": Exit Function
'        If Age = 48 Then CalculateAgeInWords = "4" & " Years": Exit Function
'        If Age > 36 Then CalculateAgeInWords = "3" & " Years and " & Age - 36 & " months": Exit Function
'        If Age = 36 Then CalculateAgeInWords = "3" & " Years": Exit Function
'        If Age > 24 Then CalculateAgeInWords = "2" & " Years and " & Age - 24 & " months": Exit Function
'        If Age = 24 Then CalculateAgeInWords = "2" & " Years": Exit Function
'        If Age > 12 Then CalculateAgeInWords = "1" & " Years and " & Age - 12 & " months": Exit Function
'        If Age = 12 Then CalculateAgeInWords = "1" & " Year": Exit Function
'        If Age >= 1 Then CalculateAgeInWords = Age & " Months": Exit Function
'        Age = DateDiff("d", PatientDOB, Now)
'        CalculateAgeInWords = Age & " Days": Exit Function
'        Exit Function
'    End If
'End Function


Public Function CalculateAgeInWords(PatientDOB As Date) As String
    Dim Age As Long
    Dim AgeInMonths As Integer
    Dim AgeInDays As Long
    Dim AgeInYears As Integer
    
    Dim todaysDate As Date
    
    todaysDate = Date
    
    Dim monthBirthday As Date
    
    
    
    AgeInYears = DateDiff("yyyy", PatientDOB, todaysDate)
    AgeInMonths = DateDiff("m", PatientDOB, todaysDate)
    AgeInDays = DateDiff("d", PatientDOB, todaysDate)
    
    
    If AgeInYears > 12 Then
        CalculateAgeInWords = AgeInYears & " Years"
        Exit Function
    End If
    
    If AgeInMonths >= 12 Then
        Age = DateDiff("m", PatientDOB, todaysDate)
        
        If Age > 144 Then CalculateAgeInWords = "12" & " Years and " & Age - 144 & " months": Exit Function
        If Age = 144 Then CalculateAgeInWords = "12" & " Years": Exit Function
        If Age > 132 Then CalculateAgeInWords = "11" & " Years and " & Age - 132 & " months": Exit Function
        If Age = 132 Then CalculateAgeInWords = "11" & " Years": Exit Function
        
        If Age > 120 Then CalculateAgeInWords = "10" & " Years and " & Age - 120 & " months": Exit Function
        If Age = 120 Then CalculateAgeInWords = "10" & " Years": Exit Function
        If Age > 108 Then CalculateAgeInWords = "9" & " Years and " & Age - 108 & " months": Exit Function
        If Age = 108 Then CalculateAgeInWords = "9" & " Years": Exit Function
        If Age > 96 Then CalculateAgeInWords = "8" & " Years and " & Age - 96 & " months": Exit Function
        If Age = 96 Then CalculateAgeInWords = "8" & " Years": Exit Function
        
        If Age > 84 Then CalculateAgeInWords = "7" & " Years and " & Age - 84 & " months": Exit Function
        If Age = 84 Then CalculateAgeInWords = "7" & " Years": Exit Function
        If Age > 72 Then CalculateAgeInWords = "6" & " Years and " & Age - 72 & " months": Exit Function
        If Age = 72 Then CalculateAgeInWords = "6" & " Years": Exit Function
        If Age > 60 Then CalculateAgeInWords = "5" & " Years and " & Age - 60 & " months": Exit Function
        If Age = 60 Then CalculateAgeInWords = "5" & " Years": Exit Function
        
        
        If Age > 48 Then CalculateAgeInWords = "4" & " Years and " & Age - 48 & " months": Exit Function
        If Age = 48 Then CalculateAgeInWords = "4" & " Years": Exit Function
        If Age > 36 Then CalculateAgeInWords = "3" & " Years and " & Age - 36 & " months": Exit Function
        If Age = 36 Then CalculateAgeInWords = "3" & " Years": Exit Function
        If Age > 24 Then CalculateAgeInWords = "2" & " Years and " & Age - 24 & " months": Exit Function
        If Age = 24 Then CalculateAgeInWords = "2" & " Years": Exit Function
        If Age > 12 Then CalculateAgeInWords = "1" & " Years and " & Age - 12 & " months": Exit Function
        If Age = 12 Then CalculateAgeInWords = "1" & " Year": Exit Function
    
    End If
    
    If todaysDate = PatientDOB Then
        CalculateAgeInWords = "Born just today"
        Exit Function
    End If
    
    monthBirthday = DateSerial(Year(todaysDate), Month(todaysDate), Day(PatientDOB))
    
    
    If AgeInDays < 30 Then
        CalculateAgeInWords = AgeInDays & " Days"
        Exit Function
    End If
    
    If monthBirthday = todaysDate Then
        CalculateAgeInWords = "One month"
        Exit Function
    End If
    
    
    If AgeInMonths >= 1 Then
        
        If monthBirthday = todaysDate Then
            CalculateAgeInWords = AgeInMonths & " months"
        ElseIf monthBirthday > todaysDate Then
            CalculateAgeInWords = AgeInMonths - 1 & " months and " & 30 + Day(todaysDate) - Day(PatientDOB) & " days"
        Else
            CalculateAgeInWords = AgeInMonths & " months and " & Day(todaysDate) - Day(PatientDOB) & " days"
        End If
        Exit Function
    
    End If
    

    
    


End Function



Public Function FindTradeNameID(TradeName As String) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblItem where IsTradeName= True and Deleted = False and Item = '" & TradeName & "'"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        
        Else
            .AddNew
            !IsTradeName = True
            !Item = TradeName
            .Update
        End If
        FindTradeNameID = !ItemID
        .Close
        
    End With
    
End Function

Public Function FindGenericNameID(GenericName As String) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblItem where IsGenericName= True and Deleted = False and Item = '" & GenericName & "'"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        
        Else
            .AddNew
            !IsGenericName = True
            !Item = GenericName
            .Update
        End If
        FindGenericNameID = !ItemID
        .Close
        
    End With
    
End Function

Public Function ChangeDrawerBalances(DrawerID As Long, PaymentMethodID As Long, Value As Double, ToAdd As Boolean, ToDeduct As Boolean, ToReset As Boolean, Optional NoMinus As Boolean) As Boolean
    ChangeDrawerBalances = True
    Dim rsStock As New ADODB.Recordset
    With rsStock
        If .State = 1 Then .Close
        temSQL = "Select * from tblDrawerBalance where DrawerID = " & DrawerID & " AND PaymentMethodID = " & PaymentMethodID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
            !DrawerID = DrawerID
            !PaymentMethodID = PaymentMethodID
        End If
        If ToAdd = True Then
            !DrawerBalance = !DrawerBalance + Value
        ElseIf ToDeduct = True Then
            If !DrawerBalance < Value Then
                If NoMinus = True Then
                    ChangeDrawerBalances = False
                Else
                    !DrawerBalance = !DrawerBalance - Value
                End If
            Else
                !DrawerBalance = !DrawerBalance - Value
            End If
        ElseIf ToReset = True Then
            !DrawerBalance = Value
        End If
        .Update
        .Close
    End With
End Function

Public Function KeyCodeToString(KeyCode As Integer, Shift As Integer) As String
    Select Case KeyCode
    
        Case vbKey0: KeyCodeToString = "0"
        Case vbKey1: KeyCodeToString = "1"
        Case vbKey2: KeyCodeToString = "2"
        Case vbKey3: KeyCodeToString = "3"
        Case vbKey4: KeyCodeToString = "4"
        Case vbKey5: KeyCodeToString = "5"
        Case vbKey6: KeyCodeToString = "6"
        Case vbKey7: KeyCodeToString = "7"
        Case vbKey8: KeyCodeToString = "8"
        Case vbKey9: KeyCodeToString = "9"
        
        Case vbKeyA: KeyCodeToString = "A"
        Case vbKeyB: KeyCodeToString = "B"
        Case vbKeyC: KeyCodeToString = "C"
        Case vbKeyD: KeyCodeToString = "D"
        Case vbKeyE: KeyCodeToString = "E"
        Case vbKeyF: KeyCodeToString = "F"
        Case vbKeyG: KeyCodeToString = "G"
        Case vbKeyH: KeyCodeToString = "H"
        Case vbKeyI: KeyCodeToString = "I"
        Case vbKeyJ: KeyCodeToString = "J"
        Case vbKeyK: KeyCodeToString = "K"
        Case vbKeyL: KeyCodeToString = "L"
        Case vbKeyM: KeyCodeToString = "M"
        Case vbKeyN: KeyCodeToString = "N"
        Case vbKeyO: KeyCodeToString = "O"
        Case vbKeyP: KeyCodeToString = "P"
        Case vbKeyQ: KeyCodeToString = "Q"
        Case vbKeyR: KeyCodeToString = "R"
        Case vbKeyS: KeyCodeToString = "S"
        Case vbKeyT: KeyCodeToString = "T"
        Case vbKeyU: KeyCodeToString = "U"
        Case vbKeyV: KeyCodeToString = "V"
        Case vbKeyW: KeyCodeToString = "W"
        Case vbKeyX: KeyCodeToString = "X"
        Case vbKeyY: KeyCodeToString = "Y"
        Case vbKeyZ: KeyCodeToString = "Z"
        
        
    
    End Select
End Function


Public Function MyDraweBalance(Optional DrawerID As Long, Optional PaymentMethodID As Long) As Double
    MyDraweBalance = 0
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT Sum(tblDrawerBalance.DrawerBalance) AS SumOfDrawerBalance FROM tblDrawerBalance "
        If DrawerID <> 0 And PaymentMethodID <> 0 Then
            temSQL = temSQL & "WHERE (((tblDrawerBalance.DrawerID)=" & DrawerID & ") AND ((tblDrawerBalance.PaymentMethodID)=" & PaymentMethodID & "))"
        ElseIf DrawerID <> 0 Then
            temSQL = temSQL & "WHERE tblDrawerBalance.DrawerID  = " & DrawerID
        ElseIf PaymentMethodID <> 0 Then
            temSQL = temSQL & "WHERE tblDrawerBalance.PaymentMethodID = " & PaymentMethodID
        Else
        
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfDrawerBalance) = False Then
                MyDraweBalance = !SumOfDrawerBalance
            End If
        End If
        .Close
    End With
End Function

Public Function FindBatch(ItemID As Long, Batch As String) As ItemBatch
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemBatch where ItemID = " & ItemID & " AND ItemBatch = '" & Batch & "'"
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            FindBatch.DOE = !DOE
            FindBatch.DOM = !DOM
            FindBatch.ItemID = ItemID
            FindBatch.BatchID = !ItemBatchID
            FindBatch.Exists = True
            FindBatch.ItemBatch = !ItemBatch
        Else
            FindBatch.Exists = False
        End If
        .Close
    End With
End Function

Public Function GetControlType(MyControl As Control) As ControlType
    GetControlType = Unknown
    If TypeOf MyControl Is Textbox Then
        GetControlType = Textbox
    ElseIf TypeOf MyControl Is ComboBox Then
        GetControlType = ComboBox
    ElseIf TypeOf MyControl Is Button Then
        GetControlType = Button
    ElseIf TypeOf MyControl Is CheckBox Then
        GetControlType = CheckBox
    ElseIf TypeOf MyControl Is DataCombo Then
        GetControlType = DataCombo
    ElseIf TypeOf MyControl Is DataList Then
        GetControlType = DataList
    ElseIf TypeOf MyControl Is DTPicker Then
        GetControlType = DateTimePicker
    ElseIf TypeOf MyControl Is MSFlexGrid Then
        GetControlType = Grid
    ElseIf TypeOf MyControl Is Label Then
        GetControlType = Label
    ElseIf TypeOf MyControl Is ListBox Then
        GetControlType = ListBox
    ElseIf TypeOf MyControl Is Menu Then
        GetControlType = MenuItem
    ElseIf TypeOf MyControl Is OptionButton Then
        GetControlType = OptionButton
    ElseIf TypeOf MyControl Is SSTab Then
        GetControlType = SSTab
    End If

End Function

Public Function GetControlText(MyControl As Control) As String
    GetControlText = Empty
    If TypeOf MyControl Is Textbox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is ComboBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is Button Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is CheckBox Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is DataCombo Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is DataList Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is DTPicker Then
        GetControlText = Right(MyControl.Name, Len(MyControl.Name) - 3)
    ElseIf TypeOf MyControl Is MSFlexGrid Then
        GetControlText = Right(MyControl.Name, Len(MyControl.Name) - 4)
    ElseIf TypeOf MyControl Is Label Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is ListBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is Menu Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is OptionButton Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is SSTab Then
        GetControlText = MyControl.Caption
    End If
End Function

Public Function GetFormID(FormName As String, FormText As String) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblForm where Form = '" & FormName & "'"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !FormText = FormText
        Else
            .AddNew
            !FormText = FormText
            !Form = FormName
        End If
        .Update
        GetFormID = !FormID
        .Close
    End With
End Function


Public Sub EnableControls(MyForm As Form)
    Dim MyControl As Control
    Dim temText As String
    Dim rsTem As New ADODB.Recordset
    On Error Resume Next
    For Each MyControl In MyForm.Controls
        With rsTem
            If .State = 1 Then .Close
            temSQL = "Select * from tblUserAuthorityControl where AuthorityID = " & UserAuthority & " AND ControlID = " & GetControlID(GetFormID(MyForm.Name, MyForm.Caption), MyControl)
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                MyControl.Enabled = !Enabled
            Else
                MyControl.Enabled = True
            End If
            .Close
        End With
    Next
End Sub

Public Sub VisibleControls(MyForm As Form)
    Dim MyControl As Control
    Dim temText As String
    Dim rsTem As New ADODB.Recordset
    On Error Resume Next
    For Each MyControl In MyForm.Controls
        If AllMenusInvisible(GetControlID(GetFormID(MyForm.Name, MyForm.Caption), MyControl)) = True Then
            MyControl.Visible = False
        Else
            With rsTem
                If .State = 1 Then .Close
                temSQL = "Select * from tblUserAuthorityControl where AuthorityID = " & UserAuthority & " AND ControlID = " & GetControlID(GetFormID(MyForm.Name, MyForm.Caption), MyControl)
                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    MyControl.Visible = !Visible
                Else
                    MyControl.Visible = True
                End If
                .Close
            End With
        End If
    Next
End Sub

Private Function AllMenusInvisible(ControlID As Long) As Boolean
    AllMenusInvisible = False
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblControl WHERE ControlID = " & ControlID & " "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If !MainMenu = True And !ControlType = ControlType.MenuItem Then
                If .State = 1 Then .Close
                temSQL = "SELECT Count(tblUserAuthorityControl.UserAuthorityControlID) AS CountOfUserAuthorityControlID " & _
                            "FROM tblUserAuthorityControl RIGHT JOIN (tblControl RIGHT JOIN tblControl AS tblMainMenu ON tblControl.MainMenuID = tblMainMenu.ControlID) ON tblUserAuthorityControl.ControlID = tblControl.ControlID " & _
                            "WHERE (((tblUserAuthorityControl.Visible)=True) AND ((tblMainMenu.ControlID)=" & ControlID & ") AND ((tblMainMenu.MainMenu)=True) AND ((tblMainMenu.ControlType)=" & ControlType.MenuItem & ") AND ((tblControl.ControlType)=" & ControlType.MenuItem & ")  AND ((tblUserAuthorityControl.AuthorityID)=" & UserAuthority & "))"
                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If IsNull(!CountOfUserAuthorityControlID) = False Then
                        If !CountOfUserAuthorityControlID < 1 Then
                            AllMenusInvisible = True
                        Else
                            AllMenusInvisible = False
                        End If
                    Else
                        AllMenusInvisible = False
                    End If
                Else
                    AllMenusInvisible = False
                End If
            Else
                AllMenusInvisible = False
            End If
        Else
            AllMenusInvisible = False
        End If
        .Close
    End With
End Function

Private Function GetControlID(FormID As Long, MyControl As Control) As Long
    Dim i As Integer
    GetControlID = 0
    Dim rsForm As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
            With rsForm
                If TypeOf MyControl Is SSTab Then
                    For i = 0 To MyControl.Tabs - 1
                        If .State = 1 Then .Close
                        temSQL = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "' AND ControlIndex = " & i
                        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                        MyControl.Tab = i
                        If .RecordCount > 0 Then
                            !ControlText = GetControlText(MyControl)
                        Else
                            .AddNew
                            !FormID = FormID
                            !Control = MyControl.Name
                            !ControlType = GetControlType(MyControl)
                            !ControlText = GetControlText(MyControl)
                            !ControlIndex = i
                        End If
                        .Update
                        GetControlID = !ControlID
                        .Close
                    Next i
                Else
                    If .State = 1 Then .Close
                    temSQL = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "'"
                    .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !ControlText = GetControlText(MyControl)
                    Else
                        .AddNew
                        !FormID = FormID
                        !Control = MyControl.Name
                        !ControlType = GetControlType(MyControl)
                        !ControlText = GetControlText(MyControl)
                    End If
                    .Update
                    GetControlID = !ControlID
                    .Close
                End If
            End With
End Function


