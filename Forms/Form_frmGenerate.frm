VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnCreate_Click()
generate
End Sub

Private Sub cmbCustomer_AfterUpdate()
Dim sql As String

Me.cmbWarehouse = Null

If Not IsNull(Me.cmbCustomer) Then
    sql = "SELECT sh.shipToId, sh.shipToString + ' ' + cd.companyName as name, cd.companyCity, cd.companyCountry " _
        & "FROM tbShipTo sh LEFT JOIN tbCompanyDetails cd ON cd.companyId = sh.companyId " _
        & "WHERE cd.companyId Is Not Null And cd.isActive <> 0 And sh.SoldTo = " & Me.cmbCustomer
    populateListboxFromSQL sql, Me.cmbWarehouse
End If
End Sub

Private Sub cmbWarehouse_AfterUpdate()
Me.cmbCarrier = Null
populateCmb Me.cmbWarehouse.Column(3)
End Sub


Sub populateCmb(Optional country As Variant)
Dim sql As String

If Not IsMissing(country) Then
    If IsNull(country) Then
        sql = "SELECT car.carrierId, cd.companyName, cd.companyAddress FROM tbCarriers car LEFT JOIN tbCompanyDetails cd ON car.companyId = cd.companyId WHERE cd.companyId IS NOT NULL"
    Else
       sql = "SELECT carrierId, companyName, companyAddress FROM " _
            & "(SELECT TOP 1000 car.carrierId, cd.companyName, cd.companyAddress, " _
            & "CASE WHEN car.carrierId IN (SELECT DISTINCT custSh.PrimaryCarrier FROM tbShipTo custSh LEFT JOIN tbCompanyDetails custCd ON custSh.companyId = custCd.companyId WHERE custCd.companyCountry = '" & country & "') THEN 1 ELSE 0 END as PrimaryCarrier, " _
            & "CASE WHEN car.carrierId IN (SELECT DISTINCT custSh.supportiveCarrier FROM tbShipTo custSh LEFT JOIN tbCompanyDetails custCd ON custSh.companyId = custCd.companyId WHERE custCd.companyCountry = '" & country & "') THEN 1 ELSE 0 END as SupportiveCarrier " _
            & "FROM tbCarriers car LEFT JOIN tbCompanyDetails cd ON car.companyId = cd.companyId " _
            & "WHERE cd.companyId Is Not Null ORDER BY PrimaryCarrier DESC, SupportiveCarrier DESC) t"
    End If
Else
    sql = "SELECT car.carrierId, cd.companyName, cd.companyAddress FROM tbCarriers car LEFT JOIN tbCompanyDetails cd ON car.companyId = cd.companyId WHERE cd.companyId IS NOT NULL"
End If
populateListboxFromSQL sql, Me.cmbCarrier

End Sub

Private Sub cmbWeek_AfterUpdate()
setWeek CInt(Me.cmbWeek), CLng(Me.cmbYear)
End Sub

Private Sub cmbYear_AfterUpdate()
setWeek CInt(Me.cmbWeek), CLng(Me.cmbYear)
End Sub

Private Sub Form_Load()
Dim sql As String
Call killForm("frmNotify")
Call populateListboxSelected(Me, Me.cmbYear, getArray(2015, 2030), year(Date))
'Call setCombobox(Me.cmbYear, Year(Date))
Call populateListboxSelected(Me, Me.cmbWeek, getArray(1, weeksInYear(CLng(Me.cmbYear.value))), IsoWeekNumber(Date))
setWeek CInt(Me.cmbWeek), CLng(Me.cmbYear)
deployPallets
sql = "SELECT s.soldToId, s.soldToString + ' ' + cd.companyName as name, cd.companyCity, cd.companyCountry " _
    & "FROM tbSoldTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId = s.companyId " _
    & "WHERE cd.companyId Is Not Null And cd.isActive <> 0"
populateListboxFromSQL sql, Me.cmbCustomer
'Call setCombobox(Me.cmbWeek, 1)
'Call setWeek(IsoWeekNumber(Date), Year(Date))
End Sub


Private Function getArray(first As Long, last As Long) As String()
Dim nr As Long
Dim i As Long
Dim values() As String
If last >= first Then
    nr = last - first + 1
    ReDim values(nr - 1) As String
    For i = 1 To nr
        values(i - 1) = CStr(first + i - 1)
    Next i
    getArray = values
End If
End Function

Private Function weeksInYear(y As Long) As Integer
Dim w As Integer
Dim m As Integer
w = CInt(IsoWeekNumber(DateSerial(y, 12, 31)))
If w < 50 Then
    m = 1
    Do Until w > 50
        w = CInt(IsoWeekNumber(DateSerial(y, 12, 31 - m)))
        m = m + 1
    Loop
End If
weeksInYear = w
End Function

Sub setWeek(week As Integer, y As Long)
Dim i As Integer
Dim ind As Integer
Dim firstDate As Date

firstDate = Week2Date(CLng(week), y)

Me.txtMon = firstDate
Me.txtTue = DateAdd("d", 1, firstDate)
Me.txtWed = DateAdd("d", 2, firstDate)
Me.txtThu = DateAdd("d", 3, firstDate)
Me.txtFri = DateAdd("d", 4, firstDate)
Me.txtSat = DateAdd("d", 5, firstDate)
clearForm
'showHolidays
End Sub

Sub clearForm()
Me.txtMonA = ""
Me.txtTueA = ""
Me.txtWedA = ""
Me.txtThuA = ""
Me.txtFriA = ""
Me.txtSatA = ""
Me.txtMonNotes = ""
Me.txtTueNotes = ""
Me.txtWedNotes = ""
Me.txtThuNotes = ""
Me.txtFriNotes = ""
Me.txtSatNotes = ""
End Sub

Private Sub nextWeek_Click()
If CInt(Me.cmbWeek.value) > 51 Then
    If CInt(Me.cmbWeek.value) = weeksInYear(CLng(Me.cmbYear.value)) Then
        Call setCombobox(Me.cmbYear, CStr(CInt(Me.cmbYear.value) + 1))
    End If
End If
Call setCombobox(Me.cmbWeek, CStr(CInt(Me.cmbWeek.value) + 1))
Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value))
End Sub


Sub setCombobox(ctl As ComboBox, value As String)
Dim ind As Long
Dim i As Long

ctl.SetFocus
For i = 0 To ctl.ListCount
    If ctl.ItemData(i) = value Then
        ind = i
        Exit For
    End If
Next i
ctl.SetFocus
ctl.value = ctl.ItemData(ind)
End Sub

Private Sub prevWeek_Click()
If CInt(Me.cmbWeek.value) = 1 Then
    Call setCombobox(Me.cmbYear, CStr(CInt(Me.cmbYear.value) - 1))
    Call setCombobox(Me.cmbWeek, CStr(weeksInYear(Me.cmbYear.value)))
Else
    Call setCombobox(Me.cmbWeek, CStr(CInt(Me.cmbWeek.value) - 1))
End If

Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value))
End Sub

Private Sub deployPallets()
Me.txtMonP = 33
Me.txtTueP = 33
Me.txtWedP = 33
Me.txtThuP = 33
Me.txtFriP = 33
Me.txtSatP = 33
End Sub

Private Sub generate()
Dim db As DAO.Database
Set db = CurrentDb
Dim sql As String
Dim tb As TextBox
Dim truck As TextBox
Dim amount As TextBox
Dim note As TextBox
Dim ttruck As Integer
Dim tNow As Integer
Dim tMax As Integer
Dim rs As ADODB.Recordset
Dim mStr As String
Dim counter As Integer
Dim lasting As Integer
Dim thisCust As Integer
Dim currentCust As Integer
Dim newName As String
Dim x As Integer
Dim isError As Boolean
Dim dday As String
Dim mmonth As String
Dim aStr As String
Dim transportId As Long
Dim locations As String
Dim sDate As String
Dim highest As Integer
Dim sId As String
Dim v() As String

Dim i As Integer

newNotify "Tworzę zlecenia transportowe.. Proszę czekać.."

If validate Then
    'tMax = CInt(adoDLookup("settingValue", "tbSettings", "settingName='MAX_TRUCKS_DAILY'"))
    mStr = "Pomyślnie utworzono zlecenia: " & vbNewLine
    locations = getLocations()
    For i = 1 To 6
        Select Case i
        Case 1
            Set tb = Me.txtMon
            Set truck = Me.txtMonA
            Set amount = Me.txtMonP
            Set note = Me.txtMonNotes
        Case 2
            Set tb = Me.txtTue
            Set truck = Me.txtTueA
            
            Set amount = Me.txtTueP
            Set note = Me.txtTueNotes
        Case 3
            Set tb = Me.txtWed
            Set truck = Me.txtWedA
            Set amount = Me.txtWedP
            Set note = Me.txtWedNotes
        Case 4
            Set tb = Me.txtThu
            Set truck = Me.txtThuA
            Set amount = Me.txtThuP
            Set note = Me.txtThuNotes
        Case 5
            Set tb = Me.txtFri
            Set truck = Me.txtFriA
            Set amount = Me.txtFriP
            Set note = Me.txtFriNotes
        Case 6
            Set tb = Me.txtSat
            Set truck = Me.txtSatA
            Set amount = Me.txtSatP
            Set note = Me.txtSatNotes
        End Select
        
        If Len(truck.value) = 0 Then
            ttruck = 0
        Else
            ttruck = CInt(truck.value)
        End If
        
        counter = 0
        currentCust = 0
        tNow = 0
        If ttruck > 0 Then
            tMax = getMaxSlot(tb.value)
            If Day(tb.value) < 10 Then
                dday = "0" & CStr(Day(tb.value))
            Else
                dday = CStr(Day(tb.value))
            End If
            If Month(tb.value) < 10 Then
                mmonth = "0" & CStr(Month(tb.value))
            Else
                mmonth = CStr(Month(tb.value))
            End If
            Set rs = newRecordset("SELECT * FROM tbTransport WHERE transportDate = '" & tb.value & "'")
            Set rs.ActiveConnection = Nothing
            If Not rs.EOF Then
                'rs.MoveLast
                rs.MoveFirst
                tNow = rs.RecordCount
                tNow = tNow + restrictionsOnDate(tb.value)
                highest = 0
                
                sDate = year(tb.value) & mmonth & dday
                Do Until rs.EOF
                    If InStr(1, rs.fields("transportNumber"), locations, vbTextCompare) > 0 And InStr(1, rs.fields("transportNumber"), sDate, vbTextCompare) > 0 Then
                        v = Split(rs.fields("transportNumber"), "-", , vbTextCompare)
                        If UBound(v) > 2 Then
                            If IsNumeric(v(3)) Then
                                If CInt(v(3)) > highest Then highest = CInt(v(3))
                            End If
                        End If
                    End If
                    rs.MoveNext
                Loop
            Else
                tNow = 0 + restrictionsOnDate(tb.value) 'currently we have 0 trucks on the day, but there are some restrictions
            End If
            rs.Close
            Set rs = Nothing
            If tNow < tMax Then
                'we can add at least 1 truck
                lasting = ttruck
                
                updateConnection
                Do Until tNow = tMax Or lasting = 0
                    If highest + counter + 1 < 10 Then
                        sId = "0" & CStr(highest + counter + 1)
                    Else
                        sId = CStr(highest + counter + 1)
                    End If
                    newName = CStr(year(tb.value)) & mmonth & dday & "-M024-" & Trim(locations) & "-" & sId
                    sql = "INSERT INTO tbTransport (transportNumber, transportDate, transportStatus, carrierId, createdBy, initDate, Notes) "
                    sql = sql & "VALUES('" & newName & "','" & tb.value & "', 1," & Me.cmbCarrier & "," & whoIsLogged & ",'" & tb.value & "','" & note.value & "')"
                    Set rs = adoConn.Execute(sql & ";SELECT SCOPE_IDENTITY()")
'                    sql = sql & ";SELECT SCOPE_IDENTITY()"
'                    Set rs = New ADODB.Recordset
'                    rs.Open sql, adoConn, adOpenKeyset, adLockOptimistic
                    Set rs = rs.NextRecordset
                    transportId = rs.fields(0).value
                    rs.Close
                    Set rs = Nothing
                    saveCmr transportId, i
                    tNow = tNow + 1
                    lasting = lasting - 1
                    counter = counter + 1
                Loop
                mStr = mStr & vbNewLine & WeekdayName(Weekday(tb.value, vbMonday)) & ": " & counter & " aut"
                If counter < truck Then
                    If Len(aStr) = 0 Then
                        aStr = "Z powodu ograniczonej liczby slotów nie wszystkie zlecenia zostały utworzone!" & vbNewLine
                    End If
                    mStr = mStr & " (" & ttruck - counter & " mniej)"
                End If
            Else
                mStr = mStr & vbNewLine & WeekdayName(Weekday(tb.value, vbMonday)) & ": Wszystkie dostępne sloty na ten dzień (" & tMax & ") są już zajęte.. Dodano 0 zleceń"
            End If
        End If
    Next i
    MsgBox aStr & mStr
End If

killForm "frmNotify"

End Sub

Private Function validate() As Boolean
Dim bool As Boolean

bool = True

If IsNull(Me.cmbCustomer) Then
    MsgBox "Wybierz klienta z rozwijanej listy!", vbOKOnly + vbExclamation, "Klient nie został wybrany"
    bool = False
Else
    If IsNull(Me.cmbWarehouse) Then
        MsgBox "Wybierz magazyn z rozwijanej listy!", vbOKOnly + vbExclamation, "Magazyn nie został wybrany"
        bool = False
    Else
        If IsNull(Me.cmbCarrier) Then
            MsgBox "Wybierz przewoźnika z rozwijanej listy!", vbOKOnly + vbExclamation, "Przewoźnik nie został wybrany"
            bool = False
        End If
    End If
End If

validate = bool
End Function

Sub saveCmr(transportId As Long, dday As Integer)
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim i As Long
Dim transport As Long
Dim amount As TextBox
Dim sql As String

On Error GoTo err_trap

Select Case dday
    Case 1
        Set amount = Me.txtMonP
    Case 2
        Set amount = Me.txtTueP
    Case 3
        Set amount = Me.txtWedP
    Case 4
        Set amount = Me.txtThuP
    Case 5
        Set amount = Me.txtFriP
    Case 6
        Set amount = Me.txtSatP
End Select

updateConnection

'check if there's amount next to every truck number
If Not IsNull(Me.txtMonA) And IsNull(Me.txtMonP) Then Me.txtMonP = 33
If Not IsNull(Me.txtTueA) And IsNull(Me.txtTueP) Then Me.txtTueP = 33
If Not IsNull(Me.txtWedA) And IsNull(Me.txtWedP) Then Me.txtWedP = 33
If Not IsNull(Me.txtThuA) And IsNull(Me.txtThuP) Then Me.txtThuP = 33
If Not IsNull(Me.txtFriA) And IsNull(Me.txtFriP) Then Me.txtFriP = 33
If Not IsNull(Me.txtSatA) And IsNull(Me.txtSatP) Then Me.txtSatP = 33

sql = "INSERT INTO tbDeliveryDetail (soldToId, shipToId, numberPall) VALUES (" & Me.Controls("cmbCustomer").Column(0) & "," & Me.Controls("cmbWarehouse").Column(0) & "," & amount & ")"

Set rs = adoConn.Execute(sql & ";SELECT SCOPE_IDENTITY()")
Set rs = rs.NextRecordset
i = rs.fields(0).value

rs.Close

Set rs1 = newRecordset("tbCmr", True)
rs1.AddNew
rs1.fields("cmrCreated") = Now
rs1.fields("userId") = whoIsLogged
rs1.fields("cmrLastModified") = Now
rs1.fields("transportId") = transportId
rs1.fields("detailId") = i
i = rs1.fields("cmrId")
rs1.update
rs1.Close

exit_here:
Set rs = Nothing
Set rs1 = Nothing
Exit Sub

err_trap:
MsgBox "Error in saveCmr. " & Err.number & ", " & Err.description
Resume exit_here

End Sub

Private Function getLocations() As String
Dim rs As ADODB.Recordset
Dim sql As String

    sql = "SELECT cs.location, sh.shipToString " _
            & "FROM tbShipTo sh LEFT JOIN tbCustomerString cs ON cs.companyId=sh.companyId " _
            & "WHERE sh.shipToId IN (" & Me.cmbWarehouse.value & ")"
Set rs = newRecordset(sql)
If Not rs.EOF Then
    rs.MoveFirst
    If IsNull(rs.fields("location")) Then
        getLocations = Trim(rs.fields("shipToString"))
    Else
        getLocations = Trim(rs.fields("location"))
    End If
End If
rs.Close
Set rs = Nothing

End Function
