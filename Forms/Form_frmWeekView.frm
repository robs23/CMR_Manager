VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmWeekView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private transportOrders As New Collection 'collection of transportOrders
Private restrictions As New Collection 'collection of restrictions
Private startFrom As Integer 'weekday that will be placed in first column.


Private Sub Box1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim tranOr As clsTransportOrder
For Each tranOr In transportOrders
    If tranOr.highlighted Then tranOr.highlighted = False
Next tranOr

Dim res As clsRestriction
For Each res In restrictions
    If res.highlighted Then res.highlighted = False
Next res
End Sub

Private Sub Box2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tranOr As clsTransportOrder
For Each tranOr In transportOrders
    If tranOr.highlighted Then tranOr.highlighted = False
Next tranOr

Dim res As clsRestriction
For Each res In restrictions
    If res.highlighted Then res.highlighted = False
Next res
End Sub



Private Sub Box3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tranOr As clsTransportOrder
For Each tranOr In transportOrders
    If tranOr.highlighted Then tranOr.highlighted = False
Next tranOr

Dim res As clsRestriction
For Each res In restrictions
    If res.highlighted Then res.highlighted = False
Next res
End Sub

Private Sub Box4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tranOr As clsTransportOrder
For Each tranOr In transportOrders
    If tranOr.highlighted Then tranOr.highlighted = False
Next tranOr

Dim res As clsRestriction
For Each res In restrictions
    If res.highlighted Then res.highlighted = False
Next res
End Sub

Private Sub Box5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tranOr As clsTransportOrder
For Each tranOr In transportOrders
    If tranOr.highlighted Then tranOr.highlighted = False
Next tranOr

Dim res As clsRestriction
For Each res In restrictions
    If res.highlighted Then res.highlighted = False
Next res

End Sub

Private Sub Box6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tranOr As clsTransportOrder
For Each tranOr In transportOrders
    If tranOr.highlighted Then tranOr.highlighted = False
Next tranOr

Dim res As clsRestriction
For Each res In restrictions
    If res.highlighted Then res.highlighted = False
Next res

End Sub

Private Sub btnDelete_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("TRANSPORT_DELETE"), whoIsLogged) Then
    Call killForm("frmNotify")
    DeleteTransport
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnFinish_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("TRANSPORT_EDIT"), whoIsLogged) Then
    Call killForm("frmNotify")
    finishTransport
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnRefresh_Click()
busy True
Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value), startFrom)
busy False
End Sub

Private Sub btnShift_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("TRANSPORT_EDIT"), whoIsLogged) Then
    Call killForm("frmNotify")
    shiftTransport
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnThisWeek_Click()
busy True
Call setCombobox(Me.cmbYear, CStr(year(Date)))
Call setCombobox(Me.cmbWeek, IsoWeekNumber(Date))
Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value))
busy False
End Sub

Private Sub cmbWeek_AfterUpdate()
busy True
Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value))
busy False
End Sub

Private Sub cmbYear_AfterUpdate()
busy True
Call populateListboxSelected(Me, Me.cmbWeek, getArray(1, weeksInYear(CLng(Me.cmbYear.value))), 1)
Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value))
busy False
'Call setCombobox(Me.cmbWeek, 1)
End Sub


Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tranOr As clsTransportOrder
For Each tranOr In transportOrders
    If tranOr.highlighted Then tranOr.highlighted = False
Next tranOr
End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
Call populateListboxSelected(Me, Me.cmbYear, getArray(2015, 2030), year(Date))
'Call setCombobox(Me.cmbYear, Year(Date))
Call populateListboxSelected(Me, Me.cmbWeek, getArray(1, weeksInYear(CLng(Me.cmbYear.value))), IsoWeekNumber(Date))
'Call setCombobox(Me.cmbWeek, 1)
Call setWeek(IsoWeekNumber(Date), year(Date))
Me.txtTipper.visible = False
busy False
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

Sub setWeek(week As Integer, y As Long, Optional firstDay As Variant)
Dim i As Integer
Dim ind As Integer

If IsMissing(firstDay) Then
    firstDay = 1
End If

Select Case firstDay
Case 1
    Me.txtWeekB.visible = False
    Me.weekSplitter.visible = False
    Me.txtWeekA.value = "tydzień " & week
    Me.txtWeekA.Width = Me.BoxWeek.Width
Case Else
    Me.txtWeekB.visible = True
    Me.weekSplitter.visible = True
    Me.txtWeekA.value = "tydzień " & week
    If week > 50 Then
        If week = weeksInYear(y) Then
            Me.txtWeekB.value = "tydzień 1"
        Else
            Me.txtWeekB.value = "tydzień " & week + 1
        End If
    Else
        Me.txtWeekB.value = "tydzień " & week + 1
    End If
    Me.txtWeekA.Width = (6 - firstDay + 1) * Me.Box1.Width
    Me.txtWeekB.Left = Me.txtWeekA.Left + Me.txtWeekA.Width + 100
    Me.weekSplitter.Left = Me.BoxWeek.Left + Me.txtWeekA.Width - ((7 - firstDay) * 10)
    Me.txtWeekB.Width = Me.BoxWeek.Width - (Me.txtWeekA.Left + Me.txtWeekA.Width + 200)
    
End Select

For i = 1 To 6
    ind = i + (firstDay - 1)
    If ind > 6 Then
        If ind / 7 >= 1 And i = 1 Then week = week + Int(ind / 7)
        If ind Mod 7 = 0 And i <> 1 Then week = week + Int(ind / 7)
        ind = ind - firstDay - (6 - firstDay)
    End If
    Me.Controls("txtDay" & i).value = WeekdayName(CLng(ind))
    Me.Controls("txtDate" & i).value = DateAdd("d", ind - 1, Week2Date(CLng(week), y))
    startFrom = CInt(firstDay)
Next i
distributeOrders
adjustToMax
distributeRestrictions
showHolidays
showWorkload week, y
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

Private Sub Form_Unload(Cancel As Integer)
Dim tranOr As clsTransportOrder
Dim res As clsRestriction

For Each tranOr In transportOrders
    transportOrders.Remove CStr(tranOr.ID)
    Set tranOr = Nothing
Next tranOr

For Each res In restrictions
    restrictions.Remove CStr(res.Name)
    Set res = Nothing
Next res
End Sub

Private Sub nextDay_Click()
busy True
If startFrom = 6 Then
    Call setCombobox(Me.cmbWeek, CInt(Me.cmbWeek.value) + 1)
    Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value), 1)
Else
    Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value), startFrom + 1)
End If
busy False
End Sub

Private Sub nextWeek_Click()
busy True
If CInt(Me.cmbWeek.value) > 51 Then
    If CInt(Me.cmbWeek.value) = weeksInYear(CLng(Me.cmbYear.value)) Then
        Call setCombobox(Me.cmbYear, CStr(CInt(Me.cmbYear.value) + 1))
    End If
End If
Call setCombobox(Me.cmbWeek, CStr(CInt(Me.cmbWeek.value) + 1))
Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value))
busy False
End Sub

Private Sub prevDay_Click()
busy True
If startFrom = 1 Then
    Call setCombobox(Me.cmbWeek, CInt(Me.cmbWeek.value) - 1)
    Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value), 6)
Else
    Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value), startFrom - 1)
End If
busy False
End Sub

Private Sub prevWeek_Click()
busy True
If CInt(Me.cmbWeek.value) = 1 Then
    Call setCombobox(Me.cmbYear, CStr(CInt(Me.cmbYear.value) - 1))
    Call setCombobox(Me.cmbWeek, CStr(weeksInYear(Me.cmbYear.value)))
Else
    Call setCombobox(Me.cmbWeek, CStr(CInt(Me.cmbWeek.value) - 1))
End If

Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value))
busy False
End Sub

Private Sub distributeRestrictions()
Dim res As clsRestriction
Dim i As Integer
Dim sDate As Date
Dim eDate As Date
Dim rs As ADODB.Recordset
Dim ind As String
Dim max As Integer

For i = 1 To 6
    Me.Controls("res" & i).visible = False
Next i

For Each res In restrictions
    restrictions.Remove (CStr(res.Name))
Next res

sDate = CDate(Me.Controls("txtDate1"))
eDate = CDate(Me.Controls("txtDate6"))

Set rs = newRecordset("SELECT * FROM tbCalendarRestrictions WHERE calDate BETWEEN '" & sDate & "' AND '" & eDate & "'")
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set res = New clsRestriction
        With res
            .ID = rs.fields("calendarRestrictionsId")
            .resDate = rs.fields("calDate")
            .Name = Trim(rs.fields("name"))
            .Length = rs.fields("slotsTaken")
            .description = Trim(rs.fields("description"))
            restrictions.Add res, res.Name
        End With
        rs.MoveNext
    Loop
End If

rs.Close
Set rs = Nothing

i = 1
Do Until sDate > eDate
    ind = "res_" & Day(sDate) & "_" & Month(sDate) & "_" & year(sDate)
    If inCollection(ind, restrictions) Then
        max = trucksOnDate(sDate)
        Me.Controls("res" & i).visible = True
        Me.Controls("res" & i).value = restrictions(ind).description
        Me.Controls("res" & i).TOP = Me.Controls("mn" & max + 1).TOP
        Me.Controls("res" & i).Height = (restrictions(ind).Length * (Me.Controls("mn2").TOP - Me.Controls("mn1").TOP)) - (Me.Controls("mn2").TOP - Me.Controls("mn1").TOP - Me.Controls("mn1").Height)
        restrictions(ind).setTextBox Me.Controls("res" & i)
        restrictions(ind).setTipper Me.txtTipper
    End If
    i = i + 1
    sDate = DateAdd("d", 1, sDate)
Loop

End Sub

Private Sub distributeOrders()
Dim i As Integer
Dim theDate As Date
Dim n As Integer
Dim rs As ADODB.Recordset
Dim ctl As Access.Control
Set db = CurrentDb
Dim theDay As String
Dim tranOrder As clsTransportOrder

For Each tranOrder In transportOrders
    transportOrders.Remove (CStr(tranOrder.ID))
    Set tranOrder = Nothing
Next tranOrder


For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If Left(ctl.Name, 3) <> "txt" And Left(ctl.Name, 3) <> "res" Then
            ctl.value = ""
            ctl.visible = False
        End If
    End If
Next ctl

For i = 1 To 6
    Select Case i
    Case 1
        theDay = "mn"
    Case 2
        theDay = "tu"
    Case 3
        theDay = "we"
    Case 4
        theDay = "th"
    Case 5
        theDay = "fr"
    Case 6
        theDay = "sa"
    End Select
    
    theDate = CDate(Me.Controls("txtDate" & i))
    Set rs = newRecordset("SELECT transportId, transportNumber, transportStatus,truckNumbers , CASE WHEN c.companyId IS NULL THEN NULL ELSE cd.companyName END as Carrier FROM tbTransport t LEFT JOIN tbCarriers c ON c.carrierId=t.carrierId LEFT JOIN tbCompanyDetails cd ON cd.companyId=c.companyId WHERE transportDate = '" & theDate & "'")
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        rs.MoveFirst
        n = 1
        Do Until rs.EOF
            Me.Controls(theDay & n).value = rs.fields("transportNumber")
            If rs.fields("transportStatus") = 1 Then
                Me.Controls(theDay & n).BorderColor = RGB(200, 200, 200)
                Me.Controls(theDay & n).BackColor = RGB(200, 200, 200)
                Set tranOrder = factory.CreateTransportOrder(rs.fields("transportNumber"), rs.fields("transportId"), False, Me.Controls(theDay & n), Me.txtTipper)
            ElseIf rs.fields("transportStatus") = 2 Then
                Me.Controls(theDay & n).BorderColor = RGB(96, 222, 53)
                Me.Controls(theDay & n).BackColor = RGB(96, 222, 53)
                Set tranOrder = factory.CreateTransportOrder(rs.fields("transportNumber"), rs.fields("transportId"), True, Me.Controls(theDay & n), Me.txtTipper)
            End If
            If Not IsNull(rs.fields("truckNumbers")) Then
                tranOrder.TruckNumbers = rs.fields("truckNumbers")
            End If
            If Not IsNull(rs.fields("Carrier")) Then
                tranOrder.carrierString = rs.fields("Carrier")
            End If
            transportOrders.Add tranOrder, CStr(tranOrder.ID)
            
            
            Me.Controls(theDay & n).visible = True
            Me.Controls(theDay & n).BackStyle = 1
            Me.Controls(theDay & n).ForeColor = vbBlack
            Me.Controls(theDay & n).FontWeight = 700
            rs.MoveNext
            n = n + 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
Next i

Set db = Nothing
End Sub

Private Function getHolidays(dDate As Date) As Variant
Dim rs As ADODB.Recordset
Dim str As Variant


Set rs = newRecordset("SELECT tbHolidays.holidayDate, tbTransportLane.transportLaneInitials FROM tbHolidays LEFT JOIN tbTransportLane ON tbHolidays.laneId = tbTransportLane.transportLaneId WHERE holidayDate = '" & dDate & "'")
str = Null
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If IsNull(str) Then
            str = rs.fields("transportLaneInitials")
        Else
            str = str & ", " & rs.fields("transportLaneInitials")
        End If
        rs.MoveNext
    Loop
End If

getHolidays = str

rs.Close

Set rs = Nothing
Set db = Nothing

End Function

Private Sub showHolidays()
Dim i As Integer
Dim var As Variant

For i = 1 To 6
    var = getHolidays(CDate(Me.Controls("txtDate" & i).value))
    If IsNull(var) Then
        Me.Controls("imgHol" & i).visible = False
    Else
        Me.Controls("imgHol" & i).visible = True
        Me.Controls("imgHol" & i).ControlTipText = "Święto w: " & var
    End If
Next i

End Sub

Private Sub showWorkload(week As Integer, year As Long)
Dim sql As String
Dim rs As ADODB.Recordset
Dim str As Variant

On Error GoTo err_trap

sql = "SELECT SUM(tots.SlotsBooked) as  SlotsTaken, (SELECT CASE WHEN SUM(cr.slotsTaken) IS NULL THEN 0 ELSE SUM(cr.slotsTaken) END FROM tbCalendarRestrictions cr WHERE DATEPART(ISO_WEEK, cr.calDate)=" & week & " AND YEAR(cr.calDate)=" & year & ") As Restrictions, (SELECT TOP(1) newValue from tbSettingChanges s WHERE s.settingId=1 ORDER BY s.modificationDate DESC)*5 as SlotsPossible FROM " _
    & "(SELECT trans.transportDate, trans.SlotsBooked, CASE WHEN cr.slotsTaken IS NULL THEN 0 ELSE cr.slotsTaken END AS slotsTaken FROM " _
    & "(SELECT t.transportDate, COUNT(t.transportDate) AS SlotsBooked " _
    & "FROM tbTransport t " _
    & "WHERE DatePart(ISO_WEEK, t.transportDate) = " & week & " And Year(t.transportDate) = " & year _
    & "GROUP BY t.transportDate) trans LEFT JOIN tbCalendarRestrictions cr ON cr.calDate=trans.transportDate) tots"

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Me.lblWorkload.Caption = rs.fields("SlotsTaken").value & " (" & Round(((rs.fields("SlotsTaken").value + rs.fields("Restrictions").value) / rs.fields("SlotsPossible").value) * 100, 1) & "%)"
End If

exit_here:
rs.Close
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ShowWorkload. " & Err.description, vbOKOnly + vbCritical
Resume exit_here

End Sub

Private Sub finishTransport()
Dim t As clsTransportOrder
Dim rs As ADODB.Recordset
Dim res As clsRestriction

i = 0

For Each t In transportOrders
    If t.selected Then
        i = i + 1
        Exit For
    End If
Next t

If i > 0 Then

    Call newNotify("Wybrana akcja w toku.. Proszę czekać..")
    
    For Each t In transportOrders
        If t.selected Then
            Set rs = newRecordset("SELECT u.userName + ' ' + u.userSurname as theUser FROM tbTransport t LEFT JOIN tbUsers u ON t.isBeingEditedBy = u.userId WHERE t.transportNumber = '" & t.number & "' AND t.isBeingEditedBy IS NOT NULL")
            Set rs.ActiveConnection = Nothing
            If rs.EOF Then
                'mark as finished
                adoConn.Execute "UPDATE tbTransport SET transportStatus = 2, lastModifiedBy = " & whoIsLogged & ", lastModifiedOn = '" & Now & "' WHERE transportNumber = '" & t.number & "'"
            Else
                rs.MoveFirst
                'is edited by "user", skip this one
                MsgBox "Zlecenie " & t.number & " jest w tej chwili edytowane przez użytkownika " & rs.fields("theUser") & ". Z tego powodu zostanie ono pominięte", vbOKOnly + vbInformation, "Zlecenie w edycji"
            End If
            rs.Close
            Set rs = Nothing
        End If
    Next t
    
    For Each res In restrictions
        If res.selected Then
            res.selectMe False
        End If
    Next res
    
    Call killForm("frmNotify")
    
    Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value), startFrom)
Else
    MsgBox "Żadne zlecenie nie jest zaznaczone. Aby zmienić status zlecenia, najpierw zaznacz je kliknięciem", vbOKOnly + vbInformation, "Brak zaznaczenia"
End If
Set db = Nothing
End Sub


Private Sub DeleteTransport()
Dim t As clsTransportOrder
Dim db As DAO.Database
Dim i As Integer
Dim iRes As Integer
Dim rs As ADODB.Recordset
Dim var As Variant
Dim detailId As Variant
Dim r As clsRestriction

Set db = CurrentDb

i = 0

For Each t In transportOrders
    If t.selected Then
        i = i + 1
        Exit For
    End If
Next t

iRes = 0
For Each r In restrictions
    If r.selected Then
        iRes = iRes + 1
    End If
Next r

If i > 0 Or iRes > 0 Then

    Call newNotify("Wybrana akcja w toku.. Proszę czekać..")
    
    res = MsgBox("Wszystkie zaznaczone zlecenia zostaną usunięte. Tego kroku nie będzie można cofnąć. Czy chcesz kontynuować?", vbYesNo + vbExclamation, "Potwierdź usunięcie")
    If res = vbYes Then
        For Each t In transportOrders
            If t.selected Then
                Set rs = newRecordset("SELECT u.userName + ' ' + u.userSurname as theUser FROM tbTransport t LEFT JOIN tbUsers u ON t.isBeingEditedBy = u.userId WHERE t.transportNumber = '" & t.number & "' AND t.isBeingEditedBy IS NOT NULL")
                Set rs.ActiveConnection = Nothing
                If rs.EOF Then
                    'check if dependent CMRs are editable
                    rs.Close
                    Set rs = Nothing
                    Set rs = newRecordset("SELECT u.userName + ' ' + u.userSurname as theUser FROM tbTransport t LEFT JOIN tbCmr c ON c.transportId = t.transportId LEFT JOIN tbUsers u ON c.isBeingEditedBy = u.userId WHERE t.transportNumber = '" & t.number & "' AND c.isBeingEditedBy IS NOT NULL")
                    Set rs.ActiveConnection = Nothing
                    If rs.EOF Then
                        'delete
                        detailId = adoDLookup("detailId", "tbCmr", "transportId=" & t.ID)
                        updateConnection
                        If Not IsNull(detailId) Then
                            adoConn.Execute "DELETE FROM tbDeliveryDetail WHERE cmrDetailId=" & detailId
                        End If
                        adoConn.Execute "DELETE FROM tbCmr WHERE transportId = " & t.ID
                        adoConn.Execute "DELETE FROM tbTransport WHERE transportId = " & t.ID
                        t.selectMe False
                        transportOrders.Remove CStr(t.ID)
                    Else
                        rs.MoveFirst
                        'is edited by "user", skip this one
                        MsgBox "Jeden z dokumentów CMR powiązanych ze zleceniem transportowym " & t.number & " jest w tej chwili edytowany przez użytkownika " & rs.fields("theUser") & ". Z tego powodu zlecenie transportowe " & t.number & " nie zostanie usunięte", vbOKOnly + vbInformation, "Zlecenie w edycji"
                        End If
                Else
                    rs.MoveFirst
                    'is edited by "user", skip this one
                    MsgBox "Zlecenie " & t.number & " jest w tej chwili edytowane przez użytkownika " & rs.fields("theUser") & ". Z tego powodu zostanie ono pominięte", vbOKOnly + vbInformation, "Zlecenie w edycji"
                End If
            End If
        Next t
        If iRes > 0 Then
            For Each r In restrictions
                If r.selected Then
                    Set rs = newRecordset("SELECT u.userName + ' ' + u.userSurname as theUser FROM tbCalendarRestrictions r LEFT JOIN tbUsers u ON r.isBeingEditedBy = u.userId WHERE r.calendarRestrictionsId = " & r.ID & " AND r.isBeingEditedBy IS NOT NULL")
                    Set rs.ActiveConnection = Nothing
                    If rs.EOF Then
                        'check if dependent CMRs are editable
                        rs.Close
                        Set rs = Nothing
                        updateConnection
                        adoConn.Execute "DELETE FROM tbCalendarRestrictions WHERE calendarRestrictionsId = " & r.ID
                        r.selectMe False
                        restrictions.Remove CStr(r.Name)
                    Else
                        rs.MoveFirst
                        'is edited by "user", skip this one
                        MsgBox "Ograniczenie z dnia " & res.resDate & " jest w tej chwili edytowane przez użytkownika " & rs.fields("theUser") & ". Z tego powodu zostanie ono pominięte", vbOKOnly + vbInformation, "Zlecenie w edycji"
                    End If
                End If
            Next r
        End If
    End If
    
    Call killForm("frmNotify")
    
    Call setWeek(CInt(Me.cmbWeek.value), CLng(Me.cmbYear.value), startFrom)
Else
    MsgBox "Żadne zlecenie nie jest zaznaczone. Aby usunąć zlecenie, najpierw zaznacz je kliknięciem", vbOKOnly + vbInformation, "Brak zaznaczenia"
End If
Set db = Nothing
End Sub

Private Sub shiftTransport()
Dim t As clsTransportOrder
Dim r As clsRestriction

Dim ords As String
Dim i As Integer

i = 0

For Each t In transportOrders
    If t.selected Then
        i = i + 1
        ords = ords & t.ID & ","
    End If
Next t

For Each r In restrictions
    If r.selected Then
        r.selectMe False
    End If
Next r


If i > 0 Then
    ords = Left(ords, Len(ords) - 1)
    launchForm "frmDate", ords
Else
    MsgBox "Żadne zlecenie nie jest zaznaczone. Aby zmienić datę realizacji jakiegoś zlecenia, najpierw zaznacz je kliknięciem", vbOKOnly + vbInformation, "Brak zaznaczenia"
End If
End Sub

Private Sub adjustToMax()
Dim rs As ADODB.Recordset
Dim max As Integer
Dim displayed As Integer 'how many slots are displayed
Dim eDate As Date

eDate = Me.txtDate6

Set rs = newRecordset("SELECT TOP(1) newValue FROM tbSettingChanges WHERE settingId=1 AND modificationDate < '" & eDate & "' ORDER BY modificationDate DESC")
Set rs.ActiveConnection = Nothing

displayed = 16
max = rs.fields("newValue")

Me.lnMax.visible = False
Me.lMax.visible = False
    
rs.Close
Set rs = Nothing

'--------------------------------red line max-------------------------------------

'If max < displayed Then
'    Me.lnMax.visible = True
'    Me.lMax.visible = True
'    Me.lnMax.TOP = Me.Controls("mn" & max).TOP + Me.Controls("mn" & max).Height + 100
'    Me.lMax.TOP = Me.lnMax.TOP + Me.lnMax.Height + 50
'Else
'    Me.lnMax.visible = False
'    Me.lMax.visible = False
'End If
'


'--------------------------------variable box height max-------------------------------------
For i = displayed To 1 Step -1
    If i > max Then
        Me.Controls("lp" & i).visible = False
    Else
        Me.Controls("lp" & i).visible = True
    End If
Next i


For i = 1 To 6
    Me.Controls("Box" & i).Height = (Me.Controls("mn" & max).TOP + Me.Controls("mn" & max).Height + 50) - Me.Controls("Box" & i).TOP
Next i

End Sub

Private Sub busy(state As Boolean)

If state Then
    DoCmd.Hourglass True
    prompter.visible = True
Else
    DoCmd.Hourglass False
    prompter.visible = False
    toFront Me
End If

DoEvents

End Sub

