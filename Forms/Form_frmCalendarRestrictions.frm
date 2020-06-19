VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCalendarRestrictions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private descriptionLimit As Integer
Private restrictionId As Integer
Private mode As Integer '1-create,2-edit

Private Sub btnCreate_Click()
If verify Then
    If mode = 1 Then
        createRestrictions
    Else
        editRestriction
    End If
End If
End Sub

Private Sub bringRestriction()
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT * FROM tbCalendarRestrictions WHERE calendarRestrictionsId=" & restrictionId)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Me.txtDateTo = rs.fields("calDate")
    Me.txtHours = rs.fields("slotsTaken")
    Me.txtDescription = rs.fields("description")
End If

rs.Close
Set rs = Nothing
End Sub

Private Sub Form_Load()
killForm "frmNotify"
If Not IsNull(Me.openArgs) Then
    restrictionId = Me.openArgs
    mode = 2
    Me.txtDateFrom.visible = False
    Me.lblDateFrom.visible = False
    Me.lblDateTo.Caption = "Data ograniczenia"
    Me.btnCreate.Caption = "Zapisz zmiany"
    bringRestriction
Else
    mode = 1
    Me.txtDateFrom.visible = True
    Me.lblDateFrom.visible = True
    Me.lblDateTo.Caption = "Koniec ograniczenia"
    Me.btnCreate.Caption = "Utwórz"
End If
descriptionLimit = 150
Me.txtLeft = "Pozostało " & descriptionLimit & " znaków"
End Sub

Private Sub txtDescription_AfterUpdate()
If Len(Me.txtDescription) > descriptionLimit Then
    MsgBox "Pole może zawierać maksymalnie " & descriptionLimit & " znaków. Opis zostanie ograniczony do pierwszych " & descriptionLimit & " znaków..", vbOKOnly + vbInformation, "Ograniczenie"
    Me.txtDescription = Left(Me.txtDescription, CLng(descriptionLimit))
    Me.txtLeft = "Pozostało 0 znaków"
End If
End Sub

Private Sub txtDescription_Change()
If Len(Me.txtDescription.Text) > descriptionLimit Then
    Me.txtLeft = "Przekroczono limit znaków (" & descriptionLimit & ")"
Else
    Me.txtLeft = "Pozostało " & descriptionLimit - Len(Me.txtDescription.Text) & " znaków"
End If
End Sub

Private Function verify() As Boolean
Dim bool As Boolean

bool = False
If mode = 1 Then
    If DateDiff("d", Me.txtDateFrom, Me.txtDateTo) >= 0 Then
        If Me.txtDateFrom >= Date Then
            If IsNumeric(Me.txtHours) Then
                If CInt(Me.txtHours) > 0 Then
                    If Len(Me.txtDescription) > 0 Then
                        bool = True
                    Else
                        MsgBox "Podaj opis ograniczenia (powód)", vbExclamation + vbOKOnly, "Niepełne dane"
                    End If
                Else
                    MsgBox "Długość ograniczenia musi być większa niż 0!", vbExclamation + vbOKOnly, "Błąd typu danych"
                End If
            Else
                MsgBox "Długość ograniczenia musi być liczbą!", vbExclamation + vbOKOnly, "Błąd typu danych"
            End If
        Else
            MsgBox "Początek ograniczenia nie może znajdować się w przeszłości!", vbExclamation + vbOKOnly, "Błąd zakresu"
        End If
    Else
        MsgBox "Koniec ograniczenia nie może być wcześniej niż początek ograniczenia!", vbExclamation + vbOKOnly, "Błąd zakresu"
    End If
Else
    If Me.txtDateTo >= Date Then
        If IsNumeric(Me.txtHours) Then
            If CInt(Me.txtHours) > 0 Then
                If Len(Me.txtDescription) > 0 Then
                    bool = True
                Else
                    MsgBox "Podaj opis ograniczenia (powód)", vbExclamation + vbOKOnly, "Niepełne dane"
                End If
            Else
                MsgBox "Długość ograniczenia musi być większa niż 0!", vbExclamation + vbOKOnly, "Błąd typu danych"
            End If
        Else
            MsgBox "Długość ograniczenia musi być liczbą!", vbExclamation + vbOKOnly, "Błąd typu danych"
        End If
    Else
        MsgBox "Data ograniczenia nie może znajdować się w przeszłości!", vbExclamation + vbOKOnly, "Błąd zakresu"
    End If
End If
verify = bool

End Function

Private Sub editRestriction()
Dim sql As String
Dim spare As Integer
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT * FROM tbCalendarRestrictions WHERE calDate='" & Me.txtDateTo & "' AND calendarRestrictionsId<>" & restrictionId)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    MsgBox "Nie mogę przenieść bieżącego ograniczenia na dzień " & Me.txtDateTo & ", ponieważ w tym dniu istnieje już ograniczenie. Edytuj lub usuń ograniczenie o id=" & rs.fields("calendarRestrictionsId") & " aby kontynuować.", vbExclamation + vbOKOnly, "Istniejące ograniczenie"
    rs.Close
    Set rs = Nothing
Else
    rs.Close
    Set rs = Nothing
    spare = getMaxSlot(Me.txtDateTo) - trucksOnDate(Me.txtDateTo)
    If spare <= 0 Then
        MsgBox "W wybranym dniu nie ma już wolnych slotów. Wybierz inną datę", vbExclamation + vbOKOnly, "Brak wolnych slotów"
    Else
        If Me.txtHours > spare Then
            MsgBox "Liczba wolnych slotów w wybranym dniu jest mniejsza niż żądana długość ograniczenia. Ograniczenie zosanie automatycznie skrócone do liczby wolnych slotów", vbInformation + vbOKOnly, "Ograniczona liczba slotów"
            Me.txtHours = spare
        End If
        updateConnection
        
        sql = "UPDATE tbCalendarRestrictions SET calDate='" & Me.txtDateTo & "', name='res_" & Day(Me.txtDateTo) & "_" & Month(Me.txtDateTo) & "_" & year(Me.txtDateTo) & "',"
        sql = sql & "slotsTaken=" & Me.txtHours & ", description='" & Me.txtDescription & "' WHERE calendarRestrictionsId=" & restrictionId
        adoConn.Execute sql
        MsgBox "Edycja zakończona powodzeniem", vbInformation + vbOKOnly, "Powodzenie"
        killForm Me.Name
    End If
End If

End Sub

Private Sub createRestrictions()
Dim iSql As String
Dim curDate As Date
Dim rs As ADODB.Recordset
Dim nRes As clsRestriction
Dim ind As String
Dim ex As String
Dim restrictions As New Collection
Dim mStr As String
Dim uStr As String
Dim i As Integer
Dim resp As VbMsgBoxResult
Dim tMax As Integer
Dim tNow As Integer
Dim rMsg As String 'restricted day's string
Dim h As Integer 'number of hours/slots to book

newNotify "Tworzę ograniczenia dla wybranego okresu.. Proszę czekać.."
Set rs = newRecordset("SELECT * FROM tbCalendarRestrictions WHERE calDate BETWEEN '" & Me.txtDateFrom & "' AND '" & Me.txtDateTo & "'")
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set nRes = New clsRestriction
        nRes.resDate = rs.fields("calDate")
        nRes.ID = rs.fields("calendarRestrictionsId")
        nRes.Name = rs.fields("name")
        restrictions.Add nRes, nRes.Name
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

curDate = Me.txtDateFrom

updateConnection

Do Until curDate > Me.txtDateTo
    ind = "res_" & Day(curDate) & "_" & Month(curDate) & "_" & year(curDate)
    If Not inCollection(ind, restrictions) Then
        'add new
        tNow = trucksOnDate(curDate)
        tMax = getMaxSlot(curDate)
        If tNow < tMax Then
            If (tMax - tNow) < Me.txtHours Then
                ' we have to shorten the restriction
                h = tMax - tNow
                'save the date that you have just shortened
                rMsg = rMsg & curDate & ","
            Else
                h = Me.txtHours
            End If
            i = i + 1
            iSql = "INSERT INTO tbCalendarRestrictions (calDate, slotsTaken, name, description, dateAdded, addedBy) VALUES ("
            iSql = iSql & "'" & curDate & "'," & h & ",'" & ind & "','" & Me.txtDescription & "','" & Now & "'," & whoIsLogged & ")"
            adoConn.Execute iSql
        End If
    End If
    curDate = DateAdd("d", 1, curDate)
Loop

If restrictions.count > 0 Then
    resp = MsgBox("Dla " & restrictions.count & " dni w wybranym terminie istnieją już ograniczenia. Czy chcesz je zastąpić?", vbQuestion + vbYesNo, "Istniejące ograniczenia")
    If resp = vbYes Then
        For Each nRes In restrictions
            tNow = trucksOnDate(nRes.resDate)
            tMax = getMaxSlot(nRes.resDate)
            If tNow < tMax Then
                If (tMax - tNow) < Me.txtHours Then
                    ' we have to shorten the restriction
                    h = tMax - tNow
                    'save the date that you have just shortened
                    rMsg = rMsg & nRes.resDate & ","
                Else
                    h = Me.txtHours
                End If
                uStr = "UPDATE tbCalendarRestrictions SET slotsTaken=" & h & ", description='" & Me.txtDescription & "' WHERE calendarRestrictionsId=" & nRes.ID
                adoConn.Execute uStr
                i = i + 1
            End If
        Next nRes
    End If
End If

If Len(rMsg) > 0 Then
    rMsg = vbNewLine & "Ilość zarezerwowanych godzin dla ograniczenia została dostosowana do ilości wolnych slotów dla następujących dni: " & Left(rMsg, Len(rMsg) - 1)
End If
mStr = "Dodano " & i & " ograniczeń w zakresie od " & Me.txtDateFrom & " do " & Me.txtDateTo & rMsg

killForm "frmNotify"
MsgBox mStr, vbInformation + vbOKOnly, "Powodzenie"

End Sub
