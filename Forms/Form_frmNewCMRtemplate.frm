VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNewCMRtemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mode As Integer '1 new Template, 2 edit Template, 3 preview Template
Private printable As Boolean
Private edit_Id As Long 'number of temp that is being edited

    
Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call fResetBorders(116)
End Sub

Private Sub Form_Load()
Dim Template As Integer
Me.Controls("lblPage").visible = False
Me.Controls("lblPageDesc").visible = False
Call fSetupMouseCtls
If mode = 1 Then
    deployVars
    Me.Caption = "Nowy szablon CMR"
    Me.ShortcutMenu = False
ElseIf mode = 2 Then
    If isTheFormLoaded("frmDeliveryTemplate") Then
        If Not IsNull(Forms("frmDeliveryTemplate").Controls("cmbCmrTemplate")) Then
            Template = Forms("frmDeliveryTemplate").Controls("cmbCmrTemplate").value
            bringTemplate (Template)
            Me.Caption = "Edytuj szablon CMR"
            Me.ShortcutMenu = False
        End If
    ElseIf isTheFormLoaded("frmTemplates") Then
        Template = Forms("frmTemplates").Controls("subFrmTemplates").Form.Controls("cmrId").value
        bringTemplate (Template)
        Me.Caption = "Edytuj szablon CMR"
        Me.ShortcutMenu = False
    End If
    edit_Id = Template
    editOn (Template)
ElseIf mode = 3 Then
    If currentCmr.templateId > 0 Then
        Template = currentCmr.templateId
        previewCMR (Template)
        Me.Caption = "Podgląd wydruku"
        Me.ShortcutMenu = True
    End If
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
If IsNull(Me.openArgs) Then
    mode = 1
Else
    Select Case Me.openArgs
        Case Is = "New"
            mode = 1
        Case Is = "Edit"
            mode = 2
        Case Is = "Preview"
            mode = 3
    End Select
End If
'Me.lbl1.Caption = "Nadawca (nazwisko lub nazwa, adres, kraj)" & vbNewLine & "Absender (Name, Anschrift, Land)" & vbNewLine & "Sender (name, address, country)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim res As VbMsgBoxResult
Dim temp As Integer

If mode = 1 Then
    Dim db As DAO.Database
    Dim rs As ADODB.Recordset
    Dim TemplateName As String
    Dim prevTemp As String
    Dim dupl As Boolean
    Dim isOk As Boolean
    Cancel = True
    dupl = False
    isOk = False
    res = MsgBox("Czy chcesz zapisać bieżący układ jako szablon?", vbQuestion + vbYesNo, "Tworzenie nowego szablonu")
    If res = vbYes Then
        Do
            TemplateName = InputBox("Pod jaką nazwą chcesz zapisać ten szablon?", "Podaj nazwę")
            If TemplateName = "" And StrPtr(TemplateName) <> 0 Then
                Do While TemplateName = "" And StrPtr(TemplateName) <> 0
                    TemplateName = InputBox("Nazwa szablonu nie może być pusta. Pod jaką nazwą chcesz zapisać ten szablon?", "Podaj nazwę")
                Loop
            ElseIf StrPtr(TemplateName) = 0 Then
                Exit Do
            End If
            If TemplateName <> "" Then
                Set db = CurrentDb
                Do
                    If TemplateName <> "" Then
                        Set rs = newRecordset("SELECT * FROM tbCmrTemplate WHERE tempName = '" & TemplateName & "'")
                        Set rs.ActiveConnection = Nothing
                        If Not rs.EOF Then
                            dupl = True
                            prevTemp = TemplateName
                            TemplateName = InputBox("Podana nazwa jest już w użyciu, podaj inną", "Nazwa w użyciu", prevTemp)
                            If StrPtr(TemplateName) = 0 Then
                                isOk = True
                                Exit Do
                            End If
                        Else
                            dupl = False
                            isOk = True
                            Call saveTemplate(TemplateName)
                            If isTheFormLoaded("frmDeliveryTemplate") Then
                                Forms("frmDeliveryTemplate").Requery
                                Forms("frmDeliveryTemplate").Refresh
                            ElseIf isTheFormLoaded("frmEditCompany") Then
                                Forms("frmEditCompany").Requery
                                Forms("frmEditCompany").Refresh
                            ElseIf isTheFormLoaded("frmTemplates") Then
                                Forms("frmTemplates").Requery
                                Forms("frmTemplates").Refresh
                            ElseIf isTheFormLoaded("frmAddCompany") Then
                                Forms("frmAddCompany").Requery
                                Forms("frmAddCompany").Refresh
                            End If
                        End If
                        rs.Close
                        Set rs = Nothing
                    Else
                        dupl = False
                    End If
                Loop While dupl
                Set db = Nothing
            End If
        Loop While isOk = False
    End If
    Cancel = False
ElseIf mode = 2 Then
    Cancel = True
    res = MsgBox("Czy chcesz zapisać zmiany w układzie bieżącego szablonu?", vbQuestion + vbYesNo, "Edycja szablonu")
    If res = vbYes Then
        If isTheFormLoaded("frmDeliveryTemplate") Then
            If Not IsNull(Forms("frmDeliveryTemplate").Controls("cmbTemplate")) Then
                temp = Forms("frmDeliveryTemplate").Controls("cmbCmrTemplate")
                editTemplate (temp)
                currentCmr.Reload
            End If
        ElseIf isTheFormLoaded("frmTemplates") Then
            temp = Forms("frmTemplates").Controls("subFrmTemplates").Form.Controls("cmrId").value
            editTemplate (temp)
            Forms("frmTemplates").Requery
            Forms("frmTemplates").Refresh
        End If
        Cancel = False
    Else
        Cancel = False
    End If
    editOff (edit_Id)
ElseIf mode = 3 Then
    
End If
End Sub

Private Function fSetupMouseCtls() As Boolean
  Dim ctl As Control
  For Each ctl In Me
    If InStr(1, ctl.Name, "in", vbTextCompare) = 1 Then
      ctl.OnMouseMove = "=fBorderColor(""" & ctl.Name & """,255)"    ' vbRed = 255
    End If
  Next
  fSetupMouseCtls = (Err = 0)
End Function

Private Function fBorderColor(strCtlName As String, lColor As Long) As Boolean
Dim ctl As Control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox And InStr(1, ctl.Name, "in", vbTextCompare) = 1 Then
        If ctl.Name <> strCtlName Then
            If ctl.BorderColor <> 116 Then
                ctl.BorderColor = 116
                ctl.BorderWidth = 1
            End If
        End If
    End If
Next ctl

  With Me(strCtlName)
     If .BorderColor <> lColor Then .BorderColor = lColor
    .BorderWidth = 2
    fBorderColor = .BorderColor = lColor
  End With
End Function

Private Function fResetBorders(Optional lDefaultColor As Long) As Boolean
  Dim ctl As Control
  For Each ctl In Me
        If InStr(1, ctl.Name, "in", vbTextCompare) = 1 Then
          If ctl.BorderColor <> lDefaultColor Then
            ctl.BorderColor = lDefaultColor
            ctl.BorderWidth = 1
        End If
    End If
  Next
  fResetBorders = (Err = 0)
End Function

Function removeBorders() As Boolean
  Dim ctl As Control
  For Each ctl In Me
        If InStr(1, ctl.Name, "in", vbTextCompare) = 1 Then
            ctl.BorderStyle = 0
        End If
  Next
  removeBorders = (Err = 0)
End Function

Private Sub deployVars()

Me.in1a.value = "Jacobs Douwe Egberts PL sp. z o.o. <br> 02-677 Warszawa, ul.Taśmowa 7 <br> NIP 527-27-17-861 REGON 147353687 <br> PALARNIA KAWY w Sułaszewie <br>64-830 Margonin <br>tel. (67)284-71-50, fax (67) 284-71-08"
Me.in2a.value = "[KLIENT]"
Me.in3a.value = "[MAGAZYN]"
Me.in4a.value = "Sułaszewo, Poland, [DATA]"
Me.in5a.value = "[NR_TRANSPORTU]"
Me.in5b.value = "Delivery Note [DELIVERY_NOTE]"
Me.in6a.value = "Kawa mielona na paletach - "
Me.in6b.value = "Waga netto - "
Me.in6c.value = "[ILOSC_PALET] pal."
Me.in6d.value = "[WAGA_N] Kg"
Me.in6e.value = "[WAGA_B] Kg"
Me.in21a.value = "Sułaszewo"
Me.in21b.value = "[DATA]"
Me.in16a.value = "[PRZEWOZNIK]"
Me.in17a.value = "[NUMERY_AUTA]"
End Sub

Private Sub saveTemplate(TemplateName As String)
Dim rs As ADODB.Recordset
Dim i As Integer
Dim Nazwa As String
Dim Index As Integer
Dim iSql As String
Dim ctl As Control
Dim fields As String
Dim values As String

On Error GoTo err_trap

updateConnection

iSql = "INSERT INTO tbCmrTEMPDetail (<fields>) VALUES (<values>)"

For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If InStr(1, ctl.Name, "in", vbTextCompare) > 0 Then
            If Len(Me.Controls(ctl.Name).value & "") <> 0 Then
                'field is filled in so we have to add its name & content to INSERT string iSql
                fields = fields & ctl.Name & ","
                values = values & "'" & Me.Controls(ctl.Name).value & "'" & ","
            End If
        End If
    End If
Next ctl

If Len(fields) > 0 Then
    fields = Left(fields, Len(fields) - 1)
    values = Left(values, Len(values) - 1)
    iSql = Replace(iSql, "<fields>", fields)
    iSql = Replace(iSql, "<values>", values)
    Set rs = adoConn.Execute(iSql & ";SELECT SCOPE_IDENTITY()")
    Index = rs.fields(0).value
    rs.Close
    Set rs = Nothing
    iSql = ""
    iSql = "INSERT INTO tbCmrTemplate (cmrDate, tempName, detailId, userId) VALUES ('" & Now & "','" & TemplateName & "'," & Index & "," & whoIsLogged & ")"
End If

adoConn.Execute iSql


Exit_here:
If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error number " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Private Sub previewCMR(cmrTemp As Integer)
Dim s() As String
Dim n As Integer
Dim i As Integer
Dim fldName As String
Dim fldValue As String
Dim db As DAO.Database
Dim rs As ADODB.Recordset
Dim fld As DAO.Field
Dim forwarder As Variant

On Error GoTo err_trap

s = parseCustVars(currentCmr.templateId)
Set db = CurrentDb
Set rs = newRecordset("SELECT ctd.* FROM tbCmrTemplate ct LEFT JOIN tbCmrTEMPDetail ctd ON ct.detailId = ctd.cmrDetailId WHERE ct.cmrId=" & cmrTemp)
Set rs.ActiveConnection = Nothing

For i = 0 To rs.fields.count - 1
    If InStr(1, rs.fields(i).Name, "in", vbTextCompare) = 1 Then
        If Not IsNull(rs.fields(i)) Then
            Forms("frmNewCMRTemplate").Controls(rs.fields(i).Name).value = rs.fields(i).value
        End If
    End If
Next i

For n = LBound(s) To UBound(s)
    fldName = ""
    fldValue = ""
    With currentCmr
        Select Case s(n)
            Case Is = "KLIENT"
                fldValue = .SoldToString
            Case Is = "MAGAZYN"
                fldValue = .shipToString
            Case Is = "DATA"
                fldValue = .TransportationDate
            Case Is = "DELIVERY_NOTE"
                fldValue = .deliveryNumbers
            Case Is = "ILOSC_PALET"
                fldValue = .numberOfPallets
            Case Is = "WAGA_N"
                fldValue = .netWeight
            Case Is = "WAGA_B"
                fldValue = .grossWeight
            Case Is = "PRZEWOZNIK"
                fldValue = .ForwarderString
            Case Is = "NR_TRANSPORTU"
                fldValue = .transportNumber
            Case Is = "SPEDYTOR"
                fldValue = .carrierString
            Case Is = "NUMERY_AUTA"
                fldValue = .TruckNumbers
            Case Else
                If Not .getCustomValues(s(n)) Is Nothing Then
                    If Not IsNull(.getCustomValues(s(n)).value) Then
                        fldValue = CStr(.getCustomValues(s(n)).value)
                    Else
                        fldValue = ""
                    End If
                End If
        End Select
    End With
    If fldValue <> "" Then
        Call ReplaceVar(s(n), fldValue)
    End If
Next n


Exit_here:
Set rs = Nothing
Set db = Nothing
Exit Sub

err_trap:
MsgBox "Error in ""previewCMR"". " & Err.number & ", " & Err.description
Resume Exit_here
End Sub

Private Sub ReplaceVar(varName As String, value As String)
Dim ctl As Access.Control
Dim found As Boolean

On Error GoTo err_trap

For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If InStr(1, ctl.Name, "in", vbTextCompare) = 1 Then
            If Not IsNull(ctl) Then
                Me.Controls(ctl.Name).value = Replace(Me.Controls(ctl.Name).value, "[" & varName & "]", value)
            End If
        End If
    End If
Next ctl

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""ReplaceVar"". " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Private Sub bringTemplate(Template As Integer)
Dim rs As ADODB.Recordset
Dim i As Integer

On Error GoTo err_trap

Set rs = newRecordset("SELECT ctd.* FROM tbCmrTemplate ct LEFT JOIN tbCmrTEMPDetail ctd ON ctd.cmrDetailId = ct.detailId WHERE ct.cmrId=" & Template)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    For i = 0 To rs.fields.count - 1
        If InStr(1, rs.fields(i).Name, "in", vbTextCompare) = 1 Then
            If Not IsNull(rs.fields(i)) Then
                Me.Controls(rs.fields(i).Name).value = rs.fields(i).value
            End If
        End If
    Next i
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in ""bringTemplate"". " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Private Sub editTemplate(Template As Integer)
Dim rs As ADODB.Recordset
Dim i As Integer

On Error GoTo err_trap

Set rs = newRecordset("SELECT ctd.* FROM tbCmrTemplate ct LEFT JOIN tbCmrTEMPDetail ctd ON ctd.cmrDetailId = ct.detailId WHERE ct.cmrId=" & Template, True)

If Not rs.EOF Then
    For i = 0 To rs.fields.count - 1
        If InStr(1, rs.fields(i).Name, "in", vbTextCompare) = 1 Then
            rs.fields(i).value = Me.Controls(rs.fields(i).Name).value
        End If
    Next i
    rs.UpdateBatch
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error number " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Sub editOn(temp As Long)

updateConnection
adoConn.Execute "UPDATE tbCmrTemplate SET isBeingEditedBy=" & whoIsLogged & " WHERE cmrId = " & temp

End Sub

Sub editOff(temp As Long)

updateConnection
adoConn.Execute "UPDATE tbCmrTemplate SET isBeingEditedBy=NULL WHERE cmrId = " & temp

End Sub


Public Sub in16a_AfterUpdate()
Dim dec As VbMsgBoxResult

If mode = 3 Then
        If Not Len(Me.in16a.value) = 0 And Len(currentDelivery.NUMERY_REJESTRACYJNE) > 0 Then
            dec = MsgBox("Czy powiązać wprowadzone dane firmy transportowej z numerami rejestracyjnymi " & currentDelivery.NUMERY_REJESTRACYJNE & "?", vbQuestion + vbYesNo, "Zapamiętać firmę?")
            If dec = vbYes Then
                saveCompany Me.in16a.value, currentDelivery.NUMERY_REJESTRACYJNE
            End If
        End If
End If
End Sub

Sub saveCompany(company As String, TruckNumbers As String)
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim i As Long
Dim n As Integer
Dim v() As String

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbForwarder WHERE forwarderData = '" & Trim(company) & "'", dbOpenDynaset, dbSeeChanges)
If rs.EOF Then
    rs.AddNew
    rs.fields("forwarderData") = Trim(company)
    i = rs.fields("forwarderID")
    rs.update
Else
    rs.MoveFirst
    i = rs.fields("forwarderID")
End If
rs.Close

v() = Split(TruckNumbers, "/", , vbTextCompare)

For n = LBound(v) To UBound(v)
    Set rs = db.OpenRecordset("SELECT * FROM tbTrucks WHERE plateNumbers = '" & Replace(v(n), " ", "") & "'", dbOpenDynaset, dbSeeChanges)
    If Not rs.EOF Then
        'one of numbers already exists
        rs.MoveFirst
        rs.edit
        rs.fields("forwarderId") = i
        rs.update
    End If
    rs.Close
    Set rs = Nothing
Next n

Set rs = Nothing
Set db = Nothing
End Sub

