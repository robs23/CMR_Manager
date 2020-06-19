VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDodajUsera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub dodajUzytkownika()
Dim Name, surname, mail, haslo As String
Dim mailBody As String
Dim Index As Integer
Dim mailBody1 As String
Dim sqlString As String
Dim rs As ADODB.Recordset

If Len(Me.txtImie & vbNullString) = 0 Or Len(Me.txtNazwisko & vbNullString) = 0 Or Len(Me.txtMail & vbNullString) = 0 Or Len(Me.txtHaslo & vbNullString) = 0 Or Len(Me.txtPowtorzHaslo & vbNullString) = 0 Then
    MsgBox "Wszystkie pola muszą być wypełnione!"
ElseIf Me.txtHaslo <> Me.txtPowtorzHaslo Then
    MsgBox "Pola HASŁO i POWTÓRZ HASŁO muszą być takie same!", vbOKOnly + vbExclamation, "Różne hasła"
    Me.txtHaslo.value = ""
    Me.txtPowtorzHaslo.value = ""
Else
    Name = txtImie.value
    surname = txtNazwisko.value
    mail = txtMail.value
    haslo = txtHaslo.value
'    If checkIfStringExist("tbUsers", "userMail", mail) Then
'        MsgBox "Ten adres e-mail jest już używany. Podaj inny adres!", vbOKOnly + vbExclamation, "Adres w użyciu"
'        Me.txtMail.value = ""
'    Else
    Call newNotify("Proszę czekać..")
    DoEvents
    Set rs = newRecordset("tbUsers", True)
    rs.AddNew
    rs.fields("userMail") = mail
    rs.fields("userName") = Name
    rs.fields("userSurname") = surname
    rs.fields("userPassword") = haslo
    rs.update
    rs.Close
    Set rs = Nothing
    txtImie.value = ""
    txtNazwisko.value = ""
    txtMail.value = ""
    txtHaslo.value = ""
    txtPowtorzHaslo.value = ""
    MsgBox "Konto zostało dodane"
    Call killForm(Me.Name)
    DoCmd.OpenForm "frmLogin", acNormal, , , acFormEdit, acWindowNormal
'    End If
End If
End Sub


Private Sub przyciskDodajUsera_Click()
dodajUzytkownika
End Sub

