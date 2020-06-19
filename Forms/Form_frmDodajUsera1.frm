VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDodajUsera1"
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
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb


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
    If checkIfStringExist("tbUsers", "userMail", mail) Then
        MsgBox "Ten adres e-mail jest już używany. Podaj inny adres!", vbOKOnly + vbExclamation, "Adres w użyciu"
        Me.txtMail.value = ""
    Else
        Call newNotify("Proszę czekać..")
        DoEvents
        Set rs = db.OpenRecordset("tbUsers", dbOpenDynaset, dbSeeChanges)
        rs.AddNew
        rs.fields("userMail") = mail
        rs.fields("userName") = Name
        rs.fields("userSurname") = surname
        rs.fields("userPassword") = haslo
        rs.update
        mailBody = toHtml("Utworzono nowe konto dla użytkownika ") & toHtml(Name & " " & surname, True) & toHtml(". Przypisz tego użytkownika do odpowiedniej grupy by mógł korzystać z odpowiednich funkcji.") & "<br><br>" & toHtml("Wiadomość wysłana automatycznie, prosimy nie odpowiadać", True)
        Call SendMail(mailBody, "[CMR] Nowy użytkownik", "robert.roszak@demb.com", , True)
        mailBody1 = toHtml("Twoje konto zostało utworzone i jest już aktywne! Automatycznie otrzymałeś uprawnienia do przeglądania podstawowych formularzy. Jeśli zostaną Ci przyznane dodatkowe role/uprawnienia, zostaniesz o tym poinformowany w osobnej wiadomości.")
        mailBody1 = mailBody1 & "<br><br>" & toHtml("Wiadomość wysłana automatycznie, prosimy nie odpowiadać", True)
        Call SendMail(mailBody1, "[CMR] Witaj!", txtMail.value, , True)
'        sqlString = "INSERT INTO tbRoleAssign (userId, roleId) VALUES (" & userIdentity(txtImie.value & " " & txtNazwisko.value) & ", 21)"
'        db.Execute sqlString
        txtImie.value = ""
        txtNazwisko.value = ""
        txtMail.value = ""
        txtHaslo.value = ""
        txtPowtorzHaslo.value = ""
        MsgBox "Konto zostało dodane, jednak dostęp do poszczególnych funkcji będzie ograniczony póki użytkownik nie zostanie przypisany do odpowiedniej grupy. O przypisaniu do grupy poinformujemy mailowo."
        Call killForm(Me.Name)
        DoCmd.OpenForm "frmLogin", acNormal, , , acFormEdit, acWindowNormal
    End If
End If
End Sub


Private Sub przyciskDodajUsera_Click()
dodajUzytkownika
End Sub

