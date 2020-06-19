VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Sub LadujListeUserow()

Dim rs As ADODB.Recordset

On Error GoTo err_trap

Set rs = New ADODB.Recordset

rs.Open "SELECT * FROM tbUsers ORDER BY userSurname ASC", adoConn, adOpenKeyset, adLockOptimistic
listaLoginow.RowSourceType = "Value List"
listaLoginow.RowSource = ""
If Not (rs.EOF) Then
    rs.MoveFirst
    Do Until rs.EOF = True
    listaLoginow.AddItem rs!UserName & " " & rs!userSurname
    rs.MoveNext
    Loop

End If
Me.przyciskZaloguj.Enabled = False
rs.Close 'Close the recordset

Exit_here:
Set rs = Nothing 'Clean up
Exit Sub

err_trap:
If Err.number = 3151 Then
    MsgBox "Nie udało się nawiązać połączenia z bazą danych. Sprawdź swoje połączenie internetowe, jeśli łączysz się z domu upewnij się, że nawiązałeś połączenie VPN.", vbOKOnly + vbCritical, "Błąd połączenia"
Else
    MsgBox "Error in ""LadujListeUserow"" of frmLogin. Error number: " & Err.number & ", " & Err.description, vbOKOnly + vbExclamation, "Błąd"
End If
Resume Exit_here

End Sub



Private Sub Form_Current()
DoEvents
Call killForm("frmNotify")
DoEvents

End Sub

Private Sub Form_Load()
LadujListeUserow
End Sub

Sub aktywujPrzyciskLogin()
Me.przyciskZaloguj.Enabled = True
End Sub



Private Sub Label7_Click()

End Sub

Private Sub lblPassForgotten_Click()
Dim recipient As Long
Dim bodyText As String
If IsNull(Me.listaLoginow.value) Then
    MsgBox "Zaznacz swojego użytkownika powyżej i kliknij jeszcze raz"
Else
    DoEvents
    Call launchForm("frmNotify")
    DoEvents
    Call newNotify("Proszę czekać..")
    DoEvents
    recipient = userIdentity(Me.listaLoginow.value)
    bodyText = toHtml("Tego maila dostajesz ponieważ poprosiłeś o przypomnienie hasła. Twoje hasło to: ") & toHtml(userPassword(recipient), True) & "<br><br>" & toHtml("Wiadomość wysłana automatycznie, prosimy nie odpowiadać", True)
    Call SendMail(bodyText, "[NPD] Przypomnienie hasła", getMail(recipient), , True)
    DoEvents
    Call killForm("frmNotify")
    DoEvents
    MsgBox "Twoje hasło zostało wysłane na Twój adres mailowy"
End If

End Sub

Private Sub listaLoginow_Click()
aktywujPrzyciskLogin
End Sub

Sub kliknijZaloguj()

DoCmd.OpenForm "frmHaslo"

End Sub



Private Sub listaLoginow_DblClick(Cancel As Integer)
kliknijZaloguj
End Sub

Private Sub przyciskNowyUzytkownik_Click()
DoCmd.Close acForm, "frmLogin", acSaveNo
DoCmd.OpenForm "frmDodajUsera"
End Sub

Private Sub przyciskZaloguj_Click()
kliknijZaloguj
End Sub



