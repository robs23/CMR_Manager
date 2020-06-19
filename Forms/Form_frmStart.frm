VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub btnAddComp_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("COMPANY_CREATE"), whoIsLogged) Then
    Call launchForm("frmEditCompany")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnAddCont_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("CONTACT_CREATE"), whoIsLogged) Then
    Call launchForm("frmEditContact")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnBrowseCmrs_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("CMR_BROWSE"), whoIsLogged) Then
    Call launchForm("frmCmrs")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnBrowseComp_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("COMPANY_BROWSE"), whoIsLogged) Then
    Call launchForm("frmBrowseCompany")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnBrowseCont_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("CONTACT_BROWSE"), whoIsLogged) Then
    Call launchForm("frmBrowseContacts")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnCalendarView_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("CMR_BROWSE"), whoIsLogged) Then
    Call launchForm("frmWeekView")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnComponents_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("MATERIAL_BROWSE"), whoIsLogged) Then
    Call launchForm("frmMaterials")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnDeveloper_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("DEVELOPER_MODE"), whoIsLogged) Then
    Call launchForm("frmDBSettings")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnGenerate_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("TRANSPORT_CREATE"), whoIsLogged) Then
    Call launchForm("frmGenerate")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnNewShipment_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("TRANSPORT_CREATE"), whoIsLogged) Then
    Call launchForm("frmTransport")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnPrinterSetup_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("PRINTER_BROWSE"), whoIsLogged) Then
    Call launchForm("frmPrinterSetup")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnPrivilages_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(17, whoIsLogged) Then
    Call launchForm("frmPrivAssign")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnProducts_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("PRODUCT_BROWSE"), whoIsLogged) Then
    Call launchForm("frmZfinOverview")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnReports_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("REPORT_BROWSE"), whoIsLogged) Then
    Call launchForm("frmReportChoose")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnReqs_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("REQS_BROWSE"), whoIsLogged) Then
    Call launchForm("frmReqs")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnRestrictions_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("RESTRICTIONS_CHANGE"), whoIsLogged) Then
    Call launchForm("frmCalendarRestrictions")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnSettings_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("SETTINGS_BROWSE"), whoIsLogged) Then
    Call launchForm("frmSettings")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnTemplates_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("TEMPLATE_BROWSE"), whoIsLogged) Then
    Call launchForm("frmTemplates")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub Form_Load()
Dim cor As Long
cor = Me.InsideWidth - Me.frmGraphContainer.Width - 300
If cor <= Me.btnAddComp.Width + 800 Then
    cor = Me.btnAddComp.Width + 1000
End If
Me.frmGraphContainer.Left = cor
End Sub
