VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmZFIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private mode As Integer '1-add,2-edit,3-preview
Private zfinId As Long 'zfinId in tbZfin

Private Sub btnEdit_Click()
If authorize(getFunctionId("PRODUCT_EDIT"), whoIsLogged) Then
    If productEdited = False Then
        mode = 2
        updateConnection
        adoConn.Execute "UPDATE tbZfin SET isBeingEditedBy=" & whoIsLogged & "  WHERE zfinId=" & zfinId
        changeLock True
    End If
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Function productEdited() As Boolean
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT * FROM tbZfin WHERE zfinId=" & zfinId)
Set rs.ActiveConnection = Nothing

If rs.EOF Then
    productEdited = False
Else
    If Not IsNull(rs.fields("isBeingEditedBy")) Then
        MsgBox "Produkt jest w tej chwili edytowany przez " & getUserName(rs.fields("isBeingEditedBy")) & ". Spróbuj ponownie później", vbInformation + vbOKOnly, "Produkt w edycji"
        productEdited = True
    Else
        productEdited = False
    End If
End If

rs.Close
Set rs = Nothing

End Function

Private Sub btnNewZfor_Click()
If authorize(getFunctionId("MATERIAL_CREATE"), whoIsLogged) Then
    Call launchForm("frmMaterial")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnSave_Click()
SaveZfin
End Sub

Private Sub cmbPalletType_AfterUpdate()
If Not IsNull(Me.cmbPalletType) Then
    Me.txtPalletType = "Rozmiar: " & Me.cmbPalletType.Column(2) & ", " & Me.cmbPalletType.Column(3)
    Me.txtPalletType.visible = True
Else
    Me.txtPalletType = ""
    Me.txtPalletType.visible = False
End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.ActiveControl.Name = "txtKillbox1" Then
    Select Case KeyCode
    Case vbKeyO
    Me.tab.Pages("pgOverview").SetFocus
    Case vbKeyU
    Me.tab.Pages("pgUOM").SetFocus
    Case vbKeyW
    Me.tab.Pages("pgProperties").SetFocus
    Case vbKeyD
    Me.tab.Pages("pgDelivery").SetFocus
    Case vbKeyP
    Me.tab.Pages("pgProduction").SetFocus
    End Select
    KeyCode = 0
    Me.txtKillbox1.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim sql As String

sql = "SELECT custStringId, custString FROM tbCustomerString"

populateListboxFromSQL sql, Me.cmbKlient

sql = "SELECT zfinId, CONVERT(nvarchar,zfinIndex) + ' ' + zfinName as zfinName FROM tbZfin WHERE zfinType='zfor'"

populateListboxFromSQL sql, Me.cmbZfor

sql = "SELECT p.palletId, p.palletName, CONVERT(varchar,p.palletWidth) + 'x' + CONVERT(varchar,p.palletLength) as palletSize, CASE WHEN p.palletChep=1 THEN 'CHEP' ELSE CASE WHEN p.palletChep=0 THEN 'NIE CHEP' ELSE 'B/D' END END as palletChep FROM tbPallets p ORDER BY p.palletId;"

populateListboxFromSQL sql, Me.cmbPalletType

If IsNull(Me.openArgs) Then
    mode = 1
    Me.txtCreationDetails.visible = False
    Me.txtUpdateDetails.visible = False
Else
    If IsNumeric(Me.openArgs) Then
        Me.txtCreationDetails.visible = False
        Me.txtUpdateDetails.visible = False
        mode = 3
        zfinId = CLng(Me.openArgs)
        bringZFIN
    Else
        mode = 1
        Me.txtCreationDetails.visible = False
        Me.txtUpdateDetails.visible = False
    End If
End If
If mode = 3 Then
    changeLock (False)
Else
    changeLock (True)
End If

Call killForm("frmNotify")

Me.txtKillbox1.SetFocus
End Sub

Sub SaveZfin()
Dim rs As ADODB.Recordset
Dim bool As Boolean
Dim weight As Double
Dim pcPal As Long
Dim pcLay As Integer
Dim pcBox As Integer
Dim pallType As Integer
Dim uStr As String

On Error GoTo err_trap


If mode = 1 Then
    newNotify "Przygotowanie do dodania produktu.. Proszę czekać.."
    If Len(Me.txtIndex.value) > 0 And Len(Me.txtDescription.value) > 0 Then
        If Not IsNull(adoDLookup("zfinIndex", "tbZfin", "zfinIndex=" & Me.txtIndex.value)) Then
            MsgBox """Numer Zfin"" musi być unikatowy! Wprowadź unikatowy numer w polu ""Numer Zfin"" by kontynuować.", vbOKOnly + vbExclamation, "Wykryto duplikat"
            killForm "frmNotify"
        Else
            uStr = "INSERT INTO tbZfin (zfinIndex, zfinName, zfinType, prodStatus, creationDate, createdBy, custString) VALUES ("
            uStr = uStr & fSql(Me.txtIndex.value, True) & ","
            uStr = uStr & fSql(Me.txtDescription.value) & ","
            uStr = uStr & "'zfin','pr','" & Now & "'," & whoIsLogged & ","
            uStr = uStr & fSql(Me.cmbKlient.value, True) & ")"
            updateConnection
            Set rs = adoConn.Execute(uStr & ";SELECT SCOPE_IDENTITY()")
            Set rs = rs.NextRecordset
            zfinId = rs.fields(0).value
            rs.Close
            Set rs = Nothing
            If Not IsNull(Me.cmbZfor) Then
                newNotify "Dodawanie informacji Zfin - Zfor.. Proszę czekać.."
                updateConnection
                adoConn.Execute "INSERT INTO tbZfinZfor (zfinId, zforId) VALUES (" & zfinId & ", " & Me.cmbZfor & ")"
            End If
            newNotify "Sprawdzanie poprawności danych UoM.. Proszę czekać.."
            If Len(Me.txtPcWeight.value) > 0 Or Len(Me.txtPcPal.value) > 0 Or Len(Me.txtpcBox.value) > 0 Or Len(Me.txtPcLay.value) > 0 Or Not IsNull(Me.cmbPalletType) Then
                bool = True
                If Len(Me.txtPcWeight.value) > 0 Then
                    If IsNumeric(Me.txtPcWeight.value) Then
                        weight = Me.txtPcWeight.value
                    Else
                        bool = False
                        MsgBox "Proszę zmienić zawartość pola ""Waga szt."" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                    End If
                End If
                If Len(Me.txtPcPal.value) > 0 Then
                    If IsNumeric(Me.txtPcPal.value) Then
                        pcPal = Me.txtPcPal.value
                    Else
                        bool = False
                        MsgBox "Proszę zmienić zawartość pola ""PC/PAL"" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                    End If
                End If
                If Len(Me.txtpcBox.value) > 0 Then
                    If IsNumeric(Me.txtpcBox.value) Then
                        pcBox = Me.txtpcBox.value
                    Else
                        bool = False
                        MsgBox "Proszę zmienić zawartość pola ""PC/BOX"" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                    End If
                End If
                If Len(Me.txtPcLay.value) > 0 Then
                    If IsNumeric(Me.txtPcLay.value) Then
                        pcLay = Me.txtPcLay.value
                    Else
                        bool = False
                        MsgBox "Proszę zmienić zawartość pola ""PC/LAY"" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                    End If
                End If
                If Not IsNull(Me.cmbPalletType) Then
                    pallType = Me.cmbPalletType.value
                End If
                If bool Then
                    newNotify "Dodawanie danych UoM.. Proszę czekać.."
                    Set rs = newRecordset("tbUom", True)
                    rs.AddNew
                    rs.fields("unitWeight") = weight
                    rs.fields("pcPerBox") = pcBox
                    rs.fields("pcPerPallet") = pcPal
                    rs.fields("pcLayer") = pcLay
                    rs.fields("zfinId") = zfinId
                    If pallType > 0 Then rs.fields("palletType") = pallType
                    rs.update
                    rs.Close
                    Set rs = Nothing
                    newNotify "Dodawanie właściwości produktu.. Proszę czekać.."
                    Set rs = newRecordset("tbZfinProperties", True)
                    rs.AddNew
                    rs.fields("zfinId") = zfinId
                    If Not IsNull(Me.cboxBeans) Then rs.fields("beans?") = Me.cboxBeans
                    If Not IsNull(Me.cboxDecafe) Then rs.fields("decafe?") = Me.cboxDecafe
                    If Not IsNull(Me.cboxEco) Then rs.fields("eco?") = Me.cboxEco
                    If Not IsNull(Me.cboxSingle) Then rs.fields("single-origin?") = Me.cboxSingle
                    If Not IsNull(Me.cboxUtz) Then rs.fields("utz?") = Me.cboxUtz
                    If Not IsNull(Me.cboxAromatic) Then rs.fields("aromatic?") = Me.cboxAromatic
                    rs.update
                    rs.Close
                    Set rs = Nothing
                End If
            End If
            clearall
            MsgBox "Zapis zakończony powodzeniem", vbOKOnly + vbInformation, "Zapisano!"
        End If
    Else
        MsgBox "Pola ""Numer ZFIN"" i ""Opis"" w zakładce Ogólne muszą być wypełnione by kontynuować", vbOKOnly + vbInformation, "Brakujące dane"
    End If
ElseIf mode = 2 Then
    newNotify "Przygotowanie do edycji produktu.. Proszę czekać.."
    If Len(Me.txtIndex.value) > 0 And Len(Me.txtDescription.value) > 0 Then
        Set rs = newRecordset("SELECT * FROM tbZfin WHERE zfinId = " & zfinId, True)
        If Not rs.EOF Then
            rs.fields("zfinIndex") = Me.txtIndex.value
            rs.fields("zfinName") = Me.txtDescription.value
            rs.fields("lastUpdate") = Now
            rs.fields("lastUpdateBy") = whoIsLogged()
            If Not IsNull(Me.cmbKlient.value) Then rs.fields("custString") = Me.cmbKlient.value
            rs.update
        End If
        rs.Close
        Set rs = Nothing
        If Not IsNull(Me.cmbZfor) Then
            newNotify "Edycja informacji Zfin - Zfor.. Proszę czekać.."
            Set rs = newRecordset("SELECT zfinId FROM tbZfinZfor WHERE zfinId = " & zfinId)
            Set rs.ActiveConnection = Nothing
            If rs.EOF Then
                bool = True
            Else
                bool = False
            End If
            rs.Close
            Set rs = Nothing
            updateConnection
            If Not bool Then
                adoConn.Execute "UPDATE tbZfinZfor SET zforId = " & Me.cmbZfor & " WHERE zfinId = " & zfinId
            Else
                adoConn.Execute "INSERT INTO tbZfinZfor (zfinId, zforId) VALUES (" & zfinId & ", " & Me.cmbZfor & ")"
            End If
        End If
        newNotify "Sprawdzanie poprawności danych UoM.. Proszę czekać.."
        If Len(Me.txtPcWeight.value) > 0 Or Len(Me.txtPcPal.value) > 0 Or Len(Me.txtpcBox.value) > 0 Or Len(Me.txtPcLay.value) > 0 Or Not IsNull(Me.cmbPalletType) Then
            bool = True
            If Len(Me.txtPcWeight.value) > 0 Then
                If IsNumeric(Me.txtPcWeight.value) Then
                    weight = Me.txtPcWeight.value
                Else
                    bool = False
                    MsgBox "Proszę zmienić zawartość pola ""Waga szt."" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                End If
            End If
            If Len(Me.txtPcPal.value) > 0 Then
                If IsNumeric(Me.txtPcPal.value) Then
                    pcPal = Me.txtPcPal.value
                Else
                    bool = False
                    MsgBox "Proszę zmienić zawartość pola ""PC/PAL"" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                End If
            End If
            If Len(Me.txtpcBox.value) > 0 Then
                If IsNumeric(Me.txtpcBox.value) Then
                    pcBox = Me.txtpcBox.value
                Else
                    bool = False
                    MsgBox "Proszę zmienić zawartość pola ""PC/BOX"" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                End If
            End If
            If Len(Me.txtPcLay.value) > 0 Then
                If IsNumeric(Me.txtPcLay.value) Then
                    pcLay = Me.txtPcLay.value
                Else
                    bool = False
                    MsgBox "Proszę zmienić zawartość pola ""PC/LAY"" na wartość numeryczną", vbOKOnly + vbInformation, "Nieprawidłowy typ danych"
                End If
            End If
            If Not IsNull(Me.cmbPalletType) Then
                pallType = Me.cmbPalletType.value
            End If
            If bool Then
                newNotify "Edycja danych UoM.. Proszę czekać.."
                Set rs = newRecordset("SELECT * FROM tbUom WHERE zfinId = " & zfinId, True)
                If rs.EOF Then
                    rs.AddNew
                End If
                rs.fields("unitWeight") = weight
                rs.fields("pcPerBox") = pcBox
                rs.fields("pcPerPallet") = pcPal
                rs.fields("pcLayer") = pcLay
                rs.fields("zfinId") = zfinId
                If pallType > 0 Then
                    rs.fields("palletType") = pallType
                Else
                    rs.fields("palletType") = Null
                End If
                rs.update
                rs.Close
                Set rs = Nothing
                newNotify "Edycja właściwości produktu.. Proszę czekać.."
                Set rs = newRecordset("SELECT * FROM tbZfinProperties WHERE zfinId = " & zfinId, True)
                If rs.EOF Then
                    rs.AddNew
                End If
                rs.fields("zfinId") = zfinId
                rs.fields("beans?") = Me.cboxBeans
                rs.fields("decafe?") = Me.cboxDecafe
                rs.fields("eco?") = Me.cboxEco
                rs.fields("single-origin?") = Me.cboxSingle
                rs.fields("utz?") = Me.cboxUtz
                rs.fields("aromatic?") = Me.cboxAromatic
                rs.update
                rs.Close
                Set rs = Nothing
            End If
        End If
        clearall
        MsgBox "Zapis zakończony powodzeniem", vbOKOnly + vbInformation, "Zapisano!"
    Else
        MsgBox "Pola ""Numer ZFIN"" i ""Opis"" w zakładce Ogólne muszą być wypełnione by kontynuować", vbOKOnly + vbInformation, "Brakujące dane"
    End If
End If


Exit_here:
killForm "frmNotify"
Exit Sub

err_trap:
MsgBox "Error in ""SaveZfin"" of frmZfin. Error number: " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Sub clearall()
Dim ctl As Control

For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If ctl.Enabled And ctl.ControlSource = "" Then Me.Controls(ctl.Name).value = ""
    ElseIf ctl.ControlType = acComboBox Then
        If ctl.Enabled Then Me.Controls(ctl.Name) = Null
    ElseIf ctl.ControlType = acCheckBox Then
        If ctl.Enabled Then Me.Controls(ctl.Name).value = 0
    End If
Next ctl

End Sub

Sub calcKgPal()
If Len(Me.txtPcWeight.value) > 0 And Len(Me.txtPcPal.value) > 0 Then
    If IsNumeric(Me.txtPcWeight.value) And IsNumeric(Me.txtPcPal.value) Then
        Me.txtKgPal.value = Me.txtPcWeight.value * Me.txtPcPal.value
    End If
End If
End Sub

Sub calcLayPal()
If Len(Me.txtPcLay.value) > 0 And Len(Me.txtPcPal.value) > 0 And Me.txtPcLay.value > 0 And Me.txtPcPal.value > 0 Then
    If IsNumeric(Me.txtPcLay.value) And IsNumeric(Me.txtPcPal.value) Then
        Me.txtLayPal.value = Me.txtPcPal.value / Me.txtPcLay.value
    End If
End If
End Sub


Private Sub Form_Resize()
Me.tab.Width = Me.InsideWidth - 400
Me.tab.Height = Me.InsideHeight - 700
Me.btnSave.Left = Me.tab.Left + Me.tab.Width - Me.btnSave.Width
Me.btnEdit.Left = Me.btnSave.Left - Me.btnEdit.Width - 100
Me.subFrmBom.Width = Me.tab.Width - 200
Me.subFrmBom.Height = Me.tab.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
If mode = 2 Then
    updateConnection
    adoConn.Execute "UPDATE tbZfin SET isBeingEditedBy=NULL  WHERE zfinId=" & zfinId
End If
End Sub

Private Sub tab_KeyUp(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'Case vbKeyO
'Me.tab.Pages("pgOverview").SetFocus
'Case vbKeyU
'Me.tab.Pages("pgUOM").SetFocus
'Case vbKeyW
'Me.tab.Pages("pgProperties").SetFocus
'Case vbKeyD
'Me.tab.Pages("pgDelivery").SetFocus
'Case vbKeyP
'Me.tab.Pages("pgProduction").SetFocus
'End Select
End Sub


Private Sub txtPcLay_AfterUpdate()
calcLayPal
End Sub

Private Sub txtPcPal_AfterUpdate()
calcKgPal
End Sub

Private Sub txtPcWeight_AfterUpdate()
calcKgPal
calcLayPal
End Sub

Sub bringZFIN()
Dim rs As ADODB.Recordset
Dim str As String
Dim sql As String
Dim dFrom As Date
Dim dTo As Date

Set rs = newRecordset("SELECT * FROM tbZfin WHERE zfinId = " & zfinId)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    Me.txtIndex.value = rs.fields("zfinIndex")
    Me.txtDescription.value = rs.fields("zfinName")
    If Not IsNull(rs.fields("creationDate")) Then
        If Not IsNull(rs.fields("createdBy")) Then
            Me.txtCreationDetails.value = "Utworzony w dniu <b>" & rs.fields("creationDate") & "</b> przez <b>" & getUserName(rs.fields("createdBy")) & "</b>"
            Me.txtCreationDetails.visible = True
        End If
    End If
    If Not IsNull(rs.fields("lastUpdate")) Then
        If Not IsNull(rs.fields("lastUpdateBy")) Then
            Me.txtUpdateDetails.value = "Ostatnio edytowany w dniu <b>" & rs.fields("lastUpdate") & "</b> przez <b>" & getUserName(rs.fields("lastUpdateBy")) & "</b>"
            Me.txtUpdateDetails.visible = True
        End If
    End If
    If Not IsNull(rs.fields("custString")) Then Me.cmbKlient.value = rs.fields("custString")
End If
rs.Close
Set rs = Nothing

Set rs = newRecordset("SELECT zforId FROM tbZfinZfor WHERE zfinId = " & zfinId)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    Me.cmbZfor = rs.fields("zforId")
End If
rs.Close
Set rs = Nothing

Set rs = newRecordset("SELECT * FROM tbUom WHERE zfinId = " & zfinId)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    Me.txtPcWeight.value = rs.fields("unitWeight")
    Me.txtPcPal.value = rs.fields("pcPerPallet")
    Me.txtpcBox.value = rs.fields("pcPerBox")
    Me.txtPcLay.value = rs.fields("pcLayer")
    If Not IsNull(rs.fields("palletType")) Then Me.cmbPalletType = rs.fields("palletType")
End If

rs.Close
Set rs = Nothing

sql = "SELECT cs.custStringId " _
    & "FROM tbZfin z LEFT JOIN tbCustomerString cs ON cs.custStringId=z.custString " _
    & "WHERE z.zfinId = " & zfinId

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    Me.cmbKlient = rs.fields("custStringId")
End If
rs.Close
Set rs = Nothing

sql = "SELECT zz.zforId as zfor " _
    & "FROM tbZfin zfin LEFT JOIN tbZfinZfor zz ON zfin.zfinId=zz.zfinId " _
    & "WHERE zfin.zfinId = " & zfinId

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    If Not IsNull(rs.fields("zfor")) Then
        Me.cmbZfor = rs.fields("zfor")
    End If
End If
rs.Close
Set rs = Nothing


Set rs = newRecordset("SELECT * FROM tbZfinProperties WHERE zfinId=" & zfinId)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Me.cboxAromatic = rs.fields("aromatic?")
    Me.cboxBeans = rs.fields("beans?")
    Me.cboxDecafe = rs.fields("decafe?")
    Me.cboxEco = rs.fields("eco?")
    Me.cboxSingle = rs.fields("single-origin?")
    Me.cboxUtz = rs.fields("utz?")
End If

rs.Close
Set rs = Nothing

dTo = Date
dFrom = DateAdd("m", -6, dTo)

Do Until year(dFrom) < 2016

    sql = "SELECT m.machineName, SUM(od.plAmount) as Amount " _
        & "FROM tbOperations o LEFT JOIN tbOperationData od ON od.operationId=o.operationId LEFT JOIN tbZfin z ON z.zfinId=o.zfinId LEFT JOIN tbMachine m ON m.machineId=od.plMach " _
        & "WHERE z.zfinId=" & zfinId & " AND od.plMoment BETWEEN '" & dFrom & "' AND '" & dTo & "' " _
        & "GROUP BY m.machineName " _
        & "ORDER BY Amount DESC"
    
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        ' there's production found, exit loop
        Exit Do
    Else
        'no production, digg dipper
        dFrom = DateAdd("m", -6, dFrom)
        rs.Close
        Set rs = Nothing
    End If
Loop

If rs Is Nothing Then
    'no production found
    Me.txtboxMainLine = "Produkt nieprodukowany"
    Me.lblStat.Caption = "Na podstawie danych o produkcji w okresie " & dFrom & " - " & dTo
Else
    rs.MoveFirst
    Me.txtboxMainLine = rs.fields("machineName")
    rs.MoveNext
    Do Until rs.EOF
        Me.listOtherLine.AddItem rs.fields("machineName")
        rs.MoveNext
    Loop
    Me.lblStat.Caption = "Na podstawie danych o produkcji w okresie " & dFrom & " - " & dTo
    rs.Close
    Set rs = Nothing
End If

sql = "DECLARE @index int " _
    & "SET @index = " & zfinId & " " _
    & "SELECT comp.zfinIndex, comp.zfinName, bom.amount, bom.unit, comp.zfinType, mt.materialTypeName " _
    & "FROM tbZfin zfin LEFT JOIN tbBom bom ON bom.zfinId=zfin.zfinId LEFT JOIN tbZfin comp ON comp.zfinId=bom.materialId LEFT JOIN tbMaterialType mt ON mt.materialTypeId=comp.materialType " _
    & "WHERE zfin.zfinId=@index AND bom.bomRecId=(SELECT TOP(1) br.bomRecId FROM tbBomReconciliation br LEFT JOIN tbBom bm ON bm.bomRecId=br.bomRecId LEFT JOIN tbZfin z ON z.zfinId=bm.zfinId WHERE z.zfinId=@index ORDER BY br.dateAdded DESC)"

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    With Me.subFrmBom.Form
        Set .Recordset = rs
        .Controls("txtIndex").ControlSource = "zfinIndex"
        .Controls("txtName").ControlSource = "zfinName"
        .Controls("txtAmount").ControlSource = "amount"
        .Controls("txtUnit").ControlSource = "unit"
        .Controls("txtType").ControlSource = "zfinType"
        .Controls("txtMaterialType").ControlSource = "materialTypeName"
        .Controls("txtName").ColumnWidth = -2
    End With
End If

rs.Close
Set rs = Nothing

txtPcLay_AfterUpdate
txtPcPal_AfterUpdate
txtPcWeight_AfterUpdate
cmbPalletType_AfterUpdate

If Len(Me.txtIndex.value) > 0 Then
    str = Me.txtIndex.value
End If
If Len(Me.txtDescription.value) > 0 Then
    str = str & " " & Me.txtDescription.value
End If
If str <> "" Then Me.Caption = Me.Caption & " || " & str

End Sub

Sub changeLock(bool As Boolean)
Dim ctl As Control

For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acCheckBox Or ctl.ControlType = acListBox Or ctl.ControlType = acCommandButton Then
        If ctl.Name <> "txtKgPal" And ctl.Name <> "txtLayPal" And ctl.Name <> "txtKillbox1" And ctl.Name <> "txtPalletType" And ctl.Name <> "txtCreationDetails" And ctl.Name <> "txtUpdateDetails" Then
            ctl.Enabled = bool
            If ctl.ControlType = acCommandButton Then
                ctl.UseTheme = bool
            End If
        End If
    End If
Next ctl

If bool Then
    Me.btnEdit.Enabled = False
    Me.btnEdit.UseTheme = False
Else
    Me.btnEdit.Enabled = True
    Me.btnEdit.UseTheme = True
End If
Me.txtboxMainLine.Enabled = False

End Sub


