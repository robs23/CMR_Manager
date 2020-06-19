VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPrinterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub btnSave_Click()
If authorize(getFunctionId("PRINTER_EDIT"), whoIsLogged) Then
    If Not IsNull(Me.cmbDocs) And Not IsNull(Me.cmbPrinters) And Not IsNull(Me.cmbTrays) Then
        saveSettings
        Dim doc As Integer
        MsgBox "Zapis zakończony powodzeniem", vbOKOnly + vbInformation, "Zapisano"
        doc = Me.cmbDocs.value
        Call Form_Load
        Me.cmbDocs = doc
        Call cmbDocs_AfterUpdate
        
    Else
        MsgBox "Wszystkie pola muszą być uzupełnione", vbOKOnly + vbExclamation, "Uzupełnij"
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub cmbDocs_AfterUpdate()
Dim prt As Variant
Dim tray As Variant
Dim i As Integer
Dim ind As Integer
Dim found As Boolean

found = False
If Not IsNull(Me.cmbDocs) Then
    prt = printerSetup(Me.cmbDocs)
    If Not IsNull(prt) Then
        For i = 0 To Me.cmbPrinters.ListCount
            If Me.cmbPrinters.Column(1, i) = prt Then
                found = True
                ind = i
            End If
        Next i
        If found Then
            'selected printer was found
            Me.imgPrinter.visible = False
            Me.cmbPrinters.SetFocus
            Me.cmbPrinters.value = Me.cmbPrinters.ItemData(ind)
            Call cmbPrinters_AfterUpdate
        Else
            'selected printer wasn't found
            Me.imgPrinter.visible = True
            Me.cmbPrinters.value = prt
        End If
    Else
        'selected printer wasn't found
        Me.cmbPrinters.value = prt
    End If
    tray = registryKeyExists("document_" & Me.cmbDocs & "\TrayName")
    If Not tray = False Then
        ind = 0
        found = False
        For i = 0 To Me.cmbTrays.ListCount
            If Me.cmbTrays.Column(1, i) = tray Then
                found = True
                ind = i
            End If
        Next i
        If found Then
            'selected printer was found
            Me.imgTray.visible = False
            Me.cmbTrays.SetFocus
            Me.cmbTrays.value = Me.cmbTrays.ItemData(ind)
        Else
            'selected printer wasn't found
            Me.imgTray.visible = True
            Me.cmbTrays.value = tray
        End If
    Else
        Me.cmbTrays.value = Null
    End If
End If
End Sub

Private Sub cmbPrinters_AfterUpdate()
If Not IsNull(Me.cmbPrinters) Then
    populateTrays (Me.cmbPrinters.Column(1))
    Me.Requery
    Me.Refresh
End If
End Sub

Private Sub Form_Load()
Me.imgPrinter.visible = False
Me.imgTray.visible = False
Call killForm("frmNotify")
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTrays"
DoCmd.SetWarnings True

populateListboxFromSQL "SELECT [tbDocs].[documentId], [tbDocs].[documentName] FROM tbDocs ORDER BY [documentId]", Me.cmbDocs
populatePrinters

Me.Requery
Me.Refresh
End Sub

Sub populatePrinters()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim prt As printer

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbDrivers"
DoCmd.SetWarnings True

Set db = CurrentDb
Set rs = db.OpenRecordset("tbDrivers", dbOpenDynaset, dbSeeChanges)

For Each prt In Printers
    If Not prt.DeviceName = "" Then
        rs.AddNew
        rs.fields("driverName") = prt.DeviceName
        rs.update
    End If
Next prt
rs.Close
Set rs = Nothing
Set db = Nothing
End Sub

Sub populateTrays(printer As String)
Dim db As DAO.Database
Dim rs1 As DAO.Recordset
Dim str As String
Dim newLine() As String
Dim newTab() As String
Dim n As Integer

Set db = CurrentDb

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTrays"
DoCmd.SetWarnings True


Set rs1 = db.OpenRecordset("tbTrays", dbOpenDynaset, dbSeeChanges)
str = ""
str = GetBinList(printer)
If str <> "" Then
    newLine = Split(str, vbCrLf, , vbTextCompare)
    For n = 1 To UBound(newLine)
        newTab = Split(newLine(n), vbTab, , vbTextCompare)
        rs1.AddNew
        rs1.fields("trayValue") = newTab(0)
        rs1.fields("trayName") = newTab(1)
        rs1.update
        Erase newTab
    Next n
End If
rs1.Close

Set rs1 = Nothing
Set db = Nothing

End Sub

Private Sub saveSettings()

updateRegistry "document_" & Me.cmbDocs & "\PrinterName", Me.cmbPrinters.Column(1)
updateRegistry "document_" & Me.cmbDocs & "\tray", Me.cmbTrays.value
updateRegistry "document_" & Me.cmbDocs & "\trayName", Me.cmbTrays.Column(1)
End Sub

Function printerFound(prt As String) As Boolean
Dim pt As printer
Dim bool As Boolean
bool = False

For Each pt In Printers
    If pt.DeviceName = prt Then
        bool = True
    End If
Next pt

printerFound = bool

End Function
