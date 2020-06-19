VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmPrinterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
'Dim prins() As String
'Dim prt As Printer
'
'For Each prt In Printers
'    If isArrayEmpty(prins) Then
'        ReDim prins(0) As String
'        prins(0) = prt.DeviceName
'    Else
'        ReDim Preserve prins(UBound(prins) + 1) As String
'        prins(UBound(prins)) = prt.DeviceName
'    End If
'Next prt
'
'Call populateListbox(Me, Me.printerName, prins)
'
'Me.Parent.Requery
'Me.Parent.Refresh
End Sub
