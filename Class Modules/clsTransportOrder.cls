VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransportOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public WithEvents tb As Access.TextBox
Attribute tb.VB_VarHelpID = -1
Public WithEvents tipper As Access.TextBox
Attribute tipper.VB_VarHelpID = -1

Private transportNumber As String 'transport number from Transwide
Private cmrs As New Collection 'collection of all associated cmrs' data
Private transportId As Long
Private pTruckNumbers As String
Private isFinished As Boolean 'has the order been loaded already?
Private isHighlighted As Boolean 'is mouse over it?
Private isSelected As Boolean 'has it been selected by the user?
Private pCarrierString As String
Private isFocused As Boolean 'if control has been clicked lately, it has focus. We have to move the focus so it's not displayed "on the top" of everything else

Public Property Let number(trNum As String)
If trNum <> "" Then
    transportNumber = trNum
End If
End Property

Public Property Get number() As String
number = transportNumber
End Property

Public Property Let carrierString(val As String)
pCarrierString = val
End Property

Public Property Get carrierString() As String
carrierString = pCarrierString
End Property

Public Property Let TruckNumbers(val As String)
pTruckNumbers = val
End Property

Public Property Get TruckNumbers() As String
TruckNumbers = pTruckNumbers
End Property

Public Property Let ID(tranId As Long)
transportId = tranId
End Property

Public Property Get ID() As Long
ID = transportId
End Property

Public Property Let finished(value As Boolean)
isFinished = value
End Property

Public Property Get finished() As Boolean
finished = isFinished
End Property

Public Property Let highlighted(value As Boolean)
isHighlighted = value

On Error GoTo err_trap

If value Then
    tb.SpecialEffect = 1
    If tb.Left + tipper.Width < Forms("frmWeekView").InsideWidth Then
        tipper.Left = tb.Left + tb.Width
    Else
        tipper.Left = Forms("frmWeekView").InsideWidth - 100 - tipper.Width
    End If
    If tb.TOP + tb.Height + tipper.Height < Forms("frmWeekView").InsideHeight Then
        tipper.TOP = tb.TOP + tb.Height
    Else
        tipper.TOP = Forms("frmWeekView").InsideHeight - 100 - tipper.Height
    End If
    If cmrs.count = 0 Then setCmrs
    tipper.value = getCmrs
    tipper.visible = True
Else
    If isFocused Then
        isFocused = False
        Forms("frmWeekView").Controls("cmbWeek").SetFocus
    End If
    tb.SpecialEffect = 0
    tipper.visible = False
End If

Exit_here:
Exit Property

err_trap:
If Err.number = 2165 Then
    Forms("frmWeekView").Controls("cmbWeek").SetFocus
    tb.SpecialEffect = 0
    tipper.visible = False
    Resume Exit_here
End If

End Property

Public Property Get highlighted() As Boolean
highlighted = isHighlighted
End Property

Public Property Get selected() As Boolean
selected = isSelected
End Property

Public Sub init_properties(tNumber As String, tId As Long, tIsFinished As Boolean, ttb As Access.TextBox, ttip As Access.TextBox)
transportNumber = tNumber
transportId = tId
isFinished = tIsFinished
isSelected = False
Set tipper = ttip
Set tb = ttb
tb.OnMouseMove = "[Event Procedure]"
tb.OnClick = "[Event Procedure]"
tb.OnDblClick = "[Event Procedure]"
End Sub

Public Sub tb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not Me.highlighted Then Me.highlighted = True
End Sub

Public Sub tipper_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.highlighted Then Me.highlighted = False
End Sub

Public Sub tb_DblClick(Cancel As Integer)
selectMe False
Forms("frmWeekView").Controls("cmbWeek").SetFocus
openTransport
End Sub


Public Sub tb_Click()

If Not isSelected Then
    selectMe True
Else
    selectMe False
    
End If

End Sub

Public Sub selectMe(bool As Boolean)
isSelected = bool
isFocused = True
If bool Then
    tb.SpecialEffect = 0
    tb.BorderColor = vbYellow
    tb.BorderStyle = 1
    tb.BorderWidth = 3
Else
    tb.BorderColor = 13158600
    tb.BorderStyle = 0
    tb.BorderWidth = 1
End If
End Sub

Private Sub openTransport()
If authorize(getFunctionId("TRANSPORT_PREVIEW"), whoIsLogged) Then
    Call launchForm("frmTransport", Me.ID)
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub displayTip()
Sleep 1000
If Me.highlighted Then
    Call launchForm("frmTipper", "Du<b>pa</b>")
End If

End Sub

Private Sub setCmrs()
Dim rs As ADODB.Recordset
Dim cmr As clsCmr
Dim str As String

On Error GoTo err_trap


str = "SELECT [tbCmr].[cmrId] as id, [tbShipTo].[shipToString] + ' ' + [tbCompanyDetails].[companyCity] + ', ' + [tbCompanyDetails].[companyCountry] AS dest, [tbDeliveryDetail].[deliveryNote] as doc, tbDeliveryDetail.numberPall AS pal, tbDeliveryDetail.weightNet AS Net, tbDeliveryDetail.weightGross AS Gross " _
    & "FROM (((tbTransport LEFT JOIN tbCmr ON tbTransport.transportId = tbCmr.transportId) LEFT JOIN tbDeliveryDetail ON tbCmr.detailId = tbDeliveryDetail.cmrDetailId) LEFT JOIN tbShipTo ON tbDeliveryDetail.shipToId = tbShipTo.shipToId) LEFT JOIN tbCompanyDetails ON tbShipTo.companyId = tbCompanyDetails.companyId " _
    & "WHERE (((tbTransport.transportId)=" & transportId & "));"

Set rs = newRecordset(str)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set cmr = New clsCmr
        If Not IsNull(rs.fields("id")) Then
            cmr.ID = rs.fields("id")
            cmr.destination = rs.fields("dest")
            If Not IsNull(rs.fields("doc")) Then cmr.deliveryNumbers = rs.fields("doc")
            If Not IsNull(rs.fields("pal")) Then cmr.numberOfPallets = rs.fields("pal")
            If Not IsNull(rs.fields("Net")) Then cmr.netWeight = rs.fields("Net")
            If Not IsNull(rs.fields("Gross")) Then cmr.grossWeight = rs.fields("Gross")
            cmrs.Add cmr, CStr(rs.fields("id"))
        End If
        rs.MoveNext
    Loop
    rs.Close
End If

Exit_here:
Set rs = Nothing
Exit Sub

err_trap:
MsgBox "Error in ""setCmrs"" of clsTransportOrder. Error number: " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Private Function getCmrs() As String
Dim c As clsCmr
Dim str As String
Dim i As Integer

tipper.Height = 1600

i = 0

For Each c In cmrs
    str = str & c.toString & "<br>"
    i = i + 1
Next c

str = str & "<b>Numery auta:</b> " & pTruckNumbers
str = str & "<br><b>Spedycja:</b> " & pCarrierString

tipper.Height = tipper.Height * i

getCmrs = str

End Function
