Attribute VB_Name = "WinApi2"
Option Compare Database
Option Explicit

' Declaration for the DeviceCapabilities function API call.
Private Declare Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, _
    ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
    ByVal lpDevMode As Long) As Long
    
' DeviceCapabilities function constants.
Private Const DC_PAPERNAMES = 16
Private Const DC_PAPERS = 2
Private Const DC_BINNAMES = 12
Private Const DC_BINS = 6
Private Const DEFAULT_VALUES = 0

Public Function GetBinList(strName As String) As String
' Uses the DeviceCapabilities API function to display a
' message box with the name of the default printer and a
' list of the paper bins it supports.

    Dim lngBinCount As Long
    Dim lngCounter As Long
    Dim hPrinter As Long
    Dim strDeviceName As String
    Dim strDevicePort As String
    Dim strBinNamesList As String
    Dim strBinName As String
    Dim intLength As Integer
    Dim strMsg As String
    Dim aintNumBin() As Integer
    
    On Error GoTo GetBinList_Err
    
    ' Get name and port of the default printer.
    strDeviceName = Application.Printers(strName).DeviceName
    strDevicePort = Application.Printers(strName).Port
    
    ' Get count of paper bin names supported by the printer.
    lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_BINNAMES, _
        lpOutput:=ByVal vbNullString, _
        lpDevMode:=DEFAULT_VALUES)
    
    ' Re-dimension the array to count of paper bins.
    If lngBinCount > 0 Then
        ReDim aintNumBin(1 To lngBinCount)
        
        ' Pad variable to accept 24 bytes for each bin name.
        strBinNamesList = String(number:=24 * lngBinCount, Character:=0)
    
        ' Get string buffer of paper bin names supported by the printer.
        lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
            lpPort:=strDevicePort, _
            iIndex:=DC_BINNAMES, _
            lpOutput:=ByVal strBinNamesList, _
            lpDevMode:=DEFAULT_VALUES)
            
        ' Get array of paper bin numbers supported by the printer.
        lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
            lpPort:=strDevicePort, _
            iIndex:=DC_BINS, _
            lpOutput:=aintNumBin(1), _
            lpDevMode:=0)
            
        ' List available paper bin names.
        strMsg = ""
        For lngCounter = 1 To lngBinCount
            
            ' Parse a paper bin name from string buffer.
            strBinName = Mid(String:=strBinNamesList, _
                start:=24 * (lngCounter - 1) + 1, _
                Length:=24)
            intLength = VBA.InStr(start:=1, _
                String1:=strBinName, String2:=Chr(0)) - 1
            strBinName = Left(String:=strBinName, _
                    Length:=intLength)
    
            ' Add bin name and number to text string for message box.
            strMsg = strMsg & vbCrLf & aintNumBin(lngCounter) _
                & vbTab & strBinName
                
        Next lngCounter
    End If
    GetBinList = strMsg
    ' Show paper bin numbers and names in message box.
    
GetBinList_End:
    Exit Function
GetBinList_Err:
    MsgBox Prompt:=Err.description, Buttons:=vbCritical & vbOKOnly, _
        Title:="Error Number " & Err.number & " Occurred"
    Resume GetBinList_End
End Function

Public Sub deleteCmr(cmr As Long)
Dim detailId As Long

detailId = DLookup("detailId", "tbCmr", "cmrId=" & cmr)
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbCustomVars WHERE CmrId = " & cmr
DoCmd.RunSQL "DELETE * FROM tbCmr WHERE cmrId = " & cmr
DoCmd.RunSQL "DELETE * FROM tbDeliveryDetail WHERE cmrDetailId = " & detailId
DoCmd.SetWarnings True
End Sub
