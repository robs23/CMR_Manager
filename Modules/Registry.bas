Attribute VB_Name = "Registry"
Option Compare Database


Public Sub updateRegistry(key As String, value As Variant)
'key in form e.g. "HKEY_CURRENT_USER\Software\RWsoft\backEndPath"
Dim reg As Object 'registry itself
Dim theType As Variant
Dim regPath As String

On Error GoTo err_trap

regPath = "HKEY_CURRENT_USER\Software\cmr_manager\"
key = regPath & key

Select Case VarType(value)
    Case 0 To 1
    theType = Null
    Case 2
    theType = "REG_DWORD"
    Case 3
    theType = "REG_QWORD"
    Case 7
    value = CLng(value)
    theType = "REG_QWORD"
    Case 8
    theType = "REG_SZ"
    Case 11
    theType = "REG_BINARY"
    Case Else
    theType = Null
End Select

If theType = Null Then
    MsgBox "Type of variable ""value"" passed to ""createRegistryKey"" could not be determined or is unsuported. No key has been created", vbOKOnly + vbExclamation
Else
    Set reg = CreateObject("WScript.Shell")
    reg.RegWrite key, value, theType
End If


exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""updateRegistry"". Error number: " & Err.number & ", " & Err.description
Resume exit_here

End Sub

Public Function registryKeyExists(key As String) As Variant
Dim bool As Variant
Dim reg As Variant
Dim regPath As String

On Error GoTo err_trap

bool = False
regPath = "HKEY_CURRENT_USER\Software\cmr_manager\"

key = regPath & key

reg = CreateObject("WScript.Shell").RegRead(key)

bool = reg

exit_here:
registryKeyExists = bool
Exit Function

err_trap:
Resume exit_here

End Function



