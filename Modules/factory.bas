Attribute VB_Name = "factory"
Option Compare Database
Option Explicit

Public Function CreateSearch(mForm As Access.Form, sForm As Access.subForm, ctrl As Access.TextBox, searchO As String, Optional ruleOut As Variant) As search

On Error GoTo err_trap

    Set CreateSearch = New search
    If Not IsMissing(ruleOut) Then
        CreateSearch.init_properties mainForm:=mForm, subForm:=sForm, searchTxt:=ctrl, sObject:=searchO, exclude:=ruleOut
    Else
        CreateSearch.init_properties mainForm:=mForm, subForm:=sForm, searchTxt:=ctrl, sObject:=searchO
    End If
    
Exit_here:
   Exit Function
   
err_trap:
    MsgBox "Error in FuN ""CreateSearch"" of factory." & vbNewLine & "Err no. " & Err.number & ", description: " & Err.description
    Resume Exit_here
    
End Function

Public Function CreateTransportOrder(tNumber As String, tId As Long, tIsFinished As Boolean, ttb As Access.TextBox, tip As Access.TextBox) As clsTransportOrder
On Error GoTo err_trap

    Set CreateTransportOrder = New clsTransportOrder
    
    CreateTransportOrder.init_properties tNumber:=tNumber, tId:=tId, tIsFinished:=tIsFinished, ttb:=ttb, ttip:=tip
   
    
Exit_here:
   Exit Function
   
err_trap:
    MsgBox "Error in FuN ""CreateTransportOrder"" of factory." & vbNewLine & "Err no. " & Err.number & ", description: " & Err.description
    Resume Exit_here
    
End Function


Public Function CreateClsListener(ctrl As Access.Control) As clsListener

On Error GoTo err_trap

    Set CreateClsListener = New clsListener
    CreateClsListener.init_properties ctrl:=ctrl
    
    
Exit_here:
   Exit Function
   
err_trap:
    MsgBox "Error in FuN ""CreateClsListener"" of factory." & vbNewLine & "Err no. " & Err.number & ", description: " & Err.description
    Resume Exit_here
    
End Function

Public Function CreatePowerSearch(srchTxt As Access.TextBox, srchQuery As String, ret As String, Optional ruleOut As Variant, Optional columnWidths As Variant) As clsPowerSearch

On Error GoTo err_trap

    Set CreatePowerSearch = New clsPowerSearch
    If Not IsMissing(ruleOut) Then
        If IsMissing(columnWidths) Then
            CreatePowerSearch.init_properties searchTxt:=srchTxt, searchQuery:=srchQuery, retField:=ret, exclude:=ruleOut
        Else
            CreatePowerSearch.init_properties searchTxt:=srchTxt, searchQuery:=srchQuery, retField:=ret, exclude:=ruleOut, colWidths:=columnWidths
        End If
    Else
        If IsMissing(columnWidths) Then
            CreatePowerSearch.init_properties searchTxt:=srchTxt, searchQuery:=srchQuery, retField:=ret
        Else
            CreatePowerSearch.init_properties searchTxt:=srchTxt, searchQuery:=srchQuery, retField:=ret, colWidths:=columnWidths
        End If
    End If
    
Exit_here:
   Exit Function
   
err_trap:
    MsgBox "Error in FuN ""CreatePowerSearch"" of factory." & vbNewLine & "Err no. " & Err.number & ", description: " & Err.description
    Resume Exit_here
    
End Function

Public Function CreateClsMultiSelect(mForm As Access.Form) As clsMultiSelect

On Error GoTo err_trap

    Set CreateClsMultiSelect = New clsMultiSelect
    CreateClsMultiSelect.init_properties mForm

    
Exit_here:
   Exit Function
   
err_trap:
    MsgBox "Error in FuN ""CreateClsMultiSelect"" of factory." & vbNewLine & "Err no. " & Err.number & ", description: " & Err.description
    Resume Exit_here
    
End Function
