VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private pName As String
Private pValue As Variant

Public Property Get Name() As String
Name = pName
End Property

Public Property Get value() As Variant
value = pValue
End Property

Public Property Let Name(val As String)
pName = val
End Property

Public Property Let value(val As Variant)
pValue = val
End Property
