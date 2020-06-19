VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_xxx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Dim arr() As Variant
ReDim arr(2, 4) As Variant
arr(0, 0) = "Ten miesiąc"
arr(1, 0) = DateSerial(year(Date), Month(Date), 1)
arr(2, 0) = DateAdd("m", 1, DateSerial(year(Date), Month(Date), 0))
arr(0, 1) = "Poprzedni miesiąc"
arr(1, 1) = DateSerial(year(Date), Month(DateAdd("m", -1, Date)), 1)
arr(2, 1) = DateSerial(year(Date), Month(Date), 0)
arr(0, 2) = "Poprzedni kwartał"
arr(1, 2) = DateAdd("m", -3, Date)
arr(2, 2) = Date
arr(0, 3) = "Poprzednie pół roku"
arr(1, 3) = DateAdd("m", -6, Date)
arr(2, 3) = Date
arr(0, 4) = "Poprzedni rok"
arr(1, 4) = DateAdd("m", -12, Date)
arr(2, 4) = Date
Call populateCombo(Me.cmbX, arr, "Ten miesiąc", 3.5, 2)
End Sub
