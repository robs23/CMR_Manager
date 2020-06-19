Attribute VB_Name = "Global vars"
Option Compare Database

Public selCompany As Long
Public selCompanyType As Integer
Public selContact As Long
Public PUB_VARS(10) As String
Public currentBEPath As String 'ścieżka do back end
Public networkBEPath As String 'ścieżka do back end na sieci
Public currentBEName As String 'nazwa back end
Public adoConn As ADODB.Connection 'connection to sql server
Public backEndPass As String 'hasło back end
Public vbaPassword As String 'haslo do projektu vba
Public currentCmr As clsCmr

Public Type delivery
    KLIENT_ID As Long
    MAGAZYN_ID As Long
    DATA As Variant
    DELIVERY_NOTE As String
    ILOSC_PALET As Integer
    WAGA_N As Double
    WAGA_B As Double
    PRZEWOZNIK_ID As Long
    GRANICA_WJAZD As String
    GRANICA_WYJAZD As String
    PRZEWOZNIK_KONTAKT_ID As Long
    PRZEZ_NIEMCY As Boolean
    NUMERY_REJESTRACYJNE As String
End Type

Public currentDelivery As delivery



Public Sub Load_PUB_VARS()
PUB_VARS(0) = "KLIENT"
PUB_VARS(1) = "MAGAZYN"
PUB_VARS(2) = "DATA"
PUB_VARS(3) = "DELIVERY_NOTE"
PUB_VARS(4) = "ILOSC_PALET"
PUB_VARS(5) = "WAGA_N"
PUB_VARS(6) = "WAGA_B"
PUB_VARS(7) = "PRZEWOZNIK"
PUB_VARS(8) = "NR_TRANSPORTU"
PUB_VARS(9) = "SPEDYTOR"
PUB_VARS(10) = "NUMERY_AUTA"
End Sub

