Attribute VB_Name = "WinApi"
Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
  ByRef lpPoint As POINT) As Long ' returns a BOOL

Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
  ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long ' returns a BOOL

Private Declare PtrSafe Function GetDC Lib "user32" ( _
  ByVal hwnd As LongPtr) As LongPtr ' returns a HDC - Handle to a Device Context

Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
  ByVal hDC As LongPtr, ByVal nIndex As Long) As Long ' returns a C/C++ int

Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
  ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long ' also returns an int

Private Const LOGPIXELSX = 88 ' sticking to the original names is less confusing IMO
Private Const LOGPIXELSY = 90 ' ditto

Private Const TwipsPerInch = 1440

Private Declare Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hwnd As Long, lpRect As typRect) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As typRect) As Long
Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' Type declarations:
Private Type typRect
    Left As Long
    TOP As Long
    Right As Long
    Bottom As Long
End Type
' Constant declarations:
Private Const SW_RESTORE = 9
Private Const SWP_NOSIZE = &H1 ' Don't alter the size
Private Const SWP_NOZORDER = &H4 ' Don't change the Z-order
Private Const SWP_SHOWWINDOW = &H40 ' Display the window


Type POINT
  x As Long
  y As Long
End Type

Type RECT
  Left As Long
  TOP As Long
  Right As Long
  Bottom As Long
End Type

' **************************************
' * Center a form in the Access window *
' **************************************
' Created 1-22-2002 by Peter M. Schroeder
Public Function centerForm(parForm As Form) As Boolean
    Dim varAccess As typRect, varForm As typRect
    Dim varX As Long, varY As Long
    
    On Error GoTo CenterForm_Error
    Call apiGetClientRect(hWndAccessApp, varAccess) ' Get the Access client area coordinate
    Call apiGetWindowRect(parForm.hwnd, varForm) ' Get the form window coordinates
    varX = CLng((varAccess.Left + varAccess.Right) / 2) - CLng((varForm.Right - varForm.Left) / 2) ' Calculate a new left for the form
    varY = CLng((varAccess.TOP + varAccess.Bottom) / 2) - CLng((varForm.Bottom - varForm.TOP) / 2) ' Calculate a new top for the form
    varY = varY - 45 ' Adjust top for true center
    varY = varY - 20 ' Adjust top for appearance
    Call apiShowWindow(parForm.hwnd, SW_RESTORE) ' Restore form window
    Call apiSetWindowPos(parForm.hwnd, 0, varX, varY, (varForm.Right - varForm.Left), (varForm.Bottom - varForm.TOP), SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSIZE) ' Set new form coordinates
    centerForm = True
    Exit Function
    
CenterForm_Error:
    centerForm = False
End Function

Function PixelsToTwips(ByVal x As Long, ByVal y As Long) As POINT
  Dim ScreenDC As LongPtr
  ScreenDC = GetDC(0)
  PixelsToTwips.x = x / GetDeviceCaps(ScreenDC, LOGPIXELSX) * TwipsPerInch
  PixelsToTwips.y = y / GetDeviceCaps(ScreenDC, LOGPIXELSY) * TwipsPerInch
  ReleaseDC 0, ScreenDC
End Function

Function TwipsToPixels(ByVal x As Long, ByVal y As Long) As POINT
  Dim ScreenDC As LongPtr
  ScreenDC = GetDC(0)
  TwipsToPixels.x = x / TwipsPerInch * GetDeviceCaps(ScreenDC, LOGPIXELSX)
  TwipsToPixels.y = y / TwipsPerInch * GetDeviceCaps(ScreenDC, LOGPIXELSY)
  ReleaseDC 0, ScreenDC
End Function

Sub MoveFormToScreenPixelPos(Form As Access.Form, PixelX As Long, PixelY As Long)
  Dim FormWR As RECT, AccessWR As RECT, Offset As POINT, NewPos As POINT
  ' firstly need to calculate what the coords passed to Move are relative to
  GetWindowRect Application.hWndAccessApp, AccessWR
  GetWindowRect Form.hwnd, FormWR
  Offset = PixelsToTwips(FormWR.Left - AccessWR.Left, FormWR.TOP - AccessWR.TOP)
  Offset.x = Offset.x - Form.WindowLeft
  Offset.y = Offset.y - Form.WindowTop
  ' next convert our desired position to twips and set it
  NewPos = PixelsToTwips(PixelX - AccessWR.Left, PixelY - AccessWR.TOP)
  Form.Move NewPos.x - Offset.x, NewPos.y - Offset.y
End Sub

Sub MoveFormToCursorPos(Form As Access.Form)
  Dim Pos As POINT
  GetCursorPos Pos
  MoveFormToScreenPixelPos Form, Pos.x, Pos.y
End Sub

