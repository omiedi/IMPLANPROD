Attribute VB_Name = "WEB_CAM"
Option Explicit
' declaraciones Api, constantes, variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const ws_child = &H40000000
Public Const ws_visible = &H10000000
Public Const WM_USER = 1024
Public Const wm_cap_driver_connect = WM_USER + 10
Public Const wm_cap_set_preview = WM_USER + 50
Public Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Public Const WM_CAP_DRIVER_DISCONNECT = WM_USER + 11
Public Const WM_CAP_DLG_VIDEOFORMAT = WM_USER + 41
Public Const WM_CAP_DLG_VIDEOCONFIG = WM_USER + 42
Public Const WM_CAP_SET_SCALE = WM_USER + 53

'Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER
Public Const WM_CAP_EDIT_COPY As Long = WM_CAP_START + 30
Public Const WM_CAP_SET_VIDEOFORMAT As Long = WM_USER + 45
  
'Api para crear la ventana de captura
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" _
    Alias "capCreateCaptureWindowA" ( _
    ByVal lpszWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hwndParent As Long, _
    ByVal nID As Long) As Long
  
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
  
Public Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
'Solo 16 Bits (vb2, vb3 y vb4  de 16)
'Declare Function SendMessage Lib "User" ( _
    ByVal hWnd As Integer, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
'Api para crear la ventana de captura
'Declare Function capCreateCaptureWindow Lib "avicap.dll" ( _
    ByVal lpszWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Integer, _
    ByVal y As Integer, _
    ByVal nWidth As Integer, _
    ByVal nHeight As Integer, _
    ByVal hwndParent As Integer, _
    ByVal nID As Integer) As Long
'Declare Function DestroyWindow Lib "User" (ByVal hndw As Integer) As Integer
Public Declare Function SendMessageAsBitMap Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As BITMAPINFO) As Long

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As Integer
End Type
    
Public hwdc As Long
Public startcap As Integer

