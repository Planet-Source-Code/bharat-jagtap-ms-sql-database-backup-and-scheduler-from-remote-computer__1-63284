Attribute VB_Name = "ShellModule"
'for system tray icon
Public Declare Function Shell_NotifyIcon Lib _
"shell32.dll" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Global NotifyIcon As NOTIFYICONDATA

Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200




