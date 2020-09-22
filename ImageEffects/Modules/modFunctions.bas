Attribute VB_Name = "modFunctions"
'***************************************************************************************
' Module:     modFunctions
' DateTime:   06/07/2006 13:33
' Author:     BioHazardMX
' Purpose:    Generic Helper Module
'***************************************************************************************

Private Type OSVERSIONINFO
  dwVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(0 To 127) As Byte
End Type

Private Const SC_CLOSE As Long = &HF060&
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_MINIMIZE As Long = &HF020&
Private Const SC_MOVE As Long = &HF010&
Private Const SC_SEPARATOR As Long = &HF00F&
Private Const SC_SIZE As Long = &HF000&
Private Const MF_REMOVE = &H1000&
Private Const MF_BYPOSITION = &H400&

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Public Sub DisableClose(ByVal hWnd As Long)
  Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, False)
  DeleteMenu hMenu, SC_CLOSE, MF_BYCOMMAND
  DeleteMenu hMenu, 5, MF_BYPOSITION
End Sub

Public Sub DisableResize(ByVal hWnd As Long)
  Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, 0)
  DeleteMenu hMenu, SC_SIZE, MF_BYCOMMAND
End Sub

Public Function IsXPOrAbove() As Boolean
  Dim OSVer As OSVERSIONINFO
  OSVer.dwVersionInfoSize = Len(OSVer)
  GetVersionEx OSVer
  If (OSVer.dwMajorVersion > 5) Then
    IsXPOrAbove = True
    ElseIf (OSVer.dwMajorVersion = 5) Then
      If (OSVer.dwMinorVersion >= 1) Then
        IsXPOrAbove = True
      End If
  End If
End Function

