VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***************************************************************************************
' Module:     cImageFX
' DateTime:   06/07/2006 10:21
' Author:     BioHazardMX
' Purpose:    Entry Level Image Effects using DIB's
'
' License:
' Feel free to use this code in your projects. This
' code is free and may be used and/or modified without
' an express autorization of the autor. Please do not
' remove the copyright notices and claim that you wrote
' it. It would be nice (but not required) if you add
' a notice to your about box.
'
' This code contains portions of cDIBSections class
' written by Steve McMahon (www.vbaccelerator.com)
'
' Written by BioHazardMX (biohazardmx.tripod.com)
'
' If you improve the code (adding more effects or
' fixing bugs) please contact me to add your improvements
' to the class.
'
' Hint: To improve the speed, always compile to native
' code. PCode-compiled version will have the same speed
' as in the IDE
'
' Copyright ® 2006 BioHazardMX
'
'***************************************************************************************

Option Explicit

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const HALFTONE As Long = 4

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Private Type SAFEARRAYBOUND
  cElements As Long
  lLbound As Long
End Type

Private Type SAFEARRAY2D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DrawIcon Lib "user32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function GetCursor Lib "user32.dll" () As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Event InitProgress(ByVal lMax As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)

Private lStart As Long, lFinal As Long, lElapsed As Long

Public Sub FxGrayScale(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, SrcDC As Long, SrcX As Long, SrcY As Long, SrcWidth As Long, SrcHeight As Long, Optional MaskColor As Long = -1)

Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long
Dim lB As Long, lG As Long, lR As Long
Dim mtBI As BITMAPINFO, mBytesPerScanLine As Long
Dim lGray As Long

Dim sColor() As Byte
Dim tColor() As Byte
Dim tSA As SAFEARRAY2D
Dim sSA As SAFEARRAY2D
Dim sDC As Long, tDC As Long
Dim sPtr As Long, tPtr As Long
Dim sObj As Long, tObj As Long
Dim sDib As Long, tDib As Long

On Error GoTo ErrorHandler

  lStart = GetTickCount()
  mBytesPerScanLine = (DstWidth * 3 + 3) And &HFFFFFFFC
    
  With mtBI.bmiHeader
    .biSize = Len(mtBI.bmiHeader)
    .biWidth = DstWidth
    .biHeight = DstHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = mBytesPerScanLine * .biHeight
  End With
    
  sDC = CreateCompatibleDC(SrcDC)
  tDC = CreateCompatibleDC(DstDC)
  sDib = CreateDIBSection(SrcDC, mtBI, DIB_RGB_COLORS, sPtr, 0, 0)
  tDib = CreateDIBSection(DstDC, mtBI, DIB_RGB_COLORS, tPtr, 0, 0)
  sObj = SelectObject(sDC, sDib)
  tObj = SelectObject(tDC, tDib)
    
  BitBlt sDC, 0, 0, DstWidth, DstWidth, SrcDC, SrcX, SrcY, vbSrcCopy
  BitBlt tDC, 0, 0, DstWidth, DstHeight, DstDC, DstX, DstY, vbSrcCopy
  
  With sSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = sPtr
  End With
  CopyMemory ByVal VarPtrArray(sColor), VarPtr(sSA), 4
    
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = tPtr
  End With
  CopyMemory ByVal VarPtrArray(tColor), VarPtr(tSA), 4
        
  yMax = DstHeight - 1
  xMax = DstWidth - 1
    
  MaskColor = GetLngColor(MaskColor)
    
  RaiseEvent InitProgress(xMax)
  For X = 0 To (xMax * 3) Step 3
    For Y = 0 To yMax
      lB = sColor(X, Y)
      lG = sColor(X + 1, Y)
      lR = sColor(X + 2, Y)
      If (RGB(lR, lG, lB) <> MaskColor) Then
        lGray = (222 * lR + 707 * lG + 71 * lB) / 1000
        tColor(X, Y) = lGray
        tColor(X + 1, Y) = lGray
        tColor(X + 2, Y) = lGray
      End If
    Next Y
    RaiseEvent Progress(X)
  Next X
    
  BitBlt DstDC, DstX, DstY, DstWidth, DstHeight, tDC, SrcX, SrcY, vbSrcCopy
    
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
  
  GoTo CleanUp
    
Exit Sub
ErrorHandler:

  MsgBox Err.Description, vbCritical, "Error: " & Err.Number
  
Exit Sub
CleanUp:
  
  Erase sColor, tColor
  SelectObject sDC, sObj
  DeleteObject sDib
  SelectObject tDC, tObj
  DeleteObject tDib
  DeleteDC sDC
  DeleteDC tDC
  sDC = 0: sObj = 0: sDib = 0: sPtr = 0
  tDC = 0: tObj = 0: tDib = 0: tPtr = 0

End Sub

Public Sub FxAlphaBlend(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, SrcDC As Long, SrcX As Long, SrcY As Long, SrcWidth As Long, SrcHeight As Long, Alpha As Byte, Optional MaskColor As Long = -1)

Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long, tAlpha As Long
Dim sB As Long, sG As Long, sR As Long
Dim tB As Long, tG As Long, tR As Long
Dim mtBI As BITMAPINFO, mBytesPerScanLine As Long

Dim sColor() As Byte
Dim tColor() As Byte
Dim tSA As SAFEARRAY2D
Dim sSA As SAFEARRAY2D
Dim sDC As Long, tDC As Long
Dim sPtr As Long, tPtr As Long
Dim sObj As Long, tObj As Long
Dim sDib As Long, tDib As Long

On Error GoTo ErrorHandler

  lStart = GetTickCount()
  mBytesPerScanLine = (DstWidth * 3 + 3) And &HFFFFFFFC
    
  With mtBI.bmiHeader
    .biSize = Len(mtBI.bmiHeader)
    .biWidth = DstWidth
    .biHeight = DstHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = mBytesPerScanLine * .biHeight
  End With
    
  sDC = CreateCompatibleDC(SrcDC)
  tDC = CreateCompatibleDC(DstDC)
  sDib = CreateDIBSection(SrcDC, mtBI, DIB_RGB_COLORS, sPtr, 0, 0)
  tDib = CreateDIBSection(DstDC, mtBI, DIB_RGB_COLORS, tPtr, 0, 0)
  sObj = SelectObject(sDC, sDib)
  tObj = SelectObject(tDC, tDib)
    
  BitBlt sDC, 0, 0, DstWidth, DstWidth, SrcDC, SrcX, SrcY, vbSrcCopy
  BitBlt tDC, 0, 0, DstWidth, DstHeight, DstDC, DstX, DstY, vbSrcCopy
  
  With sSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = sPtr
  End With
  CopyMemory ByVal VarPtrArray(sColor), VarPtr(sSA), 4
    
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = tPtr
  End With
  CopyMemory ByVal VarPtrArray(tColor), VarPtr(tSA), 4
        
  yMax = DstHeight - 1
  xMax = DstWidth - 1
  tAlpha = Alpha
    
  MaskColor = GetLngColor(MaskColor)
    
  RaiseEvent InitProgress(xMax)
  For X = 0 To (xMax * 3) Step 3
    For Y = 0 To yMax
      sB = sColor(X, Y)
      sG = sColor(X + 1, Y)
      sR = sColor(X + 2, Y)
      tB = tColor(X, Y)
      tG = tColor(X + 1, Y)
      tR = tColor(X + 2, Y)
      If (RGB(sR, sG, sB) <> MaskColor) Then
        tColor(X, Y) = ((tB * tAlpha) / 255) + ((sB * (255 - tAlpha)) / 255)
        tColor(X + 1, Y) = ((tG * tAlpha) / 255) + ((sG * (255 - tAlpha)) / 255)
        tColor(X + 2, Y) = ((tR * tAlpha) / 255) + ((sR * (255 - tAlpha)) / 255)
      End If
    Next Y
    RaiseEvent Progress(X)
  Next X
    
  BitBlt DstDC, DstX, DstY, DstWidth, DstHeight, tDC, SrcX, SrcY, vbSrcCopy
    
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
  
  GoTo CleanUp
    
Exit Sub
ErrorHandler:

  MsgBox Err.Description, vbCritical, "Error: " & Err.Number
  
Exit Sub
CleanUp:
  
  Erase sColor, tColor
  SelectObject sDC, sObj
  DeleteObject sDib
  SelectObject tDC, tObj
  DeleteObject tDib
  DeleteDC sDC
  DeleteDC tDC
  sDC = 0: sObj = 0: sDib = 0: sPtr = 0
  tDC = 0: tObj = 0: tDib = 0: tPtr = 0

End Sub

Public Sub FxSaturation(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, SrcDC As Long, SrcX As Long, SrcY As Long, SrcWidth As Long, SrcHeight As Long, Value As Byte, Optional MaskColor As Long = -1)

Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long, tAlpha As Long
Dim sB As Long, sG As Long, sR As Long
Dim mtBI As BITMAPINFO, mBytesPerScanLine As Long
Dim lGray As Long

Dim sColor() As Byte
Dim tColor() As Byte
Dim tSA As SAFEARRAY2D
Dim sSA As SAFEARRAY2D
Dim sDC As Long, tDC As Long
Dim sPtr As Long, tPtr As Long
Dim sObj As Long, tObj As Long
Dim sDib As Long, tDib As Long

On Error GoTo ErrorHandler

  lStart = GetTickCount()
  mBytesPerScanLine = (DstWidth * 3 + 3) And &HFFFFFFFC
    
  With mtBI.bmiHeader
    .biSize = Len(mtBI.bmiHeader)
    .biWidth = DstWidth
    .biHeight = DstHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = mBytesPerScanLine * .biHeight
  End With
    
  sDC = CreateCompatibleDC(SrcDC)
  tDC = CreateCompatibleDC(DstDC)
  sDib = CreateDIBSection(SrcDC, mtBI, DIB_RGB_COLORS, sPtr, 0, 0)
  tDib = CreateDIBSection(DstDC, mtBI, DIB_RGB_COLORS, tPtr, 0, 0)
  sObj = SelectObject(sDC, sDib)
  tObj = SelectObject(tDC, tDib)
    
  BitBlt sDC, 0, 0, DstWidth, DstWidth, SrcDC, SrcX, SrcY, vbSrcCopy
  BitBlt tDC, 0, 0, DstWidth, DstHeight, DstDC, DstX, DstY, vbSrcCopy
  
  With sSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = sPtr
  End With
  CopyMemory ByVal VarPtrArray(sColor), VarPtr(sSA), 4
    
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = tPtr
  End With
  CopyMemory ByVal VarPtrArray(tColor), VarPtr(tSA), 4
        
  yMax = DstHeight - 1
  xMax = DstWidth - 1
  tAlpha = Abs(255 - Value)
    
  MaskColor = GetLngColor(MaskColor)
  
  RaiseEvent InitProgress(xMax)
  For X = 0 To (xMax * 3) Step 3
    For Y = 0 To yMax
      sB = sColor(X, Y)
      sG = sColor(X + 1, Y)
      sR = sColor(X + 2, Y)
      If (RGB(sR, sG, sB) <> MaskColor) Then
        lGray = (222 * sR + 707 * sG + 71 * sB) / 1000
        tColor(X, Y) = ((lGray * tAlpha) / 255) + ((sB * (255 - tAlpha)) / 255)
        tColor(X + 1, Y) = ((lGray * tAlpha) / 255) + ((sG * (255 - tAlpha)) / 255)
        tColor(X + 2, Y) = ((lGray * tAlpha) / 255) + ((sR * (255 - tAlpha)) / 255)
      End If
    Next Y
    RaiseEvent Progress(X)
  Next X
    
  BitBlt DstDC, DstX, DstY, DstWidth, DstHeight, tDC, SrcX, SrcY, vbSrcCopy
    
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
  
  GoTo CleanUp
    
Exit Sub
ErrorHandler:

  MsgBox Err.Description, vbCritical, "Error: " & Err.Number
  
Exit Sub
CleanUp:
  
  Erase sColor, tColor
  SelectObject sDC, sObj
  DeleteObject sDib
  SelectObject tDC, tObj
  DeleteObject tDib
  DeleteDC sDC
  DeleteDC tDC
  sDC = 0: sObj = 0: sDib = 0: sPtr = 0
  tDC = 0: tObj = 0: tDib = 0: tPtr = 0

End Sub

Public Sub FxTransBlt(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, SrcDC As Long, SrcX As Long, SrcY As Long, SrcWidth As Long, SrcHeight As Long, Optional MaskColor As Long = -1)

Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long, tAlpha As Long
Dim sB As Long, sG As Long, sR As Long
Dim mtBI As BITMAPINFO, mBytesPerScanLine As Long

Dim sColor() As Byte
Dim tColor() As Byte
Dim tSA As SAFEARRAY2D
Dim sSA As SAFEARRAY2D
Dim sDC As Long, tDC As Long
Dim sPtr As Long, tPtr As Long
Dim sObj As Long, tObj As Long
Dim sDib As Long, tDib As Long

On Error GoTo ErrorHandler

  lStart = GetTickCount()
  mBytesPerScanLine = (DstWidth * 3 + 3) And &HFFFFFFFC
    
  With mtBI.bmiHeader
    .biSize = Len(mtBI.bmiHeader)
    .biWidth = DstWidth
    .biHeight = DstHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = mBytesPerScanLine * .biHeight
  End With
    
  sDC = CreateCompatibleDC(SrcDC)
  tDC = CreateCompatibleDC(DstDC)
  sDib = CreateDIBSection(SrcDC, mtBI, DIB_RGB_COLORS, sPtr, 0, 0)
  tDib = CreateDIBSection(DstDC, mtBI, DIB_RGB_COLORS, tPtr, 0, 0)
  sObj = SelectObject(sDC, sDib)
  tObj = SelectObject(tDC, tDib)
    
  BitBlt sDC, 0, 0, DstWidth, DstWidth, SrcDC, SrcX, SrcY, vbSrcCopy
  BitBlt tDC, 0, 0, DstWidth, DstHeight, DstDC, DstX, DstY, vbSrcCopy
  
  With sSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = sPtr
  End With
  CopyMemory ByVal VarPtrArray(sColor), VarPtr(sSA), 4
    
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = tPtr
  End With
  CopyMemory ByVal VarPtrArray(tColor), VarPtr(tSA), 4
        
  yMax = DstHeight - 1
  xMax = DstWidth - 1
      
  MaskColor = GetLngColor(MaskColor)

  RaiseEvent InitProgress(xMax)
  For X = 0 To (xMax * 3) Step 3
    For Y = 0 To yMax
      sB = sColor(X, Y)
      sG = sColor(X + 1, Y)
      sR = sColor(X + 2, Y)
      If (RGB(sR, sG, sB) <> MaskColor) Then
        tColor(X, Y) = sB
        tColor(X + 1, Y) = sG
        tColor(X + 2, Y) = sR
      End If
    Next Y
    RaiseEvent Progress(X)
  Next X
    
  BitBlt DstDC, DstX, DstY, DstWidth, DstHeight, tDC, SrcX, SrcY, vbSrcCopy
    
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
  
  GoTo CleanUp
    
Exit Sub
ErrorHandler:

  MsgBox Err.Description, vbCritical, "Error: " & Err.Number
  
Exit Sub
CleanUp:
  
  Erase sColor, tColor
  SelectObject sDC, sObj
  DeleteObject sDib
  SelectObject tDC, tObj
  DeleteObject tDib
  DeleteDC sDC
  DeleteDC tDC
  sDC = 0: sObj = 0: sDib = 0: sPtr = 0
  tDC = 0: tObj = 0: tDib = 0: tPtr = 0

End Sub

Public Function FxBitBlt(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, SrcDC As Long, SrcX As Long, SrcY As Long)
  lStart = GetTickCount()
  Call BitBlt(DstDC, DstX, DstY, DstWidth, DstHeight, SrcDC, SrcX, SrcY, vbSrcCopy)
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
End Function

Public Function FxStretchBlt(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, SrcDC As Long, SrcX As Long, SrcY As Long, SrcWidth As Long, SrcHeight As Long)
  lStart = GetTickCount()
  Call SetStretchBltMode(DstDC, HALFTONE)
  Call StretchBlt(DstDC, DstX, DstY, DstWidth, DstHeight, SrcDC, SrcX, SrcY, SrcWidth, SrcHeight, vbSrcCopy)
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
End Function

Public Function FxScreenShot(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, Optional SrcX As Long = 0, Optional SrcY As Long = 0, Optional SrcWidth As Long = 0, Optional SrcHeight As Long = 0, Optional ByVal IncludeCursor As Boolean = False)
  
Dim sWnd As Long, sRect As RECT, sPos As POINTAPI
Dim tDC As Long, sDC As Long, sCur As Long
Dim tDib As Long, tObj As Long
Dim mtBI As BITMAPINFO
  
  lStart = GetTickCount()
  
  sWnd = GetDesktopWindow()
  sDC = GetDC(sWnd)
  Call GetWindowRect(sWnd, sRect)
  If SrcWidth = 0 Then SrcWidth = sRect.Right - SrcX
  If SrcHeight = 0 Then SrcHeight = sRect.Bottom - SrcY
  
  mtBI.bmiHeader.biSize = Len(mtBI.bmiHeader)
  mtBI.bmiHeader.biWidth = sRect.Right - sRect.Left
  mtBI.bmiHeader.biHeight = sRect.Bottom - sRect.Top
  mtBI.bmiHeader.biPlanes = 1
  mtBI.bmiHeader.biBitCount = 24
  
  tDC = CreateCompatibleDC(DstDC)
  tDib = CreateCompatibleBitmap(DstDC, mtBI.bmiHeader.biWidth, mtBI.bmiHeader.biHeight)
  tObj = SelectObject(tDC, tDib)
  
  Call BitBlt(tDC, 0, 0, mtBI.bmiHeader.biWidth, mtBI.bmiHeader.biHeight, sDC, 0, 0, vbSrcCopy)
  Call ReleaseDC(sWnd, sDC)
  
  If IncludeCursor Then
    sCur = GetCursor
    Call GetCursorPos(sPos)
    Call DrawIcon(tDC, sPos.X, sPos.Y, sCur)
    Call DeleteObject(sCur)
  End If
  
  Call SetStretchBltMode(DstDC, HALFTONE)
  Call StretchBlt(DstDC, DstX, DstY, DstWidth, DstHeight, tDC, SrcX, SrcY, SrcWidth, SrcHeight, vbSrcCopy)
  
  Call SelectObject(tDC, tObj)
  Call DeleteDC(tDC)
  
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
  
End Function

Public Sub FxColorFill(ByVal DstDC As Long, DstX As Long, DstY As Long, DstWidth As Long, DstHeight As Long, SrcDC As Long, SrcX As Long, SrcY As Long, SrcWidth As Long, SrcHeight As Long, Value As Byte, Color As Long, Optional MaskColor As Long = -1)

Dim X As Long, Y As Long
Dim xMax As Long, yMax As Long, tAlpha As Long
Dim sB As Long, sG As Long, sR As Long
Dim tB As Long, tG As Long, tR As Long
Dim mtBI As BITMAPINFO, mBytesPerScanLine As Long

Dim sColor() As Byte
Dim tColor() As Byte
Dim tSA As SAFEARRAY2D
Dim sSA As SAFEARRAY2D
Dim sDC As Long, tDC As Long
Dim sPtr As Long, tPtr As Long
Dim sObj As Long, tObj As Long
Dim sDib As Long, tDib As Long

On Error GoTo ErrorHandler

  lStart = GetTickCount()
  mBytesPerScanLine = (DstWidth * 3 + 3) And &HFFFFFFFC
    
  With mtBI.bmiHeader
    .biSize = Len(mtBI.bmiHeader)
    .biWidth = DstWidth
    .biHeight = DstHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = mBytesPerScanLine * .biHeight
  End With
    
  sDC = CreateCompatibleDC(SrcDC)
  tDC = CreateCompatibleDC(DstDC)
  sDib = CreateDIBSection(SrcDC, mtBI, DIB_RGB_COLORS, sPtr, 0, 0)
  tDib = CreateDIBSection(DstDC, mtBI, DIB_RGB_COLORS, tPtr, 0, 0)
  sObj = SelectObject(sDC, sDib)
  tObj = SelectObject(tDC, tDib)
    
  BitBlt sDC, 0, 0, DstWidth, DstWidth, SrcDC, SrcX, SrcY, vbSrcCopy
  BitBlt tDC, 0, 0, DstWidth, DstHeight, DstDC, DstX, DstY, vbSrcCopy
  
  With sSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = sPtr
  End With
  CopyMemory ByVal VarPtrArray(sColor), VarPtr(sSA), 4
    
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DstHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = mBytesPerScanLine
    .pvData = tPtr
  End With
  CopyMemory ByVal VarPtrArray(tColor), VarPtr(tSA), 4
        
  yMax = DstHeight - 1
  xMax = DstWidth - 1
  tAlpha = Value
  
  Color = GetLngColor(Color)
  MaskColor = GetLngColor(MaskColor)
    
  RaiseEvent InitProgress(xMax)
  For X = 0 To (xMax * 3) Step 3
    For Y = 0 To yMax
      sB = sColor(X, Y)
      sG = sColor(X + 1, Y)
      sR = sColor(X + 2, Y)
      tB = (Color And &HFF0000) \ &H10000
      tG = (Color And &HFF00&) \ &H100
      tR = (Color And &HFF&)
      If (RGB(sR, sG, sB) <> MaskColor) Then
        tColor(X, Y) = ((tB * tAlpha) / 255) + ((sB * (255 - tAlpha)) / 255)
        tColor(X + 1, Y) = ((tG * tAlpha) / 255) + ((sG * (255 - tAlpha)) / 255)
        tColor(X + 2, Y) = ((tR * tAlpha) / 255) + ((sR * (255 - tAlpha)) / 255)
      End If
    Next Y
    RaiseEvent Progress(X)
  Next X
    
  BitBlt DstDC, DstX, DstY, DstWidth, DstHeight, tDC, SrcX, SrcY, vbSrcCopy
    
  lFinal = GetTickCount()
  RaiseEvent Complete(lFinal - lStart)
  
  GoTo CleanUp
    
Exit Sub
ErrorHandler:

  MsgBox Err.Description, vbCritical, "Error: " & Err.Number
  
Exit Sub
CleanUp:
  
  Erase sColor, tColor
  SelectObject sDC, sObj
  DeleteObject sDib
  SelectObject tDC, tObj
  DeleteObject tDib
  DeleteDC sDC
  DeleteDC tDC
  sDC = 0: sObj = 0: sDib = 0: sPtr = 0
  tDC = 0: tObj = 0: tDib = 0: tPtr = 0

End Sub

Private Function GetLngColor(Color As Long) As Long
If (Color And &H80000000) Then
GetLngColor = GetSysColor(Color And &H7FFFFFFF)
Else
GetLngColor = Color
End If
End Function
