Attribute VB_Name = "Module1"

Public Const CF_UNICODETEXT As Long = 13&
Public Const CF_TEXT As Long = 1&
Public Const CF_BITMAP = 2

Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MOVEABLE = &H2
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long


Public Type rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Type PAINTSTRUCT
        hDC As Long
        fErase As Long
        rcPaint As rect
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type


Public Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long


Public Function WriteBitmapToFile(ByVal sPath As String)

    Dim mBitmap As Long
    OpenClipboard 0
    mGdip.InitGDIPlus
    mGdip.GdipCreateBitmapFromHBITMAP GetImageClipBoard, 0, mBitmap
    mGdip.SaveImageToPNG mBitmap, sPath
    CloseClipboard
End Function


Function GetImageClipBoard() As Long
    Dim hClipBoard As Long
    Dim hBitmap As Long
    hBitmap = GetClipboardData(2)
    If hBitmap = 0 Then GoTo exit_error
        GetImageClipBoard = hBitmap
        Exit Function
exit_error:
    GetImageClipBoard = -1
End Function


Function GetTextClipboard()
Dim hTxtPtr As Long
Dim hDataPtr As Long
Dim sClipboardText As String

Dim iCliboardSize As Long
Dim bTextData() As Byte
If (OpenClipboard(0)) Then
    If (IsClipboardFormatAvailable(CF_TEXT)) Then
        hTxtPtr = GetClipboardData(CF_TEXT)
        Call CopyMemory(hDataPtr, ByVal hTxtPtr, &H4)
        iCliboardSize = lstrlen(hTxtPtr)
        If iCliboardSize > 0 Then
            ReDim bTextData(0 To CLng(iCliboardSize) - CLng(1)) As Byte
            CopyMemory bTextData(0), ByVal GlobalLock(hTxtPtr), iCliboardSize
            sClipboardText = StrConv(bTextData, vbUnicode)
        Else
            MsgBox "无数据", vbOKOnly, "提示"
            MsgBox GetClipBoard
            
        End If
    End If
Call CloseClipboard
End If

GetTextClipboard = sClipboardText
End Function
 
 


Public Function IsString() As Boolean
    OpenClipboard 0&
    IsString = (IsClipboardFormatAvailable(CF_UNICODETEXT)) Or (IsClipboardFormatAvailable(CF_TEXT))
    CloseClipboard
End Function

