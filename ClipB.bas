Attribute VB_Name = "ClipB"
Option Explicit
Private Const MAXSIZE = 4096
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const CF_TEXT = 1
Private Const CF_LOCALE = 16
Private Const CF_UNICODETEXT As Long = 13
Private Const CF_LINK As Long = 49627
Private Const CF_LINKSOURCE As Long = 49165
#If VBA7 Then
 'https://github.com/ReneNyffenegger/WinAPI-4-VBA/blob/master/Win32API_PtrSafe.txt
 'https://docs.microsoft.com/ru-ru/office/troubleshoot/office-suite-issues/win32api_ptrsafe-with-64-bit-support
 Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
 Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
 'Declare PtrSafe Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
 Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr 'koka
 'Declare PtrSafe Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As LongPtr
 Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr 'koka
 Private Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
 Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
 Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
 
 Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
 Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
 Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
 Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
 
 Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr 'koka
 Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr 'koka
 
 Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
 
 Private Declare PtrSafe Function GetUserDefaultLCID Lib "kernel32" () As Long
#Else
 Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
 Private Declare Function CloseClipboard Lib "user32" () As Long
 Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
 Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
 Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
 Private Declare Function EmptyClipboard Lib "user32" () As Long
 Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
 
 Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
 Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
 Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
 Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
 
 'https://docs.microsoft.com/en-us/office/vba/access/concepts/windows-api/retrieve-information-from-the-clipboard
 Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
 Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
 
 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
#End If

Private Sub SetClipboardUT(sUniText As String) 'https://docs.microsoft.com/en-us/office/vba/access/Concepts/Windows-API/send-information-to-the-clipboard
 'set sUniText as CF_UNICODETEXT and CF_TEXT to ñlipboard
 'set GetKeyboardLayout as CF_LOCALE to ñlipboard
 #If VBA7 Then 'koka
  Dim iStrPtr As LongPtr
  Dim iLen As LongPtr
  Dim iLock As LongPtr
 #Else
  Dim iStrPtr As Long
  Dim iLen As Long
  Dim iLock As Long
 #End If
 On Error GoTo Finally 'koka
Try:
 OpenClipboard 0&
 EmptyClipboard
 iLen = LenB(sUniText) + 2&
 iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
 iLock = GlobalLock(iStrPtr)
 lstrcpyW iLock, StrPtr(sUniText)
 GlobalUnlock iStrPtr
 SetClipboardData CF_UNICODETEXT, iStrPtr
Finally:
 CloseClipboard
End Sub


Sub SetClipboard(sUniText As String, Optional ClipboardFormat As Long = CF_UNICODETEXT, Optional KeepClipboard As Boolean = False)
 'set sUniText as ClipboardFormat to ñlipboard
 #If VBA7 Then 'koka
  Dim iStrPtr As LongPtr
  Dim iLen As LongPtr
  Dim iLock As LongPtr
 #Else
  Dim iStrPtr As Long
  Dim iLen As Long
  Dim iLock As Long
 #End If
 Dim lLCID As Long
 Dim Buffer() As Byte
 On Error GoTo Finally 'koka
Try:
 Select Case ClipboardFormat
 Case CF_UNICODETEXT, CF_TEXT
  SetClipboardUT sUniText
  'set GetUserDefaultLCID as CF_LOCALE to ñlipboard
  lLCID = GetUserDefaultLCID
  SetClipboard Chr(lLCID And &HFF) & Chr((lLCID And &HFF00) \ 256) & String$(2, vbNullChar), CF_LOCALE, True
  Exit Sub
 End Select
 OpenClipboard 0&
 If Not KeepClipboard Then EmptyClipboard
 Buffer = StrConv(sUniText, vbFromUnicode)
 iLen = UBound(Buffer) + 1
 iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
 iLock = GlobalLock(iStrPtr)
 CopyMemory ByVal iLock, Buffer(0), iLen
 GlobalUnlock iStrPtr
 SetClipboardData ClipboardFormat, iStrPtr
Finally:
 CloseClipboard
End Sub

Private Function GetClipboardUT() As String 'https://docs.microsoft.com/en-us/office/vba/access/Concepts/Windows-API/send-information-to-the-clipboard
 'get unicode text as CF_UNICODETEXT from clipboard
 #If VBA7 Then
  Dim iStrPtr As LongPtr
  Dim iLen As LongPtr
  Dim iLock As LongPtr
 #Else
  Dim iStrPtr As Long
  Dim iLen As Long
  Dim iLock As Long
 #End If
 Dim sUniText As String
 On Error GoTo Finally 'koka
Try:
 OpenClipboard 0&
 If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
  iStrPtr = GetClipboardData(CF_UNICODETEXT)
  If iStrPtr Then
   iLock = GlobalLock(iStrPtr)
   iLen = GlobalSize(iStrPtr)
   sUniText = String$(iLen \ 2& - 1&, vbNullChar)
   lstrcpyW StrPtr(sUniText), iLock
   GlobalUnlock iStrPtr
  End If
  GetClipboardUT = sUniText
 End If
Finally:
 CloseClipboard
End Function

Function GetClipboardU(ClipboardFormat As Long) As String
 'get unicode text as ClipboardFormat from clipboard
 #If VBA7 Then
  Dim iStrPtr As LongPtr
  Dim iLen As LongPtr
  Dim iLock As LongPtr
 #Else
  Dim iStrPtr As Long
  Dim iLen As Long
  Dim iLock As Long
 #End If
 Dim sUniText As String
 On Error GoTo Finally
Try:
 OpenClipboard 0&
 If IsClipboardFormatAvailable(ClipboardFormat) Then
  iStrPtr = GetClipboardData(ClipboardFormat)
  If iStrPtr Then
   iLock = GlobalLock(iStrPtr)
   iLen = GlobalSize(iStrPtr)
   If iLen Then
    sUniText = String$(iLen \ 2& - 1&, vbNullChar)
    CopyMemory ByVal StrPtr(sUniText), ByVal iLock, iLen
   End If
   GlobalUnlock iStrPtr
  End If
  GetClipboardU = sUniText
 End If
Finally:
 CloseClipboard
End Function
Function GetClipboard(Optional ClipboardFormat As Long = CF_UNICODETEXT) As String
 'get text as ClipboardFormat from clipboard
 #If VBA7 Then
  Dim iStrPtr As LongPtr
  Dim iLen As LongPtr
  Dim iLock As LongPtr
 #Else
  Dim iStrPtr As Long
  Dim iLen As Long
  Dim iLock As Long
 #End If
 Dim sUniText As String
 Dim Buffer() As Byte
 On Error GoTo Finally
Try:
 If ClipboardFormat = CF_UNICODETEXT Then
  GetClipboard = GetClipboardUT
  Exit Function
 End If
 OpenClipboard 0&
 If IsClipboardFormatAvailable(ClipboardFormat) Then
  iStrPtr = GetClipboardData(ClipboardFormat)
  If iStrPtr Then
   iLock = GlobalLock(iStrPtr)
   iLen = GlobalSize(iStrPtr)
   If iLen Then
    ReDim Buffer(0 To (iLen - 1)) As Byte
    CopyMemory Buffer(0), ByVal iLock, iLen
    sUniText = StrConv(Buffer, vbUnicode)
   End If
   GlobalUnlock iStrPtr
  End If
  GetClipboard = sUniText
 End If
Finally:
 CloseClipboard
End Function

Public Sub EnumClipboard(iL As Long)
 'debug.print CF_ with len=iL
 #If VBA7 Then
  Dim iStrPtr As LongPtr
  Dim iLen As LongPtr
  Dim iLock As LongPtr
 #Else
  Dim iStrPtr As Long
  Dim iLen As Long
  Dim iLock As Long
 #End If
 Dim RetVal As Long
 Dim iCF As Long
 On Error GoTo Finally
Try:
 OpenClipboard 0&
 iCF = 0&
 Do
  iCF = EnumClipboardFormats(iCF)
  If iCF = 0 Then Exit Do
  iStrPtr = GetClipboardData(iCF)
  If Not IsNull(iStrPtr) Then
   iLock = GlobalLock(iStrPtr)
   iLen = GlobalSize(iStrPtr)
   If iLen = iL Then
    Debug.Print iCF
    Debug.Print iLen
   End If
   GlobalUnlock iStrPtr
  End If
 Loop
Finally:
 CloseClipboard
End Sub

Public Function ClipBoard_GetData() 'https://docs.microsoft.com/en-us/office/vba/access/concepts/windows-api/retrieve-information-from-the-clipboard
 'get text as CF_TEXT from clipboard
  #If VBA7 Then 'koka
   Dim hClipMemory As LongPtr
   Dim lpClipMemory As LongPtr
  #Else
   Dim hClipMemory As Long
   Dim lpClipMemory As Long
  #End If
   Dim MyString As String
   Dim RetVal As Long
   
   On Error GoTo OutOfHere 'koka
   
   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If
          
   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If
 
   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)
 
   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MAXSIZE)
      RetVal = lstrcpy(MyString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)
       
      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If
 
OutOfHere:
 
   RetVal = CloseClipboard()
   ClipBoard_GetData = MyString
 
End Function

Function GetClipboardT() As String
 'get text as CF_TEXT from clipboard
 GetClipboardT = GetClipboard(CF_TEXT)
 GetClipboardT = Left(GetClipboardT, InStr(1, GetClipboardT, vbNullChar, 0) - 1)
End Function

Function GetClipboardLink() As Range
 'get range as CF_LINK from clipboard (Excel D:\test\[+^C.xlsb]Ëèñò1 R6C6:R6C7)
 Dim sLink As String
 Dim aLink() As String
 Dim aRange() As String
 Dim aRC() As String
 Dim aRC2() As String
 On Error GoTo Finally
 sLink = GetClipboard(CF_LINK)
 sLink = Replace(sLink, "[", vbNullChar)
 sLink = Replace(sLink, "]", vbNullChar)
 aLink = Split(sLink, vbNullChar)
 If UBound(aLink) <> 6 Then Exit Function
 aRange = Split(aLink(4), ":")
 aRC = Split(Replace(Replace(aRange(0), "R", ""), "C", " "))
 With Application.Workbooks(aLink(2)).Worksheets(aLink(3))
  If UBound(aRange) = 1 Then
   aRC2 = Split(Replace(Replace(aRange(1), "R", ""), "C", " "))
   Set GetClipboardLink = .Range(.Cells(CLng(aRC(0)), CLng(aRC(1))), .Cells(CLng(aRC2(0)), CLng(aRC2(1))))
  Else
   Set GetClipboardLink = .Cells(CLng(aRC(0)), CLng(aRC(1)))
  End If
 End With
Finally:
End Function

