'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' paste.bas
' V1.0 2013-01-18 by Tim Gremalm 
' ==============================
' Makes sure you won't paste multiplpe lines by misstake into your terminal
' 
' To install this script:
' 	* Place in directory %ProgramFiles%\AniTa
'	* AniTa's menu Settings->Options->Mouse
'		In textfield Right button click, Add:
'		%copy%%escript%paste.bas
'	* AniTa's menu Settings->Options->Keyboard
'		In textfield Alt under V, Add:
'		%escript%paste.bas
' 
' Requirements
'	* AniTa Terminal emulator
'	  http://www.april.se/english/ftp.asp#AniTa
' 
' Tim Gremalm
' Conmel Data AB
' tim@gremalm.se
' 0735-444 293
' 
' Copyright (c) 2013 Tim Gremalm
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' 
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Clipboard Manager Functions
Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long

' Other APIs
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long

Declare Function GetLastError Lib "kernel32" () As Long
Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1
Const ERROR_BAD_USERNAME = 2202&

Const GMEM_FIXED                    As Long = &H0

Const MB_OK							As Long = 0 'Display OK button only.
Const MB_OKCANCEL					As Long = 1 'Display OK and Cancel buttons.
Const MB_ABORTRETRYIGNORE			As Long = 2 'Display Abort, Retry, and Ignore buttons.
Const MB_YESNOCANCEL				As Long = 3 'Display Yes, No, and Cancel buttons.
Const MB_YESNO						As Long = 4 'Display Yes and No buttons.
Const MB_RETRYCANCEL				As Long = 5 'Display Retry and Cancel buttons.

Const MB_ICONSTOP 					As Long = 16
Const MB_ICONQUESTION				As Long = 32
Const MB_ICONEXCLAMATION			As Long = 48
Const MB_ICONINFORMATION			As Long = 64

Const MB_DEFBUTTON1					As Long = 0 'First button is default.
Const MB_DEFBUTTON2					As Long = 256 'Second button is default.
Const MB_DEFBUTTON3					As Long = 512 'Third button is default.

Const MB_APPLMODAL					As Long = 0 'Application modal. The user must respond to the message box before continuing work in the current application 
Const MB_SYSTEMMODAL				As Long = 4096 'System modal. All applications are suspended until the user responds to the message box.

Const IDOK							As Long = 1 'OK button selected.
Const IDCANCEL						As Long = 2 'Cancel button selected.
Const IDABORT						As Long = 3 'Abort button selected.
Const IDRETRY						As Long = 4 'Retry button selected.
Const IDIGNORE						As Long = 5 'Ignore button selected.
Const IDYES							As Long = 6 'Yes button selected.
Const IDNO							As Long = 7 'No button selected.

' Predefined Clipboard Formats
Const CF_TEXT                       As Long = 1
Const CF_BITMAP                     As Long = 2
Const CF_METAFILEPICT               As Long = 3
Const CF_SYLK                       As Long = 4
Const CF_DIF                        As Long = 5
Const CF_TIFF                       As Long = 6
Const CF_OEMTEXT                    As Long = 7
Const CF_DIB                        As Long = 8
Const CF_PALETTE                    As Long = 9
Const CF_PENDATA                    As Long = 10
Const CF_RIFF                       As Long = 11
Const CF_WAVE                       As Long = 12
Const CF_UNICODETEXT                As Long = 13
Const CF_ENHMETAFILE                As Long = 14
Const CF_HDROP                      As Long = 15
Const CF_LOCALE                     As Long = 16
Const CF_MAX                        As Long = 17
Const CF_OWNERDISPLAY               As Long = &H80
Const CF_DSPTEXT                    As Long = &H81
Const CF_DSPBITMAP                  As Long = &H82
Const CF_DSPMETAFILEPICT            As Long = &H83
Const CF_DSPENHMETAFILE             As Long = &H8E

Sub main()
On Error Goto ErrHandler:
	Dim vbNewline As String
	Dim vbCarriagereturn As String
	Dim vbCarriagereturnNewline As String
	vbNewline = Chr(10)
	vbCarriagereturn = Chr(13)
	vbCarriagereturnNewline = vbCarriagereturn & vbNewline
	
	Dim sError As String
	Dim sClipboard As String
	sError = PasteTextFromClipboard(sClipboard)
	If Len(sError) > 0 Then
		MsgBox "Error!" & vbNewline & "======" & vbNewline & sError & vbNewline & "Exiting paste.bas"
		Err.Clear
		End
	End If
	
	'Dim iRow As Integer
	'iRow = 0
	'Dim sShow As String
	'Dim sResultarray() As String
	'Split sClipboard, vbCarriagereturnNewline, sResultArray
	'sShow = "Row-Count: " & (UBound(sResultArray)-LBound(sResultArray)+1) 
	'For Each s In sResultArray
	'	sShow = sShow & vbNewline & s
	'	iRow = iRow + 1
	'Next s
	'MsgBox sShow
	
	Dim sResultarray() As String
	Dim bDirtyCRLF As Boolean
	Dim bDirtyCR As Boolean
	Dim bDirtyLF As Boolean
	
	'Look for CRLF
	Split sClipboard, vbCarriagereturnNewline, sResultArray
	If (UBound(sResultArray)-LBound(sResultArray)+1) <> 1 Then
		bDirtyCRLF = True
	End If
	
	'Look for CR
	Split sClipboard, vbCarriagereturn, sResultArray
	If (UBound(sResultArray)-LBound(sResultArray)+1) <> 1 Then
		bDirtyCR = True
	End If
	
	'Look for LF
	Split sClipboard, vbNewline, sResultArray
	If (UBound(sResultArray)-LBound(sResultArray)+1) <> 1 Then
		bDirtyLF = True
	End If
	
	Dim bPaste As Boolean
	If bDirtyCR Or bDirtyCR Or bDirtyLF Then
		Dim lPasteAnyway As Long
		lPasteAnyway = MsgBox("Clipboard contains more than one row!" & vbNewline & vbNewline & "Do you want to paste anyway?" , MB_YESNO Or MB_DEFBUTTON2)
		If lPasteAnyway = IDYES Then
			bPaste = True
		Else
			bPaste = False
		End If
	Else
		bPaste = True
	End If
	
	If bPaste Then
		'MsgBox sClipboard
		AniMacro "%paste%"
	Else
		'MsgBox "I won't paste!"
	End If
	
	Exit Sub
	
ErrHandler:
    'Error handling
	MsgBox "An error occurd in main()!" & vbNewline & "Error-number: " & Err.Number & vbNewline & Err.Description
	End
End Sub

Sub Split (ByVal sInput As String, ByVal sSeparator As String, ByRef sResultArray() As String)
On Error Goto ErrHandler:
	Dim sOutputArray() As String
	Dim sTemp As String
	Dim iFound As Integer
	sTemp = sInput
	
	'MsgBox Left(sTemp, InStr(1, sTemp, sSeparator) - 1)
	'MsgBox Right(sTemp, Len(sTemp) - InStr(1, sTemp, sSeparator))
	
	Do Until InStr(1, sTemp, sSeparator) = 0
		ReDim Preserve sOutputArray(0 To iFound)
		sOutputArray(iFound) = Left(sTemp, InStr(1, sTemp, sSeparator) - 1)
		'MsgBox iFound & " " & sOutputArray(iFound)
		sTemp = Right(sTemp, Len(sTemp) - InStr(1, sTemp, sSeparator))
		iFound = iFound + 1
	Loop
	ReDim Preserve sOutputArray(0 To iFound)
	sOutputArray(iFound) = sTemp
	'MsgBox sOutputArray(0)
	
	sResultArray = sOutputArray
	Exit Sub
	
ErrHandler:
    'Error handling
	MsgBox "An error occurd in Split()!" & vbNewline & "Error-number: " & Err.Number & vbNewline & Err.Description
	End
End Sub

Function PasteTextFromClipboard(ByRef sOutput As String) As String
On Error Goto ErrHandler:
	Dim hMem As Long
	Dim pMem As Long
	Dim lMemSize As Long
	Dim sText As String
	Dim lLen As Long
	
	Dim lRetError As Long
	
	PasteTextFromClipboard = ""
	
	'Check for text on clipboard
	If IsClipboardFormatAvailable(CF_TEXT) = 0 Then
		PasteTextFromClipboard = "IsClipboardFormatAvailable is not CF_TEXT!"
		Exit Function
	End If
	
	' Open clipboard
	lRetError = OpenClipboard(0&)
	If lRetError <> 0 Then
		hMem = GetClipboardData(CF_TEXT)
		' If no text, close clipboard and exit
		If hMem = 0 Then
			PasteTextFromClipboard = "Couldn't get memmory-address from GetClipboardData(CF_TEXT)!"
			PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & hMem
			PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-address: " & hMem
		Else
			' Get memory pointer
			pMem = GlobalLock(hMem)
			If pMem = 0 Then
				PasteTextFromClipboard = "Couldn't get memory pointer from GlobalLock(hMem)!"
				PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & pMem
				PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-address: " & hMem
				PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-pointer: " & pMem
			Else
				' Get size of memory
				lMemSize = GlobalSize(hMem)
				If lMemSize = 0 Then
					PasteTextFromClipboard = "Couldn't get memory size from GlobalSize(hMem)!"
					PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & lMemSize
					PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-address: " & hMem
					PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-pointer: " & pMem
					PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-size: " & lMemSize
				Else
					' Get length of string
					lLen = lstrlenA(hMem)
					If lLen = 0 And False Then 'Ignore this (appears if pasteing from the same process as copied from), lLen is not used anyways
						PasteTextFromClipboard = "Couldn't get string length from lstrlenA(hMem)!"
						PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & lLen
						PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-address: " & hMem
						PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-pointer: " & pMem
						PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-size: " & lMemSize
						PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-string-length: " & lLen
					Else
						' Allocate local string
						'sText = String$(lLen, 90)
						sText = String$(lMemSize-1, 90)
						'sText = String$(lMemSize, 91)
						
						'MsgBox "Before lstrcpyA:" & Chr(10) & "Memory-address: " & hMem & Chr(10) & "Memory-pointer: " & pMem & Chr(10) & "Memory-size: " & lMemSize & Chr(10) & "Memory-string-length: " & lLen & Chr(10) & sText
						lRetError = lstrcpyA(sText, pMem)
						If lRetError = 0 Then
							PasteTextFromClipboard = "Couldn't get string from lstrcpyA(sText, pMem)!"
							PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & lRetError
							PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-address: " & hMem
							PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-pointer: " & pMem
							PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-size: " & lMemSize
							PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-string-length: " & lLen
							PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-string: " & sText
						Else
							' Unlock clipboard memory
							'AniMacro "%wait%1000"
							lRetError = GlobalUnlock(hMem)
							If lRetError = 0 And False Then 'Ignore this error (appears if pasteing from the same process as copied from), AniTa will probably unlock it
								PasteTextFromClipboard = "Couldn't GlobalUnlock(hMem)!"
								PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & lRetError
								PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-address: " & hMem
								PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-pointer: " & pMem
								PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-size: " & lMemSize
								PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-string-length: " & lLen
								PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Memory-string: " & sText
							Else
								'All ok!
								'MsgBox "After lstrcpyA:" & Chr(10) & "Memory-address: " & hMem & Chr(10) & "Memory-pointer: " & pMem & Chr(10) & "Memory-size: " & lMemSize & Chr(10) & "Memory-string-length: " & lLen & Chr(10) & sText
								
								sOutput = sText
								PasteTextFromClipboard = ""
							End If
						End If
					End If
				End If
			End If
		End If
		
		' Close clipboard
		lRetError = CloseClipboard()
		If lRetError = 0 Then
			If Len(PasteTextFromClipboard) > 0 Then
				PasteTextFromClipboard = PasteTextFromClipboard & vbNewline
			End If
			PasteTextFromClipboard = PasteTextFromClipboard & "Couldn't CloseClipboard()!"
			PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & lRetError
		End If
	Else
		PasteTextFromClipboard = "Couldn't OpenClipboard(0&)!"
		PasteTextFromClipboard = PasteTextFromClipboard & vbNewline & "Returned: " & lRetError
	End If
	Exit Function
	
ErrHandler:
    'Error handling
	If Len(PasteTextFromClipboard) > 0 Then
		PasteTextFromClipboard = PasteTextFromClipboard & vbNewline
	End If
	PasteTextFromClipboard = "An error occurd in PasteTextFromClipboard()!" & vbNewline & "Error-number: " & Err.Number & vbNewline & Err.Description
	CloseClipboard
	
	Dim Buffer As String
	Buffer = Space(200)
	FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, 0&, GetLastError(), LANG_NEUTRAL, Buffer, Len(Buffer), 0&
	MsgBox "Last win32api-error: " & Buffer
	
	Exit Function
End Function
