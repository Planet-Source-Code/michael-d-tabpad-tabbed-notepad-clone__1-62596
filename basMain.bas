Attribute VB_Name = "basMain"
'Written by Michael Dickens ( LFI.net )
'Source code released under GPL
'Please leave credit to the author (and my website) in any modifications made
'All modifications must be opensource, with source code available with binary

'Project page: www.LFI.net/LFI/prjTP.htm
'MichaelDickens@gmail.com
Option Explicit
Public UntitledCount As Integer
Public StatusBarOn As Boolean
Public StatusBarRandomTips As Boolean

Public MaxRecentDocs As Integer
Public RecentDocs As New Collection 'Our recent documents list
Public RecentTS As New Collection 'Our recent tabsets list

Public TabInfo() As TabInf 'Tab information array
Public LastTabIndex As Integer 'Last index selected
Public TabKeyMethod As Integer 'Method for when tabkey is pressed

Public RandomTips As New Collection
Public ShuffledRandomTips As New Collection

Public Type TabInf
    Text As String
    SelStart As Long
    SelLength As Long
    FilePath As String
End Type

'URL Support
Declare Function ShellEx Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal VKEY As Long) As Integer

'clsDialog crap:
Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type
Public Const OFN_ALLOWMULTISELECT   As Long = &H200
Public Const OFN_CREATEPROMPT       As Long = &H2000
Public Const OFN_EXPLORER           As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST      As Long = &H1000
Public Const OFN_HIDEREADONLY       As Long = &H4
Public Const OFN_LONGNAMES          As Long = &H200000
Public Const OFN_NOCHANGEDIR        As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_OVERWRITEPROMPT    As Long = &H2
Public Const OFN_PATHMUSTEXIST      As Long = &H800
Public Const OFN_READONLY           As Long = &H1
'End clsDialog


'Below code obtained from:
'http://www.xtremevbtalk.com/showpost.php?p=1062313&postcount=4
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
  ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long
    
Private Const EM_LINEINDEX = &HBB ' Retrieves the character index of the first
  '                                   character of a specified line

Private Const EM_LINEFROMCHAR As Long = &HC9 'Retrieves the index of the line that
'                                             contains the specified character index
'End obtained code

Function GetFileName(Path As String) As String
'From LFI.net's Winnap (Main module)
If Path = "" Then Exit Function
GetFileName = Split(Path, "\")(CharCount(Path, "\")) 'Get the filename from path
GetFileName = Replace(GetFileName, Chr(0), "") 'Remove vbNull characters if there are any (odd bug)
End Function

Function CharCount(Str As String, Char As String) As Long
'From LFI.net's Winnap (Main module)
CharCount = UBound(Split(LCase(Str), LCase(Char))) 'Get character count in a string
End Function

Sub MakePlaceHolder()
On Error GoTo errTrap
'Makes a file for new instances to write paths into
Open App.Path & "\TPLoadList.dat" For Output As #2
Close #2
'We leave it empty, the new instance works with it
Exit Sub
errTrap:
'Failed: Probably running in read only directory or from a CD
MsgBox "Error making Placeholder for TabPad: (" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf & "The directory you are running TabPad in is either read-only or TabPad does not have permission to create files essential for seamless Explorer integration." & vbCrLf & vbCrLf & "TabPad will continue to run, but if you use it as your default text editor it will fail to open more documents via explorer.", vbExclamation
End Sub

Sub AddToPH(Path As String)
On Error GoTo errTrap
'Called by new instances to add a path to placeholder
Open App.Path & "\TPLoadList.dat" For Append As #2
    'Open it for append (adds to end of file)
    Print #2, Path
    'Add the path
Close #2 'Close
Exit Sub
errTrap:
'Failed: Probably running in read only directory or from a CD
MsgBox "Error adding to Placeholder for TabPad: (" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf & "The directory you are running TabPad in is either read-only or TabPad does not have permission to create files essential for seamless Explorer integration." & vbCrLf & vbCrLf & "TabPad will continue to run, but if you use it as your default text editor it will fail to open more documents via explorer." & vbCrLf & vbCrLf & "Tried to open " & Path, vbExclamation
End Sub

Sub KillPH()
On Error Resume Next
'Kill placeholding file (called on shutdown)
Kill App.Path & "\TPLoadList.dat"
End Sub

Sub DeleteRRK()
On Error Resume Next 'Just incase keys don't exist
'Delete Registry Recent Keys
Dim i As Integer
For i = 1 To 32
    DeleteSetting "TabPad", "Recent", "Recent" & i
Next i
End Sub

Sub DeleteRRKTS()
On Error Resume Next 'Just incase keys don't exist
'Delete Registry Recent Keys
Dim i As Integer
For i = 1 To 32
    DeleteSetting "TabPad", "RecentTS", "Recent" & i
Next i
End Sub

Sub OpenInNP(Path As String)
On Error Resume Next
Shell "C:\Windows\Notepad.exe " & Path, vbNormalFocus
'Opens the file in Notepad, if it's installed to default path
End Sub

Function FileExists(Path As String) As Boolean
'From LFI.net's Winnap (Main module)
On Error GoTo e
Dim FL As Long
FL = FileLen(Path)
FileExists = True
e:
End Function

'Below code obtained from:
'http://www.xtremevbtalk.com/showpost.php?p=1062313&postcount=4

Function GetColNum(txtBox As TextBox) As Long

  Dim LineStart As Long
  Dim LineNumber As Long

  'Get current line number
  LineNumber = GetLineFromChar(txtBox.hwnd, txtBox.SelStart) - 1

  'Get index of first character in the line
  LineStart = SendMessage(txtBox.hwnd, EM_LINEINDEX, LineNumber, 0&)

  GetColNum = txtBox.SelStart - LineStart 'Change CharIndex relative to that line

End Function


Function GetLineFromChar(lHwnd As Long, CharIndex As Long)

  'hWnd => hWnd of a TextBox
  'CharIndex =>  Specifies the character index of the character contained in the
  '              line whose number is to be retrieved. If this parameter is ?1,
  '              EM_LINEFROMCHAR retrieves either the line number of the current line
  '              (the line containing the caret) or, if there is a selection, the line
  '              number of the line containing the beginning of the selection.

  GetLineFromChar = SendMessage(lHwnd, EM_LINEFROMCHAR, CharIndex, 0&) + 1

End Function


Sub CreatePathTree(Path As String)
On Error Resume Next
Dim CPath As String
Dim i As Integer
CPath = Split(Path, "\")(0) & "\"
For i = 1 To UBound(Split(Path, "\")) - 1
CPath = CPath & Split(Path, "\")(i) & "\"
MkDir CPath
Next i
End Sub

Sub ShellDef(file_name)
Dim x
x = ShellEx(frmMain.hwnd, "open", file_name, "", "", 10)
End Sub

Sub ShuffleCollection(OriginalCol As Collection, NewCol As Collection)
'Note: This module differs from the standardized Shuffle code, which I didn't have with me
Dim RandNum As Integer
Dim Orig As New Collection
Dim i As Integer
Randomize

For i = 1 To OriginalCol.Count
    Orig.Add OriginalCol.item(i)
Next i

Do Until Orig.Count = 0
RandNum = Int(Rnd * Orig.Count) + 1

NewCol.Add Orig.item(RandNum)
Orig.Remove (RandNum)
Loop
End Sub

Sub SetRandomTips()
RandomTips.Add "Control-Clicking words will load them in a web-browser!"
RandomTips.Add "LFI.net also writes other interesting programs similar to TabPad"
RandomTips.Add "TabPad is updated (and fixed) often, so check www.LFI.net for new builds reguarly"
RandomTips.Add "Help -> About gives you a link (which you can Control-Click) to check for updates"
RandomTips.Add "Press F1 for the next random tip"
RandomTips.Add "TabPad is opensource (Visual Basic 6), so you can make your own improvements"
RandomTips.Add "Crash Recovery will almost always keep you from loosing open documents"
RandomTips.Add "You can open the last document you worked with using Control-Insert"
RandomTips.Add "Your last 10 documents are in File -> Recent Documents"
RandomTips.Add "Clicking Status allows you to have white text on a black background"
RandomTips.Add "Press Control-S often to save your current work, or Control-Q to save all"
RandomTips.Add "Control-Q will save all your open documents"
RandomTips.Add "You can customise your Tab key under the 'Edit' menu"
RandomTips.Add "Crash Recovery can be activated and deactived in File -> Settings"
RandomTips.Add "You can change your font/font size under the 'Format' menu"
RandomTips.Add "TabPad is free and we appreciate it if you tell friends and family about it"
RandomTips.Add "You can associate text files directly with TabPad and open batches easily"
RandomTips.Add "Dragging and Dropping files into the textbox will open them in a new tab"
RandomTips.Add "Control-N will open a new tab"
End Sub
