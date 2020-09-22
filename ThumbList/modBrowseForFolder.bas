Attribute VB_Name = "modBrowseForFolder"
Option Explicit
'common to both methodsprivate
Private Type BROWSEINFO
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As Long
    lParam          As Long
    iImage          As Long
End Type

Private Type RECT
   Left     As Long
   top      As Long
   Right    As Long
   bottom   As Long
End Type

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Private Const DT_SINGLELINE = &H20&
Private Const DT_PATH_ELLIPSIS = &H4000&
Private Const DT_MODIFYSTRING = &H10000

Private Const DFLT_PROMPT = "Select a folder"

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long      'specific to the STRING method
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function PathCompactPathA Lib "shlwapi" (ByVal hDC As Long, ByVal lpszPath As String, ByVal dx As Long) As Long

Public Function CompactPath(ByVal Path As String, ByVal MaxWidthPixels As Long, ByVal hDC As Long) As String
Dim R As RECT

R.Right = MaxWidthPixels
DrawText hDC, Path, -1, R, DT_PATH_ELLIPSIS Or DT_SINGLELINE Or DT_MODIFYSTRING
CompactPath = Path
End Function

Function ShortName(ByVal Path As String) As String
Dim ShortPath   As String * MAX_PATH
Dim R           As Long
    
R = GetShortPathNameA(Path, ShortPath, MAX_PATH)
ShortName = Left(ShortPath, R)
End Function
Public Function BrowseForFolder(Optional ByVal StartFolder As String = vbNullString, Optional ByVal Prompt As String = DFLT_PROMPT) As String
Dim BI          As BROWSEINFO
Dim pidl        As Long
Dim lpSelPath   As Long
Dim sz          As Long
Dim sPath       As String * MAX_PATH
Dim NuPath      As String

If StartFolder = vbNullString Then
    StartFolder = CurDir
Else
    If Not FolderExists(StartFolder) Then StartFolder = CurDir
End If


With BI
    .hOwner = 0&
    .pidlRoot = 0
    .lpszTitle = Prompt
    .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
    sz = Len(StartFolder)
    lpSelPath = LocalAlloc(LPTR, sz)
    MoveMemory ByVal lpSelPath, ByVal StartFolder, sz
    .lParam = lpSelPath
End With

pidl = SHBrowseForFolder(BI)
If pidl Then
    If SHGetPathFromIDList(pidl, sPath) Then NuPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
    Call CoTaskMemFree(pidl)
End If
Call LocalFree(lpSelPath)

If NuPath <> vbNullString Then
    If Right(NuPath, 1) <> "\" Then NuPath = NuPath & "\"
End If
BrowseForFolder = NuPath
End Function



Public Function FolderExists(ByVal FolderSpec As String) As Boolean
Const FOLDER_DELIMITER = "\"
Const ALL_FILES = "*.*"
If FolderSpec = vbNullString Then Exit Function

Dim s As String

s = FolderSpec

If Right(s, 1) <> FOLDER_DELIMITER Then
    s = s & FOLDER_DELIMITER & ALL_FILES
Else
    s = s & ALL_FILES
End If

On Error Resume Next

FolderExists = Not (Dir(s, vbDirectory) = vbNullString)
If Err.Number Then FolderExists = False

On Error GoTo 0
End Function


Public Function BrowseCallbackProcStr(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
'Callback for the Browse STRING method.
'On initialization, set the dialog's
'pre-selected folder from the pointer
'to the path allocated as bi.lParam,
'passed back to the callback as lpData param.
Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal StrFromPtrA(lpData))
    Case Else:
End Select
End Function
Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
'Callback for the Browse PIDL method.
'On initialization, set the dialog's
'pre-selected folder using the pidl
'set as the bi.lParam, and passed back
'to the callback as lpData param.

Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, False, ByVal lpData)
    Case Else:
End Select
End Function
Public Function FARPROC(pfn As Long) As Long
'A dummy procedure that receives and returns
'the value of the AddressOf operator.
'Obtain and set the address of the callback
'This workaround is needed as you can't assign
'AddressOf directly to a member of a user-
'defined type, but you can assign it to another
'long and use that (as returned here)
FARPROC = pfn
End Function





Public Function StrFromPtrA(lpszA As Long) As String
'Returns an ANSII string from a pointer to an ANSII string.
Dim sRtn As String
sRtn = String$(lstrlenA(ByVal lpszA), 0)
Call lstrcpyA(ByVal sRtn, ByVal lpszA)
StrFromPtrA = sRtn
End Function

