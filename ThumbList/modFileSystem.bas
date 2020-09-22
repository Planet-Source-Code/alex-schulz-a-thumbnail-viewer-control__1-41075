Attribute VB_Name = "modFileSystem"
Option Explicit

Private Const MAX_PATH = 260
Private Const FOLDER_DELIMITER = "\"
Private Const ALL_FILES = "*.*"
Private Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Private Const SHGFI_TYPENAME = &H400                     '  get type name
Private Const INVALID_HANDLE_VALUE = -1

Private Type SHFILEINFO
    hIcon               As Long
    iIcon               As Long
    dwAttributes        As Long
    szDisplayName       As String * MAX_PATH
    szTypeName          As String * 80
End Type



Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Long
End Type

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function SHGetFileInfo Lib "Shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTmpPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTmpFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Function GetTempName(Optional Prefix As String, Optional Path As String)
Dim Unique As Long
Dim Buffer As String

If Path = vbNullString Then Path = GetTempPath
If Prefix = vbNullString Then Prefix = "fVB"
Unique = 0
      
Buffer = Space$(100)
GetTmpFileName Path, Prefix, Unique, Buffer
GetTempName = Mid$(Buffer, 1, InStr(Buffer, vbNullChar) - 1)
End Function

Public Sub CreateFolder(ByVal Path As String)
Dim SA As SECURITY_ATTRIBUTES
Dim s As String
Dim ix As Long
Dim All As Variant

EndTrim Path
SA.nLength = LenB(SA)

All = Split(Path, FOLDER_DELIMITER)

For ix = LBound(All) To UBound(All)
    s = s & All(ix) & FOLDER_DELIMITER
    If Not FolderExists(s) Then CreateDirectory s, SA
Next

End Sub

Public Function EndFix(ByRef Path As String) As String
If Right$(Path, 1) <> FOLDER_DELIMITER Then Path = Path & FOLDER_DELIMITER
End Function


Public Function EndTrim(ByRef Path As String) As String
If Right$(Path, 1) = FOLDER_DELIMITER Then Path = Left$(Path, Len(Path) - 1)
End Function

Public Function GetTempPath()
Dim Folder As String

Folder = String(MAX_PATH, 0)
If GetTmpPath(MAX_PATH, Folder) <> 0 Then
    GetTempPath = Left(Folder, InStr(Folder, vbNullChar) - 1)
Else
    GetTempPath = vbNullString
End If
End Function

Public Function GetFolderName(ByVal Path As String) As String
Dim ix As Long

If Path = vbNullString Then Exit Function

EndTrim Path
If Right$(Path, 1) = ":" Then
    GetFolderName = Path 'Left$(Path, Len(Path) - 1)
Else
    ix = InStrRev(Path, FOLDER_DELIMITER)
    If ix Then
        GetFolderName = Mid$(Path, ix + 1)
    Else
        GetFolderName = Path
    End If
End If
End Function







Function GetFileType(ByVal FileSpec As String) As String
Dim ShellFileInfo   As SHFILEINFO
Dim x               As Long

'get the fileinfo for this file
Call SHGetFileInfo(FileSpec, 0, ShellFileInfo, Len(ShellFileInfo), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME)
With ShellFileInfo
    x = InStr(.szTypeName, vbNullChar)
    If x Then GetFileType = Left$(.szTypeName, x - 1)
End With
End Function

Public Function BuildPath(ByVal FPath As String, ByVal FName As String) As String
If (FName = vbNullString) Or (FPath = vbNullString) Then Err.Raise 5

If Right$(FPath, 1) <> FOLDER_DELIMITER Then
    BuildPath = FPath & FOLDER_DELIMITER & FName
Else
    BuildPath = FPath & FName
End If
End Function
Public Function FolderExists(ByVal FolderSpec As String) As Boolean
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
Public Function FileExists(ByVal FileSpec As String) As Boolean
If FileSpec = vbNullString Then Err.Raise 5

On Error Resume Next
FileExists = Not (Dir(FileSpec, (vbArchive Or vbHidden Or vbReadOnly Or vbSystem)) = vbNullString)
If Err.Number Then FileExists = False
On Error GoTo 0
End Function
Public Function GetBaseName(ByVal FileSpec As String) As String
If FileSpec = vbNullString Then Err.Raise 5
Dim s   As String
Dim ix  As Long

s = GetFileName(FileSpec)
ix = InStr(s, ".")
If ix > 0 Then s = Left$(s, ix - 1)
GetBaseName = s
End Function
Public Function GetExtensionName(ByVal FileSpec As String) As String
Dim ix As Long

If FileSpec = vbNullString Then Err.Raise 5
ix = InStrRev(FileSpec, ".")
If ix > 0 Then
    GetExtensionName = Mid$(FileSpec, ix)
Else
    GetExtensionName = vbNullString
End If
End Function
Public Function GetFileName(ByVal FileSpec As String) As String
If FileSpec = vbNullString Then Err.Raise 5
GetFileName = Mid$(FileSpec, Len(GetParentFolder(FileSpec)) + 1)
End Function
Public Function GetParentFolder(ByVal FileSpec As String) As String
If FileSpec = vbNullString Then Err.Raise 5
Dim ix As Long

ix = InStrRev(FileSpec, FOLDER_DELIMITER)
If ix > 0 Then
    GetParentFolder = Left$(FileSpec, ix)
Else
    ix = InStr(FileSpec, ":")
    If ix > 0 Then
        GetParentFolder = Left$(FileSpec, ix) & FOLDER_DELIMITER
    Else
        GetParentFolder = vbNullString
    End If
End If
End Function
Public Function GetSystemFolder() As String
Dim Buffer  As String * MAX_PATH
Dim Length  As Long

Length = GetSystemDirectory(Buffer, MAX_PATH)

GetSystemFolder = Left$(Buffer, Length)
End Function

Public Function GetWindowsFolder() As String
Dim Buffer  As String * MAX_PATH
Dim Length  As Long

Length = GetWindowsDirectory(Buffer, MAX_PATH)

GetWindowsFolder = Left$(Buffer, Length)
End Function






