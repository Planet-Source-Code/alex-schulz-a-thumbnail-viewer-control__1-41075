Attribute VB_Name = "modMain"
Option Explicit
Public Const YES = 0
Public Const NO = 1
Public Const CANCEL = 2
Public Const ABORT = vbNullString

Public Const FO_DELETE = &H3&
Public Const FOF_ALLOWUNDO = &H40&
Public Const FOF_SILENT = &H4&
Public Const DT_SINGLELINE = &H20&
Public Const DT_PATH_ELLIPSIS = &H4000&
Public Const DT_MODIFYSTRING = &H10000

Type SHFILEOPSTRUCT
    hwnd                    As Long
    wFunc                   As Long
    pFrom                   As String
    pTo                     As String
    fFlags                  As Integer
    fAnyOperationsAborted   As Long
    hNameMappings           As Long
    lpszProgressTitle       As String
End Type

Type RECT
    Left        As Long
    top         As Long
    Right       As Long
    bottom      As Long
End Type


Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function PathCompactPathA Lib "shlwapi" (ByVal hDC As Long, ByVal lpszPath As String, ByVal dx As Long) As Long









Public Sub Recycle(ByVal Target As String, Optional AllowUndo As Boolean = True)
Dim SHFileOp        As SHFILEOPSTRUCT
    
With SHFileOp
    .wFunc = FO_DELETE
    .pFrom = (Target & vbNullChar & vbNullChar)
    .fFlags = FOF_SILENT Or IIf(AllowUndo, FOF_ALLOWUNDO, 0)
End With

SHFileOperation SHFileOp
End Sub

Sub Main()
Dim f As frmMain

Set f = New frmMain
f.Show vbModal


UnloadAll
End Sub


Private Sub UnloadAll()
Dim f As Form

For Each f In Forms
    Unload f
    Set f = Nothing
Next
End Sub


