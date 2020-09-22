VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const EDGE_RAISED = &H5
Private Const EDGE_SUNKEN = &HA
Private Const BF_RECT = &HF
Private Const BF_SOFT = &H1000

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Name             As String
Public Path             As String
Public Key              As String
Public FileSize         As Long
Public Row              As Long
Public Col              As Long
Public Width            As Long
Public Height           As Long
Public Cloning          As Boolean

Public Pic              As StdPicture
Public Target           As PictureBox

Private m_Bounding      As RECT
Private m_Left          As Long
Private m_Top           As Long
Private m_Size          As Long
Private m_Selected      As Boolean
Private m_HasFocus      As Boolean

Private Sub CalculateRect()
SetRect m_Bounding, m_Left, m_Top, m_Left + m_Size, m_Top + m_Size
End Sub

Public Function Clicked(ByVal x As Long, ByVal y As Long) As Boolean
Clicked = (PtInRect(m_Bounding, x, y) <> 0)
End Function


Public Property Get Description() As String
Description = Name & ", " & Format(FileSize / 1024, "#### Kb.") & ", " & CStr(Width) & " Pixels x " & CStr(Height) & " Pixels"
End Property


Public Sub EraseThumb()
'Dim hBrush      As Long
'
'hBrush = CreateSolidBrush(Target.BackColor)
'FillRect Target.hdc, m_Bounding, hBrush
'DeleteObject hBrush
With m_Bounding
    Target.Line (.Left, .Top)-(.Right, .Bottom), Target.BackColor, BF
End With
Target.Refresh
End Sub

Public Sub Render()
Target.PaintPicture Pic, m_Left, m_Top, , , , , , , vbSrcCopy
If m_Selected Then DrawEdge Target.hdc, m_Bounding, EDGE_SUNKEN, BF_SOFT + BF_RECT
If m_HasFocus Then DrawFocusRect Target.hdc, m_Bounding
Target.Refresh
End Sub

Public Property Get Selected() As Boolean
Selected = m_Selected
End Property

Private Sub Class_Terminate()
Set Pic = Nothing
Set Target = Nothing
End Sub

Public Property Get Left() As Long
Left = m_Left
End Property

Public Property Get Size() As Long
Size = m_Size
End Property

Public Property Get Top() As Long
Top = m_Top
End Property
Public Property Let Left(ByVal v As Long)
m_Left = v
CalculateRect
End Property
Public Property Let Size(ByVal v As Long)
m_Size = v
CalculateRect
End Property
Public Property Let Top(ByVal v As Long)
m_Top = v
CalculateRect
End Property

Public Property Get HasFocus() As Boolean
HasFocus = m_HasFocus
End Property

Public Property Let HasFocus(ByVal b As Boolean)
If Cloning Then
    m_HasFocus = b
    Exit Property
End If

If b = True Then
    If Not m_HasFocus Then
        DrawFocusRect Target.hdc, m_Bounding
    End If
Else
    If m_HasFocus Then
        DrawFocusRect Target.hdc, m_Bounding
    End If
End If
m_HasFocus = b

End Property

Public Property Let Selected(ByVal b As Boolean)
If Cloning Then
    m_Selected = b
    Exit Property
End If

If b Then
    If Not m_Selected Then
        DrawEdge Target.hdc, m_Bounding, EDGE_SUNKEN, BF_SOFT + BF_RECT
        Target.Refresh
    End If
Else
    If m_Selected Then Render
End If
m_Selected = b
End Property
