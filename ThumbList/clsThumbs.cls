VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThumbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsThumb"
Attribute VB_Ext_KEY = "Member0" ,"clsThumb"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_Col           As Collection
Private m_TotalSize     As Double


Public Function Add(Target As PictureBox, ByVal Size As Long, ByVal Left As Long, ByVal Top As Long, ByVal Row As Long, ByVal Col As Long, ByVal Key As String) As clsThumb
'create a new object
Dim objNewMember As clsThumb
Set objNewMember = New clsThumb

With objNewMember
    Set .Target = Target
    .Size = Size
    .Left = Left
    .Top = Top
    .Row = Row
    .Col = Col
    .Key = Key
End With

m_Col.Add objNewMember, Key

'return the object created
Set Add = objNewMember

Set objNewMember = Nothing
End Function

Public Function AddThumb(ByRef Thumb As clsThumb) As clsThumb
Dim objNewMember    As clsThumb

Set objNewMember = New clsThumb

With objNewMember
    .Cloning = True
    Set .Target = Thumb.Target
    Set .Pic = Thumb.Pic
    .Size = Thumb.Size
    .Left = Thumb.Left
    .Top = Thumb.Top
    .Row = Thumb.Row
    .Col = Thumb.Col
    .Key = Thumb.Key
    .Name = Thumb.Name
    .Path = Thumb.Path
    .FileSize = Thumb.Size
    .Width = Thumb.Width
    .Height = Thumb.Height
    .Selected = Thumb.Selected
    .HasFocus = Thumb.HasFocus
    .Cloning = False
End With

m_Col.Add objNewMember, Thumb.Key

'return the object created
Set AddThumb = objNewMember

Set objNewMember = Nothing
End Function

Public Sub ClearAll()
With m_Col
    While .Count > 0
        .Remove .Count
    Wend
End With
End Sub

Public Property Get Item(Key As String) As clsThumb
Attribute Item.VB_UserMemId = 0
Set Item = m_Col(Key)
End Property


Public Property Get SelCount() As Long
Dim Thumb As clsThumb
Dim Count As Long

For Each Thumb In m_Col
    If Thumb.Selected Then Count = Count + 1
Next

SelCount = Count
End Property



Public Property Get SelSize() As Long
Dim Thumb   As clsThumb
Dim Size    As Double

For Each Thumb In m_Col
    If Thumb.Selected Then Size = Size + Thumb.FileSize
Next

SelSize = Size
End Property
Public Property Get Count() As Long
Count = m_Col.Count
End Property


Public Sub Remove(Key As String)
m_Col(Key).EraseThumb
m_Col.Remove Key
End Sub



Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_Col.[_NewEnum]
End Property


Public Function ThumbRow(ByVal Key As String) As Long
ThumbRow = m_Col(Key).Row
End Function

Public Function ThumbCol(ByVal Key As String) As Long
ThumbCol = m_Col(Key).Col
End Function
Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub


Private Sub Class_Terminate()
Set m_Col = Nothing
End Sub

