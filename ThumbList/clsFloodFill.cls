VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFloodFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Enum ALIGNMENTS
    ffLeft
    ffCenter
    ffRight
End Enum

Enum MODES
    ffPlain
    ffPercent
    ffCaption
    ffCaptionPercent
End Enum

Private m_Panel         As PictureBox
Private m_Min           As Single
Private m_Max           As Single
Private m_Value         As Single
Private m_Color         As OLE_COLOR
Private m_Alignment     As ALIGNMENTS
Private m_Mode          As MODES
Private m_Caption       As String
Private m_Completed     As Single


















Private Sub Adjust()
Clear
If m_Max < m_Min Then m_Max = m_Min + 1
m_Value = m_Min
End Sub

Public Sub Clear()
If Not (m_Panel Is Nothing) Then m_Panel.Cls
End Sub


Private Sub EstablishAlignment()
With m_Panel
    Select Case m_Alignment
        Case ffLeft
            .CurrentX = 2
            .CurrentY = (.ScaleHeight - .TextHeight(m_Caption)) / 2

        Case ffCenter
            .CurrentX = (.ScaleWidth - .TextWidth(m_Caption)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(m_Caption)) / 2
          
        Case ffRight
            .CurrentX = (.ScaleWidth - .TextWidth(m_Caption)) - 3
            .CurrentY = (.ScaleHeight - .TextHeight(m_Caption)) / 2
    End Select
End With
End Sub




Public Property Get Panel() As PictureBox
Set Panel = m_Panel
End Property
Public Property Set Panel(pb As PictureBox)
Set m_Panel = pb
If Not pb Is Nothing Then
    With m_Panel
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .BackColor = vbWhite
        .ForeColor = m_Color
        .DrawMode = vbNotXorPen
        .FillStyle = vbFSSolid
        .Cls
    End With
End If
End Property

Public Property Get Min() As Single
Min = m_Min
End Property
Private Sub CalculateCompleted()
m_Panel.Cls
m_Completed = (m_Value - m_Min) / (m_Max - m_Min)
End Sub
Private Sub ShowResults()
With m_Panel
    'print the percentage string in the text colour
     If Not (m_Mode = ffPlain) Then m_Panel.Print m_Caption
      
    'print the flood bar to the new m_Value length in the line colour
    m_Panel.Line (0, 0)-(.ScaleWidth * m_Completed, .ScaleHeight), .ForeColor, BF
     
    'allow the flood to complete drawing
    DoEvents
End With
End Sub
Public Property Get Value() As Single
Value = m_Value
End Property
Public Property Get Max() As Single
Max = m_Max
End Property
Public Property Let Min(ByVal v As Single)
m_Min = v
Adjust
End Property
Public Property Let Value(ByVal v As Single)
m_Value = v
If Not (m_Panel Is Nothing) Then
    If m_Value > m_Max Then
        Exit Property
    ElseIf m_Value < m_Min Then
        m_Value = m_Min
        m_Panel.Cls
    Else
        Dim LocalCopy As String
        LocalCopy = m_Caption
        CalculateCompleted
        
        Select Case m_Mode
            Case ffPercent
                m_Caption = Format(m_Completed, "###%")
                EstablishAlignment
            Case ffCaption
                EstablishAlignment
            Case ffCaptionPercent
                m_Caption = LocalCopy & Format(m_Completed, "   ###%")
                EstablishAlignment
            Case Else
                'nothing to do really
        End Select
        ShowResults
        m_Caption = LocalCopy
    End If
End If
End Property
Public Property Let Max(ByVal v As Single)
m_Max = v
Adjust
End Property
Public Property Get Mode() As MODES
Mode = m_Mode
End Property
Public Property Let Mode(ByVal m As MODES)
m_Mode = m
If m = ffPercent Then m_Alignment = ffCenter
End Property
Public Property Get Alignment() As ALIGNMENTS
Alignment = m_Alignment
End Property
Public Property Let Alignment(ByVal a As ALIGNMENTS)
m_Alignment = a
End Property
Public Property Get Caption() As String
Caption = m_Caption
End Property
Public Property Let Caption(ByVal s As String)
m_Caption = s
End Property
Private Sub Class_Initialize()
m_Color = &H7F0000
m_Mode = ffPercent
End Sub
Private Sub Class_Terminate()
Set m_Panel = Nothing
End Sub
Public Property Get Color() As OLE_COLOR
Color = m_Color
End Property
Public Property Let Color(ByVal c As OLE_COLOR)
m_Color = c
If Not m_Panel Is Nothing Then m_Panel.ForeColor = c
End Property
