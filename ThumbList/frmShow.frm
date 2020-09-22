VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShow 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2910
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3180
   Icon            =   "frmShow.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComCtl2.FlatScrollBar VScroll1 
      Height          =   1935
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3413
      _Version        =   393216
      LargeChange     =   10
      Orientation     =   1179648
   End
   Begin MSComCtl2.FlatScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Orientation     =   1179649
   End
   Begin VB.PictureBox picOuter 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.PictureBox picInner 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7455
         Left            =   0
         ScaleHeight     =   497
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   688
         TabIndex        =   1
         Top             =   -3960
         Width           =   10320
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Full Screen"
      End
      Begin VB.Menu mnuSize 
         Caption         =   "25%"
         Index           =   0
      End
      Begin VB.Menu mnuSize 
         Caption         =   "50%"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   "75%"
         Index           =   2
      End
      Begin VB.Menu mnuSize 
         Caption         =   "100%"
         Index           =   3
      End
      Begin VB.Menu mnuSize 
         Caption         =   "125%"
         Index           =   4
      End
      Begin VB.Menu mnuSize 
         Caption         =   "150%"
         Index           =   5
      End
      Begin VB.Menu mnuSize 
         Caption         =   "175%"
         Index           =   6
      End
      Begin VB.Menu mnuSize 
         Caption         =   "200%"
         Index           =   7
      End
      Begin VB.Menu mnuSize 
         Caption         =   "225%"
         Index           =   8
      End
      Begin VB.Menu mnuSize 
         Caption         =   "250%"
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SWP_HIDEWINDOW& = &H80
Private Const SWP_SHOWWINDOW& = &H40
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private semaphore               As Boolean
Private Hidden                  As Boolean
Private MagX                    As Single
Private OffX                    As Single
Private OffY                    As Single
Private CAH                     As Single
Private CAW                     As Single
Private Tx                      As Single
Private Ty                      As Single
Private ScreenWidth             As Single
Private ScreenHeight            As Single
Private m_Filepath              As String
Private m_Caption               As String

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Property Let Filepath(param As String)
Dim fs As FileSystemObject

On Error Resume Next

If param = vbNullString Then
    Set pic1.Picture = LoadPicture(vbNullString)
    Hide
Else
    Set fs = New FileSystemObject
    
    m_Filepath = param
    
    With pic1
        Set .Picture = LoadPicture(m_Filepath)
        If Err.Number = 0 Then
            m_Caption = fs.GetFileName(m_Filepath) & ", " & CStr(CInt(pic1.ScaleWidth)) & " X " & CStr(CInt(pic1.ScaleHeight)) & " Pixels, " & Format(FileLen(m_Filepath) / 1024, "####.## Kb. ")
    
            CAW = Width - ScaleWidth * Tx
            CAH = Height - ScaleHeight * Ty
    
            ShowPic
        Else
            MsgBox "Unsupported or invalid image file!", vbOKOnly + vbDefaultButton1 + vbCritical + vbMsgBoxSetForeground, "Error"
        End If
    End With
    Set fs = Nothing
End If
End Property


Property Let ScaleFactor(param As Single)
If param > 1 Then param = 1
MagX = param
End Property


Private Sub ShowPic()
Const SCROLLBARSIZE = 18
Dim hh      As Single
Dim ww      As Single
Dim nuW     As Single
Dim nuH     As Single

With pic1
    nuW = .ScaleWidth * MagX
    nuH = .ScaleHeight * MagX
End With

Do
    ww = (nuW + 20) * Tx
    hh = (nuH + 20) * Ty

    If hh > 32000 Or ww > 32000 Then
        MagX = MagX * 0.75
    Else
        Exit Do
    End If
Loop


picOuter.Move 0, 0
Caption = m_Caption & Format(MagX, "(###%)")


With picInner
    .Cls
    .Width = nuW
    .Height = nuH
    .PaintPicture pic1.Picture, 0, 0, nuW, nuH
    .Refresh
End With


'/////// bigger than display //////////
If hh > ScreenHeight And ww > ScreenWidth Then
    
    'these are in twips
    'allow a pixel on each side
    Width = ScreenWidth - (Tx + Tx)
    'allow a pixel on top & bottom
    Height = ScreenHeight - (Ty + Ty)
    
    'these are in pixels
    picOuter.Move 0, 0, ScaleWidth - 20, ScaleHeight - 25
       
    With HScroll1
        .Move 0, picOuter.Height + 2, picOuter.Width, SCROLLBARSIZE
        .Max = picInner.ScaleWidth - picOuter.ScaleWidth
        .Value = 0
        .Visible = True
    End With
    
    With VScroll1
        .Move picOuter.Width + 2, 0, SCROLLBARSIZE, picOuter.Height
        .Max = picInner.ScaleHeight - picOuter.ScaleHeight
        .Value = 0
        .Visible = True
    End With
    Move 0, 0
    

'/////// higher than display //////////
ElseIf hh > ScreenHeight And ww <= ScreenWidth Then
    
    'these are in twips
    Width = ww
    'allow a pixel on top & bottom
    Height = ScreenHeight - (Ty + Ty)
    
    'these are in pixels
    picOuter.Move 0, 0, ScaleWidth - (SCROLLBARSIZE + 2), ScaleHeight
       
    With VScroll1
        .Move picOuter.Width + 2, 0, SCROLLBARSIZE, picOuter.Height
        .Max = picInner.ScaleHeight - picOuter.ScaleHeight
        .Value = 0
        .Visible = True
    End With
    Move (Screen.Width - Width) / 2, 0
       
'/////// wider than display //////////
ElseIf hh <= ScreenHeight And ww > ScreenWidth Then
    'these are in twips
    'allow a pixel on each side
    Width = ScreenWidth - (Tx + Tx)
    Height = hh
    
    picOuter.Move 0, 0, ScaleWidth, ScaleHeight - (SCROLLBARSIZE + 2)
        
    With HScroll1
        .Move 0, picOuter.Height + 2, picOuter.Width, SCROLLBARSIZE
        .Max = picInner.ScaleWidth - picOuter.ScaleWidth
        .Value = 0
        .Visible = True
    End With
    
    Move 0, (Screen.Height - Height) / 2
    
'/////// fits the display //////////
ElseIf hh <= ScreenHeight And ww <= ScreenWidth Then
    'these are in twips
    Width = (pic1.ScaleWidth * MagX) * Tx + CAW
    Height = (pic1.ScaleHeight * MagX) * Ty + CAH
        
    'these are in pixels
    picOuter.Move 0, 0, ScaleWidth, ScaleHeight
    Move (ScreenWidth - Width) \ 2, (ScreenHeight - Height) \ 2
    
End If

DoEvents
End Sub

Private Sub Form_Initialize()
MagX = 1
OffX = 0
OffY = 0

With Screen
    Tx = .TwipsPerPixelX
    Ty = .TwipsPerPixelY
    ScreenWidth = .Width
    ScreenHeight = .Height
End With
End Sub



Private Sub Form_Resize()
picOuter.Move 0, 0
picInner.Move 0, 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set picInner.Picture = Nothing
Set pic1.Picture = Nothing
End Sub

Private Sub HScroll1_Change()
With HScroll1
    picInner.Left = -.Value
    If .Value = .Min Then
        .Arrows = cc2RightDown
    ElseIf .Value = .Max Then
        .Arrows = cc2LeftUp
    Else
        .Arrows = cc2Both
    End If
End With
End Sub

Private Sub mnuFullScreen_Click()
Dim X1      As Single
Dim X2      As Single
Dim nuW     As Single
Dim nuH     As Single
Dim Ratio   As Single


'black ou the whole screen
HScroll1.Visible = False
VScroll1.Visible = False
picOuter.Visible = False

Move 0, 0, Screen.Width, Screen.Height
BackColor = vbBlack

With pic1
    X1 = Screen.Width / (.ScaleWidth * Tx)
    X2 = Screen.Height / (.ScaleHeight * Ty)
    Ratio = IIf(X1 < X2, X1, X2)
    nuW = .ScaleWidth * Ratio
    nuH = .ScaleHeight * Ratio
End With

With picInner
    .Cls
    .Width = nuW
    .Height = nuH
    .PaintPicture pic1.Picture, 0&, 0&, nuW, nuH
    .Refresh
    picOuter.Width = .Width
    picOuter.Height = .Height
End With

With picOuter
    .Move (ScaleWidth - .Width) \ 2, (ScaleHeight - .Height) \ 2
    .Visible = True
    .Refresh
End With

DoEvents
semaphore = True
End Sub

Private Sub mnuSize_Click(Index As Integer)
MagX = 0.25 + Index * 0.25
ShowPic
End Sub

Private Sub picInner_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    PopupMenu mnuPopup
ElseIf semaphore Then
    semaphore = False
    ShowPic
End If
End Sub

Private Sub picInner_Resize()

picInner.Move 0, 0


End Sub


Private Sub VScroll1_Change()
With VScroll1
    picInner.Top = -.Value
    If .Value = .Min Then
        .Arrows = cc2RightDown
    ElseIf .Value = .Max Then
        .Arrows = cc2LeftUp
    Else
        .Arrows = cc2Both
    End If
End With

End Sub


