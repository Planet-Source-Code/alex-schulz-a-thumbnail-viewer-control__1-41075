VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ViewerPane 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ViewerPane.ctx":0000
   Begin VB.PictureBox StatusPane 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   2
      Top             =   3240
      Width           =   4215
   End
   Begin VB.PictureBox Frame 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   240
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3600
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer ViewTimer 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   2400
         Top             =   1560
      End
      Begin MSComCtl2.FlatScrollBar vs1 
         Height          =   2055
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   3625
         _Version        =   393216
         Appearance      =   2
         Orientation     =   1179648
         SmallChange     =   4
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   2520
         Pattern         =   "*.jpg;*.gif;*.bmp"
         TabIndex        =   3
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Container 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   480
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuProps 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuPaper 
         Caption         =   "Set as WallPaper"
      End
      Begin VB.Menu mnuShellView 
         Caption         =   "ShellView"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "ViewerPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'private constants
Private Const COLORONCOLOR = 3
Private Const FIRST = 0
Private Const THUMB_OFFSET = 1
Private Const NOT_DEFINED = vbNullString
Private Const SW_SHOW = 5
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Private Const DEFAULT_BACKCOLOR = vbWhite
Private Const DEFAULT_PATTERN = "*.jp*;*.gif;*.bmp"
Private Const DEFAULT_WALLPAPER_NAME = "ThumbList"
Private Const FOLDER_DELIMITER = "\"
Private Const MAX_PATH = 260
Private Const ALL_FILES = "*.*"
'===================================================
'public enums
'===================================================
Enum ThumbSizes
    Small = 0
    Medium = 1
    Large = 2
    [Extra Large] = 3
End Enum

Enum BORDER_STYLES
    None = 0
    [Fixed Single] = 1
End Enum

'Private types
Private Type ThumbPos
    Row         As Long
    Col         As Long
End Type

'member variables
Private m_BackColor         As Long
Private m_Loaded            As Boolean      'true when loaded
Private m_OffsetY           As Long         'space required for 1 thumb y
Private m_OffsetX           As Long         'space required for 1 thumb x
Private m_StartX            As Long         'where thumbs start on left
Private m_Cols              As Long         'total cols
Private m_Row               As Long         'current row
Private m_Col               As Long         'current col
Private m_X                 As Long         'current position x
Private m_Y                 As Long         'current position y
Private m_Size              As ThumbSizes   'current size
Private m_ThumbSize         As Long         'actual thumbsize
Private m_Last              As String
Private m_Current           As String
Private m_WallPaperName     As String
Private m_LastKey           As String
Private m_TotalSize         As Double       'total bytes selected
Private m_Busy              As Boolean      'control is busy
Private m_AllowCopy         As Boolean
Private m_AllowMove         As Boolean
Private m_AllowDelete       As Boolean
Private m_AllowShellView    As Boolean
Private m_AllowPaper        As Boolean
Private m_ClickX            As Long
Private m_ClickY            As Long
Private m_Thumbs            As clsThumbs


'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Container,Container,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Container,Container,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Container,Container,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Container,Container,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Container,Container,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Container,Container,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event DblClick(Path As String)
Event SelectionChange(ByVal ItemSelected As Boolean)
Event Scroll()
Event Delete(Permit As Boolean, ByVal Target As String)
Event Move(Permit As Boolean, ByVal Target As String)
Event Copy(ByVal Target As String)
Event Change(ByVal Count As Long)
Event IsBusy(ByVal State As Boolean)

'API 's
Private Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SystemParametersInfoA Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Sub CopyFiles()
mnuCopy_Click
End Sub

Public Sub DeletFiles()
mnuDelete_Click
End Sub

Private Sub KeyCheck(Key As String)
If m_Current = Key Then
    m_Current = NOT_DEFINED
    StatusPanePrint
End If
End Sub

Private Sub LoadThumbs()
'first clear all existing thumbs
Clear

'see if ther are nay files and act accordingly
If (File1.ListCount = 0) Then
    'no files here
    StatusPanePrint "No files found."
Else
    Dim ix          As Long
    
    'we have files, disable clicks
    SayBusy
    
    'setup the progressbar
    Dim ff As clsFloodFill
    Set ff = New clsFloodFill
    
    With ff
        .Min = 0
        .Max = File1.ListCount
        .Mode = ffCaption
        .Alignment = ffCenter
        Set .Panel = StatusPane
    End With

    'add the thumbs
    With File1
        'for each file
        For ix = 0 To .ListCount - 1
            'show the name in the progressbar
            ff.Caption = .List(ix)
            
            'add the thumb
            AddThumb ix
            
            'show the progress
            ff.Value = ix + 1
        Next
    End With
    
    'all done, clear the progressbar and dispose of it
    ff.Clear
    Set ff = Nothing
    
    'say there are thumbs loaded
    m_Loaded = True
    
    'restore statuspane to its normal color
    StatusPane.ForeColor = vbButtonText
    StatusPanePrint CStr(m_Thumbs.Count) & " images loaded."
    're-enable clicks
    SayBusy False
End If
End Sub

Public Sub MoveFiles()
mnuMove_Click
End Sub

Public Property Get Selected() As Boolean
Selected = (m_Current <> NOT_DEFINED)
End Property
Public Function SelectedItemPath() As String
If m_Current = NOT_DEFINED Then
    SelectedItemPath = NOT_DEFINED
Else
    SelectedItemPath = m_Thumbs(m_Current).Path
End If
End Function
Public Property Get AllowCopy() As Boolean
AllowCopy = m_AllowCopy
End Property
Public Property Get AllowPaper() As Boolean
AllowPaper = m_AllowPaper
End Property
Public Property Get AllowDelete() As Boolean
AllowDelete = m_AllowDelete
End Property
Public Property Get AllowMove() As Boolean
AllowMove = m_AllowMove
End Property
Public Property Let AllowCopy(ByVal b As Boolean)
m_AllowCopy = b
PropertyChanged "AllowCopy"
End Property
Public Property Let AllowMove(ByVal b As Boolean)
m_AllowMove = b
PropertyChanged "AllowMove"
End Property
Public Property Let AllowDelete(ByVal b As Boolean)
m_AllowDelete = b
PropertyChanged "AllowDelete"
End Property

Public Property Let AllowPaper(ByVal b As Boolean)
m_AllowPaper = b
PropertyChanged "AllowPaper"
End Property
Public Property Get AllowShellView() As Boolean
AllowShellView = m_AllowShellView
End Property

Public Property Let AllowShellView(ByVal b As Boolean)
m_AllowShellView = b
PropertyChanged "AllowShellView"
End Property
Private Sub DeSelectAll()
m_Current = NOT_DEFINED
SetFocusTo m_Current
ClearSelected
SetFocusTo m_Current
StatusPanePrint

End Sub

Public Property Get SelCount() As Long
SelCount = m_Thumbs.SelCount
End Property
Public Property Get Busy() As Boolean
Busy = m_Busy
End Property


Private Sub CalculateOffsets()
Dim T As Long

m_OffsetX = m_ThumbSize + THUMB_OFFSET
m_OffsetY = m_OffsetX

T = Container.ScaleWidth - vs1.Width

m_Cols = T \ m_OffsetX - 1
m_StartX = (T Mod m_OffsetX) \ (m_Cols + 2)
m_OffsetX = m_ThumbSize + m_StartX
End Sub

Public Sub Clear()
'cleanup the picturebox container
PrepareToLoad

'empty the collection
m_Thumbs.ClearAll

'say not loaded
m_Loaded = False

'say no thumb currently selected or with focus
m_Last = NOT_DEFINED
m_Current = NOT_DEFINED

RaiseEvent Change(m_Thumbs.Count)
End Sub

Private Sub ClearSelected()
Dim Thumb As clsThumb

For Each Thumb In m_Thumbs
    Thumb.Selected = False
Next

m_TotalSize = 0
StatusPane.Cls
End Sub

Public Sub Display()
mnuView_Click
End Sub

Private Sub GetNextPosition()
' This function updates the m_X, m_Y coordinates
' as needed. It also keeps track of the current row
' and column which are stored in m_Row and m_Col
' respectively.

'update the left corner position
m_X = m_OffsetX + m_X
'update the column
m_Col = m_Col + 1

'see if we are going to run out of the container
'horizontally
If m_Col > m_Cols Then
    'if so increment the top corner position
    m_Y = m_Y + m_OffsetY
    'if we need more container get it now
    If m_Y + m_OffsetY > Frame.ScaleHeight Then Container.Height = m_Y + m_OffsetY
    'reset the left corner position
    m_X = m_StartX
    'reset the column
    m_Col = FIRST
    'and increment the row
    m_Row = m_Row + 1
End If
End Sub

Public Sub Launch()
mnuShellView_Click
End Sub



Private Function MakeKey(Row, Col) As String
MakeKey = Format(Row, "0000") & Format(Col, "000_")
End Function





Public Sub SetWallpaper()
If Selected Then mnuPaper_Click
End Sub

Public Sub ShellView()
If Selected Then mnuShellView_Click
End Sub

Public Sub ShowProperties()
mnuProps_Click
End Sub

Private Sub ShowSelectedInfo()
StatusPanePrint Format(SelCount, "#### images selected.") & ", " & Format(m_Thumbs.SelSize / 1024#, "#,###,###.## Kb. total")
End Sub



Private Function ThumbPosition(Key As String) As ThumbPos
With ThumbPosition
    .Row = m_Thumbs.ThumbRow(Key)
    .Col = m_Thumbs.ThumbCol(Key)
End With
End Function

Private Sub PrepareToLoad()
'wipe the container.clean
Container.Cls

'and reset its size
Container.Height = Frame.ScaleHeight

'reset the corner positions
m_X = m_StartX
m_Y = THUMB_OFFSET

'and the row/col positions
m_Row = FIRST
m_Col = FIRST
End Sub



Private Sub StatusPanePrint(Optional ByVal s As String = vbNullString)
With StatusPane
    .Cls
    If s <> vbNullString Then
        .CurrentX = (.ScaleWidth - .TextWidth(s)) / 2
        .CurrentY = (.ScaleHeight - .TextHeight(s)) / 2
        StatusPane.Print s;
    End If
End With
End Sub



Public Sub View()
mnuView_Click
End Sub

Private Sub Container_DblClick()
If m_Loaded Then
    Dim Thumb   As clsThumb
    Dim s       As String
        
    For Each Thumb In m_Thumbs
        If Thumb.Clicked(m_ClickX, m_ClickY) Then
            s = Thumb.Path
            Exit For
        End If
    Next
    
    Set Thumb = Nothing

    RaiseEvent DblClick(s)
End If
End Sub

Private Sub Container_Resize()
With vs1
    .Min = 0
    .Max = IIf(Container.ScaleHeight > Frame.ScaleHeight, Container.ScaleHeight - Frame.ScaleHeight, Container.ScaleHeight)
    .Value = 0
    .Enabled = Container.ScaleHeight > Frame.ScaleHeight
    .SmallChange = 4
End With
End Sub

Private Sub File1_PathChange()
LoadThumbs
End Sub

Private Sub Frame_Resize()
With vs1
    .Move Frame.ScaleWidth - .Width, _
    0, _
    .Width, _
    Frame.ScaleHeight
End With
With Frame
    StatusPane.Move 0, .Height + 3, .ScaleWidth
    If m_Loaded Then
        Container.Move 0, 0, .ScaleWidth
    Else
        Container.Move 0, 0, .ScaleWidth, .ScaleHeight
    End If
End With
CalculateOffsets
End Sub


Private Function AddThumb(ByVal Index As Long) As clsThumb
Dim Thumb       As clsThumb
Dim Pic         As StdPicture
Dim srcDC       As Long
Dim hObj        As Long
Dim NewWidth    As Long
Dim NewHeight   As Long
Dim Ratio       As Single

'add a thumb to the collection
Set Thumb = m_Thumbs.Add(Container, m_ThumbSize, _
                         m_X, m_Y, m_Row, m_Col, _
                         MakeKey(m_Row, m_Col))

'set the particulars for the new thumb
With Thumb
    'the file name
    .Name = File1.List(Index)
    'the file path
    .Path = BuildPath(File1.Path, .Name)
    'store the size
    .FileSize = FileLen(.Path)
    'get the actual picture
    Set Pic = LoadPicture(.Path)
    
    'get the picture's dimensions
    .Width = ScaleX(Pic.Width, vbHimetric, vbPixels)
    .Height = ScaleY(Pic.Height, vbHimetric, vbPixels)
    
    'calculate a reduction ratio and reduced dimensions
    Ratio = IIf(.Width > .Height, m_ThumbSize / .Width, m_ThumbSize / .Height)
    NewWidth = .Width * Ratio
    NewHeight = .Height * Ratio
    
    'reduce the picture
    srcDC = CreateCompatibleDC(Container.hdc)
    hObj = SelectObject(srcDC, Pic.Handle)
    With pic1
        .BackColor = Container.BackColor
        .Cls
        SetStretchBltMode .hdc, COLORONCOLOR
        StretchBlt .hdc, _
                   (m_ThumbSize - NewWidth) / 2, _
                   (m_ThumbSize - NewHeight) / 2, _
                   NewWidth, _
                   NewHeight, _
                   srcDC, _
                   0, _
                   0, _
                   Thumb.Width, _
                   Thumb.Height, _
                   vbSrcCopy

        Set Thumb.Pic = .Image
    End With
    'cleanup
    SelectObject srcDC, hObj
    DeleteDC srcDC
    
    'display the newly created thumb
    .Render
End With

'update the location
'this call updates the left, top, row and col, as needed
GetNextPosition

'say thea a new thumb has been added
RaiseEvent Change(m_Thumbs.Count)

'return a reference to the newly created thumb
Set AddThumb = Thumb
Set Thumb = Nothing

End Function
Private Function BuildPath(ByVal FPath As String, ByVal FName As String) As String
If (FName = vbNullString) Or (FPath = vbNullString) Then Err.Raise 5

If Right$(FPath, 1) <> FOLDER_DELIMITER Then
    BuildPath = FPath & FOLDER_DELIMITER & FName
Else
    BuildPath = FPath & FName
End If
End Function
Public Property Get Count() As Long
Count = m_Thumbs.Count
End Property
Private Sub mnuCopy_Click()
Dim Thumb   As clsThumb
Dim Path    As String

SayBusy
For Each Thumb In m_Thumbs
    With Thumb
        If (.Selected Or .HasFocus) Then
            Path = .Path
            RaiseEvent Copy(Path)
            .Selected = False
        End If
    End With
Next

Set Thumb = Nothing

SayBusy False
End Sub

 
Private Sub mnuDelete_Click()
Dim Thumb       As clsThumb
Dim Key         As String
Dim Path        As String
Dim Permission  As Boolean
Dim Selected    As Boolean

SayBusy
For Each Thumb In m_Thumbs
    With Thumb
        Selected = (.Selected Or .HasFocus)
    End With
    
    If Selected Then
        With Thumb
            Path = .Path
            Key = .Key
        End With
    
        Permission = True
        'say delete the thumb
        RaiseEvent Delete(Permission, Path)
        'if allowed remove the thumb
        If Permission Then
            SetFocusTo NOT_DEFINED
            Set Thumb = Nothing
            m_Thumbs.Remove Key
            KeyCheck Key
        Else
            Thumb.Selected = False
        End If
    End If
Next

Set Thumb = Nothing

SayBusy False
End Sub

Private Sub mnuMove_Click()
Dim Thumb       As clsThumb
Dim Key         As String
Dim Path        As String
Dim Permission  As Boolean
Dim Selected    As Boolean

SayBusy
For Each Thumb In m_Thumbs
    With Thumb
        Selected = (.Selected Or .HasFocus)
    End With
    
    If Selected Then
        With Thumb
            Path = .Path
            Key = .Key
        End With
    
        Permission = True
        'say move the thumb
        RaiseEvent Move(Permission, Path)
        'if allowed remove the thumb
        If Permission Then
            SetFocusTo NOT_DEFINED
            Set Thumb = Nothing
            m_Thumbs.Remove Key
            KeyCheck Key
        Else
            Thumb.Selected = False
        End If
    End If
Next

Set Thumb = Nothing

SayBusy False
End Sub




Private Sub mnuPaper_Click()
Dim WallPaperPath As String

'set the path to the wallpaper bitmap
WallPaperPath = BuildPath(GetWindowsFolder, m_WallPaperName & ".bmp")


'create the wallpaper bitmap file
With pic1
    .AutoSize = True
    Set .Picture = LoadPicture(m_Thumbs(m_Current).Path)
    SavePicture .Image, WallPaperPath
    SystemParametersInfoA SPI_SETDESKWALLPAPER, 0, ByVal WallPaperPath, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
    'restore the picturebox
    .AutoSize = False
    .Cls
    .Move 0, 0, m_ThumbSize, m_ThumbSize
End With

End Sub

Public Function GetWindowsFolder() As String
Dim Buffer  As String * MAX_PATH
Dim Length  As Long

Length = GetWindowsDirectory(Buffer, MAX_PATH)

GetWindowsFolder = Left$(Buffer, Length)
End Function
Private Sub mnuProps_Click()
Dim Dlg As frmPicStats

Set Dlg = New frmPicStats
Dlg.PicturePath = m_Thumbs(m_Current).Path
Dlg.Show vbModal
Set Dlg = Nothing
End Sub


Private Sub mnuShellView_Click()
On Error Resume Next
ShellExecuteA UserControl.hWnd, _
             "open", _
             m_Thumbs(m_Current).Path, _
             vbNullString, _
             File1.Path, _
             SW_SHOW
End Sub

Private Sub mnuView_Click()
If Selected Then ViewTimer.Enabled = True
End Sub

Private Sub UserControl_Initialize()
Set m_Thumbs = New clsThumbs
End Sub

Private Sub UserControl_Resize()
With StatusPane
    .Height = .TextHeight("|") + 2
    Frame.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - .Height - 4
End With
If m_Loaded Then Refresh
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Container.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
   Container.BackColor = New_BackColor
   PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BORDER_STYLES
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal nuBs As BORDER_STYLES)
UserControl.BorderStyle() = nuBs
StatusPane.BorderStyle = nuBs
Frame.BorderStyle = nuBs
vs1.Appearance = IIf(nuBs = None, fsbFlat, fsb3D)
PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub Container_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Container_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Container_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Container_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Container_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub Container_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Thumb   As clsThumb
Dim Clicked As Boolean
Dim ix      As Long
Dim iy      As Long
Dim Offset  As Long


If m_Loaded Then

    'store the mouse co-ordinates for the dblclick event
    m_ClickX = x
    m_ClickY = y
    
    'was a thumb clicked?
    For Each Thumb In m_Thumbs
        If Thumb.Clicked(x, y) Then
            Clicked = True
            m_Current = Thumb.Key
            Exit For
        End If
    Next
        
    'yes , a thumb was clicked
    If Clicked Then
        
        'was it a right click?
        If Button = vbRightButton Then
            'if it was selected or has focus then show the popup menu
            If Thumb.Selected Or Thumb.HasFocus Then
                'show the popup menu
                mnuCopy.Visible = m_AllowCopy
                mnuDelete.Visible = m_AllowDelete
                mnuMove.Visible = m_AllowMove
                mnuView.Visible = (SelCount = 0)
                mnuProps.Visible = mnuView.Visible
                mnuPaper.Visible = m_AllowPaper
                mnuShellView.Visible = mnuView.Visible And m_AllowShellView
                PopupMenu mnuPopup
            Else
                DeSelectAll
            End If
        Else
            'no thumb has focus
            SetFocusTo NOT_DEFINED
            
            'if shift key was pressed
            If (Shift And vbShiftMask) = vbShiftMask Then
                'and a thumb was already selected
                'we have a range so select all thumbs in between
                If Not (m_Last = NOT_DEFINED) Then
                    Dim Start   As ThumbPos
                    Dim Finish  As ThumbPos
                    Dim Scol    As Long 'start column
                    Dim Ecol    As Long 'end column
               
                    'get the start/end locations
                    Start = ThumbPosition(m_Last)
                    Finish = ThumbPosition(m_Current)
                    
                    'put them descending order
                    If Start.Row > Finish.Row Then
                        Dim Holder As ThumbPos
                        
                        Holder = Start
                        Start = Finish
                        Finish = Holder
                    End If
                    
                    'for each row
                    For ix = Start.Row To Finish.Row
                        'and each column depending on the row
                        'start.row's starting column is start.col
                        'finish.row's ending column is finish.col
                        'all others do all columns
                        'unless start and finish row are the same
                        If ix = Start.Row And ix = Finish.Row Then
                            Scol = Start.Col
                            Ecol = Finish.Col
                        ElseIf ix = Start.Row Then
                            Scol = Start.Col
                            Ecol = m_Cols
                        ElseIf ix = Finish.Row Then
                            Scol = FIRST
                            Ecol = Finish.Col
                        Else
                            Scol = FIRST
                            Ecol = m_Cols
                        End If
                        For iy = Scol To Ecol
                            m_Thumbs(MakeKey(ix, iy)).Selected = True
                        Next 'col
                    Next 'row
                End If
           
                'set focus to the current thumb
                SetFocusTo m_Current
                ShowSelectedInfo
                
            
            'if control pressed select it
            ElseIf (Shift And vbCtrlMask) = vbCtrlMask Then
                Thumb.Selected = True
                
                'set focus to the current thumb
                SetFocusTo m_Current
                ShowSelectedInfo
            Else
                'set focus to the current thumb
                SetFocusTo m_Current
                
                'if the thumb is not already selected then deselect all
                If Not Thumb.Selected Then ClearSelected
                
                StatusPanePrint Thumb.Description
            End If
        End If
        m_Last = m_Current
    Else
        DeSelectAll
    End If
End If
Container.Refresh
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Private Function SetFocusTo(Key As String)

If Not (m_LastKey = NOT_DEFINED) Then
    m_Thumbs(m_LastKey).HasFocus = False
End If

m_LastKey = Key

If Not Key = (NOT_DEFINED) Then
    m_Thumbs(Key).HasFocus = True
End If

RaiseEvent SelectionChange((Key <> NOT_DEFINED))
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=File1,File1,-1,Path
Public Property Get Path() As String
Attribute Path.VB_Description = "Returns/sets the current path."
    Path = File1.Path
End Property

Public Property Let Path(ByVal New_Path As String)
If New_Path = NOT_DEFINED Then Exit Property

If Not FolderExists(New_Path) Then Err.Raise 76

File1.Path() = New_Path
PropertyChanged "Path"

End Property

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
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=File1,File1,-1,Pattern
Public Property Get Pattern() As String
Attribute Pattern.VB_Description = "Returns/sets a value indicating the filenames displayed in a control at run time."
    Pattern = File1.Pattern
End Property

Public Property Let Pattern(ByVal New_Pattern As String)
    File1.Pattern() = New_Pattern
    PropertyChanged "Pattern"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Dim NuCol       As clsThumbs
Dim Thumb       As clsThumb
Dim NuKey       As String

'reset client area
PrepareToLoad

'if nothing to do quit
If m_Thumbs.Count = 0 Or Not m_Loaded Then Exit Sub

'show the hourglass
SayBusy

'create a new collection of thumbs
Set NuCol = New clsThumbs

'for each thumb
For Each Thumb In m_Thumbs
    NuKey = MakeKey(m_Row, m_Col)
    With Thumb
        'preserve things
        If m_Current = .Key Then m_Current = NuKey
        If m_LastKey = .Key Then m_LastKey = NuKey
        'reset it's position and key
        .Key = NuKey
        .Left = m_X
        .Top = m_Y
        .Row = m_Row
        .Col = m_Col
        'draw it in the new position
        .Render
    End With
    
    'add it to the new collection
    'this is necessary to make sure the thumb's key
    'and the collection's key  match
    NuCol.AddThumb Thumb
    
    
    GetNextPosition
Next

Set m_Thumbs = NuCol
Set NuCol = Nothing

StatusPane.ForeColor = vbButtonText

RaiseEvent Change(Count)

SayBusy False
End Sub

Private Sub SayBusy(Optional ByVal ToState As Boolean = True)
m_Busy = ToState
Container.MousePointer = IIf(ToState, vbHourglass, vbDefault)
Frame.MousePointer = IIf(ToState, vbHourglass, vbDefault)
StatusPane.MousePointer = IIf(ToState, vbHourglass, vbDefault)
Container.Enabled = Not ToState
DoEvents

RaiseEvent IsBusy(ToState)
End Sub
Private Sub UserControl_Terminate()
Clear
Set m_Thumbs = Nothing

End Sub

Private Sub ViewTimer_Timer()
ViewTimer.Enabled = False
Dim f   As frmShow

Set f = New frmShow
f.Filepath = m_Thumbs(m_Current).Path
f.Show vbModal
Set f = Nothing
End Sub

Private Sub vs1_Change()
With vs1
    Container.Top = -.Value
    If .Value = .Min Then
        .Arrows = cc2RightDown
    ElseIf .Value = .Max Then
        .Arrows = cc2LeftUp
    Else
        .Arrows = cc2Both
    End If
End With
End Sub

Private Sub vs1_Scroll()
    RaiseEvent Scroll
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = DEFAULT_BACKCOLOR
    m_WallPaperName = DEFAULT_WALLPAPER_NAME
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    BackColor = PropBag.ReadProperty("BackColor", DEFAULT_BACKCOLOR)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Size = PropBag.ReadProperty("Size", Medium)
    Path = PropBag.ReadProperty("Path", "")
    File1.Pattern = PropBag.ReadProperty("Pattern", DEFAULT_PATTERN)
    m_WallPaperName = PropBag.ReadProperty("PaperName", DEFAULT_WALLPAPER_NAME)
    m_AllowCopy = PropBag.ReadProperty("AllowCopy", False)
    m_AllowMove = PropBag.ReadProperty("AllowMove", False)
    m_AllowDelete = PropBag.ReadProperty("AllowDelete", False)
    m_AllowPaper = PropBag.ReadProperty("AllowPaper", False)
    m_AllowShellView = PropBag.ReadProperty("AllowShellView", False)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Container.BackColor, DEFAULT_BACKCOLOR)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Size", m_Size, Medium)
    Call PropBag.WriteProperty("Path", File1.Path, "")
    Call PropBag.WriteProperty("Pattern", File1.Pattern, DEFAULT_PATTERN)
    Call PropBag.WriteProperty("PaperName", m_WallPaperName, DEFAULT_WALLPAPER_NAME)
    Call PropBag.WriteProperty("AllowCopy", m_AllowCopy, False)
    Call PropBag.WriteProperty("AllowMove", m_AllowMove, False)
    Call PropBag.WriteProperty("AllowDelete", m_AllowDelete, False)
    Call PropBag.WriteProperty("AllowPaper", m_AllowPaper, False)
    Call PropBag.WriteProperty("AllowShellView", m_AllowShellView, False)
End Sub


Public Property Get Size() As ThumbSizes
Size = m_Size
End Property

Public Property Let Size(ByVal sz As ThumbSizes)
If sz <> m_Size Then
    m_Size = sz
    m_ThumbSize = 64 + m_Size * 32
    pic1.Move 0, 0, m_ThumbSize, m_ThumbSize
    vs1.LargeChange = m_ThumbSize
    CalculateOffsets
    LoadThumbs
    PropertyChanged "Size"
End If
End Property


Public Property Get WallPaperName() As String
WallPaperName = m_WallPaperName
End Property

Public Property Let WallPaperName(ByVal s As String)
m_WallPaperName = s
PropertyChanged "PaperName"
End Property
