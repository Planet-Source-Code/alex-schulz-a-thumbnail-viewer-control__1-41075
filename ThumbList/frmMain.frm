VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "*\A..\..\..\..\DOCUME~1\ALEXSC~1\Desktop\THUMBL~1\ThumbList.vbp"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "CCRPFTV6.OCX"
Begin VB.Form frmMain 
   Caption         =   "Prospect"
   ClientHeight    =   5955
   ClientLeft      =   1170
   ClientTop       =   2010
   ClientWidth     =   7455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":08CA
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   1200
   End
   Begin VB.Timer dragTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A1C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FFE
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15E0
            Key             =   "select"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BD2
            Key             =   "move"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21B4
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2796
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D78
            Key             =   "view"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":335A
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":393C
            Key             =   "thumb"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "select"
            Object.ToolTipText     =   "Select source folder"
            ImageKey        =   "select"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "move"
            Object.ToolTipText     =   "Move"
            ImageKey        =   "move"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "properties"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "view"
            Object.ToolTipText     =   "View"
            ImageKey        =   "view"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "thumb"
            Object.ToolTipText     =   "Thumbnail size"
            ImageKey        =   "thumb"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Small"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Medium"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Large"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Extra large"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
            Object.ToolTipText     =   "Preview on/off"
            ImageKey        =   "preview"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh View"
            ImageKey        =   "refresh"
         EndProperty
      EndProperty
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   60
         Width           =   2895
      End
   End
   Begin VB.PictureBox SplitterBar 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3120
      ScaleHeight     =   375
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox rightPane 
      HasDC           =   0   'False
      Height          =   5055
      Left            =   2880
      ScaleHeight     =   4995
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin ThumbList.ViewerPane ViewerPane1 
         Height          =   1335
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
         _extentx        =   2143
         _extenty        =   2355
         backcolor       =   -2147483643
         path            =   "C:\Documents and Settings\Alex Schulz\Desktop\ThumbList"
         allowcopy       =   -1  'True
         allowmove       =   -1  'True
         allowdelete     =   -1  'True
         allowpaper      =   -1  'True
         allowshellview  =   -1  'True
      End
   End
   Begin VB.PictureBox leftPane 
      HasDC           =   0   'False
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5115
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   480
      Width           =   2415
      Begin CCRPFolderTV6.FolderTreeview fv1 
         Height          =   1740
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3069
      End
      Begin VB.PictureBox pbPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2175
         TabIndex        =   5
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox Splitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2760
      MouseIcon       =   "frmMain.frx":3E3E
      MousePointer    =   99  'Custom
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SPLITTER_WIDTH = 2
Private Const APPNAME = "Prospect"
Private Const SOURCEPATH = "Source"
Private Const DESTPATH = "Destination"
Private Const THUMBSIZE = "size"

Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" ()


Private m_Y1            As Long
Private m_Y2            As Long
Private m_X1            As Long
Private m_X2            As Long
Private m_W1            As Long
Private m_W2            As Long
Private m_H1            As Long
Private m_H2            As Long
Private m_Sx            As Long
Private m_frmSizeX      As Long
Private m_PreviewOn     As Long
Private m_SourcePath    As String
Private m_DestPath      As String
Private m_PreviewPath   As String
Private Function CollisionCheck(ByRef FileSpec As String) As String
'check to see if file already exists
'if so, show a dialog asking what to do
'returns a valid path or ABORT (vbNullString) if user cancels

'assume the path is ok to start with
CollisionCheck = FileSpec

If FileExists(FileSpec) Then
    'if the file exists show the dialog
    
    Dim dlg As frmFileExists
    
    Set dlg = New frmFileExists
    dlg.Show vbModal, Me
    DoEvents
    
    'act according to user selection
    Select Case dlg.ResultCode
    
        Case Cancel
            'user cancelled out
            CollisionCheck = ABORT
            
        Case YES
            'user selected overwrite
            'so kill the existing file
            Kill FileSpec
            
        Case NO
            'user wants a new file
            Dim NewPath     As String
            Dim Extension   As String
            Dim Name        As String
            Dim Path        As String
            Dim ix          As Long
        
            'get the folder part fo the filespsec
            Path = GetParentFolder(FileSpec)
            'get the name of the file
            Name = GetFileName(FileSpec)
            'get the extension of the file
            Extension = GetExtensionName(Name)
            'get the file's name w/o the extension
            Name = GetBaseName(Name)
            
            'get a new filespec
            Do
                ix = ix + 1
                NewPath = BuildPath(Path, Name & "[" & CStr(ix) & "]" & Extension)
            Loop While FileExists(NewPath)
            
            CollisionCheck = NewPath
            
    End Select
    
    Unload dlg
    Set dlg = Nothing
End If
End Function


Private Sub SetDestinationPath(Optional p As String = vbNullString)
If p = vbNullString Then p = m_DestPath
m_DestPath = p
SaveSetting APPNAME, "Settings", DESTPATH, p
fv1.SelectedFolder = p

End Sub


Private Sub SetSize(ByVal NewSize As ThumbSizes)
Dim ix As Long


For ix = 1 To 4
    Toolbar1.Buttons("thumb").ButtonMenus(ix).Enabled = True
Next
Toolbar1.Buttons("thumb").ButtonMenus(NewSize + 1).Enabled = False
ViewerPane1.Size = NewSize
SaveSetting APPNAME, "Settings", THUMBSIZE, NewSize
End Sub

Private Sub ShowPreview(ByVal Show As Boolean)
Dim Pic         As StdPicture
Dim Width       As Long
Dim Height      As Long
Dim Ratio       As Single

pbPreview.Cls

If Show Then
    If ViewerPane1.Selected Then
        Set Pic = LoadPicture(m_PreviewPath)
            
        'get the picture's dimensions
        Width = ScaleX(Pic.Width, vbHimetric, vbPixels)
        Height = ScaleY(Pic.Height, vbHimetric, vbPixels)
        'calculate a reduction ratio and reduced dimensions
        Ratio = IIf(Width > Height, m_Sx / Width, m_Sx / Height)
        Width = Width * Ratio
        Height = Height * Ratio
        'show the picture
        pbPreview.PaintPicture Pic, (m_Sx - Width) \ 2, (m_Sx - Height) \ 2, Width, Height
        pbPreview.Refresh
    
        Set Pic = Nothing
    End If
End If
End Sub

Private Sub Form_Activate()
Static Done As Boolean

If Not Done Then
    
    SetSize (GetSetting(APPNAME, "Settings", THUMBSIZE, 2))
    SetDestinationPath
    SetSourcePath
    Done = True
End If
End Sub

Private Sub Form_Initialize()
m_SourcePath = GetSetting(APPNAME, "Settings", SOURCEPATH, fv1.GetSpecialFolderName(ftvDocuments))
If Not FolderExists(m_SourcePath) Then m_SourcePath = fv1.GetSpecialFolderName(ftvCommonDesktopDir)
If Not FolderExists(m_SourcePath) Then m_SourcePath = CurDir

m_DestPath = GetSetting(APPNAME, "Settings", DESTPATH, fv1.GetSpecialFolderName(ftvDocuments))
If Not FolderExists(m_DestPath) Then m_DestPath = fv1.GetSpecialFolderName(ftvCommonDesktopDir)
If Not FolderExists(m_DestPath) Then m_SourcePath = CurDir

m_PreviewOn = True
End Sub

Private Sub Form_Load()
    m_frmSizeX = &H7FFFFFFF
    Form_Resize
End Sub



Private Sub fv1_Click()
SetDestinationPath fv1.SelectedFolder
End Sub

Private Sub fv1_DragDrop(Source As Control, x As Single, y As Single)
ViewerPane1.MoveFiles
End Sub

Private Sub fv1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
ViewerPane1.DragIcon = LoadResPicture("DROP", vbResIcon)
End Sub


Private Sub leftPane_Resize()
With leftPane
    m_Sx = .ScaleWidth - 180
    Label1.Move 0, 0
    pbPreview.Move 90, .ScaleHeight - m_Sx - 90, m_Sx, m_Sx
    fv1.Move 90, 270, m_Sx, .ScaleHeight - pbPreview.Height - 360
End With

ShowPreview m_PreviewOn

End Sub


Private Sub pbPreview_DragOver(Source As Control, x As Single, y As Single, State As Integer)
ViewerPane1.DragIcon = LoadResPicture("NODROP", vbResIcon)
End Sub


Private Sub rightPane_Resize()
With rightPane
    ViewerPane1.Move 0, 0, .ScaleWidth, .ScaleHeight
End With
End Sub


Private Sub splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If m_frmSizeX <> &H7FFFFFFF Then
    If CLng(x) <> m_frmSizeX Then
        With Splitter
            .Move .Left + x, m_Y1, SPLITTER_WIDTH, ScaleHeight - 2
            SplitterBar.Move .Left, .top, .Width, .Height
        End With
        m_frmSizeX = CLng(x)
    End If
End If
End Sub

Private Sub splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SplitterBar.Visible = False
If m_frmSizeX <> &H7FFFFFFF Then
    With Splitter
        If Not CLng(x) = m_frmSizeX Then .Move .Left + x, m_Y1, SPLITTER_WIDTH, ScaleHeight - 2
        m_frmSizeX = &H7FFFFFFF
        .BackColor = &H8000000F
        If .Left > 60 And .Left < (ScaleWidth - 60) Then
            leftPane.Width = .Left - leftPane.Left
        ElseIf .Left < 60 Then
            leftPane.Width = 60
        Else
            leftPane.Width = ScaleWidth - 60
        End If
    End With
    Form_Resize
End If
End Sub

Private Sub splitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    Splitter.BackColor = vbBlack '&H808080
    SplitterBar.Visible = True
    m_frmSizeX = CLng(x)
Else
    If m_frmSizeX <> &H7FFFFFFF Then
        splitter_MouseUp Button, Shift, x, y
    End If
    m_frmSizeX = &H7FFFFFFF
End If
End Sub

Private Sub Form_Resize()
Const BAR_WIDTH = 2
On Error Resume Next

m_Y1 = Toolbar1.Height + BAR_WIDTH
m_H1 = ScaleHeight - Toolbar1.Height - BAR_WIDTH * 2
m_X1 = BAR_WIDTH
m_W1 = leftPane.Width
m_X2 = m_X1 + leftPane.Width + SPLITTER_WIDTH - 1
m_W2 = ScaleWidth - m_X2 - BAR_WIDTH
If Not WindowState = vbMinimized Then
    leftPane.Move m_X1 - 1, m_Y1, m_W1, m_H1
    rightPane.Move m_X2, m_Y1, m_W2 + 1, m_H1
    Splitter.Move m_X1 + leftPane.Width - 1, m_Y1, SPLITTER_WIDTH, m_H1
End If
End Sub



Private Sub Splitter_Resize()
With Splitter
    SplitterBar.Move .Left, .top, .Width, .Height
End With
End Sub







Private Sub Recycle(ByVal Target As String, Optional AllowUndo As Boolean = True)
Dim SHFileOp        As SHFILEOPSTRUCT
    
With SHFileOp
    .wFunc = FO_DELETE
    .pFrom = (Target & vbNullChar & vbNullChar)
    .fFlags = FOF_SILENT Or IIf(AllowUndo, FOF_ALLOWUNDO, 0)
End With

SHFileOperation SHFileOp
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
If ViewerPane1.Selected Then ShowPreview True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    
    Case "copy"
        ViewerPane1.copyfiles
    
    Case "delete"
        ViewerPane1.DeletFiles
        
    Case "move"
        ViewerPane1.MoveFiles
        
    Case "preview"
        m_PreviewOn = (Button.Value = tbrPressed)
        m_PreviewPath = ViewerPane1.SelectedItemPath
        ShowPreview m_PreviewOn
        
    Case "properties"
        ViewerPane1.ShowProperties
        
    Case "refresh"
        ViewerPane1.Refresh
        
    Case "select"
        Dim BrwsDlg As frmBrowseDlg
        Set BrwsDlg = New frmBrowseDlg

        With BrwsDlg
            .Prompt1 = "Select a source folder"
            .PreSelectedFolder = m_SourcePath
        
            If .Browse Then
                If m_SourcePath <> .SelectedFolder Then
                    m_SourcePath = .SelectedFolder
                    SetSourcePath
                End If
            End If
        End With
        Unload BrwsDlg
        Set BrwsDlg = Nothing
    
    Case "thumb"
        With ViewerPane1
            If .Size + 1 > [Extra Large] Then
                SetSize Small
            Else
                SetSize .Size + 1
            End If
        End With
    
    Case "view"
        ViewerPane1.View
End Select
End Sub



Private Sub SetSourcePath()
Dim p As String
Dim R As RECT


p = m_SourcePath
R.Right = 188
DrawText hDC, p, -1, R, DT_PATH_ELLIPSIS Or DT_SINGLELINE Or DT_MODIFYSTRING
txtPath.Text = p
txtPath.ToolTipText = m_SourcePath
SaveSetting APPNAME, "Settings", SOURCEPATH, m_SourcePath
ViewerPane1.Path = m_SourcePath
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
SetSize ButtonMenu.Index - 1
End Sub

Private Sub Toolbar1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
ViewerPane1.DragIcon = LoadResPicture("NODROP", vbResIcon)
End Sub

Private Sub ViewerPane1_Change(ByVal Count As Long)
Toolbar1.Buttons("refresh").Enabled = (Count > 0)
End Sub

Private Sub ViewerPane1_Copy(ByVal Target As String)
Dim Dest As String

Enabled = False
Dest = CollisionCheck(BuildPath(fv1.SelectedFolder, GetFileName(Target)))
If Dest <> ABORT Then FileCopy Target, Dest
Enabled = True
End Sub

Private Sub ViewerPane1_Delete(Permit As Boolean, ByVal Target As String)
Enabled = False
Permit = True
Recycle Target
Enabled = True
End Sub


Private Sub ViewerPane1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
ViewerPane1.DragIcon = LoadResPicture(IIf(Source.SelCount > 1, "DRAGALL", "DRAG"), vbResIcon)
End Sub

Private Sub dragTimer_Timer()
dragTimer.Enabled = False
With ViewerPane1
    .DragIcon = LoadResPicture(IIf(.SelCount > 1, "DRAGALL", "DRAG"), vbResIcon)
    .Drag
End With
End Sub
Private Sub ViewerPane1_IsBusy(ByVal State As Boolean)
Enabled = Not State
End Sub

Private Sub ViewerPane1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton And ViewerPane1.Selected Then dragTimer.Enabled = True
End Sub

Private Sub ViewerPane1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dragTimer.Enabled = False
ViewerPane1.Drag vbCancel
End Sub

Private Sub ViewerPane1_Move(Permit As Boolean, ByVal Target As String)
Dim Dest As String

Enabled = False
Dest = CollisionCheck(BuildPath(fv1.SelectedFolder, GetFileName(Target)))
Permit = (Dest <> ABORT)
If Permit Then Name Target As Dest
Enabled = True
End Sub

Private Sub ViewerPane1_SelectionChange(ByVal ItemSelected As Boolean)
With Toolbar1
    .Buttons("move").Enabled = ItemSelected
    .Buttons("copy").Enabled = ItemSelected
    .Buttons("delete").Enabled = ItemSelected
    .Buttons("view").Enabled = ItemSelected
    .Buttons("properties").Enabled = ItemSelected
End With

If m_PreviewOn Then
    m_PreviewPath = ViewerPane1.SelectedItemPath
    ShowPreview ItemSelected
End If
End Sub


