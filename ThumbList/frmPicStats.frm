VERSION 5.00
Begin VB.Form frmPicStats 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Properties for "
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmPicStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPath 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   -480
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   -360
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   13
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   12
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   11
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblDimensions 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Last Opened:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Last Modified:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Created:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dimensions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   825
   End
End
Attribute VB_Name = "frmPicStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const IMAGE_SIZE = 118

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Private Const DT_SINGLELINE = &H20&
Private Const DT_PATH_ELLIPSIS = &H4000&
Private Const DT_MODIFYSTRING = &H10000

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private m_Path As String
Private Sub CompactPath(ByVal Path As String, ByRef tb As TextBox)
Dim R As RECT

R.Right = tb.Width - 4
DrawText tb.Parent.hdc, Path, -1, R, DT_PATH_ELLIPSIS Or DT_SINGLELINE Or DT_MODIFYSTRING
tb.Text = Path
End Sub
Public Property Let PicturePath(ByVal p As String)
Dim FileExtension       As String
Dim FileDateCreated     As String
Dim FileDateAccessed    As String
Dim FileDateModified    As String
Dim Filename            As String
Dim Filepath            As String
Dim ImageWidth          As Long
Dim ImageHeight         As Long
Dim FileSize            As Long
Dim Ratio               As Single
Dim fs                  As FileSystemObject
Dim f                   As File

On Error Resume Next
'ShowCursor False
Set fs = New FileSystemObject
With fs
    If Not .FileExists(p) Then Err.Raise 76
    Filepath = .GetParentFolderName(p)
    Filename = .GetFileName(p)
    m_Path = p
    Caption = Caption & Filename
    Set f = .GetFile(p)
        FileDateCreated = f.DateCreated
        FileDateAccessed = f.DateLastAccessed
        FileDateModified = f.DateLastModified
        FileSize = f.Size
    Set f = Nothing
End With
Set fs = Nothing

With pic1
    'load the picture
    Set .Picture = LoadPicture(p)
    'make sure it worked
    If Err.Number = 0 Then
        'record the size for later
        ImageWidth = .ScaleWidth
        ImageHeight = .ScaleHeight
                   
        'get the longest dimension
        Ratio = IIf(.ScaleWidth > .ScaleHeight, IMAGE_SIZE / .Width, IMAGE_SIZE / .Height)
        PaintPicture .Picture, 4, 4, Ratio * .ScaleWidth, Ratio * .ScaleHeight
        lblName.Caption = Filename
        CompactPath Filepath, txtPath
        lblPath.Caption = txtPath.Text
        lblSize.Caption = Format(CSng(FileSize / 1024), "#####.## Kb.")
        lblDimensions.Caption = CStr(ImageWidth) & " x " & CStr(ImageHeight)

        lblDate(0).Caption = FileDateCreated
        lblDate(1).Caption = FileDateModified
        lblDate(2).Caption = FileDateAccessed
    Else
        lblName.Caption = "Unable to load image!"
    End If
End With

End Property







Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    If x > 4 And y > 4 And x < IMAGE_SIZE And y < IMAGE_SIZE Then
        Dim f   As frmShow

        Set f = New frmShow
        f.Filepath = m_Path
        f.Show vbModal
        Set f = Nothing
    End If
End If
    
End Sub





