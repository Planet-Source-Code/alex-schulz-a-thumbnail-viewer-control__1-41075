VERSION 5.00
Object = "*\A..\..\..\..\DOCUME~1\ALEXSC~1\Desktop\THUMBL~1\ThumbList.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox leftPane 
      HasDC           =   0   'False
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      Begin VB.Frame Frame 
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
         Height          =   1575
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
         Begin VB.CommandButton cmdDestSelect 
            Caption         =   "Select"
            Height          =   300
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox txtDestPath 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1935
         End
         Begin VB.Image ImgDest 
            Height          =   720
            Left            =   1320
            Picture         =   "Form1.frx":0000
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Trash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   2175
         Begin VB.Image imgTrash 
            Height          =   720
            Left            =   600
            Picture         =   "Form1.frx":0ECA
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Source"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2175
         Begin VB.TextBox txtSrcPath 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmdSrcSelect 
            Caption         =   "Select"
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   900
         End
         Begin VB.Image imgSrc 
            Height          =   720
            Left            =   1320
            Picture         =   "Form1.frx":1D94
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.PictureBox pbPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2175
         TabIndex        =   3
         Top             =   4800
         Width           =   2175
      End
   End
   Begin VB.PictureBox rightPane 
      HasDC           =   0   'False
      Height          =   5055
      Left            =   4080
      ScaleHeight     =   4995
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin ThumbList.ViewerPane ViewerPane1 
         Height          =   1335
         Left            =   480
         TabIndex        =   1
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2355
         BackColor       =   -2147483643
         Path            =   "C:\Documents and Settings\Alex Schulz\Desktop\ThumbList"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

