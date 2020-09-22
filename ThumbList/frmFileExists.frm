VERSION 5.00
Begin VB.Form frmFileExists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Exists"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "No"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Do you want to replace the existing file?"
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
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "A file of the same name already exists in the destination folder!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmFileExists.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmFileExists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Result As Long

Public Property Get ResultCode() As Long
ResultCode = m_Result
End Property

Private Sub Command1_Click(Index As Integer)
m_Result = Index
Hide
End Sub


Private Sub Form_Initialize()
m_Result = CANCEL
End Sub

