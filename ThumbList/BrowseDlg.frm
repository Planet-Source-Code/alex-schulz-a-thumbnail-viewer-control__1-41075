VERSION 5.00
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "CCRPFTV6.OCX"
Begin VB.Form frmBrowseDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse For Folder"
   ClientHeight    =   4560
   ClientLeft      =   3645
   ClientTop       =   4455
   ClientWidth     =   4740
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   316
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin CCRPFolderTV6.FolderTreeview FTV1 
      Height          =   1020
      Left            =   660
      TabIndex        =   5
      Top             =   2460
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   1799
      IntegralHeight  =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3465
      TabIndex        =   1
      Top             =   4050
      WhatsThisHelpID =   28444
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   2250
      TabIndex        =   0
      Top             =   4050
      WhatsThisHelpID =   28443
      Width           =   1125
   End
   Begin VB.Label labNewFTVInfo 
      Caption         =   "If adding a new FolderTreeview control, set the following properties in designtime: Name = ""FTV1"", IntegralHeight = False."
      Height          =   765
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   3015
   End
   Begin VB.Label labPrompt2 
      Height          =   300
      Left            =   165
      TabIndex        =   3
      Top             =   660
      Width           =   3600
   End
   Begin VB.Label labPrompt1 
      Caption         =   "Select the folder where you want to begin the search."
      Height          =   495
      Left            =   165
      TabIndex        =   2
      Top             =   165
      Width           =   3600
   End
   Begin VB.Menu mnuContext 
      Caption         =   "<no caption>"
      Visible         =   0   'False
      Begin VB.Menu mnuContextWhatsThis 
         Caption         =   "&What's This?"
      End
   End
End
Attribute VB_Name = "frmBrowseDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Brought to you by Brad Martinez
'   http://www.mvps.org/ccrp/
'   news://news.mvps.org/ccrp.foldertreeview

' =========================================================
' Duplicates the appearance and behavior "Browse For Folder" dialog box.
'
' - Code was developed using (and is formatted for) 8pt. MS Sans Serif font
' =========================================================

' Pertinent form designtime properties:

'   BorderStyle          = vbFixed Dialog
'   Caption                = "Browse For Folder"
'   ClipControls          = False
'   KeyPreview          = True   - for Shift+F10 "What's This?" help popup menu functionality
'   MaxButton            = False
'   MinButton             = False
'   ScaleMode           = vbPixels
'   ShowInTaskbar    = False
'   Visible                   = False
'   WhatsThisButton = True
'   WhatsThisHelp    = True

' Form position in twips (pixels):     Left = Owner + 1095 (73), Top = 1140 (76)
' Form size in twips (pixels):          Width = 4860 (324), Height = 4965 (331)

' (all coords in pixels)
' Prompt1 position & size:              10,   11, 240,   33
' Prompt2 position & size:              10,   44, 240,   20  (Top = Prompt1.Top + Prompt1.Height), Visible = False
' FTV no Prompt2 position & size: 10,   43, 298, 208
' FTV Prompt2 position & size:      10,   67, 298, 184  (Top = + 24, Height = - 24)
' OK button position & size:         150, 270,   75,  23
' Cancel button position & size:   231, 270,   75,  23   (Left = OK.Left + OK.Width + 6)

' =========================================================

' Holds the object reference to the dialog's owner form.
Private m_frmOwner As Form

' Holds the object reference to the currently selected browse dialog Folder,
' is set to Nothing if the dialog is cancelled.
Private m_SelectedFolder As CCRPFolderTV6.Folder

' Holds the object reference to the currently right clicked Control,
' used for "What's This?" context sensitive help.
Private m_ctlSelected As Control

' Holds onto the currently set App.HelpFile value when the dialog
' is show. This value is restored when the dialog is hidden.
Private m_sOldHelpFile As String

' The single exposed event
Event SelectionChanged(Folder As Folder)
Attribute SelectionChanged.VB_Description = "Occurs when a folders selection has changed in the dialog box. Allows setting of the Prompt2 propery."
'

Private Sub Form_Load()
  
  Set Icon = Nothing

  ScaleMode = vbPixels
  
  ' Set the postion and size of the form's controls.
  labPrompt1.Move 11, 11, 240, 33
  labPrompt2.Move 11, 44, 240, 20
  labPrompt2.Visible = False
  
  ' The ShowPrompt2 property sets the FTV size and position
  ShowPrompt2 = labPrompt2.Visible
  cmdOK.Move 150, 270, 75, 23
  cmdCancel.Move 231, 270, 75, 23

' For right click content menu help (the value below were obtained from
' viewing WinHelp messages in MS Help Workshop...)
' The App.HelpFIle property is set to "windows.hlp" each time the dialog
' is shown, and the old value is restored when the dialog is dismissed.
  cmdOK.WhatsThisHelpID = 28443
  cmdCancel.WhatsThisHelpID = 28444

  ' Change a few FolderTreeview properties from their default values
  ' to match the Browse For Folder dialog's behavior and appearance.
  With FTV1
' Read-only at runtime properties, must be set at designtime
'    .Name = "FTV1"
'    .IntegralHeight = False
    .AutoUpdate = False
    .HasLinesAtRoot = True
    .HideSelection = True
    .OverlayIcons = False
    .TabIndex = 0
    .ValidateSelection = False
' Override the browse dialog's default setting so we can
' set its SelecteFolder to everything each demo views...
'    .VirtualFolders = False
    .WhatsThisHelpID = 2224
  End With
  
End Sub

' =========================================================
' Properties

' use Let instead of Set so the client doesn't have to Set

Public Property Let Owner(frm As Form)
Attribute Owner.VB_Description = "Sets the object reference of the browse dialog's parent form. The dialog will be modal and displayed 73 pixels to the right and 76 pixels below the upper left corner of the owner form."
  Set m_frmOwner = frm
End Property

Public Property Get Prompt1() As String
Attribute Prompt1.VB_Description = "Returns/sets the text that is displayed as the first (upper) dialog prompt. This property setting is only evaluated before the dialog opens."
  Prompt1 = labPrompt1
End Property

Public Property Let Prompt1(sPrompt As String)
  labPrompt1 = sPrompt
End Property

Public Property Get Prompt2() As String
Attribute Prompt2.VB_Description = "Returns/sets the text that is displayed as the second (lower) dialog prompt. ShowPrompt2 must be set to True to display this prompt. Can be set in the SelectionChanged event while the dialog is shown."
  Prompt1 = labPrompt2
End Property

Public Property Let Prompt2(sPrompt As String)
  labPrompt2 = sPrompt
End Property

Public Property Get ShowPrompt2() As Boolean
Attribute ShowPrompt2.VB_Description = "Returns/sets a value that determines whether the second (lower) dialog prompt is visible. This property setting is only evaluated before the dialog opens."
  ShowPrompt2 = labPrompt2.Visible
End Property

' Sets the FTV's size and position with respect to the specified value

Public Property Let ShowPrompt2(Show As Boolean)
  labPrompt2.Visible = Show
  If Not Show Then
    FTV1.Move 10, 43, 298, 208   ' obscures prompt2
  Else
    FTV1.Move 10, 67, 298, 184   ' reveals prompt2
  End If
End Property

Public Property Get RootFolder() As String
Attribute RootFolder.VB_Description = "Returns/sets the root folder in the dialog box. This property setting is only evaluated before the dialog opens."
  RootFolder = FTV1.RootFolder
End Property

Public Property Let RootFolder(sFolder As String)
  FTV1.RootFolder = sFolder
End Property

Public Property Get PreSelectedFolder() As String
Attribute PreSelectedFolder.VB_Description = "Returns/sets the initially selected folder in the dialog box. This property setting is only evaluated before the dialog opens."
  SelectedFolder = FTV1.SelectedFolder
End Property

Public Property Let PreSelectedFolder(sFolder As String)
  FTV1.SelectedFolder = sFolder
End Property

' =========================================================
' The Browse method

Public Function Browse() As Boolean
Attribute Browse.VB_Description = "Displays the Browse for Folder dialog box. Returns True if a folder selection was made and the OK button was clicked, returns False if the Cancel button was clicked or the dialog was close from the control menu."
  Dim x As Single
  Dim y As Single
  
  ' Hold onto the old helpfile and set the new value for right
  ' click context menu help'
  m_sOldHelpFile = App.HelpFile
  App.HelpFile = "windows.hlp"
  
  ' Establish the dialog's initial non-owner screen display position.
  ' (screen and form position coordinants are always in twips).
  x = 73 * Screen.TwipsPerPixelX
  y = 76 * Screen.TwipsPerPixelY
  
  ' If m_frmOwner is set...
  If Not (m_frmOwner Is Nothing) Then
    ' Adjust the dialog's position relative to m_frmOwner's position
    x = y + m_frmOwner.Left
    y = y + m_frmOwner.top
  
    ' If any portion of the dialog will be dispayed off the screen,
    ' adjust the dialog's position to that screen edge.
    If x < 0 Then x = 0   ' left
    If y < 0 Then y = 0   ' top
    If x + Width > Screen.Width Then x = Screen.Width - Width      ' right
    If y + Height > Screen.Height Then y = Screen.Height - Height  ' bottom
  End If
  
  ' Set the dialog's position and size.
  Move x, y, 324 * Screen.TwipsPerPixelX, 331 * Screen.TwipsPerPixelY
    
  ' Restore the FolderTreeview to it's original state after it was previously
  ' shown (the dialog is hidden when it's dismissed, it is not unloaded).
  With FTV1
    
    ' Collapse all expanded folders except for the SelectedFolder
    ' and it's parents (this call is a hidden Folder object method).
    .SelectedFolder.CollapseAllButMe
    
    ' Raise the SelectionChanged event each time the
    ' dialog is shown (per Browse For Folder functionality).
    RaiseEvent SelectionChanged(.SelectedFolder)
    
    ' Scroll to the top if the FolderTreeview.
    .RootFolder.EnsureVisible
  
    ' Make sure the SelectedFolder is visible.
    .SelectedFolder.EnsureVisible
  
  End With
  
  ' The dialog shows itself modally, owned by m_frmOwner.
  Show vbModal, m_frmOwner
  
  ' Return how the dialog was closed. m_SelectedFolder is set to the
  ' folder selected by the user if the OK button was clicked, is set to
  ' Nothing if the dialog is cancelled or closed from the control menu.
  Browse = Not (m_SelectedFolder Is Nothing)

End Function

' Ensures that the FVT has the input focus when the dialog is shown.
' (it won't have the focus if the dialog was closed by a button click below)

Private Sub Form_Activate()
  FTV1.SetFocus
End Sub

' The form is hidden and kept loaded after the OK or Cancel buttons are clicked,
' allowing the properties of the folder the user selected to remain available.

Private Sub cmdOK_Click()
  ' Set the dialog's selected folder
  Set m_SelectedFolder = FTV1.SelectedFolder
  ' Restore the old helpfile.
  App.HelpFile = m_sOldHelpFile
  Hide
End Sub

Private Sub cmdCancel_Click()
  Set m_SelectedFolder = Nothing
  App.HelpFile = m_sOldHelpFile
  Hide
End Sub

' Unload the form and clear the module level object variables only
' when the application itself is unloading.

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set m_SelectedFolder = Nothing
  App.HelpFile = m_sOldHelpFile
  If UnloadMode = vbFormControlMenu Then
    Hide
    Cancel = True
  Else
    ' Unloading...
    Set m_frmOwner = Nothing
  End If
End Sub

' =========================================================
' The one and only raised event

' Occurs when the dialog is first shown and whenever a folder
' selection has changed in the FolderTreeviw.

Private Sub FTV1_SelectionChange(Folder As CCRPFolderTV6.Folder, PreChange As Boolean, Cancel As Boolean)
  
  ' If the selection has changed, raise the client's SelectionChanged event.
  If PreChange = False Then RaiseEvent SelectionChanged(Folder)

End Sub

Public Property Get SelectedFolder() As Folder
Attribute SelectedFolder.VB_Description = "Returns the Folder object selected by the user. This property is set after the dialog is closed by an OK button click."
  Set SelectedFolder = m_SelectedFolder
End Property

Public Function FolderString(nFolder As ftvSpecialFolderConstants) As String
Attribute FolderString.VB_Description = "Returns the display name or path of a special shell folder specified by a ftvSpecialFolderConstants value. Can be assigned directly to the RootFolder and PreSelectedFolder properties."
  FolderString = FTV1.GetSpecialFolderName(nFolder)
End Function

' =========================================================
' "What's This?" help context menu

Private Sub FTV1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then Call InvokePopupMenu(FTV1)
End Sub

Private Sub cmdOK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then Call InvokePopupMenu(cmdOK)
End Sub

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then Call InvokePopupMenu(cmdCancel)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = vbKeyShift And KeyCode = vbKeyF10 Then
    Call InvokePopupMenu(ActiveControl)
  End If
End Sub

Private Sub InvokePopupMenu(ctl As Control)
  Set m_ctlSelected = ctl
  PopupMenu mnuContext, vbPopupMenuRightButton
End Sub

Private Sub mnuContextWhatsThis_Click()
  m_ctlSelected.ShowWhatsThis
  Set m_ctlSelected = Nothing
End Sub
