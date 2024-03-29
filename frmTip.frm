VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   127
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin Tips.LaVolpeButton cmdLink 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _extentx        =   2990
      _extenty        =   556
      btype           =   3
      bstyle          =   0
      optval          =   0   'False
      optmod          =   0   'False
      tx              =   ""
      enab            =   0   'False
      font            =   "frmTip.frx":000C
      coltype         =   1
      bcol            =   16763594
      fcol            =   0
      fcolo           =   0
      gstart          =   0
      gstop           =   16711680
      gstyle          =   0
      embossm         =   12632256
      embosss         =   16777215
      mptr            =   0
      micon           =   "frmTip.frx":0038
      align           =   1
      iconalign       =   0
      style           =   0
      orient          =   0
      iconsize        =   2
      showf           =   -1  'True
   End
   Begin RichTextLib.RichTextBox txtTip 
      Height          =   855
      Left            =   1380
      TabIndex        =   2
      Top             =   420
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1508
      _Version        =   393217
      BackColor       =   -2147483624
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmTip.frx":0056
   End
   Begin VB.Timer timAutoClose 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   600
      Width           =   195
   End
   Begin VB.Image imgDisplayIcon 
      Height          =   240
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgX_Up 
      Height          =   240
      Left            =   4080
      Picture         =   "frmTip.frx":00DB
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgX_Dn 
      Height          =   240
      Left            =   3480
      Picture         =   "frmTip.frx":041D
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgX 
      Height          =   240
      Left            =   3840
      Picture         =   "frmTip.frx":075F
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   1
      Left            =   960
      Picture         =   "frmTip.frx":0AA1
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmTip.frx":102B
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   2
      Left            =   0
      Picture         =   "frmTip.frx":15B5
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   2
      Left            =   480
      Picture         =   "frmTip.frx":1B3F
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   1
      Left            =   240
      Picture         =   "frmTip.frx":20C9
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmTip.frx":2653
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009EF5F3&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'All variables must be declared

Dim XY() As POINTAPI

Private Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long 'Used to round the corners of the form
    
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, _
    ByVal nPolyFillMode As Long) As Long 'Used to round corners of form

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
ByVal hRgn As Long, ByVal bRedraw As Long) As Long


Public Sub RoundCorners()
Attribute RoundCorners.VB_Description = "Rounds the corners of the form via API to create the tooltip effect"
    Me.ScaleMode = vbPixels
    mlWidth = Me.ScaleWidth
    mlHeight = Me.ScaleHeight
    
    
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, _
                (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), _
                25, 25), _
                True
End Sub

Private Sub cmdLink_Click()
    MsgBox "So you want me to do something...well how about Quitting."
    End
End Sub

Private Sub Form_Click()
'Hide me after I'm clicked on
HideBalloon
End Sub

Private Sub Form_Load()
RoundCorners ' Round the corners of this form to make it look "tool-tippy"
End Sub
Private Sub Form_Resize()

  txtTip.Move 8, lblTitle.Height + 10, Me.ScaleWidth - 20, Me.ScaleHeight _
  - lblTitle.Height - 20 - cmdLink.Height - (cmdLink.Height * 0.75)
  
  lblX.Move (Me.ScaleWidth - lblX.Width) - 13, 5 'lblX.Height - 10  '- 2
  imgX.Move (Me.ScaleWidth - lblX.Width) - 15, 2 'lblX.Height - 13  '- 5
  imgX_Dn.Move (Me.ScaleWidth - lblX.Width) - 15, 2 '  lblX.Height - 13 ' - 5
  imgX_Up.Move (Me.ScaleWidth - lblX.Width) - 15, 2 'lblX.Height - 13 '- 5
  
  imgDisplayIcon.Move 10, 2
  
  'Now, resize the title label's width to fit the balloon size:
  lblTitle.Move 0, 1, Me.ScaleWidth
  
  Me.Cls
  
  Me.DrawWidth = 1
  Me.FillStyle = 0
  Me.Line (lblTitle.Left, lblTitle.Top)-(lblTitle.Width, lblTitle.Height), &H9EF5F3, BF
  
  Me.FillStyle = 1
  Me.DrawWidth = 2
  Me.ForeColor = vbBlack
  RoundRect Me.hdc, 0, 0, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)

End Sub

Private Sub imgDisplayIcon_Click()
  ' Hide this balloon if I'm clicked
  HideBalloon
End Sub

Private Sub imgX_Click()
  HideBalloon
End Sub

Private Sub imgX_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then imgX.Picture = imgX_Dn.Picture
End Sub

Private Sub imgX_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then imgX.Picture = imgX_Up.Picture
End Sub

Private Sub lblTitle_Click()
'Hide me after I'm clicked on
'HideBalloon
End Sub
Public Sub SetBalloon(sTitle As String, sText As String, lPosX As Long, lPosY _
    As Long, Optional sIcon As String, Optional bShowClose As Boolean = False, _
    Optional lAutoCloseAfter As Long = 0, Optional lHeight As Long = 1620, _
     Optional lWidth As Long = 4680, Optional sFont = "MS Sans Serif", Optional sRTFFilename As String _
     , Optional boolButton As Boolean = False)
    
'Arguments for this Sub are explained below. What this Sub does is
'set the properties for the balloon to be displayed--text, title, etc.
'After (or right before) setting the properties, you must show the
'balloon yourself by calling <form_name>.Show
'For example, if this "template" form is frmTip, you can create a new
'instance of frmTip by doing:
'   Dim frmMyTip as New frmTip
'and then calling frmMyTip.SetBalloon using the values you want, as in:
'   frmMyTip.SetBalloon "Sample Title", "Sample Text"
'and going on with the arguments as needed (see below and the declaration
'for this Sub above).
'Then, to show the balloon, call
'   frmMyTip.show

'Here's what the arguments for this Sub do:

'sTitle: The bold title to appear above the text on the balloon (Required)

'sText: Text of balloon (Required)

'lPosX and lPosY: The horizontal and vertical, respectively, positions to
'                 show the ballon at (Required)

'sIcon: The icon to be displayed on the balloon, similar to the messagebox's.
'       They're an "i", "x", or "!". (No question mark here; you can't ask
'       on a balloon, can you?) To specifiy, pass either "i", "x", or "!" as
'       the argument, e.g., SetBalloon("Title", "Text", "!" ...
'       For none, don't pass anything. And, they'll use the XP-style icons
'       by default; to use 9x-looking icons instead, specify "i9", "x9", or "!9"
'       Look at the tooltip form (frmTip, in my example project) to see what
'       they look like; you should see the difference, but they're quite similar--
'       the XP ones just look more colorful and 3D-ish (Optional)

'bShowClose: Whether or not to show the "X" close button the user can
'            press to close the balloon. If there, click to close the
'            balloon; if it's not there (or if it is) clicking anywhere
'            in the balloon will close it. (Optional)

'lAutoCloseAfter: Specifies the amount of time (in milliseconds) after
'                 which to automatically close the balloon. Setting it
'                 to 0 will make it not automatically close.
'                 E.g., 10,000 is ten seconds. (Optional)
'lHeight and lWidth: The width and height that you want the balloon to have.
'                    It 's optional, and it will default to a "normal" size.
'                    If you have a long message, increasing the height should
'                    be good, although you can increase the width if you want, too
                     
'
'sFont: The font the text will appear in, defaulting to MS Sans Serif.
'       The other normal choice would be Tahoma, which is gives it a
'       "new" look, but some earlier Windows 9x versions may not have
'       it (Optional)
'
'boolButton: Indicates if the command button is to be shown or not. If it is then
'            the button can be used for any normal button function - close, link to
'            something. The LaVolpeButton User Control has been add to allow the button
'            style be set to XP. The Control also allows other styles (Optional)


  'Setting TITLE AND CAPTION on tip:
  lblTitle.Caption = sTitle
  If sText <> "" Then txtTip.Text = sText
  If sRTFFilename <> "" Then txtTip.FileName = sRTFFilename
  
  'Setting the X AND Y POSITIONS:
  Me.Move lPosX, lPosY
  
  'Setting the ICON:
  'First, convert the case to all lower; that way, since all Select Case
  'statements below use lowercase for identification
  sIcon = LCase(sIcon)
  
  Select Case sIcon
      Case "i": 'The "i" icon, XP-style (default)
          Me.imgDisplayIcon.Picture = Me.imgIconXP(0).Picture
          
      Case "i9": 'The "i" icon, 9x/Me-style
          imgDisplayIcon.Picture = imgIcon(0).Picture
          
      Case "x": 'The "x" icon, XP-style
          imgDisplayIcon.Picture = imgIconXP(1).Picture
          
      Case "x9": 'The "x" icon, 9x/Me-style
          imgDisplayIcon.Picture = imgIcon(1).Picture
          
      Case "!": 'The "!" icon, XP-style
          imgDisplayIcon.Picture = imgIconXP(2).Picture
          
      Case "!9": 'The "!" icon, 9x-style
          imgDisplayIcon.Picture = imgIcon(2).Picture
          
      Case Else: 'Use no icon
          Me.imgDisplayIcon.Visible = False
          Me.lblTitle.Left = imgDisplayIcon.Left 'Move title over so it looks right
  End Select
          
  'Showing/not showing THE X BUTTON:
  If bShowClose = False Then ' Then don't show the X button
      Me.imgX.Visible = False
      Me.lblX.Visible = False
  End If
  If bShowClose = True Then ' Then make the X button visible
      Me.imgX.Visible = True
      Me.lblX.Visible = True
  End If
  
  'Enabling/disabling AUTO-CLOSE:
  If lAutoCloseAfter = 0 Then ' Then we don't need to auto-close, so ...
      Me.timAutoClose.Enabled = False ' Just make sure the auto-close timer
                                      ' is disabled, since we shouldn't auto-close
  Else    ' Then we DO need to auto-close
      Me.timAutoClose.Interval = lAutoCloseAfter ' Set timer's interval so it will
                                                 ' auto-close at the right time, and...
      Me.timAutoClose.Enabled = True 'Enable the timer, so it will go off and auto-close
  End If
  
  
  'Setting HEIGHT AND WIDTH: Modified by MarkS
  If boolButton Then
    ' increase the size of the tip balloon to accomodate the button. Why *12 - don't know!
    ' it just looks best and it works.
    lHeight = lHeight + (cmdLink.Height * 12)
  Else
    Me.cmdLink.Height = 0 ' remove button from calculations where it is not going to be shown
  End If
    
  Me.Width = lWidth
  Me.Height = lHeight
  RoundCorners

  'Setting the FONT:
  Me.Font = sFont
  If sRTFFilename = "" Then Me.txtTip.Font = sFont
  Me.lblTitle.Font = sFont


End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  EasyMove Me
End Sub

Private Sub lblX_Click()
  HideBalloon
End Sub

Private Sub lblX_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then imgX.Picture = imgX_Dn.Picture
End Sub

Private Sub lblX_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then imgX.Picture = imgX_Up.Picture
End Sub


Private Sub timAutoClose_Timer()
' This timer is used to automatically close the balloon, if needed,
' after the specified number of milliseconds

  HideBalloon  'Calls HideBalloon(), which hides the balloon
End Sub
Public Sub HideBalloon()
'HideBalloon() is used to manually hide the balloon and by the
'balloon itself to hide itself when needed
  Unload Me
End Sub

Private Sub txtTip_Click()
  If lblX.Visible = False Then HideBalloon
End Sub

Private Sub txtTip_DblClick()
  HideBalloon
End Sub
