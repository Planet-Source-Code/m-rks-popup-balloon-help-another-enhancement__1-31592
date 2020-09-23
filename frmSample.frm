VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Tips Form"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd5 
      Caption         =   "Balloon 5"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Balloon 4"
      Height          =   375
      Left            =   3900
      TabIndex        =   9
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Balloon 3"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Balloon 2"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Balloon 1"
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   1440
      Width           =   1155
   End
   Begin VB.TextBox txtInformation 
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmSample.frx":0000
      Top             =   1920
      Width           =   6195
   End
   Begin VB.CommandButton cmdExample2 
      Caption         =   "&Another Example"
      Height          =   375
      Left            =   4620
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdPopIt 
      Caption         =   "&Pop It Up!"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label lblMore 
      Caption         =   "The buttons below will also display some more styles of pop-up balloons. Click them!"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   6195
   End
   Begin VB.Label lblExample2 
      Caption         =   "For another, more practical example, click:"
      Height          =   255
      Left            =   1380
      TabIndex        =   3
      Top             =   3060
      Width           =   3135
   End
   Begin VB.Label lblInformation 
      Alignment       =   2  'Center
      Caption         =   "Click the button to pop up the sample balloon to see what it looks like:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd1.hWnd, WinRect)
BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmBalloon1 As New frmTip

frmBalloon1.SetBalloon "Balloon One", "This is a balloon! My properties are set " & _
    "so that I have a close button, no icon, and will not automatically close " & _
    "after a certain amount of time. My coordinates are also set so that I " & _
    "display by the ""Balloon 1"" button.", BalloonXY.x, BalloonXY.y, , True
    
frmBalloon1.Show , Me
Me.SetFocus
End Sub
Private Sub cmd2_Click()
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd2.hWnd, WinRect)
BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmBalloon2 As New frmTip

frmBalloon2.SetBalloon "Balloon Two", "I am Balloon 2. My properties are set " & _
    "so that I do not display a close (X) button, will auto-close after ten " & _
    "seconds, have a custom height and width, appear next to the ""Balloon 2"" " & _
    "button, and display a 9x-style ""!"" icon", _
    BalloonXY.x, BalloonXY.y, "!9", , 10000, 2500, 2100
    
frmBalloon2.Show , Me
Me.SetFocus
End Sub

Private Sub cmd3_Click()

Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd1.hWnd, WinRect)
BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmBalloon3 As New frmTip

frmBalloon3.SetBalloon "Balloon Three", "I am the Balloon 3. I am set to auto-" & _
    "close after fifteen seconds, display a close (X) button, show an (XP-style" & _
    ") ""i"" icon, appear lined up with the first button in this row (but be " & _
    "about centered on this form), and show using a custom font, Tahoma.", _
    BalloonXY.x, BalloonXY.y, "i", True, 15000, , , "Tahoma"
 
frmBalloon3.Show , Me
Me.SetFocus

End Sub

Private Sub cmd4_Click()
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd1.hWnd, WinRect)
BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmBalloon4 As New frmTip

frmBalloon4.SetBalloon "RTF File In Balloon", "", _
    BalloonXY.x, BalloonXY.y, "i", _
    True, 30000, 3000, 6000, "Tahoma", App.Path & "\test.rtf"
 
frmBalloon4.Show , Me
Me.SetFocus
End Sub

Private Sub cmd5_Click()
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords

Call GetWindowRect(cmd1.hWnd, WinRect)
BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY


Dim frmBalloon1 As New frmTip

' Set ballon button properties
With frmBalloon1.cmdLink
    .Caption = "Do something"
    .Enabled = True
    .Visible = True
End With

frmBalloon1.SetBalloon "Balloon Five", "This is a balloon with a Button! My properties are set " & _
    "so that I have a close button, no icon, and will not automatically close " & _
    "after a certain amount of time, and I display a Command Button at the bottom. " & _
    "The button can be used in conjunction with all the other features shown by the " & _
    "other examples.", BalloonXY.x, BalloonXY.y, , True, , , , , , True
    
frmBalloon1.Show , Me
Me.SetFocus

End Sub

Private Sub cmdExample2_Click()
frmSample2.Show
Unload Me
End Sub
Private Sub cmdPopIt_Click()
Dim WinRect As RECT         'These are used to hold some values we
Dim WinPoint As POINTAPI    'get during the API calls, and for
Dim BalloonXY As BalloonCoords 'storing the X and Y coordinates
                            'of the balloon that we pass when showing it

'This code is used to determine the position of the balloon. We (usually)
'want it to be displayed near some type of object, and since we need to
'set the balloon's coordinates relative to the screen, not the form, we
'need to determine the screen position of the control by which we want to
'place the balloon so it will show in the right spot.

'Get coordinates of the object on it that we want to
'display the ballooon by
Call GetWindowRect(cmdPopIt.hWnd, WinRect) 'When you use this code, replace
                                    'cmdPopIt with whatever control you
                                    'want to place the balloon by.
                                    
'This is multiplied by TwipsPerPixel because VB works
'with twips by default, but the API works in pixels. We'll be assigning
'the X and Y coordinates we get (which will be the coordiate for the lower
'left-hand corner of the control we chose above) to a BalloonXY (with .X
'and .Y properties) type object so we can easily use these coordinates
'later when we call to show the balloon.
'You can just assign them to two variables or whatever if you like, or
'just use the "formula" for figuring them directly when you call SetBalloon()
'instead of calculating them and then holding them in a variable.
BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmPopUpBalloon As New frmTip 'Make a new form based on frmTip

frmPopUpBalloon.SetBalloon "Sample Balloon", "This is a sample balloon to " & _
    "demonstrate the capabilities of the pop-up balloon/tooltips that you can " & _
    "use in your programs! They can include a title, multi-line text, an optional " & _
    "close button, automatiically close after a certain amount of time, " & _
    "an icon, and show in any font!  Don't forget you can click and drag the balloon by its title!", BalloonXY.x, BalloonXY.y, "i", True, _
     30000, , , "Tahoma"  'These preceeding lines set the properties
                        '(text, etc.) for the balloon
    
frmPopUpBalloon.Show , Me 'Show the balloon, with me as the owner

Me.SetFocus 'Since the balloon is a window (a form), showing it will
            'take focus away from this form, which it's called from,
            'and we don't want that to happen. We're working around it
            'by giving me focus after showing it. There IS away to show
            'a window without giving it focus via API, but I haven't
            'gotten that to work yet.
End Sub

Private Sub txtInformation_KeyPress(KeyAscii As Integer)
'Show a balloon if you try to type in the textbox (its Locked property
'is true, so you can't edit it anyway, but this will tell you if you try)

End Sub
