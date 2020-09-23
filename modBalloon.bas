Attribute VB_Name = "modBalloon"
Option Explicit

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
        lpRect As RECT) As Long 'Used for getting positions of objects/forms
                                'to place balloons correctly

Public Type RECT   'Also used to store values for positions of balloons
   Left As Long    'after using the API to determine where
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Used to move around by caption
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Used to draw the ellipse on the form
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
    ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'Used to create the regiod around the form to shape it
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, _
        lpPoint As POINTAPI) As Long 'Also used for getting positions of
                                     'objects/forms we want to place the
                                     'balloons by
Public mlWidth As Long
Public mlHeight As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type BalloonCoords 'Used to store X and Y coordinates of balloon
    X As Long 'after using API and math operations to figure exact
    Y As Long 'coordinates regarding where to place itself
End Type

Public Sub EasyMove(frm As Form)
  If frm.WindowState <> vbMaximized Then
    ReleaseCapture
    SendMessage frm.hWnd, &HA1, 2, 0&
  End If
End Sub

