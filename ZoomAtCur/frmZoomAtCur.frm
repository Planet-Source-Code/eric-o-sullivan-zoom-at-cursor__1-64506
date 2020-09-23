VERSION 5.00
Begin VB.Form frmZoomAtCur 
   Caption         =   "Zoom"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   Icon            =   "frmZoomAtCur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timUpdate 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmZoomAtCur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
'                                MODULE DETAILS
'-------------------------------------------------------------------------------
'   Program Name:   Zoom
'  ---------------------------------------------------------------------------
'   Author:         Eric O'Sullivan
'  ---------------------------------------------------------------------------
'   Date:           25 February 2006
'  ---------------------------------------------------------------------------
'   Company:        CompApp Technologies
'  ---------------------------------------------------------------------------
'   Email:          diskjunky@hotmail.com
'  ---------------------------------------------------------------------------
'   Description:    This will zoom in the screen at the current position of
'                   the cursor
'  ---------------------------------------------------------------------------
'   Dependancies:   modSizeLimit.bas        clsBitmap.cls
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------


'all variables must be declared
Option Explicit


'-------------------------------------------------------------------------------
'                            MODULE LEVEL VARIABLES
'-------------------------------------------------------------------------------
Private mbmpBack            As clsBitmap            'holds a reference to the bitmap drawn on scree
Private Zoom                As Single               'holds the times size to zoom to


'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Private Sub Form_Load()
    'setup the initial values and size limitations
    
    Dim LastWidth           As Integer      'holds the last width of the form
    Dim LastHeight          As Integer      'holds the last height of the form
    
    
    'get the last screen sizes
    LastWidth = Val(GetSetting(App.Title, "ZoomAtCur", "LastWidth"))
    LastHeight = Val(GetSetting(App.Title, "ZoomAtCur", "LastHeight"))
    If (LastWidth = 0) Then
        LastWidth = Me.Width
    End If
    If (LastHeight = 0) Then
        LastHeight = Me.Height
    End If
    
    'reposition the form at the top right of the screen
    Call Me.Move(Screen.Width - LastWidth, 0, LastWidth, LastHeight)
    
    'set the resize limitations on the form
    Call SetResizeHook(Me.hWnd, 100, 50, Screen.Width / Screen.TwipsPerPixelX, 800, _
                       0, Screen.Width \ Screen.TwipsPerPixelX, _
                       0, Screen.Height \ Screen.TwipsPerPixelY, _
                       3, 3)
    
    'initialise the screen bitmap and the mouse
    Set mbmpBack = New clsBitmap
    Call mbmpBack.SetBitmap(Me.ScaleWidth \ Screen.TwipsPerPixelX, _
                            Me.ScaleHeight \ Screen.TwipsPerPixelY, _
                            Me.BackColor)
    
    'set the zoom scale (default *8)
    Zoom = 8#
    
    'get a screenshot at the mouse position
    Call GrabScreenShot
End Sub

Private Sub Form_Paint()
    'redraw the bitmap
    Call mbmpBack.Paint(Me.hdc)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'make sure that we clean up before exiting
    
    Set mbmpBack = Nothing
    
    Call ReleaseResizeHook(Me.hWnd)
    
    If (Me.WindowState <> vbMinimized) Then
        'save the last size
        Call SaveSetting(App.Title, "ZoomAtCur", "LastWidth", Me.Width)
        Call SaveSetting(App.Title, "ZoomAtCur", "LastHeight", Me.Height)
    End If
End Sub

Private Sub Form_Resize()
    'get another screenshot capable with the new size
    
    If mbmpBack Is Nothing Then
        Exit Sub
    End If
    
    Call mbmpBack.ReSize(Me.ScaleWidth \ Screen.TwipsPerPixelX, _
                         Me.ScaleHeight \ Screen.TwipsPerPixelY)
    Call GrabScreenShot
End Sub

Private Sub timUpdate_Timer()
    'update the screen
    If (Me.WindowState <> vbMinimized) Then
        Call GrabScreenShot
    End If
End Sub

Private Sub GrabScreenShot()
    'This will grab a screenshot at the mouse's current position and display it on the
    'form
    
    
    Static bmpGrab          As clsBitmap        'holds the screenshot
    
    Dim GrabWidth           As Integer          'holds the width of the screenshot to grab
    Dim GrabHeight          As Integer          'holds the height of the screenshot to grab
    Dim GrabLeft            As Integer          'holds the left position to grab at
    Dim GrabTop             As Integer          'holds the top position to grab at
    
    
    'do we need to create an object
    If bmpGrab Is Nothing Then
        Set bmpGrab = New clsBitmap
    End If
    
    'calculate the size of the bitmap we're to grab
    GrabWidth = (mbmpBack.Width / Zoom) + 1
    GrabHeight = (mbmpBack.Height / Zoom) + 1
    Call mbmpBack.MousePosition(GrabLeft, GrabTop)
    GrabLeft = GrabLeft - (GrabWidth \ 2)
    GrabTop = GrabTop - (GrabHeight \ 2)
    
    If (GrabWidth <> bmpGrab.Width) Or (GrabHeight <> bmpGrab.Height) Then
        'resize the bitmap to match
        Call bmpGrab.SetBitmap(GrabWidth, GrabHeight, vbBlack)
    End If
    
    'get the screen shot
    Call bmpGrab.Cls
    Call bmpGrab.GetScreenShot(GrabLeft, GrabTop)
    Call mbmpBack.PaintFrom(bmpGrab.hdc, GrabWidth, GrabHeight, 0, 0)
    
    'display the screenshot
    Call mbmpBack.Paint(Me.hdc)
End Sub
