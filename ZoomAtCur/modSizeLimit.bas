Attribute VB_Name = "modSizeLimit"
'-------------------------------------------------------------------------------
'                                MODULE DETAILS
'-------------------------------------------------------------------------------
'   Program Name:   General Use
'  ---------------------------------------------------------------------------
'   Author:         Eric O'Sullivan
'  ---------------------------------------------------------------------------
'   Date:           25 May 2004
'  ---------------------------------------------------------------------------
'   Company:        CompApp Technologies
'  ---------------------------------------------------------------------------
'   Contact:        DiskJunky@hotmail.com
'  ---------------------------------------------------------------------------
'   Description:    This will limit the size a single window can be resize to
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'require variable declaration
Option Explicit


'-------------------------------------------------------------------------------
'                              API DECLARATIONS
'-------------------------------------------------------------------------------
Private Declare Function DefWindowProc _
        Lib "user32" _
        Alias "DefWindowProcA" _
            (ByVal hWnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Long, _
             ByVal lParam As Long) _
             As Long

Private Declare Function CallWindowProc _
        Lib "user32" _
        Alias "CallWindowProcA" _
            (ByVal lpPrevWndFunc As Long, _
             ByVal hWnd As Long, _
             ByVal Msg As Long, _
             ByVal wParam As Long, _
             ByVal lParam As Long) _
             As Long
             
Private Declare Function SetWindowLong _
        Lib "user32" _
        Alias "SetWindowLongA" _
            (ByVal hWnd As Long, _
             ByVal nIndex As Long, _
             ByVal dwNewLong As Long) _
             As Long

Private Declare Sub CopyMemoryToMinMaxInfo _
        Lib "kernel32" _
        Alias "RtlMoveMemory" _
            (hpvDest As MINMAXINFO, _
             ByVal hpvSource As Long, _
             ByVal cbCopy As Long)

Private Declare Sub CopyMemoryFromMinMaxInfo _
        Lib "kernel32" _
        Alias "RtlMoveMemory" _
            (ByVal hpvDest As Long, _
             hpvSource As MINMAXINFO, _
             ByVal cbCopy As Long)

Private Declare Sub CopyMemory _
        Lib "kernel32" _
        Alias "RtlMoveMemory" _
            (lpvDest As Any, _
             lpvSource As Any, _
             ByVal cbCopy As Long)


'-------------------------------------------------------------------------------
'                           MODULE LEVEL CONSTANTS
'-------------------------------------------------------------------------------
Private Const GWL_WNDPROC       As Long = -4
Private Const WM_GETMINMAXINFO  As Long = &H24
Private Const WM_DESTROY        As Long = &H2
Private Const WM_SIZING         As Long = &H214
Private Const WM_MOVING         As Long = &H216&
Private Const WM_ENTERSIZEMOVE  As Long = &H231&
Private Const WM_EXITSIZEMOVE   As Long = &H232&
Private Const WM_ACTIVATE       As Long = &H6
Private Const WM_SIZE           As Long = &H5
Private Const WM_MOVE           As Long = &HF012


'-------------------------------------------------------------------------------
'                           USER DEFINED TYPES
'-------------------------------------------------------------------------------
Private Type POINTAPI
    X               As Long
    Y               As Long
End Type

Private Type MINMAXINFO
    ptReserved      As POINTAPI
    ptMaxSize       As POINTAPI
    ptMaxPosition   As POINTAPI
    ptMinTrackSize  As POINTAPI
    ptMaxTrackSize  As POINTAPI
End Type

Private Type TypeFormDetails
    hWnd            As Long         'holds a window handle
    lngPrevhWnd     As Long         'holds a reference to the previous window handle
    intMinWidth     As Integer      'holds the minimum width of the form
    intMinHeight    As Integer      'holds the minimum height of the form
    intMaxWidth     As Integer      'holds the maximum width of the form
    intMaxHeight    As Integer      'holds the maximum height of the form
    intLeftBound    As Integer      'holds the left most position the form can go to
    intRightBound   As Integer      'holds the right most position the form can go to
    intTopBound     As Integer      'holds the right most position the form can go to
    intBottomBound  As Integer      'holds the bottom most position the form can go to
    intSnapToBoundX As Integer      'holds the amount of pixels where the window snaps to its horizontal boundries
    intSnapToBoundY As Integer      'holds the amount of pixels where the window snaps to its veritcal boundries
End Type

Private Type Rect
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type


'-------------------------------------------------------------------------------
'                           MODULE LEVEL VARIABLES
'-------------------------------------------------------------------------------
Private mudtForm()      As TypeFormDetails  'holds details about the form
Private mintNumHooks    As Integer          'holds the number of references we are keeping track of

'-------------------------------------------------------------------------------
'                                 PROCEDURES
'-------------------------------------------------------------------------------
Private Sub InitHooks(Optional ByVal blnReset As Boolean = False)
    'This will initialise the arrays
    
    Static blnStarted   As Boolean      'flags if we have initialised the arrays
    
    'do we need to initialise the arrays
    If (Not blnStarted) Or (blnReset) Then
        mintNumHooks = 0
        ReDim mudtForm(0)
        
        'we have initialised the arrays at least once
        blnStarted = True
    End If  'do we need to initialise the arrays
End Sub

Public Sub ReleaseAllResizeHooks()
    'This will unhook all open hooks. Usually called at the end of a program
    
    Dim intCounter      As Integer      'used to cycle through the list of hooks
    
    Do While (mintNumHooks > 0)
        Call ReleaseResizeHook(mudtForm(mintNumHooks - 1).hWnd)
    Loop
End Sub

Public Sub SetResizeHook(ByVal hWnd As Long, _
                         Optional ByVal intMinWidth As Integer, _
                         Optional ByVal intMinHeight As Integer, _
                         Optional ByVal intMaxWidth As Integer, _
                         Optional ByVal intMaxHeight As Integer, _
                         Optional ByVal intLeftBound As Integer = -32767, _
                         Optional ByVal intRightBound As Integer = 32767, _
                         Optional ByVal intTopBound As Integer = -32767, _
                         Optional ByVal intBottomBound As Integer = 32767, _
                         Optional ByVal intSnapToBoundX As Integer = 0, _
                         Optional ByVal intSnapToBoundY As Integer = 0)
    'Start subclassing the specified window
    
    'validate the parameters
    If (hWnd = 0) Then
        Exit Sub
    End If
    
    'check the minimum size
    If (intMinWidth < 0) Then
        intMinWidth = 0
    End If
    If (intMinHeight < 0) Then
        intMinHeight = 0
    End If
    
    'check the maximum size
    If (intMaxWidth < 0) Then
        intMaxWidth = 0
    End If
    If (intMaxHeight < 0) Then
        intMaxHeight = 0
    End If
    
    'if we have already subclassed a window, then close that subclass before starting a new one
    If AlreadyHooked(hWnd) Then
        Exit Sub
    End If
    
    'create a new element to hold this hook in the array
    ReDim Preserve mudtForm(mintNumHooks)
    
    With mudtForm(mintNumHooks)
        'set the window and size parameters
        .hWnd = hWnd
        .intMinWidth = intMinWidth
        .intMinHeight = intMinHeight
        .intMaxWidth = intMaxWidth
        .intMaxHeight = intMaxHeight
        .intLeftBound = intLeftBound
        .intRightBound = intRightBound
        .intTopBound = intTopBound
        .intBottomBound = intBottomBound
        .intSnapToBoundX = intSnapToBoundX
        .intSnapToBoundY = intSnapToBoundY
        
        'create the hook
        .lngPrevhWnd = SetWindowLong(.hWnd, GWL_WNDPROC, AddressOf WindowProc)
    End With    'mudtform(mintNumHooks)
    mintNumHooks = mintNumHooks + 1
End Sub

Private Function AlreadyHooked(ByVal hWnd As Long) _
                               As Boolean
    'This will return True if the window is already hooked
    
    Dim intCounter      As Integer      'used to cycle through the hooks
    Dim blnFound        As Boolean      'flags if the hook is in the array
    
    'validate the parameter
    If (hWnd = 0) Then
        Exit Function
    End If
    
    'look for the hook
    For intCounter = 0 To (mintNumHooks - 1)
        If (hWnd = mudtForm(intCounter).hWnd) Then
            'we found the window
            blnFound = True
            Exit For
        End If
    Next intCounter
    
    'return whether or not we found it
    AlreadyHooked = blnFound
End Function

Private Function GetHookPos(ByVal hWnd As Long) _
                            As Integer
    'This will return the array position of the window if it exists in the array. If it doesn't then -1 is
    'returned
    
    Dim intPos      As Integer      'holds the array position of the window
    Dim intCounter  As Integer      'used to cycle through the hooks
    
    intPos = -1
    
    'valdiate the parameter
    If (hWnd = 0) Then
        GetHookPos = intPos
        Exit Function
    End If
    
    'look for the hook
    For intCounter = 0 To (mintNumHooks - 1)
        If (hWnd = mudtForm(intCounter).hWnd) Then
            'we found the window
            intPos = intCounter
            Exit For
        End If
    Next intCounter
    
    'return the position if we found it
    GetHookPos = intPos
End Function

Public Sub ReleaseResizeHook(ByVal hWnd As Long)
    'stop subclassing a form. This must be called before exiting the application - especially in the ide as
    'it will cause it to crash
    
    Dim lngTemp     As Long         'holds the returned value from an api call
    Dim intPos      As Integer      'holds the array position of the window
    Dim intCounter  As Integer      'used to cycle through the arrays
    
    'is this parameter valid
    If (hWnd = 0) Then
        Exit Sub
    End If
    
    'have we hooked this window
    If Not AlreadyHooked(hWnd) Then
        Exit Sub
    End If
    
    'get the array position of the window
    intPos = GetHookPos(hWnd)
    
    'Cease subclassing.
    lngTemp = SetWindowLong(mudtForm(intPos).hWnd, GWL_WNDPROC, mudtForm(intPos).lngPrevhWnd)
    mudtForm(intPos).hWnd = 0
    
    If (mintNumHooks > 1) Then
        'copy over the removed windows details
        For intCounter = intPos To (mintNumHooks - 2)
            'copy the next element down on top of this one
            mudtForm(intCounter) = mudtForm(intCounter + 1)
        Next intCounter
        
        'shrink the array
        mintNumHooks = mintNumHooks - 1
        ReDim Preserve mudtForm(mintNumHooks - 1)
        
    Else
        'wipe the last element
        Call InitHooks(True)
    End If
End Sub

Public Function WindowProc(ByVal hw As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) _
                           As Long
    'process resize messages for the window
    
    Dim MinMax      As MINMAXINFO   'holds the size details of the window
    Dim intPos      As Integer      'holds the array position of the window
    Dim udtBounds   As Rect         'holds the bounds of the moving area
    Dim lngResult   As Long         'holds any returned value from an api call
    Dim blnChanged  As Boolean      'flags if we have changed any details
    
    'Check for request for min/max window sizes.
    Select Case uMsg
    Case WM_GETMINMAXINFO
        
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)
        
        'get the array position of the window
        intPos = GetHookPos(hw)
        If (intPos >= 0) Then
            
            With mudtForm(intPos)
                
                blnChanged = False
                
                'Specify new minimum size for window.
                If (.intMinWidth > 0) Then
                    MinMax.ptMinTrackSize.X = .intMinWidth
                    blnChanged = True
                End If
                If (.intMinHeight > 0) Then
                    MinMax.ptMinTrackSize.Y = .intMinHeight
                    blnChanged = True
                End If
        
                'Specify new maximum size for window.
                If (.intMaxWidth > 0) Then
                    MinMax.ptMaxTrackSize.X = .intMaxWidth
                    blnChanged = True
                End If
                If (.intMaxHeight > 0) Then
                    MinMax.ptMaxTrackSize.Y = .intMaxHeight
                    blnChanged = True
                End If
            End With    'mudtForm(intPos)
            
            If blnChanged Then
                'Copy local structure back.
                CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
            End If
        End If  'have we subclassed this window
        
        'pass the message onto the next window
        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
        
    Case WM_MOVING, WM_MOVE
        'prevent the form from moving outside the specified bound
        
        ' Form is moving:
        CopyMemory udtBounds, ByVal lParam, Len(udtBounds)
        
        blnChanged = False
        
        'get the array position of the window
        intPos = GetHookPos(hw)
        If (intPos >= 0) Then
                
            'make sure that the form is within the bounds specified
            With mudtForm(intPos)
                'have we gone too far horizontally
                If (udtBounds.Right > .intRightBound) Then
                    'we have gone too far right
                    udtBounds.Left = udtBounds.Left - (udtBounds.Right - .intRightBound)
                    udtBounds.Right = .intRightBound
                    blnChanged = True
                    
                End If
                If (udtBounds.Left < .intLeftBound) Then
                    'we have gone too far left
                    udtBounds.Right = udtBounds.Right + (.intLeftBound - udtBounds.Left)
                    udtBounds.Left = .intLeftBound
                    blnChanged = True
                End If  'have we gone too far horizontally
                
                'have we gone too far vertically
                If (udtBounds.Bottom > .intBottomBound) Then
                
                    'we have gone too far down
                    udtBounds.Top = udtBounds.Top - (udtBounds.Bottom - .intBottomBound)
                    udtBounds.Bottom = .intBottomBound
                    blnChanged = True
                
                End If
                If (udtBounds.Top < .intTopBound) Then
                    'we have gone too far up
                    udtBounds.Bottom = udtBounds.Bottom + (.intTopBound - udtBounds.Top)
                    udtBounds.Top = .intTopBound
                    blnChanged = True
                End If  'have we gone too far vertically
            End With    'mudtForm(intPos)
        End If  'have we hooked this window
        
        If blnChanged Then
            'set the position and size of the form
            CopyMemory ByVal lParam, udtBounds, Len(udtBounds)
        End If
        
    Case Else
        'pass the message on to the next window
        WindowProc = CallWindowProc(mudtForm(0).lngPrevhWnd, hw, uMsg, wParam, lParam)
    End Select
End Function
