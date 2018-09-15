Attribute VB_Name = "McClosingForm"
'Module created by M.C., jan, 2001
'**********************************
'Cool App closing procedures
'**********************************
'better then similar ones ? Yes, border style of your form is not important

'call like this:
'MCCloseForm Me, number
'number can be anything from 1 to 16
'most interesting are 1 to 8

'or
'MCCloseForm Me, "Rnd"
'here number is randomly generated

'Have fun !

'Declares
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const WS_EX_TOPMOST = &H8&
Private Const WS_BORDER = &H800000
Private Const WS_SYSMENU = &H80000
Private Const WS_POPUP = &H80000000
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

Private Const SW_SHOWNORMAL = 1
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOCOPYBITS = &H100

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long


Public Sub MCCloseForm(FormToClose As Form, process As Variant)
  Dim Hwndnew As Long
  Dim R As RECT
  Dim Rwidth As Integer
  Dim Rheight As Integer
  
  'get dimensions of form closing into RECT struct.
  GetWindowRect FormToClose.hwnd, R
  'calculate it's width and height
  Rwidth = R.Right - R.Left
  Rheight = R.Bottom - R.Top
  
    ' Create the API window 'that copys form closing window image & dimensions into itself
    Hwndnew = CreateWindowEx(0, "static", "", WS_POPUPWINDOW Or WS_EX_TOPMOST, R.Top, R.Left, Rwidth, Rheight, 0, 0, App.hInstance, ByVal 0&)
    'show it
    ShowWindow Hwndnew, SW_SHOWNORMAL
    'time to kill our form that we are actualy closing
    Unload FormToClose


    'following 4 lines needed only in compiled project
    'the aim is to remove focus from our created form
    'if we wouldn't do this result would be ugly(only in compiled project) - you can try
    'Any smarter solutions to this - TELL ME: kozlicki@yahoo.com
    
    '1.move it a bit - so mouse click will 100% sure happen outside our form
    R.Left = R.Left + 1
    SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
    '2.move the mouse cursor to place outside our form & click there = focus is lost i.e.
    'set the focus to whatewer window is there - no harm done as probably no event is written
    'for clicking edge of window
    SetCursorPos 0, Screen.Height / Screen.TwipsPerPixelY / 2
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    
    
    'Closing procedure
    If process = "Rnd" Then
    Randomize
    process = Int(Rnd * 16)
    End If
    
    
    Select Case process
    Case 1 'normal - all sides shrinking
            Do
            If Rheight < 3 Then Exit Do
            R.Left = R.Left + 1
            R.Top = R.Top + 1
            Rwidth = Rwidth - 2
            Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
            
    Case 2 'TV - as 1 + something more
            Do
            If Rheight < 3 Then Exit Do
            R.Left = R.Left + 1
            R.Top = R.Top + 1
            Rwidth = Rwidth - 2
            Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
            
            Do
            If Rwidth < 0 Then Exit Do
            R.Left = R.Left + 1
            Rwidth = Rwidth - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
    Case 3 ' --->  <---
            Do
            If Rwidth < 3 Then Exit Do
            R.Left = R.Left + 1
            'R.Top = R.Top + 1
            Rwidth = Rwidth - 2
            'Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
            
            Do
            If Rheight < 0 Then Exit Do
            R.Top = R.Top + 1
            Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
    Case 4 ' opposite of 3
            Do
            If Rheight < 3 Then Exit Do
            'R.Left = R.Left + 1
            R.Top = R.Top + 1
            'Rwidth = Rwidth - 2
            Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
            
            Do
            If Rwidth < 0 Then Exit Do
            R.Left = R.Left + 1
            Rwidth = Rwidth - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
     Case 5 ' erase ---------->
            Do
            If Rwidth < 0 Then Exit Do
            R.Left = R.Left + 1
            'R.Top = R.Top + 1
            Rwidth = Rwidth - 1
            'Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
    Case 6 ' erase  <----------
            Do
            If Rwidth < 0 Then Exit Do
            'R.Left = R.Left - 1
            'R.Top = R.Top + 1
            Rwidth = Rwidth - 1
            'Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
    Case 7 ' erase  up
            Do
            If Rheight < 0 Then Exit Do
            'R.Left = R.Left - 1
            'R.Top = R.Top + 1
            'Rwidth = Rwidth - 1
            Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
    Case 8 ' erase  down
            Do
            If Rheight < 0 Then Exit Do
            'R.Left = R.Left - 1
            R.Top = R.Top + 1
            'Rwidth = Rwidth - 1
            Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents 'enable redrawing
            Loop
    Case 9 ' Form exit to the right
            Do
            If R.Left > Screen.Width / Screen.TwipsPerPixelX Then Exit Do
            R.Left = R.Left + 1
            'R.Top = R.Top + 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
    Case 10 ' Form exit to the left
            Do
            If Abs(R.Left) > Rwidth Then Exit Do
            R.Left = R.Left - 1
            'R.Top = R.Top + 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
     Case 11 ' Form exit to the top
            Do
            If Abs(R.Top) > Rheight Then Exit Do
            'R.Left = R.Left - 1
            R.Top = R.Top - 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
    Case 12 ' Form exit to the bottom
            Do
            If R.Top > Screen.Height / Screen.TwipsPerPixelY Then Exit Do
            'R.Left = R.Left - 1
            R.Top = R.Top + 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
    Case 13 ' Form exit to the top/right
            Do
            If Abs(R.Top) > Rheight Then Exit Do
            R.Left = R.Left + 1
            R.Top = R.Top - 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
    Case 14 ' Form exit to the top/left
            Do
            If Abs(R.Top) > Rheight Then Exit Do
            R.Left = R.Left - 1
            R.Top = R.Top - 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
    Case 15 ' Form exit to the bottom/left
            Do
            If R.Top > Screen.Height / Screen.TwipsPerPixelY Then Exit Do
            R.Left = R.Left - 1
            R.Top = R.Top + 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
    Case 16 ' Form exit to the bottom/right
            Do
            If R.Top > Screen.Height / Screen.TwipsPerPixelY Then Exit Do
            R.Left = R.Left + 1
            R.Top = R.Top + 1
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
    Case "Min"
            Do
            If Rheight < 25 Then Exit Do
            'R.Left = R.Left - 1
            'R.Top = R.Top + 1
            'Rwidth = Rwidth - 2
            Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
            
            Do
            If R.Top > (Screen.Height / Screen.TwipsPerPixelY) - 40 Then Exit Do
            'R.Left = R.Left - 1
            R.Top = R.Top + 2
            'Rwidth = Rwidth - 1
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
            
            Do
            If Rwidth < 26 Then Exit Do
            Rwidth = Rwidth - 4
            R.Left = R.Left + 2
            R.Top = R.Top - 2
            
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
            
            Do
            If R.Top > (Screen.Height / Screen.TwipsPerPixelY) - 40 Then Exit Do
            'Rwidth = Rwidth - 2
            R.Left = R.Left + 2
            R.Top = R.Top + 2
            
            'Rheight = Rheight - 1
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
            DoEvents 'enable redrawing
            Loop
            
            
            
    Case Else
    End Select
    
    
   'destroj API window
   DestroyWindow Hwndnew
End Sub



