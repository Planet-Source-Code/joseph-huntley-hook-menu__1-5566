Attribute VB_Name = "modHookMenu"
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111
Public Const WM_CLOSE = &H10



''Variables to store our dynamic menu's item IDs
Public gDynSubMenu1 As Long
Public gDynSubMenu2 As Long
Public gDynSubMenu3 As Long
Public gDynSubMenu4 As Long
Public gDynSubMenu5 As Long

''Variable to hold the address of the old window procedure
Public gOldProc As Long
Public Function MenuProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

       Select Case wMsg&
           Case WM_CLOSE:
               ''User has closed the window, so we should stop
               ''subclassing immediately! We do this by handing
               ''back the original window procedure.
               Call SetWindowLong(hwnd&, GWL_WNDPROC, gOldProc&)
          
           Case WM_COMMAND:
               ''WM_COMMAND is sent to the window
               ''whenever someone clicks a menu.
               ''The menu's item ID is stored in wParam.
               
                  Select Case wParam&
                      Case gDynSubMenu1&
                          Call MsgBox("You clicked Dynamic Sub Menu 1!", vbExclamation)
                      Case gDynSubMenu2&
                          Call MsgBox("You clicked Dynamic Sub Menu 2!", vbExclamation)
                      Case gDynSubMenu3&
                          Call MsgBox("You clicked Dynamic Sub Menu 3!", vbExclamation)
                      Case gDynSubMenu4&
                          Call MsgBox("You clicked Dynamic Sub Menu 4!", vbExclamation)
                      Case gDynSubMenu5&
                          Call MsgBox("You clicked Dynamic Sub Menu 5!", vbExclamation)
                  End Select
       
       End Select

    ''Call original window procedure for default processing.
    MenuProc = CallWindowProc(gOldProc&, hwnd&, wMsg&, wParam&, lParam&)

End Function
