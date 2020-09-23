VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Hook Menu"
   ClientHeight    =   3675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPremade 
      Caption         =   "Premade Menu"
      Begin VB.Menu mnuPremadeSub1 
         Caption         =   "Premade Submenu 1"
      End
      Begin VB.Menu mnuPremadeSub2 
         Caption         =   "Premade Submenu 2"
      End
      Begin VB.Menu mnuPremadeSub3 
         Caption         =   "Premade Submenu 3"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   This was requested by a couple of people plus I saw
'some questions on planet source code's discussion board
'about how to do this. Well, here it is. Enjoy =)
'
'      Joseph Huntley
'--------------------------------------------------------

Private Sub Form_Load()

   Dim lngMenu As Long, lngNewMenu As Long, lngNewSubMenu As Long

   ''Assign new menu's item IDs. This can be any
   ''number as long as the application doesn't have
   ''a menu with the same item ID.
   gDynSubMenu1& = 70
   gDynSubMenu2& = 71
   gDynSubMenu3& = 72
   gDynSubMenu4& = 73
   gDynSubMenu5& = 74

   ''Get the form's menu handle
   lngMenu& = GetMenu(Me.hwnd)

   ''Create a new popup menu to add our menus to.
   lngNewMenu& = CreatePopupMenu
   
   ''Now insert it into the place where a second menu
   ''is supposed to be on our form. (this is why the
   ''second parameter is 1 and not 0.
   
   ''NOTE: The MF_POPUP flag is used ONLY when inserting
   ''a new popup menu or new sub menu using CreatePopupMenu.
   
   ''When using MF_POPUP, the argument for the new item ID
   ''should contain the handle to the new popup menu, as shown below.
   
   ''You can use MF_SEPARATOR (without MF_STRING) if you want to add
   ''a separator line. When you do this, use vbNullString as the lpNewItem
   ''parameter.
   Call InsertMenu(lngMenu&, 1&, MF_POPUP Or MF_STRING Or MF_BYPOSITION, lngNewMenu&, "Dynamic Menu")
   
   ''Now add the sub menus
   Call InsertMenu(lngNewMenu&, 0&, MF_STRING Or MF_BYPOSITION, gDynSubMenu1&, "Dynamic Sub Menu 1")
   Call InsertMenu(lngNewMenu&, 1&, MF_STRING Or MF_BYPOSITION, gDynSubMenu2&, "Dynamic Sub Menu 2")
   
   ''The same way you create a new menu on the menu bar
   ''is the same way you create a new sub-submenu.
   lngNewSubMenu& = CreatePopupMenu
   
   Call InsertMenu(lngNewMenu&, 2&, MF_STRING Or MF_BYPOSITION Or MF_POPUP, lngNewSubMenu&, "Dynamic Sub-Submenu")
   
   ''Add two menus to our sub-submenu
   Call InsertMenu(lngNewSubMenu&, 0&, MF_STRING Or MF_BYPOSITION, gDynSubMenu3&, "Dynamic Sub Menu 3")
   Call InsertMenu(lngNewSubMenu&, 1&, MF_STRING Or MF_BYPOSITION, gDynSubMenu4&, "Dynamic Sub Menu 4")
   
   ''Now add one more menu to our original menu.
   Call InsertMenu(lngNewMenu&, 3&, MF_STRING Or MF_BYPOSITION, gDynSubMenu5&, "Dynamic Sub Menu 5")

   ''Now we want to know if it was clicked, right?
   ''So we subclass it by replacing the old window
   ''procedure with our own.
   
   ''Get the original window procedure, so we can call
   ''it and we can give it back when our program is done.
   gOldProc& = GetWindowLong(Me.hwnd, GWL_WNDPROC)
   
   ''Now replace the old window procedure
   Call SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MenuProc)

   ''Now whenever a window message is sent to the form
   ''it sends it to MenuProc

End Sub

