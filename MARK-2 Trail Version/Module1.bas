Attribute VB_Name = "Module1"
Option Explicit

Private Const SC_CLOSE As Long = &HF060&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

'*******************************************************************************
' Enables / Disables the close button on the titlebar and in the system menu
' of the form window passed.
'-------------------------------------------------------------------------------
' Return Values:
'
'    0  Close button state changed succesfully / nothing to do.
'   -1  Invalid Window Handle (hWnd argument) Passed to the function
'   -2  Failed to switch command ID of Close menu item in system menu
'   -3  Failed to switch enabled state of Close menu item in system menu
'
'-------------------------------------------------------------------------------
' Parameters:
'
'   hWnd    The window handle of the form whose close button is to be enabled/
'           disabled / greyed out.
'
'   Enable  True if the close button is to be enabled, or False if it is to
'           be disabled / greyed out.
'
'-------------------------------------------------------------------------------
' Example:
'
' Add a form window to your project, and place a button on the form. Add the
' following in the form's code window:
'
'    Option Explicit
'
'    Private m_blnCloseEnabled As Boolean
'
'    Private Sub Form_Load()
'        m_blnCloseEnabled = True
'        Command1.Caption = "Disable"
'    End Sub
'
'    Private Sub Command1_Click()
'        m_blnCloseEnabled = Not m_blnCloseEnabled
'        EnableCloseButton Me.hwnd, m_blnCloseEnabled
'
'        If m_blnCloseEnabled Then
'            Command1.Caption = "Disable"
'        Else
'            Command1.Caption = "Enable"
'        End If
'    End Sub
'
'-------------------------------------------------------------------------------

Public Function EnableCloseButton(ByVal hWnd As Long, Enable As Boolean) _
                                                                As Integer
    Const xSC_CLOSE As Long = -10

    ' Check that the window handle passed is valid
    
    EnableCloseButton = -1
    If IsWindow(hWnd) = 0 Then Exit Function
    
    ' Retrieve a handle to the window's system menu
    
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, 0)
    
    ' Retrieve the menu item information for the close menu item/button
    
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    
    If Enable Then
        MII.wID = xSC_CLOSE
    Else
        MII.wID = SC_CLOSE
    End If
    
    EnableCloseButton = -0
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Switch the ID of the menu item so that VB can not undo the action itself
    
    Dim lngMenuID As Long
    lngMenuID = MII.wID
    
    If Enable Then
        MII.wID = SC_CLOSE
    Else
        MII.wID = xSC_CLOSE
    End If
    
    MII.fMask = MIIM_ID
    EnableCloseButton = -2
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function
    
    ' Set the enabled / disabled state of the menu item
    
    If Enable Then
        MII.fState = (MII.fState Or MFS_GRAYED)
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If
    
    MII.fMask = MIIM_STATE
    EnableCloseButton = -3
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Activate the non-client area of the window to update the titlebar, and
    ' draw the close button in its new state.
    
    SendMessage hWnd, WM_NCACTIVATE, True, 0
    
    EnableCloseButton = 0
    
End Function

'*******************************************************************************
'
'-------------------------------------------------------------------------------


