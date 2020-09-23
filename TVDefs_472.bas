Attribute VB_Name = "modTreeviewDefs_472"
Option Explicit

' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'   http://www.mvps.org/ccrp
'
' Code was written in and formatted for 8pt MS San Serif

' ================================================================
' Dependent window definitions

Public Type POINTAPI   ' pt
    X As Long
    y As Long
End Type


Public Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum

' ================================================================
' Generic WM_NOTIFY notification messages.

' The lParam of all messages points to the NMHDR struct unless noted.
Public Enum CCNotifications
  NM_FIRST = -0&              ' (0U-  0U)       '  generic to all controls
  NM_LAST = -99&              ' (0U- 99U)
  NM_OUTOFMEMORY = (NM_FIRST - 1)
  NM_CLICK = (NM_FIRST - 2)
  NM_DBLCLK = (NM_FIRST - 3)
  NM_RETURN = (NM_FIRST - 4)
  NM_RCLICK = (NM_FIRST - 5)
  NM_RDBLCLK = (NM_FIRST - 6)
  NM_SETFOCUS = (NM_FIRST - 7)
  NM_KILLFOCUS = (NM_FIRST - 8)
#If (Win32_IE >= &H300) Then
  NM_CUSTOMDRAW = (NM_FIRST - 12)
  NM_HOVER = (NM_FIRST - 13)
#End If   ' 300
#If (Win32_IE >= &H400) Then
  NM_NCHITTEST = (NM_FIRST - 14)        ' lParam = NMMOUSE struct
  NM_KEYDOWN = (NM_FIRST - 15)        ' lParam =  NMKEY struct
  NM_RELEASEDCAPTURE = (NM_FIRST - 16)
  NM_SETCURSOR = (NM_FIRST - 17)     ' lParam =  NMMOUSE struct
  NM_CHAR = (NM_FIRST - 18)                ' lParam =  NMCHAR struct
#End If   '400
End Enum

' ================================================================
'  Generic WM_NOTIFY notification structures

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Public Type NMHDR
    hwndFrom As Long   ' Window handle of control sending message
    idFrom As Long        ' Identifier of control sending message
    code  As Long          ' Specifies the notification code
End Type

#If (Win32_IE >= &H400) Then

Public Type NMMOUSE
    hdr As NMHDR
    dwItemSpec As Long
    dwItemData As Long
    pt As POINTAPI
    dwHitInfo As Long ' any specifics about where on the item or control the mouse is
End Type

' Generic structure to request an object of a specific type.
Public Type NMOBJECTNOTIFY
    hdr As NMHDR
    iItem As Long
    piid As Long
    pObject As Long
    hResult As Long
    dwFlags As Long    ' control specific flags (hints as to where in iItem it hit)
End Type

' Generic structure for a key
Public Type NMKEY
    hdr As NMHDR
    nVKey As Long
    uFlags As Long
End Type

' Generic structure for a character
Public Type NMCHAR
    hdr As NMHDR
    ch As Long
    dwItemPrev As Long     ' Item previously selected
    dwItemNext As Long     ' Item to be selected
End Type
'
#End If           ' WIN32_IE >= &H400

' ==============================================================
' Shared common control messages

'#If (Win32_IE >= &H400) Then

Public Const CCM_FIRST = &H2000

Public Const CCM_SETBKCOLOR = (CCM_FIRST + 1)   ' lParam is bkColor

Public Type COLORSCHEME
  dwSize As Long
  clrBtnHighlight  As Long    ' highlight color, COLORREF
  clrBtnShadow As Long      ' shadow color, COLORREF
End Type

Public Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)     ' lParam is color scheme
Public Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)     ' fills in COLORSCHEME pointed to by lParam
Public Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
Public Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)

' For tooltips
Public Const INFOTIPSIZE = 1024

'#End If  ' WIN32_IE >= &H400

' ================================================================
' Treeview control definitions

' Classname
Public Const WC_TREEVIEW = "SysTreeView32"

' Treeview styles
Public Enum TVStyles
    TVS_HASBUTTONS = &H1
    TVS_HASLINES = &H2
    TVS_LINESATROOT = &H4
    TVS_EDITLABELS = &H8
    TVS_DISABLEDRAGDROP = &H10
    TVS_SHOWSELALWAYS = &H20
    TVS_RTLREADING = &H40
#If (Win32_IE >= &H300) Then
    TVS_NOTOOLTIPS = &H80
    TVS_CHECKBOXES = &H100
    TVS_TRACKSELECT = &H200
#If (Win32_IE >= &H400) Then
    TVS_SINGLEEXPAND = &H400
    TVS_INFOTIP = &H800
    TVS_FULLROWSELECT = &H1000
    TVS_NOSCROLL = &H2000
    TVS_NONEVENHEIGHT = &H4000
#End If   ' 400
#End If   ' 300

    TVS_SHAREDIMAGELISTS = &H0
    TVS_PRIVATEIMAGELISTS = &H400
End Enum

'typedef struct _TREEITEM FAR* HTREEITEM;

' ================================================================
' TVITEM struct

Public Type TVITEM   ' was TV_ITEM
    mask As TVITEM_mask
    hItem As Long
    state As TVITEM_state
    stateMask As Long
    pszText As String   ' pointer Use to be long a pointer
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

#If (Win32_IE >= &H400) Then
' only used for Get and Set messages (not for notifications)
Public Type TVITEMEX
    mask As TVITEM_mask
    hItem As Long
    state As TVITEM_state
    stateMask As Long
    pszText As Long   ' pointer
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
End Type
#End If

Public Enum TVITEM_mask
    TVIF_TEXT = &H1
    TVIF_IMAGE = &H2
    TVIF_PARAM = &H4
    TVIF_STATE = &H8
    TVIF_HANDLE = &H10
    TVIF_SELECTEDIMAGE = &H20
    TVIF_CHILDREN = &H40
#If (Win32_IE >= &H400) Then
    TVIF_INTEGRAL = &H80
#End If
    TVIF_DI_SETITEM = &H1000   ' Notification
End Enum

Public Const LVFI_STRING = &H2
Public Const LVM_FIRST = &H1000&
Public Const LVM_FINDITEM = (LVM_FIRST + 13)

Public Enum TVITEM_state
    TVIS_SELECTED = &H2
    TVIS_CUT = &H4
    TVIS_DROPHILITED = &H8
    TVIS_BOLD = &H10
    TVIS_EXPANDED = &H20
    TVIS_EXPANDEDONCE = &H40
#If (Win32_IE >= &H300) Then
    TVIS_EXPANDPARTIAL = &H80
#End If
    
    TVIS_OVERLAYMASK = &HF00
    TVIS_STATEIMAGEMASK = &HF000
    TVIS_USERMASK = &HF000
End Enum

' TVITEM(EX).pszText
Public Const LPSTR_TEXTCALLBACK = (-1)

' TVITEM.iImage, TVITEM.iSelectedImage
Public Const I_IMAGECALLBACK = (-1)
  
' TVITEM.cChildren
Public Const I_CHILDRENCALLBACK = (-1)

' ================================================================
' Treeview messages

Public Enum TVHandles
    TVI_ROOT = &HFFFF0000
    TVI_FIRST = &HFFFF0001
    TVI_LAST = &HFFFF0002
    TVI_SORT = &HFFFF0003
End Enum

Public Type TVINSERTSTRUCT   ' was TV_INSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
#If (Win32_IE >= &H400) Then
'    Union
'    {
         ' use larger of two structs
         itemex As TVITEMEX
'        TVITEM  item;
'    } DUMMYUNIONNAME;
#Else
    Item As TVITEM
#End If
End Type

Public Enum TVMessages
    TV_FIRST = &H1100
    
    #If UNICODE Then
      TVM_INSERTITEM = (TV_FIRST + 50)
    #Else
      TVM_INSERTITEM = (TV_FIRST + 0)
    #End If
    
    TVM_DELETEITEM = (TV_FIRST + 1)
    TVM_EXPAND = (TV_FIRST + 2)
    TVM_GETITEMRECT = (TV_FIRST + 4)
    TVM_GETCOUNT = (TV_FIRST + 5)
    TVM_GETINDENT = (TV_FIRST + 6)
    TVM_SETINDENT = (TV_FIRST + 7)
    TVM_GETIMAGELIST = (TV_FIRST + 8)
    TVM_SETIMAGELIST = (TV_FIRST + 9)
    TVM_GETNEXTITEM = (TV_FIRST + 10)
    TVM_SELECTITEM = (TV_FIRST + 11)
    
    #If UNICODE Then
      TVM_GETITEM = (TV_FIRST + 62)
      TVM_SETITEM = (TV_FIRST + 63)
      TVM_EDITLABEL = (TV_FIRST + 65)
    #Else
      TVM_GETITEM = (TV_FIRST + 12)
      TVM_SETITEM = (TV_FIRST + 13)
      TVM_EDITLABEL = (TV_FIRST + 14)
    #End If
    
    TVM_GETEDITCONTROL = (TV_FIRST + 15)
    TVM_GETVISIBLECOUNT = (TV_FIRST + 16)
    TVM_HITTEST = (TV_FIRST + 17)
    TVM_CREATEDRAGIMAGE = (TV_FIRST + 18)
    TVM_SORTCHILDREN = (TV_FIRST + 19)
    TVM_ENSUREVISIBLE = (TV_FIRST + 20)
    TVM_SORTCHILDRENCB = (TV_FIRST + 21)
    TVM_ENDEDITLABELNOW = (TV_FIRST + 22)
    
    #If UNICODE Then
      TVM_GETISEARCHSTRING = (TV_FIRST + 64)
    #Else
      TVM_GETISEARCHSTRING = (TV_FIRST + 23)
    #End If
    
'#If (Win32_IE >= &H300) Then
    TVM_SETTOOLTIPS = (TV_FIRST + 24)
    TVM_GETTOOLTIPS = (TV_FIRST + 25)
'#End If    ' 0x0300

#If (Win32_IE >= &H400) Then
    TVM_SETINSERTMARK = (TV_FIRST + 26)
    TVM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
    TVM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
    TVM_SETITEMHEIGHT = (TV_FIRST + 27)
    TVM_GETITEMHEIGHT = (TV_FIRST + 28)
    TVM_SETBKCOLOR = (TV_FIRST + 29)
    TVM_SETTEXTCOLOR = (TV_FIRST + 30)
    TVM_GETBKCOLOR = (TV_FIRST + 31)
    TVM_GETTEXTCOLOR = (TV_FIRST + 32)
    TVM_SETSCROLLTIME = (TV_FIRST + 33)
    TVM_GETSCROLLTIME = (TV_FIRST + 34)
    TVM_SETINSERTMARKCOLOR = (TV_FIRST + 37)
    TVM_GETINSERTMARKCOLOR = (TV_FIRST + 38)
#End If   ' 0x0400

End Enum   ' TVMessages
    
Public Enum TVM_EXPAND_wParam
    TVE_COLLAPSE = &H1
    TVE_EXPAND = &H2
    TVE_TOGGLE = &H3
#If (Win32_IE >= &H300) Then
    TVE_EXPANDPARTIAL = &H4000
#End If
    TVE_COLLAPSERESET = &H8000
End Enum
    
Public Enum TVM_GET_SETIMAGELIST_wParam
    TVSIL_NORMAL = 0
    TVSIL_STATE = 2
End Enum
    
Public Enum TVM_GETNEXTITEM_wParam
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
    TVGN_LASTVISIBLE = &HA

End Enum

Public Type TVHITTESTINFO   ' was TV_HITTESTINFO
    pt As POINTAPI
    flags As TVHITTESTINFO_flags
    hItem As Long
End Type
    
Public Enum TVHITTESTINFO_flags
    TVHT_NOWHERE = &H1
    TVHT_ONITEMICON = &H2
    TVHT_ONITEMLABEL = &H4
    TVHT_ONITEMINDENT = &H8
    TVHT_ONITEMBUTTON = &H10
    TVHT_ONITEMRIGHT = &H20
    TVHT_ONITEMSTATEICON = &H40
    TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
    
    TVHT_ABOVE = &H100
    TVHT_BELOW = &H200
    TVHT_TORIGHT = &H400
    TVHT_TOLEFT = &H800
End Enum

'typedef int (CALLBACK *PFNTVCOMPARE)(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort);

Public Type TVSORTCB   ' was TV_SORTCB
    hParent As Long
    lpfnCompare As Long
    lParam As Long
End Type

' ================================================================
' Treeview notifications

Public Enum TVNotifications
    TVN_FIRST = -400&   ' &HFFFFFE70   ' (0U-400U)
    TVN_LAST = -499&    ' &HFFFFFE0D    ' (0U-499U)
    
    #If UNICODE Then
      TVN_SELCHANGING = (TVN_FIRST - 50)
      TVN_SELCHANGED = (TVN_FIRST - 51)
      TVN_GETDISPINFO = (TVN_FIRST - 52)
      TVN_SETDISPINFO = (TVN_FIRST - 53)
      TVN_ITEMEXPANDING = (TVN_FIRST - 54)
      TVN_ITEMEXPANDED = (TVN_FIRST - 55)
      TVN_BEGINDRAG = (TVN_FIRST - 56)
      TVN_BEGINRDRAG = (TVN_FIRST - 57)
      TVN_DELETEITEM = (TVN_FIRST - 58)
      TVN_BEGINLABELEDIT = (TVN_FIRST - 59)
      TVN_ENDLABELEDIT = (TVN_FIRST - 60)
#If (Win32_IE >= &H400) Then
      TVN_GETINFOTIPW = (TVN_FIRST - 14)
#End If   ' 0x400
    #Else                                                      ' lParam points to:
      TVN_SELCHANGING = (TVN_FIRST - 1)          ' NMTREEVIEW
      TVN_SELCHANGED = (TVN_FIRST - 2)           ' NMTREEVIEW
      TVN_GETDISPINFO = (TVN_FIRST - 3)            ' NMTVDISPINFO
      TVN_SETDISPINFO = (TVN_FIRST - 4)            ' NMTVDISPINFO
      TVN_ITEMEXPANDING = (TVN_FIRST - 5)       ' NMTREEVIEW
      TVN_ITEMEXPANDED = (TVN_FIRST - 6)        ' NMTREEVIEW
      TVN_BEGINDRAG = (TVN_FIRST - 7)              ' NMTREEVIEW
      TVN_BEGINRDRAG = (TVN_FIRST - 8)            ' NMTREEVIEW
      TVN_DELETEITEM = (TVN_FIRST - 9)             ' NMTREEVIEW
      TVN_BEGINLABELEDIT = (TVN_FIRST - 10)    ' NMTVDISPINFO
      TVN_ENDLABELEDIT = (TVN_FIRST - 11)       ' NMTVDISPINFO
#If (Win32_IE >= &H400) Then
      TVN_GETINFOTIP = (TVN_FIRST - 13)
#End If   ' 0x400
    #End If   ' UNICODE
    TVN_KEYDOWN = (TVN_FIRST - 12)                ' NMTVKEYDOWN
#If (Win32_IE >= &H400) Then
    TVN_SINGLEEXPAND = (TVN_FIRST - 15)
#End If   ' 0x400
End Enum   ' Notifications

' 1st member of all notification structs is a NMHDR

Public Type NMTREEVIEW   ' was NM_TREEVIEW
    hdr As NMHDR
    ' Specifies a notification-specific action flag.
    ' Is NMTREEVIEW_action for TVN_SELCHANGING, TVN_SELCHANGED, TVN_SETDISPINFO
    ' Is TVM_EXPAND_wParam for TVN_ITEMEXPANDING, TVN_ITEMEXPANDED
    action As Long
    itemOld As TVITEM
    itemNew As TVITEM
    ptDrag As POINTAPI
End Type

' for TVN_SELCHANGING, TVN_SELCHANGED, TVN_SETDISPINFO
Public Enum NMTREEVIEW_action
    TVC_UNKNOWN = &H0
    TVC_BYMOUSE = &H1
    TVC_BYKEYBOARD = &H2
End Enum

Public Type NMTVDISPINFO   ' was TV_DISPINFO
    hdr As NMHDR
    Item As TVITEM
End Type

Public Type NMTVKEYDOWN   ' was TV_KEYDOWN
    hdr As NMHDR
    wVKey As Integer
    flags As Long   ' Always zero.
End Type

' for tooltips
Public Type NMTVGETINFOTIP
    hdr As NMHDR
    pszText As Long
    cchTextMax As Long
    hItem As Long
    lParam As Long
End Type

' treeview's customdraw return meaning don't draw images.
' valid on CDRF_NOTIFYITEMPREPAINT
Public Const TVCDRF_NOIMAGES = &H10000
'



' Prepares the index of a state image so that a tree view control or list
' view control can use the index to retrieve the state image for an item.
' Returns the one-based index of the state image shifted left twelve bits.
' A common control utility macro.

Public Function INDEXTOSTATEIMAGEMASK(iState As Long) As Long
' #define INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
  INDEXTOSTATEIMAGEMASK = iState * (2 ^ 12)
End Function
