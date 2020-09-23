Attribute VB_Name = "modTreeviewMacros_472"
Option Explicit

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type LVFINDINFO
  flags As Long
  psz As String
  lParam As Long
  pt As POINTAPI
  vkDirection As Long
End Type

'
' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'   http://www.mvps.org/ccrp
'
' Code was written in and formatted for 8pt MS San Serif
'
' ================================================================
' Treeview control macros
'
' 53 macros total,
'   IE2 = 37 (v4.00.950)
'   IE3 =   2 (v4,70)
'   IE4 = 14 (v4.71, v4.72)

'

' Inserts a new item in a tree-view control.
' Returns the handle to the new item if successful or 0 otherwise.

Public Function TreeView_InsertItem(hWnd As Long, lpis As TVINSERTSTRUCT) As Long
  'TreeView_InsertItem = SendMessage(hWnd, TVM_INSERTITEM, 0, lpis)
End Function

' Removes an item from a tree-view control.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_DeleteItem(hWnd As Long, hItem As Long) As Boolean
  'TreeView_DeleteItem = SendMessage(hWnd, TVM_DELETEITEM, 0, ByVal hItem)
End Function

' Removes all items from a tree-view control.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_DeleteAllItems(hWnd As Long) As Boolean
  'TreeView_DeleteAllItems = SendMessage(hWnd, TVM_DELETEITEM, 0, ByVal TVI_ROOT)
End Function

' Expands or collapses the list of child items, if any, associated with the specified parent item.
' Returns TRUE if successful or FALSE otherwise.
' (docs say TVM_EXPAND does not send the TVN_ITEMEXPANDING and
' TVN_ITEMEXPANDED notification messages to the parent window...?)

Public Function TreeView_Expand(hWnd As Long, hItem As Long, flag As TVM_EXPAND_wParam) As Boolean
  'TreeView_Expand = SendMessage(hWnd, TVM_EXPAND, ByVal flag, ByVal hItem)
End Function

' Retrieves the bounding rectangle for a tree-view item and indicates whether the item is visible.
' If the item is visible and retrieves the bounding rectangle, the return value is TRUE.
' Otherwise, the TVM_GETITEMRECT message returns FALSE and does not retrieve
' the bounding rectangle.

Public Function TreeView_GetItemRect(hWnd As Long, hItem As Long, prc As RECT, fItemRect As CBoolean) As Boolean
  prc.Left = hItem
  'TreeView_GetItemRect = SendMessage(hWnd, TVM_GETITEMRECT, ByVal fItemRect, prc)
End Function

' Returns the count of total items in a tree-view control.

Public Function TreeView_GetCount(hWnd As Long) As Long
  'TreeView_GetCount = SendMessage(hWnd, TVM_GETCOUNT, 0, 0)
End Function

' Retrieves the amount, in pixels, that child items are indented relative to their parent items.
' Returns the amount of indentation.

Public Function TreeView_GetIndent(hWnd As Long) As Long
  'TreeView_GetIndent = SendMessage(hWnd, TVM_GETINDENT, 0, 0)
End Function

' Sets the indentation pixel width for a tree-view control and redraws the control to reflect the new width.
' No return value.

Public Sub TreeView_SetIndent(hWnd As Long, iIndent As Long)
  'Call SendMessage(hWnd, TVM_SETINDENT, ByVal iIndent, 0)
End Sub

' Retrieves the handle to the normal or state image list associated with a tree-view control.
' Returns the handle to the image list.

Public Function TreeView_GetImageList(hWnd As Long, iImage As Long) As Long
  'TreeView_GetImageList = SendMessage(hWnd, TVM_GETIMAGELIST, ByVal iImage, 0)
End Function

' Sets the normal or state image list for a tree-view control and redraws the control using the new images.
' Returns the handle to the previous image list, if any, or 0 otherwise.

Public Function TreeView_SetImageList(hWnd As Long, himl As Long, iImage As Long) As Long
  'TreeView_SetImageList = SendMessage(hWnd, TVM_SETIMAGELIST, ByVal iImage, ByVal himl)
End Function

' ======= Begin TreeView_GetNextItem ===========================================

' Retrieves the tree-view item that bears the specified relationship to a specified item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextItem(hWnd As Long, hItem As Long, flag As Long) As Long
  'TreeView_GetNextItem = SendMessage(hWnd, TVM_GETNEXTITEM, ByVal flag, ByVal hItem)
End Function

' Retrieves the first child item. The hitem parameter must be NULL.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetChild(hWnd As Long, hItem As Long) As Long
  'TreeView_GetChild = TreeView_GetNextItem(hWnd, hItem, TVGN_CHILD)
End Function

' Retrieves the next sibling item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextSibling(hWnd As Long, hItem As Long) As Long
  'TreeView_GetNextSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXT)
End Function

' Retrieves the previous sibling item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetPrevSibling(hWnd As Long, hItem As Long) As Long
  'TreeView_GetPrevSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_PREVIOUS)
End Function

' Retrieves the parent of the specified item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetParent(hWnd As Long, hItem As Long) As Long
  'TreeView_GetParent = TreeView_GetNextItem(hWnd, hItem, TVGN_PARENT)
End Function

' Retrieves the first visible item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetFirstVisible(hWnd As Long) As Long
  'TreeView_GetFirstVisible = TreeView_GetNextItem(hWnd, 0, TVGN_FIRSTVISIBLE)
End Function

' Retrieves the next visible item that follows the specified item. The specified item must be visible.
' Use the TVM_GETITEMRECT message to determine whether an item is visible.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextVisible(hWnd As Long, hItem As Long) As Long
  'TreeView_GetNextVisible = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXTVISIBLE)
End Function

' Retrieves the first visible item that precedes the specified item. The specified item must be visible.
' Use the TVM_GETITEMRECT message to determine whether an item is visible.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetPrevVisible(hWnd As Long, hItem As Long) As Long
  'TreeView_GetPrevVisible = TreeView_GetNextItem(hWnd, hItem, TVGN_PREVIOUSVISIBLE)
End Function

' Retrieves the currently selected item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetSelection(hWnd As Long) As Long
  'TreeView_GetSelection = TreeView_GetNextItem(hWnd, 0, TVGN_CARET)
End Function

' Retrieves the item that is the target of a drag-and-drop operation.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetDropHilight(hWnd As Long) As Long
  'TreeView_GetDropHilight = TreeView_GetNextItem(hWnd, 0, TVGN_DROPHILITE)
End Function

' Retrieves the topmost or very first item of the tree-view control.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetRoot(hWnd As Long) As Long
  'TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)
End Function
'
' ======= End TreeView_GetNextItem =============================================
'

' ======= Begin TreeView_Select ================================================

' Selects the specified tree-view item, scrolls the item into view, or redraws the item
' in the style used to indicate the target of a drag-and-drop operation.
' If hitem is NULL, the selection is removed from the currently selected item, if any.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_Select(hWnd As Long, hItem As Long, code As Long) As Boolean
  'TreeView_Select = SendMessage(hWnd, TVM_SELECTITEM, ByVal code, ByVal hItem)
End Function

' Sets the selection to the specified item.
' Returns TRUE if successful or FALSE otherwise.

' If the specified item is already selected, a TVN_SELCHANGING *will not* be generated !!

' If the specified item is 0 (indicating to remove selection from any currrently selected item)
' and an item is selected, a TVN_SELCHANGING *will* be generated and the itemNew
' member of NMTREEVIEW will be 0 !!!

Public Function TreeView_SelectItem(hWnd As Long, hItem As Long) As Boolean
  'TreeView_SelectItem = TreeView_Select(hWnd, hItem, TVGN_CARET)
End Function

' Redraws the given item in the style used to indicate the target of a drag and drop operation.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_SelectDropTarget(hWnd As Long, hItem As Long) As Boolean
  'TreeView_SelectDropTarget = TreeView_Select(hWnd, hItem, TVGN_DROPHILITE)
End Function

' Scrolls the tree view vertically so that the given item is the first visible item.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_SelectSetFirstVisible(hWnd As Long, hItem As Long) As Boolean
  'TreeView_SelectSetFirstVisible = TreeView_Select(hWnd, hItem, TVGN_FIRSTVISIBLE)
End Function
'
' ======= End TreeView_Select ==================================================
'

' Retrieves some or all of a tree-view item's attributes.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_GetItem(hWnd As Long, pitem As TVITEM) As Boolean
  'TreeView_GetItem = SendMessage(hWnd, TVM_GETITEM, 0, pitem)
End Function

' Sets some or all of a tree-view item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise

Public Function TreeView_SetItem(hWnd As Long, pitem As TVITEM) As Boolean
  'TreeView_SetItem = SendMessage(hWnd, TVM_SETITEM, 0, pitem)
End Function

' Begins in-place editing of the specified item's text, replacing the text of the item with a single-line
' edit control containing the text. This macro implicitly selects and focuses the specified item.
' Returns the handle to the edit control used to edit the item text if successful or NULL otherwise.

Public Function TreeView_EditLabel(hWnd As Long, hItem As Long) As Long
  'TreeView_EditLabel = SendMessage(hWnd, TVM_EDITLABEL, 0, (hItem))
End Function

' Retrieves the handle to the edit control being used to edit a tree-view item's text.
' Returns the handle to the edit control if successful or NULL otherwise.

Public Function TreeView_GetEditControl(hWnd As Long) As Long
  'TreeView_GetEditControl = SendMessage(hWnd, TVM_GETEDITCONTROL, 0, 0)
End Function

' Returns the number of items that are fully visible in the client window of the tree-view control.

Public Function TreeView_GetVisibleCount(hWnd As Long) As Long
  'TreeView_GetVisibleCount = SendMessage(hWnd, TVM_GETVISIBLECOUNT, 0, 0)
End Function

' Determines the location of the specified point relative to the client area of a tree-view control.
' Returns the handle to the tree-view item that occupies the specified point or NULL if no item
' occupies the point.

Public Function TreeView_HitTest(hWnd As Long, lpht As TVHITTESTINFO) As Long
  'TreeView_HitTest = SendMessage(hWnd, TVM_HITTEST, 0, lpht)
End Function

' Creates a dragging bitmap for the specified item in a tree-view control, creates an image list
' for the bitmap, and adds the bitmap to the image list. An application can display the image
' when dragging the item by using the image list functions.
' Returns the handle of the image list to which the dragging bitmap was added if successful or
' NULL otherwise.

Public Function TreeView_CreateDragImage(hWnd As Long, hItem As Long) As Long
  'TreeView_CreateDragImage = SendMessage(hWnd, TVM_CREATEDRAGIMAGE, 0, ByVal hItem)
End Function

' Sorts the child items of the specified parent item in a tree-view control.
' Returns TRUE if successful or FALSE otherwise.
' fRecurse is reserved for future use and must be zero.

Public Function TreeView_SortChildren(hWnd As Long, hItem As Long, fRecurse As Boolean) As Boolean
  'TreeView_SortChildren = SendMessage(hWnd, TVM_SORTCHILDREN, ByVal fRecurse, ByVal hItem)
End Function

' Ensures that a tree-view item is visible, expanding the parent item or scrolling the tree-view
' control, if necessary.
' Returns TRUE if the system scrolled the items in the tree-view control to ensure that the
' specified item is visible. Otherwise, the macro returns FALSE.

Public Function TreeView_EnsureVisible(hWnd As Long, hItem As Long) As Boolean
  'TreeView_EnsureVisible = SendMessage(hWnd, TVM_ENSUREVISIBLE, 0, ByVal hItem)
End Function

' Sorts tree-view items using an application-defined callback function that compares the items.
' Returns TRUE if successful or FALSE otherwise.
' fRecurse is reserved for future use and must be zero.

Public Function TreeView_SortChildrenCB(hWnd As Long, psort As TVSORTCB, fRecurse As Boolean) As Boolean
  'TreeView_SortChildrenCB = SendMessage(hWnd, TVM_SORTCHILDRENCB, ByVal fRecurse, psort)
End Function

' Ends the editing of a tree-view item's label.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_EndEditLabelNow(hWnd As Long, fCancel) As Boolean
  'TreeView_EndEditLabelNow = SendMessage(hWnd, TVM_ENDEDITLABELNOW, ByVal fCancel, 0)
End Function

' Retrieves the incremental search string for a tree-view control. The tree-view control uses the
' incremental search string to select an item based on characters typed by the user.
' Returns the number of characters in the incremental search string.
' If the tree-view control is not in incremental search mode, the return value is zero.

Public Function TreeView_GetISearchString(hWnd As Long, lpsz As String) As Boolean
  'TreeView_GetISearchString = SendMessage(hWnd, TVM_GETISEARCHSTRING, 0, lpsz)
End Function

' ================================================================
'#If (Win32_IE >= &H300) Then
'
' returns (HWND), old?
Public Function TreeView_SetToolTips(hWnd As Long, hwndTT As Long) As Long   ' IE3
  'TreeView_SetToolTips = SendMessage(hWnd, TVM_SETTOOLTIPS, ByVal hwndTT, 0)
End Function

' returns (hWnd)
Public Function TreeView_GetToolTips(hWnd As Long) As Long   ' IE3
  'TreeView_GetToolTips = SendMessage(hWnd, TVM_GETTOOLTIPS, 0, 0)
End Function
'
'#End If     ' WIN32_IE >= &H300
' ================================================================
'

Public Function TreeView_GetLastVisible(hWnd As Long) As Long   ' IE4
  TreeView_GetLastVisible = TreeView_GetNextItem(hWnd, 0, TVGN_LASTVISIBLE)
End Function

Public Function TreeView_Find(hwndListView As Long, strSearch As String) As Long

Dim lvfi As LVFINDINFO
Dim nItem As Long

Debug.Print "Searching for " + strSearch

With lvfi
  .flags = LVFI_STRING
  .psz = Trim(strSearch)
End With

'nItem = SendMessage(hwndListView, LVM_FINDITEM, -1&, lvfi)
TreeView_Find = nItem

'If nItem = -1& Then
'  MsgBox "No match."
'Else
'  MsgBox "Found a match at index: " & nItem
'End If

End Function

