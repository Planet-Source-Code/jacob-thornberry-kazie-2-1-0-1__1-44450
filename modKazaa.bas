Attribute VB_Name = "modKazaa"
Option Explicit

Const MAX_NAME = 256
Const GWL_STYLE = (-16)

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Declare Function CopyStringA Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function VirtualQueryEx& Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const WM_USER = &H400

Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETTEXT = (WM_USER + 11)
Private Const TTM_GETTOOLINFO = (WM_USER + 8)

Private Type SYSTEM_INFO ' 36 Bytes

   dwOemID As Long
   dwPageSize As Long
   lpMinimumApplicationAddress As Long
   lpMaximumApplicationAddress As Long
   dwActiveProcessorMask As Long
   dwNumberOrfProcessors As Long
   dwProcessorType As Long
   dwAllocationGranularity As Long
   wProcessorLevel As Integer
   wProcessorRevision As Integer
   
End Type

Private Type MEMORY_BASIC_INFORMATION ' 28 bytes

   BaseAddress As Long
   AllocationBase As Long
   AllocationProtect As Long
   RegionSize As Long
   state As Long
   Protect As Long
   lType As Long
   
End Type

Const PROCESS_VM_READ = (&H10)
Const PROCESS_VM_WRITE = (&H20)
Const PROCESS_VM_OPERATION = (&H8)
Const PROCESS_QUERY_INFORMATION = (&H400)
Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION

Const MEM_PRIVATE& = &H20000
Const MEM_COMMIT& = &H1000


''Tooltip Window Types
Private Type TOOLINFO

    cbSize As Long
    uFlags As Long
    hWnd As Long
    uID As Long
    RECT As RECT
    hinst As Long
    lpszText As Long
    lParam As Long
    
End Type

Declare Function SendMessageType Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
     lpTV_ITEM As TVITEM) As Long


Public Const TTM_GETTEXTA = (WM_USER + 11)
Public Const WM_RBUTTONDOWN = &H204

Private m_hwndLV As Long   ' ListView1.hWnd
Private Const CB_MAXLVITEMTEXT = 256   ' user-defined, exceptions will happen if not big enough...


Function Find_KaZaaQue() As Long

Dim lngStatic As Long
Dim lngSysTreeView As Long
Dim hwndTT As Long, Kazaa As Long
Dim mdiclient As Long, afxframeorviews As Long
Dim afxmdiframes As Long, X As Long
Dim numItems As Long, hndToolTips As Long
Dim tvCopy As TVITEM
Dim ItemText As String, GetItemString As String
Dim hItem As Long, lngTemp As Long

Kazaa& = FindWindow("kazaa", vbNullString)
mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
X& = FindWindowEx(afxmdiframes&, X&, "#32770", vbNullString)
lngStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
lngSysTreeView& = FindWindowEx(lngStatic&, 0&, "systreeview32", vbNullString)

If lngSysTreeView <> 0 Then

    Debug.Print "Found que!!"
    Find_KaZaaQue = lngSysTreeView
    numItems = TreeView_GetCount(lngSysTreeView)
    Debug.Print "numItems = " + Str(numItems)

    'lngTemp = TreeView_
    'lngTemp = TreeView_SelectItem(lngSysTreeView, 2)
    'Debug.Print "Success = " + str(lngTemp)
    
    hItem = TreeView_GetToolTips(lngSysTreeView)
    Debug.Print "Tool tips = " + Str(hItem)
    
Else
    Debug.Print "Que not found!!"
    
End If

End Function



Public Function RunKazaaByString(StringSearch As String) As Boolean

Dim Kazaa As Long, aMenu As Long, mCount As Long
Dim LookFor As Long, sMenu As Long, sCount As Long
Dim LookSub As Long, sID As Long, sString As String
Dim Win_Buff0 As Long, LookSubSub As Long
Dim ssSub As Long, ssCount As Long

'----- Assume FALSE ------------------
RunKazaaByString = False

'----  Get yahoo program handle ------
Win_Buff0 = FindWindow("kazaa", vbNullString)

If Win_Buff0 <> 0 Then
    Kazaa& = Win_Buff0
    Debug.Print "Founds KaZaa Handle"
    
Else
    Debug.Print "Did not find KaZaa Handle"
    RunKazaaByString = False
    Exit Function
    
End If
      
aMenu& = GetMenu(Kazaa&)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        Debug.Print ">> RYS >> Menu String = " + Trim(sString$) + " >> SID >> " + Str(sID&)
        'Debug.Print ">> RsID = " + Str(sID)
        
        If sID = -1 Then  'Located SUB SUB MENU
           ssSub& = GetSubMenu(sMenu&, LookSub&) 'Check for sub sub items
           ssCount& = GetMenuItemCount(ssSub&)
           Debug.Print ">> RYS >> ssCount = " + Str(ssCount)
           For LookSubSub = 0 To ssCount& - 1
               sID& = GetMenuItemID(ssSub&, LookSubSub&)
               sString$ = String$(100, " ")
               Call GetMenuString(ssSub&, sID&, sString$, 100&, 1&)
               Debug.Print ">> RYS >> SUB SUB String = " + Trim(sString$) + " >> SID >> " + Str(sID&)
               
               If InStr(LCase(sString$), LCase(StringSearch)) Then
                Debug.Print ">> RYS >> RunMenu SUB SUB Located for " + sString$
                'Call SendMessageLong(Kazaa&, WM_COMMAND, sID&, 0&)
                Call PostMessage(Kazaa&, WM_COMMAND, sID&, 0&)
                RunKazaaByString = True
                Exit Function
              End If
           Next
        End If
        
        If InStr(LCase(sString$), LCase(StringSearch)) Then
                Debug.Print ">> RYS >> RunMenu String Located for " + sString$
                'Call SendMessageLong(Kazaa&, WM_COMMAND, sID&, 0&)
                Call PostMessage(Kazaa&, WM_COMMAND, sID&, 0&)
                RunKazaaByString = True
                Exit Function
        End If
   Next LookSub&
Next LookFor&

End Function

Public Function FindKazaaButton()

Dim Kazaa As Long
Dim KazaaPrompt As Long
Dim KazaaButton As Long

Kazaa& = FindWindow("kazaa", vbNullString)
Debug.Print "Kazaa = " + Str(Kazaa)
Debug.Print "KaZaa caption = " + GetCaption(Kazaa)
'Pause (2)

KazaaPrompt& = FindAWindow("32770", "KaZaA Lite")
Debug.Print "Kazaa Prompt = " + Str(KazaaPrompt)

'KazaaPrompt& = FindChildByTitle(Kazaa, "KaZaA Lite")
'Debug.Print "Kazaa Prompt = " + Str(KazaaPrompt)
'Pause (2)
KazaaButton = FindChildByTitle(KazaaPrompt, "&Yes")
If KazaaButton = 0 Then
    KazaaButton = FindChildByTitle(KazaaPrompt, "Button")
End If

Debug.Print "Kazaa Button = " + Str(KazaaButton)
'Create seperate code for Kazaa and Kazaa LiTe
'Pause (1)

ButtonPress (KazaaButton)

'Call SendMessage(KazaaButton, WM_KEYDOWN, VK_SPACE, 0&)
'Call SendMessage(KazaaButton, WM_KEYUP, VK_SPACE, 0&)

End Function
Public Function Kazaa_Down()

Dim hndKazaaTree As Long

'hndKazaaTree = Find_KaZaaQue
hndKazaaTree = Find_Que
Call SendMessage(hndKazaaTree, WM_KEYDOWN, VK_DOWN, 0&)
Call SendMessage(hndKazaaTree, WM_KEYUP, VK_DOWN, 0&)

End Function
Public Function Kazaa_Up()

Dim hndKazaaTree As Long

'hndKazaaTree = Find_KaZaaQue
hndKazaaTree = Find_Que
Call SendMessage(hndKazaaTree, WM_KEYDOWN, VK_UP, 0&)
Call SendMessage(hndKazaaTree, WM_KEYUP, VK_UP, 0&)

End Function


Public Function Kazaa_CancelMenu()

Dim click As Long
Dim hndKazaaTree As Long
Dim hndSelection As Long
Dim hndTest As Long

Dim X As Long

hndKazaaTree = Find_KaZaaQue
Debug.Print "Que = " + Str(hndKazaaTree)
Debug.Print "Tips Que = " + Str(TreeView_GetToolTips(hndKazaaTree))

hndSelection = TreeView_GetFirstVisible(hndKazaaTree)
Debug.Print "Selection = " + Str(hndSelection)
Debug.Print "Tips selection = " + Str(TreeView_GetToolTips(hndSelection))

X = TreeView_GetLastVisible(hndKazaaTree)
Debug.Print "X = " + Str(X)
Debug.Print "Tips X = " + Str(TreeView_GetToolTips(X))
hndSelection = X
'Pause (1)

click = PostMessage(hndSelection&, WM_RBUTTONDOWN, 0, ByVal 0&)
'click = PostMessage(hndSelection&, WM_RBUTTONUP, 0, ByVal 0&)
    
Debug.Print ">> BP >> Click = " + Str(click)

End Function

