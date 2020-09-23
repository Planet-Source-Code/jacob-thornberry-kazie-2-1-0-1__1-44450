Attribute VB_Name = "modLeech"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    
    '--- Test stuff ---
    Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
    
    Declare Function ReadProcessMemoryStr Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, lpNumberOfBytesRead As Long) As Long
    Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal _
    lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    
    Public Const PROCESS_READ = &H10
    Public Const RIGHTS_REQUIRED = &HF0000
    Const PROCESS_VM_OPERATION = &H8
    Const PROCESS_VM_READ = &H10
    Const PROCESS_VM_WRITE = &H20
    Const MEM_COMMIT = &H1000
    Const MEM_RESERVE = &H2000
    Const MEM_RELEASE = &H8000
    Const PAGE_READWRITE = &H4&
    Const LVIF_TEXT = &H1
    Private Const LVM_FIRST As Long = &H1000
    Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
    
    Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    
    Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

'---- End Test Stuff --------

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_CANCELMODE = &H1F
Public Const WM_ACTIVATE = &H6
Public Const WA_CLICKACTIVE = 2
Public Const WM_SETFOCUS = &H7

'---- Enable / disable functions ----
Global Const MF_ENABLED = &H0
Global Const MF_GRAYED = &H1
Global Const MF_DISABLED = &H2
Global Const MF_BYCOMMAND = &H0
Global Const MF_BYPOSITION = &H400

Public Const VK_C = &H43
Public Const VK_S = &H53
Public Const VK_F = &H46
Public Const VK_U = &H55
Global Const VK_UP = &H26
Global Const VK_RETURN = &HD
Const VK_MENU = &H12
Const VK_HOME = &H24

'--- Tool bar stuff -----
Public Const TB_BUTTONCOUNT = (WM_USER + 24)
Public Const TB_GETBUTTON = (WM_USER + 23)
Public Const TB_PRESSBUTTON = (WM_USER + 3)
Public Const TB_ENABLEBUTTON = WM_USER + 1
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_DEFAULT = SWP_NOMOVE Or SWP_NOSIZE

Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

Public Declare Function SetWindowPos _
Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal y As Long, _
ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long

Private Type TBBUTTON
iBitmap As Long
idCommand As Long
fsState As Byte
fsStyle As Byte
dwData As Long
iString As Long
End Type

Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, _
ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Public timStart As Date
Public timRestart As Date

Public timTemp As Long
Public timScanner As Long
Public timer0 As Long
Public Function Find_KazaaProg() As Long
    
    Dim Win_Buff0 As Long
    
    Win_Buff0 = FindWindow("kazaa", vbNullString)
    
    If Win_Buff0 <> 0 Then
        Find_Kazaa = Win_Buff0
        Debug.Print "Founds KaZaa Handle"
        
    Else
        Debug.Print "Did not find KaZaa Handle"
        Find_Kazaa = 0
        
    End If
    
End Function

Public Function Find_Que() As Long
    
    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long
    Dim hSystreeview As Long
    
    '!! NOTE THESE ARE FOR THE DOWNLOAD QUE !!
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    hStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    hSystreeview& = FindWindowEx(hStatic&, 0&, "systreeview32", vbNullString)
    
    '!! NOTE THESE ARE FOR THE UPLOAD QUE !!
    'Kazaa& = FindWindow("kazaa", vbNullString)
    'mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    'afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    'afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    'X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    'X& = FindWindowEx(afxmdiframes&, X&, "#32770", vbNullString)
    'hStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    'hSystreeview& = FindWindowEx(hStatic&, 0&, "systreeview32", vbNullString)
    
    If hSystreeview <> 0 Then
        
        Find_Que = hSystreeview
        
    Else
        
    End If
    
End Function
Public Function Find_Files() As Long
    
    Dim Kazaa&
    Dim mdiclient&
    Dim afxframeorviews&
    Dim afxmdiframes&
    Dim X&
    Dim hStatic&
    Dim hSystreeview&
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    X& = FindWindowEx(afxmdiframes&, X&, "#32770", vbNullString)
    hStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    hSystreeview& = FindWindowEx(hStatic&, 0&, "systreeview32", vbNullString)
    
    Find_Files = hSystreeview
    
End Function

Public Function Stop_Search()
    
    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long
    Dim Button As Long
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    Button& = FindWindowEx(X&, 0&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    
    If Button <> 0 Then
        Debug.Print "Caption = " + GetCaption(Button)
        'frmMain.lblStatus = "User verified! Stopping search"
        ButtonPress (Button)
        
    Else 'No button
        'frmMain.lblStatus = "Error: Did not find stop search button"
        
    End If
    
End Function

Public Function Kazaa_ClearQue()
    
    Dim Win_Buff0 As Long
    
    Dim rebarwindow As Long, tb1 As Long, tb2 As Long
    Dim Kazaa As Long, i As Long, click As Long
    Dim toolCount As Long, hTool As Long, X As Long
    Dim tbb As TBBUTTON, hNext As Long, nCount As Long
    Dim tb3 As Long, afxframeorviews As Long
    Dim afxmdiframes As Long, mdiclient As Long
    Dim hCheck As Long, hWIndow As Long, hText As Long
    Dim hButton As Long, retVal As Long
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    rebarwindow& = FindWindowEx(Kazaa&, 0&, "rebarwindow32", vbNullString)
    tb1& = FindWindowEx(rebarwindow&, 0&, "toolbarwindow32", vbNullString)
    tb2& = FindWindowEx(rebarwindow&, tb1&, "toolbarwindow32", vbNullString)
    
    If tb2& <> 0 Then
        
        click = PostMessage(tb2&, WM_COMMAND, 32962, tb1&) 'Run find more from same user
        
    Else
        
        'frmMain.lblStatus = "Error finding status bar... Can not clear que"
        'Pause (1)
        
    End If
    
End Function

Public Function FindKazaaButton() As Boolean
    
    Dim Kazaa As Long, X As Integer
    Dim KazaaPrompt As Long
    Dim KazaaButton As Long
    
    Kazaa& = Find_Kazaa
    
    Kazaa = FindWindow("#32770", "KaZaA Lite")
    If Kazaa = 0 Then
        Kazaa = FindWindow("#32770", "KaZaA Media Desktop")
    ElseIf Kazaa = 0 Then
        Kazaa = FindWindow("#32770", "KaZaA")
    ElseIf Kazaa = 0 Then
        Kazaa = FindWindow("#32770", "Kazaa Lite")
    ElseIf Kazaa = 0 Then
        Kazaa = FindWindow("32770", "")
        If InStr(1, LCase(GetCaption(Kazaa)), "kazaa") = 0 Then
            Kazaa = 0
        End If
        
    End If
    
    Debug.Print "Kazaa = " + Str(Kazaa)
    Debug.Print "KaZaa caption = " + GetCaption(Kazaa)
    
    'KazaaPrompt& = FindAWindow("32770", "KaZaA Lite")
    'Debug.Print "Kazaa Prompt = " + Str(KazaaPrompt)
    
    'KazaaPrompt& = FindChildByTitle(Kazaa, "KaZaA Lite")
    'Debug.Print "Kazaa Prompt = " + Str(KazaaPrompt)
    'Pause (2)
    KazaaButton = FindChildByTitle(Kazaa, "&Yes")
    If KazaaButton = 0 Then
        KazaaButton = FindWindowEx(Kazaa, 0, "Button", vbNullString)
    End If
    
    If KazaaButton <> 0 Then
        
        Debug.Print "Kazaa Button = " + Str(KazaaButton)
        'Create seperate code for Kazaa and Kazaa LiTe
        
        ButtonPress (KazaaButton)
        
        For X = 1 To 2
            Call SendMessage(KazaaButton, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(KazaaButton, WM_KEYUP, VK_SPACE, 0&)
            Next
            
            FindKazaaButton = True
            
        Else
            FindKazaaButton = False
            
        End If
        
    End Function

Public Function FindAWindow(Class As String, Caption As String)
    
    'This searches for Top-Level Windows
    'This can be more usefull then FindWindow
    'Because with this you don't need to know the exact
    'Caption of the window
    
    '--- 9/02 Switch to search on Visible windows! ----
    
    Dim Wind As Long, TextLen As Long, Buff As String
    Dim cntLoop As Long
    
    '-- Search only in Kazaa --
    'Wind& = GetWindow(Find_Kazaa, GW_CHILD)
    Wind& = GetWindow(Find_Kazaa, GW_HWNDFIRST)
    
    cntLoop = 0
    
    Do While Wind <> 0
        
        DoEvents
        
        If Wind& = 0 Then Exit Do
        If IsWindowVisible(Wind) Then
            TextLen& = SendMessage(Wind&, WM_GETTEXTLENGTH, 0, 0) + 1
            
            If TextLen& > 0 And TextLen < 32000 Then
                Buff$ = String(TextLen&, 0)
                Call SendMessageByString(Wind&, WM_GETTEXT, TextLen&, Buff$)
                If Len(Trim(Buff$)) > 1 Then
                    'Debug.Print ">> FAW >>" + Buff$
                End If
                
                'Debug.Print ">> Visible >>" + Buff$
                If InStr(LCase(Buff$), LCase(Caption$)) <> 0 Then 'And GetClass(Wind&) = Class$ Then
                FindAWindow = Wind&
                Exit Function
                
            End If
            
        End If 'Window has text
    End If 'Only search and process visible windows!
    
    Wind& = GetWindow(Wind&, GW_HWNDNEXT)
    
    DoEvents
    
    Loop
    
    FindAWindow = 0
    
End Function

Public Function Search_User()
    
    Dim rebarwindow As Long, tb1 As Long, tb2 As Long
    Dim Kazaa As Long, i As Long, click As Long
    Dim toolCount As Long, hTool As Long, X As Long
    Dim tbb As TBBUTTON, hNext As Long, nCount As Long
    Dim tb3 As Long, Win_Buff0 As Long, afxframeorviews As Long
    Dim afxmdiframes As Long, mdiclient As Long
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    rebarwindow& = FindWindowEx(Kazaa&, 0&, "rebarwindow32", vbNullString)
    tb1& = FindWindowEx(rebarwindow&, 0&, "toolbarwindow32", vbNullString)
    tb2& = FindWindowEx(rebarwindow&, tb1&, "toolbarwindow32", vbNullString)
    
    If tb2& <> 0 Then
        
        click = PostMessage(tb2&, WM_COMMAND, 32990, tb1&) '--- Run find more from same user 1.7.1
        click = PostMessage(tb2&, WM_COMMAND, 32909, tb1&) '--- Check user 1.7.2
        
    Else
        
        'frmMain.lblStatus = "Error finding status bar... Can not scan user"
        'Pause (1)
        
    End If
    
    '---- Wait for the switch to Kazaa - Search ------
    
    timTemp = 0
    Win_Buff0 = FindWindow("KaZaA", vbNullString)
    GetCaption (Win_Buff0)
    
    Do While LCase(GetCaption(Win_Buff0)) <> LCase("KaZaA - [Search]")
        Win_Buff0 = FindWindow("KaZaA", vbNullString)
        GetCaption (Win_Buff0)
        DoEvents
        If timTemp > 15 Then
            'frmMain.lblStatus = "Error searching this user."
            Exit Do
        End If
        
        Loop
        
    End Function
Public Function Get_PopUP()
    
    Dim X As Long, pMenu As Long
    Dim pCount As Long
    
    X& = FindWindow("#32768", vbNullString)
    
    If X <> 0 Then
        
        'pMenu& = GetMenu(x&)
        
        pCount& = GetMenuItemCount(pMenu&)
        Debug.Print "pMenu = " + Str(pMenu)
        Debug.Print "Count = " + Str(pCount)
        
    Else
        
        Debug.Print "Pop up not found!"
        
    End If
    
End Function

Public Function Get_Status() As String
    
    Dim Kazaa As Long, Msctlsstatusbar As Long
    Dim tempText As String
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    Msctlsstatusbar& = FindWindowEx(Kazaa&, 0&, "msctls_statusbar32", vbNullString)
    tempText = GetCaption(Msctlsstatusbar)
    Debug.Print "Text = " + tempText
    
End Function
Public Sub Home_Que() 'Moves to first item in que
    
    Dim hQue As Long, X As Integer
    
    hQue = Find_Que
    
    Debug.Print "Enter Home Que"
    
    Call SendMessage(hQue, WM_KEYDOWN, VK_HOME, 0&)
    Call SendMessage(hQue, WM_KEYUP, VK_HOME, 0&)
    
    'Pause (0.5)
    
    Call SendMessage(hQue, 33, 0&, 0&) '1.7.1 Select Item Message
    Call SendMessage(hQue, 33, 132&, 0&) '1.7.2 98 Version required 132& string
    
    'Call SendMessage(hButton, WM_KEYUP, VK_SPACE, 0&)
    Debug.Print "Exit Home Que"
    
End Sub

Public Function Check_Leech() As String
    
    Dim hFiles As Long, hSearch As Long
    
    hFiles = Find_Files
    hSearch = Find_Search
    
    If TreeView_GetCount(hFiles) > 1 Then
        '--- XP Test Code ---
        'frmMain.lblStatus = "Sharing " + str(TreeView_GetCount(hFiles)) + " files..)"
        'Pause (1)
        
        Check_Leech = "C" 'Cool user not a leech!
        Debug.Print "Cool, non leech detected"
        
        If IsWindowEnabled(Find_Search) Then
            Debug.Print "Stopping Search!"
            Stop_Search
        Else
            Debug.Print "-- 2.0 stop search --"
            ButtonPress (Find_Search)
            
        End If 'Stop the search if files came back
        
    ElseIf IsWindowEnabled(Find_Search) And TreeView_GetCount(hFiles) = 1 Then
        Debug.Print "Wating for files to retun"
        Check_Leech = "W" 'Waiting for files to return
        
    ElseIf IsWindowEnabled(Find_Search) = False And TreeView_GetCount(hFiles) = 1 Then
        Debug.Print "Leech detected!"
        Check_Leech = "L" 'Leech detected
        
    End If
    
End Function

Public Function Find_Search() As Long
    
    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long
    Dim Button As Long
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    Button& = FindWindowEx(X&, 0&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    
    Find_Search = Button&
    
    'Debug.Print "Find Search >> " + GetCaption(Button)
    
End Function
Public Sub View_Traffic()
    
    Dim Win_Buff0 As Long
    Dim strCaption As String
    
    RunKazaaByString ("&Traffic")
    
    '--- Wait for Traffic ----
    Win_Buff0 = FindWindow("KaZaA", vbNullString)
    strCaption = GetCaption(Win_Buff0)
        
    End Sub
Public Sub View_Search()
    
    Dim strCaption As String
    Dim Win_Buff0 As Long
    
    RunKazaaByString ("&Search")
    
    '--- Wait For Search ----
    timTemp = 0
    Win_Buff0 = FindWindow("KaZaA", vbNullString)
    strCaption = GetCaption(Win_Buff0)
    
    Do While InStr(1, LCase(strCaption), "search") = 0 '
        
        Win_Buff0 = FindWindow("KaZaA", vbNullString)
        strCaption = GetCaption(Win_Buff0)
        
        DoEvents
        If timTemp > 5 Then
            'frmMain.lblStatus = "Error switching to search view!"
            'Pause (1)
            Exit Do
        End If
        
        Loop
        
    End Sub
Public Sub Enable_ToolBar()
    
    Dim rebarwindow As Long, tb1 As Long, tb2 As Long
    Dim Kazaa As Long, i As Long, click As Long
    Dim toolCount As Long, hTool As Long, X As Long
    Dim tbb As TBBUTTON, hNext As Long, nCount As Long
    Dim tb3 As Long, Button As Long, afxframeorviews As Long
    Dim afxmdiframes As Long, mdiclient As Long
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    rebarwindow& = FindWindowEx(Kazaa&, 0&, "rebarwindow32", vbNullString)
    tb1& = FindWindowEx(rebarwindow&, 0&, "toolbarwindow32", vbNullString)
    tb2& = FindWindowEx(rebarwindow&, tb1&, "toolbarwindow32", vbNullString)
    
    'toolCount = SendMessage(toolbarwindow, TB_BUTTONCOUNT, 0, 0)
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    Button& = FindWindowEx(X&, 0&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    Button& = FindWindowEx(X&, Button&, "button", vbNullString)
    
    Debug.Print "next = " + Str(tb2)
    
    'Debug.Print "Button = " + Str(button)
    'Debug.Print "Enabled = " + Str(IsWindowEnabled(button))
    'click = PostMessage(tb1&, WM_COMMAND, 1245, 0&)
    'Debug.Print GetCaption(kazaa)
    'Exit Sub
    click = PostMessage(tb2&, WM_COMMAND, 32909, tb1&) '--- Check user 1.7.2
    MsgBox "Text 1"
    click = PostMessage(tb2&, WM_COMMAND, 32981, tb1&) '--- Send Message to user 1.7.2
    MsgBox "Text 2"
    
    Exit Sub
    For i = 32907 To 32999
        'Click = PostMessage(tb2&, WM_COMMAND, i, tb1&)
        'Pause (0.0001)
        'Pause (1)
        'Sleep 100
        Sleep 10
        'Sleep 5
        'Sleep 1
        Kazaa& = FindWindow("kazaa", vbNullString)
        'If IsWindowEnabled(button) = 1 Then
        MsgBox ">> Good find " + Str(i), vbExclamation
        'End If
        
        If Kazaa = 0 Then
            MsgBox ">> Crash at " + Str(i), vbExclamation
        End If
        'Switch to
        DoEvents
        
        Next
        
        'For i = 200 To 299
        '    click = PostMessage(tb1&, WM_COMMAND, i, tb2&)
        'Next
        
        MsgBox "Done", vbExclamation
        
        'If click = 1 Then
        '    Debug.Print "Hallo"
        '
        'End If
        
        'click = SendMessage(toolbarwindow, TB_PRESSBUTTON, i, False)
        'Debug.Print "Click = " + Str(click)
        'Debug.Print SendMessage(toolbarwindow, TB_GETBUTTON, i, tbb)
        'Debug.Print ">> " + Str(tbb.idCommand)
        'Debug.Print ">> " + Str(tbb.iString)
        
        'hTool = SendMessage(toolbarwindow, TB_GETBUTTON, i, ByVal VarPtr(lpTBBtn))
        'Next
        
        'MsgBox "Done"
        
    End Sub
Public Sub Hit_Header()
    
    Dim Kazaa As Long, mdicleint As Long
    Dim afxframeorviews As Long, X As Long
    Dim afxmdiframes As Long, mdiclient As Long
    Dim xstatic As Long, xsysheader As Long
    Dim rebarwindow As Long
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    X& = FindWindowEx(afxmdiframes&, X&, "#32770", vbNullString)
    rebarwindow& = FindWindowEx(X&, 0&, "rebarwindow32", vbNullString)
    
    Debug.Print "Header = " + Str(rebarwindow)
    ButtonPress (rebarwindow)
    
    'Call PostMessage(xsysheader, WM_LBUTTONDOWN, 10&, ByVal 210&)
    'Call PostMessage(xsysheader, WM_LBUTTONUP, 0, ByVal 0&)
    
    'Call SendMessage(xrebarwindow, WM_KEYDOWN, VK_SPACE, 0&)
    'Call SendMessage(rebarwindow, WM_KEYUP, VK_SPACE, 0&)
    
End Sub

Public Function Find_Download() As Long
    
    Dim Kazaa As Long, mdicleint As Long
    Dim afxframeorviews As Long, X As Long
    Dim afxmdiframes As Long, mdiclient As Long
    Dim xstatic As Long, xsystreeview As Long
    Dim rebarwindow As Long
    
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    X& = FindWindowEx(afxmdiframes&, X&, "#32770", vbNullString)
    rebarwindow& = FindWindowEx(X&, 0&, "rebarwindow32", vbNullString)
    
    '--- Beginning of MDI frame ---
    Find_Download = rebarwindow
    
End Function
Public Sub Pause_Download()
    
    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long, lngStatic As Long
    Dim lngSysTreeView As Long, hLoop As Long
    Dim hItem As Long
    
    'Set scan timer to scan timer
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    lngStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    lngSysTreeView& = FindWindowEx(lngStatic&, 0&, "systreeview32", vbNullString)
    
    If lngSysTreeView <> 0 Then
        
        'Home the download que
        'Call SendMessage(lngSysTreeView, WM_KEYDOWN, VK_HOME, 0&)
        'Call SendMessage(lngSysTreeView, WM_KEYUP, VK_HOME, 0&)
        Call SendMessage(lngSysTreeView, 33, 0&, 0&) '1.7.1 Select Item Message
        Call SendMessage(lngSysTreeView, 33, 132&, 0&)
        
        RunKazaaByString ("Pause Download")
        Call SendMessage(lngSysTreeView, &HEF3A, 0, 0) ' Pause Download
        '' Call SendMessage(lngSysTreeView, &HEF3B, 0)  ' To cancel download
        
    End If
    
End Sub
Public Sub Resume_Download()
    
    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long, lngStatic As Long
    Dim lngSysTreeView As Long, hLoop As Long
    Dim hItem As Long
    
    'Set scan timer to scan timer
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    lngStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    lngSysTreeView& = FindWindowEx(lngStatic&, 0&, "systreeview32", vbNullString)
    
    If lngSysTreeView <> 0 Then
        
        'Home the download que
        'Call SendMessage(lngSysTreeView, WM_KEYDOWN, VK_HOME, 0&)
        'Call SendMessage(lngSysTreeView, WM_KEYUP, VK_HOME, 0&)
        Call SendMessage(lngSysTreeView, 33, 0&, 0&) '1.7.1 Select Item Message
        Call SendMessage(lngSysTreeView, 33, 132&, 0&)
        
        RunKazaaByString ("Resume Download")
        Call PostMessage(lngSysTreeView, WM_COMMAND, &HEF3C, 0) 'Resume download
        Call PostMessage(lngSysTreeView, WM_COMMAND, &H80DA, 0) 'Find more source
        
    Else
        Debug.Print "Error finding download tree"
        
    End If
    
End Sub

Public Sub Download_Accelerate()
    
    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long, lngStatic As Long
    Dim lngSysTreeView As Long, hLoop As Long
    Dim hItem As Long
    
    'Set scan timer to scan timer
    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    lngStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    lngSysTreeView& = FindWindowEx(lngStatic&, 0&, "systreeview32", vbNullString)
    
    If lngSysTreeView <> 0 Then
        '-- Clear downloaded
        Call SendMessage(Kazaa, WM_COMMAND, &H8070, Null) '---- Clear Completed Items ---
        
        'Home the download que
        Call SendMessage(lngSysTreeView, WM_KEYDOWN, VK_HOME, 0&)
        Call SendMessage(lngSysTreeView, WM_KEYUP, VK_HOME, 0&)
        Call SendMessage(lngSysTreeView, 33, 0&, 0&) '1.7.1 Select Item Message
        Call SendMessage(lngSysTreeView, 33, 132&, 0&)
        
        '-- Count items  ---
        hLoop = TreeView_GetCount(lngSysTreeView)
        
        If hLoop > 0 Then
            'Display status message
            
            For X = 1 To hLoop
                RunKazaaByString ("Resume Download")
                Call SendMessage(Kazaa, WM_COMMAND, &HEF3C, Null)
                Call SendMessage(Kazaa, WM_COMMAND, &H80DA, Null)
                
                'Pause (0.1)
                
                Call SendMessage(lngSysTreeView, WM_KEYDOWN, VK_DOWN, 0&)
                Call SendMessage(lngSysTreeView, WM_KEYUP, VK_DOWN, 0&)
                Next
                
            End If
        End If
        
    End Sub

Public Sub Collapse_TreeView()

    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long, lngStatic As Long
    Dim lngSysTreeView As Long, hLoop As Long
    Dim hItem As Long

    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    lngStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    lngSysTreeView& = FindWindowEx(lngStatic&, 0&, "systreeview32", vbNullString)
    
    hLoop = TreeView_GetCount(lngSysTreeView)
    
    For X = 1 To hLoop
    
        Call SendMessage(lngSysTreeView, WM_KEYDOWN, VK_LEFT, 0&)
        Call SendMessage(lngSysTreeView, WM_KEYUP, VK_LEFT, 0&)
    
        'Pause (0.1)
    
        Call SendMessage(lngSysTreeView, WM_KEYDOWN, VK_DOWN, 0&)
        Call SendMessage(lngSysTreeView, WM_KEYUP, VK_DOWN, 0&)
    
    Next X
End Sub

Public Sub PauseSpecified_Download(Download As Integer)
Dim i

Home_Que

For i = 1 To Download
    Kazaa_Down
Next i

Pause_Download

End Sub

Public Sub ResumeSpecified_Download(Download As Integer)
Dim i

Home_Que

For i = 1 To Download
    Kazaa_Down
Next i

Resume_Download

End Sub

Public Sub CancelSpecified_Download(Download As Integer)
Dim i

Dim Kazaa As Long, mdiclient As Long
Dim afxframeorviews As Long, afxmdiframes As Long
Dim X As Long, hStatic As Long, lngStatic As Long
Dim lngSysTreeView As Long, hLoop As Long
Dim hItem As Long

Kazaa& = FindWindow("kazaa", vbNullString)
mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
lngStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
lngSysTreeView& = FindWindowEx(lngStatic&, 0&, "systreeview32", vbNullString)

Home_Que

For i = 1 To Download
    Kazaa_Down
Next i

RunKazaaByString ("&Hide in systray")
'Pause (0.5)
RunKazaaByString ("&Cancel Download")

SendKeys ("{ENTER}")

'Pause (1)

RunKazaaByString ("&Hide in systray")

End Sub

Public Sub CancelAll_Downloads()

Home_Que

RunKazaaByString ("Cancel &All Downloads")
RunKazaaByString ("Clear Downloaded and &Erroneous")

End Sub

Public Sub Clear_Downloaded()

Home_Que

RunKazaaByString ("Clear Downloaded and &Erroneous")

End Sub

Public Sub Kazaa_SetFocus()
Dim lFoundWindow As Long
Dim lOK As Long
Dim lOK1 As Long
Dim X As Variant
    lFoundWindow = FindWindow(vbNullString, "KaShutdown!")
    If lFoundWindow = 0 Then
        ' Did Not Find Window
        ' You Could Use The Shell Command Here If You Wanted
        ' To Start It in This Instance.
    Else
        lOK = SetForegroundWindow(lFoundWindow)
        
        ' You may only need one of these lines, This needed both
        ' Due to the nature of the App I was Selecting so I will
        ' Leave ' both lines
            lOK1 = ShowWindow(lFoundWindow, 9)
            lOK1 = ShowWindow(lFoundWindow, 10)
        ' End
        
        lFoundWindow = 0
        lOK = 0
        lOK1 = 0
    End If
End Sub

Public Sub FindQue_Length()
    Dim Kazaa As Long, mdiclient As Long
    Dim afxframeorviews As Long, afxmdiframes As Long
    Dim X As Long, hStatic As Long, lngStatic As Long
    Dim lngSysTreeView As Long, hLoop As Long
    Dim hItem As Long

    Kazaa& = FindWindow("kazaa", vbNullString)
    mdiclient& = FindWindowEx(Kazaa&, 0&, "mdiclient", vbNullString)
    afxframeorviews& = FindWindowEx(mdiclient&, 0&, "afxframeorview42s", vbNullString)
    afxmdiframes& = FindWindowEx(afxframeorviews&, 0&, "afxmdiframe42s", vbNullString)
    X& = FindWindowEx(afxmdiframes&, 0&, "#32770", vbNullString)
    lngStatic& = FindWindowEx(X&, 0&, "static", vbNullString)
    lngSysTreeView& = FindWindowEx(lngStatic&, 0&, "systreeview32", vbNullString)
    
    QueLength = TreeView_GetCount(lngSysTreeView)
End Sub
