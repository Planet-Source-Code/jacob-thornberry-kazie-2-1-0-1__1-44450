Attribute VB_Name = "modFunctions"
Option Explicit

'=========Checking OS staff=============
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowsName As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal cCmdShow As Long) As Long

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'Save info
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'Get window functions
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As _
    Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Const PROCESS_ALL_ACCESS As Long = 2035711
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'Menu functions
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long

'Key declares
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_Control = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

'List view / tree view constants
Public Const LVM_FIRST = &H1000
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)


'Windows handle declares
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2

Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_CHAR = &H102
'Public Const WM_USER = &H400
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_SYSCOMMAND = &H112
Public Const SC_CLOSE = &HF060
   
'Find window functions

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Sub ButtonPress(Button As Long)

    Dim click As Long

    click = PostMessage(Button&, WM_LBUTTONDOWN, 0, ByVal 0&)
    'Pause (0.1)
    click = PostMessage(Button&, WM_LBUTTONUP, 0, ByVal 0&)
    
    Debug.Print ">> BP >> Click = " + Str(click)

End Sub


Public Function GetCaption(Window As Long)

'Gets the caption of a window

Dim WText As String, Buff As String, leng As Long

leng& = GetWindowTextLength(Window&)
'Debug.Print "Length = " + Str(Leng)
Buff$ = String(leng&, 0)
Call GetWindowText(Window&, Buff$, leng& + 1)
GetCaption = Trim(Buff$)

End Function
Function FindChildByTitle(parentw, childhand)

Dim firs As Long
Dim firss As Long
Dim room As Long

firs = GetWindow(parentw, 5)
If UCase(GetCaption(firs)) Like UCase(childhand) Then GoTo bone
firs = GetWindow(parentw, GW_CHILD)

While firs
    firss = GetWindow(parentw, 5)
    If UCase(GetCaption(firss&)) Like UCase(childhand) & "*" Then GoTo bone
    Debug.Print "FCT1 >> +" + GetCaption(firss&)
    firs = GetWindow(firs, 2)
    If UCase(GetCaption(firs&)) Like UCase(childhand) & "*" Then GoTo bone
    Debug.Print "FCT2 >> +" + GetCaption(firss&)
    
Wend

FindChildByTitle = 0

bone:
room = firs
FindChildByTitle = room

End Function


Function FindChildByClass(parentw, childhand)

Dim firs As Long
Dim firss As Long
Dim room As Long
firs = GetWindow(parentw, 5)

If UCase(Mid(GetClass(firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
While firs
firss = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
    firs = GetWindow(firs, 2)
If UCase(Mid(GetClass(firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
    Wend
FindChildByClass = 0

bone:
room = firs
FindChildByClass = room

End Function
Public Function GetClass(Wind As Long)

'gets a window's classname

Dim Buff As String

Buff$ = String(255, 0)
Call GetClassName(Wind&, Buff$, 255)

GetClass = Replace(Buff$, Chr$(0), "")
Debug.Print "Class = " + Replace(Buff$, Chr$(0), "")

End Function
Function KillWindow(hwnd)

  Dim PROCESSID As Long
  Dim ExitCode As Long
  Dim MyProcess As Long
  
  Call GetWindowThreadProcessId(hwnd, PROCESSID)
  MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, PROCESSID)
  KillWindow = TerminateProcess(MyProcess, ExitCode)
  Call CloseHandle(MyProcess)
  
End Function


Public Sub KillProcess(TaskId As Long)

'Shutdown process.
'Where TaskId is value returned by previous call to VB Shell

    Dim hProc As Long
    
    hProc = OpenProcess(PROCESS_ALL_ACCESS, 0&, TaskId)
    If hProc <> 0& Then
        Call TerminateProcess(hProc, 0&)
        Call CloseHandle(hProc)
    End If
    
End Sub

Public Function IsWindowsNT() As Boolean

   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
   
End Function


