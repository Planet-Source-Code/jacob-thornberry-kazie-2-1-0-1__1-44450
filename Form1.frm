VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   ClientHeight    =   375
   ClientLeft      =   9480
   ClientTop       =   8595
   ClientWidth     =   5100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdexit 
      Height          =   375
      Left            =   4560
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exit KazIE"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdcanceluploads 
      Height          =   375
      Left            =   4080
      Picture         =   "Form1.frx":13CC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel all uploads"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdresumeall 
      Height          =   375
      Left            =   3720
      Picture         =   "Form1.frx":1E2E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Resume all downloads"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdpauseall 
      Height          =   375
      Left            =   3360
      Picture         =   "Form1.frx":25F0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Pause all downloads"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdsearch 
      Height          =   375
      Left            =   2880
      Picture         =   "Form1.frx":2D52
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Go to Search window"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdtraffic 
      Height          =   375
      Left            =   2400
      Picture         =   "Form1.frx":3694
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Go to Traffic window"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdcancel 
      Height          =   375
      Left            =   2040
      Picture         =   "Form1.frx":3FD6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel selected download"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdResume 
      Height          =   375
      Left            =   1680
      Picture         =   "Form1.frx":4678
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Resume selected download"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdPause 
      Height          =   375
      Left            =   1320
      Picture         =   "Form1.frx":4D1A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Pause selected download"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdDisconnect 
      Height          =   375
      Left            =   840
      Picture         =   "Form1.frx":541C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Disconnect"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdConnect 
      Height          =   375
      Left            =   360
      Picture         =   "Form1.frx":5CFE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Connect"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdKazaa 
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":6580
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Start Kazaa Lite"
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnumin 
         Caption         =   "&Minimize to System Tray             ..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "systray"
      Visible         =   0   'False
      Begin VB.Menu mnuKazaa 
         Caption         =   "Run Kazaa"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause Download"
      End
      Begin VB.Menu mnuResume 
         Caption         =   "Resume Download"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Download"
      End
      Begin VB.Menu mnuTraffic 
         Caption         =   "Show Traffic Window"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Show Search Window"
      End
      Begin VB.Menu mnuPauseall 
         Caption         =   "Pause All Downloads"
      End
      Begin VB.Menu mnuResumeall 
         Caption         =   "Resume All Downloads"
      End
      Begin VB.Menu mnuUploads 
         Caption         =   "Cancel All Uploads                           ..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' KazIE
' Written by Jacob Thornberry
' This code is distributed freely for you to enhance the features, or add new functionality.
' If you add a significant feature, please contact me (Jacob) at jakert50@hotmail.com
' Long live open source and long live Kazaa Lite!!
' You can download Kazaa Lite from http://www.k-lite.tk
' Download source code and new versions of KazIE from jaker.edskes.com
' KazIE is distributed as freeware.  This means that you may use the program and/or
' source code in any of your own software without any fee, subject to acknowledgement
' as appropriate.  It does not mean that you can attempt to pass this work off as
' your own, and it is unethical to do.
' "SuperCodeXP" code by Shantibhushan (s.naik@ ebsolutech.com)
' SysTray code by Michael Cowell (VorTech Software Web: www.vortech.freeservers.com) Copyrighted 2000-2001

Option Explicit
Private mSystem       As SystemInteroperatability.System ' = Nothing
Private ontop As New clsOnTop
Dim kazaaexe As String

Private Property Set System(ByRef System As SystemInteroperatability.System)
  Rem /* NOTE: Please remember to unitialize and dereference any objects used,
  Rem  *       We need to also indicate the interface that we no longer wish
  Rem  *       our window to be used, so we must call UnInit() to deinitialize
  Rem  *       it,
  Rem  */
  If (System Is Nothing) Then
    Call mSystem.Window.UnInit    ' /* Indicate the interface that we no longer
                                  '  * need it to keep track of our form window,
                                  '  */
  End If
  
  Set mSystem = System
End Property

Private Property Get System() As SystemInteroperatability.System
  Rem /* NOTE: SuperCode users, please pay close attention to this property!
  Rem  * To use sophisticated and advanced features, you must correctly indi-
  Rem  * the interface which window it has to use. So an important call is
  Rem  * Init() method for Window Class. You must set it to point to your
  Rem  * window,
  Rem  */
  If (mSystem Is Nothing) Then  ' /* Check if we are accessing this member for first
                                '  * time,
                                '  */
    Set mSystem = New SystemInteroperatability.System ' /* Create a new instance, */
    Call mSystem.Window.Init(Me)  ' /* This call instructs the interface to use our
                                  '  * form window,
                                  '  */
  End If
  
  Set System = mSystem          ' /* Pass reference, */
End Property

Private Sub cmdCancel_Click()
    RunKazaaByString ("&Traffic")
    RunKazaaByString ("&Cancel Download")
End Sub
Private Sub cmdcancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\light\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub
Private Sub cmdcanceluploads_Click()
    RunKazaaByString ("&Traffic")
    RunKazaaByString ("Cancel All Upload&s")
    End Sub
Private Sub cmdcanceluploads_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\light\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub
Private Sub cmdConnect_Click()
    RunKazaaByString ("&Connect")
End Sub
Private Sub cmdConnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\light\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub
Private Sub cmdDisconnect_Click()
    RunKazaaByString ("&Disconnect")
End Sub
Private Sub cmdDisconnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\light\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub
Private Sub cmdExit_Click()
    SaveSetting App.Title, "positions", "Yposition", Form1.Top
    SaveSetting App.Title, "positions", "Xposition", Form1.Left
    End
End Sub

Private Sub cmdexit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\light\exit.bmp"))
End Sub

Private Sub cmdKazaa_Click()
        kazaaexe = ReadKey("HKLM\Software\Kazaa\Cloudload\exedir")
        Shell (kazaaexe)
End Sub


Private Sub cmdKazaa_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\light\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub


Private Sub cmdPause_Click()
    RunKazaaByString ("&Traffic")
    RunKazaaByString ("&Pause Download")
End Sub

Private Sub cmdPause_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\light\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub

Private Sub cmdPauseall_Click()
    Dim retint As Integer
    retint = SendPauseCommands
End Sub

Private Sub cmdpauseall_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\light\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub

Private Sub cmdResume_Click()
    RunKazaaByString ("&Traffic")
    RunKazaaByString ("&Resume Download")
End Sub

Private Sub cmdResume_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\light\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub

Private Sub cmdResumeall_Click()
    Dim retint As Integer
    retint = SendResumeCommands
End Sub

Private Sub cmdresumeall_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\light\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub
Private Sub cmdSearch_Click()
    RunKazaaByString ("&Search")
End Sub

Private Sub cmdsearch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\light\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub

Private Sub cmdTraffic_Click()
    RunKazaaByString ("&Traffic")
End Sub

Private Sub cmdtraffic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\light\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub

Private Sub Form_Load()
    Set ontop = New clsOnTop
    ontop.MakeTopMost hWnd
    On Error Resume Next
    Form1.Top = GetSetting(App.Title, "positions", "Yposition")
    Form1.Left = GetSetting(App.Title, "positions", "xposition")
    If App.PrevInstance = True Then
    MsgBox "Another instance of KazIE is already running!", vbCritical + vbOKOnly, "KazIE is already running!"
    End
    Else
    End If
    Set ontop = New clsOnTop
    ontop.MakeTopMost hWnd
    Form1.Height = 405
    If (System.OS > WINDOWSNT400) Then Call System.Window.Animate
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FormDrag Form1 'Makes form1 movable
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case x
        Case 7755:   'Right Click
            PopupMenu Menu
        Case 7725:    'Dbl Left Click
            Form1.Show
    End Select
    cmdKazaa.Picture = LoadPicture(App.Path & ("\Skins\Default\default\green.bmp"))
    cmdConnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\connect.bmp"))
    cmdDisconnect.Picture = LoadPicture(App.Path & ("\Skins\Default\default\disconnect.bmp"))
    cmdPause.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pause.bmp"))
    cmdResume.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resume.bmp"))
    cmdcancel.Picture = LoadPicture(App.Path & ("\Skins\Default\default\cancel.bmp"))
    cmdtraffic.Picture = LoadPicture(App.Path & ("\Skins\Default\default\traffic.bmp"))
    cmdsearch.Picture = LoadPicture(App.Path & ("\Skins\Default\default\search.bmp"))
    cmdpauseall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\pauseall.bmp"))
    cmdresumeall.Picture = LoadPicture(App.Path & ("\Skins\Default\default\resumeall.bmp"))
    cmdcanceluploads.Picture = LoadPicture(App.Path & ("\Skins\Default\default\upload.bmp"))
    cmdexit.Picture = LoadPicture(App.Path & ("\Skins\Default\default\exit.bmp"))
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
    PopupMenu mnu
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuConnect_Click()
    RunKazaaByString ("&Connect")
End Sub

Private Sub mnuDisconnect_Click()
    RunKazaaByString ("&Disconnect")
End Sub

Private Sub mnuEnd_Click()
    SaveSetting App.Title, "positions", "Yposition", Form1.Top
    SaveSetting App.Title, "positions", "Xposition", Form1.Left
    End
End Sub

Private Sub mnuExit_Click()
    SaveSetting App.Title, "positions", "Yposition", Form1.Top
    SaveSetting App.Title, "positions", "Xposition", Form1.Left
    End
End Sub

Private Sub mnuontop_Click()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (System.OS > WINDOWSNT400) Then Call System.Window.Animate(DEACTIVATE_SLIDE_FADE_TRANSITION)
    If Not (System Is Nothing) Then Set System = Nothing
    SaveSetting App.Title, "positions", "Yposition", Form1.Top
    SaveSetting App.Title, "positions", "Xposition", Form1.Left
    End Sub

Private Sub mnuKazaa_Click()
        kazaaexe = ReadKey("HKLM\Software\Kazaa\Cloudload\exedir")
        Shell (kazaaexe)
End Sub

Private Sub mnumin_Click()
    Form1.Hide
    try.cbSize = Len(try)
    try.hWnd = Me.hWnd
    try.uID = vbNull
    try.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    try.uCallBackMessage = WM_MOUSEMOVE

    'To Change the Icon Displayed in the systray
    'Change the Forms Icon
    'This uses whatever Icon the Form Displays
    try.hIcon = Me.Icon

    'Tool Tip
    try.szTip = "KazIE MiniMode" & vbNullChar

    Call Shell_NotifyIcon(NIM_ADD, try)
    Call Shell_NotifyIcon(NIM_MODIFY, try)

    'If u just want the systay icon to appear at start Hide the Form
    'Me.Hide
End Sub

Private Sub mnuPause_Click()
    RunKazaaByString ("&Traffic")
    RunKazaaByString ("&Pause Download")
End Sub

Private Sub mnuPauseall_Click()
    Dim retint As Integer
    retint = SendPauseCommands
End Sub

Private Sub mnuRestore_Click()
    Form1.Show
End Sub

Private Sub mnuResumeall_Click()
    Dim retint As Integer
    retint = SendResumeCommands
End Sub

Private Sub mnuSearch_Click()
    RunKazaaByString ("&Search")
End Sub

Private Sub mnuTraffic_Click()
    RunKazaaByString ("&Traffic")
End Sub

Private Sub mnuUploads_Click()
    RunKazaaByString ("&Traffic")
    RunKazaaByString ("Cancel All Upload&s")
End Sub
