VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Main 
   Caption         =   "Special Folders Example"
   ClientHeight    =   4350
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8160
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   2400
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lst_Directorys 
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_Browse 
         Caption         =   "&Browse Directory"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFile_Save 
         Caption         =   "&Save to File"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFile_Blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAbout_Vote 
         Caption         =   "&Vote on PSC"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuAbout_Blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout_Show 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Sub Launch(ByVal Path As String)
Dim hWnd As Long
    Let hWnd& = GetDesktopWindow&
    Call ShellExecute(hWnd&, vbNullString, Path$, vbNullString, vbNullString, vbNormalFocus)
End Sub
Private Sub Form_Load()
    Call lst_Directorys.AddItem("Common Administrative Tools: " & GetFolderPath(FOLDER_COMMON_ADMIN_TOOLS))
    Call lst_Directorys.AddItem("Common Application Data: " & GetFolderPath(FOLDER_COMMON_APPLICATION_DATA))
    Call lst_Directorys.AddItem("Common Desktop: " & GetFolderPath(FOLDER_COMMON_DESKTOP))
    Call lst_Directorys.AddItem("Common Documents: " & GetFolderPath(FOLDER_COMMON_DOCUMENTS))
    Call lst_Directorys.AddItem("Common Favorites: " & GetFolderPath(FOLDER_COMMON_FAVORITES))
    Call lst_Directorys.AddItem("Common Program Files: " & GetFolderPath(FOLDER_COMMON_PROGRAM_FILES))
    Call lst_Directorys.AddItem("Common Programs: " & GetFolderPath(FOLDER_COMMON_PROGRAMS))
    Call lst_Directorys.AddItem("Common Start Menu: " & GetFolderPath(FOLDER_COMMON_START_MENU))
    Call lst_Directorys.AddItem("Common Startup: " & GetFolderPath(FOLDER_COMMON_STARTUP))
    Call lst_Directorys.AddItem("Common Temlates: " & GetFolderPath(FOLDER_COMMON_TEMPLATES))
    Call lst_Directorys.AddItem("All Users Desktop: " & GetFolderPath(FOLDER_ALL_USERS_DESKTOP))
    Call lst_Directorys.AddItem("Application Data: " & GetFolderPath(FOLDER_APPLICATION_DATA))
    Call lst_Directorys.AddItem("Cookies: " & GetFolderPath(FOLDER_COOKIES))
    Call lst_Directorys.AddItem("Desktop: " & GetFolderPath(FOLDER_DESKTOP))
    Call lst_Directorys.AddItem("Favorites: " & GetFolderPath(FOLDER_FAVORITES))
    Call lst_Directorys.AddItem("Fonts: " & GetFolderPath(FOLDER_FONTS))
    Call lst_Directorys.AddItem("History: " & GetFolderPath(FOLDER_HISTORY))
    Call lst_Directorys.AddItem("Internet Cache: " & GetFolderPath(FOLDER_INTERNET_CACHE))
    Call lst_Directorys.AddItem("Local Settings: " & GetFolderPath(FOLDER_LOCAL_SETTINGS))
    Call lst_Directorys.AddItem("My Documents: " & GetFolderPath(FOLDER_MY_DOCUMENTS))
    Call lst_Directorys.AddItem("My Pictures: " & GetFolderPath(FOLDER_MY_PICTURES))
    Call lst_Directorys.AddItem("User Profile: " & GetFolderPath(FOLDER_PROFILE))
    Call lst_Directorys.AddItem("Program Files: " & GetFolderPath(FOLDER_PROGRAM_FILES))
    Call lst_Directorys.AddItem("Start Menu Programs: " & GetFolderPath(FOLDER_PROGRAMS))
    Call lst_Directorys.AddItem("Recent Files: " & GetFolderPath(FOLDER_RECENT))
    Call lst_Directorys.AddItem("Send To: " & GetFolderPath(FOLDER_SEND_TO))
    Call lst_Directorys.AddItem("Start Menu: " & GetFolderPath(FOLDER_START_MENU))
    Call lst_Directorys.AddItem("Start Up: " & GetFolderPath(FOLDER_STARTUP))
    Call lst_Directorys.AddItem("System: " & GetFolderPath(FOLDER_SYSTEM))
    Call lst_Directorys.AddItem("Temperary Internet Files: " & GetFolderPath(FOLDER_TEMP_INTERNET))
    Call lst_Directorys.AddItem("Templates: " & GetFolderPath(FOLDER_TEMPLATES))
    Call lst_Directorys.AddItem("Windows Folder: " & GetFolderPath(FOLDER_WINDOWS))
    Call Form_Resize
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Msg As VbMsgBoxResult
    Let Msg = MsgBox("Do you really intend on leaving?", vbYesNo + vbQuestion, "Special Folders Example")
    If Msg = vbNo Then
        Let Cancel = True
    End If
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 8280& Then
            Let Me.Width = 8280&
        ElseIf Me.Height < 4860& Then
            Let Me.Height = 4860&
        End If
        Let lst_Directorys.Top = 0&
        Let lst_Directorys.Left = 0&
        Let lst_Directorys.Width = Me.Width - 100&
        Let lst_Directorys.Height = Me.Height - 640&
    End If
End Sub
Private Sub lst_Directorys_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Position As Integer
    If lst_Directorys.List(lst_Directorys.ListIndex) <> "" Then
        Let Position% = InStr(1&, lst_Directorys.List(lst_Directorys.ListIndex), ":\", vbTextCompare) - 1&
        Let lst_Directorys.ToolTipText = Mid$(lst_Directorys.List(lst_Directorys.ListIndex), Position%, Len(lst_Directorys.List(lst_Directorys.ListIndex)) - Position% + 1&)
    End If
End Sub
Private Sub mnuAbout_Show_Click()
Dim dspAbout As String
    Let dspAbout$ = "Special Folders Example by: dacryonic" & vbNewLine & vbNewLine & "Just add this module to your project and you will be able to get up to 43 different special folders like: System and Windows Folders." & vbNewLine & vbNewLine & "You may add this module to your project, modify, what ever you choose without notifying me what-so-ever. this is to learn from after all so play with it." & vbNewLine & vbNewLine & "thanks dacryonic"
    Call MsgBox(dspAbout$, vbInformation, "About")
End Sub
Private Sub mnuAbout_Vote_Click()
Dim theURL As String
    Let theURL$ = "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=56570&lngWId=1"
    Call Launch(theURL$)
End Sub
Private Sub mnuFile_Browse_Click()
Dim Position As Integer
    If lst_Directorys.List(lst_Directorys.ListIndex) <> "" Then
        Let Position% = InStr(1&, lst_Directorys.List(lst_Directorys.ListIndex), ":\", vbTextCompare) - 1&
        Call Launch(Mid$(lst_Directorys.List(lst_Directorys.ListIndex), Position%, Len(lst_Directorys.List(lst_Directorys.ListIndex)) - Position% + 1&))
    End If
End Sub
Private Sub mnuFile_Save_Click()
Dim Doit As Long, intFreeFile As Integer, Msg As VbMsgBoxResult
Dim Filename As String
    On Error Resume Next
    Let cDlg.CancelError = True
    Let cDlg.Filter = "Text Files|*.txt"
    Call cDlg.ShowSave
    If Not Err Then
        Let Filename$ = cDlg.Filename
        If Right(LCase(Filename$), 4&) <> ".txt" Then
            Let Filename$ = Mid$(Filename$, 1&, Len(Filename$) - 4&) & ".txt"
        End If
        If Len(Dir$(Filename$, vbHidden)) > 0& Then
            Let Msg = MsgBox("The file you selected exists, Do you want to overwrite?", vbYesNo + vbQuestion, "File Exists")
            If Msg = vbYes Then
                Call Kill(Filename$)
            Else
                Exit Sub
            End If
        End If
        Let intFreeFile% = FreeFile
        Open Filename$ For Output As #intFreeFile
            For Doit& = 0& To lst_Directorys.ListCount - 1&
                DoEvents
                Print #intFreeFile, lst_Directorys.List(Doit&)
            Next Doit&
        Close #intFreeFile%
        Call Launch(Filename$)
    Else
        Call MsgBox("There was an error saveing." & vbNewLine & vbNewLine & Err.Description, vbCritical + vbInformation, "Error #" & Err.Number)
        Call Err.Clear
    End If
Resume Next: End Sub
Private Sub mnuFile_Copy_Click()
Dim Position As Integer
    If lst_Directorys.List(lst_Directorys.ListIndex) <> "" Then
        Let Position% = InStr(1&, lst_Directorys.List(lst_Directorys.ListIndex), ":\", vbTextCompare) - 1&
        Call Clipboard.Clear
        Call Clipboard.SetText(Mid$(lst_Directorys.List(lst_Directorys.ListIndex), Position%, Len(lst_Directorys.List(lst_Directorys.ListIndex)) - Position% + 1&))
    End If
End Sub
Private Sub mnuFile_Exit_Click()
    Call Unload(Me)
End Sub
