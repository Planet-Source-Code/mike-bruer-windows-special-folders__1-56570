Attribute VB_Name = "mod_SpecialFolders"
Option Explicit
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type SHTEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkID As SHTEMID
End Type
Public Enum SHFolders
    FOLDER_TEMP_INTERNET = &H20
    FOLDER_DESKTOP = &H0
    FOLDER_INTERNET_EXPLORER = &H1
    FOLDER_PROGRAMS = &H2
    FOLDER_CONTROL_PANEL = &H3
    FOLDER_PRINTERS = &H4
    FOLDER_MY_DOCUMENTS = &H5
    FOLDER_FAVORITES = &H6
    FOLDER_STARTUP = &H7
    FOLDER_RECENT = &H8
    FOLDER_SEND_TO = &H9
    FOLDER_RECYCLE_BIN = &HA
    FOLDER_START_MENU = &HB
    FOLDER_ALL_USERS_DESKTOP = &H10
    FOLDER_MY_COMPUTER = &H11
    FOLDER_NETWORK_NEIGHBOURHOOD = &H12
    FOLDER_FONTS = &H14
    FOLDER_TEMPLATES = &H15
    FOLDER_APPLICATION_DATA = &H1A
    FOLDER_LOCAL_SETTINGS = &H1C
    FOLDER_INTERNET_CACHE = &H20
    FOLDER_COOKIES = &H21
    FOLDER_HISTORY = &H22
    FOLDER_WINDOWS = &H24
    FOLDER_SYSTEM = &H25
    FOLDER_PROGRAM_FILES = &H26
    FOLDER_MY_PICTURES = &H27
    FOLDER_PROFILE = &H28
    FOLDER_COMMON_APPLICATION_DATA = &H23
    FOLDER_COMMON_START_MENU = &H16
    FOLDER_COMMON_PROGRAMS = &H17
    FOLDER_COMMON_STARTUP = &H18
    FOLDER_COMMON_DESKTOP = &H19
    FOLDER_COMMON_TEMPLATES = &H2D
    FOLDER_COMMON_DOCUMENTS = &H2E
    FOLDER_COMMON_ADMIN_TOOLS = &H2F
    FOLDER_COMMON_FAVORITES = &H1F
    FOLDER_COMMON_PROGRAM_FILES = &H2B
    FOLDER_NONLOCAL_STARTUP = &H1D
    FOLDER_NONLOCAL_COMMON_STARTUP = &H1E
    FOLDER_X86_SYSTEM = &H29
    FOLDER_X86_PROGRAM_FILES = &H2A
    FOLDER_X86_COMMON_PROGRAM_FILES = &H2C
End Enum
Public Function GetFolderPath(ByVal ID As SHFolders) As String
Dim Doit As Long, idList As ITEMIDLIST, FolderPath As String
    Let Doit& = SHGetSpecialFolderLocation(100&, ID, idList)
    If Doit& = 0& Then
        Let FolderPath$ = Space$(512&)
        Let Doit& = SHGetPathFromIDList(ByVal idList.mkID.cb, ByVal FolderPath$)
        Let GetFolderPath = Left$(FolderPath$, InStr(FolderPath$, Chr$(0&)) - 1&) & "\"
        Exit Function
    End If
    Let GetFolderPath$ = ""
End Function
