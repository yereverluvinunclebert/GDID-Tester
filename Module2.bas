Attribute VB_Name = "Module2"
'---------------------------------------------------------------------------------------
' Module    : Module2
' Author    : beededea
' Date      : 08/08/2026
' Purpose   : general purpose functions for registry, folders, files and properties for same
'---------------------------------------------------------------------------------------

Option Explicit



'------------------------------------------------------ STARTS
'constants and APIs defined for querying the registry
Public Const HKEY_CURRENT_USER As Long = &H80000001
Private Const REG_SZ  As Long = 1                          ' Unicode nul terminated string

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' APIs for useful functions START
#If twinbasic Then
    Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongLong, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If
' APIs for useful functions END
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' API and enums for acquiring the special folder paths
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Public Enum FolderEnum ' has to be public
    feCDBurnArea = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feCommonAppData = 35 ' \Docs & Settings\All Users\Application Data
    feCommonAdminTools = 47 ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
    feCommonDesktop = 25 ' \Docs & Settings\All Users\Desktop
    feCommonDocs = 46 ' \Docs & Settings\All Users\Documents
    feCommonPics = 54 ' \Docs & Settings\All Users\Documents\Pictures
    feCommonMusic = 53 ' \Docs & Settings\All Users\Documents\Music
    feCommonStartMenu = 22 ' \Docs & Settings\All Users\Start Menu
    feCommonStartMenuPrograms = 23 ' \Docs & Settings\All Users\Start Menu\Programs
    feCommonTemplates = 45 ' \Docs & Settings\All Users\Templates
    feCommonVideos = 55 ' \Docs & Settings\All Users\Documents\My Videos
    feLocalAppData = 28 ' \Docs & Settings\User\Local Settings\Application Data
    feLocalCDBurning = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feLocalHistory = 34 ' \Docs & Settings\User\Local Settings\History
    feLocalTempInternetFiles = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feProgramFiles = 38 ' \Program Files
    feProgramFilesCommon = 43 ' \Program Files\Common Files
    'feRecycleBin = 10 ' ???
    feUser = 40 ' \Docs & Settings\User
    feUserAdminTools = 48 ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
    feUserAppData = 26 ' \Docs & Settings\User\Application Data
    feUserCache = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feUserCookies = 33 ' \Docs & Settings\User\Cookies
    feUserDesktop = 16 ' \Docs & Settings\User\Desktop
    feUserDocs = 5 ' \Docs & Settings\User\My Documents
    feUserFavorites = 6 ' \Docs & Settings\User\Favorites
    feUserMusic = 13 ' \Docs & Settings\User\My Documents\My Music
    feUserNetHood = 19 ' \Docs & Settings\User\NetHood
    feUserPics = 39 ' \Docs & Settings\User\My Documents\My Pictures
    feUserPrintHood = 27 ' \Docs & Settings\User\PrintHood
    feUserRecent = 8 ' \Docs & Settings\User\Recent
    feUserSendTo = 9 ' \Docs & Settings\User\SendTo
    feUserStartMenu = 11 ' \Docs & Settings\User\Start Menu
    feUserStartMenuPrograms = 2 ' \Docs & Settings\User\Start Menu\Programs
    feUserStartup = 7 ' \Docs & Settings\User\Start Menu\Programs\Startup
    feUserTemplates = 21 ' \Docs & Settings\User\Templates
    feUserVideos = 14  ' \Docs & Settings\User\My Documents\My Videos
    feWindows = 36 ' \Windows
    feWindowFonts = 20 ' \Windows\Fonts
    feWindowsResources = 56 ' \Windows\Resources
    feWindowsSystem = 37 ' \Windows\System32
End Enum
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' APIs, constants and types defined for determining existence of files and folders
Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
 
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
     
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                            lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
'API Functions to read/write information from INI File
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------ ENDS



' General member property variables declared

Private m_sgsSettingsDir As String
Private m_sgsSettingsFile As String

'---------------------------------------------------------------------------------------
' Procedure : writeRegistry
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : write to the registry
'---------------------------------------------------------------------------------------
'
Public Sub writeRegistry(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String, ByRef strData As String)

    Dim keyhand As Long: keyhand = 0
    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    On Error GoTo writeRegistry_Error

    unusedReturnValue = RegCreateKey(hKey, strPath, keyhand)
    unusedReturnValue = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
    unusedReturnValue = RegCloseKey(keyhand)

   On Error GoTo 0
   Exit Sub

writeRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistry of module module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getstring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function getstring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String) As String

    Dim keyhand As Long: keyhand = 0
    Dim lResult As Long: lResult = 0
    Dim strBuf As String: strBuf = vbNullString
    Dim lDataBufSize As Long: lDataBufSize = 0
    Dim intZeroPos As Integer: intZeroPos = 0
    Dim rvar As Integer: rvar = 0

    Dim lValueType As Variant ' cannot initialise

    On Error GoTo getstring_Error

    rvar = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
        Dim ERROR_SUCCESS As Variant
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

getstring_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getstring of Module Common"
End Function




'---------------------------------------------------------------------------------------
' Procedure : gsSettingsDir
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsSettingsDir() As String

    On Error GoTo gsSettingsDir_Error

    gsSettingsDir = m_sgsSettingsDir

    On Error GoTo 0
    Exit Property

gsSettingsDir_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSettingsDir of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsSettingsDir
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsSettingsDir(ByVal sgsSettingsDir As String)

    On Error GoTo gsSettingsDir_Error

    m_sgsSettingsDir = sgsSettingsDir

    On Error GoTo 0
    Exit Property

gsSettingsDir_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSettingsDir of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsSettingsFile
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsSettingsFile() As String

    On Error GoTo gsSettingsFile_Error

    gsSettingsFile = m_sgsSettingsFile

    On Error GoTo 0
    Exit Property

gsSettingsFile_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSettingsFile of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsSettingsFile
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsSettingsFile(ByVal sgsSettingsFile As String)

    On Error GoTo gsSettingsFile_Error

    m_sgsSettingsFile = sgsSettingsFile

    On Error GoTo 0
    Exit Property

gsSettingsFile_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSettingsFile of Module Module1"

End Property



'---------------------------------------------------------------------------------------
' Procedure : fFExists
' Author    : RobDog888 https://www.vbforums.com/member.php?17511-RobDog888
' Date      : 19/07/2023
' Purpose   : Test for file existence using the OpenFile API
'---------------------------------------------------------------------------------------
'
Public Function fFExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    On Error GoTo fFExists_Error
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        fFExists = True
    Else
        fFExists = False
    End If

   On Error GoTo 0
   Exit Function

fFExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFExists of Module Module1"
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : zeezee https://www.vbforums.com/member.php?90054-zeezee
' Date      : 19/07/2023
' Purpose   : Test for file existence using the PathFileExists API
'---------------------------------------------------------------------------------------
'
Public Function fDirExists(ByVal pstrFolder As String) As Boolean
   On Error GoTo fDirExists_Error

    fDirExists = (PathFileExists(pstrFolder) = 1)
    If fDirExists Then fDirExists = (PathIsDirectory(pstrFolder) <> 0)

   On Error GoTo 0
   Exit Function

fDirExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDirExists of Module Module1"
End Function



'
'---------------------------------------------------------------------------------------
' Procedure : fGetINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Get the INI Setting from the File
'---------------------------------------------------------------------------------------
'
Public Function fGetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
   On Error GoTo fGetINISetting_Error
    Const cparmLen As Integer = 500 ' maximum no of characters allowed in the returned string
    Dim sReturn As String * cparmLen ' not going to initialise this with a 500 char string
    Dim sDefault As String * cparmLen
    Dim lLength As Long: lLength = 0

    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
    fGetINISetting = Mid$(sReturn, 1, lLength)

   On Error GoTo 0
   Exit Function

fGetINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetINISetting of module module1"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : sPutINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save a specific INI setting to a specific section of the filename supplied using a key to identify the setting
'---------------------------------------------------------------------------------------
'
Public Sub sPutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)

   On Error GoTo sPutINISetting_Error

    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    unusedReturnValue = WritePrivateProfileString(sHeading, sKey, sSetting, sINIFileName)

   On Error GoTo 0
   Exit Sub

sPutINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sPutINISetting of module module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fSpecialFolder
' Author    : si_the_geek vbforums
' Date      : 17/10/2019
' Purpose   : Returns the path to the specified special folder (AppData etc)
'---------------------------------------------------------------------------------------
'
Public Function fSpecialFolder(ByVal pfe As FolderEnum) As String
    Const MAX_PATH As Integer = 260
    Dim strPath As String: strPath = vbNullString
    Dim strBuffer As String: strBuffer = vbNullString
    
    On Error GoTo fSpecialFolder_Error

    strBuffer = Space$(MAX_PATH)
    If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
    fSpecialFolder = strPath

    On Error GoTo 0
    Exit Function

fSpecialFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fSpecialFolder of Module Module1"
End Function



'---------------------------------------------------------------------------------------
' Procedure : writeLogFile
' Author    : beededea
' Date      : 22/12/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub writeLogFile(ByVal inputStr As String, Optional ByVal timestamp As String)

    Dim FN As Integer: FN = 0

    On Error GoTo writeLogFile_Error

    FN = FreeFile
    
    ' write the error to the log file
    Open gsSettingsDir & "\GDIDChangeLog.log" For Append As FN
    Print #FN, timestamp & " " & inputStr
    Close FN
        
    On Error GoTo 0
    Exit Sub

writeLogFile_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeLogFile of Module common"
            Resume Next
          End If
    End With
End Sub
