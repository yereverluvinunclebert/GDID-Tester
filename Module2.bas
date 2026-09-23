Attribute VB_Name = "Module2"
'---------------------------------------------------------------------------------------
' Module    : Module2
' Author    : beededea
' Date      : 08/08/2026
' Purpose   : general purpose functions for registry, folders, files and properties for same
'---------------------------------------------------------------------------------------

Option Explicit

' ** requires windevlib package for 64bit operation using TwinBasic **

'------------------------------------------------------ STARTS
'constants and APIs defined for querying the registry
Public Const HKEY_CURRENT_USER As Long = &H80000001
Private Const REG_SZ  As Long = 1                          ' Unicode nul terminated string

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_READ As Long = &H20019
Private Const KEY_WRITE As Long = &H20006

Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_NO_MORE_ITEMS As Long = 259
Private Const ERROR_MORE_DATA As Long = 234

#If Not WIN64 Then ' VB6 only
    'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long  ' hKey LongPtr, phkResult As LongPtr *
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
    
    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long ' hKey As LongPtr, lpReserved LongPtr *
    Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long ' hKey LongPtr *
    Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long ' hKey LongPtr, phkResult LongPtr *
    Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long  ' hKey LongPtr *
    
    Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
        ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        lpcchName As Long, _
        ByVal lpReserved As Long, _
        ByVal lpClass As String, _
        lpcchClass As Long, _
        lpftLastWriteTime As Any) As Long
        
#End If

'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' APIs for useful functions START
#If Not WIN64 Then ' VB6 only
    Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long  ' hWnd required as longPtr, returns longPtr *
#End If
' APIs for useful functions END
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' API and enums for acquiring the special folder paths
#If Not WIN64 Then ' VB6 only
    Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long '  1 hwnd required As LongPtr, 3 ByVal hToken As LongPtr ' *
#End If

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

#If Not WIN64 Then ' VB6 only
    Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                                lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long ' no longPtrs
    Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long ' no longPtrs
    Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long ' no longPtrs
#End If
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
'API Function to read/write information from INI File start

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long ' ANSI version has no longPtrs, Unicode does, ' UNICODE cparmLen will require a longptr*

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long ' ANSI version has no longPtrs, Unicode does ' UNICODE cparmLen will require a longptr*
    
'API Function to read/write information from INI File start
'------------------------------------------------------ ENDS



' General member property variables declared

Private m_sgsSettingsDir As String
Private m_sgsSettingsFile As String


'---------------------------------------------------------------------------------------
' Procedure : WriteRegistryString
' Author    : chatGPT
' Date      : 22/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function WriteRegistryString( _
    ByVal hKey As Long, _
    ByVal sValueName As String, _
    ByVal sValue As String) As Boolean

    Dim lResult As Long
    Dim lDataSize As Long

    On Error GoTo WriteRegistryString_Error
    
    'REG_SZ requires the terminating NULL to be included.
    lDataSize = Len(sValue) + 1

    lResult = RegSetValueEx( _
                    hKey, _
                    sValueName, _
                    0, _
                    REG_SZ, _
                    ByVal sValue & vbNullChar, _
                    lDataSize)

    WriteRegistryString = (lResult = ERROR_SUCCESS) ' VB6 boolean success

    On Error GoTo 0
    Exit Function

WriteRegistryString_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteRegistryString of Module Module2"

End Function

'---------------------------------------------------------------------------------------
' Procedure : regCreateKeyWriteStringClose
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : write to the registry
'---------------------------------------------------------------------------------------
'
Public Function regCreateKeyWriteStringClose(ByVal hKey As Long, ByVal strPath As String, ByVal sValueName As String, ByVal sValue As String) As Boolean
    ' hKey required As LongPtr
 
    Dim keyhand As Long: keyhand = 0 ' keyhand required As LongPtr
    Dim unusedReturnValue As Long
    Dim lResult As Long
    Dim lDataSize As Long
    
    On Error GoTo regCreateKeyWriteStringClose_Error

    unusedReturnValue = RegCreateKey(hKey, strPath, keyhand)  ' hKey, keyhand required As LongPtr
'    lresult = RegSetValueEx(keyhand, sValueName, 0, REG_SZ, ByVal sValue, Len(sValue)) ' keyhand required as longPtr
    
    'REG_SZ requires the terminating NULL to be included.
    lDataSize = Len(sValue) + 1

    lResult = RegSetValueEx( _
                    keyhand, _
                    sValueName, _
                    0, _
                    REG_SZ, _
                    ByVal sValue & vbNullChar, _
                    lDataSize)

    regCreateKeyWriteStringClose = (lResult = ERROR_SUCCESS) ' VB6 boolean success
    unusedReturnValue = RegCloseKey(keyhand) ' keyhand required As LongPtr

   On Error GoTo 0
   Exit Function

regCreateKeyWriteStringClose_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure regCreateKeyWriteStringClose of module module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : regOpenKeyGetString
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
' hKey required As LongPtr
Public Function regOpenKeyGetString(ByVal hKey As Long, ByVal strPath As String, ByVal strvalue As String) As String

    Dim keyhand As Long: keyhand = 0 ' required As LongPtr
    Dim lResult As Long: lResult = 0
    Dim strBuf As String: strBuf = vbNullString
    Dim lDataBufSize As Long: lDataBufSize = 0
    Dim intZeroPos As Integer: intZeroPos = 0
    Dim lValueType As Long: lValueType = 0

    On Error GoTo regOpenKeyGetString_Error

    ' hKey required As LongPtr, keyhand  required As LongPtr
    'lResult = RegOpenKey(hKey, strPath, keyhand)
    
    ' hKey required As LongPtr, keyhand  required As LongPtr
    'Open the parent key using the newer RegOpenKeyEx with more control
    lResult = RegOpenKeyEx( _
                    hKey, _
                    strPath, _
                    0, _
                    KEY_READ Or KEY_WRITE, _
                    keyhand)
                    
    If lResult <> ERROR_SUCCESS Then
        Debug.Print "Unable to open key. " & strPath & " Error: "; lResult
        Exit Function
    End If
    
    'First call obtains the required buffer size.
    ' keyhand required As LongPtr, 3rd value 0& lpReserved 0& As LongPtr
    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
        
    If lResult <> ERROR_SUCCESS Then Exit Function

    If lValueType <> REG_SZ Then Exit Function
    
    If lValueType = REG_SZ Then
        'Registry size is in bytes. For the ANSI API, one byte per character.
        strBuf = String$(lDataBufSize, vbNullChar)
        ' keyhand  required As LongPtr, 3rd value 0& phkResult required As LongPtr
        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
            
               'Remove the terminating NULL.
                regOpenKeyGetString = Left$(strBuf, intZeroPos - 1)
            Else
                regOpenKeyGetString = strBuf
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

regOpenKeyGetString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure regOpenKeyGetString of Module Common"
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
    
    '  1 hwnd required As LongPtr, 3 ByVal hToken As LongPtr
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



'---------------------------------------------------------------------------------------
' Procedure : ReplaceMatchingDeviceIds
' Author    : chatGPT/beededea
' Date      : 22/09/2026
' Purpose   : Within the registry location there are sub keys located of the form : {12E984BD-5803-4D78-9EFB-BED7B9212C26}
'             For each sub key found, there is a DeviceId value that needs to be extracted. That contains a 16 character GDID value.
'             Extract the DeviceId, match it with our known original string GDID and see which match. Then we replace that GDID value
'             with a temporary 16 character string that we generate elsewhere. Then we write back each maching subkey to the registry.
'
'---------------------------------------------------------------------------------------
'
Public Sub ReplaceMatchingDeviceIds(ByVal OriginalGDID As String, ByVal TemporaryGDID As String)

    Dim nowValue As Date
        
    Const TOKEN_KEY As String = _
        "SOFTWARE\Microsoft\IdentityCRL\Immersive\production\Token"

    Dim hTokenKey As Long
    Dim hSubKey As Long
    Dim lResult As Long
    Dim lIndex As Long
    Dim sSubKeyName As String
    Dim sDeviceId As String
    Dim lNameLength As Long

    On Error GoTo ReplaceMatchingDeviceIds_Error

    'Open the parent Token key.
    lResult = RegOpenKeyEx( _
                    HKEY_CURRENT_USER, _
                    TOKEN_KEY, _
                    0, _
                    KEY_READ Or KEY_WRITE, _
                    hTokenKey)

    If lResult <> ERROR_SUCCESS Then
        Debug.Print "Unable to open Token key. Error: "; lResult
        Exit Sub
    End If

    lIndex = 0

    Do
        'Registry subkey names can be up to 255 characters.
        sSubKeyName = String$(256, vbNullChar)
        lNameLength = 255
        lResult = RegEnumKeyEx( _
                        hTokenKey, _
                        lIndex, _
                        sSubKeyName, _
                        lNameLength, _
                        0, _
                        vbNullString, _
                        0, _
                        ByVal 0&)

        If lResult = ERROR_NO_MORE_ITEMS Then
            Exit Do
        End If

        If lResult = ERROR_SUCCESS Then
            sSubKeyName = Left$(sSubKeyName, lNameLength)
            Debug.Print "Checking: "; sSubKeyName

            'Open this particular GUID subkey.
            hSubKey = 0
            lResult = RegOpenKeyEx( _
                            hTokenKey, _
                            sSubKeyName, _
                            0, _
                            KEY_QUERY_VALUE Or KEY_SET_VALUE, _
                            hSubKey)

            If lResult = ERROR_SUCCESS Then
                'Read DeviceId.
                sDeviceId = vbNullString

                If ReadRegistryString(hSubKey, "DeviceId", sDeviceId) Then
                    Debug.Print "    DeviceId: "; sDeviceId

                    'Compare with our known GDID.
                    If StrComp(sDeviceId, OriginalGDID, vbBinaryCompare) = 0 Then
                        nowValue = Now()
                        Debug.Print "    *** MATCH ***"

                        'Replace with temporary GDID.
                        If WriteRegistryString( _
                                    hSubKey, _
                                    "DeviceId", _
                                    TemporaryGDID) Then

                            Debug.Print "    DeviceId replaced with: "; TemporaryGDID
                            Call writeLogFile("DeviceId replaced with - " & TemporaryGDID & " in Immersive\production\Token", CStr(nowValue))
                        Else
                            Debug.Print "    ERROR writing DeviceId"
                        End If
                    End If
                Else
                    Debug.Print "    DeviceId not found/readable."
                End If

                RegCloseKey hSubKey
                hSubKey = 0
            Else
                Debug.Print "    Unable to open subkey. Error: "; lResult
            End If

            lIndex = lIndex + 1
        Else
            Debug.Print "RegEnumKeyEx error: "; lResult
            Exit Do
        End If
    Loop

    RegCloseKey hTokenKey

    On Error GoTo 0
    Exit Sub

ReplaceMatchingDeviceIds_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReplaceMatchingDeviceIds of Module Module2"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : ReadRegistryString
' Author    : beededea
' Date      : 22/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function ReadRegistryString( _
    ByVal hKey As Long, _
    ByVal ValueName As String, _
    ByRef Value As String) As Boolean

    Dim lResult As Long
    Dim lType As Long
    Dim lDataSize As Long

    Dim sBuffer As String

    'First call obtains the required buffer size.
    On Error GoTo ReadRegistryString_Error

    lResult = RegQueryValueEx( _
                    hKey, _
                    ValueName, _
                    0, _
                    lType, _
                    ByVal 0&, _
                    lDataSize)

    If lResult <> ERROR_SUCCESS Then Exit Function

    If lType <> REG_SZ Then Exit Function

    If lDataSize <= 0 Then
        Value = vbNullString
        ReadRegistryString = True
        Exit Function
    End If

    'Registry size is in bytes. For the ANSI API, one byte per character.
    sBuffer = String$(lDataSize, vbNullChar)

    lResult = RegQueryValueEx( _
                    hKey, _
                    ValueName, _
                    0, _
                    lType, _
                    ByVal sBuffer, _
                    lDataSize)

    If lResult <> ERROR_SUCCESS Then Exit Function

    'Remove the terminating NULL.
    If lDataSize > 0 Then
        Value = Left$(sBuffer, lDataSize - 1)
    Else
        Value = vbNullString
    End If

    ReadRegistryString = True

    On Error GoTo 0
    Exit Function

ReadRegistryString_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadRegistryString of Module Module2"

End Function

