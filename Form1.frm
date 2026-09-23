VERSION 5.00
Object = "{13E244CC-5B1A-45EA-A5BC-D3906B9ABB79}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GDID Tester"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTicker 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4770
      Top             =   3570
   End
   Begin CCRSlider.Slider sliGDIDInterval 
      Height          =   405
      Left            =   180
      TabIndex        =   21
      Top             =   2250
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   714
      LargeChange     =   1
   End
   Begin VB.CommandButton btnTicks 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   5820
      Picture         =   "Form1.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2250
      Width           =   405
   End
   Begin VB.CommandButton btnLogfile 
      Caption         =   "View log"
      Height          =   435
      Left            =   6210
      TabIndex        =   19
      Top             =   4230
      Width           =   1515
   End
   Begin VB.TextBox txtOriginalGDID 
      Height          =   345
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Original GDID"
      Top             =   180
      Width           =   2295
   End
   Begin VB.CheckBox chkAutomaticGeneration 
      Caption         =   "Enable Automatic Replacement of GDID"
      Height          =   405
      Left            =   330
      TabIndex        =   16
      Top             =   4410
      Width           =   3285
   End
   Begin VB.CheckBox chkWindowsStartup 
      Caption         =   "Enable at Windows Startup"
      Height          =   255
      Left            =   330
      TabIndex        =   15
      Top             =   4860
      Width           =   2535
   End
   Begin VB.CommandButton btnClear 
      Height          =   285
      Left            =   7440
      Picture         =   "Form1.frx":1543
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   720
      Width           =   285
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate"
      Height          =   405
      Left            =   6210
      TabIndex        =   13
      Top             =   3780
      Width           =   1515
   End
   Begin VB.CheckBox chkAutomaticBlanking 
      Caption         =   "Enable Automatic Removal (blanking)"
      Height          =   405
      Left            =   330
      TabIndex        =   11
      Top             =   4020
      Width           =   3345
   End
   Begin VB.CheckBox chkRegularTesting 
      Caption         =   "Enable Regular Testing"
      Height          =   405
      Left            =   330
      TabIndex        =   10
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CheckBox chkAlertMsgBox 
      Caption         =   "Enable Automatic Alert pop-up when found"
      Height          =   405
      Left            =   330
      TabIndex        =   8
      Top             =   3630
      Width           =   3675
   End
   Begin VB.CommandButton btnReadRegistry 
      Caption         =   "Read GDID"
      Height          =   405
      Left            =   6210
      TabIndex        =   4
      Top             =   2820
      Width           =   1515
   End
   Begin VB.CommandButton btnDismiss 
      Caption         =   "Dismiss"
      Height          =   435
      Left            =   6210
      TabIndex        =   2
      Top             =   4710
      Width           =   1515
   End
   Begin VB.Timer tmrGDIDTester 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   3570
   End
   Begin VB.CommandButton btnRemoveRegValue 
      Caption         =   "Remove"
      Height          =   405
      Left            =   6210
      TabIndex        =   1
      Top             =   3300
      Width           =   1515
   End
   Begin VB.TextBox txtRegistryValue 
      Height          =   345
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Current Registry Value"
      Top             =   690
      Width           =   2295
   End
   Begin VB.ComboBox cmbDateTime 
      Height          =   315
      ItemData        =   "Form1.frx":1770
      Left            =   5130
      List            =   "Form1.frx":1777
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "none found"
      Top             =   720
      Width           =   2235
   End
   Begin VB.Label Label4 
      Caption         =   "10 secs"
      Height          =   255
      Left            =   4410
      TabIndex        =   23
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label lblOneSecond 
      Caption         =   "0 sec"
      Height          =   255
      Left            =   330
      TabIndex        =   22
      Top             =   2880
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Original Key Value (GDID)"
      Height          =   435
      Left            =   300
      TabIndex        =   17
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label lblMilliseconds 
      Caption         =   "Timer interval "
      Height          =   435
      Left            =   1350
      TabIndex        =   9
      Top             =   2880
      Width           =   2625
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   780
      Width           =   975
   End
   Begin VB.Label lblGDIDLink 
      Caption         =   "GDID Information"
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   5220
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   240
      Width           =   1545
   End
   Begin VB.Label lblCheckValue 
      Caption         =   $"Form1.frx":1787
      Height          =   645
      Left            =   300
      TabIndex        =   5
      Top             =   1260
      Width           =   6645
   End
   Begin VB.Label lblKey 
      Caption         =   "Current Key Value"
      Height          =   435
      Left            =   300
      TabIndex        =   3
      Top             =   750
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' Author    : beededea
' Date      : 08/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private GDID As String


'---------------------------------------------------------------------------------------
' Procedure : btnLogfile_Click
' Author    : beededea
' Date      : 11/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnLogfile_Click()

    On Error GoTo btnLogfile_Click_Error

    Call ShellExecute(Form1.hWnd, "Open", gsSettingsDir & "\GDIDChangeLog.log", vbNullString, vbNullString, 1)
    
    On Error GoTo 0
    Exit Sub

btnLogfile_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnLogfile_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTicks_Click
' Author    : beededea
' Date      : 11/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnTicks_Click()

    On Error GoTo btnTicks_Click_Error
    
    btnTicks.Tag = "Hidden"
    btnTicks.Visible = False

    On Error GoTo 0
    Exit Sub

btnTicks_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnTicks_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Initialize
' Author    : beededea
' Date      : 08/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Initialize()

    On Error GoTo Form_Initialize_Error

    ' general storage variables declared
    gsSettingsDir = vbNullString
    gsSettingsFile = vbNullString
    gbFirstTimeRun = False
    gbStartupFlg = True
    
    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Error
    
    ' prevent two instances running simultaneously
    Call checkPreviousInstance
    
    ' get the location of the tool settings file
    Call getToolSettingsFile
        
    ' read the program settings from the configuration settings file
    Call readSettingsFile("Software\GDIDTester", gsSettingsFile)
    
    ' open the log and touch it.
    Call writeLogFile(" ")
    Call writeLogFile("Starting GDIDTester ", Now())

    ' validate and set any missing inputs
    Call validateInputs
    
    ' adjust all the controls
    Call adjustControls
    
    ' set the tooltips
    Call setTooltips
    
    ' check the first time run status
    Call setFirstRunStatus
    
    ' the main timer is started in chkRegularTesting_Click() in adjustControls
    
    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnClear_Click
' Author    : beededea
' Date      : 07/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClear_Click()

    On Error GoTo btnClear_Click_Error

    cmbDateTime.Clear

    On Error GoTo 0
    Exit Sub

btnClear_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClear_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDismiss_Click
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDismiss_Click()

    Dim nowValue As Date
    
    On Error GoTo btnDismiss_Click_Error
    
    tmrGDIDTester.Enabled = False
    
    nowValue = Now()
    
    Call writeLogFile("GDIDTester shutting down with GDID " & GDID, CStr(nowValue))
    
    Unload Form1

    On Error GoTo 0
    Exit Sub

btnDismiss_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDismiss_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnGenerate_Click
' Author    : beededea
' Date      : 07/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGenerate_Click()

    On Error GoTo btnGenerate_Click_Error
    
    Call writeLogFile("Generated a GDID manually ", Now())
    Call generateGDID(False)
        
    On Error GoTo 0
    Exit Sub

btnGenerate_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGenerate_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : generateGDID
' Author    : beededea
' Date      : 09/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub generateGDID(auto As Boolean)

    Dim newGDID As String
    Dim bResult As Boolean
    
    On Error GoTo generateGDID_Error

    newGDID = SecureRandomHex64
    
    ' store the generated GDID
    If auto = True Then gsAutoGeneratedGDID = newGDID

    bResult = regCreateKeyWriteStringClose(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid", newGDID)
    If bResult = False Then
        Exit Sub
    End If
    GDID = readRegistryExtendedPropertiesLid
        
    Form1.Caption = "GDID Tester " & newGDID
    
    On Error GoTo 0
    Exit Sub

generateGDID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generateGDID of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnReadRegistry_Click
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnReadRegistry_Click()

    On Error GoTo btnReadRegistry_Click_Error

    Call modifyExtendedPropertiesLid
    If GDID = "" Then
        MsgBox "No GDID found in the registry"
    Else
        MsgBox "GDID found in the registry " & GDID
    End If
    On Error GoTo 0
    Exit Sub

btnReadRegistry_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnReadRegistry_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnRemoveRegValue_Click
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnRemoveRegValue_Click()
    Dim bResult As Boolean
    
    On Error GoTo btnRemoveRegValue_Click_Error
    
    Call writeLogFile("Removed the GDID manually ", Now())
    
    GDID = readRegistryExtendedPropertiesLid
    
    If txtRegistryValue.Text = "" Then
        MsgBox "No GDID found in the registry"
    Else
        
        bResult = regCreateKeyWriteStringClose(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid", "")
        If bResult = False Then
            Exit Sub
        End If
        GDID = readRegistryExtendedPropertiesLid
    End If
    
    On Error GoTo 0
    Exit Sub

btnRemoveRegValue_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnRemoveRegValue_Click of Form Form1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkAutomaticGeneration_Click
' Author    : beededea
' Date      : 09/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutomaticGeneration_Click()

    On Error GoTo chkAutomaticGeneration_Click_Error
    
    gsAutomaticGeneration = CStr(chkAutomaticGeneration.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "AutomaticGeneration", gsAutomaticGeneration, gsSettingsFile
    End If

    If chkAutomaticGeneration.Value = 1 Then
        chkAutomaticBlanking.Value = 0
        chkRegularTesting.Value = 1
    End If

    If gbStartupFlg = False Then Call writeLogFile("Changing the Automatic Generation status manually " & gsAutomaticGeneration, Now())

    On Error GoTo 0
    Exit Sub

chkAutomaticGeneration_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAutomaticGeneration_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkWindowsStartup_Click
' Author    : beededea
' Date      : 08/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkWindowsStartup_Click()
    Dim bResult As Boolean

    On Error GoTo chkWindowsStartup_Click_Error

    gsWindowsStartup = CStr(chkWindowsStartup.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "WindowsStartup", gsWindowsStartup, gsSettingsFile
    End If
    
    If gsWindowsStartup = "1" Then
        bResult = regCreateKeyWriteStringClose(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "GDIDTester", """" & App.Path & "\" & "GDIDTester" & ".exe""")
    Else
        bResult = regCreateKeyWriteStringClose(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "GDIDTester", vbNullString)
    End If
    
    If gbStartupFlg = False Then Call writeLogFile("Changing the Windows Startup status manually " & gsWindowsStartup, Now())
    
    On Error GoTo 0
    Exit Sub

chkWindowsStartup_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWindowsStartup_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkAlertMsgBox_Click
' Author    : beededea
' Date      : 08/08/2026
' Purpose   : save the alert value
'---------------------------------------------------------------------------------------
'
Private Sub chkAlertMsgBox_Click()

    On Error GoTo chkAlertMsgBox_Click_Error
    
    gsAlertMsgBox = CStr(chkAlertMsgBox.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "AlertMsgBox", gsAlertMsgBox, gsSettingsFile
    End If
    
    If gbStartupFlg = False Then Call writeLogFile("Changing the Auto Alert status manually to " & gsAlertMsgBox, Now())
    
    On Error GoTo 0
    Exit Sub

chkAlertMsgBox_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAlertMsgBox_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkAutomaticBlanking_Click
' Author    : beededea
' Date      : 05/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutomaticBlanking_Click()

    On Error GoTo chkAutomaticBlanking_Click_Error
    
    gsAutomaticRemoval = CStr(chkAutomaticBlanking.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "AutomaticRemoval", gsAutomaticRemoval, gsSettingsFile
    End If

    If chkAutomaticBlanking.Value = 1 Then chkRegularTesting.Value = 1
    If chkAutomaticBlanking.Value = 1 Then chkAutomaticGeneration.Value = 0
    
    If gbStartupFlg = False Then Call writeLogFile("Changing the Automatic Removal status manually " & gsAutomaticRemoval, Now())
   
    On Error GoTo 0
    Exit Sub

chkAutomaticBlanking_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAutomaticBlanking_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkRegularTesting_Click
' Author    : beededea
' Date      : 05/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkRegularTesting_Click()

    On Error GoTo chkRegularTesting_Click_Error
    
    tmrGDIDTester.Enabled = chkRegularTesting.Value
    
    If gbStartupFlg = True Then Exit Sub

    gsRegularTesting = CStr(chkRegularTesting.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "RegularTesting", gsRegularTesting, gsSettingsFile
    End If
    
    If chkRegularTesting.Value = 0 Then
        chkAutomaticBlanking.Value = 0
        chkAutomaticGeneration.Value = 0
        tmrTicker.Enabled = False
        btnTicks.Visible = False
    End If

    lblMilliseconds.Caption = tmrGDIDTester.Interval & " milliseconds, running = " & tmrGDIDTester.Enabled
    
    If gbStartupFlg = False Then Call writeLogFile("Changing the Regular Testing status manually " & gsRegularTesting, Now())

    On Error GoTo 0
    Exit Sub

chkRegularTesting_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkRegularTesting_Click of Form Form1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : checkPreviousInstance
' Author    : beededea
' Date      : 10/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub checkPreviousInstance()
  
    On Error GoTo checkPreviousInstance_Error

    If App.PrevInstance = True Then
        MsgBox ("Previous instance aready running, please close the other instance and try again.")
        Unload Form1
    End If
        
    On Error GoTo 0
    Exit Sub

checkPreviousInstance_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkPreviousInstance of Form Form1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 10/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub validateInputs()

    On Error GoTo validateInputs_Error

    ' general
    If gsWindowsStartup = "" Then gsWindowsStartup = "0"
    
    ' configuration
    If gsAlertMsgBox = "" Then gsAlertMsgBox = "0"
    If gsRegularTesting = "" Then gsRegularTesting = "0"
    If gsAutomaticRemoval = "" Then gsAutomaticRemoval = "0"
    If gsAutomaticGeneration = "" Then gsAutomaticGeneration = "0"
    If gsGDIDInterval = "" Then gsGDIDInterval = "3"
    
    On Error GoTo 0
    Exit Sub

validateInputs_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : adjustControls
' Author    : beededea
' Date      : 08/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub adjustControls()

    On Error GoTo adjustControls_Error

    ' general
    chkWindowsStartup.Value = CInt(gsWindowsStartup)
    
    ' configuration
    chkAlertMsgBox.Value = CInt(gsAlertMsgBox)
    chkRegularTesting.Value = CInt(gsRegularTesting) ' implicitly starts the main timer
    chkAutomaticBlanking.Value = CInt(gsAutomaticRemoval)
    chkAutomaticGeneration.Value = CInt(gsAutomaticGeneration)
    sliGDIDInterval.Value = CInt(gsGDIDInterval)
    
    txtOriginalGDID.Text = gsOriginalGDID
    
    lblMilliseconds.Caption = tmrGDIDTester.Interval & " milliseconds, running = " & tmrGDIDTester.Enabled

    On Error GoTo 0
    Exit Sub

adjustControls_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustControls of Form Form1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readRegistryExtendedPropertiesLid
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function readRegistryExtendedPropertiesLid() As String

    Dim thisGDID As String
    
    On Error GoTo readRegistryExtendedPropertiesLid_Error

    thisGDID = regOpenKeyGetString(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid")
    
    txtRegistryValue.Text = thisGDID
    
    readRegistryExtendedPropertiesLid = thisGDID ' return

    On Error GoTo 0
    Exit Function

readRegistryExtendedPropertiesLid_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in function readRegistryExtendedPropertiesLid of Form Form1"

End Function

'---------------------------------------------------------------------------------------
' Procedure : setFirstRunStatus
' Author    : beededea
' Date      : 09/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setFirstRunStatus()

    Dim encryptedString As String: encryptedString = vbNullString

    On Error GoTo setFirstRunStatus_Error
        
    If gbFirstTimeRun = True Then
        gsOriginalGDID = GDID
        
        gbFirstTimeRun = False
        
        ' save the first time run state AND the original GDID
        If fFExists(gsSettingsFile) Then
            sPutINISetting "Software\GDIDTester", "FirstTimeRun", gbFirstTimeRun, gsSettingsFile
            
            encryptedString = encryptstr(gsOriginalGDID)
            sPutINISetting "Software\GDIDTester", "OriginalGDID", encryptedString, gsSettingsFile
            
            'sPutINISetting "Software\GDIDTester", "OriginalGDID", gsOriginalGDID, gsSettingsFile

        End If
        
        Call writeLogFile("The Original GDID found and stored on first run - " & gsOriginalGDID, Now())
    Else
    
        Call writeLogFile("The original GDID Key Value as stored " & gsOriginalGDID) ' read from the registry.

    End If

    On Error GoTo 0
    Exit Sub

setFirstRunStatus_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFirstRunStatus of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Form_Unload_Error

    Set Form1 = Nothing
    
    End

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Form1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    Dim iFileNo As Integer: iFileNo = 0

    On Error GoTo getToolSettingsFile_Error
    
    gsSettingsDir = fSpecialFolder(feUserAppData) & "\GDIDTester"
    gsSettingsFile = gsSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gsSettingsDir) Then
        MkDir gsSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gsSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gsSettingsFile For Output As #iFileNo
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblGDIDLink_DblClick
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGDIDLink_DblClick()

    On Error GoTo lblGDIDLink_DblClick_Error

        Call ShellExecute(Form1.hWnd, "Open", "https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-i", vbNullString, vbNullString, 1)

    On Error GoTo 0
    Exit Sub

lblGDIDLink_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGDIDLink_DblClick of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblGDIDLink_MouseMove
' Author    : beededea
' Date      : 05/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGDIDLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo lblGDIDLink_MouseMove_Error

    lblGDIDLink.ForeColor = &HC00000

    On Error GoTo 0
    Exit Sub

lblGDIDLink_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGDIDLink_MouseMove of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliGDIDInterval_Change
' Author    : beededea
' Date      : 24/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliGDIDInterval_Change()

    On Error GoTo sliGDIDInterval_Change_Error
    
    tmrGDIDTester.Interval = sliGDIDInterval.Value * 1000

    If gbStartupFlg = True Then Exit Sub

    gsGDIDInterval = CStr(sliGDIDInterval.Value)
    
    If sliGDIDInterval.Value <= 0 Then
        tmrGDIDTester.Enabled = False
        btnTicks.Visible = False
        tmrTicker.Enabled = False
        chkRegularTesting.Value = 0
    Else
        If gsRegularTesting = "1" Then
            chkRegularTesting.Value = 1
            tmrGDIDTester.Enabled = True
        End If
    End If
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "GDIDInterval", gsGDIDInterval, gsSettingsFile
    End If
    
    lblMilliseconds.Caption = tmrGDIDTester.Interval & " milliseconds, running = " & tmrGDIDTester.Enabled
    
    If gbStartupFlg = False Then Call writeLogFile("Changed the testing interval manually " & gsGDIDInterval, Now())

    On Error GoTo 0
    Exit Sub

sliGDIDInterval_Change_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliGDIDInterval_Change of Form Form1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : tmrGDIDTester_tmrGDIDTester
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrGDIDTester_Timer()
 
    On Error GoTo tmrGDIDTester_tmrGDIDTester_Error
    
    tmrTicker.Enabled = True
    
    If fGDIDStatusChanged = True Then
        Call modifyExtendedPropertiesLid
        Call ReplaceMatchingDeviceIds(gsOriginalGDID, gsAutoGeneratedGDID)
    End If

    gbStartupFlg = False

    On Error GoTo 0
    Exit Sub

tmrGDIDTester_tmrGDIDTester_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrGDIDTester_tmrGDIDTester of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fGDIDStatusChanged
' Author    : beededea
' Date      : 23/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function fGDIDStatusChanged() As Boolean
    Dim nowValue As Date
    Dim bResult As Boolean

    On Error GoTo fGDIDStatusChanged_Error
    
    gbOldRegValue = GDID
    nowValue = Now()
    
    GDID = readRegistryExtendedPropertiesLid
    
    ' if it has been made blank, move on
    If GDID <> "" Then
         ' if the current auto generated GDID is the same as the stored GDID that we have to compare then there is no need to generate new
         ' the first time the old value will be blank
        If GDID = gbOldRegValue Then
            fGDIDStatusChanged = False
        Else
            If gbOldRegValue = "" Then gbOldRegValue = GDID
            fGDIDStatusChanged = True
        End If
    End If

    On Error GoTo 0
    Exit Function

fGDIDStatusChanged_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGDIDStatusChanged of Form Form1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : modifyExtendedPropertiesLid
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub modifyExtendedPropertiesLid()

    Dim nowValue As Date
    Dim bResult As Boolean

    On Error GoTo modifyExtendedPropertiesLid_Error

    nowValue = Now()
    
    If chkAutomaticBlanking.Value = 1 Then
        bResult = regCreateKeyWriteStringClose(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid", "")
        If bResult <> 0 Then
            Exit Sub
        End If
        GDID = readRegistryExtendedPropertiesLid
        Exit Sub
    End If
    
    
    If chkAutomaticGeneration.Value = 1 Then
        
            Call generateGDID(True)
           
            If gbStartupFlg = True Then
                Call writeLogFile("GDID auto-generated at start up - " & GDID, CStr(nowValue))
            Else
                Call writeLogFile("GDID auto-generated " & GDID, CStr(nowValue))
            End If
            
            Call writeCombo(nowValue)
            
            If gbStartupFlg = False Then
                If chkAlertMsgBox.Value = 1 Then MsgBox "GDID has been auto-generated"
            End If
            Call writeLogFile("Changed from - " & gbOldRegValue)

    Else
    
        Call writeCombo(nowValue)
        Call writeLogFile("GDID Changed " & GDID, CStr(nowValue))
        
        If chkAlertMsgBox.Value = 1 Then
            If gbStartupFlg = False Then MsgBox "GDID has been changed"
        End If
        Call writeLogFile("Changed from - " & gbOldRegValue)
    End If

    On Error GoTo 0
    Exit Sub

modifyExtendedPropertiesLid_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure modifyExtendedPropertiesLid of Form Form1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : testImmersiveProductionTokenKeys
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub testImmersiveProductionTokenKeys()

    Dim nowValue As Date

    On Error GoTo testImmersiveProductionTokenKeys_Error

    gbOldRegValue = GDID
    nowValue = Now()
    
    'Loop through all the keys
    'Call readRegistryImmersiveProductionTokenKeys
    
    GDID = regOpenKeyGetString(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\Immersive\production\Token", "lid")
    
'    {12E984BD-5803-4D78-9EFB-BED7B9212C26}
'
'extract the DeviceId, match it with the known original GDID and see which match.
'Change the DeviceID to match the new generated GDID
    
    If gbOldRegValue <> "" Then
        
        If chkAutomaticBlanking.Value = 1 Then
            'bResult = regCreateKeyWriteStringClose(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\Immersive\production\Token\", "lid", "")
            'Call readRegistryImmersiveProductionTokenKeys
        End If
                
    End If

    On Error GoTo 0
    Exit Sub

testImmersiveProductionTokenKeys_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testImmersiveProductionTokenKeys of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : writeCombo
' Author    : beededea
' Date      : 17/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub writeCombo(ByVal thisTime As Date)

    On Error GoTo writeCombo_Error

    If cmbDateTime.Text = "none found" Then
        cmbDateTime.RemoveItem 0
        cmbDateTime.AddItem CStr(thisTime), 0
    Else
        cmbDateTime.AddItem CStr(thisTime)
    End If

    cmbDateTime.Text = CStr(thisTime)
                        
    On Error GoTo 0
    Exit Sub

writeCombo_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeCombo of Form Form1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal Location As String, ByVal gsSettingsFile As String)
    
    Dim boolCheck As String: boolCheck = vbNullString
    Dim encryptedString As String: encryptedString = vbNullString
    
    On Error GoTo readSettingsFile_Error

    If fFExists(gsSettingsFile) Then
    
        boolCheck = fGetINISetting(Location, "FirstTimeRun", gsSettingsFile)
        If boolCheck <> "True" Then boolCheck = False
        
        ' general
        gbFirstTimeRun = CBool(boolCheck)
        gsWindowsStartup = fGetINISetting(Location, "WindowsStartup", gsSettingsFile)
        
        encryptedString = fGetINISetting(Location, "OriginalGDID", gsSettingsFile)
        gsOriginalGDID = decryptstr(encryptedString)

        ' configuration
        gsAlertMsgBox = fGetINISetting(Location, "AlertMsgBox", gsSettingsFile)
        gsRegularTesting = fGetINISetting(Location, "RegularTesting", gsSettingsFile)
        gsAutomaticRemoval = fGetINISetting(Location, "AutomaticRemoval", gsSettingsFile)
        gsAutomaticGeneration = fGetINISetting(Location, "AutomaticGeneration", gsSettingsFile)
        gsGDIDInterval = fGetINISetting(Location, "GDIDInterval", gsSettingsFile)

    End If

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setTooltips
' Author    : beededea
' Date      : 09/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setTooltips()

    On Error GoTo setTooltips_Error

    txtRegistryValue.ToolTipText = "This field shows any GDID that is inserted into the registry key, either by this tool or externally by Microsoft."
    cmbDateTime.ToolTipText = "This drop down list shows a date and time recording when the GDID was last changed since startup, for more detail view the logfile."
    chkAlertMsgBox.ToolTipText = "This check box will enable a pop-up message box to appear in the centre of the screen when a GDID is found."
    chkRegularTesting.ToolTipText = "This check box will enable the regular testing timer."
    btnClear.ToolTipText = "This button will clear any stored dates/times."
    chkAutomaticBlanking.ToolTipText = "This check box when enabled will remove any remote-generated GDID."
    chkAutomaticGeneration.ToolTipText = "This check box when enabled will replace any current registry key with a fake GDID"
    chkWindowsStartup.ToolTipText = "Checking this box will cause this utility to restart on each windows startup."
    btnGenerate.ToolTipText = "This button generates a new unique 64bit GDID, entirely random."
    btnReadRegistry.ToolTipText = "This button will read the registry key when you want to check manually."
    btnRemoveRegValue.ToolTipText = "This button will remove the registry key when you want to do so manually."
    btnDismiss.ToolTipText = "Click on me to close the utility."
    lblGDIDLink.ToolTipText = "Double click here to view a site describing the GDID tracking issue."
    txtOriginalGDID.ToolTipText = "This is the original GDID that was stored within the registry key."
    btnTicks.ToolTipText = "Click on me and this image will disappear but the timer will still run in the background."
    sliGDIDInterval.ToolTipText = "Alter the testing frequency here - 0 will stop the timer, a suitable interval is 2-3 seconds. "

    On Error GoTo 0
    Exit Sub

setTooltips_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setTooltips of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrTicker_Timer
' Author    : beededea
' Date      : 24/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrTicker_Timer()

    On Error GoTo tmrTicker_Timer_Error

      If btnTicks.Tag <> "Hidden" Then
        If btnTicks.Visible = True Then
            btnTicks.Visible = False
            tmrTicker.Enabled = False
        Else
            btnTicks.Visible = True
        End If
        
        btnTicks.Refresh
    End If
    
    On Error GoTo 0
    Exit Sub

tmrTicker_Timer_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrTicker_Timer of Form Form1"
End Sub
