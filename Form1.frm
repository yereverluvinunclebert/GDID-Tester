VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GDID Tester"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOriginalGDID 
      Height          =   345
      Left            =   2340
      TabIndex        =   20
      Text            =   "Original GDID"
      Top             =   180
      Width           =   2295
   End
   Begin VB.CheckBox chkAutomaticGeneration 
      Caption         =   "Enable Automatic Replacement of GDID"
      Height          =   405
      Left            =   330
      TabIndex        =   18
      Top             =   3840
      Width           =   3285
   End
   Begin VB.CheckBox chkWindowsStartup 
      Caption         =   "Enable at Windows Startup"
      Height          =   255
      Left            =   330
      TabIndex        =   17
      Top             =   4290
      Width           =   2535
   End
   Begin VB.CommandButton btnClear 
      Height          =   285
      Left            =   7080
      Picture         =   "Form1.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   285
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate"
      Height          =   405
      Left            =   5850
      TabIndex        =   15
      Top             =   3270
      Width           =   1515
   End
   Begin VB.CheckBox chkAutomaticRemoval 
      Caption         =   "Enable Automatic Removal"
      Height          =   405
      Left            =   330
      TabIndex        =   13
      Top             =   3450
      Width           =   2535
   End
   Begin VB.CheckBox chkRegularTesting 
      Caption         =   "Enable Regular Testing"
      Height          =   405
      Left            =   330
      TabIndex        =   12
      Top             =   3060
      Width           =   2535
   End
   Begin VB.CheckBox chkAlertMsgBox 
      Caption         =   "Enable Automatic Alert pop-up when found"
      Height          =   405
      Left            =   330
      TabIndex        =   10
      Top             =   2640
      Width           =   3675
   End
   Begin VB.CommandButton btnReadRegistry 
      Caption         =   "Read GDID"
      Height          =   405
      Left            =   5850
      TabIndex        =   5
      Top             =   2310
      Width           =   1515
   End
   Begin VB.CommandButton btnDismiss 
      Caption         =   "Dismiss"
      Height          =   435
      Left            =   5850
      TabIndex        =   2
      Top             =   4140
      Width           =   1515
   End
   Begin VB.Timer tmrGDIDTester 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   2340
   End
   Begin VB.CommandButton btnRemoveRegValue 
      Caption         =   "Remove"
      Height          =   405
      Left            =   5850
      TabIndex        =   1
      Top             =   2790
      Width           =   1515
   End
   Begin VB.TextBox txtRegistryValue 
      Height          =   345
      Left            =   2340
      TabIndex        =   0
      Text            =   "Current Registry Value"
      Top             =   690
      Width           =   2295
   End
   Begin VB.ComboBox cmbDateTime 
      Height          =   315
      ItemData        =   "Form1.frx":12F7
      Left            =   5160
      List            =   "Form1.frx":12FE
      TabIndex        =   14
      Text            =   "none found"
      Top             =   720
      Width           =   1845
   End
   Begin VB.Label Label1 
      Caption         =   "Original Key Value (GDID)"
      Height          =   435
      Left            =   300
      TabIndex        =   19
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "Seconds until next test -"
      Height          =   435
      Left            =   300
      TabIndex        =   11
      Top             =   2310
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   780
      Width           =   975
   End
   Begin VB.Label lblGDIDLink 
      Caption         =   "GDID Information"
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   5220
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   240
      Width           =   1545
   End
   Begin VB.Label lblCountdown 
      Caption         =   "Disabled"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2310
      Width           =   1935
   End
   Begin VB.Label lblCheckValue 
      Caption         =   $"Form1.frx":130E
      Height          =   645
      Left            =   300
      TabIndex        =   6
      Top             =   1650
      Width           =   6645
   End
   Begin VB.Label lblKey 
      Caption         =   "Current Key Value"
      Height          =   435
      Left            =   300
      TabIndex        =   4
      Top             =   750
      Width           =   1305
   End
   Begin VB.Label txtGDID 
      Caption         =   "HKEY_CURRENT_USER, ""SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties"", ""lid"""
      Height          =   465
      Left            =   300
      TabIndex        =   3
      Top             =   1260
      Width           =   6705
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

    On Error GoTo btnDismiss_Click_Error
    
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
    
    On Error GoTo generateGDID_Error

    newGDID = SecureRandomHex64
    
    ' store the generated GDID
    If auto = True Then gsAutoGeneratedGDID = newGDID

    Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid", newGDID)
    Call readRegistryValue
        
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

    Call testGDID
    If GDID = "" Then
        MsgBox "No GDID found in the registry"
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
    On Error GoTo btnRemoveRegValue_Click_Error
    
    Call readRegistryValue
    
    If txtRegistryValue.Text = "" Then
        MsgBox "No GDID found in the registry"
    Else
        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid", "")
        Call readRegistryValue
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
        chkAutomaticRemoval.Value = 0
        chkRegularTesting.Value = 1
    End If

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

    On Error GoTo chkWindowsStartup_Click_Error

    gsWindowsStartup = CStr(chkWindowsStartup.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "WindowsStartup", gsWindowsStartup, gsSettingsFile
    End If
    
    If gsWindowsStartup = "1" Then
        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "GDIDTester", """" & App.Path & "\" & "GDIDTester" & ".exe""")
    Else
        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "GDIDTester", vbNullString)
    End If
    
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
    
    On Error GoTo 0
    Exit Sub

chkAlertMsgBox_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAlertMsgBox_Click of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkAutomaticRemoval_Click
' Author    : beededea
' Date      : 05/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutomaticRemoval_Click()

    On Error GoTo chkAutomaticRemoval_Click_Error
    
    gsAutomaticRemoval = CStr(chkAutomaticRemoval.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "AutomaticRemoval", gsAutomaticRemoval, gsSettingsFile
    End If

    If chkAutomaticRemoval.Value = 1 Then chkRegularTesting.Value = 1
    If chkAutomaticRemoval.Value = 1 Then chkAutomaticGeneration.Value = 0
    
    On Error GoTo 0
    Exit Sub

chkAutomaticRemoval_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAutomaticRemoval_Click of Form Form1"
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
    
    gsRegularTesting = CStr(chkRegularTesting.Value)
    
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\GDIDTester", "RegularTesting", gsRegularTesting, gsSettingsFile
    End If
    
    tmrGDIDTester.Enabled = chkRegularTesting.Value
    
    If chkRegularTesting.Value = 0 Then
        chkAutomaticRemoval.Value = 0
        chkAutomaticGeneration.Value = 0
    End If

    On Error GoTo 0
    Exit Sub

chkRegularTesting_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkRegularTesting_Click of Form Form1"
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
    
    ' validate and set any missing inputs
    Call validateInputs
    
    ' adjust all the preference controls
    Call adjustControls
    
    ' set the tooltips
    Call setTooltips

    ' read the GDID registry values
    Call readRegistryValue
    
    ' check the first time run status
    Call setFirstRunStatus

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Form1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkPreviousInstance
' Author    : beededea
' Date      : 10/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub checkPreviousInstance()

    Dim ans As VbMsgBoxResult
    
    On Error GoTo checkPreviousInstance_Error

    If App.PrevInstance = True Then
        ans = MsgBox("Previous instance aready running, please close the other instance and try again.")
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
    chkRegularTesting.Value = CInt(gsRegularTesting)
    chkAutomaticRemoval.Value = CInt(gsAutomaticRemoval)
    chkAutomaticGeneration.Value = CInt(gsAutomaticGeneration)
    
    txtOriginalGDID.Text = gsOriginalGDID

    On Error GoTo 0
    Exit Sub

adjustControls_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustControls of Form Form1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readRegistryValue
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub readRegistryValue()
    
    On Error GoTo readRegistryValue_Error

    GDID = getstring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid")
    
    txtRegistryValue.Text = GDID

    On Error GoTo 0
    Exit Sub

readRegistryValue_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readRegistryValue of Form Form1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setFirstRunStatus
' Author    : beededea
' Date      : 09/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setFirstRunStatus()

    On Error GoTo setFirstRunStatus_Error

    If gbFirstTimeRun = True Then

        gsOriginalGDID = GDID
        gbFirstTimeRun = False
        
        ' save the first time run state AND the original GDID
        If fFExists(gsSettingsFile) Then
            sPutINISetting "Software\GDIDTester", "FirstTimeRun", gbFirstTimeRun, gsSettingsFile
            sPutINISetting "Software\GDIDTester", "OriginalGDID", gsOriginalGDID, gsSettingsFile
        End If
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
' Procedure : tmrGDIDTester_tmrGDIDTester
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrGDIDTester_Timer()

    Static timerValue As Long

    On Error GoTo tmrGDIDTester_tmrGDIDTester_Error

    timerValue = timerValue + 1
    
    lblCountdown.Caption = CStr(10 - timerValue)
    
    If timerValue >= 10 Then
        lblCountdown.Caption = "10 - Testing"
        timerValue = 0
        Call testGDID
    End If

    On Error GoTo 0
    Exit Sub

tmrGDIDTester_tmrGDIDTester_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrGDIDTester_tmrGDIDTester of Form Form1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : testGDID
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub testGDID()

    Dim oldRegValue As String: oldRegValue = vbNullString
    Dim nowValue As Date

    On Error GoTo testGDID_Error

    oldRegValue = GDID
    
    Call readRegistryValue
    
    If oldRegValue <> "" Then
        
        If chkAutomaticRemoval.Value = 1 Then
            Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid", "")
            Call readRegistryValue
        End If
        
        If chkAutomaticGeneration.Value = 1 Then
            
            ' if the GDID is the same as the stored generated GDID then there is no need to generate new
            If gsAutoGeneratedGDID <> txtRegistryValue.Text Then
                Call generateGDID(True)
            End If
        Else
            nowValue = Now()
            
            If cmbDateTime.Text = "none found" Then
                cmbDateTime.RemoveItem 0
                cmbDateTime.AddItem CStr(nowValue), 0
            Else
                cmbDateTime.AddItem CStr(nowValue)
            End If
            
            cmbDateTime.Text = CStr(nowValue)
                    
        End If
        
        If chkAlertMsgBox.Value = 1 Then MsgBox "GDID has changed"
        
    End If

    On Error GoTo 0
    Exit Sub

testGDID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testGDID of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal Location As String, ByVal gsSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(gsSettingsFile) Then
        
        ' general
        gbFirstTimeRun = fGetINISetting(Location, "FirstTimeRun", gsSettingsFile)
        gsWindowsStartup = fGetINISetting(Location, "WindowsStartup", gsSettingsFile)
        gsOriginalGDID = fGetINISetting(Location, "OriginalGDID", gsSettingsFile)
        
        ' configuration
        gsAlertMsgBox = fGetINISetting(Location, "AlertMsgBox", gsSettingsFile)
        gsRegularTesting = fGetINISetting(Location, "RegularTesting", gsSettingsFile)
        gsAutomaticRemoval = fGetINISetting(Location, "AutomaticRemoval", gsSettingsFile)
        gsAutomaticGeneration = fGetINISetting(Location, "AutomaticGeneration", gsSettingsFile)

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

    txtRegistryValue.ToolTipText = "This field shows any GDID that is inserted into the registry key."
    cmbDateTime.ToolTipText = "This drop down list shows a new date and time recording when the GDID was last changed."
    chkAlertMsgBox.ToolTipText = "This check box will enable a pop-up message box to appear in the centre of the screen when a GDID is found."
    chkRegularTesting.ToolTipText = "This check box will enable the regular testing timer."
    btnClear.ToolTipText = "This button will clear any stored dates/times."
    chkAutomaticRemoval.ToolTipText = "This check box when enabled will remove any remote-generated GDID."
    chkAutomaticGeneration.ToolTipText = "This check box when enabled will populate the registry key with a fake GDID"
    chkWindowsStartup.ToolTipText = "Checking this box will cause this utility to restart on each windows startup."
    btnGenerate.ToolTipText = "This button generates a new unique 64bit GDID, entirely random."
    btnReadRegistry.ToolTipText = "This button will read the registry key when you want to check manually."
    btnRemoveRegValue.ToolTipText = "This button will remove the registry key when you want to do so manually."
    btnDismiss.ToolTipText = "Click on me to close the utility."
    lblGDIDLink.ToolTipText = "Double click here to view a site describing the GDID tracking issue."
    txtOriginalGDID.ToolTipText = "This is the original GDID that was stored within the registry key."

    On Error GoTo 0
    Exit Sub

setTooltips_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setTooltips of Form Form1"
End Sub
