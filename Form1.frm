VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GDID Tester"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Enable at Windows Startup"
      Height          =   405
      Left            =   330
      TabIndex        =   17
      Top             =   3420
      Width           =   2535
   End
   Begin VB.CommandButton btnClear 
      Height          =   285
      Left            =   7080
      Picture         =   "Form1.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   270
      Width           =   285
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate"
      Height          =   405
      Left            =   4200
      TabIndex        =   15
      Top             =   2850
      Width           =   1515
   End
   Begin VB.CheckBox chkAutomaticRemoval 
      Caption         =   "Enable Automatic Removal"
      Height          =   405
      Left            =   330
      TabIndex        =   13
      Top             =   3030
      Width           =   2535
   End
   Begin VB.CheckBox chkRegularTesting 
      Caption         =   "Enable Regular Testing"
      Height          =   405
      Left            =   330
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CheckBox chkAlertMsgBox 
      Caption         =   "Enable Automatic Alert pop-up when found"
      Height          =   405
      Left            =   330
      TabIndex        =   10
      Top             =   2220
      Width           =   3675
   End
   Begin VB.CommandButton btnReadRegistry 
      Caption         =   "Read GDID"
      Height          =   405
      Left            =   5850
      TabIndex        =   5
      Top             =   2370
      Width           =   1515
   End
   Begin VB.CommandButton btnDismiss 
      Caption         =   "Dismiss"
      Height          =   435
      Left            =   5850
      TabIndex        =   2
      Top             =   3300
      Width           =   1515
   End
   Begin VB.Timer tmrGDIDTester 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4290
      Top             =   1830
   End
   Begin VB.CommandButton btnRemoveRegValue 
      Caption         =   "Remove"
      Height          =   405
      Left            =   5850
      TabIndex        =   1
      Top             =   2850
      Width           =   1515
   End
   Begin VB.TextBox txtRegistryValue 
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Text            =   "Registry Value"
      Top             =   240
      Width           =   2805
   End
   Begin VB.ComboBox cmbDateTime 
      Height          =   315
      ItemData        =   "Form1.frx":12F7
      Left            =   5160
      List            =   "Form1.frx":12FE
      TabIndex        =   14
      Text            =   "none found"
      Top             =   270
      Width           =   1845
   End
   Begin VB.Label Label3 
      Caption         =   "Seconds until next test -"
      Height          =   435
      Left            =   300
      TabIndex        =   11
      Top             =   1890
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "Date/Time"
      Height          =   435
      Left            =   4200
      TabIndex        =   9
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label lblGDIDLink 
      Caption         =   "GDID Information"
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   4320
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      ToolTipText     =   "Double click here to view a site describing the GDID tracking issue."
      Top             =   3480
      Width           =   1545
   End
   Begin VB.Label lblCountdown 
      Caption         =   "Disabled"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1890
      Width           =   1935
   End
   Begin VB.Label lblCheckValue 
      Caption         =   $"Form1.frx":130E
      Height          =   645
      Left            =   300
      TabIndex        =   6
      Top             =   1230
      Width           =   6645
   End
   Begin VB.Label lblKey 
      Caption         =   "Key Value"
      Height          =   435
      Left            =   300
      TabIndex        =   4
      Top             =   300
      Width           =   1035
   End
   Begin VB.Label txtGDID 
      Caption         =   "HKEY_CURRENT_USER, ""SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties"", ""lid"""
      Height          =   465
      Left            =   300
      TabIndex        =   3
      Top             =   840
      Width           =   6705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\IdentityCRL\ExtendedProperties", "lid", "0018AAAABBBBCCCC")
        Call readRegistryValue
        
    On Error GoTo 0
    Exit Sub

btnGenerate_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGenerate_Click of Form Form1"
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
    
    ' textDateTime.Text = ""
    
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
' Procedure : chkAutomaticRemoval_Click
' Author    : beededea
' Date      : 05/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkAutomaticRemoval_Click()

    On Error GoTo chkAutomaticRemoval_Click_Error

    If chkAutomaticRemoval.Value = 1 Then chkRegularTesting.Value = 1

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

    tmrGDIDTester.Enabled = chkRegularTesting.Value
    
    If chkRegularTesting.Value = 0 Then chkAutomaticRemoval.Value = 0

    On Error GoTo 0
    Exit Sub

chkRegularTesting_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkRegularTesting_Click of Form Form1"
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

    Call readRegistryValue

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Form1"
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
' Procedure : lblGDIDLink_DblClick
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGDIDLink_DblClick()

    On Error GoTo lblGDIDLink_DblClick_Error

        Call ShellExecute(Form1.hwnd, "Open", "https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-i", vbNullString, vbNullString, 1)

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
        
        nowValue = Now()
        
        If cmbDateTime.Text = "none found" Then
            cmbDateTime.RemoveItem 0
            cmbDateTime.AddItem CStr(nowValue), 0
        Else
            cmbDateTime.AddItem CStr(nowValue)
        End If
        
        cmbDateTime.Text = CStr(nowValue)
        
        If chkAlertMsgBox.Value = 1 Then MsgBox "GDID has changed"
        
    End If

    On Error GoTo 0
    Exit Sub

testGDID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testGDID of Form Form1"
End Sub
