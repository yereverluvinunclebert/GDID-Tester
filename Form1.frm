VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GDID Tester"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDateTime 
      Height          =   345
      Left            =   5160
      TabIndex        =   10
      Text            =   "None Found"
      Top             =   240
      Width           =   1845
   End
   Begin VB.CommandButton btnReadRegistry 
      Caption         =   "Read GDID"
      Height          =   405
      Left            =   5490
      TabIndex        =   5
      Top             =   1230
      Width           =   1515
   End
   Begin VB.CommandButton btnDismiss 
      Caption         =   "Dismiss"
      Height          =   435
      Left            =   5490
      TabIndex        =   2
      Top             =   2220
      Width           =   1515
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   4350
      Top             =   2190
   End
   Begin VB.CommandButton btnRemoveRegValue 
      Caption         =   "Remove"
      Height          =   405
      Left            =   5490
      TabIndex        =   1
      Top             =   1740
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
   Begin VB.Label Label2 
      Caption         =   "Date/Time"
      Height          =   435
      Left            =   4200
      TabIndex        =   11
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label lblGDILink 
      Caption         =   "GDID Information"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3390
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      ToolTipText     =   "Double click here to view a site describing the GDID tracking issue."
      Top             =   2400
      Width           =   1545
   End
   Begin VB.Label lblCountdown 
      Caption         =   "10 - Testing"
      Height          =   255
      Left            =   300
      TabIndex        =   8
      Top             =   2370
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "If you use Edge or visit any MS site that accesses login.live.com, then this value may be re-populated with a GDID."
      Height          =   645
      Left            =   300
      TabIndex        =   7
      ToolTipText     =   "Microsoft account, Store, OneDrive, Microsoft 365, account-linked UWP apps"
      Top             =   1590
      Width           =   4365
   End
   Begin VB.Label lblCheckValue 
      Caption         =   "Checking the above value every ten seconds"
      Height          =   645
      Left            =   300
      TabIndex        =   6
      Top             =   1230
      Width           =   3915
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
' Procedure : btnReadRegistry_Click
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnReadRegistry_Click()

    On Error GoTo btnReadRegistry_Click_Error

    If GDID = "" Then MsgBox "No GDID found in the registry"

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
' Procedure : lblGDILink_DblClick
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGDILink_DblClick()

    On Error GoTo lblGDILink_DblClick_Error

        Call ShellExecute(Form1.hwnd, "Open", "https://www.it-connect.tech/windows-gdid-impossible-to-delete-but-you-can-block-i", vbNullString, vbNullString, 1)

    On Error GoTo 0
    Exit Sub

lblGDILink_DblClick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGDILink_DblClick of Form Form1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Timer_Timer
' Author    : beededea
' Date      : 03/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Timer_Timer()

    Static timerValue As Long

    On Error GoTo Timer_Timer_Error

    timerValue = timerValue + 1
    
    lblCountdown.Caption = CStr(10 - timerValue)
    
    If timerValue >= 10 Then
        lblCountdown.Caption = "10 - Testing"
        timerValue = 0
        Call testGDID
    End If

    On Error GoTo 0
    Exit Sub

Timer_Timer_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer_Timer of Form Form1"
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

    On Error GoTo testGDID_Error

    oldRegValue = GDID
    
    Call readRegistryValue
    
    If oldRegValue <> GDID Then
        txtDateTime.Text = Now()
        MsgBox "GDID has changed"
    End If

    On Error GoTo 0
    Exit Sub

testGDID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testGDID of Form Form1"
End Sub
