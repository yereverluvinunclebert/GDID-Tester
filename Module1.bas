Attribute VB_Name = "Module1"
'---------------------------------------------------------------------------------------
' Module    : Module1
' Author    : beededea
' Date      : 08/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function RtlGenRandom Lib "advapi32.dll" Alias "SystemFunction036" ( _
    ByRef Buffer As Any, _
    ByVal Length As Long) As Long


' General member property variables declared

Private m_sgsWindowsStartup As String
Private m_sgsAlertMsgBox As String
Private m_sgsRegularTesting As String
Private m_sgsAutomaticRemoval As String




'---------------------------------------------------------------------------------------
' Procedure : gsWindowsStartup
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWindowsStartup() As String

    On Error GoTo gsWindowsStartup_Error

    gsWindowsStartup = m_sgsWindowsStartup

    On Error GoTo 0
    Exit Property

gsWindowsStartup_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWindowsStartup of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWindowsStartup
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWindowsStartup(ByVal sgsWindowsStartup As String)

    On Error GoTo gsWindowsStartup_Error

    m_sgsWindowsStartup = sgsWindowsStartup

    On Error GoTo 0
    Exit Property

gsWindowsStartup_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWindowsStartup of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsAlertMsgBox
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsAlertMsgBox() As String

    On Error GoTo gsAlertMsgBox_Error

    gsAlertMsgBox = m_sgsAlertMsgBox

    On Error GoTo 0
    Exit Property

gsAlertMsgBox_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsAlertMsgBox of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsAlertMsgBox
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsAlertMsgBox(ByVal sgsAlertMsgBox As String)

    On Error GoTo gsAlertMsgBox_Error

    m_sgsAlertMsgBox = sgsAlertMsgBox

    On Error GoTo 0
    Exit Property

gsAlertMsgBox_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsAlertMsgBox of Module Module1"

End Property




'---------------------------------------------------------------------------------------
' Procedure : gsAutomaticRemoval
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsAutomaticRemoval() As String

    On Error GoTo gsAutomaticRemoval_Error

    gsAutomaticRemoval = m_sgsAutomaticRemoval

    On Error GoTo 0
    Exit Property

gsAutomaticRemoval_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsAutomaticRemoval of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsAutomaticRemoval
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsAutomaticRemoval(ByVal sgsAutomaticRemoval As String)

    On Error GoTo gsAutomaticRemoval_Error

    m_sgsAutomaticRemoval = sgsAutomaticRemoval

    On Error GoTo 0
    Exit Property

gsAutomaticRemoval_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsAutomaticRemoval of Module Module1"

End Property





'---------------------------------------------------------------------------------------
' Procedure : gsRegularTesting
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsRegularTesting() As String

    On Error GoTo gsRegularTesting_Error

    gsRegularTesting = m_sgsRegularTesting

    On Error GoTo 0
    Exit Property

gsRegularTesting_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsRegularTesting of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsRegularTesting
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsRegularTesting(ByVal sgsRegularTesting As String)

    On Error GoTo gsRegularTesting_Error

    m_sgsRegularTesting = sgsRegularTesting

    On Error GoTo 0
    Exit Property

gsRegularTesting_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsRegularTesting of Module Module1"

End Property



'---------------------------------------------------------------------------------------
' Procedure : SecureRandomHex64
' Author    : beededea
' Date      : 09/08/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SecureRandomHex64() As String

    Dim b(7) As Byte
    Dim i As Long

    On Error GoTo SecureRandomHex64_Error

    If RtlGenRandom(b(0), 8) <> 0 Then
        For i = 0 To 7
            SecureRandomHex64 = SecureRandomHex64 & Right$("0" & Hex$(b(i)), 2)
        Next
    End If

    On Error GoTo 0
    Exit Function

SecureRandomHex64_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SecureRandomHex64 of Module Module1"

End Function
