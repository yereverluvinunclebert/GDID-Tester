Attribute VB_Name = "ENCRYPT"
'---------------------------------------------------------------------------------------
' Module    : ENCRYPT
' DateTime  : 04/08/2006 14:29
' Author    : Dean
' Purpose   :
'---------------------------------------------------------------------------------------
'   ============================================================
'    ----------------------------------------------------------
'     Application Name: FireStrike Military Simulation
'                       Software solution for the tabletop Wargamer
'     Developer/Programmer: Dean Beedell
'    ----------------------------------------------------------
'     Module Name: ENCRYPT
'     Module File: Encrypt.bas
'     Module Type: Form
'     Module Description:
'    ----------------------------------------------------------
'     © Copyright LightQuick 2006
'    ----------------------------------------------------------
'   ============================================================
Option Explicit

'----------------------------------------
'Name: encryptstr
'Description:
'----------------------------------------
Public Function encryptstr(messagetext As String) As String
    Dim code1 As String
    Dim Encryptcode As String
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    Dim d As Integer
    Dim cr As String

 
    code1 = "GDUID"
    Encryptcode = "6F09E"
    a = 0: b = 0: C = 0: cr = vbNullString


    Do While a < Len(messagetext)
        a = a + 1
        b = b + 1: If b > Len(code1) Then b = 1
        C = C + 1: If C > Len(Encryptcode) Then C = 1
        d = Asc(Mid$(messagetext, a, 1)) + Asc(Mid$(code1, b, 1)) + Asc(Mid$(Encryptcode, C, 1))
Loop2:
        If d > 255 Then d = d - 255: GoTo Loop2
        cr = cr + Chr$(d)
    Loop

    encryptstr = cr                              ' set the return value

    '======================================================
    'END routine error handler
    '======================================================
    On Error GoTo 0: Exit Function

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure encryptstr  " ' call a custom error-handler subroutine


End Function

'---------------------------------------------------------------------------------------
' Procedure : decryptstr
' Author    : beededea
' Date      : 27/04/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function decryptstr(messagetext As String) As String

    Dim code1 As String
    Dim Encryptcode As String
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    Dim d As Integer
    Dim cr As String
 
    On Error GoTo decryptstr_Error

    code1 = "GDUID"
    Encryptcode = "6F09E"

    a = 0
    b = 0
    C = 0

    '========================
    ' decrypt
    '========================
    Do While a < Len(messagetext)
        a = a + 1
        b = b + 1
        If b > Len(code1) Then b = 1
        C = C + 1
        If C > Len(Encryptcode) Then C = 1
        d = Asc(Mid$(messagetext, a, 1)) - (Asc(Mid$(code1, b, 1)) + Asc(Mid$(Encryptcode, C, 1)))
Loop3:
        If d < 1 Then d = d + 255: GoTo Loop3
        cr = cr + Chr$(d)
    Loop

    decryptstr = cr                              ' set the return value


    On Error GoTo 0
    Exit Function

decryptstr_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure decryptstr of Module ENCRYPT"

End Function

