VERSION 5.00
Begin VB.Form messageBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GDIDTester Message Box"
   ClientHeight    =   1890
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4500
      TabIndex        =   0
      Top             =   1410
      Width           =   1215
   End
   Begin VB.Frame fraBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   -30
      TabIndex        =   1
      Top             =   -90
      Width           =   5805
      Begin VB.TextBox txtMessageBox 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   150
         TabIndex        =   2
         Text            =   "txtMessageBox"
         Top             =   540
         Width           =   5745
      End
   End
End
Attribute VB_Name = "messageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_sgsMsgboxText As String

'---------------------------------------------------------------------------------------
' Procedure : gsMsgboxText
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsMsgboxText() As String

    On Error GoTo gsMsgboxText_Error

    gsMsgboxText = m_sgsMsgboxText

    On Error GoTo 0
    Exit Property

gsMsgboxText_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsMsgboxText of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsMsgboxText
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsMsgboxText(ByVal sgsMsgboxText As String)

    On Error GoTo gsMsgboxText_Error

    m_sgsMsgboxText = sgsMsgboxText
    
    txtMessageBox.Text = m_sgsMsgboxText
    messageBox.Show

    On Error GoTo 0
    Exit Property

gsMsgboxText_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsMsgboxText of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 25/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Error

    'gsMsgboxText = "Empty Caption"

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form messageBox"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : OKButton_Click
' Author    : beededea
' Date      : 25/09/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub OKButton_Click()

    On Error GoTo OKButton_Click_Error

    Me.Hide

    On Error GoTo 0
    Exit Sub

OKButton_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OKButton_Click of Form messageBox"
End Sub
