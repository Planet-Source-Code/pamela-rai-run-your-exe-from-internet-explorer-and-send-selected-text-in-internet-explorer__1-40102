VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Click here to VOTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3015
      TabIndex        =   1
      Top             =   750
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetDesktopWindow Lib "USER32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31
Private Sub Form_Load()
Text1.Text = Command
End Sub
Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
   hWndDesk = GetDesktopWindow()
     success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
If success = SE_ERR_NOASSOC Then
    MsgBox "Couldn't load the default application"
 
    Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF0000
End Sub

Private Sub Label1_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=40102&lngWId=1"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF00&
End Sub
