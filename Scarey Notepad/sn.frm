VERSION 5.00
Begin VB.Form sn 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "sikz"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "sn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "<----- the cursor"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "sn.frx":030A
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "sn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'scarey notepad by sikz
'demonstrates the following (all pretty simple)
'very simple api wait function
'some simple cursor manipulation
'the sendkeys function
'if you know of a better way, either mail me
'sikz187@hushmail.com
'or leave a msg on the psc board

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Form_Load()
Me.Height = 1260
FixLabel
End Sub

Private Function FixLabel()                         'this function is entirely for this app only
With Label1
    .Alignment = 2                          'center
    .Caption = "scarey notepad by sikz" & vbCrLf & "do not close notepad while effect is in progress" & vbCrLf & "click with finger to start"
    .FontUnderline = True
    .ForeColor = &HFF0000                   'blue
End With
End Function

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 0                     'vbdefault
End Sub

Private Sub Label1_Click()
MakeNotepadScarey
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Screen
    .MouseIcon = Image1.Picture             'set the cursor up
    .MousePointer = 99                      'vbcustom
End With
End Sub

Private Sub Wait(TimeToWait As Variant)
Dim RetConst1 As Variant
Dim RetConst2 As Integer
Dim RetConst3 As Variant
RetConst1 = GetTickCount()
Do
RetConst2% = DoEvents()
RetConst3 = GetTickCount()
Loop Until RetConst3 - RetConst1 >= TimeToWait * 1000
End Sub

Private Function MakeNotepadScarey()
'DO NOT CLOSE NOTEPAD WHILE THIS FUNCTION IS EXECUTING
WindowState = vbMinimized                   'get out the way
Shell ("notepad.exe"), vbNormalFocus        'any text editor will do, but notepad is pretty universal
Wait 0.1                                    'wait to make sure the window is properly focused
    SendKeys "s "
Wait 0.5
    SendKeys "c "
Wait 0.5
    SendKeys "a "
Wait 0.5
    SendKeys "r "
Wait 0.5
    SendKeys "e "
Wait 0.5
    SendKeys "y"
Wait 0.5
    SendKeys "{TAB}"
Wait 0.5
    SendKeys "n "
Wait 0.5
    SendKeys "o "
Wait 0.5
    SendKeys "t "
Wait 0.5
    SendKeys "e "
Wait 0.5
    SendKeys "p "
Wait 0.5
    SendKeys "a "
Wait 0.5
    SendKeys "d "
Wait 0.5
    SendKeys ". "
Wait 0.5
    SendKeys ". "
Wait 0.5
    SendKeys "."
End Function
