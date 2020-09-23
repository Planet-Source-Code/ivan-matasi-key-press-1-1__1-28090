VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   Icon            =   "frmPress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "="
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Key pressed"
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Key code"
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INS"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUM"
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CAPS"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SCROLL"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CTRL"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ALT"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Napravio Ivan Matasiæ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   960
         Width           =   2415
      End
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Close CD"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Open CD"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "**************************************************************"
      Height          =   135
      Left            =   0
      TabIndex        =   14
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API Delaration
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'End API



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Label15.Caption = "Non-ASCII character"
  khr = Chr(KeyCode)
  ask = Asc(khr)
  '-----------------
  If KeyCode = 16 Then '
    Label12.Enabled = True
    Label12.FontBold = True
    Label12.BackColor = vbGreen
  End If
  If KeyCode = 17 Then '
    Label13.Enabled = True
    Label13.FontBold = True
    Label13.BackColor = vbGreen
  End If
  If KeyCode = 18 Then '
    Label14.Enabled = True
    Label14.FontBold = True
    Label14.BackColor = vbGreen
  End If
  '-----------------
  If KeyCode = vbKeyLeft Then Label2.Caption = "Left"
  If KeyCode = vbKeyRight Then Label2.Caption = "Right"
  If KeyCode = vbKeyUp Then Label2.Caption = "Up"
  If KeyCode = vbKeyDown Then Label2.Caption = "Down"
  '-----------------
  If KeyCode = vbKeyDelete Then Label2.Caption = "Delete"
  If KeyCode = vbKeyHome Then Label2.Caption = "Home"
  If KeyCode = vbKeyEnd Then Label2.Caption = "End"
  If KeyCode = vbKeyPageUp Then Label2.Caption = "PageUp"
  If KeyCode = vbKeyPageDown Then Label2.Caption = "PageDown"
  '-----------------
  If KeyCode = vbKeyPrint Then Label2.Caption = "PrintScr"
  If KeyCode = vbKeyPause Then Label2.Caption = "Pause"
  '-----------------
  For fbr = 112 To 123
    If KeyCode = fbr Then Label2.Caption = "F" & fbr - 111
  Next fbr
  '-----------------
  If KeyCode = 91 Then Label2.Caption = "Win Start L"
  If KeyCode = 92 Then Label2.Caption = "Win Start R"
  If KeyCode = 93 Then Label2.Caption = "Menu"
  '-----------------
  If KeyCode = vbKeyInsert Then Label2.Caption = "Insert"
  If KeyCode = vbKeyNumlock Then Label2.Caption = "NumLock"
  If KeyCode = vbKeyCapital Then Label2.Caption = "Caps"
  If KeyCode = vbKeyScrollLock Then Label2.Caption = "Scroll"
  '-----------------
  If CapsLockOn = True Then Label10.Enabled = True Else Label10.Enabled = False
  If NumLockOn = True Then Label9.Enabled = True Else Label9.Enabled = False
  If ScrollLockOn = True Then Label11.Enabled = True Else Label11.Enabled = False
  If InsertOn = True Then Label8.Enabled = True Else Label8.Enabled = False
  '-----------------
  
  Label1.Caption = ask
End Sub

Public Function NumLockOn() As Boolean
   Dim iKeyState As Integer
   iKeyState = GetKeyState(vbKeyNumlock)
   NumLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function

Public Function CapsLockOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyCapital)
    CapsLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function

Public Function ScrollLockOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyScrollLock)
    ScrollLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function

Public Function InsertOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyInsert)
    InsertOn = (iKeyState = 1 Or iKeyState = -127)
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)

Label15.Caption = "ASCII character"
khr = Chr(KeyAscii)
ask = Asc(khr)
Label1.Caption = ask
Label2.Caption = khr

If ask = 27 Then Label2.Caption = "Escape"
If ask = 8 Then Label2.Caption = "Backspace"
If ask = 13 Then Label2.Caption = "Enter"
If ask = 9 Then Label2.Caption = "Tab"
If ask = 10 Then Label2.Caption = "Line feed"
If ask = 32 Then Label2.Caption = "Space"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 Then '
    Label12.Enabled = False
    Label12.FontBold = False
    Label12.BackColor = vbRed
  End If
  If KeyCode = 17 Then '
    Label13.Enabled = False
    Label13.FontBold = False
    Label13.BackColor = vbRed
  End If
  If KeyCode = 18 Then '
    Label14.Enabled = False
    Label14.FontBold = False
    Label14.BackColor = vbRed
  End If
End Sub

Private Sub Form_Load()
If CapsLockOn = True Then Label10.Enabled = True Else Label10.Enabled = False
If NumLockOn = True Then Label9.Enabled = True Else Label9.Enabled = False
If ScrollLockOn = True Then Label11.Enabled = True Else Label11.Enabled = False
If InsertOn = True Then Label8.Enabled = True Else Label8.Enabled = False

End Sub

Private Sub Label16_Click()
Label17.Enabled = True
Label16.Enabled = False
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Sub

Private Sub Label17_Click()
Label17.Enabled = False
Label16.Enabled = True
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door closed", strReturn, 127, 0)
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub Label5_Click()
Label15.Visible = True
Label16.ForeColor = vbGreen
Label17.ForeColor = vbGreen
End Sub

