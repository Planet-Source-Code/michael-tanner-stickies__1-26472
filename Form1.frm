VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNoteText 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   2440
   End
   Begin VB.PictureBox picContractButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   1800
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   120
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   0
      Width           =   150
   End
   Begin VB.PictureBox picExpandButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2040
      Picture         =   "Form1.frx":0142
      ScaleHeight     =   120
      ScaleWidth      =   150
      TabIndex        =   1
      Top             =   0
      Width           =   150
   End
   Begin VB.PictureBox picExitButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   2280
      Picture         =   "Form1.frx":0284
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   0
      Width           =   165
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "User32" (ByVal _
    hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Dim myForms() As Form1

Sub SetTopmostWindow(ByVal hwnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hwnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_Load()
    If NoteCount = 0 Then
        Call AddToTray(Me.Icon, Me.Caption, Me)
        NoteCount = 1
        ReDim myForms(1 To NoteCount)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault

    If RespondToTray(X) <> 0 Then
        NoteCount = NoteCount + 1
        ReDim Preserve myForms(1 To NoteCount)
        Set myForms(NoteCount) = New Form1
        Call ShowFormAgain(myForms(NoteCount))
        SetTopmostWindow myForms(NoteCount).hwnd
        myForms(NoteCount).Left = (Screen.Width / 2) - (myForms(NoteCount).Width / 2)
        myForms(NoteCount).Top = 60
        myForms(NoteCount).lblTitle.Caption = InputBox("Enter Sticky Note title.", "Title", "Sticky Notes!")
        myForms(NoteCount).Show
        myForms(NoteCount).txtNoteText.SetFocus
    End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbCustom
    Me.MouseIcon = LoadPicture(App.Path & "\hand.ico")
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub picExitButton_Click()
    Call RemoveFromTray(Me)
    Unload Me
    NoteCount = NoteCount - 1
    If NoteCount = 1 Then
        Call RemoveFromTray(Me)
        End
    End If
End Sub

Private Sub picContractButton_Click()
    Me.Height = 250
End Sub

Private Sub picExpandButton_Click()
    Me.Height = 1635
End Sub

Private Sub txtNoteText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbIbeam
End Sub
