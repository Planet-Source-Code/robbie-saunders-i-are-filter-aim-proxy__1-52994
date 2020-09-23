VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2040
      Top             =   2160
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmAlert.frx":0000
      Stretch         =   -1  'True
      Top             =   80
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "robbies signed on."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API Declarations
Private Declare Function GetSystemMetrics& Lib "User32" (ByVal nIndex As Long)
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Constants
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 17   ' Height of window client area
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

' Declarations
Private fX As Long
Private fY As Long
Private lngScaleX As Long
Private lngScaleY As Long
Private AlertIndex As Long

Public Sub DisplayAlert(MessageText As String, Duration As Long)

    Dim wFlags As Long, X As Long
    
    ' Set the message
    Label1 = MessageText

    ' Set the duration
    Timer1.interval = Duration

    ' Get the system metrics we need
    fX = GetSystemMetrics(SM_CXFULLSCREEN)
    fY = GetSystemMetrics(SM_CYFULLSCREEN)
    lngScaleX = Me.Width - Me.ScaleWidth
    lngScaleY = Me.Height - Me.ScaleHeight
    
    ' Size the form
    Me.Height = Shape1.Height + 10
    Me.Width = Shape1.Width + 10
    Me.Left = fX * Screen.TwipsPerPixelX - Me.Width
    Me.Top = (fY * Screen.TwipsPerPixelY) - ((Shape1.Height + lngScaleY)) + 400
    Me.Show

    ' Open the alert box
    Timer2.Enabled = True
    
    StayOnTop Me
    
End Sub

Private Sub Timer1_Timer()
    Timer3.Enabled = True
End Sub

Private Sub Timer2_Timer()
    Dim curHeight As Long
    Dim newHeight As Long
    curHeight = Me.Height
    If curHeight < Shape1.Height + lngScaleY Then
        newHeight = curHeight + 30
        If newHeight > Shape1.Height + lngScaleY Then newHeight = Shape1.Height + lngScaleY
        Me.Height = Me.Height + (newHeight - curHeight)
        Me.Top = Me.Top - (newHeight - curHeight) - 300
    Else
        Timer2.Enabled = False
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer3_Timer()
    Dim curHeight As Long
    curHeight = Me.Height
    If curHeight > 50 Then
        Me.Height = curHeight - 30
        Me.Top = Me.Top + 30
    Else
        Unload Me
    End If
End Sub

