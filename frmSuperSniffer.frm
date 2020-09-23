VERSION 5.00
Begin VB.Form frmSuperSniffer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " I ARE OSCAR ANALYZER"
   ClientHeight    =   5295
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSuperSniffer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1530
      ItemData        =   "frmSuperSniffer.frx":000C
      Left            =   720
      List            =   "frmSuperSniffer.frx":0073
      TabIndex        =   28
      Top             =   6000
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Caption         =   "Packet Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.Label Label14 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   540
         Width           =   2655
      End
      Begin VB.Label Label13 
         Caption         =   "Length:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "6"
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Sub Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "4"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Family Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "CHANNEL_MSG_TOHOST"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   1260
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Sub:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "ICBM"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   780
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Family:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Sequence:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Capture Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Text            =   "6"
         Top             =   1415
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Text            =   "4"
         Top             =   1050
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Capture Only:"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2505
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Incoming Packets"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Outgoing Packets"
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "Sub Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Family Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Viewing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   6615
      Begin VB.OptionButton Option5 
         Caption         =   "Sexy"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "ASC/Chra"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ascii"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "HEX Chra"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ASC Chra"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Packets"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   4680
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin IAREFILTER.aPacket aPacket 
      Height          =   615
      Index           =   0
      Left            =   5640
      TabIndex        =   32
      Top             =   6960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   255
      Left            =   1920
      TabIndex        =   29
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Menu ListMenu 
      Caption         =   "ListMenu"
      Visible         =   0   'False
      Begin VB.Menu removetheitem 
         Caption         =   "Remove Packet"
      End
      Begin VB.Menu clearpackets 
         Caption         =   "Clear Packets"
      End
   End
End
Attribute VB_Name = "frmSuperSniffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim currentView As String

Private Sub aPacket_gotInfo(Index As Integer, intID As Integer, strPacket As String, intSequence As Integer, intLength As Integer, intFamily As Integer, intSubType As Integer, strFamily As String, strSub As String, boolOut As Boolean)
    If intID = "0" Then
        If boolOut = True Then
            List1.AddItem "OUT :: " & intFamily & " x " & intSubType & String(20, " ") & Chr(1) & Index
        Else
            List1.AddItem "IN  :: " & intFamily & " x " & intSubType & String(20, " ") & Chr(1) & Index
        End If
    ElseIf intID = "1" Then
        Label4 = intSequence
        Label14 = intLength
        Label6 = strFamily
        Label10 = intFamily
        Label8 = strSub
        Label12 = intSubType
        currentView = strPacket
        updateViewer
    End If
End Sub

Public Sub updateViewer()
    On Error Resume Next
    If currentView = "" Then Exit Sub
    Dim B2, B3, B4, B5, B6, B7, B8, B9
    If Option3.Value Then
        Text3 = Replace(currentView, Chr(0), Chr(1))
    ElseIf Option1.Value Then
        B2 = ""
        For i = 1 To Len(currentView)
            B2 = B2 & Asc(Mid(currentView, i, 1)) & " "
            DoEvents
        Next i
        Text3 = B2
    ElseIf Option2.Value Then
        B2 = ""
        For i = 1 To Len(currentView)
            B2 = B2 & Hex(Asc(Mid(currentView, i, 1))) & " "
            DoEvents
        Next i
        Text3 = B2
    ElseIf Option4.Value Then
        B2 = ""
        For i = 0 To (Len(currentView) / 8) - 1
            For X = 1 To 8
                B2 = B2 & fatString(Asc(Mid(currentView, X + (i * 8), 1)), " ", 4)
                DoEvents
            Next X
            B2 = B2 & vbCrLf
            For X = 1 To 8
                B2 = B2 & Replace(fatString(Mid(currentView, X + (i * 8), 1), " ", 4), Chr(0), Chr(1))
                DoEvents
            Next X
            B6 = (i * 8) + 9
            B2 = B2 & vbCrLf
        Next i
        B4 = ""
        B5 = ""
        For i = B6 To Len(currentView)
            B4 = B4 & fatString(Asc(Mid(currentView, i, 1)), " ", 4)
            B5 = B5 & Replace(fatString(Mid(currentView, i, 1), " ", 4), Chr(0), Chr(1))
            DoEvents
        Next i
        Text3 = B2 & B4 & vbCrLf & B5
    ElseIf Option5.Value Then
        If Len(currentView) >= 10 Then
            B2 = "SNAC Header: " & vbCrLf
            For i = 7 To 10
                B2 = B2 & fatString(Asc(Mid(currentView, i, 1)), " ", 4)
                DoEvents
            Next i
            B2 = B2 & vbCrLf
            For i = 7 To 10
                B2 = B2 & Replace(fatString(Mid(currentView, i, 1), " ", 4), Chr(0), Chr(1))
                DoEvents
            Next i
            If Len(currentView) >= 16 Then
                B2 = B2 & vbCrLf & vbCrLf & "Request ID:" & vbCrLf
                For i = 11 To 16
                    B2 = B2 & fatString(Asc(Mid(currentView, i, 1)), " ", 4)
                    DoEvents
                Next i
                B2 = B2 & vbCrLf
                For i = 11 To 16
                    B2 = B2 & Replace(fatString(Mid(currentView, i, 1), " ", 4), Chr(0), Chr(1))
                    DoEvents
                Next i
            End If
            If Len(currentView) >= 20 Then
                B2 = B2 & vbCrLf & vbCrLf & "TLV Guesses:" & vbCrLf
                For i = 17 To Len(currentView) - 3
                    If Mid(currentView, i, 1) = Chr(0) And Mid(currentView, i + 1, 1) <> Chr(0) And Mid(currentView, i + 2, 1) = Chr(0) And Mid(currentView, i + 3, 1) <> Chr(0) Then
                        B7 = Asc(Mid(currentView, i + 1, 1))
                        B8 = Asc(Mid(currentView, i + 3, 1))
                        If (B8 + i + 3) <= Len(currentView) Then
                            B9 = Mid(currentView, i + 4, B8)
                            B2 = B2 & "[" & B7 & "]::[" & Replace(B9, Chr(0), Chr(1)) & "]" & vbCrLf
                        End If
                    End If
                    DoEvents
                Next i
            End If
        End If
        Text3 = B2
    End If
End Sub

Public Sub newPack(strPacket As String, boolOut As Boolean)
    On Error Resume Next
    Label15 = Label15 + 1
    Load aPacket(Label15)
    If FileExist("C:\aimsnacers.htm") Then
        aPacket(Label15).setPacket strPacket, List2, text_read("C:\aimsnacers.htm"), boolOut
    Else
        aPacket(Label15).setPacket strPacket, List2, text_read(MyTruePath & "aimsnacers.htm"), boolOut
    End If
    aPacket(Label15).getInfo "0"
End Sub

Private Sub clearpackets_Click()
    On Error Resume Next
    Dim B1
    For i = 0 To List1.ListCount - 1
        B1 = Split(List1.List(i), Chr(1))
        Unload aPacket(B1(1))
        DoEvents
    Next i
    List1.Clear
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If List1.text <> "" Then
            Dim B1
            B1 = Split(List1.text, Chr(1))
            aPacket(B1(1)).getInfo "1"
        End If
    ElseIf Button = 2 Then
        PopupMenu Me.ListMenu
    End If
End Sub

Private Sub Option1_Click()
    updateViewer
End Sub

Private Sub Option2_Click()
    updateViewer
End Sub

Private Sub Option3_Click()
    updateViewer
End Sub

Private Sub Option4_Click()
    updateViewer
End Sub

Private Sub Option5_Click()
    updateViewer
End Sub

Private Sub removetheitem_Click()
    Dim B1
    B1 = Split(List1.text, Chr(1))
    List1.RemoveItem List1.ListIndex
    Unload aPacket(B1(1))
End Sub
