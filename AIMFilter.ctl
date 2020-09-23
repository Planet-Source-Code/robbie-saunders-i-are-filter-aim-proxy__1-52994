VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinSck.ocx"
Begin VB.UserControl AIMFilter 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ScaleHeight     =   3480
   ScaleWidth      =   4785
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4200
      Top             =   2880
   End
   Begin VB.TextBox SBYTE 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox SBYTE 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2280
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1320
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1800
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "AIMFilter.ctx":0000
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "AIMFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim TheData(3) As String, IncomingStuff(3) As String, PacketLen(3), LeftOver(3) As Boolean
Dim JoMama1, JoMama2, JoMama3, JoMama4
Dim directConnectName As String
Public Event LostConnection()
Public Event incomingPacket(TheIndex As Integer, TheStuff As String)
Public Event incomingPacketDC(TheIndex As Integer, TheStuff As String, TheName As String)
Public Event FreePort(ThePort As Integer)

Public Sub DoTheCrap(requestID As Long)
Winsock1(0).Close
Winsock1(1).Close
Winsock1(0).Accept requestID
'Winsock1(1).Connect TheServer, ThePort
End Sub

Public Sub sendPacket(Index As Integer, TheStuff As String, TheTimes)
On Error Resume Next
If Left(TheStuff, 1) <> "*" Then TheStuff = ChrA("H2A 2 0 0 0 0") & TheStuff
For i = 1 To TheTimes
    If Winsock1(Index).State = sckConnected Then
        SBYTE(Index) = SBYTE(Index) + 1
        If SBYTE(Index) > 65535 Then SBYTE(Index) = 0
        JoMama4 = SBYTE(Index)
        Mid(TheStuff, 3, 2) = IntegerToBase256(CInt(SBYTE(Index)))
        Mid(TheStuff, 5, 2) = IntegerToBase256(Len(TheStuff) - 6) 'correct any length jackups
        Winsock1(Index).SendData TheStuff
    Else
        RaiseEvent LostConnection
    End If
Next i
End Sub

Private Sub Timer1_Timer()
    If Winsock1(0).State <> sckConnected Or Winsock1(1).State <> sckConnected Then
        RaiseEvent LostConnection
    End If
End Sub

Private Sub Winsock1_Connect(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        Winsock1(0).SendData Chr(4) & Chr(90) & JoMama3
        Timer1.Enabled = True
    End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'Incoming Packet Splitter/Joiner
On Error Resume Next
Winsock1(Index).GetData IncomingStuff(Index) 'grab new info
TheData(Index) = TheData(Index) & IncomingStuff(Index) 'add it to cache
NextOne:
If Left(TheData(Index), 1) = "*" And Asc(Mid(TheData(Index), 2, 1)) <= 5 Then
    If LeftOver(Index) = False Then PacketLen(Index) = GetLength(Chr(0) & Mid(TheData(Index), 5, 2)) + 6 'grab the instruction length
    If PacketLen(Index) > Len(TheData(Index)) Then
        LeftOver(Index) = True 'make sure we don't regrab the length bytes and slow us down
        Exit Sub
    End If
    RaiseEvent incomingPacket(Index, Left(TheData(Index), PacketLen(Index))) 'process the instruction
    TheData(Index) = Right(TheData(Index), Len(TheData(Index)) - PacketLen(Index)) 'remove it from the cache
    PacketLen(Index) = 0
    LeftOver(Index) = False 're-enable len grabbing
    If Len(TheData(Index)) >= 6 Then GoTo NextOne: 'if we have the complete header of the next packet then keep going
ElseIf Left(TheData(Index), 2) = Chr(4) & Chr(1) Then
    JoMama1 = GetLength(Chr(0) & Mid(TheData(Index), 3, 2))
    JoMama2 = Asc(Mid(TheData(Index), 5, 1)) & "." & Asc(Mid(TheData(Index), 6, 1)) & "." & Asc(Mid(TheData(Index), 7, 1)) & "." & Asc(Mid(TheData(Index), 8, 1))
    JoMama3 = Mid(TheData(Index), 3, 6)
    Winsock1(1).Connect JoMama2, JoMama1
    TheData(Index) = ""
ElseIf Left(TheData(Index), 4) = "ODC2" Then
            If Index = 1 Then
                'get whose name it is
                directConnectName = Replace(Mid(TheData(Index), 45, 16), Chr(0), "")
            End If
    If Mid(TheData(Index), 29, 2) = ChrA("0 0") And Asc(Mid(TheData(Index), 31, 1)) < 9 And Asc(Mid(TheData(Index), 32, 1)) <> 0 Then

            If Index = 0 And directConnectName <> "" Then
                If Label1 > 30000 Then
                    RaiseEvent incomingPacketDC(Index, TheData(Index), directConnectName) 'process the instruction
                    TheData(Index) = ""
                    Exit Sub
                End If
            End If
    End If
    
        If Index = 0 Then
            Winsock1(1).SendData TheData(Index)
        Else
            Winsock1(0).SendData TheData(Index)
        End If
    TheData(Index) = ""
Else
    If Index = 0 Then
        Winsock1(1).SendData TheData(Index)
    Else
        Winsock1(0).SendData TheData(Index)
    End If
    TheData(Index) = ""
End If
End Sub

Private Sub Winsock1_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Label1 = bytesRemaining
End Sub
