VERSION 5.00
Begin VB.UserControl aPacket 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Windowless      =   -1  'True
End
Attribute VB_Name = "aPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim thePacket As String, intSequence As Integer, intLength As Integer
Dim intFamily As Integer, intSubType As Integer
Dim strFamily As String, strSub As String, boolOute As Boolean
Dim A1, A2, A3, A4, A5, A6, A7, A8, A9
Public Event gotInfo(intID As Integer, strPacket As String, intSequence As Integer, intLength As Integer, intFamily As Integer, intSubType As Integer, strFamily As String, strSub As String, boolOut As Boolean)

Public Sub getInfo(intID As Integer)
    RaiseEvent gotInfo(intID, thePacket, intSequence, intLength, intFamily, intSubType, strFamily, strSub, boolOute)
End Sub

Public Function justData()
    justData = thePacket
End Function

Public Sub setPacket(strPacket As String, lstDefs As ListBox, Snacers As String, boolOut As Boolean)
    
    'sequence bytes
    intSequence = GetLength(ChrA("0 0") & Mid(strPacket, 3, 2))
    
    'packet length
    intLength = MessyLength(ChrA("0 0") & Mid(strPacket, 5, 2))
    
    If Len(strPacket) < 10 Then Exit Sub 'if it's just a keep-alive ignore it
    
    
    'family / type
    intFamily = GetLength(Chr(0) & Mid(strPacket, 7, 2))
    Select Case intFamily
        
        Case 0
            strFamily = "*undefined*"
        
        Case Is <= 33
            strFamily = lstDefs.List(intFamily - 1)
            
        Case 64
            strFamily = "AOL"
        
        Case 98
            strFamily = "PLOT"
            
        Case 1027
            strFamily = "LOGIN"
            
        Case 1035
            strFamily = "TUNNEL"
        
        Case 1057
            strFamily = "SECURID"
        
        Case 1098
            strFamily = "ARS"
            
        Case Else
            strFamily = "*undefined*"
        
    End Select

    'sub / type
    intSubType = GetLength(Chr(0) & Mid(strPacket, 9, 2))
    If strFamily <> "*undefined*" Then
        A1 = InStr(1, Snacers, "#define " & strFamily & "__ERR 1")
        A2 = InStr(A1, Snacers, CStr(intSubType)) 'find the next type
        'walk backwards until we hit a space
        A3 = A2 - 2
        Do While Mid(Snacers, A3, 1) <> " "
            A3 = A3 - 1
            DoEvents
        Loop
        'cut it up nicely
        strSub = Replace(Mid(Snacers, A3 + 1, A2 - (A3 + 2)), strFamily & "__", "")
    Else
        strSub = "*undefined*"
    End If
    
    thePacket = strPacket
    boolOute = boolOut

End Sub
