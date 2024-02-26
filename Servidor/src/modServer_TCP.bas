Attribute VB_Name = "modServer_TCP"
Option Explicit

' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ******************************************
Public ServerBuffer As clsBuffer

' Initiate TCP Settings
Public Sub MainTcpInit()
    ' Initate our Player Buffer
    Set ServerBuffer = New clsBuffer
    
    ' Set the connection settings
    frmServer.Server_Socket.RemoteHost = "localhost"
    frmServer.Server_Socket.RemotePort = 9005
    
    main_InitMessages
End Sub

Public Sub DestroyMainTCP()
    ' Close socket
    frmServer.Server_Socket.Close
End Sub

' This function start the connection between server and client
Public Function ConnectToServer() As Boolean
Dim Wait As Long
    
    ' Check to see if we are already connected, if so just exit
    If IsServerConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmServer.Server_Socket.Close
    frmServer.Server_Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsServerConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsServerConnected
End Function

' This function check if our socket is connected
Public Function IsServerConnected() As Boolean
    ' Check if socket is connected
    If frmServer.Server_Socket.State = sckConnected Then IsServerConnected = True
End Function

' Receive Incomming Data
Public Sub main_IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    ' Get data from Socket
    frmServer.Server_Socket.GetData buffer, vbUnicode, DataLength
    
    ' Prevent from hacking
    If ((Not buffer) = -1) Or ((Not buffer) = 0) Then Exit Sub
    
    ServerBuffer.WriteBytes buffer()
    
    If ServerBuffer.Length >= 4 Then pLength = ServerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= ServerBuffer.Length - 4
        If pLength <= ServerBuffer.Length - 4 Then
            ServerBuffer.ReadLong
            main_HandleData ServerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If ServerBuffer.Length >= 4 Then pLength = ServerBuffer.ReadLong(False)
    Loop
    ServerBuffer.Trim
    DoEvents
End Sub

' This send the data to it's connected Socket
Public Sub main_SendData(ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim count As Long

    If IsServerConnected Then
        ' Be sure we don't send a incorrect packets
        If ((Not Data) = -1) Or ((Not Data) = 0) Then Exit Sub
        Set buffer = New clsBuffer
        ' Get the data count
        count = (UBound(Data) - LBound(Data)) + 1
        ' Resize the buffer
        buffer.PreAllocate 4 + count
        ' Place Data
        buffer.WriteLong count
        buffer.WriteBytes Data()
        ' Send
        frmServer.Server_Socket.SendData buffer.ToArray()
        Set buffer = Nothing
    End If
End Sub

' /////////////////////////////////
' //// Outgoing Client Packets ////
' /////////////////////////////////
Public Sub main_SendCheckConnection()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    ' Allocate the buffer
    buffer.PreAllocate 9
    ' Place Our Packet
    buffer.WriteLong Server_Packets.main_CCheckPing
    ' Send our port
    buffer.WriteLong frmServer.Socket(0).LocalPort
    ' Send
    main_SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
