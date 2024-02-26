Attribute VB_Name = "modServer_HandleData"
Option Explicit

Public Sub main_InitMessages()
    Server_PacketHandler(main_SSendPing) = GetAddress(AddressOf Server_HandlePing)
End Sub

Public Sub main_HandleData(ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long
    
    ' Prevent from receiving a empty data
    If ((Not Data) = -1) Or ((Not Data) = 0) Then Exit Sub

    ' Init Buffer
    Set buffer = New clsBuffer
    ' Get size from data
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong

    ' Prevent Hacking
    If MsgType < 0 Then Exit Sub
    If MsgType >= MainServerPacket_Count Then Exit Sub
    
    CallWindowProc Server_PacketHandler(MsgType), 1, buffer.ReadBytes(buffer.Length), 0, 0
    
    ' Clear Buffer
    Set buffer = Nothing
End Sub

Private Sub Server_HandlePing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
End Sub
