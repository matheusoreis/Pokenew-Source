Attribute VB_Name = "modServer_Packet"
Option Explicit

Public Enum Main_Server_Packets
    main_SSendPing = 1
    MainServerPacket_Count
End Enum

Public Enum Server_Packets
    main_CCheckPing = 1
    ServerPacket_Count
End Enum

Public Server_PacketHandler(MainServerPacket_Count) As Long
