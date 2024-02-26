Attribute VB_Name = "SharedFunctions"
Public Function ByteToBoolean(ByVal Value As Integer) As Boolean
    If Value = 0 Then
        ByteToBoolean = False
    Else
        ByteToBoolean = True
    End If
End Function

Public Function BooleanToByte(ByVal Value As Boolean) As Byte
    If Value = True Then
        BooleanToByte = 1
    Else
        BooleanToByte = 0
    End If
End Function

Public Function AlertToPlayer(ByVal Index As Long, ByVal TextPT As Long, ByVal TextEN As Long, ByVal TextES As Long)
    Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, TextPT, White
        Case LANG_EN: AddAlert Index, TextEN, White
        Case LANG_ES: AddAlert Index, TextES, White
    End Select
End Function
