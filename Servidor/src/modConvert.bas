Attribute VB_Name = "modConvert"
Option Explicit


Public Conversas(1 To MAX_CONVERSATION) As ConversationRec

Private Type TextLangRec
    Text As String * 255
    tReply(1 To 3) As String * 100
End Type

Private Type ConvDataRec
    TextLang(1 To 3) As TextLangRec
    '//Others
    NoText As Byte
    NoReply As Byte
    CustomScript As Byte
    CustomScriptData As Long
    CustomScriptData2 As Long
    MoveNext As Byte
    tReplyMove(1 To 3) As Byte
    CustomScriptData3 As Long
End Type


Private Type ConversationRec
    Name As String * NAME_LENGTH
    
    '//Data
    ConvData(1 To MAX_CONV_DATA) As ConvDataRec
End Type
Public Sub ConversationConv()
    Dim i As Long, d As Long, f As Long, l As Long

    For i = 1 To MAX_CONVERSATION
        Conversas(i).Name = Conversation(i).Name

        For d = 1 To 10
            Conversas(i).ConvData(d).CustomScript = Conversation(i).ConvData(d).CustomScript
            Conversas(i).ConvData(d).CustomScriptData = Conversation(i).ConvData(d).CustomScriptData
            Conversas(i).ConvData(d).CustomScriptData2 = Conversation(i).ConvData(d).CustomScriptData2
            Conversas(i).ConvData(d).CustomScriptData3 = Conversation(i).ConvData(d).CustomScriptData3
            Conversas(i).ConvData(d).MoveNext = Conversation(i).ConvData(d).MoveNext
            Conversas(i).ConvData(d).NoReply = Conversation(i).ConvData(d).NoReply
            Conversas(i).ConvData(d).NoText = Conversation(i).ConvData(d).NoText
            For f = 1 To 3
            
                Dim fBost As Integer
                fBost = f
                If i >= 1 And i <= 2 Then fBost = i Else: fBost = 2
                
                Conversas(i).ConvData(d).TextLang(f).Text = Conversation(i).ConvData(d).TextLang(fBost).Text
                For l = 1 To 3
                    Conversas(i).ConvData(d).TextLang(f).tReply(l) = Conversation(i).ConvData(d).TextLang(fBost).tReply(l)
                Next l
                Conversas(i).ConvData(d).tReplyMove(f) = Conversation(i).ConvData(d).tReplyMove(fBost)
            Next f
        Next d

    Next i

    ChkDir App.Path & "\data\", "conversas"

    For i = 1 To MAX_CONVERSATION
        Call SaveConversas(i)
    Next i
End Sub

Public Sub SaveConversas(ByVal ConversationNum As Long)
Dim filename As String
Dim f As Long
    
    filename = App.Path & "\data\conversas\conversationdata_" & ConversationNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Conversas(ConversationNum)
    Close #f
    DoEvents
End Sub
