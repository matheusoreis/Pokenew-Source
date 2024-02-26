Attribute VB_Name = "ItemDatabase"
Public Sub ClearItem(ByVal ItemNum As Long)
    Call ZeroMemory(ByVal VarPtr(Item(ItemNum)), LenB(Item(ItemNum)))
    Item(ItemNum).Name = vbNullString
    Item(ItemNum).Description = vbNullString
End Sub

Public Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEM
        ClearItem i
    Next
End Sub

Public Sub LoadItem(ByVal ItemNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\items\itemdata_" & ItemNum & ".dat"
    f = FreeFile
    
    If Not FileExist(filename) Then
        ClearItem ItemNum
        SaveItem ItemNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Item(ItemNum)
    Close #f

    DoEvents
End Sub

Public Sub LoadItems()
Dim i As Long

    For i = 1 To MAX_ITEM
        LoadItem i
    Next
End Sub

Public Sub SaveItem(ByVal ItemNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\items\itemdata_" & ItemNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Item(ItemNum)
    Close #f
    DoEvents
End Sub

Public Sub SaveItems()
Dim i As Long

    For i = 1 To MAX_ITEM
        SaveItem i
    Next
End Sub
