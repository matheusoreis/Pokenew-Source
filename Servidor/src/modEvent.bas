Attribute VB_Name = "modEvent"
Option Explicit

Public InRumbleTournament As Boolean
Public TournamentStarted As Boolean
Public BattleStarted As Boolean
Public Participant(1 To MAX_PLAYER) As Boolean
Public ParticipantIndex(1 To MAX_PLAYER) As Long
Public ParticipantSlot(1 To MAX_PLAYER) As Long
Public PokemonAlive(1 To MAX_PLAYER) As Long
Public ParticipantScore(1 To MAX_PLAYER) As Long
Public ParticipantCount As Long
Public TournamentStartingTime As Long
Public TourStartCount As Long
Public EventTimer As Long
Public BattleStartTimer As Long

'//////////
Public Event_TData As Event_TDataRec
Public EventSchedule As EventScheduleRec

Public Type Event_TDataRec
    EventMap As Long
    EventX As Long
    EventY As Long
End Type

Private Type ScheduleDataRec
    sType As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
End Type

Public Type EventScheduleRec
    MaxSchedule As Long
    ScheduleData() As ScheduleDataRec
End Type

Public Sub LoadEventData()
Dim FileName As String
Dim I As Long

    FileName = App.Path & "\data\event_data.ini"
    If Not FileExist(FileName) Then
        ClearEventData
        SaveEventData
    Else
        Event_TData.EventMap = Val(GetVar(FileName, "General", "EventMap"))
        Event_TData.EventX = Val(GetVar(FileName, "General", "EventX"))
        Event_TData.EventY = Val(GetVar(FileName, "General", "EventY"))
    End If
End Sub

Public Sub ClearEventData()
    Event_TData.EventMap = 1
    Event_TData.EventX = 10
    Event_TData.EventY = 10
End Sub

Public Sub SaveEventData()
Dim FileName As String
Dim I As Long

    FileName = App.Path & "\data\event_data.ini"
    PutVar FileName, "General", "EventMap", Str(Event_TData.EventMap)
    PutVar FileName, "General", "EventX", Str(Event_TData.EventX)
    PutVar FileName, "General", "EventY", Str(Event_TData.EventY)
End Sub

Public Sub LoadEventSched()
Dim FileName As String
Dim I As Long

    FileName = App.Path & "\data\event_schedule.ini"
    If Not FileExist(FileName) Then
        ClearEventSched
        SaveEventSched
    Else
        EventSchedule.MaxSchedule = Val(GetVar(FileName, "General", "MaxSchedule"))
        If EventSchedule.MaxSchedule <= 0 Then Exit Sub
        ReDim EventSchedule.ScheduleData(1 To EventSchedule.MaxSchedule) As ScheduleDataRec
        For I = 1 To EventSchedule.MaxSchedule
            EventSchedule.ScheduleData(I).sType = Val(GetVar(FileName, "Schedule " & I, "Type"))
            EventSchedule.ScheduleData(I).Data1 = Val(GetVar(FileName, "Schedule " & I, "Data1"))
            EventSchedule.ScheduleData(I).Data2 = Val(GetVar(FileName, "Schedule " & I, "Data2"))
            EventSchedule.ScheduleData(I).Data3 = Val(GetVar(FileName, "Schedule " & I, "Data3"))
            EventSchedule.ScheduleData(I).Data4 = Val(GetVar(FileName, "Schedule " & I, "Data4"))
        Next
    End If
End Sub

Public Sub ClearEventSched()
    ' /////// Schedule
    EventSchedule.MaxSchedule = 2
    ReDim EventSchedule.ScheduleData(1 To 2) As ScheduleDataRec
    
    ' Type
    ' 1 = Every Week
        ' Data 1 = Week Day ([1 Sunday], [2 - 6 Mon - Fri], [7 Saturday])
        ' Data 2 = None
    ' 2 = Every Day
        ' Data 1 = None
        ' Data 2 = None
    ' 3 = Every Month
        ' Data 1 = Day
        ' Data 2 = None
    ' 4 = Set Date
        ' Data 1 = Month
        ' Data 2 = Day
    
    ' Data 3 = Game Hour
    ' Data 4 = Game Minute
    
    EventSchedule.ScheduleData(1).sType = 1 ' Every Week
    EventSchedule.ScheduleData(1).Data1 = 1 ' Sunday
    EventSchedule.ScheduleData(1).Data2 = 0
    EventSchedule.ScheduleData(1).Data3 = 0
    EventSchedule.ScheduleData(1).Data4 = 0

    EventSchedule.ScheduleData(2).sType = 1 ' Every Week
    EventSchedule.ScheduleData(2).Data1 = 7 ' Saturday
    EventSchedule.ScheduleData(2).Data2 = 0
    EventSchedule.ScheduleData(2).Data3 = 0
    EventSchedule.ScheduleData(2).Data4 = 0
End Sub

Public Sub SaveEventSched()
Dim FileName As String
Dim I As Long

    FileName = App.Path & "\data\event_schedule.ini"
    PutVar FileName, "General", "MaxSchedule", Str(EventSchedule.MaxSchedule)
    For I = 1 To EventSchedule.MaxSchedule
        PutVar FileName, "Schedule " & I, "Type", Str(EventSchedule.ScheduleData(I).sType)
        PutVar FileName, "Schedule " & I, "Data1", Str(EventSchedule.ScheduleData(I).Data1)
        PutVar FileName, "Schedule " & I, "Data2", Str(EventSchedule.ScheduleData(I).Data2)
        PutVar FileName, "Schedule " & I, "Data3", Str(EventSchedule.ScheduleData(I).Data3)
        PutVar FileName, "Schedule " & I, "Data4", Str(EventSchedule.ScheduleData(I).Data4)
    Next
End Sub

Public Function InTournament(ByVal Index As Integer) As Boolean
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    If Not TournamentStarted Then Exit Function
    If Not Participant(Index) Then Exit Function
    If Not BattleStarted Then Exit Function
    InTournament = True
End Function

Public Sub TournamentLogic()
Dim myDate As String
Dim GameHour As Long, GameMinute As Long
Dim sWeekDay As Long, curDay As Long, curMonth As Long
Dim I As Long

    If InRumbleTournament Then
        If Not TournamentStarted Then
            InitTournament
        Else
            If Not BattleStarted Then
                If BattleStartTimer <= GetTickCount Then
                    BattleStarted = True
                End If
            Else
                If ParticipantCount <= 1 Then
                    EndTournament
                Else
                    If EventTimer <= GetTickCount Then
                        EndTournament
                    End If
                End If
            End If
        End If
    Else
        myDate = Format$(Date, DateFormat)
        GameHour = Hour(Now)
        GameMinute = Minute(Now)
        sWeekDay = Weekday(myDate)
        curDay = Day(myDate)
        curMonth = Month(myDate)
        StartTournament
        
        If EventSchedule.MaxSchedule > 0 Then
            For I = 1 To EventSchedule.MaxSchedule
                With EventSchedule.ScheduleData(I)
                    Select Case .sType
                        Case 1 ' Every Week
                            If Not .Data1 = sWeekDay Then GoTo continue
                        Case 2 ' Every Day
                        Case 3 ' Every Month
                            If Not .Data1 = curDay Then GoTo continue
                        Case 4 ' Set Date
                            If Not .Data1 = curMonth Then GoTo continue
                            If Not .Data2 = curDay Then GoTo continue
                    End Select
                    If Not .Data3 = GameHour Then GoTo continue
                    If Not (.Data4 >= GameMinute - 5 And .Data4 <= GameMinute + 5) Then GoTo continue
                    GoTo inittour
                End With
continue:
            Next
        End If
        
        Exit Sub
inittour:
        StartTournament
    End If
End Sub

Public Sub EndTournament()
Dim wIndex As Long, xInd As Long
Dim I As Long
Dim sScore As Long

    If ParticipantCount <= 1 Then
        wIndex = ParticipantIndex(1)
        If wIndex <= 0 Then
            ' Failed
            SendGlobalMsg "Event ended without a winner", White
            GoTo clearTournament
        End If
        SendGlobalMsg Trim$(Player(wIndex, TempPlayer(wIndex).UseChar).Name) & " win", White
    Else
        sScore = 0
        wIndex = 0
        For I = 1 To ParticipantCount
            xInd = ParticipantIndex(I)
            If xInd > 0 Then
                If ParticipantScore(xInd) > sScore Then
                    sScore = ParticipantScore(xInd)
                    wIndex = xInd
                ElseIf ParticipantScore(xInd) = sScore Then
                    sScore = ParticipantScore(xInd)
                    wIndex = 0
                End If
            End If
        Next
        If wIndex <= 0 Then
            ' Failed
            SendGlobalMsg "Event ended without a winner", White
            GoTo clearTournament
        End If
        SendGlobalMsg Trim$(Player(wIndex, TempPlayer(wIndex).UseChar).Name) & " win", White
    End If
    
clearTournament:
    For I = 1 To MAX_PLAYER
        Participant(I) = False
        ParticipantIndex(I) = 0
        ParticipantSlot(I) = 0
        ParticipantScore(I) = 0
    Next
    InRumbleTournament = False
    ParticipantCount = 0
    TournamentStarted = False
    EventTimer = 0
    BattleStarted = False
End Sub

Public Function AddParticipantCount(ByVal Index As Integer) As Boolean
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    If ParticipantCount >= MAX_PLAYER Then Exit Function
    If Participant(Index) Then Exit Function
    If Not InRumbleTournament Then Exit Function
    Participant(Index) = True
    ParticipantCount = ParticipantCount + 1
    ParticipantIndex(ParticipantCount) = Index
    ParticipantSlot(Index) = ParticipantCount
    ParticipantScore(Index) = 0
    AddParticipantCount = True
End Function

Public Sub RemoveParticipantCount(ByVal Index As Integer)
Dim Slot As Long
Dim I As Long
Dim LastIndex As Long
Dim pIndex As Long

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If ParticipantCount <= 0 Then Exit Sub
    If ParticipantSlot(Index) <= 0 Then Exit Sub
    If Not Participant(Index) Then Exit Sub
    Participant(Index) = False
    LastIndex = ParticipantCount
    Slot = ParticipantSlot(Index)
    ParticipantSlot(Index) = 0
    ParticipantScore(Index) = 0
    If Slot <= 0 Or Slot > MAX_PLAYER Then Exit Sub
    ParticipantIndex(Slot) = 0
    ParticipantCount = ParticipantCount - 1
    
    For I = Slot To LastIndex
        ParticipantIndex(I) = ParticipantIndex(I + 1)
        pIndex = ParticipantIndex(I)
        If pIndex > 0 And pIndex <= MAX_PLAYER Then
            ParticipantSlot(pIndex) = I
        End If
    Next
    
    ' Warp Out
    If Player(pIndex, TempPlayer(pIndex).UseChar).Map = Event_TData.EventMap Then
        ' Warp Out
        ' Temp
        SendPlayerMsg pIndex, "You have been warped out of the event", White
    End If
End Sub

Public Sub StartTournament()
Dim I As Long
Dim Msg As String

    If InRumbleTournament Then Exit Sub
    
    For I = 1 To MAX_PLAYER
        Participant(I) = False
        ParticipantIndex(I) = 0
        ParticipantSlot(I) = 0
        ParticipantScore(I) = 0
    Next
    ParticipantCount = 0
    TournamentStarted = False
    EventTimer = 0
    BattleStarted = False
    
    SendGlobalMsg "A event has started, if you want to participate type /participate", White
    
    TournamentStartingTime = GetTickCount + 60000
    TourStartCount = 5
    Msg = TourStartCount & " Minute/s"
    SendGlobalMsg "Event will start in " & Msg, White
    
    InRumbleTournament = True
End Sub

Public Sub InitTournament()
Dim Msg As String
Dim I As Long
Dim pIndex As Long

    If Not InRumbleTournament Then Exit Sub
    If TournamentStarted Then Exit Sub
    If TournamentStartingTime <= GetTickCount Then
        TourStartCount = TourStartCount - 1
        If TourStartCount > 0 Then
            Msg = TourStartCount & " Minute/s"
            SendGlobalMsg "Event will start in " & Msg, White
            TournamentStartingTime = GetTickCount + 60000
            Exit Sub
        End If
    End If
    If ParticipantCount < 2 Then Exit Sub
    
    For I = 1 To ParticipantCount
        pIndex = ParticipantIndex(I)
        If pIndex > 0 And pIndex <= MAX_PLAYER Then
            If Participant(pIndex) Then
                'PlayerWarp pIndex, Event_TData.EventMap, Event_TData.EventX, Event_TData.EventY, DIR_DOWN
                ' Temp
                SendPlayerMsg pIndex, "You have been warped to event", White
                PokemonAlive(pIndex) = CountPlayerPokemonAlive(pIndex)
            End If
        End If
    Next
    
    EventTimer = GetTickCount + 300000
    BattleStartTimer = GetTickCount + 2000
    TournamentStarted = True
    SendGlobalMsg "Event has started", White
End Sub
