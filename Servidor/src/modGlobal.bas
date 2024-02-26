Attribute VB_Name = "SharedVariables"
Option Explicit

'//General
Public AppRunning As Boolean          '//Controls whether the program is running or not...
Public DebugMode As Boolean           '//Check whether the program is on debug mode or not...

'//CPS
Public GameCPS As Long
Public CPSUnlock As Boolean

'//Index
Public Player_HighIndex As Long
Public Pokemon_HighIndex As Long

'//Map
Public PlayerOnMap(1 To MAX_MAP) As Byte

'//Shutdowng
Public isShuttingDown As Boolean
Public Secs As Long

Public CountText As Long

Public MAX_PLAYER As Integer

'// Hora desvinculada ao sistema operacional, recebe apenas ao ligar o servidor.
Public GameHour As Byte
Public GameMinute As Byte
Public GameSecs As Byte
Public GameSecs_Velocity As Byte

Public AlertPT As Long
Public AlertEN As Long
Public AlertES As Long
