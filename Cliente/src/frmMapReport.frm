VERSION 5.00
Begin VB.Form frmMapReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Mapas"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstIndex 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      ItemData        =   "frmMapReport.frx":0000
      Left            =   120
      List            =   "frmMapReport.frx":0007
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox txtSearch 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "frmMapReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lstIndex_DblClick()
    If lstIndex.ListIndex < 0 Then Exit Sub
    SendWarpTo Val(lstIndex.List(lstIndex.ListIndex))
End Sub

Private Sub txtSearch_Change()
    Dim Find As String, i As Long

    ' Clear the list
    lstIndex.Clear

    Find = UCase$(Trim$(txtSearch.Text))
    '  If Len(Find) <= 2 And Not Find = "" Then
    '  lstIndex.AddItem "Search string too small."
    '   Exit Sub
    ' End If

    For i = 1 To MAX_MAP
        If Not Find = "" Then
            If InStr(1, UCase$(Trim$(MapReport(i))), Find) > 0 Then
                lstIndex.AddItem i & ": " & Trim$(MapReport(i))
            End If
        Else
            lstIndex.AddItem i & ": " & Trim$(MapReport(i))
        End If
    Next

    If lstIndex.ListCount > 0 Then lstIndex.ListIndex = 0
End Sub

