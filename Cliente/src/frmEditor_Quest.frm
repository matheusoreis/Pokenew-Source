VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Editor"
   ClientHeight    =   4470
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9255
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   4215
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   4575
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5280
         ScaleHeight     =   32
         ScaleMode       =   0  'User
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   360
         Width           =   480
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   0
         TabIndex        =   5
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstIndex 
         Height          =   3375
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   120
         List            =   "frmEditor_Quest.frx":0002
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit(Esc)"
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIndexSearch_Click()
Dim FindChar As String
Dim clBound As Long, cuBound As Long
Dim i As Long
Dim ComboText As String
Dim indexString As String
Dim stringLength As Long

    If Len(Trim$(txtIndexSearch.Text)) > 0 Then
        FindChar = Trim$(txtIndexSearch.Text)
        clBound = 1
        cuBound = MAX_QUEST
        
        For i = clBound To cuBound
            ComboText = Trim$(lstIndex.List(i - 1))
            indexString = i & ": "
            stringLength = Len(ComboText) - Len(indexString)
            If stringLength >= 0 Then
                ComboText = Mid$(ComboText, Len(indexString) + 1, stringLength)
                If LCase(ComboText) = LCase(FindChar) Then
                    lstIndex.ListIndex = (i - 1)
                    Exit Sub
                End If
            End If
        Next
        
        MsgBox "Index not found", vbCritical
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseQuestEditor
    End If
End Sub

Private Sub Form_Load()
    txtName.MaxLength = NAME_LENGTH
End Sub

Private Sub lstIndex_Click()
    QuestEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestQuest
    End If
    CloseQuestEditor
End Sub

Private Sub mnuExit_Click()
    CloseQuestEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_QUEST
        If QuestChange(i) Then
            SendSaveQuest i
            QuestChange(i) = False
        End If
    Next
    '//reset
    EditorChange = False
    CloseQuestEditor
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub
