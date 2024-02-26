VERSION 5.00
Begin VB.Form frmEditor_Animation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Editor"
   ClientHeight    =   6255
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9510
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   6135
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   5535
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   12
         Top             =   3000
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   3135
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   10
         Top             =   2400
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   9
         Top             =   1800
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Index           =   1
         Left            =   3360
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   8
         Top             =   3480
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Index           =   0
         Left            =   120
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   4
         Top             =   3480
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   23
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Layer 1 (Above Player)"
         Height          =   180
         Left            =   3360
         TabIndex        =   21
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   20
         Top             =   2160
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   19
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   18
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Layer 0 (Below Player)"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Index"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.ListBox lstIndex 
         Height          =   5715
         ItemData        =   "frmEditor_Animation.frx":0000
         Left            =   120
         List            =   "frmEditor_Animation.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2295
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
Attribute VB_Name = "frmEditor_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseAnimationEditor
    End If
End Sub

Private Sub Form_Load()
Dim i As Long

    For i = 0 To 1
        scrlSprite(i).max = Count_Animation
        scrlLoopCount(i).max = 100
        scrlFrameCount(i).max = 100
        scrlLoopTime(i).max = 1000
    Next
End Sub

Private Sub lstIndex_Click()
    AnimationEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestAnimation
    End If
    CloseAnimationEditor
End Sub

Private Sub mnuExit_Click()
    CloseAnimationEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_ANIMATION
        If AnimationChange(i) Then
            SendSaveAnimation i
            AnimationChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseAnimationEditor
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Animation(EditorIndex).Name = (txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub scrlFrameCount_Change(Index As Integer)
    lblFrameCount(Index).Caption = "Frame Count: " & scrlFrameCount(Index).value
    Animation(EditorIndex).Frames(Index) = scrlFrameCount(Index).value
    EditorChange = True
End Sub

Private Sub scrlLoopCount_Change(Index As Integer)
    lblLoopCount(Index).Caption = "Loop Count: " & scrlLoopCount(Index).value
    Animation(EditorIndex).LoopCount(Index) = scrlLoopCount(Index).value
    EditorChange = True
End Sub

Private Sub scrlLoopTime_Change(Index As Integer)
    lblLoopTime(Index).Caption = "Loop Time: " & scrlLoopTime(Index).value
    Animation(EditorIndex).looptime(Index) = scrlLoopTime(Index).value
    EditorChange = True
End Sub

Private Sub scrlSprite_Change(Index As Integer)
    lblSprite(Index).Caption = "Sprite: " & scrlSprite(Index).value
    Animation(EditorIndex).Sprite(Index) = scrlSprite(Index).value
    EditorChange = True
End Sub
