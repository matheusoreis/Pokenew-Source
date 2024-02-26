VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   13110
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   874
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   994
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      Height          =   7335
      Left            =   7560
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CheckBox chkNoCure 
         Caption         =   "No Medicine?"
         Height          =   255
         Left            =   4320
         TabIndex        =   106
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cmbWeather 
         Height          =   315
         ItemData        =   "frmEditor_Map.frx":0000
         Left            =   240
         List            =   "frmEditor_Map.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   3480
         Width           =   2415
      End
      Begin VB.HScrollBar scrlSpriteType 
         Height          =   255
         Left            =   5160
         Max             =   3
         TabIndex        =   87
         Top             =   4200
         Width           =   1815
      End
      Begin VB.HScrollBar scrlCaveLight 
         Height          =   255
         Left            =   4680
         Max             =   5
         TabIndex        =   77
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CheckBox chkCave 
         Caption         =   "Cave?"
         Height          =   255
         Left            =   6120
         TabIndex        =   75
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chkKillPlayer 
         Caption         =   "Kill Player?"
         Height          =   255
         Left            =   2880
         TabIndex        =   74
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmEditor_Map.frx":004C
         Left            =   3720
         List            =   "frmEditor_Map.frx":005F
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Frame Frame4 
         Caption         =   "Map Npc"
         Height          =   1815
         Left            =   2760
         TabIndex        =   38
         Top             =   2280
         Width           =   4215
         Begin VB.ComboBox cmbMapNpc 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1320
            Width           =   3975
         End
         Begin VB.ListBox lstMapNpc 
            Height          =   1035
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.ComboBox cmbMusic 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   840
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         Caption         =   "Map Size"
         Height          =   1215
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   2415
         Begin VB.TextBox txtMaxY 
            Height          =   285
            Left            =   960
            TabIndex        =   25
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtMaxX 
            Height          =   285
            Left            =   960
            TabIndex        =   24
            Text            =   "0"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Labels 
            Caption         =   "Max Y:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Labels 
            Caption         =   "Max X:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Links"
         Height          =   1455
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2415
         Begin VB.TextBox txtLinkDown 
            Height          =   285
            Left            =   840
            TabIndex        =   20
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtLinkRight 
            Height          =   285
            Left            =   1560
            TabIndex        =   19
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtLinkLeft 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtLinkUp 
            Height          =   285
            Left            =   840
            TabIndex        =   17
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdPropertiesSave 
         Caption         =   "Input"
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   6720
         Width           =   1095
      End
      Begin VB.CommandButton cmdPropertiesCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         Top             =   6720
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   3720
         TabIndex        =   13
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Start Weather:"
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblSpriteType 
         Caption         =   "Sprite Type: None"
         Height          =   255
         Left            =   2760
         TabIndex        =   88
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label lblCaveLight 
         Caption         =   "Cave Light: 0"
         Height          =   255
         Left            =   2760
         TabIndex        =   76
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Moral"
         Height          =   255
         Left            =   2760
         TabIndex        =   72
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Labels 
         Caption         =   "Music:"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   26
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Labels 
         Caption         =   "Name:"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   12
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.Frame fraAttribute 
      Height          =   7335
      Left            =   7560
      TabIndex        =   41
      Top             =   7680
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame fraCheckpoint 
         Caption         =   "Warp Properties"
         Height          =   2535
         Left            =   1680
         TabIndex        =   93
         Top             =   2280
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton cmdAttributeCancel 
            Caption         =   "Cancel"
            Height          =   375
            Index           =   4
            Left            =   1920
            TabIndex        =   98
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton cmdAttributeOkay 
            Caption         =   "Okay"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   99
            Top             =   1920
            Width           =   1695
         End
         Begin VB.HScrollBar scrlCheckMap 
            Height          =   255
            Left            =   1440
            Max             =   0
            TabIndex        =   97
            Top             =   360
            Width           =   2175
         End
         Begin VB.HScrollBar scrlCheckX 
            Height          =   255
            Left            =   1440
            Max             =   100
            TabIndex        =   96
            Top             =   720
            Width           =   2175
         End
         Begin VB.HScrollBar scrlCheckY 
            Height          =   255
            Left            =   1440
            Max             =   100
            TabIndex        =   95
            Top             =   1080
            Width           =   2175
         End
         Begin VB.HScrollBar scrlCheckDir 
            Height          =   255
            Left            =   1440
            Max             =   3
            TabIndex        =   94
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblCheckX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblCheckMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblCheckY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblCheckDir 
            Caption         =   "Dir: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   1440
            Width           =   855
         End
      End
      Begin VB.Frame fraWarp 
         Caption         =   "Warp Properties"
         Height          =   2535
         Left            =   1680
         TabIndex        =   58
         Top             =   2280
         Visible         =   0   'False
         Width           =   3855
         Begin VB.HScrollBar scrlWarpDir 
            Height          =   255
            Left            =   1440
            Max             =   3
            TabIndex        =   68
            Top             =   1440
            Width           =   2175
         End
         Begin VB.HScrollBar scrlWarpY 
            Height          =   255
            Left            =   1440
            Max             =   100
            TabIndex        =   66
            Top             =   1080
            Width           =   2175
         End
         Begin VB.HScrollBar scrlWarpX 
            Height          =   255
            Left            =   1440
            Max             =   100
            TabIndex        =   64
            Top             =   720
            Width           =   2175
         End
         Begin VB.HScrollBar scrlWarpMap 
            Height          =   255
            Left            =   1440
            Max             =   0
            TabIndex        =   63
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdAttributeCancel 
            Caption         =   "Cancel"
            Height          =   375
            Index           =   2
            Left            =   1920
            TabIndex        =   59
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton cmdAttributeOkay 
            Caption         =   "Okay"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   60
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblWarpDir 
            Caption         =   "Dir: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblWarpMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame fraConvoTile 
         Caption         =   "Warp Properties"
         Height          =   1575
         Left            =   1680
         TabIndex        =   82
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton cmdAttributeCancel 
            Caption         =   "Cancel"
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   84
            Top             =   960
            Width           =   1695
         End
         Begin VB.HScrollBar scrlConvoTileNum 
            Height          =   255
            Left            =   240
            Max             =   0
            TabIndex        =   83
            Top             =   600
            Width           =   3375
         End
         Begin VB.CommandButton cmdAttributeOkay 
            Caption         =   "Okay"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   85
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblConvoTileNum 
            Caption         =   "Conversation: None"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn Properties"
         Height          =   1695
         Left            =   1680
         TabIndex        =   43
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         Begin VB.ComboBox cmbNpcSpawnDir 
            Height          =   315
            ItemData        =   "frmEditor_Map.frx":0085
            Left            =   1080
            List            =   "frmEditor_Map.frx":0095
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   720
            Width           =   2535
         End
         Begin VB.CommandButton cmdAttributeCancel 
            Caption         =   "Cancel"
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   47
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdAttributeOkay 
            Caption         =   "Okay"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   46
            Top             =   1080
            Width           =   1695
         End
         Begin VB.ComboBox cmbNpcSpawn 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Direction:"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Npc:"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox picTileset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   240
      ScaleHeight     =   382
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   32
      Top             =   360
      Width           =   4800
   End
   Begin VB.Frame fraTiles 
      Height          =   6375
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   5295
      Begin VB.HScrollBar scrlTileX 
         Height          =   255
         Left            =   120
         Max             =   0
         TabIndex        =   34
         Top             =   6000
         Width           =   4815
      End
      Begin VB.VScrollBar scrlTileY 
         Height          =   5775
         Left            =   4920
         Max             =   0
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdExpand 
         Height          =   255
         Left            =   4920
         TabIndex        =   31
         Top             =   6000
         Width           =   255
      End
   End
   Begin VB.Frame fraTileset 
      Caption         =   "Tileset - 1"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   6600
      Width           =   5295
      Begin VB.HScrollBar scrlTileset 
         Height          =   255
         Left            =   120
         Max             =   1
         Min             =   1
         TabIndex        =   10
         Top             =   360
         Value           =   1
         Width           =   5055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   5520
      TabIndex        =   1
      Top             =   6600
      Width           =   1815
      Begin VB.OptionButton optType 
         Caption         =   "Attribute"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "Layers"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   6375
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton optLayer 
         Caption         =   "Lights"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   70
         Top             =   1440
         Width           =   1335
      End
      Begin VB.HScrollBar scrlMapAnim 
         Height          =   255
         Left            =   120
         Max             =   0
         TabIndex        =   53
         Top             =   3720
         Width           =   1575
      End
      Begin VB.HScrollBar scrlFreq 
         Height          =   255
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   51
         Top             =   4320
         Value           =   1
         Width           =   1575
      End
      Begin VB.CommandButton cmdRandom 
         Caption         =   "Random"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CheckBox chkAnimated 
         Caption         =   "Animated"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmdLayerClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   5880
         Width           =   1575
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe 2"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask 2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdLayerFill 
         Caption         =   "Fill"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label lblMapAnim 
         Caption         =   "Map Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lblFreq 
         Caption         =   "Freq: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   4080
         Width           =   1575
      End
   End
   Begin VB.Frame fraAttributes 
      Caption         =   "Attributes"
      Height          =   6375
      Left            =   5520
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton optAttribute 
         Caption         =   "Fish Spot"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   105
         Top             =   3120
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Warp Check"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   104
         Top             =   2880
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Checkpoint"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   92
         Top             =   2640
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Slide"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   89
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Convo Tile"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   81
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Poke Storage"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   80
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Inv Storage"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   79
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Both Storage"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   78
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Heal"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Warp"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Npc Avoid"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearAttribute 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton cmdFillAttribute 
         Caption         =   "Fill"
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   5520
         Width           =   1575
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Npc Spawn"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optAttribute 
         Caption         =   "Blocked"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel(Esc)"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAnimated_Click()
    IsAnimated = chkAnimated.Value
End Sub

Private Sub cmdAttributeCancel_Click(Index As Integer)
    ClearMapAttribute
End Sub

Private Sub cmdAttributeOkay_Click(Index As Integer)
    Select Case Index
        Case 1 '//NpcSpawn
            EditorData1 = cmbNpcSpawn.ListIndex + 1
            EditorData2 = cmbNpcSpawnDir.ListIndex
        Case 2 '//Warp
            EditorData1 = scrlWarpMap.Value
            EditorData2 = scrlWarpX.Value
            EditorData3 = scrlWarpY.Value
            EditorData4 = scrlWarpDir.Value
        Case 3 '//ConvoTile
            EditorData1 = scrlConvoTileNum.Value
        Case 4 '//Checkpoint
            EditorData1 = scrlCheckMap.Value
            EditorData2 = scrlCheckX.Value
            EditorData3 = scrlCheckY.Value
            EditorData4 = scrlCheckDir.Value
    End Select
    
    fraAttribute.Visible = False
End Sub

Private Sub cmdClearAttribute_Click()
    MapEditorClearAttribute
End Sub

Private Sub cmdFillAttribute_Click()
    MapEditorFillAttribute
End Sub

Private Sub cmdRandom_Click()
    RandomPlaceLayer scrlFreq.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseMapEditor True
    End If
End Sub

Private Sub lstMapNpc_Click()
    cmbMapNpc.ListIndex = EditorTmpNpc(lstMapNpc.ListIndex + 1)
End Sub

Private Sub cmbMapNpc_Click()
Dim xIndex As Long
Dim tmpIndex As Long
    
    xIndex = lstMapNpc.ListIndex + 1
    tmpIndex = lstMapNpc.ListIndex
    If xIndex <= 0 Then Exit Sub
    EditorTmpNpc(xIndex) = cmbMapNpc.ListIndex
    lstMapNpc.RemoveItem xIndex - 1
    If EditorTmpNpc(xIndex) > 0 Then
        lstMapNpc.AddItem xIndex & ": " & Npc(EditorTmpNpc(xIndex)).Name, xIndex - 1
    Else
        lstMapNpc.AddItem xIndex & ": None", xIndex - 1
    End If
    lstMapNpc.ListIndex = tmpIndex
End Sub

Private Sub cmdExpand_Click()
    TileExpand = Not TileExpand
    
    If TileExpand Then
        picTileset.Width = 448
        picTileset.Height = 448
        scrlTileY.Height = 6735
        scrlTileX.Width = 6735
        scrlTileY.Left = 6840
        scrlTileX.top = 6960
        cmdExpand.Left = 6840
        cmdExpand.top = 6960
        fraTiles.Width = 481
        fraTiles.Height = 489
        
        EditorTileX = 0
        EditorTileY = 0
        EditorTileWidth = 1
        EditorTileHeight = 1
    Else
        picTileset.Width = 320
        picTileset.Height = 384
        scrlTileY.Height = 5775
        scrlTileX.Width = 4815
        scrlTileY.Left = 4920
        scrlTileX.top = 6000
        cmdExpand.Left = 4920
        cmdExpand.top = 6000
        fraTiles.Width = 353
        fraTiles.Height = 425
        
        EditorTileX = 0
        EditorTileY = 0
        EditorTileWidth = 1
        EditorTileHeight = 1
    End If
End Sub

Private Sub cmdLayerClear_Click()
    MapEditorClearLayer
End Sub

Private Sub cmdLayerFill_Click()
    MapEditorFillLayer
End Sub

Private Sub cmdPropertiesCancel_Click()
    mnuData.Visible = True
    Me.Height = 8340
    fraProperties.Visible = False
End Sub

Private Sub cmdPropertiesSave_Click()
Dim X As Long, x2 As Long
Dim Y As Long, Y2 As Long
Dim tempArr() As TileRec
Dim i As Long
    '//Input Data
    
    '//General
    Map.Name = Trim$(txtName.Text)
    
    '//Map Link
    If IsNumeric(txtLinkUp.Text) Then Map.LinkUp = Val(txtLinkUp.Text)
    If IsNumeric(txtLinkDown.Text) Then Map.LinkDown = Val(txtLinkDown.Text)
    If IsNumeric(txtLinkLeft.Text) Then Map.LinkLeft = Val(txtLinkLeft.Text)
    If IsNumeric(txtLinkRight.Text) Then Map.LinkRight = Val(txtLinkRight.Text)
    
    '//Map Size
    If Not IsNumeric(txtMaxX.Text) Then txtMaxX.Text = Map.MaxX
    If Val(txtMaxX.Text) < MAX_MAPX Then txtMaxX.Text = MAX_MAPX
    If Val(txtMaxX.Text) > 80 Then txtMaxX.Text = 50
    If Not IsNumeric(txtMaxY.Text) Then txtMaxY.Text = Map.MaxY
    If Val(txtMaxY.Text) < MAX_MAPY Then txtMaxY.Text = MAX_MAPY
    If Val(txtMaxY.Text) > 80 Then txtMaxY.Text = 50
    
    '//set the data before changing it
    tempArr = Map.Tile
    x2 = Map.MaxX
    Y2 = Map.MaxY
    
    Map.MaxX = Val(txtMaxX.Text)
    Map.MaxY = Val(txtMaxY.Text)
    
    If x2 > Map.MaxX Then x2 = Map.MaxX
    If Y2 > Map.MaxY Then Y2 = Map.MaxY
    
    '//redim the map size
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To x2
        For Y = 0 To Y2
            Map.Tile(X, Y) = tempArr(X, Y)
        Next
    Next
    
    '//Music
    If cmbMusic.ListIndex >= 0 Then
        Map.Music = cmbMusic.List(cmbMusic.ListIndex)
    Else
        Map.Music = "None."
    End If
    
    For i = 1 To MAX_MAP_NPC
        '//Npc
        Map.Npc(i) = EditorTmpNpc(i)
    Next
    
    '//Moral
    Map.Moral = cmbMoral.ListIndex
    Map.KillPlayer = chkKillPlayer.Value
    Map.IsCave = chkCave.Value
    Map.CaveLight = scrlCaveLight.Value
    Map.SpriteType = scrlSpriteType.Value
    Map.StartWeather = cmbWeather.ListIndex
    
    Map.NoCure = chkNoCure.Value
    
    '//Hide properties
    mnuData.Visible = True
    Me.Height = 8340
    fraProperties.Visible = False
End Sub

Private Sub Form_Load()
    Me.Width = 7515
    Me.Height = 8340
    fraProperties.Left = 8
    fraProperties.top = 8
    fraAttribute.top = 8
    fraAttribute.Left = 8
    
    scrlMapAnim.max = Count_MapAnim
    scrlWarpMap.max = MAX_MAP
    scrlCheckMap.max = MAX_MAP
    scrlConvoTileNum.max = MAX_CONVERSATION
    
    txtName.MaxLength = NAME_LENGTH
End Sub

Private Sub mnuCancel_Click()
    CloseMapEditor True
End Sub

Private Sub mnuProperties_Click()
Dim i As Long

    '//General
    txtName.Text = Trim$(Map.Name)
    
    '//Map Link
    txtLinkUp.Text = Str(Map.LinkUp)
    txtLinkDown.Text = Str(Map.LinkDown)
    txtLinkLeft.Text = Str(Map.LinkLeft)
    txtLinkRight.Text = Str(Map.LinkRight)
    
    '//Map Size
    txtMaxX.Text = Str(Map.MaxX)
    txtMaxY.Text = Str(Map.MaxY)
    
    '//Music
    cmbMusic.Clear
    cmbMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        cmbMusic.AddItem Trim$(musicCache(i))
    Next
    
    '//find the music we have set
    If cmbMusic.ListCount >= 0 Then
        cmbMusic.ListIndex = 0
        For i = 0 To cmbMusic.ListCount
            If Trim$(cmbMusic.List(i)) = Trim$(Map.Music) Then
                cmbMusic.ListIndex = i
            End If
        Next
    End If
    
    '//Npc
    cmbMapNpc.Clear
    cmbMapNpc.AddItem "None"
    For i = 1 To MAX_NPC
        cmbMapNpc.AddItem i & ": " & Trim$(Npc(i).Name)
    Next

    '//Npc
    lstMapNpc.Clear

    For i = 1 To MAX_MAP_NPC
        '//Npc
        EditorTmpNpc(i) = Map.Npc(i)
        
        If EditorTmpNpc(i) > 0 Then
            lstMapNpc.AddItem i & ": " & Trim$(Npc(EditorTmpNpc(i)).Name)
        Else
            lstMapNpc.AddItem i & ": None"
        End If
    Next
    
    '//Npc
    lstMapNpc.ListIndex = 0
    cmbMapNpc.ListIndex = Map.Npc(lstMapNpc.ListIndex + 1)
    
    '//Moral
    cmbMoral.ListIndex = Map.Moral
    chkKillPlayer.Value = Map.KillPlayer
    chkCave.Value = Map.IsCave
    scrlCaveLight.Value = Map.CaveLight
    scrlSpriteType.Value = Map.SpriteType
    cmbWeather.ListIndex = Map.StartWeather
    
    chkNoCure.Value = Map.NoCure

    '//Init
    mnuData.Visible = False
    Me.Height = 8040
    fraProperties.Visible = True
End Sub

Private Sub mnuSave_Click()
    MapEditorSend
End Sub

Private Sub optAttribute_Click(Index As Integer)
    CurAttribute = Index
    
    ClearMapAttribute
    
    Select Case Index
        Case MapAttribute.NpcSpawn
            fraAttribute.Visible = True
            fraNpcSpawn.Visible = True
            
            cmbNpcSpawn.ListIndex = 0
            cmbNpcSpawnDir.ListIndex = 0
        Case MapAttribute.Warp
            fraAttribute.Visible = True
            fraWarp.Visible = True
            
            scrlWarpMap.Value = 0
            scrlWarpX.Value = 0
            scrlWarpY.Value = 0
            scrlWarpDir.Value = 0
        Case MapAttribute.ConvoTile
            fraAttribute.Visible = True
            fraConvoTile.Visible = True
            
            scrlConvoTileNum.Value = 0
        Case MapAttribute.Checkpoint
            fraAttribute.Visible = True
            fraCheckpoint.Visible = True
            
            scrlCheckMap.Value = 0
            scrlCheckX.Value = 0
            scrlCheckY.Value = 0
            scrlCheckDir.Value = 0
    End Select
End Sub

Private Sub optLayer_Click(Index As Integer)
    CurLayer = Index
End Sub

Private Sub optType_Click(Index As Integer)
    fraLayers.Visible = False
    fraAttributes.Visible = False
    
    Select Case Index
        Case 1: fraLayers.Visible = True
        Case 2: fraAttributes.Visible = True
    End Select
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MapEditorChooseTile(Button, X, Y)
End Sub

Private Sub picTileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MapEditorChooseTile(Button, X, Y, True)
End Sub

Private Sub scrlCaveLight_Change()
    lblCaveLight.Caption = "Cave Light: " & scrlCaveLight.Value
End Sub

Private Sub scrlConvoTileNum_Change()
    If scrlConvoTileNum.Value > 0 Then
        lblConvoTileNum.Caption = "Conversation: " & Trim$(Conversation(scrlConvoTileNum.Value).Name)
    Else
        lblConvoTileNum.Caption = "Conversation: None"
    End If
End Sub

Private Sub scrlFreq_Change()
    lblFreq.Caption = "Freq: " & scrlFreq.Value
End Sub

Private Sub scrlMapAnim_Change()
    If scrlMapAnim.Value > 0 Then
        lblMapAnim.Caption = "Map Anim: " & scrlMapAnim.Value
    Else
        lblMapAnim.Caption = "Map Anim: None"
    End If
    editorMapAnim = scrlMapAnim.Value
End Sub

Private Sub scrlSpriteType_Change()
    Select Case scrlSpriteType.Value
        Case TEMP_SPRITE_GROUP_DIVE
            lblSpriteType.Caption = "Sprite Type: Dive"
        Case TEMP_SPRITE_GROUP_BIKE
            lblSpriteType.Caption = "Sprite Type: Bike"
        Case TEMP_SPRITE_GROUP_SURF
            lblSpriteType.Caption = "Sprite Type: Surf"
        Case TEMP_SPRITE_GROUP_MOUNT
            lblSpriteType.Caption = "Sprite Type: Mount"
        Case TEMP_FISH_MODE
            lblSpriteType.Caption = "Sprite Type: Fish"
        Case Else
            lblSpriteType.Caption = "Sprite Type: None"
    End Select
End Sub

Private Sub scrlTileset_Change()
    fraTileset.Caption = "Tileset - " & scrlTileset.Value
    LoadTileset scrlTileset.Value
End Sub

Private Sub scrlTileY_Change()
    EditorScrollY = scrlTileY.Value
End Sub

Private Sub scrlTileY_Scroll()
    EditorScrollY = scrlTileY.Value
End Sub

Private Sub scrlTileX_Change()
    EditorScrollX = scrlTileX.Value
End Sub

Private Sub scrlTileX_Scroll()
    EditorScrollX = scrlTileX.Value
End Sub

Private Sub scrlWarpDir_Change()
    Select Case scrlWarpDir.Value
        Case DIR_UP: lblWarpDir.Caption = "Dir: Up"
        Case DIR_DOWN: lblWarpDir.Caption = "Dir: Down"
        Case DIR_LEFT: lblWarpDir.Caption = "Dir: Left"
        Case DIR_RIGHT: lblWarpDir.Caption = "Dir: Right"
    End Select
End Sub

Private Sub scrlWarpMap_Change()
    lblWarpMap.Caption = "Map: " & scrlWarpMap.Value
End Sub

Private Sub scrlWarpX_Change()
    lblWarpX.Caption = "X: " & scrlWarpX.Value
End Sub

Private Sub scrlWarpY_Change()
    lblWarpY.Caption = "Y: " & scrlWarpY.Value
End Sub

Private Sub scrlCheckDir_Change()
    Select Case scrlCheckDir.Value
        Case DIR_UP: lblCheckDir.Caption = "Dir: Up"
        Case DIR_DOWN: lblCheckDir.Caption = "Dir: Down"
        Case DIR_LEFT: lblCheckDir.Caption = "Dir: Left"
        Case DIR_RIGHT: lblCheckDir.Caption = "Dir: Right"
    End Select
End Sub

Private Sub scrlCheckMap_Change()
    lblCheckMap.Caption = "Map: " & scrlCheckMap.Value
End Sub

Private Sub scrlCheckX_Change()
    lblCheckX.Caption = "X: " & scrlCheckX.Value
End Sub

Private Sub scrlCheckY_Change()
    lblCheckY.Caption = "Y: " & scrlCheckY.Value
End Sub

