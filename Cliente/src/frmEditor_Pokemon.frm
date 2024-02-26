VERSION 5.00
Begin VB.Form frmEditor_Pokemon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokemon Editor"
   ClientHeight    =   9120
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14025
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   608
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   935
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   9015
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   79
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   8250
         ItemData        =   "frmEditor_Pokemon.frx":0000
         Left            =   120
         List            =   "frmEditor_Pokemon.frx":0002
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   9015
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.CheckBox chkLendary 
         Caption         =   "Lendary?"
         Height          =   255
         Left            =   4920
         TabIndex        =   99
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   1440
         Width           =   3495
      End
      Begin VB.VScrollBar scrlOffSetY 
         Height          =   1215
         Left            =   4800
         Max             =   200
         Min             =   -200
         TabIndex        =   96
         Top             =   120
         Width           =   135
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   2880
         TabIndex        =   92
         Top             =   720
         Width           =   735
      End
      Begin VB.Frame Frame8 
         Caption         =   "Item Moveset"
         Height          =   2055
         Left            =   6120
         TabIndex        =   86
         Top             =   4800
         Width           =   4575
         Begin VB.ListBox lstItemMoveset 
            Height          =   840
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   4335
         End
         Begin VB.ComboBox cmbItemMove 
            Height          =   315
            ItemData        =   "frmEditor_Pokemon.frx":0004
            Left            =   1800
            List            =   "frmEditor_Pokemon.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CommandButton cmdItemMoveFind 
            Caption         =   "Find"
            Height          =   255
            Left            =   3480
            TabIndex        =   88
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtItemMoveFind 
            Height          =   285
            Left            =   1800
            TabIndex        =   87
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label28 
            Caption         =   "Move:"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1560
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Item Drop"
         Height          =   1815
         Left            =   6120
         TabIndex        =   80
         Top             =   7080
         Width           =   4575
         Begin VB.TextBox txtItemSearch 
            Height          =   285
            Left            =   1200
            TabIndex        =   94
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtItemDropRate 
            Height          =   285
            Left            =   1800
            TabIndex        =   85
            Text            =   "0"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.ComboBox cmbItemNum 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   1080
            Width           =   2655
         End
         Begin VB.ListBox lstItemDrop 
            Height          =   840
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label27 
            Caption         =   "Rarity 0% - 100%"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   1440
            Width           =   4335
         End
         Begin VB.Label Label26 
            Caption         =   "Item:"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.Frame fraEvolve 
         Caption         =   "Evolve - 1"
         Height          =   1815
         Left            =   120
         TabIndex        =   62
         Top             =   5400
         Width           =   5775
         Begin VB.TextBox txtSearch 
            Height          =   285
            Left            =   4680
            TabIndex        =   93
            Top             =   720
            Width           =   855
         End
         Begin VB.HScrollBar scrlEvolveCondition 
            Height          =   255
            Left            =   3000
            Max             =   0
            TabIndex        =   70
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtEvolveLevel 
            Height          =   285
            Left            =   1440
            TabIndex        =   67
            Text            =   "0"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtEvolveConditionData 
            Height          =   285
            Left            =   1440
            TabIndex        =   66
            Text            =   "0"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.HScrollBar scrlEvolve 
            Height          =   255
            Left            =   2400
            Max             =   0
            TabIndex        =   64
            Top             =   720
            Width           =   2175
         End
         Begin VB.HScrollBar scrlEvolveIndex 
            Height          =   255
            Left            =   120
            Max             =   1
            Min             =   1
            TabIndex        =   63
            Top             =   240
            Value           =   1
            Width           =   5535
         End
         Begin VB.Label lblEvolveCondition 
            Caption         =   "Condition: None"
            Height          =   255
            Left            =   3000
            TabIndex        =   71
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label21 
            Caption         =   "Evolve Level:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Condition Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblEvolve 
            Caption         =   "Evolve To: None"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   5295
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5640
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1560
         Max             =   10
         TabIndex        =   60
         Top             =   8640
         Width           =   4335
      End
      Begin VB.TextBox txtSpecies 
         Height          =   285
         Left            =   1560
         TabIndex        =   58
         Top             =   8280
         Width           =   4335
      End
      Begin VB.TextBox txtPokedexEntry 
         Height          =   855
         Left            =   1560
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   7320
         Width           =   4335
      End
      Begin VB.Frame Frame6 
         Caption         =   "Egg Move"
         Height          =   2055
         Left            =   6120
         TabIndex        =   53
         Top             =   2520
         Width           =   4575
         Begin VB.TextBox txtEggMoveFind 
            Height          =   285
            Left            =   1800
            TabIndex        =   77
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdEggMoveFind 
            Caption         =   "Find"
            Height          =   255
            Left            =   3480
            TabIndex        =   76
            Top             =   1200
            Width           =   975
         End
         Begin VB.ComboBox cmbEggMoveNum 
            Height          =   315
            ItemData        =   "frmEditor_Pokemon.frx":0008
            Left            =   1800
            List            =   "frmEditor_Pokemon.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1560
            Width           =   2655
         End
         Begin VB.ListBox lstEggMove 
            Height          =   840
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label lblEggMoveNum 
            Caption         =   "Move:"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1560
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Moveset"
         Height          =   2175
         Left            =   6120
         TabIndex        =   49
         Top             =   240
         Width           =   4575
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find"
            Height          =   255
            Left            =   3480
            TabIndex        =   75
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtFind 
            Height          =   285
            Left            =   1800
            TabIndex        =   74
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox cmbMoveNum 
            Height          =   315
            ItemData        =   "frmEditor_Pokemon.frx":000C
            Left            =   1800
            List            =   "frmEditor_Pokemon.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtMoveLevel 
            Height          =   285
            Left            =   3600
            TabIndex        =   52
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.ListBox lstMoveset 
            Height          =   1035
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label lblMovesetNum 
            Caption         =   "Move: "
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1680
            Width           =   2175
         End
      End
      Begin VB.TextBox txtWeight 
         Height          =   285
         Left            =   4440
         TabIndex        =   48
         Text            =   "0"
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   4440
         TabIndex        =   45
         Text            =   "0"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.ComboBox cmbGrowthRate 
         Height          =   315
         ItemData        =   "frmEditor_Pokemon.frx":0010
         Left            =   1440
         List            =   "frmEditor_Pokemon.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtBaseExp 
         Height          =   285
         Left            =   1440
         TabIndex        =   42
         Text            =   "0"
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtEVYeildVal 
         Height          =   285
         Left            =   4440
         TabIndex        =   40
         Text            =   "0"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox cmbEVYeildType 
         Height          =   315
         ItemData        =   "frmEditor_Pokemon.frx":0064
         Left            =   4440
         List            =   "frmEditor_Pokemon.frx":007A
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   3960
         Width           =   1455
      End
      Begin VB.ComboBox cmbEggGroup 
         Height          =   315
         ItemData        =   "frmEditor_Pokemon.frx":009F
         Left            =   1440
         List            =   "frmEditor_Pokemon.frx":00D0
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtEggCycle 
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Text            =   "0"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtFemaleRate 
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Text            =   "0"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtCatchRate 
         Height          =   285
         Left            =   4320
         TabIndex        =   29
         Text            =   "0"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Type"
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   2775
         Begin VB.ComboBox cmbSecondaryType 
            Height          =   315
            ItemData        =   "frmEditor_Pokemon.frx":0153
            Left            =   1080
            List            =   "frmEditor_Pokemon.frx":0190
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox cmbPrimaryType 
            Height          =   315
            ItemData        =   "frmEditor_Pokemon.frx":021E
            Left            =   1080
            List            =   "frmEditor_Pokemon.frx":025B
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Secondary:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Primary:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Base Stats"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   5775
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   6
            Left            =   4800
            TabIndex        =   23
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   5
            Left            =   2760
            TabIndex        =   21
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   19
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   3
            Left            =   4800
            TabIndex        =   17
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   2
            Left            =   2760
            TabIndex        =   15
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Spd:"
            Height          =   255
            Left            =   4080
            TabIndex        =   22
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "SpDef:"
            Height          =   255
            Left            =   2040
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "SpAtk:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Def:"
            Height          =   255
            Left            =   4080
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Atk:"
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "HP:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox chkScale 
         Caption         =   "Scale"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   5040
         ScaleHeight     =   66.065
         ScaleMode       =   0  'User
         ScaleWidth      =   68.266
         TabIndex        =   9
         Top             =   480
         Width           =   960
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   0
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   315
         ItemData        =   "frmEditor_Pokemon.frx":02E9
         Left            =   1200
         List            =   "frmEditor_Pokemon.frx":02F9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label23 
         Caption         =   "Cries:"
         Height          =   255
         Left            =   240
         TabIndex        =   97
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblOffSetY 
         Caption         =   "Name OffSetY: 0"
         Height          =   255
         Left            =   3120
         TabIndex        =   95
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblRange 
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   8640
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "Species:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   8280
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Pokedex Entry:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   7320
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   6000
         X2              =   6000
         Y1              =   8880
         Y2              =   240
      End
      Begin VB.Label Label20 
         Caption         =   "Weight:"
         Height          =   255
         Left            =   3120
         TabIndex        =   47
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Height:"
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Growth Rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Base Exp:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "EV Yeild Val:"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "EV Yeild Type:"
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Egg Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Egg Cycle:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Female Rate:"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Catch Rate:"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Behaviour:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
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
Attribute VB_Name = "frmEditor_Pokemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MoveIndex As Long
Private EggMoveIndex As Long
Private ItemMoveIndex As Long
Private EvolveIndex As Long
Private ItemDropIndex As Long

Private Sub chkLendary_Click()
    Pokemon(EditorIndex).Lendary = chkLendary
    EditorChange = True
End Sub

Private Sub chkScale_Click()
    Pokemon(EditorIndex).ScaleSprite = chkScale.value
    EditorChange = True
End Sub

Private Sub cmbBehaviour_Click()
    Pokemon(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    EditorChange = True
End Sub

Private Sub cmbEggGroup_Click()
    Pokemon(EditorIndex).EggGroup = cmbEggGroup.ListIndex
    EditorChange = True
End Sub

Private Sub cmbEggMoveNum_Click()
Dim tmpIndex As Long

    If EggMoveIndex = 0 Then Exit Sub
    tmpIndex = lstEggMove.ListIndex
    lstEggMove.RemoveItem EggMoveIndex - 1
    Pokemon(EditorIndex).EggMoveset(EggMoveIndex) = cmbEggMoveNum.ListIndex
    If cmbEggMoveNum.ListIndex > 0 Then
        lstEggMove.AddItem (EggMoveIndex) & ": " & Trim$(PokemonMove(cmbEggMoveNum.ListIndex).Name), EggMoveIndex - 1
    Else
        lstEggMove.AddItem (EggMoveIndex) & ": None", EggMoveIndex - 1
    End If
    lstEggMove.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub cmbEVYeildType_Click()
    Pokemon(EditorIndex).EvYeildType = cmbEVYeildType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbGrowthRate_Click()
    Pokemon(EditorIndex).GrowthRate = cmbGrowthRate.ListIndex
    EditorChange = True
End Sub

Private Sub cmbItemMove_Click()
Dim tmpIndex As Long

    If ItemMoveIndex = 0 Then Exit Sub
    tmpIndex = lstItemMoveset.ListIndex
    lstItemMoveset.RemoveItem ItemMoveIndex - 1
    Pokemon(EditorIndex).ItemMoveset(ItemMoveIndex) = cmbItemMove.ListIndex
    If cmbItemMove.ListIndex > 0 Then
        lstItemMoveset.AddItem (ItemMoveIndex) & ": " & Trim$(PokemonMove(cmbItemMove.ListIndex).Name), ItemMoveIndex - 1
    Else
        lstItemMoveset.AddItem (ItemMoveIndex) & ": None", ItemMoveIndex - 1
    End If
    lstItemMoveset.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub cmbItemNum_Click()
Dim tmpIndex As Long

    If ItemDropIndex = 0 Then Exit Sub
    tmpIndex = lstItemDrop.ListIndex
    lstItemDrop.RemoveItem ItemDropIndex - 1
    Pokemon(EditorIndex).DropNum(ItemDropIndex) = cmbItemNum.ListIndex
    If cmbItemNum.ListIndex > 0 Then
        lstItemDrop.AddItem (ItemDropIndex) & ": " & Trim$(Item(cmbItemNum.ListIndex).Name) & " Rate:" & Pokemon(EditorIndex).DropRate(ItemDropIndex), ItemDropIndex - 1
    Else
        lstItemDrop.AddItem (ItemDropIndex) & ": None", ItemDropIndex - 1
    End If
    lstItemDrop.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub cmbMoveNum_Click()
Dim tmpIndex As Long

    If MoveIndex = 0 Then Exit Sub
    tmpIndex = lstMoveset.ListIndex
    lstMoveset.RemoveItem MoveIndex - 1
    Pokemon(EditorIndex).Moveset(MoveIndex).MoveNum = cmbMoveNum.ListIndex
    If cmbMoveNum.ListIndex > 0 Then
        lstMoveset.AddItem (MoveIndex) & ": " & Trim$(PokemonMove(cmbMoveNum.ListIndex).Name) & " Lv:" & Pokemon(EditorIndex).Moveset(MoveIndex).MoveLevel, MoveIndex - 1
    Else
        lstMoveset.AddItem (MoveIndex) & ": None", MoveIndex - 1
    End If
    lstMoveset.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub cmbPrimaryType_Click()
    Pokemon(EditorIndex).PrimaryType = cmbPrimaryType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbSecondaryType_Click()
    Pokemon(EditorIndex).SecondaryType = cmbSecondaryType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbSound_Click()
    If EditorStart = True Then Exit Sub
    '//Sound
    If cmbSound.ListIndex >= 0 Then
        Pokemon(EditorIndex).Sound = Trim$(cmbSound.List(cmbSound.ListIndex))
    Else
        Pokemon(EditorIndex).Sound = "None."
    End If
    EditorChange = True
End Sub

Private Sub cmdEggMoveFind_Click()
Dim FindChar As String
Dim clBound As Long, cuBound As Long
Dim i As Long
Dim ComboText As String
Dim indexString As String
Dim stringLength As Long

    If Len(Trim$(txtEggMoveFind.Text)) > 0 Then
        FindChar = Trim$(txtEggMoveFind.Text)
        clBound = 0
        cuBound = MAX_POKEMON_MOVE
        
        For i = clBound To cuBound
            If cmbEggMoveNum.List(i) <> "None" Then
                ComboText = Trim$(cmbEggMoveNum.List(i))
                indexString = i & ": "
                stringLength = Len(ComboText) - Len(indexString)
                If stringLength >= 0 Then
                    ComboText = Mid$(ComboText, Len(indexString) + 1, stringLength)
                    If LCase(ComboText) = LCase(FindChar) Then
                        cmbEggMoveNum.ListIndex = i
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        MsgBox "Index not found", vbCritical
    End If
End Sub

Private Sub cmdFind_Click()
Dim FindChar As String
Dim clBound As Long, cuBound As Long
Dim i As Long
Dim ComboText As String
Dim indexString As String
Dim stringLength As Long

    If Len(Trim$(txtFind.Text)) > 0 Then
        FindChar = Trim$(txtFind.Text)
        clBound = 0
        cuBound = MAX_POKEMON_MOVE
        
        For i = clBound To cuBound
            If cmbMoveNum.List(i) <> "None" Then
                ComboText = Trim$(cmbMoveNum.List(i))
                indexString = i & ": "
                stringLength = Len(ComboText) - Len(indexString)
                If stringLength >= 0 Then
                    ComboText = Mid$(ComboText, Len(indexString) + 1, stringLength)
                    If LCase(ComboText) = LCase(FindChar) Then
                        cmbMoveNum.ListIndex = i
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        MsgBox "Index not found", vbCritical
    End If
End Sub

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
        cuBound = MAX_POKEMON
        
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

Private Sub cmdItemMoveFind_Click()
Dim FindChar As String
Dim clBound As Long, cuBound As Long
Dim i As Long
Dim ComboText As String
Dim indexString As String
Dim stringLength As Long

    If Len(Trim$(txtItemMoveFind.Text)) > 0 Then
        FindChar = Trim$(txtItemMoveFind.Text)
        clBound = 0
        cuBound = MAX_POKEMON_MOVE
        
        For i = clBound To cuBound
            If cmbItemMove.List(i) <> "None" Then
                ComboText = Trim$(cmbItemMove.List(i))
                indexString = i & ": "
                stringLength = Len(ComboText) - Len(indexString)
                If stringLength >= 0 Then
                    ComboText = Mid$(ComboText, Len(indexString) + 1, stringLength)
                    If LCase(ComboText) = LCase(FindChar) Then
                        cmbItemMove.ListIndex = i
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        MsgBox "Index not found", vbCritical
    End If
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    
    For i = 1 To MAX_POKEMON
        
    Next i
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        ClosePokemonEditor
    End If
End Sub

Private Sub Form_Load()
    scrlSprite.max = Count_Pokemon
    scrlEvolve.max = MAX_POKEMON
    scrlEvolveIndex.max = MAX_EVOLVE
    scrlEvolveCondition.max = MAX_EVOLVE_CONDT
    
    txtSpecies.MaxLength = NAME_LENGTH
    txtName.MaxLength = NAME_LENGTH
    
    '//Set Index
    MoveIndex = lstMoveset.ListIndex + 1
    EggMoveIndex = lstEggMove.ListIndex + 1
    ItemMoveIndex = lstItemMoveset.ListIndex + 1
    ItemDropIndex = lstItemDrop.ListIndex + 1
    EvolveIndex = scrlEvolveIndex.value
End Sub

Private Sub lstEggMove_Click()
    EggMoveIndex = lstEggMove.ListIndex
    If EggMoveIndex = 0 Then Exit Sub
    cmbEggMoveNum.ListIndex = Pokemon(EditorIndex).EggMoveset(EggMoveIndex)
End Sub

Private Sub lstIndex_Click()
    PokemonEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub lstItemDrop_Click()
    ItemDropIndex = lstItemDrop.ListIndex + 1
    If ItemDropIndex = 0 Then Exit Sub
    cmbItemNum.ListIndex = Pokemon(EditorIndex).DropNum(ItemDropIndex)
    txtItemDropRate.Text = Pokemon(EditorIndex).DropRate(ItemDropIndex)
End Sub

Private Sub lstItemMoveset_Click()
    ItemMoveIndex = lstItemMoveset.ListIndex + 1
    If ItemMoveIndex = 0 Then Exit Sub
    cmbItemMove.ListIndex = Pokemon(EditorIndex).ItemMoveset(ItemMoveIndex)
End Sub

Private Sub lstMoveset_Click()
    MoveIndex = lstMoveset.ListIndex + 1
    If MoveIndex = 0 Then Exit Sub
    cmbMoveNum.ListIndex = Pokemon(EditorIndex).Moveset(MoveIndex).MoveNum
    txtMoveLevel.Text = Pokemon(EditorIndex).Moveset(MoveIndex).MoveLevel
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestPokemon
    End If
    ClosePokemonEditor
End Sub

Private Sub mnuExit_Click()
    ClosePokemonEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_POKEMON
        If PokemonChange(i) Then
            SendSavePokemon i
            PokemonChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'ClosePokemonEditor
End Sub

Private Sub scrlEvolve_Change()
    If scrlEvolve.value > 0 Then
        lblEvolve.Caption = "Evolve To: " & Trim$(Pokemon(scrlEvolve.value).Name)
    Else
        lblEvolve.Caption = "Evolve To: None"
    End If
    Pokemon(EditorIndex).evolveNum(EvolveIndex) = scrlEvolve.value
    EditorChange = True
End Sub

Private Sub scrlEvolveCondition_Change()
    Select Case scrlEvolveCondition.value
        Case EVOLVE_CONDT_TIME: lblEvolveCondition.Caption = "Condition: Time"
        Case EVOLVE_CONDT_HAPPINESS: lblEvolveCondition.Caption = "Condition: Happiness"
        Case EVOLVE_CONDT_TRADE: lblEvolveCondition.Caption = "Condition: Trade"
        Case EVOLVE_CONDT_GENDER: lblEvolveCondition.Caption = "Condition: Gender"
        Case EVOLVE_CONDT_ITEM: lblEvolveCondition.Caption = "Condition: Item"
        Case EVOLVE_CONDT_KNOWMOVE: lblEvolveCondition.Caption = "Condition: Know Move"
        Case EVOLVE_CONDT_AREA: lblEvolveCondition.Caption = "Condition: Area"
        Case Else: lblEvolveCondition.Caption = "Condition: None"
    End Select
    Pokemon(EditorIndex).EvolveCondition(EvolveIndex) = scrlEvolveCondition.value
    EditorChange = True
End Sub

Private Sub scrlEvolveIndex_Change()
    fraEvolve.Caption = "Evolve - " & scrlEvolveIndex.value
    EvolveIndex = scrlEvolveIndex.value
    If EvolveIndex <= 0 Then Exit Sub
    
    '//Set Data
    scrlEvolve.value = Pokemon(EditorIndex).evolveNum(EvolveIndex)
    txtEvolveLevel.Text = Pokemon(EditorIndex).EvolveLevel(EvolveIndex)
    txtEvolveConditionData.Text = Pokemon(EditorIndex).EvolveConditionData(EvolveIndex)
    scrlEvolveCondition.value = Pokemon(EditorIndex).EvolveCondition(EvolveIndex)
End Sub

Private Sub scrlOffSetY_Change()
    Pokemon(EditorIndex).NameOffSetY = scrlOffSetY
    lblOffSetY = "Name OffSetY: " & scrlOffSetY
    EditorChange = True
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Range: " & scrlRange.value
    Pokemon(EditorIndex).Range = scrlRange.value
    EditorChange = True
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    Pokemon(EditorIndex).Sprite = scrlSprite.value
    EditorChange = True
End Sub

Private Sub txtBaseExp_Change()
    If IsNumeric(txtBaseExp.Text) Then
        Pokemon(EditorIndex).BaseExp = Val(txtBaseExp.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtBaseStat_Change(Index As Integer)
    If IsNumeric(txtBaseStat(Index).Text) Then
        Pokemon(EditorIndex).BaseStat(Index) = Val(txtBaseStat(Index).Text)
        EditorChange = True
    End If
End Sub

Private Sub txtCatchRate_Change()
    If IsNumeric(txtCatchRate.Text) Then
        Pokemon(EditorIndex).CatchRate = Val(txtCatchRate.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtEggCycle_Change()
    If IsNumeric(txtEggCycle.Text) Then
        Pokemon(EditorIndex).EggCycle = Val(txtEggCycle.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtEvolveConditionData_Change()
    If IsNumeric(txtEvolveConditionData.Text) Then
        Pokemon(EditorIndex).EvolveConditionData(EvolveIndex) = Val(txtEvolveConditionData.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtEvolveLevel_Change()
    If IsNumeric(txtEvolveLevel.Text) Then
        Pokemon(EditorIndex).EvolveLevel(EvolveIndex) = Val(txtEvolveLevel.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtEVYeildVal_Change()
    If IsNumeric(txtEVYeildVal.Text) Then
        Pokemon(EditorIndex).EvYeildVal = Val(txtEVYeildVal.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtFemaleRate_Change()
    If IsNumeric(txtFemaleRate.Text) Then
        Pokemon(EditorIndex).FemaleRate = Val(txtFemaleRate.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtHeight_Change()
    If IsNumeric(txtHeight.Text) Then
        Pokemon(EditorIndex).Height = Val(txtHeight.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtID_Change()
    If Not IsNumeric(txtID) Then
        txtID = 0
    End If
    If txtID < 0 Then
        txtID = 0
    End If
    If txtID > Count_Pokemon Then
        txtID = Count_Pokemon
    End If
    
    scrlSprite = txtID
End Sub

Private Sub txtItemDropRate_Change()
Dim tmpIndex As Long

    If ItemDropIndex = 0 Then Exit Sub
    If Not IsNumeric(txtItemDropRate.Text) Then Exit Sub
    If txtItemDropRate > 100 Or txtItemDropRate < 0 Then
        txtItemDropRate = 0
    End If
    
    tmpIndex = lstItemDrop.ListIndex
    lstItemDrop.RemoveItem ItemDropIndex - 1
    Pokemon(EditorIndex).DropRate(ItemDropIndex) = Val(txtItemDropRate.Text)
    If Pokemon(EditorIndex).DropNum(ItemDropIndex) > 0 Then
        lstItemDrop.AddItem (ItemDropIndex) & ": " & Trim$(Item(Pokemon(EditorIndex).DropNum(ItemDropIndex)).Name) & " Rate:" & Val(txtItemDropRate.Text), ItemDropIndex - 1
    Else
        lstItemDrop.AddItem (ItemDropIndex) & ": None", ItemDropIndex - 1
    End If
    lstItemDrop.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtItemSearch_Change()
    Dim Find As String, i As Long

    If Not IsNumeric(txtItemSearch) Then
        Find = UCase$(Trim$(txtItemSearch.Text))
        If Len(Find) <= 2 And Not Find = "" Then
            'lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_ITEM
            If Not Find = "" Then
                If InStr(1, UCase$(Trim$(Item(i).Name)), Find) > 0 Then
                    cmbItemNum.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
    Else
        If txtItemSearch > MAX_ITEM Then
            txtItemSearch = MAX_ITEM
        ElseIf txtItemSearch <= 0 Then
            txtItemSearch = 1
        End If
        cmbItemNum.ListIndex = txtItemSearch
    End If
End Sub

Private Sub txtMoveLevel_Change()
Dim tmpIndex As Long

    If MoveIndex = 0 Then Exit Sub
    If Not IsNumeric(txtMoveLevel.Text) Then Exit Sub
    tmpIndex = lstMoveset.ListIndex
    lstMoveset.RemoveItem MoveIndex - 1
    Pokemon(EditorIndex).Moveset(MoveIndex).MoveLevel = Val(txtMoveLevel.Text)
    If Pokemon(EditorIndex).Moveset(MoveIndex).MoveNum > 0 Then
        lstMoveset.AddItem (MoveIndex) & ": " & Trim$(PokemonMove(Pokemon(EditorIndex).Moveset(MoveIndex).MoveNum).Name) & " Lv:" & Val(txtMoveLevel.Text), MoveIndex - 1
    Else
        lstMoveset.AddItem (MoveIndex) & ": None", MoveIndex - 1
    End If
    lstMoveset.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Pokemon(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pokemon(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtPokedexEntry_Change()
    Pokemon(EditorIndex).PokeDexEntry = txtPokedexEntry.Text
    EditorChange = True
End Sub

Private Sub txtSearch_Change()
    Dim Find As String, i As Long

    If Not IsNumeric(txtSearch) Then
        Find = UCase$(Trim$(txtSearch.Text))
        If Len(Find) <= 2 And Not Find = "" Then
            'lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_POKEMON
            If Not Find = "" Then
                If InStr(1, UCase$(Trim$(Pokemon(i).Name)), Find) > 0 Then
                    scrlEvolve = i
                    Exit Sub
                End If
            End If
        Next
    Else
        If txtSearch > MAX_POKEMON Then
            txtSearch = MAX_POKEMON
        ElseIf txtSearch < 0 Then
            txtSearch = 0
        End If
        scrlEvolve = txtSearch
    End If
End Sub

Private Sub txtSpecies_Change()
    Pokemon(EditorIndex).Species = txtSpecies.Text
    EditorChange = True
End Sub

Private Sub txtWeight_Change()
    If IsNumeric(txtWeight.Text) Then
        Pokemon(EditorIndex).Weight = Val(txtWeight.Text)
        EditorChange = True
    End If
End Sub
