VERSION 5.00
Begin VB.Form frmEditor_Move 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Move Editor"
   ClientHeight    =   8895
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   11790
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   8775
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   7860
         ItemData        =   "frmEditor_Move.frx":0000
         Left            =   120
         List            =   "frmEditor_Move.frx":0002
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   8775
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.ComboBox cmbSelfStatusReq 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":0004
         Left            =   6240
         List            =   "frmEditor_Move.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CheckBox chkProtect 
         Alignment       =   1  'Right Justify
         Caption         =   "Protect"
         Height          =   255
         Left            =   6720
         TabIndex        =   76
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cmbReflectType 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":004B
         Left            =   6240
         List            =   "frmEditor_Move.frx":0058
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CheckBox chkStatusToSelf 
         Caption         =   "Self?"
         Height          =   255
         Left            =   3480
         TabIndex        =   73
         Top             =   7920
         Width           =   1335
      End
      Begin VB.ComboBox cmbDecreaseWeather 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":0075
         Left            =   6480
         List            =   "frmEditor_Move.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox cmbStatusReq 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":00B7
         Left            =   6240
         List            =   "frmEditor_Move.frx":00D0
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ComboBox cmbBoostWeather 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":010A
         Left            =   6480
         List            =   "frmEditor_Move.frx":0120
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cmbWeather 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":014C
         Left            =   6240
         List            =   "frmEditor_Move.frx":0165
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtAbsorbDamage 
         Height          =   285
         Left            =   6240
         TabIndex        =   63
         Text            =   "0"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtRecoilDamage 
         Height          =   285
         Left            =   6240
         TabIndex        =   61
         Text            =   "0"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtStatusChance 
         Height          =   285
         Left            =   1560
         TabIndex        =   56
         Text            =   "0"
         Top             =   8280
         Width           =   2535
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":0198
         Left            =   1560
         List            =   "frmEditor_Move.frx":01B4
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   7920
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Buff / Debuff"
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   6120
         Width           =   4695
         Begin VB.TextBox txtBuffDebuff 
            Height          =   285
            Index           =   6
            Left            =   3600
            TabIndex        =   46
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtBuffDebuff 
            Height          =   285
            Index           =   5
            Left            =   2160
            TabIndex        =   45
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtBuffDebuff 
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   44
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtBuffDebuff 
            Height          =   285
            Index           =   3
            Left            =   3600
            TabIndex        =   43
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtBuffDebuff 
            Height          =   285
            Index           =   2
            Left            =   2160
            TabIndex        =   42
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtBuffDebuff 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   41
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Spd"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   52
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "SpDef"
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   51
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "SpAtk"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Def"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   49
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Atk"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   48
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "HP"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CheckBox chkPlaySelf 
         Caption         =   "Play on self?"
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   3720
         Width           =   2655
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":01F8
         Left            =   1560
         List            =   "frmEditor_Move.frx":01FA
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   7560
         Width           =   3135
      End
      Begin VB.TextBox txtAmountOfAttack 
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Text            =   "0"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtCooldown 
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Text            =   "0"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtCastTime 
         Height          =   285
         Left            =   3840
         TabIndex        =   32
         Text            =   "0"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtDuration 
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Text            =   "0"
         Top             =   7200
         Width           =   975
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2760
         Max             =   0
         TabIndex        =   28
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   3720
         TabIndex        =   26
         Text            =   "0"
         Top             =   7200
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Target Type"
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   4560
         Width           =   4695
         Begin VB.OptionButton optTargetType 
            Caption         =   "Spray"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   60
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optTargetType 
            Caption         =   "Self"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTargetType 
            Caption         =   "AoE"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optTargetType 
            Caption         =   "Linear"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbAttackType 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":01FC
         Left            =   1560
         List            =   "frmEditor_Move.frx":020C
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4080
         Width           =   3135
      End
      Begin VB.TextBox txtDescription 
         Height          =   645
         Left            =   1560
         MaxLength       =   150
         TabIndex        =   18
         Top             =   2280
         Width           =   3135
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1560
         Max             =   10
         TabIndex        =   15
         Top             =   5520
         Width           =   3135
      End
      Begin VB.TextBox txtPower 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Text            =   "0"
         Top             =   5880
         Width           =   3135
      End
      Begin VB.TextBox txtMaxPP 
         Height          =   285
         Left            =   3840
         TabIndex        =   12
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtPP 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":0231
         Left            =   1560
         List            =   "frmEditor_Move.frx":0241
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmEditor_Move.frx":0269
         Left            =   1560
         List            =   "frmEditor_Move.frx":02A6
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label26 
         Caption         =   "Self Status:"
         Height          =   255
         Left            =   4920
         TabIndex        =   78
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label25 
         Caption         =   "Reflect Type:"
         Height          =   255
         Left            =   4920
         TabIndex        =   74
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "Decrease Weather:"
         Height          =   255
         Left            =   4920
         TabIndex        =   72
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Status:"
         Height          =   255
         Left            =   4920
         TabIndex        =   70
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "Boost Weather:"
         Height          =   255
         Left            =   4920
         TabIndex        =   68
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Change Weather:"
         Height          =   255
         Left            =   4920
         TabIndex        =   66
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Absorb Damage:"
         Height          =   255
         Left            =   4920
         TabIndex        =   64
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "Recoil Damage:"
         Height          =   255
         Left            =   4920
         TabIndex        =   62
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "%"
         Height          =   255
         Left            =   4200
         TabIndex        =   57
         Top             =   8280
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Status Chance:"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   8280
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Status:"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   7920
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   7560
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Amount of Attack:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "Cooldown:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Cast Time:"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Duration:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Animation: None"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "Interval:"
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Attack Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblRange 
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Power:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Max PP:"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "PP:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Catergory:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
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
Attribute VB_Name = "frmEditor_Move"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPlaySelf_Click()
    PokemonMove(EditorIndex).SelfAnim = chkPlaySelf.value
    EditorChange = True
End Sub

Private Sub chkProtect_Click()
    PokemonMove(EditorIndex).CastProtect = chkProtect.value
    EditorChange = True
End Sub

Private Sub chkStatusToSelf_Click()
    PokemonMove(EditorIndex).StatusToSelf = chkStatusToSelf.value
    EditorChange = True
End Sub

Private Sub cmbAttackType_Click()
    PokemonMove(EditorIndex).AttackType = cmbAttackType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbBoostWeather_Change()
    PokemonMove(EditorIndex).BoostWeather = cmbBoostWeather.ListIndex
    EditorChange = True
End Sub

Private Sub cmbCategory_Click()
    PokemonMove(EditorIndex).Category = cmbCategory.ListIndex
    EditorChange = True
End Sub

Private Sub cmbDecreaseWeather_Click()
    PokemonMove(EditorIndex).DecreaseWeather = cmbDecreaseWeather.ListIndex
    EditorChange = True
End Sub

Private Sub cmbReflectType_Click()
    PokemonMove(EditorIndex).ReflectType = cmbReflectType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbSelfStatusReq_Click()
    PokemonMove(EditorIndex).SelfStatusReq = cmbSelfStatusReq.ListIndex
    EditorChange = True
End Sub

Private Sub cmbSound_Click()
    If EditorStart = True Then Exit Sub
    '//Sound
    If cmbSound.ListIndex >= 0 Then
        PokemonMove(EditorIndex).Sound = Trim$(cmbSound.List(cmbSound.ListIndex))
    Else
        PokemonMove(EditorIndex).Sound = "None."
    End If
    EditorChange = True
End Sub

Private Sub cmbStatus_Click()
    PokemonMove(EditorIndex).pStatus = cmbStatus.ListIndex
    EditorChange = True
End Sub

Private Sub cmbStatusReq_Click()
    PokemonMove(EditorIndex).StatusReq = cmbStatusReq.ListIndex
    EditorChange = True
End Sub

Private Sub cmbType_Click()
    PokemonMove(EditorIndex).Type = cmbType.ListIndex
    EditorChange = True
End Sub

Private Sub cmbWeather_Click()
    PokemonMove(EditorIndex).ChangeWeather = cmbWeather.ListIndex
    EditorChange = True
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
        cuBound = MAX_POKEMON_MOVE
        
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
        ClosePokemonMoveEditor
    End If
End Sub

Private Sub Form_Load()
    txtName.MaxLength = NAME_LENGTH
    scrlAnimation.max = MAX_ANIMATION
End Sub

Private Sub lstIndex_Click()
    PokemonMoveEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestPokemonMove
    End If
    ClosePokemonMoveEditor
End Sub

Private Sub mnuExit_Click()
    ClosePokemonMoveEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_POKEMON_MOVE
        If PokemonMoveChange(i) Then
            SendSavePokemonMove i
            PokemonMoveChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'ClosePokemonMoveEditor
End Sub

Private Sub optTargetType_Click(Index As Integer)
    If optTargetType(Index).value = True Then
        PokemonMove(EditorIndex).targetType = Index
    End If
End Sub

Private Sub scrlAnimation_Change()
    If scrlAnimation.value > 0 Then
        lblAnimation.Caption = "Animation: #" & scrlAnimation.value & " " & Trim$(Animation(scrlAnimation.value).Name)
    Else
        lblAnimation.Caption = "Animation: None"
    End If
    PokemonMove(EditorIndex).Animation = scrlAnimation.value
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Range: " & scrlRange.value
    PokemonMove(EditorIndex).Range = scrlRange.value
    EditorChange = True
End Sub

Private Sub txtAbsorbDamage_Change()
    If IsNumeric(txtAbsorbDamage.Text) Then
        PokemonMove(EditorIndex).AbsorbDamage = Val(txtAbsorbDamage.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtAmountOfAttack_Change()
    If IsNumeric(txtAmountOfAttack.Text) Then
        PokemonMove(EditorIndex).AmountOfAttack = Val(txtAmountOfAttack.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtBuffDebuff_Change(Index As Integer)
    If IsNumeric(txtBuffDebuff(Index).Text) Then
        PokemonMove(EditorIndex).dStat(Index) = Val(txtBuffDebuff(Index).Text)
        EditorChange = True
    End If
End Sub

Private Sub txtCastTime_Change()
    If IsNumeric(txtCastTime.Text) Then
        PokemonMove(EditorIndex).CastTime = Val(txtCastTime.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtCooldown_Change()
    If IsNumeric(txtCooldown.Text) Then
        PokemonMove(EditorIndex).Cooldown = Val(txtCooldown.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtDescription_Change()
    PokemonMove(EditorIndex).Description = Trim$(txtDescription.Text)
    EditorChange = True
End Sub

Private Sub txtDuration_Change()
    If IsNumeric(txtDuration.Text) Then
        PokemonMove(EditorIndex).Duration = Val(txtDuration.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtInterval_Change()
    If IsNumeric(txtInterval.Text) Then
        PokemonMove(EditorIndex).Interval = Val(txtInterval.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    PokemonMove(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & PokemonMove(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtPower_Change()
    If IsNumeric(txtPower.Text) Then
        PokemonMove(EditorIndex).Power = Val(txtPower.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtPP_Change()
    If IsNumeric(txtPP.Text) Then
        PokemonMove(EditorIndex).PP = Val(txtPP.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtMaxPP_Change()
    If IsNumeric(txtMaxPP.Text) Then
        PokemonMove(EditorIndex).MaxPP = Val(txtMaxPP.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtRecoilDamage_Change()
    If IsNumeric(txtRecoilDamage.Text) Then
        PokemonMove(EditorIndex).RecoilDamage = Val(txtRecoilDamage.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtStatusChance_Change()
    If IsNumeric(txtStatusChance.Text) Then
        PokemonMove(EditorIndex).pStatusChance = Val(txtStatusChance.Text)
        EditorChange = True
    End If
End Sub
