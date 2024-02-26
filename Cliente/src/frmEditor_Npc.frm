VERSION 5.00
Begin VB.Form frmEditor_Npc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Editor"
   ClientHeight    =   6900
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9210
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Spawn Day's"
      Height          =   1335
      Left            =   120
      TabIndex        =   43
      Top             =   5520
      Width           =   2895
      Begin VB.CheckBox chkWeekDay 
         Caption         =   "Sabado"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   50
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkWeekDay 
         Caption         =   "Sexta"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   49
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkWeekDay 
         Caption         =   "Quinta"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkWeekDay 
         Caption         =   "Quarta"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   47
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkWeekDay 
         Caption         =   "Terça"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   46
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox chkWeekDay 
         Caption         =   "Segunda"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   45
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkWeekDay 
         Caption         =   "Domingo"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   6855
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.OptionButton optRebattle 
         Caption         =   "Rebattle Never"
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   41
         Top             =   6360
         Width           =   1455
      End
      Begin VB.OptionButton optRebattle 
         Caption         =   "Rebattle Lose"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   40
         Top             =   6120
         Width           =   1335
      End
      Begin VB.OptionButton optRebattle 
         Caption         =   "None"
         Height          =   195
         Index           =   0
         Left            =   4200
         TabIndex        =   39
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox txtRewardExp 
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         Text            =   "0"
         Top             =   6120
         Width           =   1815
      End
      Begin VB.HScrollBar scrlWinConvo 
         Height          =   255
         Left            =   2520
         Max             =   0
         TabIndex        =   28
         Top             =   6480
         Width           =   1575
      End
      Begin VB.TextBox txtReward 
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         Text            =   "0"
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pokemon"
         Height          =   3375
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   5535
         Begin VB.TextBox txtFindItem 
            Height          =   285
            Left            =   3480
            TabIndex        =   42
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CheckBox chkIv 
            Caption         =   "IV Full"
            Height          =   255
            Left            =   1560
            TabIndex        =   36
            Top             =   2760
            Width           =   975
         End
         Begin VB.CheckBox chkShiny 
            Caption         =   "Shiny"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox cmbNature 
            Height          =   315
            Left            =   240
            TabIndex        =   33
            Text            =   "Combo1"
            Top             =   2400
            Width           =   2535
         End
         Begin VB.ComboBox cmbItem 
            Height          =   315
            Left            =   2880
            TabIndex        =   32
            Text            =   "Combo1"
            Top             =   3000
            Width           =   2415
         End
         Begin VB.TextBox txtFind 
            Height          =   285
            Left            =   1080
            TabIndex        =   30
            Top             =   1320
            Width           =   1695
         End
         Begin VB.ComboBox cmbMoveset 
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2400
            Width           =   1815
         End
         Begin VB.TextBox txtFindMoveset 
            Height          =   285
            Left            =   3480
            TabIndex        =   23
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   1200
            TabIndex        =   22
            Text            =   "0"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.ListBox lstMoveset 
            Height          =   840
            Left            =   2880
            TabIndex        =   20
            Top             =   1320
            Width           =   2415
         End
         Begin VB.ComboBox cmbPokeNum 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1560
            Width           =   2535
         End
         Begin VB.ListBox lstPokemon 
            Height          =   1035
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   5055
         End
         Begin VB.Label Label9 
            Caption         =   "Nature"
            Height          =   255
            Left            =   960
            TabIndex        =   34
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Item"
            Height          =   255
            Left            =   2880
            TabIndex        =   31
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Move"
            Height          =   255
            Left            =   2880
            TabIndex        =   25
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Level:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Pokemon"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   1320
            Width           =   2415
         End
      End
      Begin VB.ComboBox cmbNpcType 
         Height          =   315
         ItemData        =   "frmEditor_Npc.frx":0000
         Left            =   1200
         List            =   "frmEditor_Npc.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2040
         Width           =   4575
      End
      Begin VB.HScrollBar scrlConvo 
         Height          =   255
         Left            =   3000
         Max             =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   315
         ItemData        =   "frmEditor_Npc.frx":003B
         Left            =   1200
         List            =   "frmEditor_Npc.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   3855
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   0
         TabIndex        =   7
         Top             =   720
         Width           =   3855
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   5160
         ScaleHeight     =   66.065
         ScaleMode       =   0  'User
         ScaleWidth      =   34.133
         TabIndex        =   5
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label10 
         Caption         =   "Exp:"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Money:"
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label lblWinConvo 
         Caption         =   "Win Convo:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   6480
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblConvo 
         Caption         =   "Conversation: None"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Label Label2 
         Caption         =   "Behaviour:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   4740
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2655
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
Attribute VB_Name = "frmEditor_Npc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PokemonIndex As Long
Private MoveIndex As Long

Private Sub chkIv_Click()
    If PokemonIndex = 0 Then Exit Sub
    Npc(EditorIndex).PokemonIvFull(PokemonIndex) = chkIv
    EditorChange = True
End Sub

Private Sub chkShiny_Click()
    If PokemonIndex = 0 Then Exit Sub
    Npc(EditorIndex).PokemonIsShiny(PokemonIndex) = chkShiny
    EditorChange = True
End Sub

Private Sub chkWeekDay_Click(Index As Integer)
    If PokemonIndex = 0 Then Exit Sub
    Npc(EditorIndex).SpawnWeekDay(Index + 1) = chkWeekDay(Index).value
    EditorChange = True
End Sub

Private Sub cmbBehaviour_Click()
    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    EditorChange = True
End Sub

Private Sub cmbItem_Click()

    If PokemonIndex = 0 Then Exit Sub
    If cmbItem.ListIndex > 0 Then
        If Item(cmbItem.ListIndex).NotEquipable = YES Then
            MsgBox "O item selecionado não pode ser equipado por pokemon, altere no editor de item"
            Exit Sub
        End If
    End If

    Npc(EditorIndex).PokemonItem(PokemonIndex) = cmbItem.ListIndex
    EditorChange = True
End Sub

Private Sub cmbMoveset_Click()
Dim tmpIndex As Long

    If PokemonIndex = 0 Then Exit Sub
    If MoveIndex = 0 Then Exit Sub
    tmpIndex = lstMoveset.ListIndex
    lstMoveset.RemoveItem MoveIndex - 1
    Npc(EditorIndex).PokemonMoveset(PokemonIndex, MoveIndex) = cmbMoveset.ListIndex
    If cmbMoveset.ListIndex > 0 Then
        lstMoveset.AddItem MoveIndex & ": " & Trim$(PokemonMove(cmbMoveset.ListIndex).Name), MoveIndex - 1
    Else
        lstMoveset.AddItem MoveIndex & ": None", MoveIndex - 1
    End If
    lstMoveset.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub cmbNature_Click()
    If PokemonIndex = 0 Then Exit Sub
    Npc(EditorIndex).PokemonNature(PokemonIndex) = cmbNature.ListIndex - 1
    EditorChange = True
End Sub

Private Sub cmbPokeNum_Click()
Dim tmpIndex As Long

    If PokemonIndex = 0 Then Exit Sub
    tmpIndex = lstPokemon.ListIndex
    lstPokemon.RemoveItem PokemonIndex - 1
    Npc(EditorIndex).PokemonNum(PokemonIndex) = cmbPokeNum.ListIndex
    If cmbPokeNum.ListIndex > 0 Then
        lstPokemon.AddItem PokemonIndex & ": " & Trim$(Pokemon(cmbPokeNum.ListIndex).Name) & " Lv: " & Npc(EditorIndex).PokemonLevel(PokemonIndex), PokemonIndex - 1
    Else
        lstPokemon.AddItem PokemonIndex & ": None", PokemonIndex - 1
    End If
    lstPokemon.ListIndex = tmpIndex
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
        cuBound = MAX_NPC
        
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
        CloseNpcEditor
    End If
End Sub

Private Sub Form_Load()
    scrlSprite.max = Count_Character
    txtName.MaxLength = NAME_LENGTH
    scrlConvo.max = MAX_CONVERSATION
    scrlWinConvo.max = MAX_CONVERSATION
    
    txtFind.Text = "Digite um Nome ou ID..."
    txtFind.ForeColor = vbGrayText ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
    
    txtFindMoveset.Text = "Digite um Nome ou ID..."
    txtFindMoveset.ForeColor = vbGrayText ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
    
    txtFindItem.Text = "Digite um Nome ou ID..."
    txtFindItem.ForeColor = vbGrayText ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
End Sub

Private Sub lstIndex_Click()
    NpcEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub lstMoveset_Click()
    MoveIndex = lstMoveset.ListIndex + 1
    
    If PokemonIndex <= 0 Then Exit Sub
    If MoveIndex <= 0 Then Exit Sub
    
    cmbMoveset.ListIndex = Npc(EditorIndex).PokemonMoveset(PokemonIndex, MoveIndex)
End Sub

Private Sub lstPokemon_Click()
Dim X As Byte

    PokemonIndex = lstPokemon.ListIndex + 1
    
    If PokemonIndex <= 0 Then Exit Sub
    
    cmbPokeNum.ListIndex = Npc(EditorIndex).PokemonNum(PokemonIndex)
    txtLevel.Text = Npc(EditorIndex).PokemonLevel(PokemonIndex)
    cmbItem.ListIndex = Npc(EditorIndex).PokemonItem(PokemonIndex)
    cmbNature.ListIndex = Npc(EditorIndex).PokemonNature(PokemonIndex) + 1
    chkShiny = Npc(EditorIndex).PokemonIsShiny(PokemonIndex)
    chkIv = Npc(EditorIndex).PokemonIvFull(PokemonIndex)
    lstMoveset.Clear
    For X = 1 To MAX_MOVESET
        If Npc(EditorIndex).PokemonMoveset(PokemonIndex, X) > 0 Then
            lstMoveset.AddItem X & ": " & Trim$(PokemonMove(Npc(EditorIndex).PokemonMoveset(PokemonIndex, X)).Name)
        Else
            lstMoveset.AddItem X & ": None"
        End If
    Next
    lstMoveset.ListIndex = 0
    cmbMoveset.ListIndex = Npc(EditorIndex).PokemonMoveset(PokemonIndex, 1)
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestNpc
    End If
    CloseNpcEditor
End Sub

Private Sub mnuExit_Click()
    CloseNpcEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_NPC
        If NpcChange(i) Then
            SendSaveNpc i
            NpcChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseNpcEditor
End Sub

Private Sub optRebattle_Click(Index As Integer)
    Npc(EditorIndex).Rebatle = Index
    EditorChange = True
End Sub

Private Sub scrlConvo_Change()
    If scrlConvo.value > 0 Then
        lblConvo.Caption = "Conversation: " & Trim$(Conversation(scrlConvo.value).Name)
    Else
        lblConvo.Caption = "Conversation: None"
    End If
    Npc(EditorIndex).Convo = scrlConvo.value
    EditorChange = True
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    Npc(EditorIndex).Sprite = scrlSprite.value
    EditorChange = True
End Sub

Private Sub scrlWinConvo_Change()
    If scrlWinConvo.value > 0 Then
        lblWinConvo.Caption = "Win Convo: " & Trim$(Conversation(scrlWinConvo.value).Name)
    Else
        lblWinConvo.Caption = "Win Convo: None"
    End If
    Npc(EditorIndex).WinEvent = scrlWinConvo.value
    EditorChange = True
End Sub

Private Sub txtFind_Click()
    txtFind.Text = ""
    txtFind.ForeColor = vbWindowText ' Altera a cor do texto para preto para indicar que é uma mensagem descritiva
End Sub

Private Sub txtFind_LostFocus()
    txtFind.Text = "Digite um Nome ou ID..."
    txtFind.ForeColor = vbGrayText    ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
End Sub

Private Sub txtFind_Change()
    Dim Find As String, i As Long
    Dim MAX_INDEX As Integer, MinChar As Byte
    
    ' Maior Índice  \/
    MAX_INDEX = MAX_POKEMON
    
    ' Quantidade Mínima de caracteres pra procurar
    MinChar = 2
    
    ' Nome deste controle
    If Not IsNumeric(txtFind) Then
        ' Nome deste controle
        Find = UCase$(Trim$(txtFind))
        If Len(Find) <= MinChar And Not Find = "" Then
            'lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_INDEX
            If Not Find = "" Then
                ' Atribuição da estrutura em procura
                If InStr(1, UCase$(Trim$(Pokemon(i).Name)), Find) > 0 Then
                    ' Nome do controle a ser alterado
                    cmbPokeNum.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
    Else
        ' Nome deste controle
        If txtFind > MAX_INDEX Then
            ' Nome deste controle
            txtFind = MAX_INDEX
            ' Nome deste controle
        ElseIf txtFind <= 0 Then
            ' Nome deste controle
            txtFind = 1
        End If
        ' Nome do controle a ser alterado & Nome deste controle
        cmbPokeNum.ListIndex = txtFind
    End If
End Sub

Private Sub txtFindItem_Change()
    Dim Find As String, i As Long
    Dim MAX_INDEX As Integer, MinChar As Byte
    
    ' Maior Índice  \/
    MAX_INDEX = MAX_ITEM
    
    ' Quantidade Mínima de caracteres pra procurar
    MinChar = 2
    
    ' Nome deste controle
    If Not IsNumeric(txtFindItem) Then
        ' Nome deste controle
        Find = UCase$(Trim$(txtFindItem))
        If Len(Find) <= MinChar And Not Find = "" Then
            'lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_INDEX
            If Not Find = "" Then
                ' Atribuição da estrutura em procura
                If InStr(1, UCase$(Trim$(Item(i).Name)), Find) > 0 Then
                    ' Nome do controle a ser alterado
                    cmbItem.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
    Else
        ' Nome deste controle
        If txtFindItem > MAX_INDEX Then
            ' Nome deste controle
            txtFindItem = MAX_INDEX
            ' Nome deste controle
        ElseIf txtFindItem <= 0 Then
            ' Nome deste controle
            txtFindItem = 1
        End If
        ' Nome do controle a ser alterado & Nome deste controle
        cmbItem.ListIndex = txtFindItem
    End If
End Sub

Private Sub txtFindItem_Click()
    txtFindItem.Text = ""
    txtFindItem.ForeColor = vbWindowText ' Altera a cor do texto para preto para indicar que é uma mensagem descritiva
End Sub

Private Sub txtFindItem_LostFocus()
    txtFindItem.Text = "Digite um Nome ou ID..."
    txtFindItem.ForeColor = vbGrayText    ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
End Sub

Private Sub txtFindMoveset_Click()
    txtFindMoveset.Text = ""
    txtFindMoveset.ForeColor = vbWindowText ' Altera a cor do texto para preto para indicar que é uma mensagem descritiva
End Sub

Private Sub txtFindMoveset_LostFocus()
    txtFindMoveset.Text = "Digite um Nome ou ID..."
    txtFindMoveset.ForeColor = vbGrayText    ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
End Sub

Private Sub txtFindMoveset_Change()
    Dim Find As String, i As Long
    Dim MAX_INDEX As Integer, MinChar As Byte
    
    ' Maior Índice  \/
    MAX_INDEX = MAX_POKEMON_MOVE
    
    ' Quantidade Mínima de caracteres pra procurar
    MinChar = 2
    
    ' Nome deste controle
    If Not IsNumeric(txtFindMoveset) Then
        ' Nome deste controle
        Find = UCase$(Trim$(txtFindMoveset))
        If Len(Find) <= MinChar And Not Find = "" Then
            'lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_INDEX
            If Not Find = "" Then
                ' Atribuição da estrutura em procura
                If InStr(1, UCase$(Trim$(PokemonMove(i).Name)), Find) > 0 Then
                    ' Nome do controle a ser alterado
                    cmbMoveset.ListIndex = i
                    Exit Sub
                End If
            End If
        Next
    Else
        ' Nome deste controle
        If txtFindMoveset > MAX_INDEX Then
            ' Nome deste controle
            txtFindMoveset = MAX_INDEX
            ' Nome deste controle
        ElseIf txtFindMoveset <= 0 Then
            ' Nome deste controle
            txtFindMoveset = 1
        End If
        ' Nome do controle a ser alterado & Nome deste controle
        cmbMoveset.ListIndex = txtFindMoveset
    End If
End Sub

Private Sub txtLevel_Change()
Dim tmpIndex As Long

    If PokemonIndex = 0 Then Exit Sub
    If Not IsNumeric(txtLevel.Text) Then Exit Sub
    tmpIndex = lstPokemon.ListIndex
    lstPokemon.RemoveItem PokemonIndex - 1
    Npc(EditorIndex).PokemonLevel(PokemonIndex) = Val(txtLevel.Text)
    If Npc(EditorIndex).PokemonNum(PokemonIndex) > 0 Then
        lstPokemon.AddItem PokemonIndex & ": " & Trim$(Pokemon(Npc(EditorIndex).PokemonNum(PokemonIndex)).Name) & " Lv: " & Trim$(txtLevel.Text), PokemonIndex - 1
    Else
        lstPokemon.AddItem PokemonIndex & ": None", PokemonIndex - 1
    End If
    lstPokemon.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtReward_Change()
    If IsNumeric(txtReward.Text) Then
        Npc(EditorIndex).Reward = Val(txtReward.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtRewardExp_Change()
    If IsNumeric(txtRewardExp.Text) Then
        Npc(EditorIndex).RewardExp = Val(txtRewardExp.Text)
        EditorChange = True
    End If
End Sub
