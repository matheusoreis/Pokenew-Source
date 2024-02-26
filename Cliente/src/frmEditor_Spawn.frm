VERSION 5.00
Begin VB.Form frmEditor_Spawn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokemon Spawn Editor"
   ClientHeight    =   5865
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   5775
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.CheckBox chkFish 
         Caption         =   "Fishing?"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmbNature 
         Height          =   315
         Left            =   1200
         TabIndex        =   43
         Text            =   "PokeNature"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Frame Frame4 
         Caption         =   "Spawn Time(Hour)"
         Height          =   975
         Left            =   120
         TabIndex        =   34
         Top             =   4680
         Width           =   5055
         Begin VB.CommandButton Command6 
            Caption         =   "Night 18-04"
            Height          =   255
            Left            =   3720
            TabIndex        =   41
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Day 10-18"
            Height          =   255
            Left            =   3720
            TabIndex        =   40
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Morning 04-10"
            Height          =   255
            Left            =   3720
            TabIndex        =   39
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtSpawnMin 
            Height          =   285
            Left            =   1080
            TabIndex        =   36
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtSpawnMax 
            Height          =   285
            Left            =   2640
            TabIndex        =   35
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "to"
            Height          =   255
            Left            =   2160
            TabIndex        =   38
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Spawn Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Text            =   "ItemEquipped"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPokeBuff 
         Height          =   255
         Left            =   1200
         Max             =   100
         TabIndex        =   27
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CheckBox chkNoExp 
         Caption         =   "Cannot Give Exp?"
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   2520
         Width           =   1695
      End
      Begin VB.ComboBox cmbPokemonNum 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkCanCatch 
         Caption         =   "Cannot Catch?"
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtRarity 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Location:"
         Height          =   1335
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   5055
         Begin VB.CommandButton Command3 
            Caption         =   "Click On Map"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox chkRandomXY 
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkRandomMap 
            Caption         =   "Random"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtY 
            Height          =   285
            Left            =   2880
            TabIndex        =   17
            Text            =   "0"
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtX 
            Height          =   285
            Left            =   2880
            TabIndex        =   15
            Text            =   "0"
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtMap 
            Height          =   285
            Left            =   2880
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Y:"
            Height          =   255
            Left            =   1440
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "X:"
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Map:"
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtRespawn 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "0"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtMaxLevel 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtMinLevel 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Nature:"
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Item Equipped:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblBuff 
         Caption         =   "Poke Buff: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Common  [ 0 ~~ 100000 ] Rare"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Rarity:"
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "ms"
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Respawn:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Level Range:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblPokemon 
         Caption         =   "Pokemon: "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   280
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Paste"
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copy"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   5400
         Width           =   975
      End
      Begin VB.ListBox lstMapPokemon 
         Height          =   5130
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
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
   End
End
Attribute VB_Name = "frmEditor_Spawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CopySpawn As SpawnRec

Private Sub chkCanCatch_Click()
    Spawn(EditorIndex).CanCatch = chkCanCatch.value
End Sub

Private Sub chkFish_Click()
    Spawn(EditorIndex).Fishing = chkFish
    EditorChange = True
End Sub

Private Sub chkNoExp_Click()
    Spawn(EditorIndex).NoExp = chkNoExp.value
    EditorChange = True
End Sub

Private Sub chkRandomMap_Click()
    Spawn(EditorIndex).randomMap = chkRandomMap.value
    EditorChange = True
End Sub

Private Sub chkRandomXY_Click()
    Spawn(EditorIndex).randomXY = chkRandomXY.value
    EditorChange = True
End Sub

Private Sub cmbItem_Click()
    If cmbItem.ListIndex > 0 Then
        If Item(cmbItem.ListIndex).NotEquipable = YES Then
            MsgBox "O item selecionado não pode ser equipado por pokemon, altere no editor de item"
            Exit Sub
        End If
    End If

    Spawn(EditorIndex).HeldItem = cmbItem.ListIndex
    EditorChange = True
End Sub

Private Sub cmbNature_Click()
    Spawn(EditorIndex).Nature = cmbNature.ListIndex - 1
    EditorChange = True
End Sub

Private Sub cmbPokemonNum_Click()
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstMapPokemon.ListIndex
    lstMapPokemon.RemoveItem EditorIndex - 1
    Spawn(EditorIndex).PokeNum = cmbPokemonNum.ListIndex
    If cmbPokemonNum.ListIndex > 0 Then
        lstMapPokemon.AddItem EditorIndex & ": " & Trim$(Pokemon(cmbPokemonNum.ListIndex).Name), EditorIndex - 1
    Else
        lstMapPokemon.AddItem EditorIndex & ": ", EditorIndex - 1
    End If
    lstMapPokemon.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub Command1_Click()
    CopySpawn = Spawn(EditorIndex)
End Sub

Private Sub Command2_Click()
    Spawn(EditorIndex) = CopySpawn
    SpawnEditorLoadIndex EditorIndex
End Sub

Private Sub Command3_Click()
    SpawnSet = Not SpawnSet
    
    If SpawnSet Then
        Command3.Caption = "Cancel"
    Else
        Command3.Caption = "Click On Map"
    End If
End Sub

Private Sub Command4_Click()
    txtSpawnMin = 4
    txtSpawnMax = 10
End Sub

Private Sub Command5_Click()
    txtSpawnMin = 10
    txtSpawnMax = 18
End Sub

Private Sub Command6_Click()
    txtSpawnMin = 18
    txtSpawnMax = 4
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseSpawnEditor
    End If
End Sub

Private Sub Form_Load()
    txtFind.Text = "Digite um Nome ou ID..."
    txtFind.ForeColor = vbGrayText ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
End Sub

Private Sub lstMapPokemon_Click()
    SpawnEditorLoadIndex lstMapPokemon.ListIndex + 1
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestSpawn
    End If
    CloseSpawnEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        If SpawnChange(i) Then
            SendSaveSpawn i
            SpawnChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseSpawnEditor
End Sub

Private Sub scrlPokeBuff_Change()
    lblBuff.Caption = "Poke Buff: " & scrlPokeBuff.value
    Spawn(EditorIndex).PokeBuff = scrlPokeBuff.value
    EditorChange = True
End Sub

Private Sub txtFind_Change()
    Dim Find As String, i As Long
    Dim MAX_INDEX As Integer, MinChar As Byte
    
    ' Maior Índice  \/
    MAX_INDEX = MAX_ITEM
    
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
                    cmbPokemonNum.ListIndex = i
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
        cmbPokemonNum.ListIndex = txtFind
    End If
End Sub

Private Sub txtFind_Click()
    txtFind.Text = ""
    txtFind.ForeColor = vbWindowText ' Altera a cor do texto para preto para indicar que é uma mensagem descritiva
End Sub

Private Sub txtFind_LostFocus()
    txtFind.Text = "Digite um Nome ou ID..."
    txtFind.ForeColor = vbGrayText    ' Altera a cor do texto para cinza para indicar que é uma mensagem descritiva
End Sub

Private Sub txtMap_Change()
    If IsNumeric(txtMap.Text) Then
        Spawn(EditorIndex).MapNum = Val(txtMap.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtMaxLevel_Change()
    If IsNumeric(txtMaxLevel.Text) Then
        If Val(txtMaxLevel.Text) <= 0 Then txtMaxLevel.Text = 0
        If Val(txtMaxLevel.Text) >= 100 Then txtMaxLevel.Text = 100
        Spawn(EditorIndex).MaxLevel = Val(txtMaxLevel.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtMinLevel_Change()
    If IsNumeric(txtMinLevel.Text) Then
        If Val(txtMinLevel.Text) <= 0 Then txtMinLevel.Text = 0
        If Val(txtMinLevel.Text) >= 100 Then txtMinLevel.Text = 100
        Spawn(EditorIndex).MinLevel = Val(txtMinLevel.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtRarity_Change()
    If IsNumeric(txtRarity.Text) Then
        Spawn(EditorIndex).Rarity = Val(txtRarity.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtRespawn_Change()
    If IsNumeric(txtRespawn.Text) Then
        Spawn(EditorIndex).Respawn = Val(txtRespawn.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtSpawnMax_Change()
    If IsNumeric(txtSpawnMax.Text) Then
        Spawn(EditorIndex).SpawnTimeMax = Val(txtSpawnMax.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtSpawnMin_Change()
    If IsNumeric(txtSpawnMin.Text) Then
        Spawn(EditorIndex).SpawnTimeMin = Val(txtSpawnMin.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtX_Change()
    If IsNumeric(txtX.Text) Then
        Spawn(EditorIndex).MapX = Val(txtX.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtY_Change()
    If IsNumeric(txtY.Text) Then
        Spawn(EditorIndex).MapY = Val(txtY.Text)
        EditorChange = True
    End If
End Sub
