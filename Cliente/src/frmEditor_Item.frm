VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   10035
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   23595
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   669
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1573
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Paste"
         Height          =   255
         Left            =   1920
         TabIndex        =   77
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copy"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   3375
         ItemData        =   "frmEditor_Item.frx":0000
         Left            =   120
         List            =   "frmEditor_Item.frx":0002
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   8295
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   18495
      Begin VB.TextBox txtDelay 
         Height          =   285
         Left            =   4200
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3480
         TabIndex        =   47
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkNEquipable 
         Caption         =   "No Poke Equip"
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkLinked 
         Caption         =   "Vinculado"
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkIsCash 
         Caption         =   "Is Cash?"
         Height          =   195
         Left            =   4920
         TabIndex        =   35
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtDesc 
         Height          =   525
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Text            =   "0"
         Top             =   1150
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   0
         TabIndex        =   11
         Top             =   1440
         Width           =   135
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmEditor_Item.frx":0004
         Left            =   1200
         List            =   "frmEditor_Item.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox chkStock 
         Caption         =   "Acumular?"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5400
         ScaleHeight     =   32
         ScaleMode       =   0  'User
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   120
         Width           =   480
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   0
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Frame fraKeyItem 
         Caption         =   "Key Item Properties"
         Height          =   2295
         Left            =   11400
         TabIndex        =   28
         Top             =   5880
         Visible         =   0   'False
         Width           =   6735
         Begin VB.CheckBox chkPassiva 
            Caption         =   "Usar passiva de corrida com Shift?"
            Height          =   195
            Left            =   360
            TabIndex        =   75
            Top             =   1440
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.HScrollBar scrlExp 
            Height          =   255
            Left            =   1680
            Max             =   200
            TabIndex        =   74
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.HScrollBar scrlFish 
            Height          =   255
            Left            =   3840
            Max             =   2
            TabIndex        =   55
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox cmbKeyItemType 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":0088
            Left            =   2280
            List            =   "frmEditor_Item.frx":0092
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   360
            Width           =   3015
         End
         Begin VB.HScrollBar scrlSpriteType 
            Height          =   255
            Left            =   2280
            Max             =   5
            TabIndex        =   29
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label lblExp 
            AutoSize        =   -1  'True
            Caption         =   "Exp Bonus%:"
            Height          =   195
            Left            =   240
            TabIndex        =   73
            Top             =   1080
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblFish 
            BackStyle       =   0  'Transparent
            Caption         =   "Sprite:"
            Height          =   255
            Left            =   3120
            TabIndex        =   54
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Key Item Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblSpriteType 
            Caption         =   "Sprite Type: None"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   720
            Width           =   5055
         End
      End
      Begin VB.Frame fraMysteryBox 
         Caption         =   "Mystery Box"
         Height          =   1695
         Left            =   1320
         TabIndex        =   63
         Top             =   5760
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtChance 
            Height          =   285
            Left            =   720
            TabIndex        =   69
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtQuant 
            Height          =   285
            Left            =   720
            TabIndex        =   67
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox cmbItems 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":00A9
            Left            =   120
            List            =   "frmEditor_Item.frx":00AB
            TabIndex        =   65
            Text            =   "Combo1"
            Top             =   240
            Width           =   1575
         End
         Begin VB.ListBox lstItems 
            Height          =   1230
            Left            =   1800
            TabIndex        =   64
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblChanceF 
            AutoSize        =   -1  'True
            Caption         =   "Faltam: 0%"
            Height          =   195
            Left            =   3480
            TabIndex        =   72
            Top             =   120
            Width           =   765
         End
         Begin VB.Label lblChance 
            AutoSize        =   -1  'True
            Caption         =   "Chance Total: 0%"
            Height          =   195
            Left            =   1800
            TabIndex        =   71
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label Label16 
            Caption         =   "Chance:"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblQuant 
            Caption         =   "Quant:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame fraItemP 
         Caption         =   "Item Properties"
         Height          =   1455
         Left            =   11640
         TabIndex        =   58
         Top             =   480
         Visible         =   0   'False
         Width           =   5535
         Begin VB.Frame Frame4 
            Caption         =   "Pokemon Center"
            Height          =   1215
            Left            =   1200
            TabIndex        =   59
            Top             =   120
            Width           =   2775
            Begin VB.OptionButton OptData 
               Caption         =   "None"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   62
               Top             =   240
               Width           =   1815
            End
            Begin VB.OptionButton OptData 
               Caption         =   "Open Item Storage"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   61
               Top             =   600
               Width           =   1815
            End
            Begin VB.OptionButton OptData 
               Caption         =   "Open Poke Storage"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   60
               Top             =   960
               Width           =   1815
            End
         End
      End
      Begin VB.Frame fraPowerBracer 
         Caption         =   "Power Bracer"
         Height          =   1455
         Left            =   13920
         TabIndex        =   49
         Top             =   2160
         Width           =   4695
         Begin VB.TextBox txtPowerValue 
            Height          =   285
            Left            =   1440
            TabIndex        =   51
            Text            =   "0"
            Top             =   720
            Width           =   3015
         End
         Begin VB.ComboBox cmbPowerType 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":00AD
            Left            =   1440
            List            =   "frmEditor_Item.frx":00C6
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label14 
            Caption         =   "Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Type:"
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.Frame fraBerrie 
         Caption         =   "Berries/Proteins"
         Height          =   1455
         Left            =   6600
         TabIndex        =   42
         Top             =   3960
         Width           =   4695
         Begin VB.ComboBox cmbBerrieType 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":0115
            Left            =   1440
            List            =   "frmEditor_Item.frx":012E
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox txtBerrieValue 
            Height          =   285
            Left            =   1440
            TabIndex        =   43
            Text            =   "0"
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Berr./Prot. Type:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame fraTMHM 
         Caption         =   "TM/HM"
         Height          =   1215
         Left            =   5280
         TabIndex        =   38
         Top             =   2520
         Width           =   4695
         Begin VB.CheckBox chkTakeItem 
            Caption         =   "Take Item?"
            Height          =   255
            Left            =   1440
            TabIndex        =   40
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox cmbMoveList 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":0159
            Left            =   1440
            List            =   "frmEditor_Item.frx":015B
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label9 
            Caption         =   "Move List"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraPokeball 
         Caption         =   "Pokeball Properties"
         Height          =   1695
         Left            =   6120
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CheckBox chkAutoCatch 
            Caption         =   "Auto Catch?"
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   1080
            Width           =   2655
         End
         Begin VB.HScrollBar scrlBallSprite 
            Height          =   255
            Left            =   1680
            Max             =   15
            TabIndex        =   16
            Top             =   720
            Width           =   3615
         End
         Begin VB.TextBox txtCatchRate 
            Height          =   285
            Left            =   1680
            TabIndex        =   14
            Text            =   "0"
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblBallSprite 
            Caption         =   "Ball Sprite: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Catch Rate"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraMedicine 
         Caption         =   "Medicine"
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   4695
         Begin VB.CheckBox chkLevelUp 
            Caption         =   "Level Up"
            Height          =   255
            Left            =   1440
            TabIndex        =   26
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtValue 
            Height          =   285
            Left            =   1440
            TabIndex        =   23
            Text            =   "0"
            Top             =   720
            Width           =   3015
         End
         Begin VB.ComboBox cmbMedicineType 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":015D
            Left            =   1440
            List            =   "frmEditor_Item.frx":0179
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label7 
            Caption         =   "Value:"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Medicine Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delay(ms)"
         Height          =   195
         Left            =   4320
         TabIndex        =   57
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label12 
         Caption         =   "ID:"
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Description"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1455
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
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Option Explicit

' Private CopyItem As ItemRec

' Private Sub chkAutoCatch_Click()
'     Item(EditorIndex).Data3 = chkAutoCatch.Value
'     EditorChange = True
' End Sub

' Private Sub chkIsCash_Click()
'     Item(EditorIndex).IsCash = chkIsCash
'     EditorChange = True
' End Sub

' Private Sub chkLevelUp_Click()
'     Item(EditorIndex).Data3 = chkLevelUp.Value
'     EditorChange = True
' End Sub

' Private Sub chkLinked_Click()
'     Item(EditorIndex).Linked = chkLinked
'     EditorChange = True
' End Sub

' Private Sub chkNEquipable_Click()
'     Item(EditorIndex).NotEquipable = chkNEquipable.Value
'     EditorChange = True
' End Sub

' Private Sub chkPassiva_Click()
'     Item(EditorIndex).Data5 = chkPassiva.Value
'     EditorChange = True
' End Sub

' Private Sub chkStock_Click()
'     Item(EditorIndex).Stock = chkStock.Value
'     EditorChange = True
' End Sub

' Private Sub chkTakeItem_Click()
'     Item(EditorIndex).Data2 = chkTakeItem.Value
'     EditorChange = True
' End Sub

' Private Sub cmbBerrieType_Click()
'     Item(EditorIndex).Data1 = cmbBerrieType.ListIndex
'     EditorChange = True
' End Sub

' Private Sub cmbKeyItemType_Click()
'     Item(EditorIndex).Data1 = cmbKeyItemType.ListIndex
'     EditorChange = True
' End Sub

' Private Sub cmbMedicineType_Click()
'     Item(EditorIndex).Data1 = cmbMedicineType.ListIndex
'     EditorChange = True
' End Sub

' Private Sub cmbMoveList_Click()
'     Item(EditorIndex).Data1 = cmbMoveList.ListIndex
'     EditorChange = True
' End Sub

' Private Sub cmbPowerType_Click()
'     Item(EditorIndex).Data1 = cmbPowerType.ListIndex
'     EditorChange = True
' End Sub

' Private Sub cmbType_Click()
'     Item(EditorIndex).Type = cmbType.ListIndex
    
'     If Item(EditorIndex).Type = ItemTypeEnum.PokeBall Then
'         fraPokeball.Visible = True
'     Else
'         fraPokeball.Visible = False
'     End If
    
'     If Item(EditorIndex).Type = ItemTypeEnum.Medicine Then
'         fraMedicine.Visible = True
'     Else
'         fraMedicine.Visible = False
'     End If
    
'     If Item(EditorIndex).Type = ItemTypeEnum.keyItems Then
'         fraKeyItem.Visible = True
'     Else
'         fraKeyItem.Visible = False
'     End If
    
'     If Item(EditorIndex).Type = ItemTypeEnum.TM_HM Then
'         fraTMHM.Visible = True
'     Else
'         fraTMHM.Visible = False
'     End If
    
'     If Item(EditorIndex).Type = ItemTypeEnum.Berries Then
'         fraBerrie.Visible = True
'     Else
'         fraBerrie.Visible = False
'     End If
    
'     If Item(EditorIndex).Type = ItemTypeEnum.PowerBracer Then
'         fraPowerBracer.Visible = True
'     Else
'         fraPowerBracer.Visible = False
'     End If
    
'     If Item(EditorIndex).Type = ItemTypeEnum.Items Then
'         fraItemP.Visible = True
'     Else
'         fraItemP.Visible = False
'     End If
    
'     If Item(EditorIndex).Type = ItemTypeEnum.MysteryBox Then
'         fraMysteryBox.Visible = True
'     Else
'         fraMysteryBox.Visible = False
'     End If
    
'     EditorChange = True
' End Sub

' Private Sub cmdAdd_Click()
'     Dim tmpString() As String
'     Dim X As Long, tmpIndex As Long, Chance As Double

'     ' exit out if needed
'     If Not cmbItems.ListCount > 0 Then Exit Sub
'     If Not lstItems.ListCount > 0 Then Exit Sub

'     ' set the combo box properly
'     tmpString = Split(cmbItems.List(cmbItems.ListIndex))
'     ' make sure it's not a clear
'     If Not cmbItems.List(cmbItems.ListIndex) = "No Items" Then
'         Item(EditorIndex).Item(lstItems.ListIndex + 1) = cmbItems.ListIndex
'         Item(EditorIndex).ItemValue(lstItems.ListIndex + 1) = txtQuant.Text
'         Item(EditorIndex).ItemChance(lstItems.ListIndex + 1) = txtChance.Text
'     Else
'         Item(EditorIndex).Item(lstItems.ListIndex + 1) = 0
'         Item(EditorIndex).ItemValue(lstItems.ListIndex + 1) = 0
'         Item(EditorIndex).ItemChance(lstItems.ListIndex + 1) = 0
'     End If

'     ' re-load the list
'     tmpIndex = lstItems.ListIndex
'     lstItems.Clear
'     For X = 1 To MAX_MYSTERY_BOX
'         If Item(EditorIndex).Item(X) > 0 Then
'             lstItems.AddItem X & ": " & Item(EditorIndex).ItemValue(X) & "x - " & Trim$(Item(Item(EditorIndex).Item(X)).Name) & Item(EditorIndex).ItemChance(X) & "%"
'             Chance = Chance + Item(EditorIndex).ItemChance(X)
'         Else
'             lstItems.AddItem X & ": No Items"
'         End If
        
'     Next
    
'     lblChance = "Chance total: " & Chance & "%"
    
'     lblChanceF = "Faltam: " & (100 - Chance) & "%"
'     lstItems.ListIndex = tmpIndex

' End Sub

' Private Sub cmdIndexSearch_Click()
' Dim FindChar As String
' Dim clBound As Long, cuBound As Long
' Dim i As Long
' Dim ComboText As String
' Dim indexString As String
' Dim stringLength As Long

'     If Len(Trim$(txtIndexSearch.Text)) > 0 Then
'         FindChar = Trim$(txtIndexSearch.Text)
'         clBound = 1
'         cuBound = MAX_ITEM
        
'         For i = clBound To cuBound
'             ComboText = Trim$(lstIndex.List(i - 1))
'             indexString = i & ": "
'             stringLength = Len(ComboText) - Len(indexString)
'             If stringLength >= 0 Then
'                 ComboText = Mid$(ComboText, Len(indexString) + 1, stringLength)
'                 If LCase(ComboText) = LCase(FindChar) Then
'                     lstIndex.ListIndex = (i - 1)
'                     Exit Sub
'                 End If
'             End If
'         Next
        
'         MsgBox "Index not found", vbCritical
'     End If
' End Sub

' Private Sub Command1_Click()
'     CopyItem = Item(EditorIndex)
' End Sub

' Private Sub Command2_Click()
'     Item(EditorIndex) = CopyItem
'     ItemEditorLoadIndex EditorIndex
' End Sub

' Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'     If KeyCode = vbKeyEscape Then
'         CloseItemEditor
'     End If
' End Sub

' Private Sub Form_Load()
'     scrlSprite.max = Count_Item
'     txtName.MaxLength = NAME_LENGTH
' End Sub

' Private Sub lstIndex_Click()
'     ItemEditorLoadIndex lstIndex.ListIndex + 1
' End Sub

' Private Sub mnuCancel_Click()
'     '//Check if something was edited
'     If EditorChange Then
'         '//Request old data
'         SendRequestItem
'     End If
'     CloseItemEditor
' End Sub

' Private Sub mnuExit_Click()
'     CloseItemEditor
' End Sub

' Private Sub mnuSave_Click()
' Dim i As Long

'     For i = 1 To MAX_ITEM
'         If ItemChange(i) Then
'             SendSaveItem i
'             ItemChange(i) = False
'         End If
'     Next
'     MsgBox "Data was saved!", vbOKOnly
'     '//reset
'     EditorChange = False
'     'CloseItemEditor
' End Sub

' Private Sub OptData_Click(Index As Integer)
'     Item(EditorIndex).Data1 = Index
'     EditorChange = True
' End Sub

' Private Sub scrlBallSprite_Change()
'     lblBallSprite.Caption = "Ball Sprite: " & scrlBallSprite.Value
'     Item(EditorIndex).Data2 = scrlBallSprite.Value
'     EditorChange = True
' End Sub

' Private Sub scrlExp_Change()
'     lblExp = "Exp: " & scrlExp & "%"
'     Item(EditorIndex).Data4 = scrlExp.Value
'     EditorChange = True
' End Sub

' Private Sub scrlFish_Change()
'     lblFish = "Sprite: " & scrlFish
'     Item(EditorIndex).Data3 = scrlFish.Value
'     EditorChange = True
' End Sub

' Private Sub scrlSprite_Change()
'     lblSprite.Caption = "Sprite: " & scrlSprite.Value
'     Item(EditorIndex).Sprite = scrlSprite.Value
'     EditorChange = True
' End Sub

' Private Sub scrlSpriteType_Change()
'     scrlFish.Value = 0
'     scrlExp.Value = 0
'     chkPassiva.Value = 0
    
'     scrlFish.Visible = False
'     lblFish.Visible = False
'     scrlExp.Visible = False
'     lblExp.Visible = False
'     chkPassiva.Visible = False
    
    
'     Select Case scrlSpriteType.Value
'         Case TEMP_SPRITE_GROUP_DIVE
'             lblSpriteType.Caption = "Sprite Type: Dive"
'         Case TEMP_SPRITE_GROUP_BIKE
'             lblSpriteType.Caption = "Sprite Type: Bike"
'         Case TEMP_SPRITE_GROUP_SURF
'             lblSpriteType.Caption = "Sprite Type: Surf"
'         Case TEMP_SPRITE_GROUP_MOUNT
'             lblSpriteType.Caption = "Sprite Type: Mount"
'             scrlFish.Visible = True
'             lblFish.Visible = True
'             scrlFish.max = Count_PlayerSprite_M(1)
'             scrlFish = Item(EditorIndex).Data3
'             scrlExp.Visible = True
'             lblExp.Visible = True
'             scrlExp.Value = Item(EditorIndex).Data4
'             chkPassiva.Visible = True
'             chkPassiva.Value = Item(EditorIndex).Data5
'         Case TEMP_FISH_MODE
'             lblSpriteType.Caption = "Sprite Type: Fish"
'             scrlFish.Visible = True
'             lblFish.Visible = True
'             scrlFish = Item(EditorIndex).Data3
'         Case Else
'             lblSpriteType.Caption = "Sprite Type: None"
'     End Select
'     Item(EditorIndex).Data2 = scrlSpriteType.Value
'     EditorChange = True
' End Sub

' Private Sub txtBerrieValue_Change()
'     If IsNumeric(txtBerrieValue) Then
'         Item(EditorIndex).Data2 = Val(txtBerrieValue)
'         EditorChange = True
'     End If
' End Sub

' Private Sub txtCatchRate_Change()
'     If IsNumeric(txtCatchRate.Text) Then
'         Item(EditorIndex).Data1 = Val(txtCatchRate.Text)
'         EditorChange = True
'     End If
' End Sub

' Private Sub txtDelay_Change()
'     If Not IsNumeric(txtDelay) Then
'         txtDelay = 0
'     End If
    
'     If txtDelay < 0 Then
'         txtDelay = 0
'     End If
    
'     Item(EditorIndex).Delay = txtDelay
'     EditorChange = True
' End Sub

' Private Sub txtDesc_Change()
'     If EditorIndex < 0 Or EditorIndex > MAX_ITEM Then Exit Sub
'     Item(EditorIndex).Desc = Trim$(txtDesc.Text)
'     EditorChange = True
' End Sub

' Private Sub txtID_Change()
'     If Not IsNumeric(txtID) Then
'         txtID = 0
'     End If
    
'     If txtID < 0 Then
'         txtID = 0
'     End If
    
'     If txtID > Count_Item Then
'         txtID = Count_Item - 1
'     End If
    
'     scrlSprite.Value = CInt(txtID)
' End Sub

' Private Sub txtName_Validate(Cancel As Boolean)
' Dim tmpIndex As Long

'     If EditorIndex = 0 Then Exit Sub
'     tmpIndex = lstIndex.ListIndex
'     Item(EditorIndex).Name = Trim$(txtName.Text)
'     lstIndex.RemoveItem EditorIndex - 1
'     lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
'     lstIndex.ListIndex = tmpIndex
'     EditorChange = True
' End Sub

' Private Sub txtPowerValue_Change()
'     If IsNumeric(txtPowerValue) Then
'         Item(EditorIndex).Data2 = Val(txtPowerValue)
'         EditorChange = True
'     End If
' End Sub

' Private Sub txtPrice_Change()
'     If IsNumeric(txtPrice.Text) Then
'         Item(EditorIndex).Price = Val(txtPrice.Text)
'         EditorChange = True
'     End If
' End Sub

' Private Sub txtValue_Change()
'     If IsNumeric(txtValue.Text) Then
'         Item(EditorIndex).Data2 = Val(txtValue.Text)
'         EditorChange = True
'     End If
' End Sub
