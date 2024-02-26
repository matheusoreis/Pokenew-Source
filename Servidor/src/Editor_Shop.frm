VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmEditor_Itens 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Item"
   ClientHeight    =   9495
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabCategory 
      Height          =   3975
      Left            =   3720
      TabIndex        =   38
      Top             =   5400
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Pokebola"
      TabPicture(0)   =   "Editor_Shop.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameTabs(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cura"
      TabPicture(1)   =   "Editor_Shop.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameTabs(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Proteinas"
      TabPicture(2)   =   "Editor_Shop.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameTabs(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Chaves"
      TabPicture(3)   =   "Editor_Shop.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frameTabs(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Habilidades"
      TabPicture(4)   =   "Editor_Shop.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frameTabs(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Braceletes"
      TabPicture(5)   =   "Editor_Shop.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "frameTabs(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Gacha"
      TabPicture(6)   =   "Editor_Shop.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "frameTabs(6)"
      Tab(6).ControlCount=   1
      Begin VB.Frame frameTabs 
         Height          =   3495
         Index           =   6
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   7695
         Begin VB.HScrollBar scrollChanceGachaBox 
            Height          =   320
            Left            =   120
            Max             =   100
            TabIndex        =   52
            Top             =   1680
            Width           =   3615
         End
         Begin VB.HScrollBar scrollQuantyGachaBox 
            Height          =   320
            Left            =   120
            Max             =   100
            TabIndex        =   50
            Top             =   1080
            Width           =   3615
         End
         Begin VB.CommandButton buttonAddGachaBox 
            Caption         =   "Adicionar"
            Height          =   615
            Left            =   120
            TabIndex        =   49
            Top             =   2640
            Width           =   3615
         End
         Begin VB.ListBox listItensGachaBox 
            Height          =   2790
            Left            =   3840
            TabIndex        =   48
            Top             =   480
            Width           =   3735
         End
         Begin VB.ComboBox comboItensGachaBox 
            Height          =   315
            ItemData        =   "Editor_Shop.frx":00C4
            Left            =   120
            List            =   "Editor_Shop.frx":00C6
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label labelTotalChanceGachaBox 
            Caption         =   "Chance Total: 100%"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   2280
            Width           =   2280
         End
         Begin VB.Label labelMissingChanceGachaBox 
            Caption         =   "Faltam: 100%"
            Height          =   255
            Left            =   2520
            TabIndex        =   54
            Top             =   2280
            Width           =   1200
         End
         Begin VB.Label labelChanceGachaBox 
            Caption         =   "Chance do item:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   2760
         End
         Begin VB.Label labelQuantyGachaBox 
            Caption         =   "Quantidade do item:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   2760
         End
         Begin VB.Label Label5 
            Caption         =   "Lista de itens:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame frameTabs 
         Height          =   1575
         Index           =   5
         Left            =   -74880
         TabIndex        =   44
         Top             =   360
         Width           =   7695
         Begin VB.TextBox textValueBracelet 
            Height          =   320
            Left            =   120
            TabIndex        =   79
            Top             =   1080
            Width           =   7455
         End
         Begin VB.ComboBox comboTypeBracelet 
            Height          =   315
            ItemData        =   "Editor_Shop.frx":00C8
            Left            =   120
            List            =   "Editor_Shop.frx":00E7
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   480
            Width           =   7455
         End
         Begin VB.Label labelValueBracelet 
            Caption         =   "Valor do poder:"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   840
            Width           =   2760
         End
         Begin VB.Label labelTypeBracelet 
            Caption         =   "Tipo de Poder:"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame frameTabs 
         Height          =   1335
         Index           =   4
         Left            =   -74880
         TabIndex        =   43
         Top             =   360
         Width           =   7695
         Begin VB.CheckBox checkConsumeSkills 
            Caption         =   "Consumir o item?"
            Height          =   375
            Left            =   120
            TabIndex        =   76
            Top             =   840
            Width           =   7455
         End
         Begin VB.ComboBox comboMovesSkills 
            Height          =   315
            ItemData        =   "Editor_Shop.frx":014C
            Left            =   120
            List            =   "Editor_Shop.frx":016B
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   480
            Width           =   7455
         End
         Begin VB.Label labelMovesTmHm 
            Caption         =   "Lista de Movimentos:"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame frameTabs 
         Height          =   3495
         Index           =   3
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   7695
         Begin VB.Frame frameSkin 
            Caption         =   "Skin"
            Height          =   2535
            Left            =   120
            TabIndex        =   83
            Top             =   840
            Width           =   7455
            Begin VB.HScrollBar textBonusExperience 
               Height          =   320
               Left            =   120
               Max             =   100
               TabIndex        =   88
               Top             =   1080
               Width           =   7215
            End
            Begin VB.HScrollBar scrollSpriteSkin 
               Height          =   320
               Left            =   120
               Max             =   100
               TabIndex        =   87
               Top             =   480
               Width           =   6615
            End
            Begin VB.PictureBox pictureSpriteSkin 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   6840
               ScaleHeight     =   32
               ScaleMode       =   0  'User
               ScaleWidth      =   32
               TabIndex        =   86
               Top             =   360
               Width           =   480
            End
            Begin VB.CheckBox checkShiftKey 
               Caption         =   "Correr com Shift?"
               Height          =   375
               Left            =   120
               TabIndex        =   85
               Top             =   2085
               Width           =   7215
            End
            Begin VB.HScrollBar scrollBonusMoney 
               Height          =   320
               Left            =   120
               Max             =   100
               TabIndex        =   84
               Top             =   1680
               Width           =   7215
            End
            Begin VB.Label labelBonusExperience 
               Caption         =   "Bonus de experiência:"
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   840
               Width           =   2760
            End
            Begin VB.Label labelSpriteSkin 
               Caption         =   "Sprite:"
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   3480
            End
            Begin VB.Label labelBonusMoney 
               Caption         =   "Bonus de dinheiro:"
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   1440
               Width           =   2760
            End
         End
         Begin VB.ComboBox comboTypeKey 
            Height          =   315
            ItemData        =   "Editor_Shop.frx":01D0
            Left            =   120
            List            =   "Editor_Shop.frx":01E0
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   480
            Width           =   7455
         End
         Begin VB.Label labelTypeKey 
            Caption         =   "Tipo do item:"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Frame frameTabs 
         Height          =   1575
         Index           =   2
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   7695
         Begin VB.HScrollBar scrollProteinValue 
            Height          =   320
            Left            =   120
            Max             =   100
            TabIndex        =   82
            Top             =   1080
            Width           =   7455
         End
         Begin VB.ComboBox comboProteinType 
            Height          =   315
            ItemData        =   "Editor_Shop.frx":0211
            Left            =   120
            List            =   "Editor_Shop.frx":022A
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   480
            Width           =   7455
         End
         Begin VB.Label labelProteinValue 
            Caption         =   "Valor da proteina:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   840
            Width           =   7440
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo da proteina:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   7440
         End
      End
      Begin VB.Frame frameTabs 
         Height          =   1935
         Index           =   1
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   7695
         Begin VB.CheckBox checkLevelUp 
            Caption         =   "Aumentar o nível?"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   7455
         End
         Begin VB.HScrollBar scrollCureValue 
            Height          =   320
            Left            =   120
            Max             =   100
            TabIndex        =   66
            Top             =   1080
            Width           =   7455
         End
         Begin VB.ComboBox comboTypeCure 
            Height          =   315
            ItemData        =   "Editor_Shop.frx":0255
            Left            =   120
            List            =   "Editor_Shop.frx":026E
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   480
            Width           =   7455
         End
         Begin VB.Label labelCureValue 
            Caption         =   "Valor da medicina:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   840
            Width           =   7440
         End
         Begin VB.Label labelTypeCure 
            Caption         =   "Tipo da medicina:"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   7440
         End
      End
      Begin VB.Frame frameTabs 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   7695
         Begin VB.CheckBox checkPerfectCapture 
            Caption         =   "Captura pokemon imediatamente?"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   1440
            Width           =   7455
         End
         Begin VB.PictureBox pictureSpritePokeball 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   7080
            ScaleHeight     =   32
            ScaleMode       =   0  'User
            ScaleWidth      =   32
            TabIndex        =   60
            Top             =   960
            Width           =   480
         End
         Begin VB.HScrollBar scrollSpritePokeball 
            Height          =   320
            Left            =   120
            Max             =   100
            TabIndex        =   59
            Top             =   1080
            Width           =   6855
         End
         Begin VB.HScrollBar scrollChancePokeball 
            Height          =   320
            Left            =   120
            Max             =   100
            TabIndex        =   57
            Top             =   480
            Width           =   7455
         End
         Begin VB.Label labelSpritePokeball 
            Caption         =   "Sprite:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   6840
         End
         Begin VB.Label labelChancePokeball 
            Caption         =   "Chance de sucesso:"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   7440
         End
      End
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Detalhes"
      Height          =   5295
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   7935
      Begin TabDlg.SSTab tabProperties 
         Height          =   4695
         Left            =   3240
         TabIndex        =   14
         Top             =   405
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   8281
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Cooldown"
         TabPicture(0)   =   "Editor_Shop.frx":02C0
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "frameCooldown"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Restrições"
         TabPicture(1)   =   "Editor_Shop.frx":02DC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frameRestriction"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Requisitos"
         TabPicture(2)   =   "Editor_Shop.frx":02F8
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "tabPropertiesCategory"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin TabDlg.SSTab tabPropertiesCategory 
            Height          =   4095
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7223
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Pokémon"
            TabPicture(0)   =   "Editor_Shop.frx":0314
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame4"
            Tab(0).Control(1)=   "scrollPokemonLevel"
            Tab(0).Control(2)=   "labelPokemonLevel"
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Treinador"
            TabPicture(1)   =   "Editor_Shop.frx":0330
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "labelPlayerLevel"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Label8"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label9"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "scrollPlayerLevel"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "comboBadge"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "listMaps"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "buttonAddMap"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "comboMap"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).ControlCount=   8
            Begin VB.ComboBox comboMap 
               Height          =   315
               ItemData        =   "Editor_Shop.frx":034C
               Left            =   120
               List            =   "Editor_Shop.frx":0353
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   1200
               Width           =   4095
            End
            Begin VB.CommandButton buttonAddMap 
               Caption         =   "Adicionar"
               Height          =   615
               Left            =   120
               TabIndex        =   63
               Top             =   1560
               Width           =   4095
            End
            Begin VB.ListBox listMaps 
               Height          =   1035
               ItemData        =   "Editor_Shop.frx":035D
               Left            =   120
               List            =   "Editor_Shop.frx":035F
               TabIndex        =   62
               Top             =   2280
               Width           =   4095
            End
            Begin VB.ComboBox comboBadge 
               Height          =   315
               ItemData        =   "Editor_Shop.frx":0361
               Left            =   120
               List            =   "Editor_Shop.frx":0380
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   3600
               Width           =   4095
            End
            Begin VB.HScrollBar scrollPlayerLevel 
               Height          =   320
               Left            =   120
               Max             =   100
               TabIndex        =   33
               Top             =   600
               Width           =   4095
            End
            Begin VB.Frame Frame4 
               Caption         =   "Tipo do pokémon"
               Height          =   1095
               Left            =   -74880
               TabIndex        =   27
               Top             =   960
               Width           =   4095
               Begin VB.ComboBox comboSecondaryType 
                  Height          =   315
                  ItemData        =   "Editor_Shop.frx":03CB
                  Left            =   1680
                  List            =   "Editor_Shop.frx":0408
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   600
                  Width           =   2295
               End
               Begin VB.ComboBox comboPrimaryType 
                  Height          =   315
                  ItemData        =   "Editor_Shop.frx":0496
                  Left            =   1680
                  List            =   "Editor_Shop.frx":04D3
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.Label Label10 
                  Caption         =   "Secundário:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   31
                  Top             =   630
                  Width           =   1575
               End
               Begin VB.Label Label6 
                  Caption         =   "Primário:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   30
                  Top             =   285
                  Width           =   1575
               End
            End
            Begin VB.HScrollBar scrollPokemonLevel 
               Height          =   320
               Left            =   -74880
               Max             =   100
               TabIndex        =   26
               Top             =   600
               Width           =   4095
            End
            Begin VB.Label Label9 
               Caption         =   "Insignia:"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   3360
               Width           =   4080
            End
            Begin VB.Label Label8 
               Caption         =   "Mapa:"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   960
               Width           =   4080
            End
            Begin VB.Label labelPlayerLevel 
               Caption         =   "Level:"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   360
               Width           =   4080
            End
            Begin VB.Label labelPokemonLevel 
               Caption         =   "Level:"
               Height          =   255
               Left            =   -74880
               TabIndex        =   25
               Top             =   360
               Width           =   4080
            End
         End
         Begin VB.Frame frameRestriction 
            Height          =   1695
            Left            =   -74880
            TabIndex        =   20
            Top             =   360
            Width           =   4335
            Begin VB.CheckBox checkRestriction 
               Caption         =   "Apenas administradores"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   37
               Top             =   1200
               Width           =   4095
            End
            Begin VB.CheckBox checkRestriction 
               Caption         =   "Item pode agrupar"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   23
               Top             =   120
               Width           =   4095
            End
            Begin VB.CheckBox checkRestriction 
               Caption         =   "Pokemon não pode segurar"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   22
               Top             =   480
               Width           =   4095
            End
            Begin VB.CheckBox checkRestriction 
               Caption         =   "Vinculado ao usuário"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   21
               Top             =   840
               Width           =   4095
            End
         End
         Begin VB.Frame frameCooldown 
            Height          =   1575
            Left            =   -74880
            TabIndex        =   15
            Top             =   360
            Width           =   4335
            Begin VB.TextBox textCooldown 
               Height          =   320
               Left            =   120
               TabIndex        =   17
               Top             =   480
               Width           =   4095
            End
            Begin VB.ComboBox comboCooldown 
               Height          =   315
               ItemData        =   "Editor_Shop.frx":0561
               Left            =   120
               List            =   "Editor_Shop.frx":0571
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1080
               Width           =   4095
            End
            Begin VB.Label Label1 
               Caption         =   "Tipo do tempo:"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   840
               Width           =   4095
            End
            Begin VB.Label labelItemCooldown 
               Caption         =   "Tempo:"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   4095
            End
         End
      End
      Begin VB.ComboBox comboRarity 
         Height          =   315
         ItemData        =   "Editor_Shop.frx":059B
         Left            =   120
         List            =   "Editor_Shop.frx":05B1
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox textDescription 
         Height          =   2235
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   3015
      End
      Begin VB.ComboBox comboCategory 
         Height          =   315
         ItemData        =   "Editor_Shop.frx":05E7
         Left            =   120
         List            =   "Editor_Shop.frx":0603
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   3015
      End
      Begin VB.PictureBox pictureItemSprite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2640
         ScaleHeight     =   32
         ScaleMode       =   0  'User
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrollSprite 
         Height          =   320
         Left            =   120
         Max             =   100
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox textName 
         Height          =   320
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label labelItemRarity 
         Caption         =   "Raridade:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label labelDescription 
         Caption         =   "Descrição do Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   3000
      End
      Begin VB.Label labelType 
         Caption         =   "Tipo do Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   2760
      End
      Begin VB.Label labelSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label labelItemName 
         Caption         =   "Nome do item:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2760
      End
   End
   Begin VB.Frame frameIndex 
      Caption         =   "Lista de Itens"
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.ListBox listIndex 
         Height          =   9030
         ItemData        =   "Editor_Shop.frx":0648
         Left            =   120
         List            =   "Editor_Shop.frx":064A
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Menu menuFile 
      Caption         =   "Arquivo"
      Index           =   1
      Begin VB.Menu menuSave 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu menuExit 
         Caption         =   "Sair"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menuItem 
      Caption         =   "Item"
      Begin VB.Menu menuCopy 
         Caption         =   "Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu menuPaste 
         Caption         =   "Colar"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Ajuda"
   End
End
Attribute VB_Name = "frmEditor_Itens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    If comboCategory.listIndex <= O Or comboCategory.listIndex = 0 Then
        tabCategory.Visible = False
    Else
        Call UpdateTabsCategory
    End If

End Sub

Private Sub listIndex_Click()
    Call ItensEditorInit
End Sub

Private Sub menuExit_Click()
    Unload frmEditor_Itens
    Call ItensEditorClear
    Exit Sub
End Sub

Private Sub menuSave_Click()
    MsgBox "Deseja salvar os itens?", vbYesNo
    
    If vbYes Then
        Call ItensEditorSave
    Else
        Exit Sub
    End If
End Sub

Private Sub textName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    tmpIndex = listIndex.listIndex
    Item(EditorItemIndex).Name = Trim$(textName.Text)
    listIndex.RemoveItem EditorItemIndex - 1
    listIndex.AddItem EditorItemIndex & ": " & Item(EditorItemIndex).Name, EditorItemIndex - 1
    listIndex.listIndex = tmpIndex
End Sub

Private Sub scrollSprite_Change()
    labelSprite.Caption = "Sprite: " & scrollSprite.Value
    Item(EditorItemIndex).SpriteID = scrollSprite.Value
End Sub

Private Sub comboRarity_Click()
    Item(EditorItemIndex).Rarity = comboRarity.listIndex
End Sub

Private Sub comboCategory_Click()
    Call UpdateTabsCategory
End Sub

'Private Sub comboAction_Click()
'    Item(EditorItemIndex).ExecutionType = comboAction.listIndex
'End Sub

Private Sub textDescription_Validate(Cancel As Boolean)
    Item(EditorItemIndex).Description = Trim$(textDescription.Text)
End Sub

Private Sub tabProperties_Click(PreviusTab As Integer)
Dim tabIndex As Integer

    tabIndex = tabProperties.Tab
    
    Select Case tabIndex
        Case 0
            tabProperties.Height = 2055
        Case 1
            tabProperties.Height = 2175
        Case 2
            tabProperties.Height = 4695
    End Select

End Sub

Private Sub textCooldown_Validate(Cancel As Boolean)
    Item(EditorItemIndex).CooldownData.Value = Trim$(textCooldown.Text)
End Sub

Private Sub comboCooldown_Click()
    Item(EditorItemIndex).CooldownData.Type = comboCooldown.listIndex
End Sub

Private Sub checkRestriction_Click(Index As Integer)
    Dim CanStack As Boolean
    Dim CanHold As Boolean
    Dim IsConnected As Boolean
    Dim IsAdminItem As Boolean
    
    CanStack = ByteToBoolean(checkRestriction(Index).Value)
    CanHold = ByteToBoolean(checkRestriction(Index).Value)
    IsConnected = ByteToBoolean(checkRestriction(Index).Value)
    IsAdminItem = ByteToBoolean(checkRestriction(Index).Value)
    
    With Item(EditorItemIndex).RestrictionData
        .CanStack = CanStack
        .CanHold = CanHold
        .IsConnected = IsConnected
        .IsAdminItem = IsAdminItem
    End With
    
End Sub

Private Sub scrollPokemonLevel_Change()
    labelPokemonLevel.Caption = "Level: " & scrollPokemonLevel.Value
    
    With Item(EditorItemIndex).PokemonRequirementData
        .RequiredLevel = scrollPokemonLevel.Value
    End With
End Sub

Private Sub comboPrimaryType_Click()
    With Item(EditorItemIndex).PokemonRequirementData
        .PrimaryType = comboPrimaryType.listIndex
    End With
End Sub

Private Sub comboSecondaryType_Click()
    With Item(EditorItemIndex).PokemonRequirementData
        .SecondaryType = comboSecondaryType.listIndex
    End With
End Sub

Private Sub scrollPlayerLevel_Change()
    labelPlayerLevel.Caption = "Level: " & scrollPlayerLevel.Value
    
    With Item(EditorItemIndex).PlayerRequirementData
        .RequiredLevel = scrollPlayerLevel.Value
    End With
End Sub

Private Sub buttonAddMap_Click()
    Dim tmpString() As String
    Dim tmpIndex As Long
    
    If Not comboMap.ListCount > 0 Then Exit Sub
    If Not listMaps.ListCount > 0 Then Exit Sub

    With Item(EditorItemIndex).PlayerRequirementData
    
        tmpString = Split(comboMap.List(comboMap.listIndex))
        
        tmpIndex = listMaps.listIndex
        
        If tmpIndex >= 0 And tmpIndex < MAX_MAPS_REQUIREMENTS Then
            If Not comboMap.List(comboMap.listIndex) = "Nada" Then
                .RequiredMaps(tmpIndex + 1) = comboMap.listIndex
            Else
                .RequiredMaps(tmpIndex + 1) = 0
            End If
        End If
        
        listMaps.Clear
        For Index = 1 To MAX_MAPS_REQUIREMENTS
            
            If .RequiredMaps(Index) > 0 Then
                listMaps.AddItem Index & ": " & Map(.RequiredMaps(Index)).Name
            Else
                listMaps.AddItem Index & ": Nada"
            End If
            
        Next
    
    End With
    
    listMaps.listIndex = tmpIndex
End Sub

Private Sub comboBadge_Click()
    With Item(EditorItemIndex).PlayerRequirementData
        .RequiredBadge = comboBadge.listIndex
    End With
End Sub

Private Sub UpdateTabsCategory()
    Dim Index As Integer
    Dim TabEnabledStatus(6) As Boolean
    Dim ActiveTabIndex As Integer
    
    Item(EditorItemIndex).Category = comboCategory.listIndex

    If Item(EditorItemIndex).Category = ItemCategoryEnum.None Then
        tabCategory.Visible = False
    Else
        TabEnabledStatus(Item(EditorItemIndex).Category - 1) = True
    End If
    
    tabCategory.Visible = (Item(EditorItemIndex).Category <> ItemCategoryEnum.None)
    
    For Index = 0 To 6
    
        tabCategory.TabEnabled(Index) = TabEnabledStatus(Index)
        
        If TabEnabledStatus(Index) Then
            tabCategory.Tab = Index
            
            Select Case Index
                Case 0
                    tabCategory.Height = 2415
                Case 1
                    tabCategory.Height = 2415
                Case 2
                    tabCategory.Height = 2055
                Case 3
                    tabCategory.Height = 3975
                Case 4
                    tabCategory.Height = 1815
                Case 5
                    tabCategory.Height = 2055
                Case 6
                    tabCategory.Height = 3975
            End Select
            
            Call ItensEditorLoadCategory
            
        End If
    
    Next Index
End Sub

Private Sub scrollChancePokeball_Change()
    labelChancePokeball.Caption = "Chance de sucesso: " & scrollChancePokeball.Value
    
    With Item(EditorItemIndex).PokeballData
        .CaptureChance = scrollChancePokeball.Value
    End With
End Sub

Private Sub scrollSpritePokeball_Change()
    labelSpritePokeball.Caption = "Sprite: " & scrollSpritePokeball.Value
    
    With Item(EditorItemIndex).PokeballData
        .SpriteID = scrollSpritePokeball.Value
    End With
End Sub

Private Sub checkPerfectCapture_Click()
    Dim HasPerfectCapture As Boolean
    
    HasPerfectCapture = ByteToBoolean(checkPerfectCapture.Value)
    
    With Item(EditorItemIndex).PokeballData
        .HasPerfectCapture = HasPerfectCapture
    End With
    
End Sub

Private Sub comboTypeCure_Click()
    With Item(EditorItemIndex).MedicineData
        .Type = comboTypeCure.listIndex
    End With
End Sub

Private Sub scrollCureValue_Change()
    labelCureValue.Caption = "Valor da medicina: " & scrollCureValue.Value
    
    With Item(EditorItemIndex).MedicineData
        .Value = scrollCureValue.Value
    End With
End Sub

Private Sub checkLevelUp_Click()
    Dim HasLeveledUp As Boolean
    
    HasLeveledUp = ByteToBoolean(checkLevelUp.Value)
    
    With Item(EditorItemIndex).MedicineData
        .HasLeveledUp = HasLeveledUp
    End With
End Sub

Private Sub comboProteinType_Click()
    With Item(EditorItemIndex).ProteinsData
        .Type = comboProteinType.listIndex
    End With
End Sub

Private Sub scrollProteinValue_Change()
    labelProteinValue.Caption = "Valor da proteina: " & scrollProteinValue.Value
    
    With Item(EditorItemIndex).ProteinsData
        .Value = scrollProteinValue.Value
    End With
End Sub

Private Sub comboTypeKey_Click()
    With Item(EditorItemIndex).KeyData
        .Type = comboTypeKey.listIndex
        
        Select Case comboTypeKey.listIndex
        
            Case KeyTypeEnum.None
                frameSkin.Visible = False
                
            Case KeyTypeEnum.Sprite
                frameSkin.Visible = True
                
            Case KeyTypeEnum.OpenBank
                frameSkin.Visible = False
                
            Case KeyTypeEnum.OpenComputer
                frameSkin.Visible = False
                
            Case Else
                frameSkin.Visible = False
    
        End Select
    End With
End Sub

Private Sub scrollTypeKey_Change()

    
    
    'Select Case scrollTypeKey.Value
    '    Case TEMP_SPRITE_GROUP_DIVE
    '        labelTypeSpriteKey.Caption = "Tipo da Sprite: Dive"
    '    Case TEMP_SPRITE_GROUP_BIKE
    '        labelTypeSpriteKey.Caption = "Tipo da Sprite: Bike"
    '    Case TEMP_SPRITE_GROUP_SURF
    '        labelTypeSpriteKey.Caption = "Tipo da Sprite: Surf"
    '    Case TEMP_SPRITE_GROUP_MOUNT
            'lblSpriteType.Caption = "Tipo da Sprite: Mount"
            'scrlFish.Visible = True
            'lblFish.Visible = True
            'scrlFish.Max = Count_PlayerSprite_M(1)
            'scrlFish = Item(EditorIndex).Data3
            'scrlExp.Visible = True
            'lblExp.Visible = True
            'scrlExp.Value = Item(EditorIndex).Data4
            'chkPassiva.Visible = True
            'chkPassiva.Value = Item(EditorIndex).Data5
   '     Case TEMP_FISH_MODE
   '         labelTypeSpriteKey.Caption = "Tipo da Sprite: Fish"
            'scrlFish.Visible = True
            'lblFish.Visible = True
            'scrlFish = Item(EditorIndex).Data3
   '     Case Else
   '         labelTypeSpriteKey.Caption = "Tipo da Sprite: Nada"
   ' End Select
    
    
End Sub

Private Sub scrollSpriteSkin_Change()
    labelSpriteSkin.Caption = "Sprite: " & scrollSpriteSkin.Value
    
    With Item(EditorItemIndex).KeyData
        .Sprite = scrollSpriteSkin.Value
    End With
End Sub

Private Sub textBonusExperience_Change()
    labelBonusExperience.Caption = "Bonus de experiência: " & textBonusExperience.Value
    
    With Item(EditorItemIndex).KeyData
        .ExperienceBonusAmount = textBonusExperience.Value
    End With
End Sub

Private Sub scrollBonusMoney_Change()
    labelBonusMoney.Caption = "Bonus de experiência: " & scrollBonusMoney.Value
    
    With Item(EditorItemIndex).KeyData
        .MoneyBonusAmount = scrollBonusMoney.Value
    End With
End Sub

Private Sub checkShiftKey_Click()
    Dim IsShiftRunning As Boolean
    
    IsShiftRunning = ByteToBoolean(checkShiftKey.Value)
    
    With Item(EditorItemIndex).KeyData
        .IsShiftRunning = IsShiftRunning
    End With
End Sub
