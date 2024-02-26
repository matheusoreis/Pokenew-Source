VERSION 5.00
Begin VB.Form frmEditor_Shop 
   AutoRedraw      =   -1  'True
   Caption         =   "Editor de Lojas"
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameProperties 
      Caption         =   "Propriedades"
      Height          =   6135
      Left            =   3720
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      Begin VB.Frame frameShop 
         Caption         =   "Lista de itens:"
         Height          =   3495
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   3975
         Begin VB.PictureBox pictureShopItemName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   3360
            ScaleHeight     =   32
            ScaleMode       =   0  'User
            ScaleWidth      =   32
            TabIndex        =   17
            Top             =   2880
            Width           =   480
         End
         Begin VB.HScrollBar scrollShopItemName 
            Height          =   320
            Left            =   120
            Max             =   0
            TabIndex        =   15
            Top             =   3000
            Width           =   3135
         End
         Begin VB.ListBox listShopItens 
            Height          =   2010
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label labelShopItemName 
            Caption         =   "item: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2715
            Width           =   3135
         End
         Begin VB.Label labelShopItens 
            Caption         =   "Lista dos Itens:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   320
            Width           =   3615
         End
      End
      Begin VB.Frame frameDetails 
         Caption         =   "Detalhes"
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3975
         Begin VB.TextBox textShopItemValue 
            Height          =   320
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   3735
         End
         Begin VB.OptionButton optionShopCurrency 
            Caption         =   "Moeda Comum"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   3735
         End
         Begin VB.OptionButton optionShopCurrency 
            Caption         =   "Moeda Premium"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   3735
         End
         Begin VB.TextBox textShopName 
            Height          =   320
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label labelShopItemValue 
            Caption         =   "Valor do Item:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   3720
         End
         Begin VB.Label labelShopName 
            Caption         =   "Nome da loja:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   3720
         End
      End
   End
   Begin VB.Frame frameIndex 
      Caption         =   "Lista de Lojas"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton buttonPaste 
         Caption         =   "Colar"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   5400
         Width           =   3255
      End
      Begin VB.CommandButton buttonCopy 
         Caption         =   "Copiar"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   4755
         Width           =   3255
      End
      Begin VB.ListBox listIndex 
         Height          =   4350
         ItemData        =   "Editor_Item.frx":0000
         Left            =   120
         List            =   "Editor_Item.frx":0002
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
   Begin VB.Menu menuHelp 
      Caption         =   "Ajuda"
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
