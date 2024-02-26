VERSION 5.00
Begin VB.Form frmEditorShop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   3930
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   3735
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox Check1 
         Caption         =   "Is Cash?"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Text            =   "0"
         Top             =   3360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         Left            =   2520
         Max             =   0
         TabIndex        =   10
         Top             =   2760
         Width           =   3255
      End
      Begin VB.ListBox lstShopItem 
         Height          =   1620
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   360
         Width           =   480
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   0
         TabIndex        =   3
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         Caption         =   "Money"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblItemNum 
         Caption         =   "Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   5535
      End
      Begin VB.Label Label2 
         Caption         =   "Shop Items"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.ListBox lstIndex 
         Height          =   3180
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
Attribute VB_Name = "frmEditorShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        CloseShopEditor
    End If
End Sub

Private Sub Form_Load()
    txtName.MaxLength = NAME_LENGTH
    scrlItemNum.max = MAX_ITEM
End Sub

Private Sub lstIndex_Click()
    ShopEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub lstShopItem_Click()
    scrlItemNum.value = Shop(EditorIndex).ShopItem(lstShopItem.ListIndex + 1).Num
    'txtPrice.Text = Shop(EditorIndex).ShopItem(lstShopItem.ListIndex + 1).Price
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestShop
    End If
    CloseShopEditor
End Sub

Private Sub mnuExit_Click()
    CloseShopEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_SHOP
        If ShopChange(i) Then
            SendSaveShop i
            ShopChange(i) = False
        End If
    Next
    MsgBox "Data was saved!", vbOKOnly
    '//reset
    EditorChange = False
    'CloseShopEditor
End Sub

Private Sub scrlItemNum_Change()
Dim tmpIndex As Long
Dim shopIndex As Long
Dim Nomenclatura As String

    shopIndex = lstShopItem.ListIndex + 1
    If shopIndex = 0 Then Exit Sub
    tmpIndex = lstShopItem.ListIndex
    Shop(EditorIndex).ShopItem(shopIndex).Num = scrlItemNum.value
    lstShopItem.RemoveItem shopIndex - 1
    If Shop(EditorIndex).ShopItem(shopIndex).Num > 0 Then
    
        Nomenclatura = "Money:"
        If Item(Shop(EditorIndex).ShopItem(shopIndex).Num).IsCash = YES Then Nomenclatura = "Cash:"
        
        lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - " & Nomenclatura & "$" & Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Price, shopIndex - 1
        lblItemNum.Caption = "Item: " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name)
        If Item(Shop(EditorIndex).ShopItem(shopIndex).Num).IsCash = YES Then
            lblMoney.Caption = "Cash: " & Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Price
        Else
            lblMoney.Caption = "Money: " & Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Price
        End If
    Else
        lstShopItem.AddItem shopIndex & ": None - Price: $ 0", shopIndex - 1
        lblItemNum.Caption = "Item: None"
        lblMoney = "Money:"
    End If
    lstShopItem.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtPrice_Change()
'Dim tmpIndex As Long
'Dim shopIndex As Long

'    shopIndex = lstShopItem.ListIndex + 1
'    If shopIndex = 0 Then Exit Sub
'    tmpIndex = lstShopItem.ListIndex
'    If IsNumeric(txtPrice.Text) Then
'        Shop(EditorIndex).ShopItem(shopIndex).Price = Val(txtPrice.Text)
'    End If
'    lstShopItem.RemoveItem shopIndex - 1
'    If Shop(EditorIndex).ShopItem(shopIndex).Num > 0 Then
'        lstShopItem.AddItem shopIndex & ": " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name) & " - Price: $" & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
'        lblItemNum.Caption = "Item: " & Trim$(Item(Shop(EditorIndex).ShopItem(shopIndex).Num).Name)
'    Else
'        lstShopItem.AddItem shopIndex & ": None - Price: $" & Shop(EditorIndex).ShopItem(shopIndex).Price, shopIndex - 1
'        lblItemNum.Caption = "Item: None"
'    End If
'    lstShopItem.ListIndex = tmpIndex
'    EditorChange = True
End Sub
