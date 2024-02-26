VERSION 5.00
Begin VB.Form frmEnc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryptador"
   ClientHeight    =   4545
   ClientLeft      =   8700
   ClientTop       =   5505
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbExtension 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1920
      List            =   "Form1.frx":0002
      TabIndex        =   6
      Text            =   "Extensão"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.DirListBox dirAppPath 
      Height          =   1440
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3375
   End
   Begin VB.FileListBox flbAppPath 
      Height          =   1260
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Aguardando Ação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "O que Contém:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pasta dos Gráficos:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "frmEnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbExtension_Click()
    Select Case cmbExtension.ListIndex
        Case ExtensionEnum.PNG
            flbAppPath.Pattern = "*" & DecExtension
        Case ExtensionEnum.DAT
            flbAppPath.Pattern = "*" & EncExtension
    End Select
End Sub

Private Sub Command1_Click()

    'MsgBox dirAppPath.Path & "\"
    Dim I As Long, data() As Byte
    
    Call SetStatus("Cryptografando...", Yellow)
    ConvertPNGToBinary GlobalDir, I, data
End Sub

Private Sub Command2_Click()
    Dim I As Long, data() As Byte
    
    Call SetStatus("Descryptografando...", Yellow)
    ConvertBinaryToPNG GlobalDir, I, data

End Sub

Private Sub dirAppPath_Change()
    GlobalDir = dirAppPath.Path & "\"
    flbAppPath.Path = GlobalDir
End Sub

Private Sub Form_Load()
    'drvAppPath.AddItem "C:\MeuCaminho\"
    'drvAppPath.List(drvAppPath.ListCount - 1) = "C:\MeuCaminho\"
    'drvAppPath. = "C:\"
    'flbAppPath.Path = App.Path
    
    GlobalDir = dirAppPath.Path & "\"
    cmbExtension.AddItem DecExtension, ExtensionEnum.PNG
    cmbExtension.AddItem EncExtension, ExtensionEnum.DAT
End Sub
