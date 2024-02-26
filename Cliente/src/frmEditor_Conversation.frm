VERSION 5.00
Begin VB.Form frmEditor_Conversation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversation Editor"
   ClientHeight    =   6960
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   6855
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.Frame fraConvData 
         Caption         =   "Data - 1"
         Height          =   6015
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   5775
         Begin VB.TextBox txtCustomScriptData3 
            Height          =   285
            Left            =   4320
            TabIndex        =   33
            Text            =   "0"
            Top             =   5400
            Width           =   1095
         End
         Begin VB.TextBox txtCustomScriptData2 
            Height          =   285
            Left            =   4320
            TabIndex        =   29
            Text            =   "0"
            Top             =   5040
            Width           =   1095
         End
         Begin VB.TextBox txtCustomScriptData 
            Height          =   285
            Left            =   4320
            TabIndex        =   27
            Text            =   "0"
            Top             =   4680
            Width           =   1095
         End
         Begin VB.CheckBox chkNoReply 
            Caption         =   "Don't use reply"
            Height          =   255
            Left            =   3360
            TabIndex        =   26
            Top             =   3960
            Width           =   2295
         End
         Begin VB.TextBox txtMoveTo 
            Height          =   285
            Left            =   960
            TabIndex        =   25
            Text            =   "0"
            Top             =   5520
            Width           =   1215
         End
         Begin VB.HScrollBar scrlCustomScript 
            Height          =   255
            Left            =   3600
            Max             =   3
            TabIndex        =   20
            Top             =   4320
            Width           =   2055
         End
         Begin VB.CheckBox chkNoText 
            Caption         =   "Don't use conversation"
            Height          =   255
            Left            =   960
            TabIndex        =   18
            Top             =   3960
            Width           =   2415
         End
         Begin VB.Frame fraLanguage 
            Caption         =   "Language - En"
            Height          =   3135
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   5535
            Begin VB.TextBox txtReplyMove 
               Height          =   285
               Index           =   3
               Left            =   4560
               TabIndex        =   23
               Text            =   "0"
               Top             =   2640
               Width           =   855
            End
            Begin VB.TextBox txtReplyMove 
               Height          =   285
               Index           =   2
               Left            =   4560
               TabIndex        =   22
               Text            =   "0"
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox txtReplyMove 
               Height          =   285
               Index           =   1
               Left            =   4560
               TabIndex        =   21
               Text            =   "0"
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox txtReply 
               Height          =   285
               Index           =   3
               Left            =   840
               TabIndex        =   16
               Top             =   2640
               Width           =   3615
            End
            Begin VB.TextBox txtReply 
               Height          =   285
               Index           =   2
               Left            =   840
               TabIndex        =   15
               Top             =   2280
               Width           =   3615
            End
            Begin VB.TextBox txtReply 
               Height          =   285
               Index           =   1
               Left            =   840
               TabIndex        =   14
               Top             =   1920
               Width           =   3615
            End
            Begin VB.TextBox txtText 
               Height          =   1095
               Left            =   840
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Top             =   720
               Width           =   4575
            End
            Begin VB.HScrollBar scrlLanguage 
               Height          =   255
               Left            =   120
               Max             =   1
               Min             =   1
               TabIndex        =   10
               Top             =   240
               Value           =   1
               Width           =   5295
            End
            Begin VB.Label Label3 
               Caption         =   "Reply:"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Text:"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   720
               Width           =   1335
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   5400
               Y1              =   600
               Y2              =   600
            End
         End
         Begin VB.HScrollBar scrlData 
            Height          =   255
            Left            =   120
            Max             =   1
            Min             =   1
            TabIndex        =   8
            Top             =   240
            Value           =   1
            Width           =   5535
         End
         Begin VB.Label Label8 
            Caption         =   "Custom Script Data 3"
            Height          =   255
            Left            =   2520
            TabIndex        =   34
            Top             =   5400
            Width           =   1935
         End
         Begin VB.Label Label7 
            Caption         =   "Custom Script Data 2"
            Height          =   255
            Left            =   2520
            TabIndex        =   30
            Top             =   5040
            Width           =   1935
         End
         Begin VB.Label Label6 
            Caption         =   "Custom Script Data"
            Height          =   255
            Left            =   2520
            TabIndex        =   28
            Top             =   4680
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Move To:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   5520
            Width           =   2535
         End
         Begin VB.Label lblCustomScript 
            Caption         =   "Custom Script: None"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   4320
            Width           =   5535
         End
         Begin VB.Label Label4 
            Caption         =   "Others:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   3960
            Width           =   1935
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5640
            Y1              =   600
            Y2              =   600
         End
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
      Begin VB.Label Label1 
         Caption         =   "Identifier:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   6105
         ItemData        =   "frmEditor_Conversation.frx":0000
         Left            =   120
         List            =   "frmEditor_Conversation.frx":0002
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
Attribute VB_Name = "frmEditor_Conversation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ConvDataIndex As Byte
Private LanguageIndex As Byte

Private Sub chkNoReply_Click()
    If ConvDataIndex <= 0 Then Exit Sub
    Conversation(EditorIndex).ConvData(ConvDataIndex).NoReply = chkNoReply.value
    EditorChange = True
End Sub

Private Sub chkNoText_Click()
    If ConvDataIndex <= 0 Then Exit Sub
    Conversation(EditorIndex).ConvData(ConvDataIndex).NoText = chkNoText.value
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
        cuBound = MAX_CONVERSATION
        
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
        CloseConversationEditor
    End If
End Sub

Private Sub Form_Load()
    txtName.MaxLength = NAME_LENGTH
    scrlData.max = MAX_CONV_DATA
    ConvDataIndex = scrlData.value
    scrlLanguage.max = MAX_LANGUAGE
    LanguageIndex = scrlLanguage.value
    scrlCustomScript.max = MAX_CONVO_SCRIPT
End Sub

Private Sub lstIndex_Click()
    ConversationEditorLoadIndex lstIndex.ListIndex + 1
End Sub

Private Sub mnuCancel_Click()
    '//Check if something was edited
    If EditorChange Then
        '//Request old data
        SendRequestConversation
    End If
    CloseConversationEditor
End Sub

Private Sub mnuExit_Click()
    CloseConversationEditor
End Sub

Private Sub mnuSave_Click()
Dim i As Long

    For i = 1 To MAX_CONVERSATION
        If ConversationChange(i) Then
            SendSaveConversation i
            ConversationChange(i) = False
        End If
    Next
    '//reset
    EditorChange = False
    CloseConversationEditor
End Sub

Private Sub scrlCustomScript_Change()
Dim Text As String
    
    Select Case scrlCustomScript.value
        Case CONVO_SCRIPT_INVSTORAGE: Text = "Custom Script: Open Inv Storage"
        Case CONVO_SCRIPT_POKESTORAGE: Text = "Custom Script: Open Poke Storage"
        Case CONVO_SCRIPT_HEAL: Text = "Custom Script: Heal"
        Case CONVO_SCRIPT_SHOP: Text = "Custom Script: Shop"
        Case CONVO_SCRIPT_SETSWITCH: Text = "Custom Script: Set Switch"
        Case CONVO_SCRIPT_GIVEPOKE: Text = "Custom Script: Give Poke"
        Case CONVO_SCRIPT_GIVEITEM: Text = "Custom Script: Give Item"
        Case CONVO_SCRIPT_WARPTO: Text = "Custom Script: Warp To"
        Case CONVO_SCRIPT_CHECKMONEY: Text = "Custom Script: Check Money"
        Case CONVO_SCRIPT_TAKEMONEY: Text = "Custom Script: Take Money"
        Case CONVO_SCRIPT_STARTBATTLE: Text = "Custom Script: Start Battle"
        Case CONVO_SCRIPT_RELEARN: Text = "Custom Script: Relearn"
        Case CONVO_SCRIPT_GIVEBADGE: Text = "Custom Script: Give Badge"
        Case CONVO_SCRIPT_CHECKBADGE: Text = "Custom Script: Check Badge"
        Case CONVO_SCRIPT_BEATPOKE: Text = "Custom Script: Beat Poke"
        Case CONVO_SCRIPT_CHECKITEM: Text = "Custom Script: Check Item"
        Case CONVO_SCRIPT_TAKEITEM: Text = "Custom Script: Take Item"
        Case CONVO_SCRIPT_RESPAWNPOKE: Text = "Custom Script: Respawn Poke"
        Case CONVO_SCRIPT_CHECKLEVEL: Text = "Custom Script: Check Level"
        Case Else: Text = "Custom Script: None"
    End Select
    lblCustomScript.Caption = Text
    
    If ConvDataIndex <= 0 Then Exit Sub
    Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScript = scrlCustomScript.value
    EditorChange = True
End Sub

Private Sub scrlData_Change()
Dim i As Long

    fraConvData.Caption = "Data - " & scrlData.value
    ConvDataIndex = scrlData.value
    
    scrlLanguage.value = 1
    LanguageIndex = scrlLanguage.value
    
    If ConvDataIndex <= 0 Then Exit Sub
    If LanguageIndex <= 0 Then Exit Sub
    
    For i = 1 To 3
        txtReply(i).Text = Trim$(Conversation(EditorIndex).ConvData(ConvDataIndex).TextLang(LanguageIndex).tReply(i))
        txtReplyMove(i).Text = (Conversation(EditorIndex).ConvData(ConvDataIndex).tReplyMove(i))
    Next
    txtText.Text = Trim$(Conversation(EditorIndex).ConvData(ConvDataIndex).TextLang(LanguageIndex).Text)
    scrlCustomScript.value = Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScript
    txtCustomScriptData.Text = Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScriptData
    txtCustomScriptData2.Text = Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScriptData2
    txtCustomScriptData3.Text = Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScriptData3
    chkNoText.value = Conversation(EditorIndex).ConvData(ConvDataIndex).NoText
    chkNoReply.value = Conversation(EditorIndex).ConvData(ConvDataIndex).NoReply
    txtMoveTo.Text = Conversation(EditorIndex).ConvData(ConvDataIndex).MoveNext
End Sub

Private Sub scrlLanguage_Change()
Dim Text As String
Dim i As Byte

    Select Case scrlLanguage.value
        Case 1: Text = "Language - Portuguese" '//Portugues
        Case 2: Text = "Language - English" '//Ingles
        Case Else: Text = "Language - Spanish" '//Spañol
    End Select
    fraLanguage.Caption = Text
    LanguageIndex = scrlLanguage.value
    
    If ConvDataIndex <= 0 Then Exit Sub
    If LanguageIndex <= 0 Then Exit Sub
    
    For i = 1 To 3
        txtReply(i).Text = Trim$(Conversation(EditorIndex).ConvData(ConvDataIndex).TextLang(LanguageIndex).tReply(i))
    Next
    txtText.Text = Trim$(Conversation(EditorIndex).ConvData(ConvDataIndex).TextLang(LanguageIndex).Text)
End Sub

Private Sub txtCustomScriptData_Change()
    If ConvDataIndex <= 0 Then Exit Sub
    
    If IsNumeric(txtCustomScriptData.Text) Then
        Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScriptData = Val(txtCustomScriptData.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtCustomScriptData2_Change()
    If ConvDataIndex <= 0 Then Exit Sub
    
    If IsNumeric(txtCustomScriptData2.Text) Then
        Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScriptData2 = Val(txtCustomScriptData2.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtCustomScriptData3_Change()
    If ConvDataIndex <= 0 Then Exit Sub
    
    If IsNumeric(txtCustomScriptData3.Text) Then
        Conversation(EditorIndex).ConvData(ConvDataIndex).CustomScriptData3 = Val(txtCustomScriptData3.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtMoveTo_Change()
    If ConvDataIndex <= 0 Then Exit Sub
    
    If IsNumeric(txtMoveTo.Text) Then
        Conversation(EditorIndex).ConvData(ConvDataIndex).MoveNext = Val(txtMoveTo.Text)
        EditorChange = True
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Conversation(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conversation(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    EditorChange = True
End Sub

Private Sub txtReply_Change(Index As Integer)
    If ConvDataIndex <= 0 Then Exit Sub
    If LanguageIndex <= 0 Then Exit Sub
    
    Conversation(EditorIndex).ConvData(ConvDataIndex).TextLang(LanguageIndex).tReply(Index) = Trim$(txtReply(Index).Text)
    EditorChange = True
End Sub

Private Sub txtReplyMove_Change(Index As Integer)
    If ConvDataIndex <= 0 Then Exit Sub
    
    If IsNumeric(txtReplyMove(Index).Text) Then
        Conversation(EditorIndex).ConvData(ConvDataIndex).tReplyMove(Index) = Val(txtReplyMove(Index).Text)
        EditorChange = True
    End If
End Sub

Private Sub txtText_Change()
    If ConvDataIndex <= 0 Then Exit Sub
    If LanguageIndex <= 0 Then Exit Sub
    
    Conversation(EditorIndex).ConvData(ConvDataIndex).TextLang(LanguageIndex).Text = Trim$(txtText.Text)
    EditorChange = True
End Sub
