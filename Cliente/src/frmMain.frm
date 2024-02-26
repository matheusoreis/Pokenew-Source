VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1950
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   87
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then Call IncomingData(bytesTotal)
End Sub

' *****************
' ** Form object **
' *****************
Private Sub Form_KeyPress(KeyAscii As Integer)
    FormKeyPress KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    FormKeyUp KeyCode, Shift
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMouseDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMouseMove Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMouseUp Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    If Not ForceExit Then
        If GUI(GuiEnum.GUI_CHOICEBOX).Visible Then Exit Sub
        
        If GUI(GuiEnum.GUI_INPUTBOX).Visible Then
            CloseInputBox
        End If
        If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
            GuiState GUI_GLOBALMENU, False
        End If
        If GUI(GuiEnum.GUI_OPTION).Visible Then
            GuiState GUI_OPTION, False
        End If
        OpenChoiceBox TextUIChoiceExit, CB_EXIT
    End If
End Sub
