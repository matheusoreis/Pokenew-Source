Attribute VB_Name = "mStatus"
Option Explicit

Public Enum ColorText
    Black = 0
    Green
    Red
    Yellow
End Enum

Public Sub SetStatus(ByVal Text As String, Colour As ColorText)
    With frmEnc.lblStatus
        .Caption = Text

        Select Case Colour
        Case Black
            .ForeColor = &H0&
        Case Green
            .ForeColor = &HC000&
        Case Red
            .ForeColor = &HFF&
        Case Yellow
            .ForeColor = &HC0C0&
        End Select
    End With
End Sub
