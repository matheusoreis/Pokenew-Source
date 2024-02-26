Attribute VB_Name = "mEncrypt"
Option Explicit

Public Const DecExtension As String = ".PNG" '0
Public Const EncExtension As String = ".DAT" '1

Public GlobalDir As String

Public Enum ExtensionEnum
    PNG '0
    DAT '1
End Enum

Sub Main()
    InitCryptographyKey
    
    frmEnc.Show
End Sub

Public Sub ConvertBinaryToPNG(sPath As String, ByRef I As Long, ByRef data() As Byte)
    Dim fso As Object
    Dim f As Long, l As Long, s As String, sFile As Object
    Dim sFolder As Object, decrypt() As Byte
    Dim NewName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each sFile In fso.GetFolder(sPath).Files
        If UCase(Right(sFile.Name, 4)) = EncExtension Then
            I = I + 1
            f = FreeFile
            
            Open sFile.Path For Binary As #f
            Get #f, , l
            ReDim data(l)
            Get #f, , data ' Pega os bytes
            Close #f
            
            ' Inicia o Decrypt AES 128bits
            decrypt = DecryptFile(data, l)
            
            NewName = Replace$(sFile.Path, LCase(EncExtension), LCase(DecExtension))
            
            SaveFile NewName, decrypt
            Kill sFile.Path

            Erase data
        End If
        
        DoEvents
    Next sFile

    For Each sFolder In fso.GetFolder(sPath).SubFolders
        If LCase(sFolder.Name) <> "fonts" Then
            ConvertBinaryToPNG sFolder.Path, I, data()
        End If
    Next sFolder
    
    ' Atualiza o diretório
    RefreshDir
    
    Call SetStatus("OK!!!", Green)

    Set fso = Nothing
End Sub

Public Sub ConvertPNGToBinary(sPath As String, ByRef I As Long, ByRef data() As Byte)
    Dim fso As Object
    Dim f As Long, l As Long, s As String, sFile As Object, NewName As String
    Dim sFolder As Object, encrypt() As Byte, length As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each sFile In fso.GetFolder(sPath).Files
        If UCase(Right(sFile.Name, 4)) = DecExtension Then
            I = I + 1
            f = FreeFile
            Open sFile.Path For Binary As #f
        
            ReDim data(LOF(f) - 1)
            Get #f, , data ' Pega os bytes
            Close #f
        
            ' Inicia a encryptação AES 128bits
            encrypt = EncryptFile(data, (UBound(data) - LBound(data)) + 1)
            length = (UBound(encrypt) - LBound(encrypt)) + 1

            NewName = Replace$(sFile.Path, LCase(DecExtension), LCase(EncExtension))
            '//Correção
            If NewName = sFile.Path Then
                NewName = Replace$(sFile.Path, UCase(DecExtension), LCase(EncExtension))
            End If
            
            SaveBinary NewName, encrypt, length
            Kill sFile.Path
            Erase data
        End If
        DoEvents
    Next sFile

    For Each sFolder In fso.GetFolder(sPath).SubFolders
        ' Ignora a pasta com o nome "fonts"
        Debug.Print sFolder.Name
        If LCase(sFolder.Name) <> "fonts" Then
            ConvertPNGToBinary sFolder.Path, I, data()
        End If
    Next sFolder

    ' Atualiza o diretório
    RefreshDir

    Call SetStatus("OK!!!", Green)

    Set fso = Nothing
End Sub


Public Sub SaveBinary(fileName As String, ByRef data() As Byte, ByVal dataLength As Long)
    ' OPS
    ' Open fileNome For Binary As #F
    
    Dim f As Long
    f = FreeFile
    Open fileName For Binary As #f
        'Put #f, , "ss" ' Camada de string
        'Put #f, , 12 ' Camada de 32 bit
        Put #f, , dataLength ' Tamanho da array
        Put #f, , data ' Array de byte
    Close #f
End Sub

Public Sub SaveFile(fileName As String, ByRef data() As Byte)
    Dim f As Long
    f = FreeFile
    
    'Debug.Print fileName
    Open fileName For Binary As #f
        Put #f, 1, data
    Close #f
End Sub

Public Function FileExist(ByVal fileName As String, Optional RAW As Boolean = False) As Boolean
    FileExist = False
    If LenB(Dir(fileName)) > 0 Then FileExist = True
End Function

Public Sub RefreshDir()
    frmEnc.dirAppPath.Refresh
    frmEnc.flbAppPath.Refresh
End Sub
