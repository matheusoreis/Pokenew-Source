Attribute VB_Name = "modCryptography"
Option Explicit

Public Const CRYPTO_KEY_LENGTH As Long = 16

Private Const PADDING As Long = 64

Public Declare Function Encrypt Lib "Encryptor.dll" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByRef DestLengthLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long
Public Declare Function decrypt Lib "Encryptor.dll" Alias "Decrypt" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long

Public Key(0 To CRYPTO_KEY_LENGTH - 1) As Byte
Public IV(0 To CRYPTO_KEY_LENGTH - 1) As Byte

Public Sub InitCryptographyKey()
    Dim i As Long, SString As String, Num As Byte

    Num = 48

    ' Inicializa a string SString com caracteres específicos
    SString = SString & Chr(7 + Num) & Chr(5 + Num) & Chr(4 + Num) & Chr(6 + Num) & Chr(0 + Num) & Chr(0 + Num) & Chr(1 + Num) & Chr(3 + Num)

    ' Define a semente aleatória com base no valor calculado a partir da string SString
    Randomize CLng(SString)

    ' Gera valores aleatórios para a chave (Key) e o vetor de inicialização (IV)
    For i = 0 To CRYPTO_KEY_LENGTH - 1
        Key(i) = Int(Rnd * 256)
        IV(i) = Int(Rnd * 256)
    Next
End Sub

Public Function EncryptPacket(ByRef data() As Byte, ByVal dataLength As Long) As Byte()
    Dim EncryptedLength As Long
    Dim Encrypted() As Byte

    ReDim Encrypted(0 To dataLength + PADDING)

    ' Criptografa os dados usando a função de criptografia externa e armazena o comprimento criptografado
    EncryptedLength = Encrypt(ByVal VarPtr(data(0)), dataLength, ByVal VarPtr(Encrypted(0)), ByVal VarPtr(EncryptedLength), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    ' Redimensiona o array de dados criptografados para corresponder ao comprimento real
    ReDim Preserve Encrypted(0 To EncryptedLength - 1)

    ' Retorna os dados criptografados
    EncryptPacket = Encrypted

End Function

Public Function DecryptPacket(ByRef data() As Byte, ByVal DataLengh As Long) As Byte()
    Dim Decrypted() As Byte
    Dim Count As Long

    ' Redimensiona o array de dados descriptografados para corresponder ao tamanho dos dados recebidos
    ReDim Decrypted(0 To DataLengh - 1)

    ' Descriptografa os dados usando a função de descriptografia externa e armazena a contagem de bytes descriptografados
    Count = decrypt(ByVal VarPtr(data(0)), DataLengh, ByVal VarPtr(Decrypted(0)), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    ' Retorna os dados descriptografados
    DecryptPacket = Decrypted
End Function

Public Function EncryptFile(ByRef data() As Byte, ByVal dataLength As Long) As Byte()
    Dim EncryptedLength As Long
    Dim Encrypted() As Byte

    ReDim Encrypted(0 To dataLength + PADDING)

    EncryptedLength = Encrypt(ByVal VarPtr(data(0)), dataLength, ByVal VarPtr(Encrypted(0)), ByVal VarPtr(EncryptedLength), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    ReDim Preserve Encrypted(0 To EncryptedLength - 1)

    EncryptFile = Encrypted

End Function

Public Function DecryptFile(ByRef data() As Byte, ByVal DataLengh As Long) As Byte()
    Dim Decrypted() As Byte
    Dim Count As Long

    ReDim Decrypted(0 To DataLengh - 1)

    Count = decrypt(ByVal VarPtr(data(0)), DataLengh, ByVal VarPtr(Decrypted(0)), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    DecryptFile = Decrypted
End Function

