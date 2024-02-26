Attribute VB_Name = "mCryptography"
Option Explicit

Public Const CRYPTO_KEY_LENGTH As Long = 16

Private Const PADDING As Long = 64

Public Declare Function encrypt Lib "Encryptor.dll" Alias "Encrypt" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByRef DestLengthLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long
Public Declare Function decrypt Lib "Encryptor.dll" Alias "Decrypt" (ByVal SourceLngPtr As Long, ByVal SourceLength As Long, ByRef DestLngPtr As Long, ByVal KeyRef As Long, ByVal IvRef As Long) As Long

Public Key(0 To CRYPTO_KEY_LENGTH - 1) As Byte
Public IV(0 To CRYPTO_KEY_LENGTH - 1) As Byte

Public Sub InitCryptographyKey()
    Dim I As Long, SString As String, Num As Byte
    
    Num = 48

    ' Inicializa a string SString com caracteres específicos
    SString = SString & Chr(7 + Num) & Chr(5 + Num) & Chr(4 + Num) & Chr(6 + Num) & Chr(0 + Num) & Chr(0 + Num) & Chr(1 + Num) & Chr(3 + Num)

    ' Define a semente aleatória com base no valor calculado a partir da string SString
    Randomize CLng(SString)

    ' Gera valores aleatórios para a chave (Key) e o vetor de inicialização (IV)
    For I = 0 To CRYPTO_KEY_LENGTH - 1
        Key(I) = Int(Rnd * 256)
        IV(I) = Int(Rnd * 256)
    Next
End Sub

Public Function EncryptPacket(ByRef data() As Byte, ByVal dataLength As Long) As Byte()
    Dim EncryptedLength As Long
    Dim Encrypted() As Byte

    ReDim Encrypted(0 To dataLength + PADDING)

    EncryptedLength = encrypt(ByVal VarPtr(data(0)), dataLength, ByVal VarPtr(Encrypted(0)), ByVal VarPtr(EncryptedLength), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    ReDim Preserve Encrypted(0 To EncryptedLength - 1)

    EncryptPacket = Encrypted

End Function

Public Function DecryptPacket(ByRef data() As Byte, ByVal DataLengh As Long) As Byte()
    Dim Decrypted() As Byte
    Dim Count As Long

    ReDim Decrypted(0 To DataLengh - 1)

    Count = decrypt(ByVal VarPtr(data(0)), DataLengh, ByVal VarPtr(Decrypted(0)), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

    DecryptPacket = Decrypted
End Function

Public Function EncryptFile(ByRef data() As Byte, ByVal dataLength As Long) As Byte()
    Dim EncryptedLength As Long
    Dim Encrypted() As Byte

    ReDim Encrypted(0 To dataLength + PADDING)

    EncryptedLength = encrypt(ByVal VarPtr(data(0)), dataLength, ByVal VarPtr(Encrypted(0)), ByVal VarPtr(EncryptedLength), ByVal VarPtr(Key(0)), ByVal VarPtr(IV(0)))

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

