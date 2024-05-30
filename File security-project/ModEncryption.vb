Imports System.Security.Cryptography
Imports System.IO
Module ModEncryption
    Public clientDESCryptoServiceProvider As New TripleDESCryptoServiceProvider
    'Returns the Salt to be saved into the stream
    Public Function SetEncKey(ByVal pwd As String) As Byte()
        Dim drive As Long
        Dim rnd As Random

        rnd = New Random

        Dim byts(7) As Byte '8 bytes
        rnd.NextBytes(byts)

        SetKey(pwd, byts)

        Return byts
    End Function

    Public Sub SetDecKey(ByVal pwd As String, ByVal salt As Byte())
        Dim drive As Long
        SetKey(pwd, salt)
    End Sub

    Public Sub SetKey(ByVal pwd As String, ByVal salt As Byte())
        Dim deriver As PasswordDeriveBytes
        deriver = New PasswordDeriveBytes(pwd, salt)
        clientDESCryptoServiceProvider.Key = deriver.GetBytes(clientDESCryptoServiceProvider.LegalKeySizes(0).MaxSize \ 8)
        clientDESCryptoServiceProvider.IV = deriver.GetBytes(clientDESCryptoServiceProvider.BlockSize \ 8)

    End Sub

    Public Sub EncFile(ByVal pwd As String, ByVal fin As String, ByVal feout As String)
        Dim fileStream As FileStream
        fileStream = File.OpenWrite(feout)

        Dim writeStream As BinaryWriter
        writeStream = New BinaryWriter(fileStream)

        Dim myByte As Byte
        Dim salt() As Byte
        salt = SetEncKey(pwd)

        'Write data to file
        Try
            'Write salt to file
            Dim i As Integer
            For i = 0 To salt.Length - 1
                writeStream.Write(salt(i))
            Next
            writeStream.Flush()
        Catch
        End Try

        'Create the encrypter
        Dim encryptor As ICryptoTransform
        encryptor = clientDESCryptoServiceProvider.CreateEncryptor()

        'Create the Encrypter Stream
        Dim encStream As CryptoStream
        encStream = New CryptoStream(fileStream, encryptor, CryptoStreamMode.Write)

        'Read the input file
        Dim readStream As BinaryReader
        readStream = New BinaryReader(File.OpenRead(fin))

        Try
            Do
                myByte = readStream.ReadByte()
                encStream.WriteByte(myByte)
            Loop
        Catch
            encStream.FlushFinalBlock()
            encStream.Flush()
        End Try

        encStream.Close()
        fileStream.Close()
        readStream.Close()
    End Sub

    Public Sub DecFile(ByVal pwd As String, ByVal fin As String, ByVal fdout As String)
        Dim fileStream As FileStream
        fileStream = File.OpenRead(fin)

        'Load Salt
        Dim salt(7) As Byte
        Try
            Dim i As Integer
            For i = 0 To 7
                salt(i) = fileStream.ReadByte()
            Next
            SetDecKey(pwd, salt)
        Catch
            fileStream.Close()
            Return
        End Try

        'Create the decryptor
        Dim encryptor As ICryptoTransform
        encryptor = clientDESCryptoServiceProvider.CreateDecryptor()

        'Create the decryptor stream
        Dim decStream As CryptoStream
        decStream = New CryptoStream(fileStream, encryptor, CryptoStreamMode.Read)

        'Stream to read the from decryptor stream
        Dim readStream As BinaryReader
        readStream = New BinaryReader(decStream)

        'The output file
        Dim writeStream As BinaryWriter
        writeStream = New BinaryWriter(File.OpenWrite(fdout))

        Dim myByte As Byte
        Try
            Do
                myByte = readStream.ReadByte()
                writeStream.Write(myByte)
            Loop
        Catch
            writeStream.Flush()
        End Try

        writeStream.Close()
        fileStream.Close()
        Try
            readStream.Close()
        Catch ex As Exception
        End Try
    End Sub

End Module
