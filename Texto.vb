Imports System.Text

Public NotInheritable Class Texto


    '#Disable Warning CA2101 ' Specify marshaling for P/Invoke string arguments
    '    ' Call this function to remove the key from memory after it is used for security.
    '    Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal Destination As String, ByVal Length As Integer)
    '#Enable Warning CA2101 ' Specify marshaling for P/Invoke string arguments

    Public Enum Linguagem
        PORTUGUES = 1
        INGLES = 2
        ESPANHOL = 3
    End Enum
    Public Enum LCIDNumber As Integer
        EN_US = 1033
        ES_ES = 1034
        PT_BR = 1046
        EN_GB = 2057
        ES_MX = 2058
        ES_AR = 11274
    End Enum

    Public Enum PadraoChaveCriptografia
        DES
        AES
    End Enum


    Public Shared Function LCIDparaLinguagem(ByVal lingua As FBR.Texto.Linguagem) As LCIDNumber
        If lingua = Linguagem.INGLES Then
            Return LCIDNumber.EN_US
        ElseIf lingua = Linguagem.ESPANHOL Then
            Return LCIDNumber.ES_ES
        Else
            Return LCIDNumber.PT_BR
        End If
    End Function

    Shared Function IsReallyEmpty(ByVal texto As String) As Boolean
        If texto Is Nothing Then
            Return True
        Else
            If texto.Length <= 0 Or texto = "NULL" Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary> 
    ''' Método que remove acentos, espaços e carateres indesejados 
    ''' </summary> 
    ''' <param name="texto">Texto a ser Submetido a retirada dos Caracteres Especiais</param> 
    ''' <returns>O texto Formatado sem Caracteres Especiais</returns> 
    Public Shared Function RemoveAcentos(ByVal texto As String) As String
        Const ComAcentos As String = "ÄÅÁÂÀÃäáâàãÉÊËÈéêëèÍÎÏÌíîïìÖÓÔÒÕöóôòõÜÚÛüúûùÇç:ñÑ"
        Const SemAcentos As String = "AAAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUuuuuCc.nN"

        For i As Integer = 0 To ComAcentos.Length - 1
            texto = texto.Replace(ComAcentos(i).ToString(), SemAcentos(i).ToString()).Trim()
        Next
        Return texto
    End Function


    ''' <summary>
    ''' Função que conta a quantidade de vezes que um caracter em uma string
    ''' </summary>
    ''' <param name="TextoLongo">String / Texto Longo</param>
    ''' <param name="caracter">Caracter a buscar (1 char)</param>
    ''' <returns>Quantidade de Ocorrencias</returns>
    Public Shared Function ContarQuantidadeOcorrencias(ByVal TextoLongo As String, ByVal caracter As Char) As Integer
        Return TextoLongo.Count(Function(c As Char) c = caracter)
    End Function


    ''' <summary>
    ''' Recebe um arraylist contendo itens de string (arraylist of string) e devolve uma string concatenada (1 item por linha de texto)
    ''' </summary>
    ''' <param name="arrList">Arraylist</param>
    ''' <returns>Uma string concatenada, cada elemento em uma linha separada</returns>
    Public Shared Function ArraylistOfString(ByVal arrList As ArrayList) As String
        Dim sb As New StringBuilder(10000)
        For Each a In arrList
            sb.AppendLine(a.ToString)
        Next
        Return sb.ToString
    End Function

    Public Shared Function QuebrarEmPacotesString(ListaDeString As List(Of String), TamanhoDoPacote As Integer) As List(Of List(Of String))
        'SplitIntoChunks
        Return ListaDeString.Select(Function(x, i) New With {Key .Index = i, Key .Value = x}).
                GroupBy(Function(x) (x.Index \ TamanhoDoPacote)).
                Select(Function(x) x.Select(Function(v) v.Value).ToList()).
                ToList()
    End Function

    Public Shared Function Maiusculo(ByVal entrada As String) As String
        Return Strings.UCase(entrada)
    End Function

    Public Shared Function Minusculo(ByVal entrada As String) As String
        Return Strings.LCase(entrada)
    End Function

    Public Shared Function PrimeiraMaiuscula(ByVal entrada As String) As String

        ' cuida das preposições e artigos
        Dim qtdeCaracteres As Integer = entrada.Length
        Select Case qtdeCaracteres
            Case <= 0
                Return entrada
            Case <= 1
                Return entrada.ToUpper
            Case Else
                Dim saida As New StringBuilder(qtdeCaracteres + 1)

                Dim ultimoEspaco As Boolean = True
                For i As Integer = 0 To qtdeCaracteres - 1
                    Dim c As String = entrada.Substring(i, 1)
                    If ultimoEspaco Then
                        saida.Append(entrada.Substring(i, 1).ToUpper)
                    Else
                        saida.Append(entrada.Substring(i, 1).ToLower)
                    End If
                    If c = " " Then ultimoEspaco = True Else ultimoEspaco = False
                Next
                Return saida.ToString()
        End Select
    End Function

    Public Shared Function PrimeiraMaiuscula(ByVal entrada As String, Optional ByVal lingua As FBR.Texto.Linguagem = Linguagem.PORTUGUES) As String
        Dim PreposicoesArtigosPortugues() As String = {"a", "ante", "após", "até", "com", "contra", "de", "desde", "em", "entre", "para", "por", "perante", "sem", "sob", "sobre", "trás", "a", "as", "os", "as", "um", "uma", "uns", "umas"}
        Dim PreposicoesArtigosIngles() As String = {"at", "of", "by", "for", "in", "on", "to", "under", "until", "up", "over", "is", "a", "an", "the"}
        Dim PreposicoesArtigosEspanhol() As String = {"el", "la", "los", "las", "un", "unos", "una", "unas", "a", "ante", "bajo", "con", "contra", "de", "desde", "durante", "en", "entre", "hacia", "hasta", "mediante", "para", "por", "según", "sin", "so", "sobre", "tras", "versus"}
        Dim PreposicoesArtigos(50) As String

        ' escolha da lingua
        If lingua = Linguagem.INGLES Then
            PreposicoesArtigosIngles.CopyTo(PreposicoesArtigos, 0)
        ElseIf lingua = Linguagem.ESPANHOL Then
            PreposicoesArtigosEspanhol.CopyTo(PreposicoesArtigos, 0)
        Else
            PreposicoesArtigosPortugues.CopyTo(PreposicoesArtigos, 0)
        End If

        ' cuida das preposições e artigos
        Dim qtdeCaracteres As Integer = entrada.Length
        Select Case qtdeCaracteres
            Case <= 0
                Return entrada
            Case <= 1
                Return entrada.ToUpper
            Case Else
                Dim saida As New StringBuilder(qtdeCaracteres + 1)
                Dim partes As String() = Split(entrada, " ")
                For Each parte As String In partes
                    Dim resultado As String = Array.Find(PreposicoesArtigos, Function(s) s = parte)
                    If String.IsNullOrEmpty(resultado) Then
                        saida.Append(parte.Substring(0, 1).ToUpper)
                        saida.Append(parte.Substring(1).ToLower)
                        saida.Append(" ")
                    Else
                        'Encontrado então não adiciona
                        saida.Append(parte.ToString.ToLower)
                        saida.Append(" ")
                    End If
                Next
                Return saida.ToString().Trim
        End Select
    End Function


    Public Shared Function LPAD(ByVal entrada As String, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return entrada.ToString.PadLeft(qtdeDigitosTotal, System.Convert.ToChar(Enchimento))
    End Function
    Public Shared Function PadLeft(ByVal entrada As Integer, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return LPAD(entrada, qtdeDigitosTotal, Enchimento)
    End Function
    ''' <summary>
    ''' Repete uma string, numero ou texto por uma determinada quantidade de vezes.
    ''' </summary>
    ''' <param name="num">Número a repetir</param>
    ''' <param name="qtde">Quantidade de repetições</param>
    ''' <returns>Retorna uma string concatenada</returns>
    Public Shared Function Repetir(ByVal num As Integer, ByVal qtde As Integer) As String
        Return Repetir(CStr(num), qtde)
    End Function
    ''' <summary>
    ''' Repete uma string, numero ou texto por uma determinada quantidade de vezes.
    ''' </summary>
    ''' <param name="str">Texto a repetir</param>
    ''' <param name="qtde">Quantidade de repetições</param>
    ''' <returns>Retorna String.Empty caso qtde zero, ou string concatenada caso sucesso</returns>
    Public Shared Function Repetir(ByVal str As String, ByVal qtde As Integer) As String
        If qtde <= 0 Then
            Return String.Empty
        ElseIf qtde = 1 Then
            Return str
        ElseIf str.Length = 1 Then
            Return Strings.StrDup(qtde, str)
        Else

            Dim resultado As New StringBuilder(str.Length * qtde + 5)
            For i As Integer = 1 To qtde
                resultado.Append(str)
            Next
            Return resultado.ToString
        End If
    End Function
    ''' <summary>
    ''' Repete uma string, numero ou texto por uma determinada quantidade de vezes.
    ''' </summary>
    ''' <param name="c">Caracter a repetir</param>
    ''' <param name="qtde">Quantidade de repetições</param>
    ''' <returns>Retorna String.Empty caso qtde zero, ou string concatenada caso sucesso</returns>
    Public Shared Function Repetir(ByVal c As Char, ByVal qtde As Integer) As String
        If qtde <= 0 Then
            Return String.Empty
        ElseIf qtde = 1 Then
            Return CStr(c)
        Else
            Return New String(c, qtde)
        End If
    End Function

    Public Shared Function SHA1paraByte(ByVal texto As String) As Byte()

        'Cria uma nova instância de UnicodeEncoding para 'converter a string em um array de bytes Unicode
        Dim UE As New ASCIIEncoding()
        'Converte o stringem um array de bytes.
        Dim MessageBytes As Byte() = UE.GetBytes(texto)
        'Cria uma instância de SHA1Managed para criar um valor  hash
        Dim SHhash As New System.Security.Cryptography.SHA1Managed()
        'Cria o valor  hash
        Dim SHA1HASHValue() As Byte = SHhash.ComputeHash(MessageBytes)
        Return SHA1HASHValue

    End Function

    Public Shared Function SHA1(ByVal texto As String) As String
        Return ByteToString(SHA1paraByte(texto))
    End Function
    Public Shared Function SHA1paraHexString(ByVal texto As String) As String
        Return ByteArrayToHexString(SHA1paraByte(texto))
    End Function
    Public Shared Function SHA256paraByte(ByVal texto As String) As Byte()

        'Cria uma nova instância de UnicodeEncoding para 'converter a string em um array de bytes Unicode
        Dim UE As New ASCIIEncoding()
        'Converte o stringem um array de bytes.
        Dim MessageBytes As Byte() = UE.GetBytes(texto)
        'Cria uma instância de SHA1Managed para criar um valor  hash
        Dim SHhash As New System.Security.Cryptography.SHA256Managed()
        'Cria o valor  hash
        Dim SHA256HASHValue() As Byte = SHhash.ComputeHash(MessageBytes)
        Return SHA256HASHValue

    End Function
    Public Shared Function SHA256(ByVal texto As String) As String
        Return ByteToString(SHA256paraByte(texto))
    End Function

    Public Shared Function ByteToString(ByVal bytes As Byte()) As String

        Dim s As New StringBuilder(bytes.Length)
        For Each b As Byte In bytes
            s.AppendFormat(ChrW(b))
        Next b
        Return s.ToString

    End Function

    Public Shared Function ByteArrayToHexString(ByVal b As Byte()) As String
        Dim saida As New StringBuilder(40)
        For Each a In b
            'saida.AppendFormat(LPAD(Hex(a), 2))
            saida.AppendFormat("{0:X2}", a)
        Next
        Return saida.ToString.ToLower
    End Function

    Public Shared Function StringToHexString(ByVal s As String) As String
        Dim saida As New StringBuilder(s.Length * 2)
        For Each a As String In s
            saida.AppendFormat("{0:X2}", Hex(Strings.Asc(a)))
        Next
        Return saida.ToString.ToLower
    End Function


    Public Shared Function StringToByteArray(ByVal s As String) As Byte()
        Return HexStringToByteArray(StringToHexString(s))
    End Function



    Public Shared Function HexStringToString(ByVal s As String) As String
        Dim TamEsperadoSaida As Integer = s.Length / 2
        Dim sbSaida As New StringBuilder(TamEsperadoSaida)
        For x = 0 To TamEsperadoSaida - 1
            sbSaida.Append(ChrW(CInt("&H" & s.Substring(x * 2, 2))))
        Next
        Return sbSaida.ToString
    End Function

    Public Shared Function HexStringToByteArray(ByVal s As String) As Byte()
        Dim TamEsperadoSaida As Integer = s.Length / 2
        Dim arrSaida(TamEsperadoSaida) As Byte
        For x As Integer = 0 To TamEsperadoSaida - 1
            arrSaida(x) = Convert.ToByte(CInt("&H" & s.Substring(x * 2, 2)))
        Next
        Return arrSaida


        ' Return Convert.ToBase64String(s, 0, s.Length, Base64FormattingOptions.InsertLineBreaks)
        ' Dim st As String = "a9993e364706816aba3e25717850c26c9cd0d89d"
        ' Dim com As String = ""
    End Function


    'Funções AES
    ' Trocado de ASCII 7 bits para UTF8
    ' Fonte: http://stackoverflow.com/questions/5987186/aes-encrypt-string-in-vb-net

    Public Shared Function AES_Criptografar(ByVal input As String, ByVal pass As String, Optional ByVal tipoCodificacao As Arquivo.Codificacoes = Arquivo.Codificacoes.US_ASCII) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim encrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Array.Empty(Of Byte)()

            If tipoCodificacao = Arquivo.Codificacoes.UTF_8 Then
                temp = Hash_AES.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pass))
            ElseIf tipoCodificacao = Arquivo.Codificacoes.US_ASCII Or tipoCodificacao = Arquivo.Codificacoes.AUTO Then
                temp = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            End If

            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = System.Security.Cryptography.CipherMode.ECB
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = Array.Empty(Of Byte)()

            If tipoCodificacao = Arquivo.Codificacoes.UTF_8 Then
                Buffer = System.Text.Encoding.UTF8.GetBytes(input)
            ElseIf tipoCodificacao = Arquivo.Codificacoes.US_ASCII Or tipoCodificacao = Arquivo.Codificacoes.AUTO Then
                Buffer = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
            End If

            'encrypted = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            encrypted = ByteArrayToHexString(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))

        Catch ex As Exception
        End Try

        Return encrypted

    End Function

    Public Shared Function AES_Descriptografar(ByVal input As String, ByVal pass As String, Optional ByVal tipoCodificacao As Arquivo.Codificacoes = Arquivo.Codificacoes.US_ASCII) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim decrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = {}
            If tipoCodificacao = Arquivo.Codificacoes.UTF_8 Then
                temp = Hash_AES.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pass))
            ElseIf tipoCodificacao = Arquivo.Codificacoes.US_ASCII Or tipoCodificacao = Arquivo.Codificacoes.AUTO Then
                temp = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            End If

            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = System.Security.Cryptography.CipherMode.ECB
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = HexStringToByteArray(input)
            If tipoCodificacao = Arquivo.Codificacoes.UTF_8 Then
                decrypted = System.Text.Encoding.UTF8.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            ElseIf tipoCodificacao = Arquivo.Codificacoes.US_ASCII Or tipoCodificacao = Arquivo.Codificacoes.AUTO Then
                decrypted = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            End If
        Catch ex As Exception
        End Try
        Return decrypted
    End Function



    ''' <summary>
    ''' ClearBuffer: Clears a byte array to ensure that it cannot be read from memory
    ''' </summary>
    ''' <param name="bytes">Variavel a ser limpa (bytearray)</param>
    Public Shared Sub ClearBuffer(ByVal bytes() As Byte)
        If bytes Is Nothing Then Exit Sub
        For i As Integer = 0 To 9
            For n As Integer = 0 To bytes.Length - 1
                bytes(n) = i
            Next
        Next
    End Sub







    ''' <summary>
    ''' Gerar uma Chave Criptografica Simétrica (DES ou AES)
    ''' </summary>
    ''' <param name="Tipo">Tipo da Chave a gerar</param>
    ''' <returns>Uma String de Hexadecimal com 8 bytes (DES) ou 32 bytes (AES)</returns>
    Public Shared Function GerarChave(Optional ByVal Tipo As PadraoChaveCriptografia = PadraoChaveCriptografia.DES) As String
        If Tipo = PadraoChaveCriptografia.AES Then
            Return ByteArrayToHexString(Security.Cryptography.AesCryptoServiceProvider.Create().Key)
        Else
            Dim desCrypto As Security.Cryptography.DESCryptoServiceProvider = Security.Cryptography.DESCryptoServiceProvider.Create()
            Return ByteArrayToHexString(desCrypto.Key)

        End If
    End Function



    Function ListAggregate(ByVal arrEntrada As ArrayList, Optional ByVal StringSeparator As String = ", ") As String
        Dim s As New System.Text.StringBuilder(4096)
        Dim q As Integer = arrEntrada.Count
        Dim i As Integer
        For i = 0 To (q - 1)
            s.Append(arrEntrada(i).ToString)
            If i < (q - 1) Then s.Append(StringSeparator)
        Next
        Return s.ToString
    End Function

    Function FormataTempoPorExtenso(ByVal intEntrada As Double) As String
        ' Recebe a entrada em segundos
        ' retorna por extenso: "1 hora, 12 minutos e 15 segundos"

        Dim intTempo As Integer = intEntrada
        Dim tmpParcial As Integer
        Dim strSaida As String = ""
        If IsNumeric(intTempo) Then
            If (intTempo >= 3600) Then
                tmpParcial = ((intTempo - (intTempo Mod 3600)) / 3600)
                strSaida = tmpParcial & " hora"
                If tmpParcial > 1 Then strSaida = strSaida & "s"
                intTempo = intTempo - (tmpParcial * 3600)

                If intTempo > 0 Then
                    strSaida = strSaida & ", "
                End If
            End If

            If (intTempo >= 60) Then
                tmpParcial = ((intTempo - (intTempo Mod 60)) / 60)
                strSaida = strSaida & tmpParcial & " minuto"
                If tmpParcial > 1 Then strSaida = strSaida & "s"
                intTempo = intTempo - (tmpParcial * 60)

                If intTempo > 0 Then
                    strSaida = strSaida & " e "
                End If
            End If

            If (intTempo > 0) Then

                If intTempo = 1 Then strSaida = strSaida & intTempo & " segundo"
                If intTempo > 1 Then strSaida = strSaida & intTempo & " segundos"
            End If
        End If
        If strSaida.Equals("") Then
            strSaida = "Desconhecido"
        End If
        Return strSaida
    End Function

    Public Shared Function TextoEntreDoisTextosXML(ByVal entrada As String, ByVal tag As String) As String
        Return TextoEntreDoisTextos(entrada, String.Format("<{0}>", tag), String.Format("</{0}>", tag))

    End Function


    Public Shared Function TextoEntreDoisTextos(ByVal entrada As String, ByVal inicio As String, ByVal fim As String) As String
        Dim a As Integer = InStr(entrada, inicio) + inicio.Length
        Dim parte2Texto As String = Mid(entrada, a)
        Dim b As Integer = InStr(parte2Texto, fim)
        Dim saida As String
        If a > 0 And b > 0 Then
            saida = Left(parte2Texto, b - 1)
        Else
            saida = ""
        End If
        Return saida
    End Function

End Class
