Imports System.IO
Imports System.Security
Imports System.Text
Imports Microsoft.Management.Infrastructure
Imports Microsoft.Management.Infrastructure.Options

' Imports System.Deployment.Application

''' <summary>
''' Pacote de funções, métodos que apoiam com a tratativa de Arquivos. Essa Classe preferencialmente não gera Exceção.
''' </summary>
Public NotInheritable Class Arquivo



    Public Const MinDate As Date = #1/1/0001#


    Public Enum Codificacoes As Integer
        AUTO = 0
        Windows_1252 = 1252
        UTF16 = 1200
        UTF32 = 12000
        MACINTOSH = 10000
        US_ASCII = 20127
        ISO_8859_1 = 28591
        UTF_8 = 65001
    End Enum

    Public Shared Function CodificacaoToEncoding(ByVal Valor As Codificacoes) As Encoding
        Try
            Return Encoding.GetEncoding(Valor)
        Catch ex As ArgumentException
            Return Encoding.UTF8
        Catch ex As NotSupportedException
            Return Encoding.UTF8
        End Try
    End Function

    ''' <summary>
    ''' Função que grava um texto dentro de um arquivo com o Encoding iso-8859-1 (Por default adiciona no final)
    ''' </summary>
    ''' <param name="NomeArquivo">Nome do Arquivo (URL Completa)</param>
    ''' <param name="strTextoAGravar">Conteúdo do arquivo a adicionar (texto)</param>
    ''' <param name="modoArquivo">Modo de Gravação (Append or Create)</param>
    ''' <param name="modoAcessoArquivo">Forma de Gravação (FileAccess)</param>
    ''' <param name="TipoCodificacao">Tipo de Codificacao (padrão é UTF8)</param>
    ''' <returns>Retorna true se a operação teve sucesso, retorna falso em caso de erro</returns>
    <CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification:="<Pending>")>
    Public Shared Function GravarArquivoTexto(ByVal NomeArquivo As String, ByVal strTextoAGravar As String,
        Optional ByVal modoArquivo As FileMode = FileMode.Append,
        Optional ByVal modoAcessoArquivo As FileAccess = FileAccess.Write,
        Optional ByVal TipoCodificacao As Codificacoes = Codificacoes.UTF_8) As Boolean
        Try
            Dim fs As FileStream = New FileStream(NomeArquivo, modoArquivo, modoAcessoArquivo)
            Dim sw As StreamWriter = New StreamWriter(fs, System.Text.Encoding.GetEncoding(CInt(TipoCodificacao)))
            sw.Write(strTextoAGravar)
            sw.Close()
            fs.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function


    ''' <summary>
    ''' Função que grava um texto dentro de um arquivo (Cria se necessário, deleta conteudo e regrava)
    ''' </summary>
    ''' <param name="NomeArquivo"></param>
    ''' <param name="strTextoAGravar"></param>
    ''' <returns>Retorna true se a operação teve sucesso, retorna falso em caso de erro</returns>
    <CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification:="<Pending>")>
    Public Shared Function CriarGravarArquivoTexto(ByVal NomeArquivo As String, ByVal strTextoAGravar As String) As Boolean
        Try
            Return GravarArquivoTexto(NomeArquivo, strTextoAGravar, FileMode.Create, FileAccess.Write)
        Catch ex As Exception
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Função que processa o formato do diretório para garantir que a saida contenha um Path Valido e terminado por "\"
    ''' </summary>
    ''' <param name="diretorio"></param>
    ''' <returns></returns>
    Public Shared Function ProcessarFormatoDiretorio(ByVal diretorio As String) As String
        If diretorio.EndsWith("/") Or diretorio.EndsWith("\") Then
            Return diretorio
        Else
            Return String.Concat(diretorio, Path.DirectorySeparatorChar).Replace("/", Path.DirectorySeparatorChar).ToString
        End If
    End Function

    ''' <summary>
    ''' Obtem o tamanho em Bytes de um Arquivo em Disco
    ''' </summary>
    ''' <param name="NomeArquivo">Nome do arquivo</param>
    ''' <returns>Retorna o numero de Bytes ou -1 caso haja um erro ou arquivo inexiste</returns>
    Public Shared Function PegarTamanhoArquivo(ByVal NomeArquivo As String) As Long
        Try
            Dim info As New FileInfo(NomeArquivo)
            Return info.Length
        Catch ex As FileNotFoundException
            Return -1
        Catch ex As IOException
            Return -1
        End Try
    End Function


    ''' <summary>
    ''' Retorna uma string com o tamanho e unidade de medida de Arquivos em formato fácil de ler (KB/MB/GB)
    ''' </summary>
    ''' <param name="arq">Tamanho do Arquivo (em bytes)</param>
    ''' <returns>String com o texto em formato simples de ler</returns>
    Public Shared Function TamanhoArquivoHumano(ByVal arq As Long) As String
        If arq <= 1024 Then
            Return arq.ToString
        ElseIf arq <= 1048576 Then
            Return Math.Round(arq / 1024, 2) & " KB"
        ElseIf arq <= 1073741824 Then
            Return Math.Round(arq / 1048576, 2) & " MB"
        Else
            Return Math.Round(arq / 1073741824, 2) & " GB"
        End If
    End Function

    ''' <summary>
    ''' Obtem a quantidade de Bytes Livres em Disco
    ''' </summary>
    ''' <param name="drive">Drive a buscar, em formato string, por exemplo C:\  (não é permitido UNC)</param>
    ''' <returns>Quantidade de Bytes Livres ou retorna número negativo em caso de erro</returns>
    Public Shared Function ObterEspacoLivreDisco(ByVal drive As String) As Long
        Dim retorno As Long = -1
        Try
            Dim di As New DriveInfo(drive)
            If di.IsReady = True Then retorno = di.AvailableFreeSpace
        Catch ex As ArgumentNullException
            retorno = -1
        Catch ex As ArgumentException
            retorno = -1
        Catch ex As UnauthorizedAccessException
            retorno = -1
        Catch ex As IOException
            retorno = 1
        End Try
        Return retorno
    End Function

    ''' <summary>
    ''' Obtem a quantidade de Bytes Total em Disco
    ''' </summary>
    ''' <param name="drive">Drive a buscar, em formato string, por exemplo C:\  (não é permitido UNC)</param>
    ''' <returns>Quantidade de Bytes Totais do Disco ou retorna número negativo em caso de erro</returns>
    Public Shared Function ObterEspacoTotal(ByVal drive As String) As Long
        Dim retorno As Long = -1
        Try
            Dim di As New DriveInfo(drive)
            If di.IsReady = True Then retorno = di.TotalSize
        Catch ex As ArgumentNullException
        Catch ex As ArgumentException
        Catch ex As UnauthorizedAccessException
        Catch ex As IOException
        End Try

        Return retorno
    End Function


    Public Shared Function ObterDataUltimaModificacaoArquivo(ByVal LocalArquivo As String, Optional ByVal ValueOnException As Date = MinDate) As Date
        Dim resultado As DateTime = ValueOnException

        Try
            If File.Exists(LocalArquivo) Then resultado = File.GetLastWriteTime(LocalArquivo)
        Catch ex As UnauthorizedAccessException
        Catch ex As ArgumentException
        Catch ex As PathTooLongException
        Catch ex As NotSupportedException
        End Try

        Return resultado
    End Function



    Public Shared Function LerConteudoArquivo(ByVal LocalArquivo As String, Optional ByVal TipoCodificacao As FBR.Arquivo.Codificacoes = Codificacoes.UTF_8, Optional ByVal ValueOnException As String = "") As String
        ' Retorna Vazio caso dê erro
        Dim strRetorno As String = ValueOnException
        Try
            strRetorno = System.IO.File.ReadAllText(LocalArquivo, System.Text.Encoding.GetEncoding(CInt(TipoCodificacao)))
        Catch ex As ArgumentNullException
        Catch ex As IOException
            ' Console.WriteLine(ex.ToString)
        Catch ex As UnauthorizedAccessException
        Catch ex As SecurityException
        End Try
        Return strRetorno
    End Function



    Public Shared Function SHA1Arquivo(ByVal LocalArquivo As String) As String
        Dim strRetorno As String = ""
        If File.Exists(LocalArquivo) Then
            Dim hashAlgorithm As Cryptography.HashAlgorithm = New Cryptography.SHA1Managed
            Dim stream As New FileStream(LocalArquivo, FileMode.Open, FileAccess.Read, FileShare.Read) With {
                .Position = 0
            }
            Dim hash As Byte() = hashAlgorithm.ComputeHash(stream)
            hashAlgorithm.Clear()
            strRetorno = Texto.ByteArrayToHexString(hash).ToString.ToLower
        End If
        Return strRetorno
    End Function


End Class
