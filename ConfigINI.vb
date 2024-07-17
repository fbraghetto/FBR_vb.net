Imports System.IO
''' <summary>
''' Classe que Trata Arquivos de Configuração
''' </summary>
Public NotInheritable Class ConfigINI
#Disable Warning CA2101 ' Specify marshaling for P/Invoke string arguments
    Private Declare Auto Function GetPrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Auto Function WritePrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
#Enable Warning CA2101 ' Specify marshaling for P/Invoke string arguments

    Private Const MAX_LENGTH As Integer = 500

    ''' <summary>
    ''' Obtem o conteúdo de uma entrada de um arquivo ini básico
    ''' </summary>
    ''' <param name="PathArquivoINI">FQDN do Arquivo INI (incluindo o nome do arquivo e extensão)</param>
    ''' <param name="Secao">Seção para Buscar</param>
    ''' <param name="Item">Item na seção para retornar</param>
    ''' <returns>Retorna conteúdo FILENOTFOUND se arquivo não encontrado, String.Empty caso item não encontrado, ou conteúdo em string do item</returns>
    Public Shared Function GetArquivoIni(ByVal PathArquivoINI As String, ByVal Secao As String, ByVal Item As String) As String

        If File.Exists(PathArquivoINI) Then
            Dim str_default As String = " "
            Dim str_builder As New Text.StringBuilder(MAX_LENGTH)
            GetPrivateProfileString(Secao, Item, str_default, str_builder, MAX_LENGTH, PathArquivoINI)
            If str_builder.ToString() = "NULL" Then
                Return String.Empty
            Else
                Return str_builder.ToString()
            End If
        Else
            Return "FILENOTFOUND: " & PathArquivoINI
        End If
    End Function

    ''' <summary>
    ''' Grava o conteúdo em uma entrada de um arquivo ini básico
    ''' </summary>
    ''' <param name="PathArquivoINI">FQDN do Arquivo INI (incluindo o nome do arquivo e extensão)</param>
    ''' <param name="Secao">Seção para Buscar</param>
    ''' <param name="Item">Item na seção para retornar, caso não existir, será criado</param>
    ''' <param name="Valor">Valor do Item a ser gravado</param>
    ''' <returns>Retorna False caso houve insucesso ou arquivo não encontrado, Retorna True em caso de sucesso</returns>
    ''' <remarks>Usa a função WritePrivateProfileString, ao qual é deprecated e é mantido por compatibilidade do Windows (Winbase.h  WritePrivateProfileStringA function)</remarks>
    Public Shared Function SetArquivoIni(ByVal PathArquivoINI As String, ByVal Secao As String, ByVal Item As String, ByVal Valor As String) As Boolean
        If File.Exists(PathArquivoINI) Then
            If WritePrivateProfileString(Secao, Item, Valor, PathArquivoINI) = 1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

End Class
