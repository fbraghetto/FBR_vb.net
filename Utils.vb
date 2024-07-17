Imports System.Runtime

Public NotInheritable Class Utils
    'DECODE compares expr to each search value one by one. If expr is equal to a search, then Oracle Database returns the corresponding result. If no match is found, then Oracle returns default. If default is omitted, then Oracle returns null.
    'https://docs.oracle.com/cd/B19306_01/server.102/b14200/functions040.htm
    'DECODE (warehouse_id, 1, 'Southlake', 2, 'San Francisco', 3, 'New Jersey', 4, 'Seattle','Non domestic') 
    'DECODE( expression , search , result [, search , result]... [, default] )


    ''' <summary>
    ''' Função Decode (padrão Oracle) que compara a expressão com a busca. Se não houver match, retorna o Padrão, Caso o padrão for omitido, retorna String.Empty
    ''' </summary>
    ''' <param name="ListaItens">Expressão para avaliação inicial, seguido de item 1 para avaliar, com a resposta 1 e assim subsequentemente. Retorna default caso nenhum seja encontrado</param>
    ''' <returns>Retorna Resultado caso encontrado e/ou String.Empty (para parametros em quantidade ímpar) ou Valor Default caso não encontrado (para parametros em quantidade par)</returns>
    Public Shared Function Decode(ParamArray ByVal ListaItens() As Object) As Object
        'Função Decode (padrão Oracle) que compara a expressão com a busca. Se não houver match, retorna o Padrão, Caso o padrão for omitido, retorna String.Empty
        ' Prerequisitos ter 3 ou mais itens
        Dim qtdeTotal As Integer = ListaItens.Length
        Dim temPadrao As Boolean = ((qtdeTotal Mod 2) = 0)
        Dim _padrao As String = String.Empty
        If temPadrao Then
            _padrao = ListaItens(qtdeTotal - 1)
        End If


        If (qtdeTotal < 3) Then
            Return String.Empty
        Else
            Dim _ItemProcurado As Object = ListaItens(0)
            Dim i As Integer
            For i = 1 To qtdeTotal - 2 Step 2
                If ListaItens(i).Equals(_ItemProcurado) Then
                    Return ListaItens(i + 1)
                End If
            Next
        End If

        Return _padrao
    End Function


    ' 3 campos
    ''' <summary>
    ''' Função Decode (padrão Oracle) que compara a expressão com a busca. Se não houver match, retorna o Padrão, Caso o padrão for omitido, retorna String.Empty
    ''' </summary>
    ''' <param name="Expressao">Expressão para avaliação Inicial</param>
    ''' <param name="Busca1">Dados para comparação</param>
    ''' <param name="Resultado1">Resultado caso a comparação anterior for verdadeira</param>
    ''' <returns>Retorna Resultado caso encontrado ou String.Empty caso não encontrado</returns>
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        Else
            Return String.Empty
        End If
    End Function
    ' 4 campos
    ''' <summary>
    ''' Função Decode (padrão Oracle) que compara a expressão com a busca. Se não houver match, retorna o Padrão
    ''' </summary>
    ''' <param name="Expressao">Expressão para avaliação Inicial</param>
    ''' <param name="Busca1">Dados para comparação</param>
    ''' <param name="Resultado1">Resultado caso a comparação anterior for verdadeira</param>
    ''' <param name="Padrao">Padrão de retorno caso não seja encontrado</param>
    ''' <returns>Retorna Resultado caso encontrado ou Padrao caso não encontrado</returns>
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object, ByVal Padrao As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        Else
            Return Padrao
        End If
    End Function
    ' 5 campos
    ''' <summary>
    ''' Função Decode (padrão Oracle) que compara a expressão com a busca. Se não houver match, retorna o Padrão, Caso o padrão for omitido, retorna String.Empty
    ''' </summary>
    ''' <param name="Expressao">Expressão para avaliação Inicial</param>
    ''' <param name="Busca1">Dados para comparação</param>
    ''' <param name="Resultado1">Resultado caso a comparação anterior for verdadeira</param>
    ''' <param name="Busca2">Dados para comparação</param>
    ''' <param name="Resultado2">Resultado caso a comparação anterior for verdadeira</param>
    ''' <returns>Retorna Resultado caso encontrado ou String.Empty caso não encontrado</returns>
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object, ByVal Busca2 As Object, ByVal Resultado2 As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        ElseIf Expressao.Equals(Busca2) Then
            Return Resultado2
        Else
            Return String.Empty
        End If
    End Function
    ' 6 campos
    ''' <summary>
    ''' Função Decode (padrão Oracle) que compara a expressão com a busca. Se não houver match, retorna o Padrão
    ''' </summary>
    ''' <param name="Expressao">Expressão para avaliação Inicial</param>
    ''' <param name="Busca1">Dados para comparação</param>
    ''' <param name="Resultado1">Resultado caso a comparação anterior for verdadeira</param>
    ''' <param name="Busca2">Dados para comparação</param>
    ''' <param name="Resultado2">Resultado caso a comparação anterior for verdadeira</param>
    ''' <param name="Padrao">Retorna Resultado caso encontrado ou Padrao caso não encontrado</param>
    ''' <returns></returns>
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object, ByVal Busca2 As Object, ByVal Resultado2 As Object, ByVal Padrao As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        ElseIf Expressao.Equals(Busca2) Then
            Return Resultado2
        Else
            Return Padrao
        End If
    End Function

    ' 7 campos
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object,
                                  ByVal Busca2 As Object, ByVal Resultado2 As Object,
                                  ByVal Busca3 As Object, ByVal Resultado3 As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        ElseIf Expressao.Equals(Busca2) Then
            Return Resultado2
        ElseIf Expressao.Equals(Busca3) Then
            Return Resultado3
        Else
            Return String.Empty
        End If
    End Function
    ' 8 campos
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object,
                                  ByVal Busca2 As Object, ByVal Resultado2 As Object,
                                  ByVal Busca3 As Object, ByVal Resultado3 As Object,
                                  ByVal Padrao As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        ElseIf Expressao.Equals(Busca2) Then
            Return Resultado2
        ElseIf Expressao.Equals(Busca3) Then
            Return Resultado3
        Else
            Return Padrao
        End If
    End Function
    ' 9 campos
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object,
                                  ByVal Busca2 As Object, ByVal Resultado2 As Object,
                                  ByVal Busca3 As Object, ByVal Resultado3 As Object,
                                  ByVal Busca4 As Object, ByVal Resultado4 As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        ElseIf Expressao.Equals(Busca2) Then
            Return Resultado2
        ElseIf Expressao.Equals(Busca3) Then
            Return Resultado3
        ElseIf Expressao.Equals(Busca4) Then
            Return Resultado4
        Else
            Return String.Empty
        End If
    End Function
    ' 10 campos
    Public Shared Function Decode(ByVal Expressao As Object, ByVal Busca1 As Object, ByVal Resultado1 As Object,
                                  ByVal Busca2 As Object, ByVal Resultado2 As Object,
                                  ByVal Busca3 As Object, ByVal Resultado3 As Object,
                                  ByVal Busca4 As Object, ByVal Resultado4 As Object,
                                  ByVal Padrao As Object) As Object
        If Expressao.Equals(Busca1) Then
            Return Resultado1
        ElseIf Expressao.Equals(Busca2) Then
            Return Resultado2
        ElseIf Expressao.Equals(Busca3) Then
            Return Resultado3
        ElseIf Expressao.Equals(Busca4) Then
            Return Resultado4
        Else
            Return Padrao
        End If
    End Function



    ''' <summary>
    ''' Estrutura de Resultado para Retorno da Função de Consulta CEP
    ''' </summary>
    Public Class CEP
        Public Sucesso As Boolean = False
        Public CEP As String = "00000000"
        Public Rua As String = String.Empty
        Public Bairro As String = String.Empty
        Public Cidade As String = String.Empty
        Public UF As String = String.Empty
        Public IBGE As String = String.Empty
        Public DDD As String = String.Empty
        Sub New(ByVal _CEP As String, ByVal _Rua As String, ByVal _Bairro As String, ByVal _Cidade As String, ByVal _UF As String, ByVal _IBGE As String, ByVal _DDD As String, Optional ByVal _Sucesso As Boolean = True)
            CEP = _CEP
            Rua = _Rua
            Bairro = _Bairro
            Cidade = _Cidade
            UF = _UF
            IBGE = _IBGE
            DDD = _DDD
            Sucesso = _Sucesso
        End Sub
        Sub New(Optional ByVal _Sucesso As Boolean = False)
            Sucesso = _Sucesso
            CEP = Texto.Repetir(0, 8)
        End Sub
    End Class


    ''' <summary>
    ''' Função que busca dados de um endereço baseado no CEP (8 caracteres)
    ''' </summary>
    ''' <param name="CEP">Recebe o CEP em formato String com 8 caracteres</param>
    ''' <returns>Retorna uma estrutura do tipo Utils.CEP contendo Rua, Bairro, Cidade, UF, DDD, etc. Caso insucesso, o flag CEP.Sucesso é igual a false e o CEP de retorno é 00000000</returns>
    ''' <remarks>Usa o serviço da WEB de ViaCEP (viacep.com.br)</remarks>
    ''' <example>Dim x as String = FBR.Utils.ConsultaCEP("13271200")</example>
    Public Shared Function ConsultaCEP(ByVal CEP As String) As CEP
        'viacep.com.br/ ws / 1001000 / Xml /
        Dim CEPProcessado As String = FBR.Texto.LPAD(CEP.Replace(" ", "").Replace("-", ""), 8)
        Dim dados As String = FBR.WEB.HTTPGet("https://viacep.com.br/ws/" & CEPProcessado & "/xml/", WEB.Codificacoes.UTF_8)
        '        Dim dados As String = "<?xml version=""1.0"" encoding=""UTF-8""?>
        '<xmlcep>
        '  <cep>13270-293</cep>
        '  <logradouro>Rua Luiza Rodella Brandini</logradouro>
        '  <complemento></complemento>
        '  <bairro>Vila São José</bairro>
        '  <localidade>Valinhos</localidade>
        '  <uf>SP</uf>
        '  <ibge>3556206</ibge>
        '  <gia>7080</gia>
        '  <ddd>19</ddd>
        '  <siafi>7225</siafi>
        '</xmlcep>"
        If dados.StartsWith(FBR.WEB.TEXTO_ERRO) Then
            Return New CEP()
        Else
            Dim _cep As New CEP
            Dim xel As XElement = XElement.Parse(dados)
            Dim itens As IEnumerable(Of XElement) = xel.Elements()

            For Each a In itens
                'Console.WriteLine("{0} > {1}", a.Name, a.Value)
                Select Case a.Name.ToString.ToLower
                    Case "cep"
                        _cep.CEP = FBR.Texto.LPAD(a.Value.ToString.Replace(" ", "").Replace("-", ""), 8)
                        _cep.Sucesso = True
                    Case "logradouro"
                        _cep.Rua = a.Value
                    Case "bairro"
                        _cep.Bairro = a.Value
                    Case "localidade"
                        _cep.Cidade = a.Value
                    Case "uf"
                        _cep.UF = a.Value
                    Case "ddd"
                        _cep.DDD = a.Value
                    Case "ibge"
                        _cep.IBGE = a.Value
                End Select

            Next a
            Return _cep

        End If

    End Function

    'Function fcnNumeroSerieWindows() As String
    '    Dim moReturn As Management.ManagementObjectCollection
    '    Dim moSearch As Management.ManagementObjectSearcher
    '    Dim mo As Management.ManagementObject
    '    Dim strNumeroSerie As String = ""

    '    moSearch = New Management.ManagementObjectSearcher("Select * from Win32_OperatingSystem")
    '    moReturn = moSearch.Get

    '    For Each mo In moReturn
    '        strNumeroSerie = mo("SerialNumber")
    '        'Dim strOut As String = String.Format("p0={0}  p1={1} p2={2}", deviceID, VolumeName, SerialNumber)
    '        'texto += strOut & vbCrLf
    '    Next
    '    Return strNumeroSerie
    'End Function
    'Function fcnDataBootWindows() As DateTime
    '    'Fonte=http://www.c-sharpcorner.com/uploadfile/scottlysle/determine-the-time-since-the-last-boot-up-in-visual-basic3/

    '    'Dim moReturn As Management.ManagementObjectCollection
    '    'Dim moSearch As Management.ManagementObjectSearcher
    '    Dim mo As Management.ManagementObject

    '    ' define a select query
    '    Dim query As New Management.SelectQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem WHERE Primary='true'")


    '    ' create a new management object searcher and pass it
    '    ' the select query
    '    Dim searcher As New Management.ManagementObjectSearcher(query)

    '    ' get the datetime value and set the local boot
    '    ' time variable to contain that value
    '    'Dim mo As ManagementObject
    '    Dim dtBootTime As DateTime

    '    For Each mo In searcher.Get()
    '        dtBootTime = Management.ManagementDateTimeConverter.ToDateTime(mo.Properties("LastBootUpTime").Value.ToString())
    '    Next

    '    searcher = Nothing
    '    mo = Nothing
    '    Return dtBootTime
    'End Function

End Class
