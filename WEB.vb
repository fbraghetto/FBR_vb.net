Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Xml
Imports Newtonsoft.Json

Public NotInheritable Class WEB
    '- Suporte a TLS e SSL3 (devido bug atualização Windows 7 em 17/02)
    ' Configuração de TLS/SSL 
    Public Shared TipoConfiguracaoSSL As System.Net.SecurityProtocolType = (System.Net.SecurityProtocolType.Tls + System.Net.SecurityProtocolType.Tls11 + System.Net.SecurityProtocolType.Tls12)
    Public Const VerificaListaRevogacaoSSL As Boolean = False
    Public Shared Property ModoOffline As Boolean = False
    Public Const TEXTO_ERRO As String = "ERRO****"

    Public Enum Codificacoes As Integer
        AUTO = 0
        Windows_1252 = 1252
        UTF32 = 12000
        US_ASCII = 20127
        ISO_8859_1 = 28591
        UTF_8 = 65001
    End Enum
    Public Enum DNS_RRType
        ALL = 0
        A
        AAAA
        CNAME
        MX
        NS
        TXT
    End Enum

    <CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification:="<Pending>")>
    Public Shared Function HTTPPost(ByVal URL As String, ByVal strPOST As String, Optional ByVal contentType As String = "application/x-www-form-urlencoded") As String
        Try
            Return HTTPPostComExcecao(URL, strPOST, contentType)
        Catch ex As Exception
            'SaidaConsole.Output("Erro [WEB.HTTPPost]: URL=" & URL & "/// strPOST=" & strPOST & "/// " & ex.ToString, System.Reflection.MethodBase.GetCurrentMethod())
            Return TEXTO_ERRO & ":" & ex.ToString
        End Try
    End Function
    Public Shared Function HTTPPostComExcecao(ByVal URL As String, ByVal strPOST As String, Optional ByVal contentType As String = "application/x-www-form-urlencoded") As String

        Dim request As WebRequest = WebRequest.Create(URL)
        request.Method = "POST"

        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(strPOST)       ' Get the POST Data and convert to ByteArray

        request.ContentType = contentType                               ' Set the ContentType property of the WebRequest.
        request.ContentLength = byteArray.Length                        ' Set the ContentLength property of the WebRequest.
        Dim dataStream As Stream = request.GetRequestStream()           ' Get the request stream.
        dataStream.Write(byteArray, 0, byteArray.Length)                ' Write the data to the request stream.
        dataStream.Close()                                              ' Close the Stream object.

        Dim response As WebResponse = request.GetResponse()             ' Get the response.
        Dim sStatus As String = (CType(response, HttpWebResponse).StatusDescription) ' Display the status.
        dataStream = response.GetResponseStream()                       ' Get the stream containing content returned by the server.
        Dim reader As New StreamReader(dataStream)                      ' Open the stream using a StreamReader for easy access.
        Dim responseFromServer As String = reader.ReadToEnd()           ' Read the content.
        reader.Close()                                                  ' Clean up the streams.
        dataStream.Close()
        response.Close()

        Return responseFromServer
    End Function

    <CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification:="<Pending>")>
    Public Shared Function HTTPGet(ByVal URL As String, Optional ByVal TipoCodificacao As Codificacoes = Codificacoes.AUTO) As String
        Try
            Return HTTPGetComExcecao(URL, TipoCodificacao)
        Catch objErr As Exception
            'MsgBox(objErr.ToString)
            'SaidaConsole.Output("GetHTML_comCharset()" & objErr.ToString, System.Reflection.MethodBase.GetCurrentMethod())
            'SaidaConsole.Output("Erro no download. URL=" & strURL & " // Erro: " & objErr.ToString, System.Reflection.MethodBase.GetCurrentMethod())
            'fcnLOG("Erro no download. URL=" & strURL & " // Erro: " & objErr.ToString, cls_LOG.tipoLOG.LOG_ERR)
            Return TEXTO_ERRO & ":" & objErr.ToString
        End Try
    End Function

    Public Shared Function HTTPGetComExcecao(ByVal URL As String, Optional ByVal TipoCodificacao As Codificacoes = Codificacoes.AUTO) As String
        If ModoOffline = 1 Then Throw New WebException("Modo Offline Ativo. Bloqueando acesso a Internet.")

        Dim myWebClient As New System.Net.WebClient
        Dim myStream As System.IO.Stream
        Dim myStreamReader As System.IO.StreamReader

        System.Net.ServicePointManager.CheckCertificateRevocationList = VerificaListaRevogacaoSSL
        System.Net.ServicePointManager.SecurityProtocol = TipoConfiguracaoSSL

        If TipoCodificacao = Nothing Then TipoCodificacao = Codificacoes.AUTO

        myStream = myWebClient.OpenRead(URL)
        If TipoCodificacao = Codificacoes.AUTO Then
            ' Se for automático o charset, então usa a definição padrão
            myStreamReader = New System.IO.StreamReader(myStream)
        Else
            ' se definir um charset pessoal, então forçar o novocharset
            myStreamReader = New System.IO.StreamReader(myStream, System.Text.Encoding.GetEncoding(CInt(TipoCodificacao)))
        End If
        Dim strResultado As String = myStreamReader.ReadToEnd()
        myStreamReader.Close()
        myStream.Close()
        Return strResultado
    End Function


    Public Shared Function GetXMLDoc(ByVal URL As String, ByVal node As String) As Xml.XmlNodeList
        'Função recebe uma URL e o nome de um nó e devolve uma estrutura
        ' XML formatada para a origem.
        ' Por exemplo GetXMLDoc("http://www.xyz.com", "rss/channel/item")
        ' URL = "http://www.tempoagora.com.br/rss/cidades/rss_prev_Valinhos-SP.xml"

        Dim tempNodeList As System.Xml.XmlNodeList = Nothing

        If ModoOffline Then
            'SaidaConsole.Output("Modo Offline Ativo. Bloqueando acesso a Internet.", System.Reflection.MethodBase.GetCurrentMethod())
            Return tempNodeList
            Exit Function
        End If
        Try
            Dim request As WebRequest = WebRequest.Create(URL)
            Dim response As WebResponse = request.GetResponse()
            Dim rssStream As Stream = response.GetResponseStream()
            Dim rssDoc As XmlDocument = New XmlDocument

            rssDoc.Load(rssStream)

            tempNodeList = rssDoc.SelectNodes(node)

        Catch ex As Exception
            ' fcnLOG("Não foi possivel obter XML " & URL & " --> " & ex.ToString, cls_LOG.tipoLOG.LOG_WARN)
        End Try

        Return tempNodeList


    End Function


    Public Shared Function URLEncode(ByVal HTML As String) As String
        'Liberados =>>>
        'a-z = 97 / 122
        'A-Z = 65 / 90
        '0-9 = 48 / 57
        '-_.~ 

        Dim sbSaida As New StringBuilder(5000)

        If HTML.Length <= 0 Then
            Return HTML
        Else
            For i As Long = 0 To (HTML.Length - 1)
                Dim s As String = HTML.Substring(i, 1)
                If s = "-" Or s = "_" Or s = "." Or s = "~" Then
                    sbSaida.Append(s)
                Else
                    If Numero.Entre(Asc(s), 48, 57) Or Numero.Entre(Asc(s), 65, 90) Or Numero.Entre(Asc(s), 97, 122) Then
                        sbSaida.Append(s)
                    Else
                        sbSaida.Append("%")
                        sbSaida.Append(Strings.Right("00" & Hex(Asc(s)), 2))
                    End If
                End If
            Next
            Return sbSaida.ToString

        End If

    End Function

    Public Shared Function RemoveTAGHTML(ByVal HTML As String) As String
        Dim tagOpen As Integer = 0
        Dim len As Integer = HTML.Length
        Dim c As String
        Dim resultado As New StringBuilder(len)

        For i As Integer = 1 To len
            c = Strings.Mid(HTML, i, 1)
            If c = "<" Then
                tagOpen += 1
            ElseIf c = ">" Then
                tagOpen -= 1
            ElseIf tagOpen <= 0 Then
                resultado.Append(c)
            End If
        Next
        Return resultado.ToString
    End Function
    Public Shared Function HtmlEncode(ByVal Entrada As String) As String
        If Entrada.Length <= 3 Then
            Return Entrada
        Else
            Dim sbSaida As New StringBuilder(Entrada.Length + 1)
            sbSaida.Append(Entrada)

            Dim origem() As String = {Chr(34), Chr(38), Chr(39), Chr(60), Chr(62), Chr(160), Chr(161), Chr(162), Chr(163), Chr(164), Chr(165), Chr(166), Chr(167), Chr(168), Chr(169), Chr(170), Chr(171), Chr(172), Chr(173), Chr(174), Chr(175), Chr(176), Chr(177), Chr(178), Chr(179), Chr(180), Chr(181), Chr(182), Chr(183), Chr(184), Chr(185), Chr(186), Chr(187), Chr(188), Chr(189), Chr(190), Chr(191), Chr(192), Chr(193), Chr(194), Chr(195), Chr(196), Chr(197), Chr(198), Chr(199), Chr(200), Chr(201), Chr(202), Chr(203), Chr(204), Chr(205), Chr(206), Chr(207), Chr(208), Chr(209), Chr(210), Chr(211), Chr(212), Chr(213), Chr(214), Chr(215), Chr(216), Chr(217), Chr(218), Chr(219), Chr(220), Chr(221), Chr(222), Chr(223), Chr(224), Chr(225), Chr(227), Chr(228), Chr(229), Chr(230), Chr(231), Chr(232), Chr(233), Chr(234), Chr(235), Chr(236), Chr(237), Chr(238), Chr(239), Chr(240), Chr(241), Chr(242), Chr(243), Chr(244), Chr(245), Chr(246), Chr(247), Chr(248), Chr(249), Chr(250), Chr(251), Chr(252), Chr(253), Chr(254), Chr(255)}
            Dim destino() As String = {"&quot;", "&amp;", "&apos;", "&lt;", "&gt;", "&nbsp;", "&iexcl;", "&cent;", "&pound;", "&curren;", "&yen;", "&brvbar;", "&sect;", "&uml;", "&copy;", "&ordf;", "&laquo;", "&not;", "&shy;", "&reg;", "&macr;", "&deg;", "&plusmn;", "&sup2;", "&sup3;", "&acute;", "&micro;", "&para;", "&middot;", "&cedil;", "&sup1;", "&ordm;", "&raquo;", "&frac14;", "&frac12;", "&frac34;", "&iquest;", "&Agrave;", "&Aacute;", "&Acirc;", "&Atilde;", "&Auml;", "&Aring;", "&AElig;", "&Ccedil;", "&Egrave;", "&Eacute;", "&Ecirc;", "&Euml;", "&Igrave;", "&Iacute;", "&Icirc;", "&Iuml;", "&ETH;", "&Ntilde;", "&Ograve;", "&Oacute;", "&Ocirc;", "&Otilde;", "&Ouml;", "&times;", "&Oslash;", "&Ugrave;", "&Uacute;", "&Ucirc;", "&Uuml;", "&Yacute;", "&THORN;", "&szlig;", "&agrave;", "&aacute;", "&atilde;", "&auml;", "&aring;", "&aelig;", "&ccedil;", "&egrave;", "&eacute;", "&ecirc;", "&euml;", "&igrave;", "&iacute;", "&icirc;", "&iuml;", "&eth;", "&ntilde;", "&ograve;", "&oacute;", "&ocirc;", "&otilde;", "&ouml;", "&divide;", "&oslash;", "&ugrave;", "&uacute;", "&ucirc;", "&uuml;", "&yacute;", "&thorn;", "&yuml;"}
            For i As Integer = 0 To origem.Length - 1
                sbSaida = sbSaida.Replace(origem(i).ToString(), destino(i).ToString())
            Next
            Return sbSaida.ToString

        End If

    End Function


    Public Shared Function HTMLDecode(ByVal StringEntradaHTML As String) As String
        If Not StringEntradaHTML.Contains("&") Then
            Return StringEntradaHTML
        Else


            Dim origem() As String = {Chr(34), Chr(38), Chr(39), Chr(60), Chr(62), Chr(160), Chr(161), Chr(162), Chr(163), Chr(164), Chr(165), Chr(166), Chr(167), Chr(168), Chr(169), Chr(170), Chr(171), Chr(172), Chr(173), Chr(174), Chr(175), Chr(176), Chr(177), Chr(178), Chr(179), Chr(180), Chr(181), Chr(182), Chr(183), Chr(184), Chr(185), Chr(186), Chr(187), Chr(188), Chr(189), Chr(190), Chr(191), Chr(192), Chr(193), Chr(194), Chr(195), Chr(196), Chr(197), Chr(198), Chr(199), Chr(200), Chr(201), Chr(202), Chr(203), Chr(204), Chr(205), Chr(206), Chr(207), Chr(208), Chr(209), Chr(210), Chr(211), Chr(212), Chr(213), Chr(214), Chr(215), Chr(216), Chr(217), Chr(218), Chr(219), Chr(220), Chr(221), Chr(222), Chr(223), Chr(224), Chr(225), Chr(227), Chr(228), Chr(229), Chr(230), Chr(231), Chr(232), Chr(233), Chr(234), Chr(235), Chr(236), Chr(237), Chr(238), Chr(239), Chr(240), Chr(241), Chr(242), Chr(243), Chr(244), Chr(245), Chr(246), Chr(247), Chr(248), Chr(249), Chr(250), Chr(251), Chr(252), Chr(253), Chr(254), Chr(255)}
            Dim destino() As String = {"&quot;", "&amp;", "&apos;", "&lt;", "&gt;", "&nbsp;", "&iexcl;", "&cent;", "&pound;", "&curren;", "&yen;", "&brvbar;", "&sect;", "&uml;", "&copy;", "&ordf;", "&laquo;", "&not;", "&shy;", "&reg;", "&macr;", "&deg;", "&plusmn;", "&sup2;", "&sup3;", "&acute;", "&micro;", "&para;", "&middot;", "&cedil;", "&sup1;", "&ordm;", "&raquo;", "&frac14;", "&frac12;", "&frac34;", "&iquest;", "&Agrave;", "&Aacute;", "&Acirc;", "&Atilde;", "&Auml;", "&Aring;", "&AElig;", "&Ccedil;", "&Egrave;", "&Eacute;", "&Ecirc;", "&Euml;", "&Igrave;", "&Iacute;", "&Icirc;", "&Iuml;", "&ETH;", "&Ntilde;", "&Ograve;", "&Oacute;", "&Ocirc;", "&Otilde;", "&Ouml;", "&times;", "&Oslash;", "&Ugrave;", "&Uacute;", "&Ucirc;", "&Uuml;", "&Yacute;", "&THORN;", "&szlig;", "&agrave;", "&aacute;", "&atilde;", "&auml;", "&aring;", "&aelig;", "&ccedil;", "&egrave;", "&eacute;", "&ecirc;", "&euml;", "&igrave;", "&iacute;", "&icirc;", "&iuml;", "&eth;", "&ntilde;", "&ograve;", "&oacute;", "&ocirc;", "&otilde;", "&ouml;", "&divide;", "&oslash;", "&ugrave;", "&uacute;", "&ucirc;", "&uuml;", "&yacute;", "&thorn;", "&yuml;"}


            Dim sbSaida As New StringBuilder(StringEntradaHTML.Length)
            sbSaida.Append(StringEntradaHTML)

            ' Fase 1 - Decode de Caracteres mais conhecidos
            For i As Integer = 0 To origem.Length - 1
                sbSaida = sbSaida.Replace(destino(i).ToString(), origem(i).ToString())
            Next

            ' Se sobrar algo para decodificar
            If sbSaida.ToString.Contains("&") Then
                Dim c As String
                Dim charEspecialAberto As Boolean = False
                Dim strSaida As String = ""
                Dim caracterEspecial As String = ""
                Dim strPreProcessado As String = sbSaida.ToString()

                For i As Integer = 1 To StringEntradaHTML.Length
                    c = Mid(strPreProcessado, i, 1)
                    If c = "&" Then
                        charEspecialAberto = True
                        caracterEspecial = ""
                    ElseIf charEspecialAberto = True Then
                        If c = ";" Then
                            If IsNumeric(caracterEspecial) Then
                                strSaida &= Chr(CInt(caracterEspecial))
                            End If
                            caracterEspecial = ""
                            charEspecialAberto = False
                        ElseIf c = "#" Then
                            ' ignora
                        Else
                            caracterEspecial &= c
                        End If
                    Else
                        strSaida &= c
                    End If
                Next
                Return strSaida
            Else
                Return sbSaida.ToString()
            End If
        End If
    End Function

    Public Shared Function ChecarFormatoEnderecoIP(ByVal IPAddr As String) As Boolean
        Dim ip As New Net.IPAddress(0)
        Return Net.IPAddress.TryParse(IPAddr, ip)
    End Function
    Public Shared Function ResolveDNS(ByVal FQDN_OR_IP As String) As String
        ''https://developers.google.com/speed/public-dns/docs/doh/json
        ''https://opensource.adobe.com/Spry/samples/data_region/JSONDataSetSample.html
        'Dim err As String = "www.uol.com.br"


        '' https://dns.google/resolve?name=uol.com.br&type=A
        'Dim s1 As String = <ss>{"Status": 0,"TC": false,"RD": true,"RA": true,"AD": false,"CD": false,"Question":[ {"name": "uol.com.br.","type": 1}],"Answer":[ {"name": "uol.com.br.","type": 1,"TTL": 15,"data": "200.147.35.149"}]}</ss>

        ''https://dns.google/resolve?name=servidor.vv8.tv.br&type=A
        'Dim s4 As String = <ss>{"Status": 0,"TC": false,"RD": true,"RA": true,"AD": false,"CD": false,"Question":[ {"name": "servidor.vv8.tv.br.","type": 1}],"Answer":[ {"name": "servidor.vv8.tv.br.","type": 5,"TTL": 59,"data": "clinicavalinhos.ddns.net."},{"name": "clinicavalinhos.ddns.net.","type": 1,"TTL": 59,"data": "179.104.43.226"}],"Comment": "Response from 200.220.152.97."}</ss>

        ''QuadA IPV6
        'Dim s5 As String = <ss>{"Status": 0,"TC": false,"RD": true,"RA": true,"AD": false,"CD": false,"Question":[ {"name": "www.google.com.br.","type": 28}],"Answer":[ {"name": "www.google.com.br.","type": 28,"TTL": 299,"data": "2800:3f0:4001:81a::2003"}],"Comment": "Response from 2001:4860:4802:34::a."}</ss>

        '' https://dns.google/resolve?name=uol.com.br&type=SOA
        'Dim s2 As String = <ss>{"Status": 0,"TC": false,"RD": true,"RA": true,"AD": false,"CD": false,"Question":[ {"name": "uol.com.br.","type": 6}],"Answer":[ {"name": "uol.com.br.","type": 6,"TTL": 3413,"data": "eliot.uol.com.br. root.uol.com.br. 2016046693 7200 3600 432000 900"}]}</ss>

        ''https://dns.google/resolve?name=uol.com.br&type=ALL
        'Dim s3 As String = <ss>{"Status": 0,"TC": false,"RD": true,"RA": true,"AD": false,"CD": false,"Question":[ {"name": "uol.com.br.","type": 255}],"Answer":[ {"name": "uol.com.br.","type": 6,"TTL": 3599,"data": "eliot.uol.com.br. root.uol.com.br. 2016046693 7200 3600 432000 900"},{"name": "uol.com.br.","type": 1,"TTL": 59,"data": "200.147.3.157"},{"name": "uol.com.br.","type": 1,"TTL": 59,"data": "200.147.35.149"},{"name": "uol.com.br.","type": 28,"TTL": 59,"data": "2804:49c:3101:401:ffff:ffff:ffff:45"},{"name": "uol.com.br.","type": 28,"TTL": 59,"data": "2804:49c:3102:401:ffff:ffff:ffff:36"},{"name": "uol.com.br.","type": 15,"TTL": 17999,"data": "10 mx.uol.com.br."},{"name": "uol.com.br.","type": 16,"TTL": 59,"data": "\"DZC=qvK68EJ\""},{"name": "uol.com.br.","type": 16,"TTL": 59,"data": "\"v=spf1 include:_net1.uol.com.br include:_net2.uol.com.br -all\""},{"name": "uol.com.br.","type": 16,"TTL": 59,"data": "\"google-site-verification=wIL3zoUpiok1eJim7TNuqy-1KE6hK1pFFopPoDLR0PY\""},{"name": "uol.com.br.","type": 16,"TTL": 59,"data": "\"google-site-verification=ZT5tj3Vi3KQFGXNuyj-D1UgbMDBBqy47-GR0ZSkXHrY\""},{"name": "uol.com.br.","type": 16,"TTL": 59,"data": "\"google-site-verification=Y3guNK11wdcxMiSpCie8oOjVwNdlpou9UtBp3b1aFg8\""},{"name": "uol.com.br.","type": 16,"TTL": 59,"data": "\"google-site-verification=PNCvwjDZiXU9zUGD8ik56KfauBWibOR2fGjMOLvfgZg\""},{"name": "uol.com.br.","type": 2,"TTL": 3599,"data": "borges.uol.com.br."},{"name": "uol.com.br.","type": 2,"TTL": 3599,"data": "eliot.uol.com.br."},{"name": "uol.com.br.","type": 2,"TTL": 3599,"data": "charles.uol.com.br."}],"Comment": "Response from 200.147.38.8."}</ss>

        ''https://dns.google/resolve?name=22.152.220.200.in-addr.arpa&type=PTR
        'Dim ptr As String = <a>{"Status":  0,"TC": false,"RD": true,"RA": true,"AD": false,"CD": false,"Question":[ {"name": "22.152.220.200.In-addr.arpa.","type": 12}],"Answer":[ {"name": "22.152.220.200.In-addr.arpa.","type": 12,"TTL": 899,"data": "dns1.doctordata.com.br."}],"Comment": "Response from 200.220.152.77."}</a>

        ''Desmontar
        'Dim s As String = s1
        ''Dim nivel As Integer = 0
        ''For i As Integer = 0 To s.Length
        ''    If s.Substring(i) = "{" Then nivel += 1
        ''    If s.Substring(i) = "}" Then nivel += 1
        ''Next

        Dim s As String = ResolveDNSparaJSON(FQDN_OR_IP)
        Dim resposta As String = String.Empty
        If s.Length > 10 Then
            Try
                Dim ddd = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DNS_A_Response)(s)
                For Each ans As Answer In ddd.Answer
                    If resposta.Length <= 0 Then resposta = ans.Data.ToString
                    If ans.Type = 1 Then
                        resposta = ans.Data.ToString
                        Exit For
                    End If
                Next
#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception
                resposta = ex.Message
#Enable Warning CA1031 ' Do not catch general exception types
            End Try
        End If
        Return resposta
    End Function
    Public Shared Function ResolveDNSparaJSON(ByVal FQDN_OR_IP As String, Optional ByVal RRType As DNS_RRType = DNS_RRType.A) As String
        Dim URL As String = String.Empty
        Select Case Uri.CheckHostName(FQDN_OR_IP)
            Case UriHostNameType.Dns
                'A domain name can be up to 63 characters (letters, numbers or combination) long plus the 4 characters used to identify the domain extension (.com, .net, .org). The only symbol character domain names can include is a hyphen (-) although the domain name cannot start or end with a hyphen nor have consecutive hyphens. Common symbols such as asterisks (*), under scores (_), and exclamation points (!) are not allowed.
                If FQDN_OR_IP.Contains(".") And Numero.Entre(FQDN_OR_IP.Length, 3, 63) And
                    Not FQDN_OR_IP.EndsWith(".") And Not FQDN_OR_IP.StartsWith(".") And
                    Not FQDN_OR_IP.Contains("*") And Not FQDN_OR_IP.Contains("_") And
                    Not FQDN_OR_IP.Contains("!") And Not FQDN_OR_IP.Contains("--") Then
                    URL = String.Format("https://dns.google/resolve?name={0}&type={1}", FQDN_OR_IP.Replace("&", ""), "A")
                End If
            Case UriHostNameType.IPv4
                Dim octetos() As String = Split(FQDN_OR_IP, ".", 4)
                URL = String.Format("https://dns.google/resolve?name={3}.{2}.{1}.{0}.in-addr.arpa&type=PTR", octetos(0), octetos(1), octetos(2), octetos(3))
            Case UriHostNameType.IPv6

            Case Else

        End Select
        If URL.Length > 0 Then
            Return FBR.WEB.HTTPGet(URL)
        Else
            Return String.Empty
        End If
    End Function

    Private Class Question

        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("type")>
        Public Property Type As Integer
    End Class

    Private Class Answer

        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("type")>
        Public Property Type As Integer

        <JsonProperty("TTL")>
        Public Property TTL As Integer

        <JsonProperty("data")>
        Public Property Data As String
    End Class

    Private Class DNS_A_Response

        <JsonProperty("Status")>
        Public Property Status As Integer

        <JsonProperty("TC")>
        Public Property TC As Boolean

        <JsonProperty("RD")>
        Public Property RD As Boolean

        <JsonProperty("RA")>
        Public Property RA As Boolean

        <JsonProperty("AD")>
        Public Property AD As Boolean

        <JsonProperty("CD")>
        Public Property CD As Boolean

        <JsonProperty("Question")>
        Public Property Question As Question()

        <JsonProperty("Answer")>
        Public Property Answer As Answer()
    End Class

    Public Function LerConteudoHTTP(ByVal URL As String) As String

        ' Melhorar Implementação

        Throw New NotImplementedException()

        Try
            ' Teste de TLS
            System.Net.ServicePointManager.CheckCertificateRevocationList = False
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls

            Dim client As System.Net.WebClient = New Net.WebClient()

            client.Headers.Clear()

            client.Headers.Add("accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3")
            client.Headers.Add("accept-language", "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7")
            client.Headers.Add("sec-fetch-mode", "navigate")
            client.Headers.Add("sec-fetch-site", "none")
            client.Headers.Add("cache-control", "private")
            client.Headers.Add("upgrade-insecure-requests", "1")
            ' client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36")
            Return client.DownloadString(URL)

        Catch ex As Exception
            Return "ERRO " & ex.Message
        End Try


    End Function
    Public Shared Function PegarEnderecoIPv4()
        Dim strSaida As String = "0.0.0.0"
        Dim ip As IPAddress() = Dns.GetHostEntry(Dns.GetHostName().ToString).AddressList()
        For Each i In ip
            If i.AddressFamily = Sockets.AddressFamily.InterNetwork And Not i.ToString.StartsWith("169") Then
                strSaida = i.ToString
                Exit For
            End If
        Next
        Return strSaida
    End Function


    Public Function EnviarEmail(ByVal Mensagem As String, ByVal Assunto As String, ByVal DestinoEmail1 As String, Optional ByVal DestinoEmail2 As String = "", Optional ByVal DestinoEmail3 As String = "",
                                Optional ByVal flagEmailHTML As Boolean = False, Optional ByVal EMAIL_FROM As String = "",
                                Optional ByVal EMAIL_SMTP_SERVER As String = "", Optional ByVal EMAIL_SMTP_PORT As String = "", Optional ByVal EMAIL_SMTP_SSL_ENABLED As String = "",
                                Optional ByVal EMAIL_SMTP_LOGIN As String = "", Optional ByVal EMAIL_SMTP_PASSWORD As String = "") As Boolean
        Dim resultado As Boolean = True

        If Mensagem.Length < 2 Or Assunto.Length < 2 Or DestinoEmail1.Length < 3 Or Not DestinoEmail1.Contains("@") Then
            resultado = False
        Else
            Try
                Dim smtp As New System.Net.Mail.SmtpClient(EMAIL_SMTP_SERVER, EMAIL_SMTP_PORT)
                Dim email As New System.Net.Mail.MailMessage
                With smtp
                    .EnableSsl = EMAIL_SMTP_SSL_ENABLED
                    .UseDefaultCredentials = False
                    .Credentials = New Net.NetworkCredential(EMAIL_SMTP_LOGIN, EMAIL_SMTP_PASSWORD)
                End With

                With email
                    .From = New Net.Mail.MailAddress(EMAIL_FROM)
                    .To.Add(DestinoEmail1)
                    .Subject = Assunto
                    .IsBodyHtml = flagEmailHTML
                    .Body = Mensagem
                End With

                If DestinoEmail2.Length > 0 And DestinoEmail2.Contains("@") Then email.To.Add(DestinoEmail2)
                If DestinoEmail3.Length > 0 And DestinoEmail3.Contains("@") Then email.To.Add(DestinoEmail3)

                smtp.Send(email)


            Catch ex As Exception
                Console.WriteLine("EnviarEmail() falhou: " & ex.Message.ToString)
                resultado = False
            End Try

        End If
        Return resultado

    End Function


End Class
