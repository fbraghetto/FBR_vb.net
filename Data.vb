Imports System.Text

Public NotInheritable Class Data

    Public Const MinDate As Date = #1/1/1900#




    Public Shared Function NomeDiaDaSemana(ByVal intWeekday As Integer, Optional ByVal lingua As FBR.Texto.Linguagem = Texto.Linguagem.PORTUGUES) As String
        Dim diaDaSemanaPortugues() As String = {"Domingo", "Segunda-Feira", "Terça-Feira", "Quarta-Feira", "Quinta-Feira", "Sexta-Feira", "Sábado", "Desconhecido"}
        Dim diaDaSemanaIngles() As String = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thrusday", "Friday", "Saturday", "Unknow"}
        Dim diaDaSemanaEspanhol() As String = {"Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Desconocido"}
        Const MAX As Integer = 7

        Select Case lingua
            Case Texto.Linguagem.PORTUGUES
                If intWeekday >= 1 And intWeekday <= MAX Then
                    Return diaDaSemanaPortugues(intWeekday - 1)
                Else
                    Return diaDaSemanaPortugues(MAX)
                End If
            Case Texto.Linguagem.INGLES
                If intWeekday >= 1 And intWeekday <= MAX Then
                    Return diaDaSemanaIngles(intWeekday - 1)
                Else
                    Return diaDaSemanaIngles(MAX)
                End If
            Case Texto.Linguagem.ESPANHOL
                If intWeekday >= 1 And intWeekday <= MAX Then
                    Return diaDaSemanaEspanhol(intWeekday - 1)
                Else
                    Return diaDaSemanaEspanhol(MAX)
                End If
            Case Else
                Return diaDaSemanaPortugues(MAX)
        End Select
    End Function
    Public Shared Function NomeMes(ByVal intMes As Integer, Optional ByVal lingua As FBR.Texto.Linguagem = Texto.Linguagem.PORTUGUES) As String
        Dim mesPortugues() As String = {"Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro", "Desconhecido"}
        Dim mesIngles() As String = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Desconocido"}
        Dim mesEspanhol() As String = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Desconocido"}
        Const MAX As Integer = 12

        Select Case lingua
            Case Texto.Linguagem.PORTUGUES
                If intMes >= 1 And intMes <= MAX Then
                    Return mesPortugues(intMes - 1)
                Else
                    Return mesPortugues(MAX)
                End If
            Case Texto.Linguagem.INGLES
                If intMes >= 1 And intMes <= MAX Then
                    Return mesIngles(intMes - 1)
                Else
                    Return mesIngles(MAX)
                End If
            Case Texto.Linguagem.ESPANHOL
                If intMes >= 1 And intMes <= MAX Then
                    Return mesEspanhol(intMes - 1)
                Else
                    Return mesEspanhol(MAX)
                End If
            Case Else
                Return mesPortugues(MAX)
        End Select
    End Function


    ''' <summary>
    ''' Retorna o Numero do Mes a partir de uma string
    ''' </summary>
    ''' <param name="nomeMes">Nome do Mes</param>
    ''' <param name="lingua">Linguagem especificada, caso não especificado será usado portugues do Brasil</param>
    ''' <param name="valueOnException">Qual o retorno esperado em caso de erro</param>
    ''' <returns>Retorna o Numero do Mes (ex. Jan=1, Fev=2 ou Janeiro=1, Fevereiro=2)</returns>
    Public Shared Function NumeroMesPeloNome(ByVal nomeMes As String, Optional lingua As FBR.Texto.Linguagem = Texto.Linguagem.PORTUGUES, Optional ByVal valueOnException As Integer = 0) As Integer
        Dim ci As New Globalization.CultureInfo(Texto.LCIDparaLinguagem(lingua))
        Dim format As String = "MMMM"
        If nomeMes.Length = 3 Then format = "MMM"
        Try
            Return DateTime.ParseExact(nomeMes.ToUpper, format, ci).Month
        Catch ex As ArgumentNullException
            Return valueOnException
        Catch ex As FormatException
            Return valueOnException
        End Try

    End Function


    Public Shared Function Entre(ByVal data As DateTime, ByVal DataInicio As DateTime, ByVal DataFinal As DateTime) As Boolean
        If data >= DataInicio And data <= DataFinal Then Return True Else Return False
    End Function
    Public Shared Function Menor(ByVal d1 As DateTime, ByVal d2 As DateTime) As DateTime
        If d1 <= d2 Then Return d1 Else Return d2
    End Function
    Public Shared Function Maior(ByVal d1 As DateTime, ByVal d2 As DateTime) As DateTime
        If d1 >= d2 Then Return d1 Else Return d2
    End Function
    Public Shared Function DiferencaEmSegundos(ByVal d1 As DateTime, ByVal d2 As DateTime) As Integer
        Dim x As TimeSpan = d1.Subtract(d2)
        Return Math.Abs(Math.Abs(x.TotalSeconds))
    End Function
    Public Shared Function QualMaior(ByVal d1 As DateTime, ByVal d2 As DateTime) As Integer
        ' retorna 1 se d1 é maior, 2 se d2 é maior, 0 se for igual
        If d1 > d2 Then
            Return 1
        ElseIf d2 > d1 Then
            Return 2
        Else
            Return 0
        End If
    End Function
    Public Shared Function Truncate(ByVal d1 As DateTime) As Date
        Return New Date(d1.Year, d1.Month, d1.Day)
    End Function
    Public Shared Function TruncarData(ByVal d1 As DateTime) As Date
        Return Truncate(d1)
    End Function


    Public Shared Function ConverteDataXMLParaDateTime(ByVal strDataHora As String, Optional ByVal ValueOnException As Date = MinDate) As DateTime
        Dim dtSaida As Date = ValueOnException

        Try
            dtSaida = System.Xml.XmlConvert.ToDateTime(strDataHora, System.Xml.XmlDateTimeSerializationMode.Local)
        Catch ex As FormatException
        Catch ex As ArgumentException
        End Try

        Return dtSaida
    End Function
    Public Shared Function ConverteDataRFCparaDateTime(ByVal DataEntrada As String, Optional ByVal ValueOnException As Date = MinDate) As DateTime
        'http://www.macoratti.net/vbn_data.htm

        Dim formato As String = "ddd, dd MMM yyyy HH:mm:ss zzz"
        Dim saida As DateTime
        If Date.TryParseExact(DataEntrada, formato, New Globalization.CultureInfo("en-US", False), Globalization.DateTimeStyles.None, saida) Then
            Return saida
        Else
            Return ValueOnException
        End If



        'Entrada no formato RFC2822 --> Tue, 04 Oct 2011 09:50:18 -0300
        'Updt = "Sun, 05 Aug 2012 12:04:49 -0300"
        'Saída no formato VB.Net 04/10/2011 09:50:18
        'Dim DataSaida As DateTime = ValueOnException
        'Try
        '    DataSaida = Date.Parse(dataEntrada, New Globalization.CultureInfo("en-US", False))
        'Catch e As ArgumentException
        'Catch e As FormatException
        'End Try

        'Return DataSaida


        'Dim arrDatas(6), tmpSaida As String
        'Dim mes As Integer
        'arrDatas = Split(dataEntrada)

        'tmpSaida = arrDatas(1) & "/"
        'Select Case arrDatas(2)
        '    Case "Jan"
        '        mes = 1
        '    Case "Feb"
        '        mes = 2
        '    Case "Mar"
        '        mes = 3
        '    Case "Apr"
        '        mes = 4
        '    Case "May"
        '        mes = 5
        '    Case "Jun"
        '        mes = 6
        '    Case "Jul"
        '        mes = 7
        '    Case "Aug"
        '        mes = 8
        '    Case "Sep"
        '        mes = 9
        '    Case "Oct"
        '        mes = 10
        '    Case "Nov"
        '        mes = 11
        '    Case "Dec"
        '        mes = 12
        '    Case Else
        '        mes = 0
        'End Select
        'tmpSaida += mes & "/" & arrDatas(3)

        'tmpSaida += " " & arrDatas(4)

        'Dim offset As Integer = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).Hours
        'Try
        '    Return CDate(tmpSaida)
        'Catch ex As Exception
        '    Return ValueOnException
        'End Try

    End Function


    ''' <summary>
    ''' Formata uma data de atualização em formato string
    ''' </summary>
    ''' <param name="dDataAtualiz">Data da Ultima Atualizaçao</param>
    ''' <returns>Uma string com a data de atualização em formato humano (exemplo Hoje as 22:00 ou Ontem as 14:00)</returns>
    Public Shared Function FormatarDataAtualizacao(ByVal dDataAtualiz As DateTime) As String
        Dim dataHoje As DateTime = DateTime.Now
        Dim diferencaDias As Integer = (DatePart(DateInterval.DayOfYear, dDataAtualiz) - DatePart(DateInterval.DayOfYear, dataHoje))
        Dim tmpSaida As String
        If diferencaDias = 0 Then
            tmpSaida = "Hoje"
        ElseIf diferencaDias = -1 Then
            tmpSaida = "Ontem"
        Else
            tmpSaida = Format(dDataAtualiz, "dd/MM")
        End If

        tmpSaida &= " às " & Format(dDataAtualiz, "Short Time")
        Return tmpSaida
    End Function


    ''' <summary>
    ''' Retorna a quantidade de Minutos desde a Meia Noite
    ''' </summary>
    ''' <param name="Horario">Horario a converter em formato datetime</param>
    ''' <returns>Um numero inteiro, em minutos. Retorna número negativo em caso de erro</returns>
    Public Shared Function ObterMinutosDesdeMeiaNoite(ByVal Horario As DateTime) As Integer
        Try
            Return (Horario.Hour * 60) + Horario.Minute
        Catch ex As ArgumentException
            Return -1
        End Try
    End Function
    Public Shared Function ConverteMinutosDesdeMeiaNoiteEmString(ByVal Minutos As Integer) As String
        If Minutos < 0 Then
            Return String.Empty
        Else
            Dim m As String = Minutos Mod 60
            Dim h As String = Math.Truncate(Minutos / 60)
            Return String.Concat(Texto.LPAD(h, 2), ":", Texto.LPAD(m, 2))
        End If
    End Function
    Public Shared Function ObterSegundosDesdeMeiaNoite(ByVal Horario As DateTime) As Integer
        Try
            Return (Horario.Hour * 3600) + (Horario.Minute * 60) + Horario.Second
        Catch ex As ArgumentException
            Return -1
        End Try
    End Function

    Public Shared Function ConverteSegundosDesdeMeiaNoiteEmString(ByVal dblSegundos As Double) As String
        ' Recebe um tempo em segundos e retorna um resultado no formato hh:mm:ss ou mm:ss
        ' Pode receber um valor negativo
        Dim m, s As Integer
        Dim intSeg As Integer = Math.Abs(Math.Round(dblSegundos))
        Dim bNegativo As String = ""

        If dblSegundos < 0 Then bNegativo = "-"

        s = Math.Round(intSeg Mod 60)
        m = Math.Round((intSeg - s) / 60)
        Return String.Concat(bNegativo, Numero.LPAD(m, 2), ";", Numero.LPAD(s, 2))
    End Function

    Public Shared Function FormatarSegundosEmString(ByVal intQtdeSegundos As Double) As String
        If intQtdeSegundos < 1 Then
            Return "00:00"
        ElseIf intQtdeSegundos < 2 Then
            Return "00:01"
        Else
            Dim minuto As Integer = (intQtdeSegundos - (intQtdeSegundos Mod 60)) / 60
            Dim segundo As Double = Math.Round(intQtdeSegundos Mod 60)
            Return FBR.Numero.LPAD(minuto, 2) & ":" & FBR.Numero.LPAD(segundo, 2)
        End If
    End Function
    Public Shared Function FormatarSegundosEmHorarioPorExtenso(ByVal intEntrada As Double) As String
        ' Recebe a entrada em segundos
        ' retorna por extenso: "1 hora, 12 minutos e 15 segundos"

        Dim intTempo As Integer = intEntrada
        Dim tmpParcial As Integer
        Dim strSaida As New StringBuilder(100)
        If IsNumeric(intTempo) Then
            If (intTempo >= 3600) Then
                strSaida.Append(((intTempo - (intTempo Mod 3600)) / 3600))
                strSaida.Append(" hora")
                If tmpParcial > 1 Then strSaida.Append("s")
                intTempo -= (tmpParcial * 3600)
                If intTempo > 0 Then strSaida.Append(", ")
            End If

            If (intTempo >= 60) Then
                strSaida.Append(((intTempo - (intTempo Mod 60)) / 60))
                strSaida.Append(" minuto")
                If tmpParcial > 1 Then strSaida.Append("s")
                intTempo -= (tmpParcial * 60)

                If intTempo > 0 Then strSaida.Append(" e ")
            End If

            If (intTempo > 0) Then
                strSaida.Append(intTempo)

                If intTempo = 1 Then strSaida.Append("segundo") Else strSaida.Append("segundos")
            End If
        End If
        If strSaida.ToString.Length < 1 Then Return "Desconhecido" Else Return strSaida.ToString()

    End Function

    Public Shared Function PegarDataAtualEmStringInverso() As String
        Dim dataAtual As DateTime = Date.Now

        Return Strings.Format(DateTime.Now, "yyyyMMdd").ToString()

    End Function
    Public Shared Function PegarDataHoraAtualEmStringInverso() As String
        Dim dataAtual As DateTime = Date.Now

        Return Strings.Format(DateTime.Now, "yyyyMMddHHmm").ToString()


    End Function



    ''' <summary>
    ''' Função que recebe uma data em texto em formato especifico e converte para data
    ''' </summary>
    ''' <param name="DataEntrada">recebe uma data em texto no formato YYYYMMDDhhmm ou YYYYMMDDhhmmss e retorna uma data</param>
    ''' <returns>Data e Hora</returns>
    Public Shared Function ConverteTextoParaData(ByVal DataEntrada As String, Optional ByVal ValueOnException As Date = MinDate) As DateTime
        'Parte 01 - É uma data apenas, ou é data hora?
        Dim formatoEsperado As String = "yyyyMMddHHmmss"
        Dim DataSaida As DateTime
        If Date.TryParseExact(DataEntrada, formatoEsperado.Substring(0, DataEntrada.Length), Globalization.CultureInfo.CreateSpecificCulture("pt-BR"), Globalization.DateTimeStyles.AllowLeadingWhite, DataSaida) Then
            Return DataSaida
        Else
            Return ValueOnException
        End If


    End Function


    ''' <summary>
    ''' Função que converte uma data em uma data XML em Padrão Brasil
    ''' </summary>
    ''' <example>Retorna 17-06-2021 20:21:00
    ''' </example>
    '''  <returns>Retorna String com a data no formato dd-MM-yyyy HH:mm:ss</returns>
    Public Shared Function ConverteDataParaTextoXMLDataBrasil(Data As Date) As String
        ' 17-06-2021 20:21:00
        Return Data.ToString("dd-MM-yyyy HH:mm:ss")
    End Function



    ''' <summary>
    ''' Tenta converter uma string para uma data com datahora (Datetime)
    ''' </summary>
    ''' <param name="DataEntrada">String contendo a data para conversão</param>
    ''' <returns>Retorna Data. Em caso de falha, retorna MinDate</returns>
    Public Shared Function TryCDateTime(ByVal DataEntrada As String, Optional ByVal ValueOnException As Date = MinDate) As DateTime

        'Parte 01 - É uma data apenas, ou é data hora?
        Dim qtdeDoisPontos As Integer = Texto.ContarQuantidadeOcorrencias(DataEntrada, ":")
        Dim formatoEsperado As String = "dd/MM/yyyy HH:mm:ss"
        If qtdeDoisPontos = 1 Then formatoEsperado = "dd/MM/yyyy HH:mm"
        If qtdeDoisPontos = 0 Then formatoEsperado = "dd/MM/yyyy"

        Dim DataSaida As DateTime
        If Date.TryParseExact(DataEntrada, formatoEsperado, Globalization.CultureInfo.CreateSpecificCulture("pt-BR"), Globalization.DateTimeStyles.AllowLeadingWhite, DataSaida) Then
            Return DataSaida
        Else
            Return ValueOnException
        End If
    End Function



    ''' <summary>
    ''' Função que verifica se o sistema está configurado com Data em Português
    ''' </summary>
    ''' <returns>Retorna True se Windows estiver com data configurado em Português Brasil</returns>
    Public Shared Function VerificarDataSistemaPortuguesBrasil() As Boolean
        Dim textoData As String = New System.DateTime(2011, 1, 1).AddDays(-1).ToString("dd/MM/yyyy")
        ' Brasil = "31/12/2010 00:00:00"
        ' EUA = "12/31/2010 12:00:00AM"
        Return Left(textoData, 10).ToString.Equals("31/12/2010")
    End Function





End Class
