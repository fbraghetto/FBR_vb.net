Imports System.IO
Public Class LOGConsole
    Public Event EventoNovoLOG()
    Private strArquivoLOG As String
    Private booEfetuaLOGArquivo As Boolean = True       ' Efetua ou não LOG em Arquivo?
    Private intIDInstancia As Integer

    ' Estrutura LOG
    Private txtArrayLOG(100) As String
    Private intArrayLOGLastPos As Integer
    Private strDebugLastLog As String = ""
    Private strDebugLastLogSemData As String = ""
    Private intDebugLastLogQtdeDuplicados As Integer = 1
    Private cNewLine As String = vbCrLf

    Public Enum TipoArquivoLOG As Integer
        LOGSISTEMA = 2
        LOGDEBUG = 3
    End Enum
    Public Enum TipoLOG As Integer
        LOG_NENHUM = -5
        LOG_MAJOR_ERR = -2
        LOG_ERR = -1
        LOG_OK = 0
        LOG_WARN = 1
        LOG_INFO = 2
        LOG_WARNCONFIG = 3
    End Enum

    Public Property ArquivoLOG() As String
        Get
            Return strArquivoLOG
        End Get
        Set(ByVal strNomeArquivo As String)
            strArquivoLOG = strNomeArquivo
            If strArquivoLOG.Length > 0 Then booEfetuaLOGArquivo = True Else booEfetuaLOGArquivo = False
        End Set
    End Property
    Public Property EfetuaLOGArquivo() As Boolean
        Get
            Return booEfetuaLOGArquivo
        End Get
        Set(ByVal status As Boolean)
            booEfetuaLOGArquivo = status
        End Set
    End Property

    Public ReadOnly Property strLinhasLOG() As String()
        Get
            Return txtArrayLOG
        End Get
    End Property
    Public ReadOnly Property getDebugLastLog() As String
        Get
            Return strDebugLastLog
        End Get
    End Property
    Public ReadOnly Property IDInstancia() As Integer
        Get
            Return intIDInstancia
        End Get
    End Property

    Public Sub New()
        'Construtor
        ' Inicializa as variáveis
        strDebugLastLog = ""
        strDebugLastLogSemData = ""
        intDebugLastLogQtdeDuplicados = 1
        booEfetuaLOGArquivo = True
        SaidaConsole.Output("[log_debug] Construtor do LOG chamado sem parametros", System.Reflection.MethodBase.GetCurrentMethod())

        ' Gera um número aleatório
        Dim seed As Integer = CInt(Strings.Right(DateTime.Now.Ticks, 6))
        Dim rand1 As New Random(seed)
        intIDInstancia = rand1.Next(100000, 999999)

    End Sub
    Public Sub New(ByVal strNomeArquivo As String)
        'Construtor 2
        strDebugLastLog = ""
        strDebugLastLogSemData = ""
        intDebugLastLogQtdeDuplicados = 1
        booEfetuaLOGArquivo = True

        ' Gera um número aleatório
        Dim seed As Integer = CInt(Strings.Right(DateTime.Now.Ticks, 6))
        Dim rand1 As New Random(seed)
        intIDInstancia = rand1.Next(100000, 999999)

        ' Setar Propriedade de LOG
        ArquivoLOG = strNomeArquivo
        SaidaConsole.Output("[log_debug] Construtor do LOG chamado com arquivo = " & strNomeArquivo, System.Reflection.MethodBase.GetCurrentMethod())
    End Sub

    Public Function fcnTipoLOGtoString(ByVal tipo As TipoLOG) As String
        Dim strSaida As String = ""
        Select Case tipo
            Case TipoLOG.LOG_MAJOR_ERR
                strSaida = "ERRO GRAVE"
            Case TipoLOG.LOG_ERR
                strSaida = "Erro"
            Case TipoLOG.LOG_OK
                strSaida = "Sucesso"
            Case TipoLOG.LOG_WARN
                strSaida = "Alerta"
            Case TipoLOG.LOG_INFO
                strSaida = "Info"
            Case TipoLOG.LOG_WARNCONFIG
                strSaida = "Alerta de Config"
            Case TipoLOG.LOG_NENHUM
                strSaida = ""
            Case Else
                strSaida = ""
        End Select
        Return strSaida
    End Function
    Public Function fcnTipoLOGtoInt(ByVal tipo As TipoLOG) As Integer
        Dim intSaida As Integer = -1000
        intSaida = tipo
        Return intSaida
    End Function

    Public Sub LOG(ByVal textoLOG As String,
                    Optional ByVal tipoLOG As TipoLOG = TipoLOG.LOG_INFO,
                    Optional ByVal naoProcessar As Boolean = False)

        Dim strMsgLOGAtual As String = ""
        Dim strTipoLOG As String
        strTipoLOG = fcnTipoLOGtoString(tipoLOG)

        If naoProcessar = True Then
            strMsgLOGAtual = strTipoLOG & textoLOG
        Else
            ' Processamento de LOG
            ' Replaces - Evita Log com caracteres "reservados"
            textoLOG = textoLOG.Replace("|", "!")
            textoLOG = textoLOG.Replace(Environment.NewLine, "//")
            textoLOG = textoLOG.Replace(vbCr, " ")
            textoLOG = textoLOG.Replace(vbLf, " ")

            strMsgLOGAtual = strTipoLOG & ": " & textoLOG
        End If

        ' Evitar Duplicação
        If strMsgLOGAtual = strDebugLastLogSemData Then
            intDebugLastLogQtdeDuplicados += 1
            SaidaConsole.Output(DateTime.Now.ToString & " - Suprimido mensagem no LOG em prol de evitar duplicação (Qtde=" & intDebugLastLogQtdeDuplicados & "): " & strMsgLOGAtual, System.Reflection.MethodBase.GetCurrentMethod())
        Else
            If intDebugLastLogQtdeDuplicados > 2 Then
                ' Falar que ja foi feito mais de uma vez
                Dim c2 As String = DateTime.Now & " - " & strTipoLOG & ": "
                fcnLOG_Simples(c2 & "A última mensagem de LOG foi repetida " & intDebugLastLogQtdeDuplicados & " vezes.")
                intDebugLastLogQtdeDuplicados = 1
            End If

            'Ultimo LOG
            strDebugLastLog = DateTime.Now & " - " & strMsgLOGAtual
            strDebugLastLogSemData = strMsgLOGAtual

            '
            If naoProcessar = True Then
                ' Sem processamento - Nao inclui data
                fcnLOG_Simples(strDebugLastLogSemData)
            Else
                ' Processamento de LOG
                fcnLOG_Simples(strDebugLastLog)
            End If

        End If
    End Sub
    Private Sub fcnLOG_Simples(ByVal textoLOG As String)
        ' Esta função faz LOG porém sem processá-las
        ' Então não deve ser chamada externamente. Deve ser usada apenas  pela função fcnLOG()

        ' Esta função serve para gerar LOGs
        ' Primeiro passo é inverter, ou seja

        SaidaConsole.Output("[log_debug] fcnLOGSimples chamado. Texto: " & textoLOG, System.Reflection.MethodBase.GetCurrentMethod())

        If intArrayLOGLastPos > 0 Then
            Dim i As Integer
            For i = 99 To 1 Step -1
                txtArrayLOG(i) = txtArrayLOG(i - 1)
            Next
        End If

        ' LOG erro em arquivo se habilitado
        If booEfetuaLOGArquivo Then
            fcnLOGArquivo(textoLOG)
        Else
            SaidaConsole.Output("Log em Arquivo desabilitado. Conteúdo: " & textoLOG, System.Reflection.MethodBase.GetCurrentMethod())
        End If


        ' Preparar para LOG erro em tela
        txtArrayLOG(0) = textoLOG
        intArrayLOGLastPos = intArrayLOGLastPos + 1
        'gfrmPrincipal.txtLOGSistema.Lines = txtArrayLOG
        'gfrmPrincipal.txtLOGSistema.Refresh()

        RaiseEvent EventoNovoLOG()
    End Sub

    Public Sub fcnLOGArquivo(ByVal strEntrada As String)
        If strArquivoLOG = "" Then
            'Throw New System.Exception("Não foi setado o arquivo a fazer log. Verifique propriedade ArquivoLOG")
            'MsgBox("Arquivo de LOG setado incorretamente = **" & strArquivoLOG & "** / Instancia=" & intIDInstancia)

        End If

        'Check for existence of logger file
        If Not System.IO.Directory.Exists(Path.GetDirectoryName(strArquivoLOG)) Then
            MsgBox("Diretório de LOG não Existe: " & Path.GetDirectoryName(strArquivoLOG))
            'Dim result1 As DialogResult = MessageBox.Show("Diretório de LOG não Existe: " & Path.GetDirectoryName(strArquivoLOG) & vbCrLf &
            '            "Criar o diretório automaticamente?", "Example", MessageBoxButtons.YesNo)
            'If result1 = DialogResult.Yes Then
            '    Try
            '        Directory.CreateDirectory(Path.GetDirectoryName(strArquivoLOG))
            '    Catch ex As Exception
            '        MsgBox("Não foi possível criar diretório : " & Path.GetDirectoryName(strArquivoLOG))
            '    End Try

            'End If
        Else
            If File.Exists(strArquivoLOG) Then
                If (FBR.Arquivo.PegarTamanhoArquivo(strArquivoLOG) > 10000000) Then
                    ' Arquivo maior que 10Mbytes
                    Try

                        Dim NovoFilename As String = Replace(strArquivoLOG, ".log", "") & "_2.log"
                        If File.Exists(NovoFilename) Then File.Delete(NovoFilename)
                        File.Move(strArquivoLOG, NovoFilename)
                    Catch ex As Exception
                        ' Erro não tratado
                        SaidaConsole.Output(Date.Now & "- mod_LOG.fcnLOGArquivo() - Excessão qdo tamanho >10MB: " & ex.ToString, System.Reflection.MethodBase.GetCurrentMethod())
                    End Try
                End If
                Try
                    Dim fs As FileStream = New FileStream(strArquivoLOG, FileMode.Append, FileAccess.Write)
                    Dim sw As StreamWriter = New StreamWriter(fs)
                    Dim strTmp As String = strEntrada.Replace(cNewLine, "").Replace(ControlChars.Lf, "").Replace(ControlChars.Cr, "")
                    sw.WriteLine(DateTime.Now + "|" + strTmp)
                    sw.Close()
                    fs.Close()
                Catch Ex As Exception
                    ' Erro não tratado
                    SaidaConsole.Output(Date.Now & "- mod_nucleo.fcnLOGArquivo() - Falha ao Gravar no arquivo de Log " & strArquivoLOG & ": " & Ex.ToString, System.Reflection.MethodBase.GetCurrentMethod())
                End Try
            Else
                'If file doesn't exist create one
                Try
                    Dim fileStream As FileStream = File.Create(strArquivoLOG)
                    Dim sw As StreamWriter = New StreamWriter(fileStream)
                    Dim strTmp As String = strEntrada.Replace(cNewLine, "").Replace(ControlChars.Lf, "").Replace(ControlChars.Cr, "")
                    sw.WriteLine(DateTime.Now + "|" + strTmp)
                    sw.Close()
                    fileStream.Close()
                Catch ex As Exception

                End Try
            End If
        End If



    End Sub

    Function fcnSimpleLOG(ByVal MensagemLOG As String) As Boolean
        Throw New NotImplementedException

        'Try
        '    Dim fi As New FileInfo(gstrArquivoLOG)
        '    If fi.Length() > 10000000 Then
        '        Dim oldLog As String = Path.Combine(Path.GetDirectoryName(gstrArquivoLOG), Path.GetFileNameWithoutExtension(gstrArquivoLOG) & "_old.log")
        '        If File.Exists(oldLog) Then File.Delete(oldLog)
        '        File.Move(gstrArquivoLOG, oldLog)
        '    End If
        'Catch ex As Exception

        'End Try
        'Return fcnGravarArquivoTexto(gstrArquivoLOG, MensagemLOG & vbCrLf)
    End Function


End Class
