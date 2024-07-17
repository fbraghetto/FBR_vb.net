Public NotInheritable Class SaidaConsole

    Public Shared Sub Output(ByVal message As String, ByVal m As System.Reflection.MethodBase)
        'Exemplo = "frmPlayerLiveProgramacaoV2.PlayerV2_Load: 20/07/2017 21:19:45: Player Iniciado Agora"
        System.Diagnostics.Trace.WriteLine(String.Concat(System.DateTime.Now.ToString, ": ", message), String.Concat("[", m.ReflectedType.Name, ".", m.Name, "]"))
    End Sub

    Public Shared Sub OutputBlankLine()
        Trace.WriteLine(" ")
    End Sub
End Class
