Public NotInheritable Class Numero
    Public Shared Function Entre(ByVal valor As Long, ByVal Minimo As Long, ByVal Maximo As Long) As Boolean
        If valor >= Minimo And valor <= Maximo Then Return True Else Return False
    End Function
    Public Shared Function Entre(ByVal valor As Decimal, ByVal Minimo As Decimal, ByVal Maximo As Decimal) As Boolean
        If valor >= Minimo And valor <= Maximo Then Return True Else Return False
    End Function
    Public Shared Function Entre(ByVal valor As Integer, ByVal Minimo As Integer, ByVal Maximo As Integer) As Boolean
        If valor >= Minimo And valor <= Maximo Then Return True Else Return False
    End Function
    Public Shared Function Menor(ByVal n1 As Integer, ByVal n2 As Integer) As Integer
        Return Math.Min(n1, n2)
    End Function
    Public Shared Function Menor(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer) As Integer
        Return Math.Min(Math.Min(n1, n2), Math.Min(n2, n3))
    End Function
    Public Shared Function Menor(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer) As Integer
        Return Math.Min(Menor(n1, n2, n3), n4)
    End Function
    Public Shared Function Maior(ByVal n1 As Integer, ByVal n2 As Integer) As Integer
        Return Math.Max(n1, n2)
    End Function
    Public Shared Function Maior(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer) As Integer
        Return Math.Max(Math.Min(n1, n2), Math.Min(n2, n3))
    End Function
    Public Shared Function Maior(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer, ByVal n4 As Integer) As Integer
        Return Math.Max(Maior(n1, n2, n3), n4)
    End Function
    Public Shared Function QualMaior(ByVal n1 As Integer, ByVal n2 As Integer) As Integer
        ' retorna 1 se n1 é maior, 2 se n1 é maior, 0 se for igual
        If n1 > n2 Then
            Return 1
        ElseIf n2 > n1 Then
            Return 2
        Else
            Return 0
        End If
    End Function
    Public Shared Function QualMaior(ByVal n1 As Integer, ByVal n2 As Integer, ByVal n3 As Integer) As Integer
        ' retorna 1 se n1 é maior, 2 se n2 é maior, 3 se o n3 é maior, 0 se for igual
        If n1 > n2 And n1 > n3 Then
            Return 1
        ElseIf n2 > n1 And n2 > n3 Then
            Return 2
        ElseIf n3 > n1 And n3 > n2 Then
            Return 3
        Else
            Return 0
        End If
    End Function

    Public Shared Function TryCLng(ByVal strString As String, Optional ByVal ValueOnException As Long = 0) As Long
        Dim resultado As Long
        If Long.TryParse(strString, resultado) Then
            Return resultado
        Else
            Return ValueOnException
        End If
    End Function

    Public Shared Function TryCInt(ByVal strString As String, Optional ByVal ValueOnException As Integer = 0) As Integer
        Try
            Return CInt(strString)
        Catch ex As OverflowException
            Return ValueOnException
        End Try
    End Function

    Public Shared Function TryCDbl(ByVal strString As String, Optional ByVal ValueOnException As Double = 0) As Double
        Try
            Return CDbl(strString)
        Catch ex As OverflowException
            Return ValueOnException
        End Try
    End Function

    Public Shared Function LPAD(ByVal entrada As Integer, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return entrada.ToString.PadLeft(qtdeDigitosTotal, System.Convert.ToChar(Enchimento))
    End Function
    Public Shared Function PadLeft(ByVal entrada As Integer, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return entrada.ToString.PadLeft(qtdeDigitosTotal, System.Convert.ToChar(Enchimento))
    End Function

    Public Shared Function LPAD(ByVal entrada As Long, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return entrada.ToString.PadLeft(qtdeDigitosTotal, System.Convert.ToChar(Enchimento))
    End Function
    Public Shared Function PadLeft(ByVal entrada As Long, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return entrada.ToString.PadLeft(qtdeDigitosTotal, System.Convert.ToChar(Enchimento))
    End Function

    Public Shared Function LPAD(ByVal entrada As Double, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return entrada.ToString.PadLeft(qtdeDigitosTotal, System.Convert.ToChar(Enchimento))
    End Function
    Public Shared Function PadLeft(ByVal entrada As Double, ByVal qtdeDigitosTotal As Integer, Optional ByVal Enchimento As String = "0") As String
        Return entrada.ToString.PadLeft(qtdeDigitosTotal, System.Convert.ToChar(Enchimento))
    End Function


    Function RetornaValorNegativo(ByVal d As Double) As Double
        Return Math.Abs(d) * -1
    End Function
    Function RetornaValorNegativo(ByVal i As Integer) As Integer
        Return Math.Abs(i) * -1
    End Function
    Function RetornaValorNegativo(ByVal d As Decimal) As Decimal
        Return Math.Abs(d) * -1
    End Function
    Function RetornaValorNegativo(ByVal l As Long) As Long
        Return Math.Abs(l) * -1
    End Function

End Class
