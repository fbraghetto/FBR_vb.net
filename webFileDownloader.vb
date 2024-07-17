Imports System.Net
Imports System.IO

Public Class WebFileDownloader
    Dim booContinue As Boolean = True
    Public Event AmountDownloadedChanged(ByVal iNewProgress As Long)
    Public Event FileDownloadSizeObtained(ByVal iFileSize As Long)
    Public Event FileDownloadComplete()
    Public Event FileDownloadFailed(ByVal ex As Exception)
    Public Event DownloadStartingUp()
    Private ReadOnly DadosLOG As New List(Of String)()



    Private mCurrentFile As String = String.Empty
    Public Function Cancel() As Boolean
        booContinue = False
        ' MsgBox("Cancelando...")
        Return True
    End Function

    Public ReadOnly Property CurrentFile() As String
        Get
            Return mCurrentFile
        End Get
    End Property
    Public Function DownloadFile(ByVal URL As String, ByVal Location As String) As Boolean
        '    Try
        Dim LocationTemp As String = Location & ".tmp"
        mCurrentFile = GetFileName(URL)
        Dim WC As New WebClient
        WC.DownloadFile(URL, LocationTemp)
        If Err.Number = 0 Then
            If File.Exists(Location) Then File.Delete(Location)
            File.Move(LocationTemp, Location)
        End If
        RaiseEvent FileDownloadComplete()
        Return True
        'Catch ex As Exception
        '    RaiseEvent FileDownloadFailed(ex)
        '    Return False
        'End Try
    End Function

    Private Function GetFileName(ByVal URL As String) As String
        Try
            Return URL.Substring(URL.LastIndexOf("/") + 1)
        Catch ex As Exception
            Return URL
        End Try
    End Function
    Public Function DownloadFileWithProgress(ByVal URL As String, ByVal Location As String) As Boolean
        Dim FS As FileStream
        Dim LocationTemp As String = Location & ".tmp"
        Dim returnValue As Boolean = False

        If File.Exists(LocationTemp) Then
            Try
                File.Delete(LocationTemp)
            Catch ex As Exception
                DadosLOG.Add("[cls_webFileDownloader.DownloadFileWithProgress] Erro ao Deletar arq temp: " & LocationTemp)
                DadosLOG.Add("[cls_webFileDownloader.DownloadFileWithProgress] Erro Detalhado: " & ex.ToString)
                RaiseEvent FileDownloadFailed(ex)
            End Try
        End If

        Try
            mCurrentFile = GetFileName(URL)
            Dim wRemote As WebRequest
            Dim bBuffer As Byte()
            ReDim bBuffer(512)
            Dim iBytesRead As Integer
            Dim iTotalBytesRead As Integer

            Dim intSegundoUltAtualizacao As Integer = System.DateTime.Now.Second

            RaiseEvent DownloadStartingUp()

            FS = New FileStream(LocationTemp, FileMode.Create, FileAccess.Write)
            wRemote = WebRequest.Create(URL)
            Dim myWebResponse As WebResponse = wRemote.GetResponse
            RaiseEvent FileDownloadSizeObtained(myWebResponse.ContentLength)
            Dim sChunks As Stream = myWebResponse.GetResponseStream
            Do
                iBytesRead = sChunks.Read(bBuffer, 0, 512)
                FS.Write(bBuffer, 0, iBytesRead)
                iTotalBytesRead += iBytesRead
                If myWebResponse.ContentLength < iTotalBytesRead Then
                    If (intSegundoUltAtualizacao <> System.DateTime.Now.Second) Then
                        ' Só levanta update uma vez a cada segundo
                        intSegundoUltAtualizacao = System.DateTime.Now.Second
                        RaiseEvent AmountDownloadedChanged(myWebResponse.ContentLength)
                    End If
                Else
                    If (intSegundoUltAtualizacao <> System.DateTime.Now.Second) Then
                        intSegundoUltAtualizacao = System.DateTime.Now.Second
                        RaiseEvent AmountDownloadedChanged(iTotalBytesRead)
                    End If
                End If
                If booContinue = False Then
                    Exit Do
                End If
            Loop While Not iBytesRead = 0
            sChunks.Close()
            FS.Close()
            If Err.Number = 0 And booContinue Then
                ' Veio aqui pois esta tudo bem
                If File.Exists(Location) Then File.Delete(Location)
                File.Move(LocationTemp, Location)
                RaiseEvent FileDownloadComplete()
                returnValue = True
            Else
                ' Passou aqui porque deu erro ou foi cancelado
                If File.Exists(LocationTemp) Then File.Delete(LocationTemp)
                If Not booContinue Then
                    returnValue = False
                End If

            End If
        Catch ex As Exception

            If File.Exists(LocationTemp) Then
                Try
                    File.Delete(LocationTemp)
                Catch ex2 As Exception
                    Console.WriteLine("Erro ao Deletar arquivo " & LocationTemp)
                End Try

            End If
            RaiseEvent FileDownloadFailed(ex)
            Return False

        End Try
        Return returnValue
    End Function

    Public Shared Function FormatFileSize(ByVal Size As Long) As String
        Try
            Dim KB As Integer = 1024
            Dim MB As Integer = 1048576
            Dim GB As Long = 1073741824
            Dim Resultado As Double
            ' Return size of file in kilobytes.
            If Size < KB Then
                Return (Math.Round(Size) & "bytes")
            ElseIf Size < MB Then
                Resultado = Size / KB
                Return (Math.Round(Resultado, 1) & "KB")
            ElseIf Size < GB Then
                Resultado = Size / MB
                Return (Math.Round(Resultado, 1) & "MB")
            Else
                Resultado = Size / GB
                Return (Math.Round(Resultado, 2) & "GB")
            End If

        Catch ex As Exception
            Return Size.ToString
        End Try
    End Function

    Private Sub LOGADD(ByVal str As String)
        If DadosLOG.Capacity <> 200 Then DadosLOG.Capacity = 200
        If DadosLOG.Count > 99 Then DadosLOG.RemoveAt(1)
        DadosLOG.Add(str)
    End Sub

    Private Function getLOG() As String()
        Return DadosLOG.ToArray()
    End Function

End Class
