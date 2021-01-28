Option Strict On
Option Explicit On

Imports System.Net
Imports Titanium.Web.Proxy
Imports Titanium.Web.Proxy.EventArguments
Imports Titanium.Web.Proxy.Models
Imports System.Text.RegularExpressions
Imports System.Text

Public Class SnifferEventArgs
    Inherits EventArgs
    Public Enum State
        Request
        Response
    End Enum
    Public Enum Timing
        Before
        After
    End Enum
    Public ReadOnly Property Method As String
    Public ReadOnly Property RequestURL As Uri
    Public ReadOnly Property Headers As New List(Of KeyValuePair(Of String, String))
    Public ReadOnly Property Key As String
    Public ReadOnly Property KeyTime As Date
    Public ReadOnly Property Traffic As State
    Public ReadOnly Property Sequence As Timing

    Public Sub New(e As SessionEventArgs, request As Boolean, Optional after As Boolean = False)

        If e Is Nothing Then Exit Sub
        _KeyTime = Now
        Dim Client As Http.HttpWebClient = e.HttpClient
        _Key = Client.UserData.ToString
        _Method = Client.Request.Method
        RequestURL = New Uri(Client.Request.Url)
        Traffic = If(request, State.Request, State.Response)
        Sequence = If(after, Timing.After, Timing.Before)
        For Each header In If(request, Client.Request.Headers, Client.Response.Headers)
            Headers.Add(New KeyValuePair(Of String, String)(header.Name, header.Value))
        Next

    End Sub
End Class
Public Class Sniffer
    ' https://github.com/justcoding121/Titanium-Web-Proxy/blob/develop/examples/Titanium.Web.Proxy.Examples.Basic/ProxyTestController.cs

    Implements IDisposable
    Public Event Alert(sender As Object, e As AlertEventArgs)
    Public Event RequestAlert(sender As Object, e As SnifferEventArgs)
    Public Event ResponseAlert(sender As Object, e As SnifferEventArgs)
    Public Event Found(sender As Object, e As SnifferEventArgs)

    Private WithEvents ProxyServer As ProxyServer

    Private ReadOnly ProxyPort As Integer = 18880

    Private SniffRequests As Integer
    Public Property Name As String
    Public Property Tag As Object
    Public ReadOnly Property Filters As New List(Of Filter)
    Public ReadOnly Property Sniffing As Boolean = False
    Public ReadOnly Property Body As String
    Public Sub New(Optional portNumber As Integer = 18880)
        ProxyPort = portNumber
    End Sub

    Public Sub StartSniffing()

        If Not Sniffing Then
            _Sniffing = True
            ProxyServer = New ProxyServer
            Dim explicitEndPoint As New ExplicitProxyEndPoint(IPAddress.Any, ProxyPort + SniffRequests, True)
            With ProxyServer
                .AddEndPoint(explicitEndPoint)
                .Start()
                .SetAsSystemHttpProxy(explicitEndPoint)
                .SetAsSystemHttpsProxy(explicitEndPoint)
            End With
            SniffRequests += 1
        End If

    End Sub
    Public Sub StopSniffing()

        _Sniffing = False
        If ProxyServer IsNot Nothing AndAlso ProxyServer.ProxyRunning Then
            Try
                ProxyServer.[Stop]()
            Catch ex As InvalidOperationException
            End Try
        End If

    End Sub

    Private Async Function Proxy_BeforeRequest(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeRequest
        Await Task.Run(Sub()
                           e.HttpClient.UserData = String.Format("{0:N}", Guid.NewGuid())
                           RaiseEvent RequestAlert(Me, New SnifferEventArgs(e, True, False))
                           RequestEvent(e, True)
                       End Sub).ConfigureAwait(False)
    End Function
    Private Async Function Proxy_BeforeResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeResponse

        RaiseEvent ResponseAlert(Me, New SnifferEventArgs(e, False, True))
        Await Task.Run(Async Function()
                           RaiseEvent ResponseAlert(Me, New SnifferEventArgs(e, False, False))
                           RequestEvent(e, False)
                           If e.HttpClient.Response.StatusCode = HttpStatusCode.OK And {"GET", "POST"}.Contains(e.HttpClient.Request.Method) Then
                               If e.HttpClient.Response.ContentType IsNot Nothing AndAlso e.HttpClient.Response.ContentType.Trim().ToLower().Contains("text/html") Then
                                   Dim bodyBytes As Byte() = Await e.GetResponseBody()
                                   e.SetResponseBody(bodyBytes)
                                   Dim responseBody As String = Await e.GetResponseBodyAsString()
                                   e.SetResponseBodyString(responseBody)
                                   Dim matchCount As Integer = 0
                                   Filters.ForEach(Sub(fltr)
                                                       Dim matchString As String = String.Empty
                                                       With fltr
                                                           If .Where = LookIn.Body Then
                                                               Dim matchBody As Match = Regex.Match(If(responseBody, String.Empty), .What, .How)
                                                               If matchBody.Success Then
                                                                   matchString = matchBody.Value
                                                                   matchCount += 1
                                                                   'StopSniffing()
                                                                   'Stop
                                                               End If
                                                           End If
                                                           If matchString.Any Then .Matches.Add(matchString)
                                                       End With
                                                   End Sub)
                                   Dim hasMatches As Boolean = matchCount > 0 And matchCount >= Filters.Count
                                   If hasMatches Then
                                       RaiseEvent Found(Me, New SnifferEventArgs(e, False, False))
                                       'StopSniffing()
                                       'Stop
                                   End If
                               End If
                           End If
                       End Function)

    End Function
    Private Async Function Proxy_AfterResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.AfterResponse

        RaiseEvent ResponseAlert(Me, New SnifferEventArgs(e, False, True))
        Await Task.Run(Sub()
                           Dim requestBody As Task(Of String) = e.GetResponseBodyAsString()
                           If requestBody IsNot Nothing Then
                           End If
                       End Sub)
    End Function
    Private Sub RequestEvent(e As SessionEventArgs, isRequest As Boolean)

        If Filters.Any Then
            RaiseEvent Alert(e.HttpClient.Request, New AlertEventArgs(e.HttpClient.Request.Url.ToUpperInvariant))
            Dim matchCount As Integer = 0
            Filters.ForEach(Sub(fltr)
                                Dim matchString As String = String.Empty
                                With fltr
                                    Select Case .Where
                                        Case LookIn.RequestHeaderNames
                                            For Each hdr In e.HttpClient.Request.Headers
                                                Dim matchHdrName As Match = Regex.Match(hdr.Name, .What, .How)
                                                If matchHdrName.Success Then
                                                    matchString &= matchHdrName.Value & "■"
                                                    matchCount += 1
                                                End If
                                            Next

                                        Case LookIn.RequestHeaderValues
                                            For Each hdr In e.HttpClient.Request.Headers
                                                Dim matchHdrValue As Match = Regex.Match(hdr.Value, .What, .How)
                                                If matchHdrValue.Success Then
                                                    matchString &= matchHdrValue.Value & "■"
                                                    matchCount += 1
                                                End If
                                            Next

                                        Case LookIn.ResponseHeaderNames
                                            For Each hdr In e.HttpClient.Response.Headers
                                                Dim matchHdrName As Match = Regex.Match(hdr.Name, .What, .How)
                                                If matchHdrName.Success Then
                                                    matchString &= matchHdrName.Value & "■"
                                                    matchCount += 1
                                                End If
                                            Next

                                        Case LookIn.ResponseHeaderValues
                                            For Each hdr In e.HttpClient.Response.Headers
                                                Dim matchHdrValue As Match = Regex.Match(hdr.Value, .What, .How)
                                                If matchHdrValue.Success Then
                                                    matchString &= matchHdrValue.Value & "■"
                                                    matchCount += 1
                                                End If
                                            Next

                                        Case LookIn.Host
                                            Dim matchHost As Match = Regex.Match(e.HttpClient.Request.RequestUri.ToString, .What, .How)
                                            If matchHost.Success Then
                                                matchCount += 1
                                                matchString = matchHost.Value
                                            End If

                                        Case LookIn.RequestURL
                                            Dim matchURL As Match = Regex.Match(e.HttpClient.Request.RequestUri.ToString, .What, .How)
                                            If matchURL.Success Then
                                                matchCount += 1
                                                matchString = matchURL.Value
                                            End If

                                    End Select
                                    If matchString.Any Then .Matches.AddRange(Split(matchString, "■"))
                                End With
                            End Sub)
            Dim hasMatches As Boolean = matchCount > 0 And matchCount >= Filters.Count
            If hasMatches Then RaiseEvent Found(Me, New SnifferEventArgs(e, isRequest, False))
        End If

    End Sub

#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                ProxyServer?.Dispose()

            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
Public Enum LookIn
    None
    Host
    RequestURL
    RequestHeaderNames
    RequestHeaderValues
    ResponseHeaderNames
    ResponseHeaderValues
    Body
End Enum

Public NotInheritable Class Filter
    Public Property What As String
    Public Property Where As LookIn
    Public Property How As RegexOptions = RegexOptions.IgnoreCase
    Public ReadOnly Property Matches As New List(Of String)
    Public Overrides Function ToString() As String
        Return $"{What} => {Where} [{Split(How.ToString, ".").First}]"
    End Function
End Class