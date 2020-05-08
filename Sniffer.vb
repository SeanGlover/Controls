Option Strict On
Option Explicit On

Imports System.Net
Imports Titanium.Web.Proxy
Imports Titanium.Web.Proxy.EventArguments
Imports Titanium.Web.Proxy.Models
Public Class SnifferEventArgs
    Inherits EventArgs
    Public ReadOnly Property Header As HttpHeader
    Public Sub New()
    End Sub
    Public Sub New(header As HttpHeader)
        Me.Header = header
    End Sub
End Class
Public Class Sniffer
    Public Event Alert(sender As Object, e As Controls.AlertEventArgs)
    Public Event Found(sender As Object, e As SnifferEventArgs)

    Private WithEvents ProxyServer As ProxyServer
    Private Const ProxyPort As Integer = 18880
    Public Property Name As String
    Public Property Tag As Object
    Public ReadOnly Property Requests As List(Of Http.Request)
        Get
            Return (From c In Clients Select c.Request).ToList
        End Get
    End Property
    Public ReadOnly Property Responses As List(Of Http.Response)
        Get
            Return (From c In Clients Select c.Response).ToList
        End Get
    End Property
    Public ReadOnly Property Clients As New List(Of Http.HttpWebClient)
    Public ReadOnly Property FindRequestHeaders As New Dictionary(Of String, List(Of String)) 'Key=Host, List of Header names
    Public ReadOnly Property FindResponseBody As New List(Of String) 'List of strings to watch for in Response.Body
    Public Property FindRequestURL As Uri
    Public ReadOnly Property FindRequestURLHeaders As New List(Of HttpHeader)
    Public ReadOnly Property ClientsString As New List(Of String)
    Public Sub New()
    End Sub

    Public Sub StartSniffing()

        ProxyServer = New ProxyServer
        Dim explicitEndPoint As New ExplicitProxyEndPoint(IPAddress.Any, ProxyPort, True)
        With ProxyServer
            .AddEndPoint(explicitEndPoint)
            .Start()
            .SetAsSystemHttpProxy(explicitEndPoint)
            .SetAsSystemHttpsProxy(explicitEndPoint)
        End With

    End Sub
    Public Sub StopSniffing()

        If ProxyServer IsNot Nothing AndAlso ProxyServer.ProxyRunning Then
            Try
                ProxyServer.[Stop]()
            Catch ex As InvalidOperationException
            End Try
        End If
        Clients.Clear()

    End Sub

    Private Async Function Proxy_BeforeRequest(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeRequest
        Await Task.Run(Sub()
                           Clients.Add(e.HttpClient)
                           Dim request As Http.Request = e.HttpClient.Request
                           RaiseEvent Alert(request, New Controls.AlertEventArgs(request.Url.ToUpperInvariant))
                           If FindRequestHeaders.ContainsKey(request.Host) Then
                               Dim lookForHeaders As New List(Of String)(From h In FindRequestHeaders(request.Host) Select h.ToUpperInvariant)
                               For Each header As HttpHeader In request.Headers
                                   If lookForHeaders.Contains(header.Name.ToUpperInvariant) Then
                                       RaiseEvent Found(Me, New SnifferEventArgs(header))
                                   End If
                               Next
                           End If
                       End Sub).ConfigureAwait(False)
    End Function
    Private Async Function Proxy_AfterRequest(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeResponse
        Await Task.Run(Sub()
                           Clients.Add(e.HttpClient)
                           Dim request As Http.Request = e.HttpClient.Request
                           Dim response As Http.Response = e.HttpClient.Response
                           RaiseEvent Alert(request, New AlertEventArgs(request.Url.ToUpperInvariant))
                           If FindRequestHeaders.ContainsKey(request.Host) Then
                               Dim lookForHeaders As New List(Of String)(From h In FindRequestHeaders(request.Host) Select h.ToUpperInvariant)
                               For Each header As HttpHeader In request.Headers
                                   If lookForHeaders.Contains(header.Name.ToUpperInvariant) Then
                                       RaiseEvent Found(Me, New SnifferEventArgs(header))
                                   End If
                               Next
                           ElseIf FindRequestURL?.ToString.ToUpperInvariant = request.RequestUri.ToString.ToUpperInvariant Then
                               FindRequestURLHeaders.AddRange(request.Headers)
                               RaiseEvent Found(Me, New SnifferEventArgs)

                           End If
                       End Sub).ConfigureAwait(False)
    End Function
    Private Async Function Proxy_AfterResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.AfterResponse
        Await Task.Run(Sub()
                           Dim response As Http.Response = e.HttpClient.Response
                           If Not response.StatusCode = 0 Then
                               RaiseEvent Alert(response, New Controls.AlertEventArgs(CStr(response.StatusCode)))
                               If response.HasBody Then
                                   For Each searchText In FindResponseBody
                                       'If response.BodyString.Contains(searchText) Then
                                       '    RaiseEvent Found(Me, New SnifferEventArgs)
                                       'End If
                                   Next
                               End If
                               ClientsString.Add(ClientToString(e.HttpClient))
                           End If
                       End Sub).ConfigureAwait(False)
    End Function
    Public Function ClientsToString() As String

        Return Join(ClientsString.ToArray, StrDup(20, vbNewLine & Controls.BlackOut & vbNewLine))

    End Function
    Private Function ClientToString(client As Titanium.Web.Proxy.Http.HttpWebClient) As String

        Dim clientData As New List(Of String)
        If client IsNot Nothing Then
            With client
                clientData.Add("Request" & StrDup(10, "="))
                clientData.Add(.Request.RequestUriString)
                For Each header In .Request.Headers
                    clientData.Add(Join({header.Name, header.Value}, "="))
                Next
                'Body
                clientData.Add("Response" & StrDup(10, "="))
                For Each header In .Response.Headers
                    clientData.Add(Join({header.Name, header.Value}, "="))
                Next
            End With
        End If
        Return Join(clientData.ToArray, vbNewLine)

    End Function
End Class