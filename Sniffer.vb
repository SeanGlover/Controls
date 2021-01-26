Option Strict On
Option Explicit On

Imports System.Net
Imports Titanium.Web.Proxy
Imports Titanium.Web.Proxy.EventArguments
Imports Titanium.Web.Proxy.Models
Imports System.Text.RegularExpressions

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
        _Key = String.Format("{0:N}", Guid.NewGuid())
        _KeyTime = Now
        Dim Client As Http.HttpWebClient = e.HttpClient
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
    Implements IDisposable
    Public Event Alert(sender As Object, e As AlertEventArgs)
    Public Event RequestAlert(sender As Object, e As SnifferEventArgs)
    Public Event ResponseAlert(sender As Object, e As SnifferEventArgs)
    Public Event Found(sender As Object, e As SnifferEventArgs)

    Private WithEvents ProxyServer As ProxyServer
    Private ReadOnly ProxyPort As Integer = 18880
    Private ReadOnly Property SniffRequests As Integer
    Public Property Name As String
    Public Property Tag As Object
    Public ReadOnly Property Clients As New List(Of Http.HttpWebClient)
    Private Client_ As Http.HttpWebClient
    Private Property Client As Http.HttpWebClient
        Get
            Return Client_
        End Get
        Set(value As Http.HttpWebClient)
            Client_ = value
            If Not Clients.Contains(value) Then Clients.Add(value)
        End Set
    End Property
    Public ReadOnly Property Filters As New List(Of Filter)
    Public ReadOnly Property ClientsString As New List(Of String)
    Public ReadOnly Property Sniffing As Boolean = False

    Public Sub New(Optional portNumber As Integer = 18880)
        ProxyPort = portNumber
    End Sub

    Public Sub StartSniffing()

        Clients.Clear()
        ClientsString.Clear()
        If Not Sniffing Then
            _Sniffing = True
            _SniffRequests += 1
            ProxyServer = New ProxyServer
            Dim explicitEndPoint As New ExplicitProxyEndPoint(IPAddress.Any, ProxyPort + SniffRequests, True)
            With ProxyServer
                .AddEndPoint(explicitEndPoint)
                .Start()
                .SetAsSystemHttpProxy(explicitEndPoint)
                .SetAsSystemHttpsProxy(explicitEndPoint)
            End With
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
                           Client = e.HttpClient
                           RaiseEvent RequestAlert(Me, New SnifferEventArgs(e, True, False))
                           RequestEvent(e, True)
                       End Sub).ConfigureAwait(False)
    End Function
    Private Async Function Proxy_BeforeResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeResponse
        Await Task.Run(Sub()
                           Client = e.HttpClient
                           RaiseEvent ResponseAlert(Me, New SnifferEventArgs(e, False, False))
                           RequestEvent(e, False)
                       End Sub).ConfigureAwait(False)
    End Function
    Private Async Function Proxy_AfterResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.AfterResponse
        Await Task.Run(Sub()
                           Client = e.HttpClient
                           RaiseEvent ResponseAlert(Me, New SnifferEventArgs(e, False, True))
                           Dim response = Client.Response
                           If Not response.StatusCode = 0 Then
                               RaiseEvent Alert(response, New AlertEventArgs(CStr(response.StatusCode)))
                               If response.HasBody Then '/// Errors trying to read BodyString as Response is still processing the reading of the Body
#Region " R E S P O N S E   B O D Y - I N A C T I V E (response.BodyString ) "
                                   '■■■■■■■■■■■■■■■■■■■■■■■■ Response body is not read yet.
                                   '■■■■■■■■■■■■■■■■■■■■■■■■ Use SessionEventArgs.GetResponseBody() Or SessionEventArgs.GetResponseBodyAsString() method to read the response body
                                   'Dim responseValues As New List(Of KeyValuePair(Of String, String))
                                   'Dim bodyResponse As String = response.BodyString
                                   'With Search
                                   '    If .By = FindyBy.Body Then
                                   '        If .Expression Is Nothing Then 'Explicit
                                   '            For Each searchValue In .Values
                                   '                If bodyResponse.Contains(searchValue) Then
                                   '                    responseValues.Add(New KeyValuePair(Of String, String)(searchValue, bodyResponse))
                                   '                End If
                                   '            Next
                                   '        Else 'Regex
                                   '            Dim bodyMatches = RegexMatches(bodyResponse, .Expression.SearchPattern, .Expression.SearchOptions)
                                   '            For Each bodyMatch In bodyMatches
                                   '                responseValues.Add(New KeyValuePair(Of String, String)(bodyMatch.Value, bodyResponse))
                                   '            Next
                                   '        End If
                                   '    End If
                                   'End With
                                   'If responseValues.Any Then
                                   '    For Each responseHead In response.Headers
                                   '        responseValues.Add(New KeyValuePair(Of String, String)(responseHead.Name, responseHead.Value))
                                   '    Next
                                   '    RaiseEvent Found(Me, New SnifferEventArgs(responseValues))
                                   'End If
#End Region
                               End If
                               ClientsString.Add(ClientToString(e.HttpClient))
                           End If
                       End Sub).ConfigureAwait(False)
    End Function
    Private Sub RequestEvent(e As SessionEventArgs, isRequest As Boolean)

        If Filters.Any Then
            Dim xxx = e.HttpClient.Request.Method
            Dim request As Http.Request = e.HttpClient.Request
            RaiseEvent Alert(request, New AlertEventArgs(request.Url.ToUpperInvariant))
            Dim headersRequest As New List(Of KeyValuePair(Of String, String))(request.Headers.Select(Function(h) New KeyValuePair(Of String, String)(h.Name, h.Value)))
            Dim headersResponse As New List(Of KeyValuePair(Of String, String))(e.HttpClient.Response.Headers.Select(Function(h) New KeyValuePair(Of String, String)(h.Name, h.Value)))
            Dim headers As New List(Of KeyValuePair(Of String, String))
            Dim matchCount As Integer = 0
            Filters.ForEach(Sub(fltr)
                                With fltr
                                    If .Where = LookIn.RequestHeaderNames Then matchCount += (From hr In headersRequest Where Regex.IsMatch(hr.Key, .What, .How)).Count
                                    If .Where = LookIn.RequestHeaderValues Then matchCount += (From hr In headersRequest Where Regex.IsMatch(hr.Value, .What, .How)).Count
                                    If .Where = LookIn.ResponseHeaderNames Then matchCount += (From hr In headersResponse Where Regex.IsMatch(hr.Key, .What, .How)).Count
                                    If .Where = LookIn.ResponseHeaderValues Then matchCount += (From hr In headersResponse Where Regex.IsMatch(hr.Value, .What, .How)).Count
                                    If .Where = LookIn.RequestURL Then matchCount += If(Regex.IsMatch(request.RequestUri.ToString, .What, .How), 1, 0)
                                    If .Where = LookIn.Host Then matchCount += If(Regex.IsMatch(request.Host, .What, .How), 1, 0)
                                End With
                            End Sub)
            Dim addHeaders As Boolean = matchCount > 0 And matchCount = Filters.Count
            If addHeaders Then RaiseEvent Found(Me, New SnifferEventArgs(e, isRequest, False))
        End If

    End Sub
    Public Function ClientsToString() As String
        Return Join((From cs In ClientsString Where cs IsNot Nothing).ToArray, StrDup(20, vbNewLine & Controls.BlackOut & vbNewLine))
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
    Public Property How As RegexOptions
End Class