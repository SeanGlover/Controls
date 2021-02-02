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
    Public ReadOnly Property Payload As String
    Public ReadOnly Property Body As String
    Public ReadOnly Property Code_cSharp As String
    Public ReadOnly Property RequestURL As Uri
    Public ReadOnly Property RequestHeaders As New List(Of KeyValuePair(Of String, String))
    Public ReadOnly Property ResponseHeaders As New List(Of KeyValuePair(Of String, String))
    Public ReadOnly Property Key As String
    Public ReadOnly Property KeyTime As Date
    Public ReadOnly Property Traffic As State

    Public Sub New(sender As Sniffer, e As SessionEventArgs, isRequest As Boolean, Optional isFound As Boolean = False)

        If e Is Nothing Then Exit Sub
        KeyTime = Now
        Dim Client As Http.HttpWebClient = e.HttpClient
        Key = Client.UserData.ToString
        Method = Client.Request.Method
        RequestURL = New Uri(Client.Request.Url)

        For Each hdr In Client.Request.Headers
            RequestHeaders.Add(New KeyValuePair(Of String, String)(hdr.Name, hdr.Value))
        Next
        For Each hdr In Client.Response.Headers
            ResponseHeaders.Add(New KeyValuePair(Of String, String)(hdr.Name, hdr.Value))
        Next
        If isRequest Then
            Traffic = State.Request
            Payload = If(sender.Payloads.ContainsKey(_Key), sender.Payloads(_Key), Nothing)

            If sender.Code_AllRequests Or sender.Code_FoundRequests And isFound Then
                Dim lines As New List(Of String) From
{
$"HttpWebRequest xRqst = (HttpWebRequest)(WebRequest.Create(""{RequestURL}""));",
$"xRqst.Method =""{_Method.ToUpperInvariant}"";"
}
                Dim contentType As Content_Type
                'application/x-json-stream
                'application/octet-stream
                'application/x-www-form-urlencoded
                'application/x-www-form-urlencoded;charset=UTF-8
                'application/binary
                RequestHeaders.ForEach(Sub(hdr)
                                           Dim headName As String = hdr.Key.ToLowerInvariant
                                           Dim headValue As String = hdr.Value.ToLowerInvariant
                                           Select Case headName
                                               Case "connection"
                                                   lines.Add($"xRqst.KeepAlive = {(headValue = "keep-alive").ToString.ToLowerInvariant};")

                                               Case "accept"
                                                   'copyRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*;q=0.8,application/signed-exchange;v=b3;q=0.9";
                                                   lines.Add($"xRqst.Accept = ""{hdr.Value}"";")

                                               Case "accept-language"
                                                   'copyRequest.Headers.Set(HttpRequestHeader.AcceptLanguage, "en-US,en;q=0.5");
                                                   lines.Add($"xRqst.Headers.Set(HttpRequestHeader.AcceptLanguage, ""{hdr.Value}"");")

                                               Case "accept-encoding"
                                                   'copyRequest.Headers.Set(HttpRequestHeader.AcceptEncoding, "gzip, deflate, br");
                                                   lines.Add($"xRqst.Headers.Set(HttpRequestHeader.AcceptEncoding, ""{hdr.Value}"");")

                                               Case "cache-control"
                                                   'copyRequest.Headers.Add("Cache-Control", "max-age=0");
                                                   lines.Add($"xRqst.Headers.Add(""Cache-Control"", ""{hdr.Value}"");")

                                               Case "content-type"
                                                   'cisRequest.ContentType = "application/x-www-form-urlencoded";
                                                   lines.Add($"xRqst.ContentType = ""{hdr.Value}"";")
                                                   For Each item In [Enum].GetNames(GetType(Content_Type))
                                                       If headValue.Contains(item.Replace("_", "-")) Then
                                                           contentType = CType([Enum].Parse(GetType(Content_Type), item, True), Content_Type)
                                                       End If
                                                   Next

                                               Case "content-length"
                                                 'Handled in POST SetPayload

                                               Case "cookie"
                                                   'copyRequest.Headers.Set(HttpRequestHeader.Cookie, IBM_Cookies);
                                                   lines.Add($"xRqst.Headers.Set(HttpRequestHeader.Cookie, ""{hdr.Value}"");")

                                               Case "expect"
                                                   lines.Add($"xRqst.Expect = ""{hdr.Value}"";")

                                               Case "host"
                                                   lines.Add($"xRqst.Host = ""{hdr.Value}"";")

                                               Case "If-modified-since"
                                                   lines.Add($"xRqst.IfModifiedSince = ""{hdr.Value}"";")

                                               Case "origin"
                                                 'Do nothing with this

                                               Case "referer"
                                                   'copyRequest.Referer = "https://w3-01.ibm.com/isc/customerfulfillment/tools/cisinvoicing/mivweb/us/SearchInvoice.wss";
                                                   lines.Add($"xRqst.Referer = ""{hdr.Value}"";")

                                               Case "sec-fetch-site", "sec-fetch-mode", "sec-fetch-user", "sec-fetch-dest", "upgrade-insecure-requests"
                                                   'copyRequest.Headers.Add("Sec-Fetch-Site", "same-origin");
                                                   'copyRequest.Headers.Add("Sec-Fetch-Mode", "navigate");
                                                   'copyRequest.Headers.Add("Sec-Fetch-User", "?1");
                                                   'copyRequest.Headers.Add("Sec-Fetch-Dest", "document");
                                                   'copyRequest.Headers.Add("Upgrade-Insecure-Requests", "1");
                                                   lines.Add($"xRqst.Headers.Add(""{hdr.Key}"", ""{hdr.Value}"");")

                                               Case "user-agent"
                                                   'copyRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36";
                                                   lines.Add($"xRqst.UserAgent = ""{hdr.Value}"";")

                                               Case Else
                                                   lines.Add($"xRqst.Headers.Add(""{hdr.Key}"", ""{hdr.Value}"");")

                                           End Select
                                           'copyRequest.ServicePoint.Expect100Continue = False;
                                       End Sub)

                If If(Payload, String.Empty).Any And {"POST", "PUT"}.Contains(Method.ToUpperInvariant) Then
                    If contentType = Content_Type.www_form_urlencoded Then
                        Try
                            Dim fields As New List(Of String())(Split(Payload, "&").Select(Function(f) Split(f, "=")))
                            fields.Sort(Function(x, y)
                                            Return x(0).CompareTo(y(0))
                                        End Function)
                            Dim payString As String = String.Join(Environment.NewLine, fields.Where(Function(f) f(1).Any).Select(Function(f) "{" & $" ""{f(0)}"", ""{f(1)}"" " & "},").ToArray())
                            payString = payString.Remove(payString.Length - 1, 1)
                            lines.Add("Dictionary<string, object> parameters = new Dictionary<string, object>() {" + $"{Environment.NewLine}{payString}{Environment.NewLine}" + "};")

                        Catch ex As IndexOutOfRangeException
                            lines.Add("Dictionary<string, object> parameters = new Dictionary<string, object>() {" + $"{Environment.NewLine}{Payload}{Environment.NewLine}" + "};")

                        End Try

                    ElseIf contentType = Content_Type.x_json_stream Then
                        lines.Add("Dictionary<string, object> parameters = new Dictionary<string, object>() {" + $"{Environment.NewLine}{Payload}{Environment.NewLine}" + "};")

                    End If
                    lines.Add("SetPayload(parameters, xRqst);")
                End If

                lines.AddRange({"try", "{", "HttpWebResponse xResponse = (HttpWebResponse)xRqst.GetResponse();", "}"})
                lines.AddRange({"catch", "{", "}"})
                Code_cSharp = Join(lines.ToArray, Environment.NewLine)
            End If

        Else
            Traffic = State.Response
            Body = If(sender.Bodies.ContainsKey(_Key), sender.Bodies(_Key), Nothing)
        End If

    End Sub
End Class
Public Enum Content_Type
    none
    x_json_stream
    octet_stream
    www_form_urlencoded '...also: application/x-www-form-urlencoded;charset=UTF-8
    binary
End Enum
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
    Public Property Code_FoundRequests As Boolean
    Public Property Code_AllRequests As Boolean
    Public Property Name As String
    Public Property Tag As Object
    Public ReadOnly Property Filters As New List(Of Filter)
    Public ReadOnly Property Sniffing As Boolean = False
    Friend ReadOnly Property Payloads As New Dictionary(Of String, String)
    Friend ReadOnly Property Bodies As New Dictionary(Of String, String)
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

#Region " Rollback "
    'Private Async Function Proxy_BeforeRequest(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeRequest
    '    Await Task.Run(Sub()
    '                       e.HttpClient.UserData = String.Format("{0:N}", Guid.NewGuid())
    '                       RaiseEvent RequestAlert(Me, New SnifferEventArgs(e, True, False))
    '                       ProxyEvent(e, True)
    '                   End Sub).ConfigureAwait(False)
    'End Function
    'Private Async Function Proxy_BeforeResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeResponse
    '    Await Task.Run(Sub()
    '                       RaiseEvent ResponseAlert(Me, New SnifferEventArgs(e, False, True))
    '                       ProxyEvent(e, False)
    '                   End Sub).ConfigureAwait(False)
    'End Function
#End Region

    Private Async Function Proxy_BeforeRequest(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeRequest

        Dim clientKey As String = String.Format("{0:N}", Guid.NewGuid())
        e.HttpClient.UserData = clientKey

        Await Task.Run(Async Function()
                           If e.HttpClient.Request.HasBody Then
                               Dim requestBody As String = Await e.GetRequestBodyAsString()
                               If requestBody.Any Then
                                   Payloads.Add(clientKey, requestBody)
                                   'LookIn.Payload Code!

                               End If
                           End If
                       End Function).ConfigureAwait(False)
        Await Task.Run(Sub()
                           RaiseEvent RequestAlert(Me, New SnifferEventArgs(Me, e, True))
                           ProxyEvent(e, True)
                       End Sub).ConfigureAwait(False)

    End Function
    Private Async Function Proxy_BeforeResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeResponse

        Await Task.Run(Async Function()
                           If e.HttpClient.Response.StatusCode = HttpStatusCode.OK And {"GET", "POST"}.Contains(e.HttpClient.Request.Method) Then
                               If e.HttpClient.Response.ContentType IsNot Nothing AndAlso e.HttpClient.Response.ContentType.Trim().ToLower().Contains("text/html") Then
                                   Dim bodyBytes As Byte() = Await e.GetResponseBody()
                                   e.SetResponseBody(bodyBytes)
                                   Dim responseBody As String = Await e.GetResponseBodyAsString()
                                   e.SetResponseBodyString(responseBody)

                                   If responseBody.Any Then
                                       Dim clientKey As String = e.HttpClient.UserData.ToString
                                       Bodies.Add(clientKey, responseBody)

                                       Dim matchCount As Integer = 0
                                       Filters.ForEach(Sub(fltr)
                                                           Dim matchString As String = String.Empty
                                                           With fltr
                                                               If .Active Then
                                                                   If .Where = LookIn.Body Then
                                                                       Dim matchBody As Match = Regex.Match(If(responseBody, String.Empty), .What, .How)
                                                                       If matchBody.Success Then
                                                                           matchString = matchBody.Value
                                                                           matchCount += 1
                                                                       End If
                                                                   End If
                                                                   If matchString.Any Then .Matches.Add(matchString)
                                                               End If
                                                           End With
                                                       End Sub)
                                       Dim hasMatches As Boolean = matchCount > 0 And matchCount >= Filters.Count
                                       If hasMatches Then RaiseEvent Found(Me, New SnifferEventArgs(Me, e, False, False))
                                   End If
                               End If
                           End If
                       End Function).ConfigureAwait(False)
        Await Task.Run(Sub()
                           RaiseEvent ResponseAlert(Me, New SnifferEventArgs(Me, e, False))
                           ProxyEvent(e, False)
                       End Sub).ConfigureAwait(False)

    End Function
    Private Async Function Proxy_AfterResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.AfterResponse

        RaiseEvent ResponseAlert(Me, New SnifferEventArgs(Me, e, False))
        Await Task.Run(Sub()
                           Dim requestBody As Task(Of String) = e.GetResponseBodyAsString()
                           If requestBody IsNot Nothing Then
                           End If
                       End Sub).ConfigureAwait(False)
    End Function
    Private Sub ProxyEvent(e As SessionEventArgs, isRequest As Boolean)

        If Filters.Any Then
            RaiseEvent Alert(e.HttpClient.Request, New AlertEventArgs(e.HttpClient.Request.Url.ToUpperInvariant))
            Dim matchCount As Integer = 0
            Dim filterCount As Integer = 0
            Filters.ForEach(Sub(fltr)
                                Dim matchString As String = String.Empty
                                With fltr
                                    If .Active Then
                                        filterCount += 1
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
                                    End If
                                End With
                            End Sub)
            Dim hasMatches As Boolean = matchCount > 0 And matchCount >= filterCount
            If hasMatches Then RaiseEvent Found(Me, New SnifferEventArgs(Me, e, isRequest, False))
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
    Payload
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
    Public Property Active As Boolean = True
    Public Overrides Function ToString() As String
        Return $"{What} => {Where} [{Split(How.ToString, ".").First}]"
    End Function
End Class