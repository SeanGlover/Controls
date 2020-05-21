Option Strict On
Option Explicit On

Imports System.Net
Imports Titanium.Web.Proxy
Imports Titanium.Web.Proxy.EventArguments
Imports Titanium.Web.Proxy.Models
Imports System.Text.RegularExpressions

Public Class SnifferEventArgs
    Inherits EventArgs
    Public ReadOnly Property Headers As List(Of KeyValuePair(Of String, String))
    Public ReadOnly Property HeadersDistinct As New Dictionary(Of String, String)
    Public Sub New()
    End Sub
    Public Sub New(headers As List(Of KeyValuePair(Of String, String)))

        Me.Headers = headers
        If headers IsNot Nothing Then
            For Each header In headers
                If Not HeadersDistinct.ContainsKey(header.Key) Then HeadersDistinct.Add(header.Key, header.Value)
            Next
        End If

    End Sub
End Class
Public Class Sniffer
    Public Event Alert(sender As Object, e As Controls.AlertEventArgs)
    Public Event Found(sender As Object, e As SnifferEventArgs)

    Private WithEvents ProxyServer As ProxyServer
    Private Const ProxyPort As Integer = 18880
    Private ReadOnly Property SniffRequests As Integer
    Public Property Name As String
    Public Property Tag As Object
    Public ReadOnly Property Clients As New List(Of Http.HttpWebClient)
    Public Property Search As New SearchRequest
    Public ReadOnly Property ClientsString As New List(Of String)
    Public ReadOnly Property Sniffing As Boolean = False
    Public Sub New()
    End Sub

    Public Sub StartSniffing()

        Clients.Clear()
        ClientsString.Clear()
        If Not Sniffing Then
            ProxyServer = New ProxyServer
            Dim explicitEndPoint As New ExplicitProxyEndPoint(IPAddress.Any, ProxyPort + SniffRequests, True)
            With ProxyServer
                .AddEndPoint(explicitEndPoint)
                .Start()
                .SetAsSystemHttpProxy(explicitEndPoint)
                .SetAsSystemHttpsProxy(explicitEndPoint)
            End With
            _Sniffing = True
            _SniffRequests += 1
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
                           RequestEvent(e)
                       End Sub).ConfigureAwait(False)
    End Function
    Private Async Function Proxy_AfterRequest(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.BeforeResponse
        Await Task.Run(Sub()
                           RequestEvent(e)
                       End Sub).ConfigureAwait(False)
    End Function
    Private Async Function Proxy_AfterResponse(Sender As Object, e As SessionEventArgs) As Task Handles ProxyServer.AfterResponse
        Await Task.Run(Sub()
                           Dim response As Http.Response = e.HttpClient.Response
                           If Not response.StatusCode = 0 Then
                               RaiseEvent Alert(response, New AlertEventArgs(CStr(response.StatusCode)))
                               If response.HasBody And 0 = 1 Then '/// Errors trying to read BodyString as Response is still processing the reading of the Body
                                   Dim responseValues As New List(Of KeyValuePair(Of String, String))
                                   With Search
                                       If .By = FindyBy.Body Then
                                           If .Expression Is Nothing Then 'Explicit
                                               For Each searchValue In .Values
                                                   If response.BodyString.Contains(searchValue) Then
                                                       responseValues.Add(New KeyValuePair(Of String, String)(searchValue, response.BodyString))
                                                   End If
                                               Next
                                           Else 'Regex
                                               Dim bodyMatches = RegexMatches(response.BodyString, .Expression.SearchPattern, .Expression.SearchOptions)
                                               For Each bodyMatch In bodyMatches
                                                   responseValues.Add(New KeyValuePair(Of String, String)(bodyMatch.Value, response.BodyString))
                                               Next
                                           End If
                                       End If
                                   End With
                                   If responseValues.Any Then
                                       RaiseEvent Found(Me, New SnifferEventArgs(responseValues))
                                   End If
                               End If
                               ClientsString.Add(ClientToString(e.HttpClient))
                           End If
                       End Sub).ConfigureAwait(False)
    End Function
    Private Sub RequestEvent(e As SessionEventArgs)

        Clients.Add(e.HttpClient)
        Dim request As Http.Request = e.HttpClient.Request
        RaiseEvent Alert(request, New AlertEventArgs(request.Url.ToUpperInvariant))
        With Search
            'GET HEADER(S) WHEN ...
            Dim requestHeaders As New List(Of KeyValuePair(Of String, String))
            If .By = FindyBy.Host And .Host IsNot Nothing Then
#Region " ... A GIVEN HOST MATCHES - AND - HAS A PARTICULAR REQUEST HEADER NAME ( ex x-authorize-token ) "
                Dim hostMatches As Boolean = If(.Expression Is Nothing, .Host.ToUpperInvariant = request.Host.ToUpperInvariant, Regex.Match(request.Host, .Expression.SearchPattern, .Expression.SearchOptions).Success)
                If hostMatches Then
                    For Each requestHeader As HttpHeader In request.Headers
                        If .UpperValues.Contains(requestHeader.Name.ToUpperInvariant) Then requestHeaders.Add(New KeyValuePair(Of String, String)(requestHeader.Name, requestHeader.Value))
                    Next
                End If
#End Region
            ElseIf .By = FindyBy.RequestURL Then
#Region " ... A GIVEN REQUEST URL MATCHES ( ex JSESSIONID ) "
                Dim lookForURL As String = .RequestURL?.ToString.ToUpperInvariant
                Dim requestURL As String = request.RequestUri.ToString.ToUpperInvariant()
                Dim urlMatches As Boolean = If(.Expression Is Nothing, lookForURL = requestURL, Regex.Match(request.RequestUri.ToString, .Expression.SearchPattern, .Expression.SearchOptions).Success)
                If urlMatches Then
                    For Each requestHeader As HttpHeader In request.Headers
                        requestHeaders.Add(New KeyValuePair(Of String, String)(requestHeader.Name, requestHeader.Value))
                    Next
                End If
#End Region
            ElseIf .By = FindyBy.None Then

            End If
            If requestHeaders.Any Then RaiseEvent Found(Me, New SnifferEventArgs(requestHeaders))
        End With

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
End Class
Public Enum FindyBy
    None
    RequestURL
    Host
    Body
End Enum
Public NotInheritable Class ByPattern
    Implements IEquatable(Of ByPattern)
    Public Property SearchPattern As String
    Public Property SearchOptions As RegexOptions = RegexOptions.None
    Public Overrides Function GetHashCode() As Integer
        Return SearchPattern.GetHashCode Xor SearchOptions.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As ByPattern) As Boolean Implements IEquatable(Of ByPattern).Equals
        If other Is Nothing Then
            Return Me Is Nothing
        Else
            Return SearchPattern = other.SearchPattern AndAlso SearchOptions = other.SearchOptions
        End If
    End Function
    Public Shared Operator =(ByVal value1 As ByPattern, ByVal value2 As ByPattern) As Boolean
        If value1 Is Nothing Then
            Return value2 Is Nothing
        ElseIf value2 Is Nothing Then
            Return value1 Is Nothing
        Else
            Return value1.Equals(value2)
        End If
    End Operator
    Public Shared Operator <>(ByVal value1 As ByPattern, ByVal value2 As ByPattern) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is ByPattern Then
            Return CType(obj, ByPattern) = Me
        Else
            Return False
        End If
    End Function
End Class
Public NotInheritable Class SearchRequest
    Public Sub New()
    End Sub
    Public Sub New(searchValue As String, searchBy As FindyBy)
        Values.Add(searchValue)
        By = searchBy
    End Sub
    Public Property Host As String
    Public Property RequestURL As Uri
    Public ReadOnly Property Values As New List(Of String)
    Public ReadOnly Property UpperValues As List(Of String)
        Get
            Return (From v In Values Select v.ToUpperInvariant).ToList
        End Get
    End Property
    Public Property Expression As ByPattern
    Private By_ As FindyBy = FindyBy.None
    Public Property By As FindyBy
        Get
            If By_ = FindyBy.None Then
                If Host IsNot Nothing Then By_ = FindyBy.Host
                If RequestURL IsNot Nothing Then By_ = FindyBy.RequestURL
            End If
            Return By_
        End Get
        Set(value As FindyBy)
            By_ = value
        End Set
    End Property
End Class