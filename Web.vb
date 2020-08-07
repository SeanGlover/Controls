Option Explicit On
Option Strict On
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO

Public NotInheritable Class WebFunctions
    Private Const ErrorFlag As String = "≠"
    Public Shared Sub SetPayload(parameters As Dictionary(Of String, Object), request As HttpWebRequest)

        If parameters IsNot Nothing And request IsNot Nothing Then
            Dim body As String = Join((From eih In parameters Select Join({eih.Key, eih.Value}, "=")).ToArray, "&")
            Dim postBytes() As Byte = Text.Encoding.UTF8.GetBytes(body)
            With request
                .ContentLength = postBytes.Length
                Using RequestStream As Stream = .GetRequestStream()
                    RequestStream.Write(postBytes, 0, postBytes.Length)
                End Using
            End With
        End If

    End Sub
    Public Shared Function GetResponse(request As HttpWebRequest) As KeyValuePair(Of HttpWebResponse, String)

        If request Is Nothing Then
            Return Nothing
        Else
            Dim response As HttpWebResponse = Nothing
            Dim responseText As String
            Try
                response = DirectCast(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(response.GetResponseStream(), Text.Encoding.ASCII)
                    responseText = reader.ReadToEnd()
                End Using

            Catch ex As WebException
                responseText = ErrorFlag & ex.Message

            End Try
            Return New KeyValuePair(Of HttpWebResponse, String)(response, responseText)
        End If

    End Function
    Public Shared Function HarToCode(path As String) As String

        If path Is Nothing Then
            Return Nothing
        Else
            Dim content As String = ReadText(path)
            If content?.Any Then
                Dim requestMatch As Match = Regex.Match(content, "(?<=""request"": ){", RegexOptions.None)
                If requestMatch.Success Then
                    Dim methodString As String = Regex.Match(content, "(?<=""method"": )""[^""]{1,}""", RegexOptions.None).Value
                    Dim url As String = Regex.Match(content, "(?<=""url"": )""[^""]{1,}""", RegexOptions.None).Value
                    Dim codeList As New List(Of String) From {"Dim harRequest=DirectCast(WebRequest.Create(" & url & "), HttpWebRequest)",
                        "With harRequest"}
                    codeList.Add(".Method=" & methodString)
                    Dim requestString As String = content.Substring(requestMatch.Index, content.Length - requestMatch.Index)
                    Dim leftBrackets As New List(Of Integer)
                    Dim rightBrackets As New List(Of Integer)
                    Dim letterIndex As Integer = requestMatch.Index
                    For Each letter As Char In requestString
                        If letter = "{" Then leftBrackets.Add(letterIndex)
                        If letter = "}" Then rightBrackets.Add(letterIndex)
                        letterIndex += 1
                        If leftBrackets.Count = rightBrackets.Count Then Exit For
                    Next
                    Dim newString = content.Substring(leftBrackets.Min, 1 + rightBrackets.Max - leftBrackets.Min)

                    Dim headersCookiesParams = RegexMatches(newString, """[a-z]{1,}"": \[[^]]{1,}", RegexOptions.Multiline)
                    Dim headers As New List(Of Match)(From hcp In headersCookiesParams Where hcp.Value.StartsWith("""headers"": [", StringComparison.InvariantCulture))
                    If headers.Any Then
                        Dim headerString As String = headers.First.Value
                        Dim namesValues = RegexMatches(headerString, """[a-zA-Z]{2,}"": ""{0,1}[^\n]{2,}""{0,1},{0,1}""", RegexOptions.Multiline)
                        Dim headerDictionary As New Dictionary(Of String, String)
                        Dim key As String = Nothing
                        Dim value As String = Nothing
                        For Each nameValue In namesValues
                            If nameValue.Value.Contains("name") Then key = Replace(Split(nameValue.Value, ": ").Last, """", "")
                            If nameValue.Value.Contains("value") Then
                                value = Replace(Split(nameValue.Value, ": ").Last, """", "")
                                headerDictionary.Add(key, value)
                                key = Nothing
                                value = Nothing
                            End If
                        Next
                        For Each requestElement In headerDictionary
                            key = requestElement.Key
                            value = requestElement.Value
                            If {"HOST", "ACCEPT", "REFERER", "USER-AGENT", "CONTENT-TYPE"}.Contains(key.ToUpperInvariant) Then
                                codeList.Add("." & Replace(key, "-", String.Empty) & "=" & """" & value & """")

                            ElseIf key.ToUpperInvariant = "CONNECTIUON" Then
                                codeList.Add(".KeepAlive=" & "" & If(value.ToUpperInvariant = "KEEP-ALIVE", "True", "False") & "")

                            ElseIf key.ToUpperInvariant = "CONTENT-LENGTH" Then
                                'Do nothing .Content-Length is set with SetPayload

                            ElseIf key.ToUpperInvariant = "COOKIE" Then
                                codeList.Add(".Headers.Set(HttpRequestHeader.Cookie, CookieMonster.JsessionID)")

                            Else
                                codeList.Add(".Headers.Add(""" & key & """, """ & value & """)")

                            End If
                        Next
                    End If
                    Dim cookies As New List(Of Match)(From hcp In headersCookiesParams Where hcp.Value.StartsWith("""cookies"": [", StringComparison.OrdinalIgnoreCase))
                    If cookies.Any Then
                        Dim cookieString As String = cookies.First.Value
                        Dim namesValues = RegexMatches(cookieString, """[a-zA-Z]{2,}"": ""{0,1}[^\n]{2,}""{0,1},{0,1}""", RegexOptions.Multiline)
                        'Stop
                    End If
                    Dim params As New List(Of Match)(From hcp In headersCookiesParams Where hcp.Value.StartsWith("""params"": [", StringComparison.OrdinalIgnoreCase))
                    If params.Any Then
                        Dim paramString As String = params.First.Value
                        Dim namesValues = RegexMatches(paramString, """[a-zA-Z]{2,}"": ""{0,1}[^\n]{2,}""{0,1},{0,1}""", RegexOptions.Multiline)
                        Dim paramDictionary As New Dictionary(Of String, String)
                        Dim key As String = Nothing
                        Dim value As String = Nothing
                        For Each nameValue In namesValues
                            If nameValue.Value.Contains("name") Then key = Split(nameValue.Value, ": ").Last
                            If Not paramDictionary.ContainsKey(key) And nameValue.Value.Contains("value") Then
                                value = Split(nameValue.Value, ": ").Last
                                paramDictionary.Add(key, value)
                                key = Nothing
                                value = Nothing
                            End If
                        Next
                        Dim dictionaryString As String = Join((From pd In paramDictionary Select "{" & Join({pd.Key, pd.Value}, ", ") & "}").ToArray, "," & vbNewLine)
                        codeList.Add("Dim paramDictionary as New Dictionary(Of String, Object) From {" & dictionaryString & "}")
                        codeList.Add("SetPayload(paramDictionary, harRequest)")
                        codeList.Add("Dim responseKVP = GetResponse(harRequest)")
                        codeList.Add("If responseKVP.Value.StartsWith(ErrorFlag) Then")
                        codeList.Add("Stop")
                        codeList.Add("Else")
                        codeList.Add("End If")
                    End If
                    codeList.Add("End With")
                    Dim codeString As String = Join(codeList.ToArray, vbNewLine)
                    Clipboard.SetText(codeString)
                    Return codeString
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Shared Function ChromeRequestToCode() As String

        Dim clipPath As String = Desktop & "\chromeRequest.txt"
        Dim chromeRequest As String = Clipboard.GetText
        Using sw As New StreamWriter(clipPath)
            sw.Write(chromeRequest)
        End Using
        Return ChromeRequestToCode(clipPath)

    End Function
    Public Shared Function ChromeRequestToCode(requestPath As String) As String

        If requestPath Is Nothing Then
            Return Nothing
        Else
            If File.Exists(requestPath) Then
                Dim chromeRequest As String = ReadText(requestPath)
                Dim urlMatch As Match = Regex.Match(chromeRequest, "(?<=Request URL: ).*", RegexOptions.None)
                If urlMatch.Success Then
                    Dim codeList As New List(Of String) From {"Dim request = CType(WebRequest.Create(""" & TrimReturn(urlMatch.Value) & """), HttpWebRequest)"}
                    codeList.Add("With request")
                    Dim typeMatch = Regex.Match(chromeRequest, "(?<=Request Method: ).*", RegexOptions.None)
                    codeList.Add(".Method = " & typeMatch.Value)

                    codeList.Add("End With")
                    Dim RequestURL As String = urlMatch.Value

                    Return codeList.First

                    'Dim request = CType(WebRequest.Create(requestUri), HttpWebRequest)
                    'With request{urlMatch.Value
                    '    .Method = "GET"
                    '    .Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
                    '    .Headers.Set(HttpRequestHeader.AcceptEncoding, "gzip, deflate, br")
                    '    .Headers.Set(HttpRequestHeader.AcceptLanguage, "fr,es;q=0.9,en;q=0.8")
                    '    .KeepAlive = True
                    '    .Headers.Set(HttpRequestHeader.Cookie, Cookie)
                    '    .Host = "www.treasury.pncbank.com"
                    '    .Referer = TokenURL.ToString
                    '    .Headers.Add("Sec-Fetch-Dest", "iframe")
                    '    .Headers.Add("Sec-Fetch-Mode", "navigate")
                    '    .Headers.Add("Sec-Fetch-Site", "same-origin")
                    '    .Headers.Add("Sec-Fetch-User", "?1")
                    '    .Headers.Add("Upgrade-Insecure-Requests", "1")
                    '    .UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36"
                    '    .ServicePoint.Expect100Continue = False

                    '    Dim response = CType(request.GetResponse, HttpWebResponse)
                    '    Using sr As StreamReader = New StreamReader(response.GetResponseStream)
                    '        responseText = sr.ReadToEnd
                    '    End Using
                    '    Stop
                    'End With

                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End If

    End Function
End Class
Public NotInheritable Class TokenEventArgs
    Public ReadOnly Property Token As Token
    Public Sub New(eventToken As Token)
        Token = eventToken
    End Sub
End Class
Public Enum SameSiteValue
    None
    Strict
    Lax
End Enum
Public NotInheritable Class Token
    Implements IEquatable(Of Token)

    Public Event Expired(sender As Object, e As TokenEventArgs)
    Public Event Expiring(sender As Object, e As TokenEventArgs)
    Private WithEvents ExpiryTimer As New Timer With {.Interval = 1000}
    Public Sub New()
    End Sub
    Public Sub New(tokenString As String)

        tokenString = If(tokenString, String.Empty)
        CookieString = tokenString

        'https://developer.mozilla.org/en-US/docs/Web/HTTP/Headers/Set-Cookie
        'A <cookie-name> can be any US-ASCII characters, except control characters, spaces, or tabs. It also must not contain a separator character like the following: ( ) < > @ , ; : \ " / [ ] ? = { }.
        'A <cookie-value> can optionally be wrapped in double quotes and include any US-ASCII characters excluding control characters, Whitespace, double quotes, comma, semicolon, and backslash. Encoding: Many implementations perform URL encoding on cookie values, however it is not required per the RFC specification. It does help satisfying the requirements about which characters are allowed for <cookie-value> though.

        'Set-Cookie: <cookie-name>=<cookie-value> 
        'Set-Cookie: <cookie-name>=<cookie-value>; Expires= <date>
        'Set-Cookie: <cookie-name>=<cookie-value>; Max-Age=<non-zero-digit>
        'Set-Cookie: <cookie-name>=<cookie-value>; Domain=<domain-value>
        'Set-Cookie: <cookie-name>=<cookie-value>; Path=<path-value>
        'Set-Cookie: <cookie-name>=<cookie-value>; Secure
        'Set-Cookie: <cookie-name>=<cookie-value>; HttpOnly

        'Set-Cookie: <cookie-name>=<cookie-value>; SameSite=Strict
        'Set-Cookie: <cookie-name>=<cookie-value>; SameSite=Lax
        'Set-Cookie: <cookie-name>=<cookie-value>; SameSite=None

        '// Multiple attributes are also possible, for example:
        'Set-Cookie: <cookie-name>=<cookie-value>; Domain=<domain-value>; Secure; HttpOnly

        Dim tokenElements As New List(Of String)(Split(tokenString, BlackOut))
        If tokenElements.Count = 3 Then
            Name = tokenElements.First
            Value = tokenElements(1)
            Expiry = StringToDateTime(tokenElements.Last)
        Else
            If tokenString.Any Then
                If Regex.Match(tokenString, "Expires|Max-Age|Domain|Path|Secure|HttpOnly|SameSite", RegexOptions.IgnoreCase).Success Then
                    '_abck=B914B596F4561731D831D13AF5CD8315~-1~YAAQJO/dF4ngLsBzAQAA8YGnwwTKCb34pA93llQL6ZTWjQjhGcR42IWTKHOor4n2mil9aYjOgk3/Cxcb/8YmCm8LlYK8jAthyOrlsGhNtv66Eh1UFuEk8x6vdNRUlq3jhQh/6MsSNwauvgXNths+gnOo07uZXZuT1mJYsaLK1HmBXm33AxeFJyl/ZT3ccc2fO0UI0IADQre9YmYycsJCHX6HT1a8rDGn87PJfrVFh6qGDWa1V3bPEFVYs5+lCLl+9N6kt6GE5Mmf8vApQWsc7SIykLmRzgsl7Giizbuf1e1uswCxwMDv+nUD/5cLGcGry/5xqMUAEho=~0~-1~-1
                    '; Domain=.pncbank.com
                    '; Path=/
                    '; Expires=Fri, 06 Aug 2021 12:03:21 GMT
                    '; Max-Age=31536000
                    '; SecureStrict-Transport-Security: max-age=31536000
                    tokenElements = Regex.Split(tokenString, "; {0,1}", RegexOptions.IgnoreCase).ToList

                    '/// First is always Cookie.Name/Value
                    Dim nameValue As String() = Split(tokenElements(0), "=")
                    Name = Trim(nameValue.First) 'Cookie name can not have a space
                    Value = Regex.Replace(nameValue.Last, """", String.Empty)
                    tokenElements.RemoveAt(0)

                    '/// Iterate the remaining values ( if any )
                    tokenElements.ForEach(Sub(te)
                                              Dim domainMatch = Regex.Match(te, "domain=", RegexOptions.IgnoreCase)
                                              If domainMatch.Success Then _Domain = Trim(Split(domainMatch.Value, "=").Last)

                                              Dim pathMatch = Regex.Match(te, "path=", RegexOptions.IgnoreCase)
                                              If pathMatch.Success Then _Path = Trim(Split(pathMatch.Value, "=").Last)

                                              Dim expiresMatch = Regex.Match(te, "(?<=expires=)[^;]{1,}", RegexOptions.IgnoreCase)
                                              If expiresMatch.Success Then
                                                  Dim parsedDate = Date.Parse(expiresMatch.Value, New Globalization.CultureInfo("en-US"))
                                                  Expiry_ = parsedDate
                                              End If

                                              Dim secureMatch = Regex.Match(te, "(?<=secure)", RegexOptions.IgnoreCase)
                                              If secureMatch.Success Then _Secure = True

                                              Dim maxAgeMatch = Regex.Match(te, "(?<=Max-Age=)[0-9]{1,}", RegexOptions.IgnoreCase)
                                              If maxAgeMatch.Success Then _MaxAge = CLng(maxAgeMatch.Value)

                                              Dim httpOnlyMatch = Regex.Match(te, "httponly", RegexOptions.IgnoreCase)
                                              If httpOnlyMatch.Success Then _HttpOnly = True

                                          End Sub)

                Else
                    '???
                End If
            Else
                ExpiryTimer.Start()
            End If
        End If

    End Sub
    Public ReadOnly Property CookieString As String
    Public Property Name As String
    Public Property Value As String
    Public ReadOnly Property MaxAge As Long 'Max-Age=<number> Optional, Number of seconds until the cookie expires. A zero or negative number will expire the cookie immediately. If both Expires and Max-Age are set, Max-Age has precedence
    Public ReadOnly Property Domain As String 'Domain=<domain-value> Optional, Host to which the cookie will be sent
    Public ReadOnly Property Path As String 'Path=<path-value> Optional, A path that must exist in the requested URL, or the browser won't send the Cookie header.
    Public ReadOnly Property Secure As Boolean  'Secure Optional, A secure cookie is only sent to the server when a request is made with the https: scheme
    Public ReadOnly Property HttpOnly As Boolean 'Forbids JavaScript from accessing the cookie, for example, through the Document.cookie property
    Public ReadOnly Property SameSite As SameSiteValue '(Strict|Lax|None) ... Can be follwed with: Strict-Transport-Security: max-age=10886400
    Private Expiry_ As Date
    Public Property Expiry As Date
        Get
            Return Expiry_
        End Get
        Set(value As Date)
            If Expiry_ <> value Then
                Expiry_ = value
                ExpiryTimer.Start()
            End If
        End Set
    End Property
    Public ReadOnly Property RemainingTime As TimeSpan
        Get
            If Valid Then
                Return Expiry.Subtract(Now)
            Else
                Return New TimeSpan()
            End If
        End Get
    End Property
    Public ReadOnly Property Valid As Boolean
        Get
            Return Now <Expiry
        End Get
    End Property
    Private Sub ExpiryTimer_Tick() Handles ExpiryTimer.Tick

        If Valid Then
            If RemainingTime.TotalSeconds < 60 Then RaiseEvent Expiring(Me, New TokenEventArgs(Me))
        Else
            ExpiryTimer.Stop()
            RaiseEvent Expired(Me, New TokenEventArgs(Me))
        End If

    End Sub

    Public Overrides Function GetHashCode() As Integer
        Return If(Name, String.Empty).GetHashCode Xor If(Value, String.Empty).GetHashCode Xor Expiry.GetHashCode Xor MaxAge.GetHashCode Xor If(Domain, String.Empty).GetHashCode Xor If(Path, String.Empty).GetHashCode Xor Secure.GetHashCode Xor HttpOnly.GetHashCode Xor SameSite.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As Token) As Boolean Implements IEquatable(Of Token).Equals
        Return other IsNot Nothing AndAlso (Name = other.Name And Value = other.Value And Expiry = other.Expiry)
    End Function
    Public Shared Operator =(ByVal Object1 As Token, ByVal Object2 As Token) As Boolean
        If Object1 Is Nothing Then
            Return Object2 Is Nothing
        ElseIf Object2 Is Nothing Then
            Return False
        Else
            Return Object1.Equals(Object2)
        End If
    End Operator
    Public Shared Operator <>(ByVal Object1 As Token, ByVal Object2 As Token) As Boolean
        Return Not Object1 = Object2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is Token Then
            Return CType(obj, Token) = Me
        Else
            Return False
        End If
    End Function

    Public Overrides Function ToString() As String
        Return Join({Name, Value, DateTimeToString(Expiry)}, BlackOut)
    End Function
End Class