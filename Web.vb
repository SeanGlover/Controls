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
                    '    .Referer = lockboxURL.ToString
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
Public NotInheritable Class Token
    Public Event Expired(sender As Object, e As TokenEventArgs)
    Public Event Expiring(sender As Object, e As TokenEventArgs)
    Private WithEvents ExpiryTimer As New Timer With {.Interval = 1000}
    Public Sub New()
    End Sub
    Public Sub New(tokenString As String)

        Dim tokenElements() As String = Split(tokenString, Delimiter)
        If tokenElements.Length = 3 Then
            Name = tokenElements.First
            Value = tokenElements(1)
            Expiry = StringToDateTime(tokenElements.Last)
        End If

    End Sub
    Public Property Name As String
    Public Property Value As String
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
            Return Now < Expiry
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
    Public Overrides Function ToString() As String
        Return Join({Name, Value, DateTimeToString(Expiry)}, Delimiter)
    End Function
End Class