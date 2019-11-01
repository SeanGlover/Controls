Option Strict On
Option Explicit On
Imports Microsoft.Data.Sqlite
#Region " COOKIES "
Public Class CookieCollection
    Inherits List(Of Cookie)
    Public ReadOnly Property CookiePath As String
        Get
            Return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\Google\Chrome\User Data\Default\Cookies"
        End Get
    End Property
    Friend ReadOnly Property Properties As Dictionary(Of String, Integer)
    Public Sub New(Optional Browser As String = "Chrome")

        If Browser = "Chrome" Then
            If CookiePath.Any Then
                Dim RowIndex As Integer
                Using conn As New SqliteConnection("Data Source=" + CookiePath)
                    conn.Open()
                    Using cmd As New SqliteCommand("SELECT * FROM Cookies", conn)
                        Dim reader As SqliteDataReader = cmd.ExecuteReader
                        While reader.Read
                            If RowIndex = 0 Then Properties = Enumerable.Range(0, reader.FieldCount - 1).ToDictionary(Function(c) reader.GetName(c), Function(i) i)
                            Dim Cookie As Cookie = New Cookie(Me, reader)
                            Add(Cookie)
                            RowIndex += 1
                        End While
                        conn.Close()
                    End Using
                End Using
                Sort(Function(f1, f2)
                         Dim Level1 = String.Compare(f1.host_key, f2.host_key, StringComparison.InvariantCulture)
                         If Level1 <> 0 Then
                             Return Level1
                         Else
                             Dim Level2 = String.Compare(f1.name, f2.name, StringComparison.InvariantCulture)
                             Return Level2
                         End If
                     End Function)
            End If
        Else

        End If

    End Sub
    Public Shadows Function Item(Host As String, Name As String, Optional Path As String = Nothing) As Cookie

        If Host IsNot Nothing And Name IsNot Nothing Then
            Dim Cookies = Where(Function(c) c.host_key = Host And c.name = Name)
            If Cookies.Any Then
                If Path Is Nothing Then
                    Return Cookies.First
                Else
                    Dim CookiesWithPath = Cookies.Where(Function(c) c.path = Path)
                    If CookiesWithPath.Any Then
                        Return CookiesWithPath.First
                    Else
                        Return Nothing
                    End If
                End If
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function
    Public Function Items(Host As String) As List(Of Cookie)

        If Host IsNot Nothing Then
            Dim Cookies As New List(Of Cookie)(Where(Function(c) c.host_key = Host))
            If Cookies.Any Then
                Return Cookies
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function
End Class
Public Structure Cookie
    Implements IEquatable(Of Cookie)
    Friend ReadOnly Property Parent As CookieCollection
    Friend Sub New(cookies As CookieCollection, reader As SqliteDataReader)

        Parent = cookies
        Dim Fields As New List(Of String)(Enumerable.Range(0, reader.FieldCount - 1).Select(Function(c) reader.GetString(c)))
        creation_utc = Long.Parse(Fields.First, InvariantCulture)
        host_key = Fields(Parent.Properties("host_key"))
        name = Fields(Parent.Properties("name"))
        path = Fields(Parent.Properties("path"))
        expires_utc = Long.Parse(Fields(Parent.Properties("expires_utc")), InvariantCulture)
        is_secure = Integer.Parse(Fields(Parent.Properties("is_secure")), InvariantCulture) = 1
        is_httponly = Integer.Parse(Fields(Parent.Properties("is_httponly")), InvariantCulture) = 1
        last_access_utc = Long.Parse(Fields(Parent.Properties("last_access_utc")), InvariantCulture)
        has_expires = Integer.Parse(Fields(Parent.Properties("has_expires")), InvariantCulture) = 1
        is_persistent = Integer.Parse(Fields(Parent.Properties("is_persistent")), InvariantCulture) = 1
        priority = Integer.Parse(Fields(Parent.Properties("priority")), InvariantCulture) = 1

        Dim EncryptedColumn As Integer = Parent.Properties("encrypted_value")
        Dim ByteStream As IO.Stream = reader.GetStream(EncryptedColumn)
        Dim Bytes As New List(Of Byte)
        For b = 0 To CInt(ByteStream.Length) - 1
            Bytes.Add(CByte(ByteStream.ReadByte()))
        Next
        encrypted_value = Bytes
        Dim decodedData = Security.Cryptography.ProtectedData.Unprotect(Bytes.ToArray, Nothing, Security.Cryptography.DataProtectionScope.CurrentUser)
        value = Text.Encoding.ASCII.GetString(decodedData)

    End Sub
    Public Overrides Function GetHashCode() As Integer
        Return host_key.GetHashCode Xor name.GetHashCode Xor path.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As Cookie) As Boolean Implements IEquatable(Of Cookie).Equals
        Return host_key = other.host_key AndAlso name = other.name AndAlso path = other.path
    End Function
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is Cookie Then
            Return CType(obj, Cookie) = Me
        Else
            Return False
        End If
    End Function
    Public Shared Operator =(ByVal value1 As Cookie, ByVal value2 As Cookie) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(ByVal value1 As Cookie, ByVal value2 As Cookie) As Boolean
        Return Not value1 = value2
    End Operator
    Friend ReadOnly creation_utc As Long
    Friend ReadOnly host_key As String
    Friend ReadOnly name As String
    Friend ReadOnly value As String
    Friend ReadOnly path As String
    Friend ReadOnly expires_utc As Long
    Friend ReadOnly is_secure As Boolean
    Friend ReadOnly is_httponly As Boolean
    Friend ReadOnly last_access_utc As Long
    Friend ReadOnly has_expires As Boolean
    Friend ReadOnly is_persistent As Boolean
    Friend ReadOnly priority As Boolean
    Friend ReadOnly encrypted_value As List(Of Byte)
    Public Overrides Function ToString() As String
        Return Join({name, value}, BlackOut)
    End Function
End Structure
#End Region