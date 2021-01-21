Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Public NotInheritable Class Login
    Inherits Control
    Public Enum Display
        Horizontal
        Vertical
    End Enum
    Public Event ValuesSubmitted(ByVal sender As Object, ByVal e As LoginEventArgs)
    Public ReadOnly Property UID As New ImageCombo With {.Image = My.Resources.UID, .HighlightOnFocus = True, .BorderStyle = Border3DStyle.Flat, .HintText = "User ID", .WrapText = False, .Name = "UID"}
    Public ReadOnly Property PWD As New ImageCombo With {.Image = My.Resources.Password, .HighlightOnFocus = True, .BorderStyle = Border3DStyle.Flat, .HintText = "Password", .WrapText = False, .Name = "PWD", .PasswordProtected = True}
    Private MyCredentials As New Credentials
    Public Sub New()

        Size = New Size(100, 60)
        HidePassword_ = True
        Controls.AddRange({UID, PWD})
        AddHandler UID.ValueSubmitted, AddressOf ValueSubmitted
        AddHandler PWD.ValueSubmitted, AddressOf ValueSubmitted

    End Sub
    Protected Overrides Sub InitLayout()

    End Sub
    Private Sub ResizeMe()

        Dim UnitHeight As Int32 = 0, UnitWidth As Int32 = 0

        If Orientation = Display.Vertical Then
            UnitHeight = CInt((Height - SeparationPadding) / 2)
            UnitWidth = Width
            With UID
                .Location = New Point(0, 0)
                .Size = New Size(UnitWidth, UnitHeight)
            End With
            With PWD
                .Left = 0
                .Top = Height - UnitHeight
                .Size = New Size(UnitWidth, UnitHeight)
            End With

        ElseIf Orientation = Display.Horizontal Then
            UnitHeight = Height
            UnitWidth = CInt((Width - SeparationPadding) / 2)
            With UID
                .Location = New Point(0, 0)
                .Size = New Size(UnitWidth, UnitHeight)
            End With
            With PWD
                .Location = New Point(Width - SeparationPadding - UID.Width, 0)
                .Size = New Size(UnitWidth, UnitHeight)
            End With

        End If

    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)

        ResizeMe()
        MyBase.OnSizeChanged(e)

    End Sub
    Private mUserId As String
    Public Property UserId As String
        Get
            Return mUserId
        End Get
        Set(value As String)
            mUserId = value
            UID.Text = value
        End Set
    End Property
    Private mPassWord As String
    Public Property PassWord As String
        Get
            Return mPassWord
        End Get
        Set(value As String)
            mPassWord = value
            PWD.Text = value
        End Set
    End Property
    Private HidePassword_ As Boolean
    Public Property HidePassWord As Boolean
        Get
            Return HidePassword_
        End Get
        Set(value As Boolean)
            If HidePassword_ <> value Then
                HidePassword_ = value
                PWD.PasswordProtected = value
            End If
        End Set
    End Property
    Private mOrientation As Display = Display.Vertical
    Public Property Orientation As Display
        Get
            Return mOrientation
        End Get
        Set(value As Display)
            mOrientation = value
            ResizeMe()
        End Set
    End Property
    Private mSeparationPadding As Int32 = 2
    Public Property SeparationPadding As Int32
        Get
            Return mSeparationPadding
        End Get
        Set(value As Int32)
            mSeparationPadding = value
            ResizeMe()
        End Set
    End Property
    Private Sub ValueSubmitted(ByVal sender As Object, e As EventArgs)

        With MyCredentials
            .UserId = UID.Text
            UserId = UID.Text
            .Password = PWD.Text
            PassWord = PWD.Text
            RaiseEvent ValuesSubmitted(Me, New LoginEventArgs(MyCredentials))
        End With

    End Sub
End Class
Public Class LoginEventArgs
    Inherits EventArgs
    Public ReadOnly Property Credentials As Credentials
    Public Sub New(Credentials As Credentials)
        Me.Credentials = Credentials
    End Sub
End Class
Public Structure Credentials
    Implements IEquatable(Of Credentials)
    Public Property UserId As String
    Public Property Password As String
    Public Sub New(UID As String, PWD As String)
        UserId = UID
        Password = PWD
    End Sub
    Public Overrides Function GetHashCode() As Integer
        Return UserId.GetHashCode Xor Password.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As Credentials) As Boolean Implements IEquatable(Of Credentials).Equals
        Return UserId = other.UserId AndAlso Password = other.Password
    End Function
    Public Shared Operator =(ByVal value1 As Credentials, ByVal value2 As Credentials) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(ByVal value1 As Credentials, ByVal value2 As Credentials) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is Credentials Then
            Return CType(obj, Credentials) = Me
        Else
            Return False
        End If
    End Function
End Structure