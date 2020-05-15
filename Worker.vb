Option Explicit On
Option Strict On
Imports System.Drawing
Imports System.Security
Imports SQLitePCL

Public NotInheritable Class Worker
    Inherits ComponentModel.BackgroundWorker
    Public Sub New()
    End Sub
    Public Property Tag As Object
    Public Property Name As String
End Class
Public NotInheritable Class Ticker
    Inherits Windows.Forms.Timer
    Private ReadOnly RunIcons As New Dictionary(Of Byte, Icon)
    Public Sub New()

        Dim icons = MyIcons()
        For Each runIcon As KeyValuePair(Of String, Icon) In MyIcons()
            If runIcon.Key.StartsWith("r", StringComparison.CurrentCulture) Then
                Dim runIndex As Byte = CByte(Replace(runIcon.Key, "r", String.Empty))
                RunIcons.Add(runIndex, runIcon.Value)
            End If
        Next

    End Sub
    Public Property Name As String
    Private Counter_ As Integer
    Public Property Counter As Integer
        Get
            Return Counter_
        End Get
        Set(value As Integer)
            If MaxCount >= 0 Then
                Counter_ = {MaxCount, value}.Min
            Else
                Counter_ = value
            End If
        End Set
    End Property
    Public Property MaxCount As Integer = -1
    Public Property Flag As Boolean
    Private Sub Ticked(sender As Object, e As EventArgs) Handles Me.Tick

        If Counter = MaxCount Then
            Counter_ = 0
        Else
            Counter_ += 1
        End If

    End Sub
    Public ReadOnly Property RunIcon As Icon
        Get
            Return RunIcons(CByte(1 + Counter Mod RunIcons.Count))
        End Get
    End Property
End Class