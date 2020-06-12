Option Explicit On
Option Strict On
Imports System.Drawing

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
    Private ReadOnly RunForm As Windows.Forms.Form
    Public Sub New()
        Initialize()
    End Sub
    Public Sub New(updateForm As Windows.Forms.Form)
        RunForm = updateForm
        Initialize()
    End Sub
    Private Sub Initialize()

        Dim icons = MyIcons()
        For Each runIcon As KeyValuePair(Of String, Icon) In MyIcons()
            If runIcon.Key.StartsWith("r", StringComparison.CurrentCulture) Then
                Dim runIndex As Byte = CByte(Replace(runIcon.Key, "r", String.Empty))
                RunIcons.Add(runIndex, runIcon.Value)
            End If
        Next

    End Sub
    Public Property Name As String
    Public Property Counter As Integer
    Public Property Flag As Boolean
    Private Sub Ticked(sender As Object, e As EventArgs) Handles Me.Tick

        Counter += 1
        If RunForm IsNot Nothing Then
            SetSafeControlPropertyValue(RunForm, "Icon", RunIcon)
            'SetSafeControlPropertyValue(RunForm, "Text", Counter.ToString)
        End If

    End Sub
    Public ReadOnly Property RunIcon As Icon
        Get
            Return RunIcons(CByte(1 + Counter Mod RunIcons.Count))
        End Get
    End Property
End Class