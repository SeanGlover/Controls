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
Public NotInheritable Class IconTimer
    Inherits Windows.Forms.Timer
    Private ReadOnly RunForm As Windows.Forms.Form
    Public Sub New()
    End Sub
    Public Sub New(updateForm As Windows.Forms.Form)
        RunForm = updateForm
    End Sub
    Public ReadOnly Property Icons As New List(Of Icon)
    Public Property Name As String
    Public Property Counter As Integer
    Public Property Flag As Boolean
    Private Sub Ticked(sender As Object, e As EventArgs) Handles Me.Tick

        Counter += 1
        If RunForm IsNot Nothing Then SetSafeControlPropertyValue(RunForm, "Icon", TickIcon)

    End Sub
    Public ReadOnly Property TickIcon As Icon
        Get
            Return If(Icons.Any, Icons(Counter Mod Icons.Count), Nothing)
        End Get
    End Property
End Class