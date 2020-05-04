Option Explicit On
Option Strict On
Public NotInheritable Class Worker
    Inherits ComponentModel.BackgroundWorker
    Public Sub New()
    End Sub
    Public Property Tag As Object
    Public Property Name As String
End Class
Public NotInheritable Class Ticker
    Inherits Windows.Forms.Timer
    Public Sub New()
    End Sub
    Public Property Name As String
    Public Property Counter As Integer
    Public Property Flag As Boolean
End Class