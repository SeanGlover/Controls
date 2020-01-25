Option Explicit On
Option Strict On
Public NotInheritable Class Worker
    Inherits ComponentModel.BackgroundWorker
    Public Sub New()
    End Sub
    Public Property Tag As Object
    Public Property Name As String
End Class