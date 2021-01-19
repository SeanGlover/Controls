Option Explicit On
Option Strict On
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Gma.System.MouseKeyHook
Public Class Hooker
    Implements IDisposable
#Region " DISPOSE "
    Dim disposed As Boolean = False
    ReadOnly Handle As SafeHandle = New Microsoft.Win32.SafeHandles.SafeFileHandle(IntPtr.Zero, True)
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposed Then Return
        If disposing Then
            Handle.Dispose()
            ' Free any other managed objects here.
            GlobalHook.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public ReadOnly Property Hooked As Boolean
    Public Property Tag As Object
    Public Event Keyed(sender As Object, e As KeyPressEventArgs)
    Public Event Moused(sender As Object, e As MouseEventExtArgs)
    Private GlobalHook As IKeyboardMouseEvents
    Private disposedValue As Boolean

    Public Sub New()
    End Sub
    Public Sub Subscribe()

        GlobalHook = Hook.GlobalEvents()
        AddHandler GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
        AddHandler GlobalHook.KeyPress, AddressOf GlobalHookKeyPress
        _Hooked = True

    End Sub
    Private Sub GlobalHookKeyPress(sender As Object, e As KeyPressEventArgs)
        RaiseEvent Keyed(Me, e)
    End Sub
    Private Sub GlobalHookMouseDownExt(sender As Object, e As MouseEventExtArgs)
        RaiseEvent Moused(Me, e)
    End Sub
    Public Sub Unsubscribe()

        If GlobalHook IsNot Nothing Then
            RemoveHandler GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
            RemoveHandler GlobalHook.KeyPress, AddressOf GlobalHookKeyPress
            'GlobalHook.Dispose()
        End If
        _Hooked = False

    End Sub
End Class