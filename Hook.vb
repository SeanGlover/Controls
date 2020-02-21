Option Explicit On
Option Strict On
Imports System.Windows.Forms
Imports Gma.System.MouseKeyHook
Public Class Hooker
    Public ReadOnly Property Hooked As Boolean
    Public Property Tag As Object
    Public Event Keyed(sender As Object, e As KeyPressEventArgs)
    Public Event Moused(sender As Object, e As MouseEventExtArgs)
    Private m_GlobalHook As IKeyboardMouseEvents
    Public Sub Subscribe()

        m_GlobalHook = Hook.GlobalEvents()
        AddHandler m_GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
        AddHandler m_GlobalHook.KeyPress, AddressOf GlobalHookKeyPress
        _Hooked = True

    End Sub
    Private Sub GlobalHookKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        RaiseEvent Keyed(Me, e)
    End Sub
    Private Sub GlobalHookMouseDownExt(ByVal sender As Object, ByVal e As MouseEventExtArgs)
        RaiseEvent Moused(Me, e)
    End Sub
    Public Sub Unsubscribe()

        If m_GlobalHook IsNot Nothing Then
            RemoveHandler m_GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
            RemoveHandler m_GlobalHook.KeyPress, AddressOf GlobalHookKeyPress
            m_GlobalHook.Dispose()
        End If
        _Hooked = False

    End Sub
End Class