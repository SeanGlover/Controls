Option Explicit On
Option Strict On
Imports System.Windows.Forms
Imports System.Drawing
Public Class Dummy
    Inherits Control
    Public Sub New()

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)

    End Sub
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If e IsNot Nothing Then
            Using backBrush As New SolidBrush(BackColor)
                e.Graphics.FillRectangle(backBrush, Bounds)
            End Using
            Using textAlignment As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .FormatFlags = StringFormatFlags.FitBlackBox}
                Using TextBrush As New SolidBrush(ForeColor)
                    e.Graphics.DrawString(
                        Text,
                        Font,
                        TextBrush,
                        Bounds,
                        textAlignment
                    )
                End Using
            End Using
        End If

    End Sub
End Class