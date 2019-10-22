﻿Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions
Public NotInheritable Class FindEventArgs
    Inherits EventArgs
    Public ReadOnly Property Text As String
    Public Sub New(FindText As String)
        Text = FindText
    End Sub
End Class
Public NotInheritable Class ZoneEventArgs
    Inherits EventArgs
    Public ReadOnly Property Zone As Zone
    Public Sub New(ClickedZone As Zone)
        Zone = ClickedZone
    End Sub
End Class
Public NotInheritable Class Zone
    Friend Sub New(ZoneName As Identifier)
        Name = ZoneName
        Select Case Name
            Case Identifier.Close
                _Image = My.Resources.CloseGray.ToBitmap
                Caption = "Close"

            Case Identifier.ExpandWidth
                _Image = Base64ToImage("iVBORw0KGgoAAAANSUhEUgAAABMAAAATCAYAAAByUDbMAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAABMSURBVDhPzY67DQAgCAUZjf2X0kYLroA8Y5RLKI6/9cfdx46VOqfvsqvwM7oEh+kSHKb/g59UnsLmylPYXPk7eFn1AIuqB1hUvTNmEyBwnSer7yuJAAAAAElFTkSuQmCC")
                Caption = "Drag to resize"

            Case Identifier.Filter
                _Image = Base64ToImage("data:_Image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS42/U4J6AAAAcVJREFUSEu9lk1LAkEYxxc8iC+n3ijwENTVW1BfoU5FYFL5kpW4vqzaquisMKdO0Rco6rt07wMEBXWI6B4FnbZ5ZmeW9XF2Ncz+8H+cl2d+fxeXQU3boiHm6JQcgoBok1B7GuYh/xEQZl6YksP/8gQR1eZfGNg8oNGjb6qGSVww6Qlj87colDykCVVIpWU96x16U2zTS2xYX81fzOMzYAm3bVvTeGGTlTRdUoXoZ+Rxu0I3WE+MOSIcm0udz+JesBfuBsiQ5T26aHTpPT6km72XnSpdhx4wjI1O/xv3YfhAAJ84gHjR7N/iw2WTvAIYXG1bn3g/2+gXMJwzByYgEXLcJNcYUmmRd+YPvJ4xrCMVnDN5YRpYdEKiuTrpYhj2QY1k/eBgpwgNbbKDOYO0VGDwfpVkfL+54LkTP6XrdK3Wtr5UAfCUom1IbtCoAAaJZEv1BwyH3wLCRZcrCZ44IF9uPsGe6OLCcM7mJUhjBmCwtFOCNCIAA7GdEiRFAFwpu6dmmu35vp5gLnfgJxQAV8mmTpOj4NJOCZInoNShV2weHwcO4p9y4ise0LgTFxn/p+AFqSzFx94FpRgwkaIzvwFL2bat/QBLBgP9DYp+awAAAABJRU5ErkJggg==")
                Caption = "Filter column"

            Case Identifier.FilterReset
                _Image = Base64ToImage("data:_Image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS42/U4J6AAAAa5JREFUSEu9VLtKA0EUTa2EYCSkEAIWNqIkiMVGIwHBzt5ecPPAHY2fYOV3WFrYhWgEq7CKEEG0iDYGYukLtfEB49yd3GFmMpMIuh44d2fuPfecPGAjlNJQGZSDSp6Gwf8LCJNBMaX/BUMPiCCOKrlXk+A3PPMmHfENWoXkMAtp6aI6yT75JHN+6k35OqHfLo66+g5jh5nHwVcEAK/dxBAL2dUX6p7zcFVO7dy50eWOG10Ewvm2MLKqaxkVcyUAebzp9HyqQ+I8X5THt2/W4gkgnGsbC1+arsccqFyQDS+9BAuKAcm+gDE3z30os0p+z+QDNDaBjfXpJFv0FSMy91Yj8+9Kr485kBcG+Qlsu7EYPMFAM5TZ1xzISxei2T3j02AcEPUm4q64mIBi9pt/DgrQIfqmIQJF1dLMvW4O/wXOZQ+5F/TloQ4UmQKq5dlHnKOHfBf9oFiAop8E2MiLBSiyBXRlQmciLxagSA+Al+NJcWJF1tjIiwUokgPYq6R5WUqN4WwQebEARRjgk/QW9nBPvsvEmbiYgOJ6KbNvepHZiAjOckMHznBxEHVQSiPfY2gC9/HDg+kAAAAASUVORK5CYII=")
                Caption = "Reset filter on column"

            Case Identifier.FiltersReset
                _Image = Base64ToImage("data:_Image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACQAAAAYCAYAAACSuF9OAAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMS42/U4J6AAAAn9JREFUSEvNls9rE0EUxxfipZat0Bq81IKHokZtigmY1BRBvRgVhELRk1DaSXbT3Ri8eVmqh/pPeO7Fm9RoBU8af2AUUbD2FFDswZ+o+BvW951k43YyuyFxt/TBJz9mZ977ZPbNEmVlpj9aY1uOhIGtK6pt20onKFh4o3TIDgMS2i0r6sfGE3oxMxBbZb1zYUBCg7Kifmy8HXrF1MOyZEEAISVr9cgKe8GFbpUyX2QJ/4eHRmyBhNSOhXDsl3Pb8jKpJTP9sWKOPr5v7K2IYLyWH2DiGlCd3Xn1uxY5ih36lt+UqbLtZ/jta4/KrSD1kkWnpVJG6v1zfejya6aewG6+zfWcEucEBfI3twpSK2zr1O1zqTfixJtm6tNTfcdFzMEvFq8HBXI3hRwpspy4Y8SftCww058hBcLoObpdifnC5LE1Qoim1Oy+astCc+xr2Tz4E5+76TkqOi6uAZBB/3AHR8QtBik8n+4W9lyXJXDw6jlKfoCIE06zJjEmzgU03pThLvylEc6gA8Qqhdg1WSLwSB9e8Oo5KpLmheqky8XxP+K8H1okO2ecveSuWX/xiQeFXaVyMfNbTAZWc73HUcyr5+gafn2C1v8Sr2Md0Xe6aCU7EsJRXNT2vxMTopcgi6RI3q7n3DgyvIDw4OxaaFFPfuDPDYqG1AiREOeJrJFBhCGE4Anr/eIp1SKDCFrInYxDp8qr5+jaSCPtvwhCCMf9Xn74Ck6hOxmHhLx6LhQhHPNn2tAFqQzwEMKOBi5UMeNV+j7hKQPWQ2hJG63h/w2e3r4yoEuhqfPWNH3ezJP4BQSWWfRkWxECgaKiEHqOxltPGCJrRQYnrX6827at/AWaftX0yNAQzgAAAABJRU5ErkJggg==")
                Caption = "Reset all filters on columns"

            Case Identifier.GotoNext
                _Image = My.Resources.Limit
                Caption = "Go to next match"

            Case Identifier.MatchCase
                _Image = Base64ToImage("iVBORw0KGgoAAAANSUhEUgAAABUAAAATCAYAAAB/TkaLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAAGTSURBVDhPrZHPK0RRGIZn40fZkCwU2UkjokmhmXtnmqamsVKysrCysEQpZkH+ACtWLK1szEIWiqzEwpKUsrGkJhvK4no+fWece++pUdfi6Zz3/d7z3jNnUkEQ/DtOMylOMylOMylO0/f9U3jM5/PLrnkzYkY2m+2iLFCuovO/EDMKhcKSFHLTmq5D0UwzYgZF5/BVLBYHtXQrmsEbgwPml3DN/tDzvHEzD4W55YAWVUWz/4AHO6O+ZE5Y12GD/RO8QqfMQ2HMNTkg76qHd0TzMc/OcaupiK5oriK6MRAovWH4bOl2CcO+nbPhmbqZ+5IjvyBeY8jXRrTgDDYJVJUXeDM5AT0JR+LrmR/Q4VKMbR2+O5ADc5KTd2dfh5pcpFwut+ENayZcinkPF0YbMplMixyAY9EUzIvO5XIzJkP5hHihUoLTai6aoI0UylzezzwT2T3o408dRd+q91uKcYfxWSqVOuwyA0WzcghWRLPuqpaiOuuq7sM/ny/2mL0Lfm6vrdPpdCtev+0ZYsZ/4DSTEaS+AQyj1eAqLTaAAAAAAElFTkSuQmCC")
                Caption = "Case-sensitive match"

            Case Identifier.MatchWord
                _Image = Base64ToImage("iVBORw0KGgoAAAANSUhEUgAAABUAAAATCAYAAAB/TkaLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAAGVSURBVDhPrZQ/S8NQFMU7KAhOIiri4KS4iEOmSJvmD/0AgrjqJOgiiOBgh0g3EQdB8SOITiLiIAh1c3ERFKcuburiouAQzynvPq4vgYp2+DX33XPuyctrSCkMw7TbMDTrNn/d6X21Wn0t6LcpZVmWAwMX4AmG1SK9E7lGuVweQJg8StPVf0OuEUXRMgOx0zNznXI9ncg1EHQFvpIkmTShqdZxUw/9E5cgCKbFY81mYBwGBtW5Rv0BHrUHwxPQDwzHmGlxBtcZ8VgzgWmDBp4r16gbZiDQPiGO4zHqYF/3f5gQegtDS637zNCh9imd/hdce3TfFjwTE3AJtmCsG57Bmx4i8KzRD23F1WwBcduEvhfA4Xnl5ROwfyM9jS1geADXWiSe5/UyFJxKD6FH7OFMR7RXaP/gj5ilCeYl10AYSB2v2WClUglYg3PMLQi+7w+LX4buEPhZq9X6RdDgvOdM0DpYNLXLpvjtIF6jIamLwA5HpebNXbSXu2x/WbDTHamLgL7rrBtgT/cEhhY9yr+wO+0eYfoNCEQi0nknzQMAAAAASUVORK5CYII=")
                Caption = "Match whole word"

            Case Identifier.Move
                _Image = My.Resources.Move
                Caption = "Drag to move"

            Case Identifier.RegEx
                _Image = Base64ToImage("iVBORw0KGgoAAAANSUhEUgAAABUAAAATCAYAAAB/TkaLAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAADNSURBVDhPrZIxCsJAEEVTWHoCK88RhCx4CQstg6XYi1h7ExsLOw/hESwECy8R1jeygWz4JoFN8ZjMy/CKkMx7PzpSpiJlKlJ2URTF2jn3VO9qpFQQK4ltmQyH+rlTnuez5p0RLV0QulqM0CPMN1Swb99GSx/ElkQ+IXpXN4aUCiIXi8EtRF9Q8bxr30ZLFwRK2MDKouaYZ/Z5886IliFYuI7+Q8pUpExFylSkNPh2E+WHICXBQ/htFup9H1ISPI4eNQhOlR+ClKlImYbPvm6MZAhmAxAaAAAAAElFTkSuQmCC")
                Caption = "Match using Regular Expressions"

            Case Identifier.ReplaceAll
                _Image = Base64ToImage("iVBORw0KGgoAAAANSUhEUgAAABkAAAANCAYAAABcrsXuAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAAGKSURBVDhPtZLNK0RhFMZnYaMoHwvJWlLEStM0c+erWErZW5HFWIgSSixYCFkoG7OyUJY2k5K1lYWiLDQbC6mbf0Bz/Z7pPdebbmmSxW/O857nnPe5M3dSURT9O4nNVikUChvQSPJEYrNVXEg9yROJzVZxIR9JnohFPp/vgkU4ZGGlWCwO+YOlUmkCbxXviDrvey6kkcvl+qlrsMV+YH48yOIbZkTdpl5KB0Ew6i6Z1tmxx8wz9dZ20QoxP4agBflxCI0x04KLXmHJO3d4uqJL+HZ9OqMtZMpm0HrQmnSzccXnX3AhjxYgeJBlCOMQ1zxgUAsX6BNqCJvyyuVyL707+dQqnEnzcw7KRyvkye5yvTnmPqUtYFxL2Wx21huqQzMEfxf9YB6XD2v+lxC9uxdpCxnRErWiM8szOoOF7KDfeYhudz6W/yNE86f6h/F+J5kJYV2+n3zuBhV2Q61BM4QXPIC+Nx/2Vf0Qdq6p/h1VuzsOEel0uj2TyXT6PR/8nqS+jx6IgLbvXpT6Avgh79zsR2LpAAAAAElFTkSuQmCC")
                Caption = "Replace all matches"

            Case Identifier.ReplaceOne
                _Image = Base64ToImage("iVBORw0KGgoAAAANSUhEUgAAABIAAAARCAYAAADQWvz5AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAAE6SURBVDhPrZK9SsRAFIXTbiEsIlaW/oAoLotICPkPSS+4ryBbaLWljZWVFsq+gYViYWFjY6lPYLHlNhY2PsN4znBnGScDsqvFx7333MmXhCRQSv0L3nARvOEieEOXLMvW8zxXVVVt+vbEG7oYUVmWfd+eeEOXuUQ43AVDcIWLRkVRbFm7mQj5MfobZAOzJ7bok4dRz1Ef2Kdpuis7LYLkFfWRPUF+3xLh0J7pCQ59gFPptQhcWPtrZnVdr3IOnhD/hhHhCTcs0ZpktRZZi0u56x36MeoXOJNdSxSGYYcZdkecjaTHMI5jHRLMU1cURdG+2SdJciAinRnRjoQnnHHnQ86uCLzzQtBFzw/yxv1MRBDeymEKX1CfgSsaSSUT/A7bLRHhe+Pxl+zMBdKVpmmW3fzH8Be84fyo4BugtOMkNlTnWwAAAABJRU5ErkJggg==")
                Caption = "Replace next match"

            Case Identifier.ShowHideReplace
                _Image = Base64ToImage(ChevronCollapsed)
                Caption = "Show or hide replace"

            Case Identifier.Tip
                _Image = My.Resources.LightOff
                Caption = "Click to turn off tips"
                Selected = True

        End Select
    End Sub
    Friend ReadOnly Caption As String
    Friend ReadOnly Property Name As Identifier
    Private ReadOnly _Image As Image
    Friend ReadOnly Property Image As Image
        Get
            If Selected Then
                If Name = Identifier.ShowHideReplace Then
                    Return Base64ToImage(ChevronExpanded)

                ElseIf Name = Identifier.Tip Then
                    Return My.Resources.LightOn

                Else
                    Return _Image

                End If
            Else
                Return _Image
            End If
        End Get
    End Property
    Private _Selected As Boolean
    Friend Property Selected As Boolean
        Get
            Return _Selected
        End Get
        Set(value As Boolean)
            Select Case Name
                Case Identifier.ShowHideReplace, Identifier.Filter, Identifier.FilterReset, Identifier.FiltersReset, Identifier.MatchCase, Identifier.MatchWord, Identifier.RegEx, Identifier.Tip
                    _Selected = value
                Case Else
                    _Selected = False
            End Select
        End Set
    End Property
    Friend Enum Identifier
        None
        ShowHideReplace
        GotoNext
        Close
        ReplaceOne
        ReplaceAll
        ExpandWidth
        MatchCase
        MatchWord
        RegEx
        Filter
        FilterReset
        FiltersReset
        Tip
        Move
    End Enum
End Class
Public Class FindReplace
    Inherits Control
    Private WithEvents FindTimer As New Timer With {.Interval = 500}
    Friend ReadOnly Tree As TreeViewer
#Region " GENERAL DECLARATIONS "
    Private ReadOnly Zones As New Dictionary(Of Zone.Identifier, Zone)
    Private ReadOnly ZonesBounds As New Dictionary(Of Zone, Rectangle)
    Private MouseOverZone As Zone
    Private MouseLocation As Point
    Private Const Spacing As Integer = 3
#End Region
#Region " Classes, Structures and Enumerations "
    Public Enum ParentType
        None
        GridControl
        TextControl
    End Enum
    Public Enum TypeGroup
        Booleans
        Dates
        Numbers
        Strings
        None
    End Enum
#End Region
#Region " Constructor "
    Public Sub New()

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)
        BackColor = Color.WhiteSmoke
        Values = New List(Of Object)
        DataType = GetType(String)
        MouseOverZone = Zones(Zone.Identifier.None)

    End Sub
#End Region
#Region " DRAWING "
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If e Is Nothing Then
        Else
            Using sb As New SolidBrush(BackColor)
                e.Graphics.FillRectangle(sb, ClientRectangle)
            End Using
            e.Graphics.SmoothingMode = SmoothingMode.HighQuality
            For Each Zone In Zones.Values
                Dim ZoneBounds As Rectangle = ZonesBounds(Zone)
                If Zone.Name = Zone.Identifier.None Then
                Else
                    e.Graphics.DrawImage(Zone.Image, ZoneBounds)
                End If
                If Not Zone.Name = Zone.Identifier.Tip And Zone.Name = MouseOverZone.Name Then
                    Using FadeBrush As New SolidBrush(Color.FromArgb(128, Color.Gold))
                        e.Graphics.FillRectangle(FadeBrush, ZoneBounds)
                    End Using
                End If
                If Not (Zone.Name = Zone.Identifier.ShowHideReplace Or Zone.Name = Zone.Identifier.Tip) And Zone.Selected Then
                    ZoneBounds.Inflate(1, 1)
                    e.Graphics.DrawRectangle(Pens.Red, ZoneBounds)
                End If
            Next
            Using Pen As New Pen(Brushes.DarkSlateBlue, Spacing)
                e.Graphics.DrawLine(Pen, 0, Height - Pen.Width, Width, Height - Pen.Width)
            End Using
            ControlPaint.DrawBorder3D(e.Graphics, ClientRectangle, Border3DStyle.RaisedOuter)
        End If

    End Sub
#End Region
#Region " PROPERTIES "
    Private ReadOnly Property ParentControl As Control
    Private WithEvents FindControl_ As Control
    Public ReadOnly Property FindControl As Control
        Get
            Return FindControl_
        End Get
    End Property
    Private WithEvents ReplaceControl_ As Control
    Public ReadOnly Property ReplaceControl As Control
        Get
            Return ReplaceControl_
        End Get
    End Property
    Private WithEvents ZoneTip As New ToolTip
    Private _DataSource As Object
    Public Property DataSource As Object
        Get
            Return _DataSource
        End Get
        Set(value As Object)
            If Not value Is _DataSource Then
                _DataSource = value
                If DataSource.GetType Is GetType(String) Then
                    DataType = GetType(String)
                    _SourceType = ParentType.TextControl

                ElseIf DataSource.GetType Is GetType(DataColumn) Then
                    Dim ColumnDataSource As DataColumn = DirectCast(DataSource, DataColumn)
                    Name = ColumnDataSource.ColumnName
                    _Values = DataColumnToList(ColumnDataSource)
                    DataType = ColumnDataSource.DataType
                    _SourceType = ParentType.GridControl

                ElseIf TypeOf DataSource Is IEnumerable Then
                    _Values = (From O In DirectCast(DataSource, IEnumerable).AsQueryable Select O).ToList
                    DataType = GetDataType(Values)
                    _SourceType = ParentType.GridControl

                End If
            End If
        End Set
    End Property
    Private _DataType As Type
    Private Property DataType As Type
        Get
            Return _DataType
        End Get
        Set(value As Type)
            If value IsNot _DataType Then
                _DataType = value
                Controls.Clear()
                Zones.Clear()

                Dim Identifiers = EnumNames(GetType(Zone.Identifier))
                For Each Item In Identifiers.Skip(0)
                    Dim ZI As Zone.Identifier = DirectCast([Enum].Parse(GetType(Zone.Identifier), Item), Zone.Identifier)
                    Zones.Add(ZI, New Zone(ZI))
                    ZonesBounds.Add(Zones(ZI), New Rectangle)
                Next

#Region " ADD NEW CONTROLS "
                Select Case Types
                    Case TypeGroup.Booleans
                        Dim Find As New CheckBox
                        With Find
                            AddHandler .CheckedChanged, AddressOf RequestMade
                        End With
                        FindControl_ = Find

                        Dim Replace As New CheckBox
                        With Replace
                            AddHandler .CheckedChanged, AddressOf RequestMade
                        End With
                        ReplaceControl_ = Replace

                    Case TypeGroup.Dates
                        Dim Find As New DatePicker
                        With Find
                            AddHandler .DateChanged, AddressOf RequestMade
                        End With
                        FindControl_ = Find

                        Dim Replace As New DatePicker
                        With Replace

                        End With
                        ReplaceControl_ = Replace
#Region " BTV "
                        'Tree = New Tree With {.Visible = False, .Margin = New Padding(0), .Padding = .Margin, .Height = 400, .Font = Font}
                        'Dim Node As Node = Tree.Nodes.Add(New Node With {.Font = Font, .Text = "Day of Week"})
                        'Node.Nodes.AddRange((From D In List Order By D.DayOfWeek Ascending Select WeekdayName(Weekday(D))).Distinct.Select(Function(T) New Node With {.Text = T, .Name = T, .CheckBox = True, .Font = Font}).ToArray)
                        'Node = Tree.Nodes.Add(New Node With {.Font = Font, .Text = "First day of the Month"})
                        'Node.Nodes.AddRange((List.Where(Function(x) x.Day = 1).Distinct.Select(Function(T) New Node With {.Text = Microsoft.VisualBasic.Format(T, Format), .Name = T.ToShortDateString, .CheckBox = True, .Font = Font}).ToArray))
                        'Node = Tree.Nodes.Add(New Node With {.Font = Font, .Text = "Last day of the Month"})
                        'Node.Nodes.AddRange((List.Where(Function(x) x.AddDays(1).Day = 1).Distinct.Select(Function(T) New Node With {.Text = Microsoft.VisualBasic.Format(T, Format), .Name = T.ToShortDateString, .CheckBox = True, .Font = Font}).ToArray))
                        'Node = Tree.Nodes.Add(New Node With {.Font = Font, .Text = "Month"})
                        'Node.Nodes.AddRange((List.Select(Function(x) MonthName(x.Month, True) & " " & x.Year.ToString).Distinct.Select(Function(T) New Node With {.Text = T, .Name = T, .CheckBox = True, .Font = Font}).ToArray))
                        'Node = Tree.Nodes.Add(New Node With {.Font = Font, .Text = "Date"})
                        'Node.Nodes.AddRange((List.Distinct.Select(Function(T) New Node With {.Text = Microsoft.VisualBasic.Format(T, Format), .Name = T.ToShortDateString, .CheckBox = True, .Font = Font}).ToArray))
                        'AddHandler Tree.NodeChecked, AddressOf RequestsMade
                        'AddHandler Tree.SizeChanged, AddressOf OnTreeSizeChanged
                        'Controls.Add(Tree)
#End Region

                    Case Else
                        Dim Find As New ImageCombo
                        With Find
                            .DropDown.CheckOnSelect = False
                            AddHandler .SelectionChanged, AddressOf RequestMade
                        End With
                        FindControl_ = Find

                        Dim Replace As New ImageCombo
                        With Replace
                            .DropDown.CheckBoxes = False
                            AddHandler .SelectionChanged, AddressOf RequestMade
                        End With
                        ReplaceControl_ = Replace

                        Select Case Types
                            Case TypeGroup.Strings

                            Case TypeGroup.Numbers

                        End Select

                End Select
                With FindControl
                    .Name = "Find"
                    .Size = New Size(200, 28)
                    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
                    Dim Zone_ShowHide = Zones(Zone.Identifier.ShowHideReplace)
                    Dim wh_ShowHide As Size = Zone_ShowHide.Image.Size
                    Dim xy_ShowHide As Point = New Point(Spacing,
                                         Spacing + Convert.ToInt32((FindControl.Height - wh_ShowHide.Height) / 2))
                    ZonesBounds(Zone_ShowHide) = New Rectangle(xy_ShowHide, wh_ShowHide)
                    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
                    .Location = New Point(Spacing + Zone_ShowHide.Image.Width + Spacing, Spacing)
                    .Visible = True
                End With
                With ReplaceControl
                    .Name = "Replace"
                    .Location = FindControl.Location
                    .Size = New Size(200, 28)
                    .Visible = False
                End With

                AddHandler FindControl.KeyDown, AddressOf On_KeyDown
                AddHandler FindControl.GotFocus, AddressOf OnControlFocus
                AddHandler FindControl.TextChanged, AddressOf OnFindTextChanged

                AddHandler ReplaceControl.KeyDown, AddressOf On_KeyDown
                AddHandler ReplaceControl.GotFocus, AddressOf OnControlFocus

                Controls.AddRange({FindControl, ReplaceControl})
#End Region
                ResizeMe()

            End If
        End Set
    End Property
    Private ReadOnly Property SourceType As ParentType
    Public ReadOnly Property Types As TypeGroup
        Get
            Select Case DataType
                Case GetType(Boolean)
                    Return TypeGroup.Booleans

                Case GetType(Date)
                    Return TypeGroup.Dates

                Case GetType(String)
                    Return TypeGroup.Strings

                Case GetType(Decimal), GetType(Double), GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                    Return TypeGroup.Numbers

                Case Else
                    Return TypeGroup.None

            End Select
        End Get
    End Property
    Public ReadOnly Property Values As List(Of Object)
    Friend ReadOnly Property Zone_Case As Zone
        Get
            Return Zones(Zone.Identifier.MatchCase)
        End Get
    End Property
    Friend ReadOnly Property Zone_Word As Zone
        Get
            Return Zones(Zone.Identifier.MatchWord)
        End Get
    End Property
    Friend ReadOnly Property Zone_Regex As Zone
        Get
            Return Zones(Zone.Identifier.RegEx)
        End Get
    End Property
    Friend ReadOnly Property Zone_Filter As Zone
        Get
            Return Zones(Zone.Identifier.Filter)
        End Get
    End Property
    Friend ReadOnly Property Zone_FilterX As Zone
        Get
            Return Zones(Zone.Identifier.FilterReset)
        End Get
    End Property
    Friend ReadOnly Property Zone_FiltersX As Zone
        Get
            Return Zones(Zone.Identifier.FiltersReset)
        End Get
    End Property
    Friend ReadOnly Property Zone_GotoClickPoint As Point
    Friend ReadOnly Property SearchPattern As String
        Get
            If Zone_Word.Selected Then
                Return "\b" & FindControl.Text & "\b"
            ElseIf Zone_Regex.Selected Then
                Return FindControl.Text
            Else
                'Make all special characters literal
                Dim Input As String = FindControl.Text
                Dim Replaceables = New List(Of String) From {"+", "&", "|", "!", "(", ")", "{", "}", "[", "]", "^", "~", "*", "?", ":", "/", "\"}
                Dim rxString As String = String.Join("|", Replaceables.Select(Function(r) "\" & r))
                Dim SpecialMatches = RegexMatches(Input, rxString, RegexOptions.IgnoreCase).OrderByDescending(Function(rm) rm.Index)
                For Each SpecialMatch In SpecialMatches
                    Input = Input.Remove(SpecialMatch.Index, SpecialMatch.Length)
                    Input = Input.Insert(SpecialMatch.Index, "\" & SpecialMatch.Value)
                Next
                Return Input
            End If
        End Get
    End Property
    Friend ReadOnly Property SearchOptions As RegexOptions
        Get
            If Zone_Case.Selected Then
                Return RegexOptions.ExplicitCapture
            Else
                Return RegexOptions.IgnoreCase
            End If
        End Get
    End Property
    Friend ReadOnly Property Matches As Dictionary(Of Integer, String)
        Get
            Dim MD As New Dictionary(Of Integer, String)
            If FindControl.Text Is Nothing Then
            Else
                Values.Clear()
                If SourceType = ParentType.TextControl Then
                    Dim SearchText As String = DirectCast(DataSource, String)
                    For Each Match In RegexMatches(SearchText, SearchPattern, SearchOptions)
                        Values.Add(Match.Value)
                        MD.Add(Match.Index, Match.Value)
                    Next
                Else

                End If
            End If
            If MD.Any Then
                FindControl.ForeColor = Color.Black
            Else
                FindControl.ForeColor = Color.Red
            End If
            Return MD
        End Get
    End Property
    Private _CurrentMatch As KeyValuePair(Of Integer, String)
    Friend ReadOnly Property CurrentMatch As KeyValuePair(Of Integer, String)
        Get
            If Matches.Any Then
                Return _CurrentMatch
            Else
                Return New KeyValuePair(Of Integer, String)(-1, String.Empty)
            End If
        End Get
    End Property
    Private ReadOnly Property CurrentMatchIndex As Integer
        Get
            Dim Indices As New List(Of Integer)(Matches.Keys)
            Return Indices.IndexOf(CurrentMatch.Key)
        End Get
    End Property
    Private ReadOnly Property NextMatchCaption As String
        Get
            Return Join({MouseOverZone.Caption, "[", "#" & CurrentMatchIndex + 1, "of", Matches.Count, "]"})
        End Get
    End Property
    Private _StartAt As Integer = -1
    Friend Property StartAt As Integer
        Get
            Return _StartAt
        End Get
        Set(value As Integer)
            _StartAt = value
            Dim MatchDictionary = Matches
            If value >= 0 And MatchDictionary.Any Then
                Dim NextMatch = From m In MatchDictionary Where m.Key > value
                If NextMatch.Any Then
                    _CurrentMatch = NextMatch.First
                Else
                    _CurrentMatch = MatchDictionary.First
                End If
            End If
        End Set
    End Property
#End Region
#Region " EVENTS "
    Friend Event ZoneClicked(sender As Object, e As ZoneEventArgs)
    Friend Event FindChanged(sender As Object, e As FindEventArgs)
    Private Sub RequestMade(sender As Object, e As EventArgs)

        If e.GetType Is GetType(DateRangeEventArgs) Or e.GetType Is GetType(KeyEventArgs) Or e.GetType Is GetType(ImageComboEventArgs) Then
        ElseIf e.GetType Is GetType(MouseEventArgs) Then
        End If

    End Sub
    Private Sub OnTreeSizeChanged(sender As Object, e As EventArgs)
        ResizeMe()
    End Sub
    Private Sub OnFindTextChanged(sender As Object, e As EventArgs)
        FindControl.ForeColor = Color.Black
        FindTimer.Stop()
        FindTimer.Start()
    End Sub
    Private Sub FindTimer_Tick(sender As Object, e As EventArgs) Handles FindTimer.Tick
        FindTimer.Stop()
        RaiseEvent FindChanged(Me, New FindEventArgs(FindControl.Text))
    End Sub
    Private Sub OnControlFocus(sender As Object, e As EventArgs)

        FindControl.BackColor = SystemColors.Control
        ReplaceControl.BackColor = SystemColors.Control
        DirectCast(sender, Control).BackColor = Color.White
        Invalidate()

    End Sub
    Private Sub On_KeyDown(sender As Object, ByVal e As KeyEventArgs)

        DirectCast(sender, Control).BackColor = Color.White
        If e.KeyCode = Keys.Enter Then
            RequestMade(sender, e)
        End If
        MyBase.OnKeyDown(e)

    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            If Not Cursor.Position = MouseLocation Then
                If MouseOverZone.Name = Zone.Identifier.GotoNext Then
                    ZoneTip.Show(NextMatchCaption, Me, New Point(0, Height))
                ElseIf Zones(Zone.Identifier.Tip).Selected Then
                    ZoneTip.Show(MouseOverZone.Caption, Me, New Point(0, Height))
                End If
                If e.Button <> MouseButtons.Left Then
                    Cursor = Cursors.Default
                    MouseOverZone = Zones(Zone.Identifier.None)
                    For Each Zone In ZonesBounds
                        If Zone.Value.Contains(e.Location) Then MouseOverZone = Zone.Key
                    Next
                    Invalidate()
                Else
                    Dim x_Delta As Integer = Cursor.Position.X - MouseLocation.X
                    Dim y_Delta As Integer = Cursor.Position.Y - MouseLocation.Y
                    Dim Zone_Move As Zone = Zones(Zone.Identifier.Move)
                    Dim Zone_Width As Zone = Zones(Zone.Identifier.ExpandWidth)
                    If MouseOverZone.Name = Zone_Move.Name Then
                        Cursor = CursorDirection(Cursor.Position, MouseLocation)
                        Left += x_Delta
                        Top += y_Delta

                    ElseIf MouseOverZone.Name = Zone_Width.Name Then
                        Cursor = CursorDirection(New Point(Cursor.Position.X, 0), New Point(MouseLocation.X, 0))
                        Left += x_Delta
                        Width -= x_Delta
                        FindControl.Width -= x_Delta
                        ReplaceControl.Width = FindControl.Width
                        ResizeMe()

                    End If
                    MouseLocation = Cursor.Position
                End If
            End If
            MyBase.OnMouseMove(e)
        End If

    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

        MouseLocation = Cursor.Position
        MouseOverZone.Selected = Not MouseOverZone.Selected

        Select Case MouseOverZone.Name
            Case Zone.Identifier.None

            Case Zone.Identifier.MatchCase, Zone.Identifier.MatchWord, Zone.Identifier.RegEx
                RaiseEvent ZoneClicked(Me, New ZoneEventArgs(MouseOverZone))

            Case Zone.Identifier.Close
                Close()

            Case Zone.Identifier.ShowHideReplace
                ReplaceControl.Visible = Not ReplaceControl.Visible
                ResizeMe()

            Case Zone.Identifier.GotoNext
                If Matches.Any Then
                    ZoneTip.Show(NextMatchCaption, Me, New Point(0, Height))
                    Dim Match0 As KeyValuePair(Of Integer, String) = Matches.First
                    If CurrentMatch.Value Is Nothing Then
                        _CurrentMatch = Match0

                    Else
                        Dim NextMatch As IEnumerable(Of KeyValuePair(Of Integer, String)) = From m In Matches Where m.Key > CurrentMatch.Key
                        If NextMatch.Any Then
                            _CurrentMatch = NextMatch.First
                        Else
                            _CurrentMatch = Match0
                        End If

                    End If
                    RaiseEvent ZoneClicked(Me, New ZoneEventArgs(MouseOverZone))
                End If

            Case Zone.Identifier.ReplaceOne
                If Matches.Any Then
                    If CurrentMatch.Value Is Nothing Then _CurrentMatch = Matches.First
                    RaiseEvent ZoneClicked(Me, New ZoneEventArgs(MouseOverZone))
                    MouseOverZone = Zones(Zone.Identifier.GotoNext)
                    OnMouseDown(e)
                    MouseOverZone = Zones(Zone.Identifier.ReplaceOne)
                End If

            Case Zone.Identifier.ReplaceAll
                If Matches.Any Then
                    If CurrentMatch.Value Is Nothing Then _CurrentMatch = Matches.First
                    RaiseEvent ZoneClicked(Me, New ZoneEventArgs(MouseOverZone))
                End If

            Case Zone.Identifier.Tip
                If Not MouseOverZone.Selected Then ZoneTip.Hide(Me)
                Invalidate()

        End Select
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        Cursor = Cursors.Default
        MyBase.OnMouseUp(e)
    End Sub
    Protected Overrides Sub OnParentChanged(e As EventArgs)

        Select Case True
            Case Parent Is Nothing
                If ParentControl Is Nothing Then
                    'No change
                Else
                    'Control changed to Nothing
                    RemoveHandler ParentControl.KeyDown, AddressOf ParentCtrlF

                End If

            Case Parent IsNot Nothing
                If ParentControl Is Nothing Then
                    'Nothing to Control
                    AddHandler Parent.KeyDown, AddressOf ParentCtrlF

                Else
                    'Changing Controls
                    RemoveHandler ParentControl.KeyDown, AddressOf ParentCtrlF
                    AddHandler Parent.KeyDown, AddressOf ParentCtrlF

                End If

        End Select
        _ParentControl = Parent
        Hide()
        MyBase.OnParentChanged(e)

    End Sub
    Private Sub ParentCtrlF(sender As Object, e As KeyEventArgs)

        If Control.ModifierKeys = Keys.Control And e.KeyCode = Keys.F Then
            Location = New Point(Parent.ClientSize.Width - Width - Spacing, Spacing)
            Show()
            FindControl.Focus()
        End If

    End Sub
#End Region
#Region " METHODS "
    Private Sub ResizeMe()

        Dim Zone_ShowHide As Zone = Zones(Zone.Identifier.ShowHideReplace)
        If ReplaceControl.Visible Then
            ReplaceControl.Top = FindControl.Top + FindControl.Height + Spacing
            Height = {Spacing, FindControl.Height, Spacing, ReplaceControl.Height, Spacing, Zone_Case.Image.Height, Spacing * 4}.Sum

        Else
            ReplaceControl.Top = FindControl.Top
            Height = {Spacing, FindControl.Height, Spacing, 0, Spacing, Zone_Case.Image.Height, Spacing * 3}.Sum

        End If
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim Zone_Goto As Zone = Zones(Zone.Identifier.GotoNext)
        Dim wh_Goto As Size = Zone_Goto.Image.Size
        Dim xy_Goto As Point = New Point({Spacing, Zone_ShowHide.Image.Width, Spacing, FindControl.Width, Spacing}.Sum,
                                         FindControl.Top + Convert.ToInt32((FindControl.Height - wh_Goto.Height) / 2))
        ZonesBounds(Zone_Goto) = New Rectangle(xy_Goto, wh_Goto)
        _Zone_GotoClickPoint = ZonesBounds(Zone_Goto).Location
        _Zone_GotoClickPoint.Offset(Convert.ToInt32(wh_Goto.Width / 2), Convert.ToInt32(wh_Goto.Height / 2))
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim Zone_Close As Zone = Zones(Zone.Identifier.Close)
        Dim wh_Close As Size = Zone_Close.Image.Size
        Dim xy_Close As Point = New Point(ZonesBounds(Zone_Goto).Right + Spacing,
                                          FindControl.Top + Convert.ToInt32((FindControl.Height - wh_Close.Height) / 2))
        ZonesBounds(Zone_Close) = New Rectangle(xy_Close, wh_Close)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim Zone_ReplaceOne As Zone = Zones(Zone.Identifier.ReplaceOne)
        Dim wh_ReplaceOne As Size = Zone_ReplaceOne.Image.Size
        Dim xy_ReplaceOne As Point = New Point(ReplaceControl.Left + If(ReplaceControl.Visible, {ReplaceControl.Width, Spacing}.Sum, 0),
                                         If(ReplaceControl.Visible, ReplaceControl.Top + Convert.ToInt32((ReplaceControl.Height - wh_ReplaceOne.Height) / 2), -100))
        ZonesBounds(Zone_ReplaceOne) = New Rectangle(xy_ReplaceOne, wh_ReplaceOne)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim Zone_ReplaceAll As Zone = Zones(Zone.Identifier.ReplaceAll)
        Dim wh_ReplaceAll As Size = Zone_ReplaceAll.Image.Size
        Dim xy_ReplaceAll As Point = New Point(ZonesBounds(Zone_ReplaceOne).Right + Spacing,
                                          If(ReplaceControl.Visible, ReplaceControl.Top + Convert.ToInt32((ReplaceControl.Height - wh_ReplaceAll.Height) / 2), -100))
        ZonesBounds(Zone_ReplaceAll) = New Rectangle(xy_ReplaceAll, wh_ReplaceAll)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim Zone_Width As Zone = Zones(Zone.Identifier.ExpandWidth)
        Dim wh_Width As Size = Zone_Width.Image.Size
        Dim xy_Width As Point = New Point(0,
                                         Height - wh_Width.Height - 4)
        ZonesBounds(Zone_Width) = New Rectangle(xy_Width, wh_Width)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim ImageTop As Integer = ReplaceControl.Top + ReplaceControl.Height
        Dim ImageBottom As Integer = Height - Spacing
        Dim wh_Case As Size = Zone_Case.Image.Size
        Dim xy_Case As Point = New Point(FindControl.Left,
                                         ImageTop + Convert.ToInt32((ImageBottom - ImageTop - wh_Case.Height) / 2))
        ZonesBounds(Zone_Case) = New Rectangle(xy_Case, wh_Case)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim wh_Word As Size = Zone_Word.Image.Size
        Dim xy_Word As Point = New Point(ZonesBounds(Zone_Case).Right + Spacing,
                                         ImageTop + Convert.ToInt32((ImageBottom - ImageTop - wh_Word.Height) / 2))
        ZonesBounds(Zone_Word) = New Rectangle(xy_Word, wh_Word)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim wh_Regex As Size = Zone_Regex.Image.Size
        Dim xy_Regex As Point = New Point(ZonesBounds(Zone_Word).Right + Spacing,
                                         ImageTop + Convert.ToInt32((ImageBottom - ImageTop - wh_Regex.Height) / 2))
        ZonesBounds(Zone_Regex) = New Rectangle(xy_Regex, wh_Regex)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim wh_Filter As Size = Zone_Filter.Image.Size
        Dim xy_Filter As Point = New Point(ZonesBounds(Zone_Regex).Right + Spacing,
                                         If(SourceType = ParentType.GridControl, ImageTop + Convert.ToInt32((ImageBottom - ImageTop - wh_Filter.Height) / 2), -100))
        ZonesBounds(Zone_Filter) = New Rectangle(xy_Filter, wh_Filter)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim wh_FilterX As Size = Zone_FilterX.Image.Size
        Dim xy_FilterX As Point = New Point(ZonesBounds(Zone_Filter).Right + Spacing,
                                         If(SourceType = ParentType.GridControl, ImageTop + Convert.ToInt32((ImageBottom - ImageTop - wh_FilterX.Height) / 2), -100))
        ZonesBounds(Zone_FilterX) = New Rectangle(xy_FilterX, wh_FilterX)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim wh_FiltersX As Size = Zone_FiltersX.Image.Size
        Dim xy_FiltersX As Point = New Point(ZonesBounds(Zone_FilterX).Right + Spacing,
                                         If(SourceType = ParentType.GridControl, ImageTop + Convert.ToInt32((ImageBottom - ImageTop - wh_FiltersX.Height) / 2), -100))
        ZonesBounds(Zone_FiltersX) = New Rectangle(xy_FiltersX, wh_FiltersX)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim Zone_Move As Zone = Zones(Zone.Identifier.Move)
        Dim wh_Move As Size = Zone_Move.Image.Size
        Dim xy_Move As Point = New Point(ZonesBounds(Zone_Close).Left - Spacing,
                                         ZonesBounds(Zone_Case).Top)
        ZonesBounds(Zone_Move) = New Rectangle(xy_Move, wh_Move)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Dim Zone_Tip As Zone = Zones(Zone.Identifier.Tip)
        Dim wh_Tip As Size = Zone_Tip.Image.Size
        Dim xy_Tip As Point = New Point(ZonesBounds(Zone_Move).Left - Spacing - wh_Tip.Width,
                                         ImageTop + Convert.ToInt32((ImageBottom - ImageTop - wh_Tip.Height) / 2))
        ZonesBounds(Zone_Tip) = New Rectangle(xy_Tip, wh_Tip)
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Width = ZonesBounds(Zone_Close).Right + Spacing * 2
        Invalidate()

    End Sub
    Friend Sub Close()
        MouseOverZone = Zones(Zone.Identifier.Close)
        Hide()
        RaiseEvent ZoneClicked(Me, New ZoneEventArgs(MouseOverZone))
        _CurrentMatch = New KeyValuePair(Of Integer, String)(-1, String.Empty)
    End Sub
#End Region
End Class