Option Explicit On
Option Strict On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class DatePicker
    Inherits Control
    Private Enum MouseRegion
        None
        Search
        Drop
        Clear
        Text
        WeekDay
        Month
        Day
        Year
    End Enum
    Private MouseOver As New MouseRegion
    Friend Toolstrip As New ToolStripDropDown With {.AutoClose = False, .AutoSize = False, .Padding = New Padding(0), .DropShadowEnabled = False, .BackColor = Color.Transparent, .Visible = False}

    Private SearchBounds As New Rectangle
    Private DateBounds As New Rectangle
    Private YearBounds As New Rectangle
    Private MonthBounds As New Rectangle
    Private DayBounds As New Rectangle

    Private ReadOnly ClearTextImage As Image = Base64ToImage(ClearTextString)
    Private ClearTextBounds As New Rectangle
    Private ClearTextDrawBounds As New Rectangle

    Private ReadOnly DropImage As Image = Base64ToImage(DropString)
    Private DropBounds As New Rectangle
    Private DropDrawBounds As New Rectangle
    Private DropRectangle As Rectangle
    Private DropPoints As Point()

    Private PaddedBounds As New Rectangle

    Private Field As Integer = 0
    'Private HOffset As Integer
    'Private VOffset As Integer
    Private InBounds As Boolean
    Private ReadOnly SB As New System.Text.StringBuilder With {.Capacity = 2}

    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, False)
        SetStyle(ControlStyles.UserMouse, True)
        BackColor = Color.WhiteSmoke
        BackColor = SystemColors.Window
        Size = New Size(200, 25)
        DropDown = New MonthCalendarDropDown(Me)
    End Sub
    Protected Overrides Sub InitLayout()
        Toolstrip.Items.Add(New ToolStripControlHost(DropDown))
        MyBase.InitLayout()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)

        If e IsNot Nothing Then
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            Using backBrush As New SolidBrush(BackColor)
                e.Graphics.FillRectangle(backBrush, ClientRectangle)
            End Using
            '// a border needs to be drawn otherwise annoying flickering
            Dim borderBounds As Rectangle = ClientRectangle
            Using Pen As New Pen(BackColor, 2)
                e.Graphics.DrawRectangle(Pen, borderBounds)
                borderBounds.Inflate(-1, -1)
                e.Graphics.DrawRectangle(Pen, borderBounds)
                borderBounds.Inflate(-1, -1)
                e.Graphics.DrawRectangle(Pen, borderBounds)
            End Using
            Using Brush As New SolidBrush(Color.FromArgb(128, Color.Blue))
                e.Graphics.FillPolygon(Brush, DropPoints)
            End Using
            Using Pen As New Pen(Brushes.Silver)
                e.Graphics.DrawLines(Pen, DropPoints)
            End Using
            If HasSearch Then
                Dim searchBounds = New Rectangle(2 + PaddedBounds.X, 0, 16, PaddedBounds.Height)
                Using searchBrush As New SolidBrush(Color.Transparent)
                    e.Graphics.FillRectangle(searchBrush, searchBounds)
                    Using sf As New StringFormat With {
                                                .Alignment = StringAlignment.Center,
                                                .LineAlignment = StringAlignment.Center
                                                }
                        Using searchFont As New Font("Tahoma", 16)
                            e.Graphics.DrawString(SearchDrawString, searchFont, Brushes.Black, searchBounds, sf)
                        End Using
                    End Using
                End Using
            End If
            If ValueIsNull Then
                If HintText IsNot Nothing Then TextRenderer.DrawText(e.Graphics, HintText, Font, DateBounds, Color.DarkGray, TextFormatFlags.VerticalCenter)
            Else
                If Not SelectionPixelStart = SelectionPixelEnd Then
                    Using foreBrush As New SolidBrush(ForeColor)
                        TextRenderer.DrawText(e.Graphics, ValueString, Font, DateBounds, ForeColor, Color.Transparent, TextFormatFlags.NoPadding Or TextFormatFlags.VerticalCenter)
                    End Using
                    Using sf As StringFormat = New StringFormat With
                    {
                    .LineAlignment = StringAlignment.Center,
                    .Alignment = If(HorizontalAlignment = HorizontalAlignment.Center, StringAlignment.Center, If(HorizontalAlignment = HorizontalAlignment.Left, StringAlignment.Near, StringAlignment.Far))
                    }
                        Dim range = New CharacterRange(0, ValueString.Length)
                        sf.SetMeasurableCharacterRanges({range})
                        Dim regions = e.Graphics.MeasureCharacterRanges(ValueString, Font, DateBounds, sf)
                        If regions.Any Then
                            Dim accurateBoundings As RectangleF = regions.First.GetBounds(e.Graphics)
                            Dim textTop As Integer = {PaddedBounds.Y + BorderWidth + 2, CInt(accurateBoundings.Y) - 4}.Min
                            Dim accurateBounds As New Rectangle(SelectionPixelStart, textTop, SelectionPixelEnd - SelectionPixelStart, PaddedBounds.Height - (1 + textTop * 2))
                            Using dateBrush As New SolidBrush(Color.Silver)
                                Using datePen As New Pen(dateBrush, 1)
                                    e.Graphics.DrawRectangle(datePen, accurateBounds)
                                End Using
                            End Using
                            Using fillBrush As New SolidBrush(Color.FromArgb(64, Color.Silver))
                                e.Graphics.FillRectangle(fillBrush, accurateBounds)
                            End Using
                            'If Value = New Date(2021, 5, 30) Then Stop
                        End If
                        'e.Graphics.DrawString(ValueString,
                        '                              Font,
                        '                              foreBrush,
                        '                              DateBounds,
                        '                              sf)
                    End Using
                End If
            End If
            If MouseOver = MouseRegion.Search Or MouseOver = MouseRegion.Drop Or MouseOver = MouseRegion.Clear Then
                Dim regionBounds As Rectangle = If(MouseOver = MouseRegion.Search, SearchBounds, If(MouseOver = MouseRegion.Drop, DropDrawBounds, ClearTextDrawBounds))
                Using Brush As New Drawing2D.LinearGradientBrush(regionBounds, Color.FromArgb(60, Color.AliceBlue), Color.FromArgb(60, Color.LightSkyBlue), linearGradientMode:=Drawing2D.LinearGradientMode.Vertical)
                    e.Graphics.FillRectangle(Brush, regionBounds)
                End Using
                Using Pen As New Pen(Brushes.SkyBlue)
                    e.Graphics.DrawRectangle(Pen, regionBounds)
                End Using
            End If

            Dim colorBorder As Color = If(HighlightBorderOnFocus And InBounds, HighlightBorderColor, If(BorderColor = Color.Transparent, BackColor, BorderColor))
            borderBounds = New Rectangle(PaddedBounds.X, PaddedBounds.Y, PaddedBounds.Width - 1, PaddedBounds.Height - 1)
            For i = 0 To BorderWidth - 1
                Using Pen As New Pen(colorBorder, 1)
                    e.Graphics.DrawRectangle(Pen, borderBounds)
                End Using
                borderBounds.Inflate(-1, -1)
            Next
            e.Graphics.DrawImage(ClearTextImage, ClearTextBounds)
        End If

    End Sub
    Private Sub Bounds_Set()

        With Margin
            PaddedBounds = New Rectangle(.Left, .Top, Width - (.Left + .Right), Height - (.Top + .Bottom))
        End With

        Const spacing As Integer = 2
        With SearchBounds
            .X = PaddedBounds.X + If(HasSearch, spacing, 0)
            .Y = PaddedBounds.Y
            .Width = If(HasSearch, 16, 0)
            .Height = PaddedBounds.Height
        End With

        Const DropArrowW As Integer = 8
        Const DropArrowH As Integer = 4

        DropRectangle = New Rectangle(PaddedBounds.Width - 1 - 16, 1, 16, PaddedBounds.Height - spacing)
        Dim LeftPt As Integer = DropRectangle.Left + CInt((DropRectangle.Width - DropArrowW) / 2)
        Dim RightPt As Integer = LeftPt + DropArrowW
        Dim MidPt As Integer = LeftPt + CInt(DropArrowW / 2)
        Dim DropY As Integer = CInt((PaddedBounds.Height - DropArrowH) / 2)
        DropPoints = {New Point(LeftPt, DropY), New Point(RightPt, DropY), New Point(MidPt, DropY + DropArrowH)}

        '// SearchBounds | TextBounds | ClearTextBounds | DropBounds

        With DropBounds
            'V LOOKS BETTER WHEN NOT RESIZED
            .X = PaddedBounds.Right - (DropImage.Width + spacing)
            .Y = {0, CInt((PaddedBounds.Height - DropImage.Height) / 2)}.Max     'Might be negative if DropImage.Height > PaddedBounds.Height
            .Width = DropImage.Width
            .Height = {PaddedBounds.Height, DropImage.Height}.Min
            DropDrawBounds.X = .X : DropDrawBounds.Y = 0 : DropDrawBounds.Width = .Width : DropDrawBounds.Height = PaddedBounds.Height
        End With
        With ClearTextBounds
            If ValueIsNull Then
                .X = DropBounds.Left
                .Y = 0
                .Width = 0
                .Height = PaddedBounds.Height
                ClearTextDrawBounds = ClearTextBounds
            Else
                'X LOOKS BETTER WHEN NOT RESIZED
                .X = PaddedBounds.Right - ({DropBounds.Width, ClearTextImage.Width}.Sum + spacing)
                .Y = {0, CInt((PaddedBounds.Height - ClearTextImage.Height) / 2)}.Max
                .Width = ClearTextImage.Width
                .Height = {PaddedBounds.Height, ClearTextImage.Height}.Min
                ClearTextDrawBounds.X = .X : ClearTextDrawBounds.Y = 0 : ClearTextDrawBounds.Width = .Width : ClearTextDrawBounds.Height = PaddedBounds.Height
            End If
        End With
        With DateBounds
            .X = SearchBounds.Right + spacing
            .Y = PaddedBounds.Top
            .Width = ClearTextBounds.Left - SearchBounds.Right
            .Height = PaddedBounds.Height
            _PixelList = { .X}.Union(Enumerable.Range(1, ValueString.Length).Select(Function(i) TextLength(ValueString.Substring(0, i)))).ToList
        End With

        Dim fieldMatches = RegexMatches(ValueString, "[A-Za-z0-9]{1,}")
        Dim indexField As Integer = If(fieldMatches.Any, fieldMatches(Field).Index, 0)
        _SelectionPixelStart = PixelList(indexField)
        _SelectionPixelEnd = PixelList(indexField + ValueSections(Field).Length)
        'If Value = New Date(2021, 5, 30) Then Stop

        Invalidate()

    End Sub

#Region " Properties & Fields "
    Public Property BorderWidth As Byte = 2
    Public Property BorderColor As Color = Color.Gainsboro
    Public Property HighlightBorderColor As Color = Color.LimeGreen
    Public Property HighlightBorderOnFocus As Boolean
    Friend ReadOnly Property DropDown As MonthCalendarDropDown
    Private ReadOnly Property ValueString As String
        Get
            Return String.Format("{0:" & _Format & "}", {_Value})
        End Get
    End Property
    Private ReadOnly Property ValueSections As List(Of String)
        Get
            Return Regex.Split(ValueString, "[^A-Za-z0-9]+").ToList
        End Get
    End Property
    Private ReadOnly Property FormatSections As List(Of String)
        Get
            Return Regex.Split(Format, "[^A-Za-z0-9]+").ToList
        End Get
    End Property
    Private _Format As String = "dddd, MMM-dd-yyyy"
    Public Property Format As String
        Get
            Return _Format
        End Get
        Set(value As String)
            If _Format <> value Then
                _Format = value
                Bounds_Set()
            End If
        End Set
    End Property
    Public ReadOnly Property ValueIsNull As Boolean
        Get
            Return Value = Date.MinValue
        End Get
    End Property
    Private _Value As Date = Today
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Layout")>
    <Description("Displayed date value - An empty value ( Date.Min ) allows for catching a user change")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Value As Date
        Get
            Return _Value
        End Get
        Set(dateValue As Date)
            If _Value <> dateValue Then
                _Value = dateValue
                If ValueIsNull Then
                    Text = Nothing
                    If HintText IsNot Nothing Then Invalidate()
                Else
                    Bounds_Set()
                    Date_Changed(New DateRangeEventArgs(dateValue, dateValue))
                End If
            End If
        End Set
    End Property
    Public Property HintText As String
    Private HasSearch_ As Boolean = False
    Public Property HasSearch As Boolean
        Get
            Return HasSearch_
        End Get
        Set(value As Boolean)
            If value <> HasSearch_ Then
                HasSearch_ = value
                _SearchItem = MathSymbol.Equals
            End If
        End Set
    End Property
    Private ReadOnly Property MathSymbols As New Dictionary(Of MathSymbol, String()) From
        {
            {MathSymbol.Equals, {"=", "="}},
            {MathSymbol.GreaterThanEquals, {"≥", ">="}},
            {MathSymbol.GreaterThan, {">", ">"}},
            {MathSymbol.LessThan, {"<", "<"}},
            {MathSymbol.LessThanEquals, {"≤", "<="}},
            {MathSymbol.NotEquals, {"≠", "<>"}}
        }
    Public Property SearchItem As MathSymbol
    Private ReadOnly Property SearchDrawString As String
        Get
            Return MathSymbols(SearchItem).First
        End Get
    End Property
    Public ReadOnly Property SearchString As String
        Get
            Return MathSymbols(SearchItem).Last
        End Get
    End Property
    Private _HorizontalAlignment As HorizontalAlignment
    Public Property HorizontalAlignment As HorizontalAlignment
        Get
            Return _HorizontalAlignment
        End Get
        Set(value As HorizontalAlignment)
            If _HorizontalAlignment <> value Then
                _HorizontalAlignment = value
                Bounds_Set()
            End If
        End Set
    End Property
    Public ReadOnly Property Selection As String
        Get
            If SelectionPixelEnd = SelectionPixelStart Or SelectionPixelStart < 0 Or SelectionPixelEnd < 0 Then
                Return String.Empty
            Else
                Return ValueString.Substring(SelectionStart, SelectionEnd - SelectionStart)
            End If
        End Get
    End Property
    Public ReadOnly Property SelectionStart As Integer
        Get
            Return PixelList.IndexOf(SelectionPixelStart)
        End Get
    End Property
    Public ReadOnly Property SelectionEnd As Integer
        Get
            Return PixelList.IndexOf(SelectionPixelEnd)
        End Get
    End Property
    Private _PixelList As New List(Of Integer)
    Public ReadOnly Property PixelList As List(Of Integer)
        Get
            Return _PixelList
        End Get
    End Property
    Private _SelectionPixelStart As Integer
    Private ReadOnly Property SelectionPixelStart As Integer
        Get
            Return {_SelectionPixelStart, _SelectionPixelEnd}.Min
        End Get
    End Property
    Private _SelectionPixelEnd As Integer
    Private ReadOnly Property SelectionPixelEnd As Integer
        Get
            Return {_SelectionPixelStart, _SelectionPixelEnd}.Max
        End Get
    End Property
#End Region
#Region " Events "
    Public Event DateChanged(sender As Object, e As DateRangeEventArgs)
    Public Event DateSubmitted(sender As Object, e As DateRangeEventArgs)
    Public Event TextPasted(sender As Object, e As EventArgs)
    Public Event TextCopied(sender As Object, e As EventArgs)
    Friend Sub Date_Changed(e As DateRangeEventArgs)
        DropDown.SelectionStart = e.Start
        RaiseEvent DateChanged(Me, e)
    End Sub
    Friend Sub DropDownDate_Changed(e As DateRangeEventArgs)
        Value = e.Start
    End Sub
#Region " Overrides "
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        InBounds = True
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        InBounds = False
    End Sub
    Protected Overrides Sub OnPreviewKeyDown(e As PreviewKeyDownEventArgs)

        If e IsNot Nothing Then
            Select Case e.KeyCode
                Case Keys.Left, Keys.Right, Keys.Up, Keys.Down
                    e.IsInputKey = True
            End Select
            MyBase.OnPreviewKeyDown(e)
        End If

    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)

        If e IsNot Nothing AndAlso e.KeyChar.ToString(Globalization.CultureInfo.InvariantCulture).Intersect("0123456789".ToCharArray).Count = 1 Then
            REM Number was Keyed
            If IsNumeric(ValueSections(Field)) Then
                REM Numberic Field
                SB.Append(e.KeyChar)
                Dim nValue As Date
                Dim ParsedOK As Boolean
                Select Case FormatSections(Field).Substring(0, 1).ToUpper(Globalization.CultureInfo.InvariantCulture)
                    Case "D"
                        ParsedOK = Date.TryParse(MonthName(_Value.Month) & "/" & SB.ToString & "/" & _Value.Year, nValue)
                    Case "M"
                        ParsedOK = Date.TryParse(SB.ToString & "/" & _Value.Day & "/" & _Value.Year, nValue)
                    Case "Y"
                        ParsedOK = Date.TryParse(MonthName(_Value.Month) & "/" & _Value.Day & "/" & SB.ToString, nValue)
                End Select
                If Not nValue = Date.MinValue Then
                    Value = nValue
                    Date_Changed(New DateRangeEventArgs(_Value, _Value))
                End If
                If SB.Length = 2 Then SB.Clear()
                Invalidate()
            End If
        End If
        MyBase.OnKeyPress(e)

    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

        Dim S As Integer = SelectionStart
        If e IsNot Nothing Then
            If e.KeyCode = Keys.Left Then
                Field = If(Field = 0, UBound(FormatSections.ToArray), Field - 1)
                Bounds_Set()

            ElseIf e.KeyCode = Keys.C AndAlso Control.ModifierKeys = Keys.Control Then
                Clipboard.Clear()
                Clipboard.SetText(Selection)
                RaiseEvent TextCopied(Me, Nothing)

            ElseIf Control.ModifierKeys = Keys.Control AndAlso e.KeyCode = Keys.V Then
                Dim pastedValue As String = Clipboard.GetText
                Dim dateMatch As Match = Regex.Match(pastedValue, "2[0-9][01][0-9][123][0-9]", RegexOptions.None)
                If dateMatch.Success Then
                    Value = Date.ParseExact(dateMatch.Value, "yyMMdd", InvariantCulture)
                Else
                    '2020-06-29
                    dateMatch = Regex.Match(pastedValue, "20[0-9]{2}-[01][0-9]-[0-3][0-9]", RegexOptions.None)
                    If dateMatch.Success Then
                        Value = Date.ParseExact(dateMatch.Value, "yyyy-MM-dd", InvariantCulture)
                    Else

                    End If
                End If
                RaiseEvent TextPasted(Me, Nothing)

            ElseIf e.KeyCode = Keys.Right Then
                Field = If(Field = UBound(FormatSections.ToArray), 0, Field + 1)
                Bounds_Set()

            ElseIf e.KeyCode = Keys.Enter Then
                RaiseEvent DateSubmitted(Me, New DateRangeEventArgs(Value, Value))

            ElseIf FormatSections(Field).Length > 0 And (e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down) Then
                Select Case FormatSections(Field).Substring(0, 1).ToUpperInvariant
                    Case "D"
                        Value = Value.AddDays(If(e.KeyCode = Keys.Up, If(Value = DateTime.MaxValue, 0, 1), If(Value = DateTime.MinValue, 0, -1)))
                    Case "M"
                        Value = Value.AddMonths(If(e.KeyCode = Keys.Up, If(Value = DateTime.MaxValue, 0, 1), If(Value = DateTime.MinValue, 0, -1)))
                    Case "Y"
                        Value = Value.AddYears(If(e.KeyCode = Keys.Up, If(Value = DateTime.MaxValue, 0, 1), If(Value = DateTime.MinValue, 0, -1)))
                End Select

            ElseIf e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete Then
                If Selection.Length = 0 Then
                    If e.KeyCode = Keys.Back Then
                        If Not S = 0 Then
                            Text = Text.Remove(S - 1, 1)
                            _SelectionPixelStart = TextLength(Text.Substring(0, S - 1))
                        End If
                    ElseIf e.KeyCode = Keys.Delete Then
                        If Not S = Text.Length Then
                            Text = Text.Remove(S, 1)
                            _SelectionPixelStart = TextLength(Text.Substring(0, S))
                        End If
                    End If
                Else
                    Text = Text.Substring(0, S) & String.Empty & Text.Substring(SelectionEnd, Text.Length - SelectionEnd)
                    _SelectionPixelStart = TextLength(Text.Substring(0, S))
                End If
                _SelectionPixelEnd = SelectionPixelStart

            End If
            Invalidate()
            MyBase.OnKeyDown(e)
        End If

    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        Bounds_Set()
        MyBase.OnSizeChanged(e)
    End Sub
    Protected Overrides Sub OnFontChanged(e As EventArgs)
        Bounds_Set()
        DropDown.Font = Font
        MyBase.OnFontChanged(e)
    End Sub
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)
        Toolstrip.Size = New Size(0, 0)
        DropDown.Visible = False
        Field = 0
        MyBase.OnVisibleChanged(e)
    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            Dim lastRegion = MouseOver
            If SearchBounds.Contains(e.Location) Then
                MouseOver = MouseRegion.Search

            ElseIf DropDrawBounds.Contains(e.Location) Then
                MouseOver = MouseRegion.Drop

            ElseIf ClearTextDrawBounds.Contains(e.Location) Then
                MouseOver = MouseRegion.Clear

            Else
                If e.X >= PixelList.First And e.X <= PixelList.Last Then
                    If Not ValueIsNull Then
                        Dim Position As Integer = (From I In PixelList Where e.X <= I).First
                        Try
                            Dim mouseMatch = Regex.Match(ValueString(PixelList.Skip(1).ToList.IndexOf(Position)), "[^A-Za-z0-9]+")
                            If mouseMatch.Length = 0 Then
                                Field = Regex.Replace(ValueString.Substring(0, PixelList.IndexOf(Position)), "[A-Za-z0-9]|[ ]", String.Empty).Length
                                MouseOver = If(Field = 0, MouseRegion.WeekDay, If(Field = 1, MouseRegion.Month, If(Field = 2, MouseRegion.Day, MouseRegion.Year)))
                                Bounds_Set()
                            Else
                                MouseOver = MouseRegion.None
                            End If
                        Catch ex As IndexOutOfRangeException
                            MouseOver = MouseRegion.None
                        End Try
                    End If
                End If

            End If
            If lastRegion <> MouseOver Then Invalidate()
            MyBase.OnMouseMove(e)
        End If

    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

        If e IsNot Nothing Then
            If MouseOver = MouseRegion.Clear Then
                Value = Date.MinValue

            ElseIf MouseOver = MouseRegion.Drop Then
                If Not DropDown.Visible Then
                    Dim Coordinates As Point
                    Coordinates = PointToScreen(New Point(0, 0))
                    Toolstrip.Show(Coordinates.X + Width - DropDown.Width, If(Coordinates.Y + DropDown.Height > My.Computer.Screen.WorkingArea.Height, Coordinates.Y - DropDown.Height, Coordinates.Y + Height))
                End If
                DropDown.Visible = Not DropDown.Visible

            ElseIf MouseOver = MouseRegion.Search Then
                '0 1 2 3
                '= > < ≠
                Dim nextIndex As Integer = MathSymbols.Keys.ToList.IndexOf(SearchItem)
                nextIndex = If(nextIndex + 1 = MathSymbols.Count, 0, nextIndex + 1)
                _SearchItem = MathSymbols.Keys.ToList(nextIndex)

            End If
            SB.Clear()
            Invalidate()
            MyBase.OnMouseDown(e)
        End If

    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        Invalidate()
        MyBase.OnMouseUp(e)
    End Sub
#End Region
#End Region

    Private Function TextLength(T As String) As Integer
        Dim spacing As Integer = If(T.Length = 0, 0, (2 * TextRenderer.MeasureText(T.First, Font).Width) - TextRenderer.MeasureText(T.First & T.First, Font).Width)
        Return DateBounds.Left + TextRenderer.MeasureText(T, Font).Width - spacing
    End Function

    Friend NotInheritable Class MonthCalendarDropDown
        Inherits MonthCalendar
        Public Sub New(DatePicker As DatePicker)

            Application.VisualStyleState = VisualStyles.VisualStyleState.NoneEnabled
            Visible = False
            Margin = New Padding(0)
            MaxSelectionCount = 1
            Parent = DatePicker
            Font = DatePicker.Font
            BackColor = SystemColors.Window
            AddHandler Parent.FontChanged, AddressOf Parent_FontChanged

        End Sub

        Private Shadows ReadOnly Parent As DatePicker
        Private Const ShadowDepth As Integer = 5

        Public Property DropShadowColor As Color = Color.Red

        Protected _ForceCapture As Boolean
        Protected Property ForceCapture() As Boolean
            Get
                Return _ForceCapture
            End Get
            Set(value As Boolean)
                _ForceCapture = value
                Capture = value
            End Set
        End Property
        Protected Overrides Sub OnDateSelected(e As DateRangeEventArgs)
            Parent.DropDownDate_Changed(e)
            Parent.Invalidate()
            Visible = False
            MyBase.OnDateSelected(e)
        End Sub
        Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
            If Not Bounds.Contains(e.Location) Then
                Visible = False
            End If
            MyBase.OnMouseDown(e)
        End Sub
        Protected Overrides Sub OnMouseCaptureChanged(e As EventArgs)
            MyBase.OnMouseCaptureChanged(e)
            Capture = _ForceCapture And Visible
        End Sub
        Protected Overrides Sub OnVisibleChanged(e As EventArgs)
            If Visible Then
                ForceCapture = True
                SelectionStart = Parent.Value
                Parent.Toolstrip.Size = Size
                Top = 0
                Dim DisplayFactor = DisplayScale()
                Dim bmp As New Bitmap(CInt((Width + ShadowDepth) * DisplayFactor), CInt((Height + ShadowDepth) * DisplayFactor))
                Using Graphics As Graphics = Graphics.FromImage(bmp)
                    Dim Point As Point = PointToScreen(New Point(0, 0))
                    Graphics.CopyFromScreen(
                        CInt(Point.X * DisplayFactor),
                        CInt(Point.Y * DisplayFactor),
                        0,
                        0,
                        bmp.Size,
                        CopyPixelOperation.SourceCopy)
                    For P = 0 To ShadowDepth - 1
                        Using Brush As New SolidBrush(Color.FromArgb(16 + (P * 5), DropShadowColor))
                            Graphics.FillRectangle(Brush, New Rectangle(ShadowDepth + P, ShadowDepth + P, Width - ShadowDepth - P * 2, Height - ShadowDepth - P * 2))
                        End Using
                    Next
                End Using
                Parent.Toolstrip.BackgroundImage = bmp
            Else
                If Not IsNothing(Parent) Then Parent.Toolstrip.Size = New Size(0, 0)
                ForceCapture = False
            End If
            MyBase.OnVisibleChanged(e)
        End Sub
        Private Sub Parent_FontChanged(sender As Object, e As EventArgs)
            Font = Parent.Font
        End Sub
    End Class
End Class