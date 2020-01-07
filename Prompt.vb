Option Strict On
Option Explicit On
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Windows.Forms
Public Class Prompt
    Inherits Form
    Private Const WM_NCACTIVATE As Integer = &H86
    Private Const WM_NCPAINT As Integer = &H85
    Private Const ButtonBarHeight As Integer = 36
    Private WithEvents Table As New DataViewer With {.Font = PreferredFont, .Visible = True, .Dock = DockStyle.Fill}
    Private WithEvents OK As New Button With {.Text = "OK", .Font = PreferredFont, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.ButtonYes, .BackColor = Color.GhostWhite, .ForeColor = Color.Black, .FlatStyle = FlatStyle.Popup}
    Private WithEvents YES As New Button With {.Text = "Yes", .Font = PreferredFont, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.ButtonYes, .BackColor = Color.GhostWhite, .ForeColor = Color.Black, .FlatStyle = FlatStyle.Popup}
    Private WithEvents NO As New Button With {.Text = "No", .Font = PreferredFont, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.ButtonNo, .BackColor = Color.GhostWhite, .ForeColor = Color.Black, .FlatStyle = FlatStyle.Popup}
    Private WithEvents PromptTimer As New Timer With {.Interval = 5000}
    Private WorkingSpace As Rectangle = Screen.PrimaryScreen.Bounds
    Public Enum IconOption

        Critical
        OK
        TimedMessage
        Warning
        YesNo

    End Enum
    Public Enum StyleOption
        Plain
        BlackGold
        Blue
        Bright
        Grey
        RedBrown
        Earth
        Psychedelic
        Custom
    End Enum
    Public Sub New()

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)
        BackColor = SystemColors.ControlLight
        FormBorderStyle = FormBorderStyle.Fixed3D
        Font = PreferredFont
        TopMost = True
        MinimizeBox = False
        MaximizeBox = False
        MinimumSize = New Size(300, 3 + 112 + 3)     'IconHeight + Padding
        MaximumSize = New Size(Convert.ToInt32(0.7 * WorkingSpace.Width), Convert.ToInt32(0.7 * WorkingSpace.Height))
        Controls.Add(Table)
        Controls.AddRange({OK, YES, NO})
        KeyPreview = True

    End Sub
#Region " PROPERTIES "
    Private ReadOnly Property PreferredFont As New Font("Segoe UI", 9)
    Private PathColors As New List(Of Color)({Color.Chocolate, Color.SaddleBrown, Color.Peru})
    Private _DataSource As Object
    Public Property Datasource As Object
        Set(value As Object)
            _DataSource = value
            Table.DataSource = value
        End Set
        Get
            Return _DataSource
        End Get
    End Property
    Public Property ColorStyle As StyleOption = StyleOption.Plain
    Private ReadOnly Property AlternatingRowColor As Color
    Private ReadOnly Property BackgroundColor As Color
    Private ReadOnly Property TextColor As Color
    Private ReadOnly Property ShadeColor As Color
    Private ReadOnly Property AccentColor As Color
    Public Property TitleMessage As String
    Public Property TitleBarImageLeftSide As Boolean = True
    Public Property TitleBarImage As Image = My.Resources.Info_White
    Public Property BodyMessage As String
    Public Property BorderColor As Color = Color.Black
    Public Property BorderForeColor As Color = Color.White
    Private Icon_ As Icon = Nothing
    Public Overloads Property Icon As Icon
        Get
            If Icon_ Is Nothing Then
                Select Case Type
                    Case IconOption.Critical
                        Return My.Resources._Error

                    Case IconOption.OK
                        Return My.Resources.Check

                    Case IconOption.TimedMessage
                        Return My.Resources.Clock

                    Case IconOption.Warning
                        Return My.Resources.Warning_

                    Case IconOption.YesNo
                        Return My.Resources.Question

                    Case Else
                        Return SystemIcons.Shield

                End Select
            Else
                Return Icon_
            End If
        End Get
        Set(value As Icon)
            Icon_ = value
        End Set
    End Property
    Public Property Type As IconOption = IconOption.OK
    Private ReadOnly Property SideBorderWidths As Integer
        Get
            Dim ScreenRectangle As Rectangle = RectangleToScreen(ClientRectangle)
            Return (ScreenRectangle.Left - Left())
        End Get
    End Property
    Private ReadOnly Property BottomBorderHeight As Integer
        Get
            Dim ScreenRectangle As Rectangle = RectangleToScreen(ClientRectangle)
            Return (Bottom - ScreenRectangle.Bottom)
        End Get
    End Property
    Private ReadOnly Property TitleBarHeight As Integer
        Get
            Dim ScreenRectangle As Rectangle = RectangleToScreen(ClientRectangle)
            Return (ScreenRectangle.Top - Top)
        End Get
    End Property
    Private ReadOnly Property TitleBarBounds As Rectangle
        Get
            Return New Rectangle(0, 0, Width, TitleBarHeight)
        End Get
    End Property
    Private Const IconPadding As Integer = 3
    Private ReadOnly Property IconBounds As New Rectangle(IconPadding, IconPadding, Icon.Width, Icon.Height)
    Private ReadOnly Property TextBounds As New Dictionary(Of Rectangle, String)
    Private ReadOnly Property GridBounds As Rectangle
        Get
            Dim GridTop As Integer = {If(TextBounds.Keys.Any, TextBounds.Keys.Last.Bottom, 10), IconBounds.Bottom + 16}.Max
            If IsNothing(Table.DataSource) Then
                Table.Visible = False
                Table.Size = New Size(1, 1)
                Return New Rectangle(0, GridTop, 0, 0)
            Else
                Table.Width = Table.Columns.Sum(Function(x) x.Width + 1)
                Table.Visible = True
                Table.Height = {3 + Table.Columns.HeadBounds.Height + (Table.Rows.RowHeight * {1, Table.Rows.Count}.Max), 360}.Min
                Return New Rectangle(0, 4 + GridTop, Table.Width, Table.Height)
            End If
        End Get
    End Property
    Private ReadOnly Property ButtonBarBounds As Rectangle
    Private ReadOnly Property MainWindow As Process
    Private ReadOnly AddressBounds As New Dictionary(Of Rectangle, String)
    Private LastBounds As New Rectangle
#End Region

#Region " PAINT "
    Protected Overrides Sub WndProc(ByRef m As Message)

        If m.Msg = WM_NCPAINT Then
            DrawTitleBar(True)
            Invalidate()
        Else
            MyBase.WndProc(m)
            If m.Msg = WM_NCACTIVATE Then
                DrawTitleBar(False)
                Invalidate()
            End If
        End If

    End Sub
    Private Sub DrawTitleBar(ByVal DrawForm As Boolean)

        Dim hdc As IntPtr = NativeMethods.GetWindowDC(Handle)
        Using g As Graphics = Graphics.FromHdc(hdc)
            g.SmoothingMode = SmoothingMode.AntiAlias
            If DrawForm Then
                Using BC As New SolidBrush(BorderColor)
                    g.FillRectangle(BC, New Rectangle(0, 0, Width, Height))
                End Using
                Using bm As New Bitmap(Width, Height)
                    DrawToBitmap(bm, New Rectangle(0, 0, Width, Height))
                    g.DrawImage(bm, 0, 0, Width, Height)
                End Using
            End If
            Using BC As New SolidBrush(BorderColor)     'Color.FromArgb(32, 32, 32)
                g.FillRectangle(BC, New Rectangle(0, 0, Width, Height))
            End Using

            If If(TitleMessage, String.Empty).Any Then
                Dim horizontalPadding As Integer = 2
                Dim MaxImageHeight As Integer = TitleBarBounds.Height - 4
                Dim ImageHeight As Integer = If(TitleBarImage.Height > MaxImageHeight, MaxImageHeight, TitleBarImage.Height)
                Dim ImageWidth As Integer = ImageHeight     'Default SQUARE

                If TitleBarImage.Width = TitleBarImage.Height Then
                    'Square, so no fancy calcs

                Else
                    Dim TextSize As Size = MeasureText(TitleMessage, PreferredFont)
                    Dim MaxImageWidth As Integer = TitleBarBounds.Width - (horizontalPadding + TextSize.Width + horizontalPadding)
                    ImageWidth = If(TitleBarImage.Width > MaxImageWidth, MaxImageWidth, TitleBarImage.Width)

                End If

                Dim yOffset As Integer = Convert.ToInt32((TitleBarBounds.Height - ImageHeight) / 2)
                Dim ImageBounds As Rectangle = Nothing
                Dim TextBounds As Rectangle = Nothing

                If TitleBarImageLeftSide Then
                    ImageBounds = New Rectangle(horizontalPadding, yOffset, ImageWidth, ImageHeight)
                    TextBounds = New Rectangle(ImageWidth + horizontalPadding, 0, Width - (ImageWidth + horizontalPadding), TitleBarBounds.Height)
                Else
                    TextBounds = New Rectangle(horizontalPadding, 0, Width - (ImageWidth + horizontalPadding), TitleBarBounds.Height)
                    ImageBounds = New Rectangle(TextBounds.Width + horizontalPadding, yOffset, ImageWidth, ImageHeight)
                End If
                g.DrawImage(TitleBarImage, ImageBounds)
                TextRenderer.DrawText(g, TitleMessage, PreferredFont, TextBounds, Color.White, BorderColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)
            End If

        End Using
        Dim Result = NativeMethods.ReleaseDC(Handle, hdc)

    End Sub
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If e IsNot Nothing Then
            If IsNothing(BackgroundImage) Then
                REM /// FILL BACKGROUND
                Using GP As New GraphicsPath
                    Dim Points(5) As Point
                    REM Below is a Square, but could be changed to calculate points for GP and better effect
                    Points(0) = New Point(0, TitleBarBounds.Bottom)
                    Points(1) = New Point(0, ClientSize.Height)
                    Points(2) = New Point(ClientSize.Width, ClientSize.Height)
                    Points(3) = New Point(ClientSize.Width, 0)
                    GP.AddPolygon(Points)
                    Using PathBrush As New PathGradientBrush(GP)
                        PathBrush.SurroundColors = PathColors.Take(2).ToArray
                        PathBrush.CenterColor = PathColors.Last
                        e.Graphics.FillPath(PathBrush, GP)
                    End Using
                End Using

            Else
                e.Graphics.DrawImage(BackgroundImage, New Point(0, 0))

            End If

            REM /// DRAW ICON IN THE UPPER LEFT CORNER
            'e.Graphics.FillRectangle(Brushes.Gainsboro, IconBounds)
            e.Graphics.DrawIcon(Icon, IconBounds)

            REM /// DRAW TEXT IN EACH RECTANGLE
            AddressBounds.Clear()
            For Each TextBound In TextBounds.Keys
                Dim Line = TextBounds(TextBound)
                If Regex.Match(Line, "https{0,1}:[^ ]{1,}", RegexOptions.None).Success Then     'Line hass URL
                    Dim Words = RegexMatches(Line, "[^\s]{1,}", RegexOptions.None)
                    Dim WordRectangles As New List(Of Rectangle)
                    Dim LastRight As Integer = TextBound.Left
                    For Each Word In Words
                        Dim WordWidth As Integer = MeasureText(Word.Value, PreferredFont).Width
                        Dim WordRectangle As New Rectangle(LastRight, TextBound.Top, WordWidth, TextBound.Height)
                        WordRectangles.Add(WordRectangle)
                        LastRight = WordRectangle.Right
                        Dim WordIsAddress As Boolean = Regex.Match(Word.Value, "https{0,1}:[^ ]{1,}", RegexOptions.None).Success
                        'e.Graphics.DrawRectangle(Pens.Red, WordRectangle)
                        If WordIsAddress Then
                            AddressBounds.Add(WordRectangle, Word.Value)
                            If WordRectangle = LastBounds Then
                                Using Underline As New Font(PreferredFont.FontFamily, PreferredFont.Size, FontStyle.Underline)
                                    TextRenderer.DrawText(e.Graphics, Word.Value, Underline, WordRectangle, Color.Blue, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                                End Using
                            Else
                                TextRenderer.DrawText(e.Graphics, Word.Value, PreferredFont, WordRectangle, Color.Blue, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                            End If
                        Else
                            TextRenderer.DrawText(e.Graphics, Word.Value, PreferredFont, WordRectangle, TextColor, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                        End If
                    Next

                Else
                    TextRenderer.DrawText(e.Graphics, Line, PreferredFont, TextBound, Color.Black, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)

                End If
            Next
            If Not Type = IconOption.TimedMessage Then
                Using ButtonBarBrush As New SolidBrush(Color.FromArgb(32, BackgroundColor))
                    e.Graphics.FillRectangle(ButtonBarBrush, ButtonBarBounds)
                    ControlPaint.DrawBorder3D(e.Graphics, ButtonBarBounds)
                End Using
            End If
        End If

    End Sub
#End Region
#Region " EVENTS "
    Private Sub Button_Click(ByVal sender As Object, ByVal e As EventArgs) Handles OK.Click, YES.Click, NO.Click

        Select Case True
            Case sender Is OK
                DialogResult = DialogResult.OK

            Case sender Is YES
                DialogResult = DialogResult.Yes

            Case sender Is NO
                DialogResult = DialogResult.No

        End Select
        Close()

    End Sub
    Private Sub PromptTimer_Tick(sender As Object, e As EventArgs) Handles PromptTimer.Tick

        PromptTimer.Stop()
        DialogResult = DialogResult.None
        Hide()

    End Sub
    Private Sub Message_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Table.DataSource = Nothing
        TitleMessage = String.Empty
        BodyMessage = String.Empty
        Dim result As Integer = NativeMethods.SetForegroundWindow(MainWindow.Handle)

    End Sub
    Protected Overrides Sub OnFontChanged(e As EventArgs)

        Table.Font = Font
        YES.Font = Font
        NO.Font = Font
        _PreferredFont = Font
        MyBase.OnFontChanged(e)

    End Sub
    Private Sub Keyed(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.C AndAlso Control.ModifierKeys = Keys.Control Then
            Clipboard.Clear()
            Clipboard.SetText(BodyMessage)
        End If

    End Sub
    Private Sub MouseOver(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        If LastBounds.Contains(e.Location) Then
        Else
            LastBounds = Nothing
            Cursor = Cursors.Default
            Invalidate()
            For Each address In AddressBounds.Keys
                If address.Contains(e.Location) Then
                    LastBounds = address
                    Cursor = Cursors.Hand
                    Invalidate()
                End If
            Next
        End If

    End Sub
    Private Sub MouseClicked(sender As Object, e As MouseEventArgs) Handles Me.MouseClick

        If LastBounds.Contains(e.Location) Then
            'Dim Programs = GetFiles("C:\Program Files", ".exe")
            'Stop
            Process.Start("C:\Program Files\Internet Explorer\IExplore.exe", AddressBounds(LastBounds))
        End If

    End Sub
#End Region

    Public Overloads Function Show(BodyMessage As String, Optional Type As IconOption = IconOption.OK, Optional ColorTheme As StyleOption = StyleOption.Plain, Optional AutoCloseSeconds As Integer = 3) As DialogResult
        Return Show(String.Empty, If(BodyMessage, String.Empty), Type, ColorTheme, AutoCloseSeconds)
    End Function
    Public Overloads Function Show(TitleMessage As String, BodyMessage As List(Of String), Optional Type As IconOption = IconOption.OK, Optional ColorTheme As StyleOption = StyleOption.Plain, Optional AutoCloseSeconds As Integer = 3) As DialogResult

        If BodyMessage Is Nothing Then
            Return Show(TitleMessage, String.Empty, Type, ColorTheme, AutoCloseSeconds)
        Else
            Return Show(TitleMessage, Join(BodyMessage.ToArray, vbNewLine), Type, ColorTheme, AutoCloseSeconds)
        End If

    End Function
    Public Overloads Function Show(TitleMessage As String,
                                    BodyMessage As String,
                                    Type As IconOption,
                                    AlternatingRowColor As Color,
                                    BackgroundColor As Color,
                                    TextColor As Color,
                                    ShadeColor As Color,
                                    AccentColor As Color,
                                    BorderColor As Color) As DialogResult
        _AlternatingRowColor = AlternatingRowColor
        _BackgroundColor = BackgroundColor
        _TextColor = TextColor
        _ShadeColor = ShadeColor
        _AccentColor = AccentColor
        _BorderColor = BorderColor

        ColorStyle = StyleOption.Custom
        Return Show(TitleMessage, BodyMessage, Type, StyleOption.Custom)

    End Function
    Public Overloads Function Show(TitleMessage As String, BodyMessage As String(), Optional Type As IconOption = IconOption.OK, Optional ColorTheme As StyleOption = StyleOption.Plain, Optional AutoCloseSeconds As Integer = 3) As DialogResult
        Return Show(TitleMessage, Join(BodyMessage, vbNewLine), Type, ColorTheme, AutoCloseSeconds)
    End Function
    Public Overloads Function Show(TitleMessage As String, BodyMessage As String, Optional Type As IconOption = IconOption.OK, Optional ColorTheme As StyleOption = StyleOption.Plain, Optional AutoCloseSeconds As Integer = 3) As DialogResult

        _MainWindow = Process.GetCurrentProcess
        Dim ProcessList As New List(Of Process)(Process.GetProcesses)
        ProcessList.Sort(Function(p1, p2)
                             Dim level1 = String.Compare(p1.MainWindowTitle, p2.MainWindowTitle, StringComparison.InvariantCulture)
                             Return level1
                         End Function)

        ControlBox = False
        Me.TitleMessage = TitleMessage
        Text = TitleMessage
        PromptTimer.Interval = 1000 * AutoCloseSeconds

        BodyMessage = If(BodyMessage, String.Empty)
        Me.BodyMessage = BodyMessage
        If BodyMessage.Length = 0 Then Me.BodyMessage = "No Message"

        Me.Type = Type
        ColorStyle = ColorTheme

        Select Case ColorTheme
            Case StyleOption.BlackGold
                _AlternatingRowColor = Color.Gold
                _BackgroundColor = Color.Black
                _TextColor = Color.White
                _ShadeColor = Color.DarkKhaki
                _AccentColor = Color.DarkGoldenrod
                BorderColor = Color.Black

            Case StyleOption.Blue
                _AlternatingRowColor = Color.LightSkyBlue
                _BackgroundColor = Color.CornflowerBlue
                _TextColor = Color.White
                _ShadeColor = Color.DarkBlue
                _AccentColor = Color.DarkSlateBlue
                BorderColor = Color.RoyalBlue

            Case StyleOption.Bright
                _AlternatingRowColor = Color.Gold
                _BackgroundColor = Color.HotPink
                _TextColor = Color.White
                _ShadeColor = Color.Fuchsia
                _AccentColor = Color.DarkOrchid
                BorderColor = Color.DarkMagenta

            Case StyleOption.Grey
                _AlternatingRowColor = Color.DarkGray
                _BackgroundColor = Color.Gainsboro
                _TextColor = Color.Black
                _ShadeColor = Color.Silver
                _AccentColor = Color.Gray
                BorderColor = Color.Black

            Case StyleOption.Earth
                _AlternatingRowColor = Color.Beige
                _BackgroundColor = Color.Green
                _TextColor = Color.White
                _ShadeColor = Color.DarkGreen
                _AccentColor = Color.DarkOliveGreen
                BorderColor = Color.Maroon

            Case StyleOption.Psychedelic
                _AlternatingRowColor = Color.Lavender
                _BackgroundColor = Color.Fuchsia
                BackgroundImageLayout = ImageLayout.Stretch
                _TextColor = Color.White
                _ShadeColor = Color.Gainsboro
                _AccentColor = Color.DarkOrange
                BorderColor = Color.YellowGreen

            Case StyleOption.Plain
                _AlternatingRowColor = Color.Gainsboro
                _BackgroundColor = Color.LightGray
                _TextColor = Color.Black
                _ShadeColor = Color.DarkGray
                _AccentColor = Color.Gainsboro
                BorderColor = Color.Silver

            Case StyleOption.RedBrown
                _AlternatingRowColor = Color.Chocolate
                _BackgroundColor = Color.Orange
                _TextColor = Color.White
                _ShadeColor = Color.Crimson
                _AccentColor = Color.Peru
                BorderColor = Color.SaddleBrown

            Case StyleOption.Custom

        End Select

        With Table
            .Columns.HeaderStyle = New CellStyle With {.BackColor = ShadeColor, .ShadeColor = ShadeColor, .ForeColor = TextColor}
            .Rows.AlternatingRowStyle = New CellStyle With {.BackColor = AlternatingRowColor, .ForeColor = Color.Black}
            .Rows.RowStyle = New CellStyle With {.BackColor = Color.GhostWhite, .ForeColor = Color.Black}
        End With
        For Each InputButton As Button In {YES, NO, OK}
            InputButton.ForeColor = TextColor
            InputButton.BackColor = BorderColor
        Next

        ForeColor = TextColor
        PathColors = {Color.Gray, BackColor, ShadeColor, AccentColor}.ToList
        ResizeMe()
        Hide()
        ShowDialog()

        Return DialogResult

    End Function
    Private Sub ResizeMe()

        TextBounds.Clear()
        Dim RowBWidth As Integer = 0
        Dim IconZoneWH As Integer = IconPadding * 2 + Icon.Height

#Region " #1 - Get TextWidth "
        'a) Width & Height as a Rectangle based on total area...if there are long words wider than the Rectangle, it will need expanding below =============
        Dim TextSize As Size = MeasureText(BodyMessage, Font)
        Dim TextArea As Integer = TextSize.Width * TextSize.Height

        Dim x2yRatio As Double = 3
        'x * y = TextArea, x is x2yRatio larger than y ∴ ( y * x2yRatio ) * y = TextArea ∴ y² = TextArea ÷ x2yRatio ∴ y=√ ( TextArea ÷ x2yRatio )
        Dim y As Double = Math.Sqrt(TextArea / x2yRatio)
        RowBWidth = CInt(y * x2yRatio)
        '==============================
        'Ensure extra long words are considered
        Dim Words As New List(Of Integer)(From rm In RegexMatches(BodyMessage, "[^ ]{1,}", RegexOptions.None) Select MeasureText(rm.Value, Font).Width)
        If Words.Max > RowBWidth Then
            'If a word is wider than the derived width the expand the width but if the word appears in the RowsA section, then add IconZoneWH as this value is subtracted below: Dim LinesA = WrapWords(BodyMessage, Font, RowBWidth - *** IconZoneWH *** )
            'Dim LinesA = WrapWords(BodyMessage, Font, RowBWidth - IconZoneWH) makes an empty String if too long ... so remove empty strings - easy fix
        End If
#End Region

        Dim Attempts As Integer
        Dim RowsABHeight As Integer = 0
        Do
            RowBWidth = {RowBWidth, SideBorderWidths + Words.Max + SideBorderWidths, MinimumSize.Width}.Max
#Region " #1 - Get TextHeight "
            Dim RowsA As New Dictionary(Of Rectangle, String)
            Dim RowsB As New Dictionary(Of Rectangle, String)
            Dim LinesA = WrapWords(BodyMessage, Font, RowBWidth - IconZoneWH)
            Dim LinesB As New Dictionary(Of Integer, String)
            Dim LineHeight As Integer = MeasureText("|".ToUpperInvariant, Font).Height
            Dim LineIndex As Integer = 0
            Dim WidenText As Boolean = False

            For Each Line In LinesA.Where(Function(l) l.Value.Any)
                If LineIndex * LineHeight >= IconZoneWH Then
                    Dim RemainingLines = Join(LinesA.Values.Skip(LineIndex).ToArray, vbNewLine)
                    LinesB = WrapWords(RemainingLines, Font, RowBWidth)
                    Exit For

                Else
                    Dim TopRightIconPoint As New Point(IconZoneWH, LineIndex * LineHeight)
                    RowsA.Add(New Rectangle(TopRightIconPoint, New Size(If(WidenText, SideBorderWidths, 0) + RowBWidth - TopRightIconPoint.X, LineHeight)), Line.Value)
                    TextBounds.Add(RowsA.Last.Key, RowsA.Last.Value)

                End If
                LineIndex += 1
            Next
            For Each Line In LinesB
                RowsB.Add(New Rectangle(New Point(0, LineIndex * LineHeight), New Size(If(WidenText, SideBorderWidths, 0) + RowBWidth, LineHeight)), Line.Value)
                TextBounds.Add(RowsB.Last.Key, RowsB.Last.Value)
                LineIndex += 1
            Next
            RowsABHeight = (RowsA.Count + RowsB.Count) * LineHeight
#End Region
            If RowsA.Any Then
                Dim RowsABottom As Integer = RowsA.Last.Key.Bottom
                If RowsABottom > IconZoneWH Then
                    'Center Icon between Text Rows
                    Dim yOffset = CInt((RowsABottom - Icon.Height) / 2)
                    _IconBounds = New Rectangle(IconPadding, yOffset, Icon.Width, Icon.Height)

                ElseIf IconZoneWH > RowsABottom And Not RowsB.Any Then
                    'Center Text between IconZone
                    Dim RowOffset As Integer = CInt((IconZoneWH - RowsABottom) / 2)
                    TextBounds.Clear()
                    For Each Row In RowsA
                        TextBounds.Add(New Rectangle(Row.Key.X, Row.Key.Y + RowOffset, Row.Key.Width, Row.Key.Height), Row.Value)
                    Next
                    _IconBounds = New Rectangle(IconPadding, IconPadding, Icon.Width, Icon.Height)

                Else
                    'Text height in top rows matches the height of the Icon
                    _IconBounds = New Rectangle(IconPadding, IconPadding, Icon.Width, Icon.Height)
                End If
            End If
            RowBWidth += CInt(Words.Average)
            Attempts += 1
        Loop While attempts < 1 Or RowBWidth / {RowsABHeight, 1}.max < 2

        Width = (SideBorderWidths * 2) + {TextBounds.Max(Function(x) x.Key.Right), GridBounds.Right}.Max
        Dim ButtonBarTop As Integer = IconPadding + {TextBounds.Keys.Last.Bottom, IconBounds.Bottom}.Max

#Region " GRID WIDTH / HEIGHT / PLACEMENT  / VISIBILITY "
        If Table.DataSource IsNot Nothing Then
            Dim TLP_Table As New TableLayoutPanel With {.ColumnCount = 1, .RowCount = 1, .CellBorderStyle = TableLayoutPanelCellBorderStyle.None, .BorderStyle = BorderStyle.None}
            With TLP_Table
                .Left = 0
                Dim GridWidth As Integer = Width - (SideBorderWidths * 2)
                .Width = GridWidth
                .Top = GridBounds.Top
                Dim GridHeight As Integer = Table.Columns.HeadBounds.Height + Table.Rows.Take(15).Count * Table.Rows.RowHeight      ' 15 Rows MAXIMUM
                .Height = GridHeight
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = GridWidth})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = GridHeight})
                .Controls.Add(Table)
                Table.Columns.DistibuteWidths()
            End With
            Controls.Add(TLP_Table)
            ButtonBarTop = TLP_Table.Bottom
        End If
#End Region

        _ButtonBarBounds = New Rectangle(0, ButtonBarTop, ClientSize.Width - 1, If(Type = IconOption.TimedMessage, 0, ButtonBarHeight))
        Height = TitleBarHeight + ButtonBarBounds.Bottom + BottomBorderHeight

#Region " BUTTON PLACEMENT / VISIBILITY "
        For Each Button In {YES, NO, OK}
            Button.Visible = False
            Button.Top = ButtonBarBounds.Top + Convert.ToInt32((ButtonBarBounds.Height - Button.Height) / 2)
        Next
        Select Case Type
            Case IconOption.Critical, IconOption.OK, IconOption.Warning
                OK.Visible = True
                OK.Left = Convert.ToInt32((ClientSize.Width - OK.Width) / 2)

            Case IconOption.TimedMessage
                PromptTimer.Start()

            Case IconOption.YesNo
                NO.Visible = True
                YES.Visible = True
                Dim ButtonSpacing As Integer = Convert.ToInt32((ClientSize.Width - YES.Width - NO.Width) / 3)
                YES.Left = ButtonSpacing
                NO.Left = YES.Bounds.Right + ButtonSpacing

        End Select
#End Region

        Invalidate()
        CenterToScreen()

    End Sub
End Class
Public Class TitleBarImage
    Inherits Form
    Private ReadOnly _frm As Form = Nothing
    Private ReadOnly _TitleBarAlignment As HorizontalAlignment = Nothing
    Private PositionOffset As Point = Point.Empty
    Private ReadOnly _Img As Image = Nothing

    ''' <summary>Displays an image in the TitleBar of the specified form.</summary>
    ''' <param name="frm">The form to display the image on.</param>
    ''' <param name="Img">The Image to display.</param>
    ''' <param name="Alignment">Aligns the image to the Left, Center, or Right.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal frm As Form, ByVal Img As Image, ByVal Alignment As HorizontalAlignment)
        StartPosition = FormStartPosition.Manual
        FormBorderStyle = FormBorderStyle.None
        Opacity = 0.0
        ShowInTaskbar = False
        BackColor = Color.Lime
        TransparencyKey = Color.Lime
        Owner = frm
        _TitleBarAlignment = Alignment
        _frm = frm
        _Img = Img
        SetHandlers()
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        If e IsNot Nothing Then
            e.Graphics.DrawImage(_Img, 0, 0, Width, Height)
            MyBase.OnPaint(e)
        End If
    End Sub

    Protected Overrides Sub OnShown(ByVal e As System.EventArgs)
        Dim ratio As Double = _Img.Width / _Img.Height
        Height = SystemInformation.CaptionHeight - 4
        Width = CInt(Height * ratio)
        SetPosition()
        Opacity = 1.0
        MyBase.OnShown(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        PositionOffset.X = MousePosition.X - _frm.Left
        PositionOffset.Y = MousePosition.Y - _frm.Top
        MyBase.OnMouseDown(e)
        _frm.Focus()
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        If MouseButtons = MouseButtons.Left Then
            _frm.Left = MousePosition.X - PositionOffset.X
            _frm.Top = MousePosition.Y - PositionOffset.Y
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Private Sub SetPosition()
        If _TitleBarAlignment = HorizontalAlignment.Left Then
            Left = _frm.Left + 10
            Top = _frm.Top + 4
        ElseIf _TitleBarAlignment = HorizontalAlignment.Center Then
            Left = _frm.Left + CInt((_frm.Width / 2) - (Width / 2))
            Top = _frm.Top + 4
        Else
            Left = _frm.Right - Width - 10
            Top = _frm.Top + 4
        End If
    End Sub

    Private Sub SetHandlers()
        AddHandler _frm.FormClosing, AddressOf Frm_Closing
        AddHandler _frm.Move, AddressOf Frm_Move
        AddHandler _frm.Resize, AddressOf Frm_Resize
    End Sub

    Private Sub Frm_Closing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        Close()
    End Sub

    Private Sub Frm_Move(ByVal sender As Object, ByVal e As System.EventArgs)
        SetPosition()
    End Sub

    Private Sub Frm_Resize(ByVal sender As Object, ByVal e As System.EventArgs)
        SetPosition()
    End Sub
End Class