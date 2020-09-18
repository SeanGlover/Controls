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
    Private Const IconPadding As Integer = 3
    Private ReadOnly Property Table As DataViewer
    Private WithEvents OK As New Button With {.Text = "OK", .Font = PreferredFont, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.ButtonYes, .BackColor = Color.GhostWhite, .ForeColor = Color.Black, .FlatStyle = FlatStyle.Popup}
    Private WithEvents YES As New Button With {.Text = "Yes", .Font = PreferredFont, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.ButtonYes, .BackColor = Color.GhostWhite, .ForeColor = Color.Black, .FlatStyle = FlatStyle.Popup}
    Private WithEvents NO As New Button With {.Text = "No", .Font = PreferredFont, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.ButtonNo, .BackColor = Color.GhostWhite, .ForeColor = Color.Black, .FlatStyle = FlatStyle.Popup}
    Private WithEvents PromptTimer As New Timer With {.Interval = 5000}
    Private ReadOnly ParentControl As Control
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
    Public Sub New(Optional parentWindow As Control = Nothing)

        ParentControl = parentWindow
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
        MaximumSize = New Size(Convert.ToInt32(0.7 * WorkingArea.Width), Convert.ToInt32(0.7 * WorkingArea.Height))
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
            _Table = New DataViewer With {
                .Font = PreferredFont,
                .Dock = DockStyle.Fill,
                .DataSource = value}
        End Set
        Get
            Return _DataSource
        End Get
    End Property
    Public Property ColorStyle As StyleOption = StyleOption.Plain
    Private ReadOnly Property AlternatingRowColor As Color
    Private ReadOnly Property BackgroundColor As Color
    Private ReadOnly Property HeaderTextColor As Color = Color.White
    Private ReadOnly Property ShadeColor As Color
    Private ReadOnly Property AccentColor As Color
    Public Property OutlineText As Boolean
    Public Property TitleMessage As String
    Public Property BodyMessage As String
    Public Property TitleBarImageLeftSide As Boolean = True
    Public Property TitleBarImage As Image = My.Resources.Info_White
    Public Property BorderColor As Color = Color.Transparent
    Public Property BorderForeColor As Color = Color.White
    Private Icon_ As Icon = Nothing
    Public Overloads Property Icon As Icon
        Get
            If Icon_ Is Nothing Then
                Select Case Type
                    Case IconOption.Critical
                        Return My.Resources.okNot

                    Case IconOption.OK
                        Return My.Resources.ok

                    Case IconOption.TimedMessage
                        Return My.Resources.Clock

                    Case IconOption.Warning
                        Return My.Resources.Warning

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
    Private AutoCloseSeconds_ As Integer = 3
    Public Property AutoCloseSeconds As Integer
        Get
            Return AutoCloseSeconds_
        End Get
        Set(value As Integer)
            AutoCloseSeconds_ = value
            PromptTimer.Interval = 1000 * value
        End Set
    End Property
    Private ReadOnly Property IconBounds As Rectangle
        Get
            Return New Rectangle(IconPadding, IconPadding, Icon.Width, Icon.Height)
        End Get
    End Property
    Private ReadOnly Property TextBounds As New Dictionary(Of Rectangle, String)
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
                Dim ImageBounds As Rectangle
                Dim TextBounds As Rectangle

                If TitleBarImageLeftSide Then
                    ImageBounds = New Rectangle(horizontalPadding, yOffset, ImageWidth, ImageHeight)
                    TextBounds = New Rectangle(ImageWidth + horizontalPadding, 0, Width - (ImageWidth + horizontalPadding), TitleBarBounds.Height)
                Else
                    TextBounds = New Rectangle(horizontalPadding, 0, Width - (ImageWidth + horizontalPadding), TitleBarBounds.Height)
                    ImageBounds = New Rectangle(TextBounds.Width + horizontalPadding, yOffset, ImageWidth, ImageHeight)
                End If
                g.DrawImage(TitleBarImage, ImageBounds)
                TextRenderer.DrawText(g, TitleMessage, PreferredFont, TextBounds, HeaderTextColor, BorderColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)
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
            For Each TextBound In TextBounds
                Dim isURL As Boolean = Regex.Match(TextBound.Value, "https{0,1}:[^ ]{1,}", RegexOptions.None).Success
                If isURL Then
                    Using urlFont As New Font(Font.FontFamily, Font.Size, FontStyle.Underline)
                        TextRenderer.DrawText(e.Graphics, TextBound.Value, urlFont, TextBound.Key.Location, Color.Blue)
                    End Using
                    AddressBounds.Add(TextBound.Key, TextBound.Value)
                Else
                    TextRenderer.DrawText(e.Graphics, TextBound.Value, Font, TextBound.Key.Location, ForeColor)
                End If
                If OutlineText Then e.Graphics.DrawRectangle(Pens.Red, TextBound.Key)
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
        Close()

    End Sub
    Protected Overrides Sub OnFontChanged(e As EventArgs)

        If Table IsNot Nothing Then Table.Font = Font
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
    Private Shadows Sub OnShown()
    End Sub
    Private Sub Message_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        _Table = Nothing
        TitleMessage = String.Empty
        BodyMessage = String.Empty
        If MainWindow IsNot Nothing Then
            Dim result As Integer = NativeMethods.SetForegroundWindow(MainWindow.Handle)
        End If

    End Sub
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
        ForeColor = TextColor
        _ShadeColor = ShadeColor
        _AccentColor = AccentColor
        _BorderColor = BorderColor

        ColorStyle = StyleOption.Custom
        Return Show(TitleMessage, BodyMessage, Type, StyleOption.Custom)

    End Function
    Public Overloads Function Show(TitleMessage As String, BodyMessage As String(), Optional Type As IconOption = IconOption.OK, Optional ColorTheme As StyleOption = StyleOption.Plain, Optional AutoCloseSeconds As Integer = 3) As DialogResult
        Return Show(TitleMessage, Join(BodyMessage, vbNewLine), Type, ColorTheme, AutoCloseSeconds)
    End Function
    Public Overloads Function Show(titleBarMessage As String, bodyMessage As String, Optional Type As IconOption = IconOption.OK, Optional ColorTheme As StyleOption = StyleOption.Plain, Optional AutoCloseSeconds As Integer = 3) As DialogResult

        ControlBox = False
        TitleMessage = titleBarMessage
        Text = TitleMessage
        PromptTimer.Interval = 1000 * AutoCloseSeconds

        bodyMessage = If(bodyMessage, String.Empty)
        Me.BodyMessage = bodyMessage
        If bodyMessage.Length = 0 Then Me.BodyMessage = "No Message"

        Me.Type = Type
        ColorStyle = ColorTheme

        Select Case ColorTheme
            Case StyleOption.BlackGold
                _AlternatingRowColor = Color.Gold
                _BackgroundColor = Color.Black
                _HeaderTextColor = Color.White
                ForeColor = Color.White
                _ShadeColor = Color.DarkKhaki
                _AccentColor = Color.DarkGoldenrod
                BorderColor = If(BorderColor = Color.Transparent, Color.Black, BorderColor)

            Case StyleOption.Blue
                _AlternatingRowColor = Color.LightSkyBlue
                _BackgroundColor = Color.CornflowerBlue
                _HeaderTextColor = Color.White
                ForeColor = Color.Black
                _ShadeColor = Color.DarkBlue
                _AccentColor = Color.DarkSlateBlue
                BorderColor = If(BorderColor = Color.Transparent, Color.RoyalBlue, BorderColor)

            Case StyleOption.Bright
                _AlternatingRowColor = Color.Gold
                _BackgroundColor = Color.HotPink
                _HeaderTextColor = Color.White
                ForeColor = Color.White
                _ShadeColor = Color.Fuchsia
                _AccentColor = Color.DarkOrchid
                BorderColor = If(BorderColor = Color.Transparent, Color.DarkMagenta, BorderColor)

            Case StyleOption.Grey
                _AlternatingRowColor = Color.DarkGray
                _BackgroundColor = Color.Gainsboro
                _HeaderTextColor = Color.White
                ForeColor = Color.Black
                _ShadeColor = Color.Silver
                _AccentColor = Color.Gray
                BorderColor = If(BorderColor = Color.Transparent, Color.Black, BorderColor)

            Case StyleOption.Earth
                _AlternatingRowColor = Color.Beige
                _BackgroundColor = Color.Green
                _HeaderTextColor = Color.White
                ForeColor = Color.White
                _ShadeColor = Color.DarkGreen
                _AccentColor = Color.DarkOliveGreen
                BorderColor = If(BorderColor = Color.Transparent, Color.Maroon, BorderColor)

            Case StyleOption.Psychedelic
                _AlternatingRowColor = Color.Lavender
                _BackgroundColor = Color.Fuchsia
                BackgroundImageLayout = ImageLayout.Stretch
                _HeaderTextColor = Color.White
                ForeColor = Color.White
                _ShadeColor = Color.Gainsboro
                _AccentColor = Color.DarkOrange
                BorderColor = If(BorderColor = Color.Transparent, Color.YellowGreen, BorderColor)

            Case StyleOption.Plain
                _AlternatingRowColor = Color.Gainsboro
                _BackgroundColor = Color.LightGray
                _HeaderTextColor = Color.White
                ForeColor = Color.Black
                _ShadeColor = Color.DarkGray
                _AccentColor = Color.Gainsboro
                BorderColor = If(BorderColor = Color.Transparent, Color.Silver, BorderColor)

            Case StyleOption.RedBrown
                _AlternatingRowColor = Color.Chocolate
                _BackgroundColor = Color.Orange
                _HeaderTextColor = Color.White
                ForeColor = Color.White
                _ShadeColor = Color.Crimson
                _AccentColor = Color.Peru
                BorderColor = If(BorderColor = Color.Transparent, Color.SaddleBrown, BorderColor)

            Case StyleOption.Custom

        End Select

        For Each InputButton As Button In {YES, NO, OK}
            InputButton.ForeColor = HeaderTextColor
            InputButton.BackColor = BorderColor
        Next

        PathColors = {Color.Gray, BackColor, ShadeColor, AccentColor}.ToList
        ResizeMe()
        If ParentControl Is Nothing Then
            ShowDialog()
        Else
            ShowDialog(ParentControl)
        End If
        Return DialogResult

    End Function
    Private Enum CharacterType
        None
        Space
        NotSpace
    End Enum
    Private Sub ResizeMe()

        Dim proposedFormSize As New Size()
        Dim proposedTextSize As New Size()
        Dim proposedGridSize As New Size()

        If BodyMessage.Any Then
#Region " #0 Get word sizes "
            Dim characterGroups As New Dictionary(Of Integer, String)
            Dim firstLetter As Char = BodyMessage.First
            Dim lastType As CharacterType = If(TrimReturn(firstLetter).Any, CharacterType.NotSpace, CharacterType.Space)
            Dim typeString As String = String.Empty
            For Each letter As Char In BodyMessage
                Dim currentType As CharacterType = If(TrimReturn(letter).Any, CharacterType.NotSpace, CharacterType.Space)
                If lastType <> currentType Then
                    characterGroups.Add(characterGroups.Count, typeString)
                    lastType = currentType
                    typeString = String.Empty
                End If
                typeString &= letter
            Next
            characterGroups.Add(characterGroups.Count, typeString)
            Dim wordSizes As New Dictionary(Of Integer, Size)
            For Each group In characterGroups
                wordSizes.Add(group.Key, MeasureText(group.Value, Font))
            Next
            Dim rowHeight As Integer = wordSizes.Values.Max(Function(w) w.Height)
            Dim largestWord As Integer = wordSizes.Values.Max(Function(w) w.Width)
#End Region
#Region " #1 Get best size for the Form "
            If Datasource Is Nothing Then
#Region " Text.Size is the main driver of the Form's Size - Get proposed text dimensions [width X height] "
                'a) Width & Height as a Rectangle based on total area...if there are long words wider than the Rectangle, it will need expanding
                Dim TextSize As Size = MeasureText(BodyMessage, Font)
                Dim TextArea As Integer = TextSize.Width * TextSize.Height
                Dim width2height_Ratio As Double = 3 'Prompt box looks good when width is 3 times the height
                'Area = Width * Height                       ex Area = 10,000 ( x * y )
                '∵ Width = 3*Height                         ex x = 3y
                'Area = Width (3*Height) * Height            ex Area = 3y * y
                'Area = 3*Height * Height                    ex y² * 3 = 10,000
                'Area = Height² * 3                          ex y² = 10,000 / 3
                'Height = √Area/3                            ex y = √3,333.33   57.73
                'Width = Height * 3                          ex x = 57.73 * 3 = 173.21   ... 57.73 * 173.21 = 10,000
                'Area of 10,000 should have a width of 173.21 and a height of 57.73
                Dim proposedTextHeight As Double = Math.Sqrt(TextArea / width2height_Ratio)
                Dim proposedTextWidth As Double = proposedTextHeight * width2height_Ratio
                proposedTextSize = New Size(Convert.ToInt32({proposedTextWidth, largestWord}.Max), Convert.ToInt32(proposedTextHeight))
#End Region
            Else
                With Table
                    .Visible = True
                    .AutoSize = True
                    With .Columns.HeaderStyle
                        .BackColor = ShadeColor
                        .ShadeColor = ShadeColor
                        .ForeColor = ForeColor
                    End With
                    With .Rows
                        With .RowStyle
                            .BackColor = Color.GhostWhite
                            .ForeColor = Color.Black
                        End With
                        With .AlternatingRowStyle
                            .BackColor = AlternatingRowColor
                            .ForeColor = Color.Black
                        End With
                    End With
                End With
#Region " Grid.Size is the main driver of the Form's Size - Get proposed text dimensions [width X height] "
                Dim idealGridWidth As Integer = Table.IdealSize.Width
                Dim proposedGridWidth As Integer = {1000, {idealGridWidth, largestWord, 200}.Max}.Min 'No smaller than 200 wide, no larger than 1000 wide
                proposedGridWidth = If(MaximumSize.IsEmpty, proposedGridWidth, {proposedGridWidth, MaximumSize.Width}.Min) 'Ensure within MaximumSize.Width, if any
                Dim gridRowsCount As Integer = {{1, Table.Rows.Count}.Max, 15}.Min
                Dim proposedGridHeight As Integer = Table.IdealSize.Height
                proposedGridHeight = If(MaximumSize.IsEmpty, proposedGridHeight, {proposedGridHeight, MaximumSize.Height}.Min) 'Ensure within MaximumSize.Height, if any
                proposedGridSize = New Size(proposedGridWidth, proposedGridHeight)
                Dim tlpTable As New TableLayoutPanel With {
                .Name = "gridContainer",
                .ColumnCount = 1,
                .RowCount = 1,
                .Margin = New Padding(0),
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                .BorderStyle = BorderStyle.None}
                With tlpTable
                    .Size = proposedGridSize
                    .ColumnStyles.Add(New ColumnStyle With {
                                      .SizeType = SizeType.Absolute,
                                      .Width = proposedGridWidth
                                      })
                    .RowStyles.Add(New RowStyle With {
                                      .SizeType = SizeType.Absolute,
                                      .Height = proposedGridHeight
                                      })
                    .Controls.Add(Table)
                End With
                TLP.SetSize(tlpTable)
                Controls.Add(tlpTable)
#End Region
            End If
#End Region
            'If no Grid, TextSize = ( Height * 3, Height ). If Grid, then Form.Width will depend on Grid.Size
            'Now check for Form Min and Max Size restrictions...
            Dim proposedClientWidth As Integer = IconBounds.Right + If(proposedGridSize.IsEmpty, proposedTextSize.Width, proposedGridSize.Width)
            proposedClientWidth = If(MaximumSize.IsEmpty, proposedClientWidth, {proposedClientWidth, MaximumSize.Width}.Min)
            proposedClientWidth = If(MinimumSize.IsEmpty, proposedClientWidth, {proposedClientWidth, MinimumSize.Width}.Max)

            'Now a Width is determined, the TextBounds collection can be filled
            'XXXXXXXXXX
            'XXXXXXXXXX  L I N E   1 ...................
            'XXXXXXXXXX_________________________________
            'XXXXXXXXXX
            'XXXXXXXXXX  L I N E   2 ...................
            'XXXXXXXXXX_________________________________
            '
            ' L I N E   3 ..............................

#Region " #3 - Get word rectangles "
            Dim lines As New Dictionary(Of Integer, List(Of String))
            Dim lineIndex As Integer = 0
            Dim leftBuffer As Integer = 6
            Dim wordBoundsLeft As Integer = IconBounds.Right + leftBuffer
            Dim pastIcon As Boolean = False
            TextBounds.Clear()
            For Each wordSize In wordSizes 'Indexed words and spaces
                Dim isReturn As Boolean = {vbNewLine, vbCrLf, vbCr}.Contains(characterGroups(wordSize.Key))
                If isReturn Or wordBoundsLeft + wordSize.Value.Width > proposedClientWidth Then
                    'Image.Width + Word.Width > Content.Width ... new line
                    pastIcon = rowHeight * lines.Count > IconBounds.Bottom
                    wordBoundsLeft = If(pastIcon, leftBuffer, IconBounds.Right + leftBuffer)
                    lineIndex += 1
                End If
                If Not lines.ContainsKey(lineIndex) Then lines.Add(lineIndex, New List(Of String))
                lines(lineIndex).Add(characterGroups(wordSize.Key))
                TextBounds.Add(New Rectangle(wordBoundsLeft, IconPadding + (lineIndex * rowHeight), wordSize.Value.Width, rowHeight), characterGroups(wordSize.Key))
                wordBoundsLeft += wordSize.Value.Width
            Next
            If pastIcon Then
                Dim relativeSizing = TextIconSizing(rowHeight)
                Dim iconAdjust As Integer = relativeSizing.Key
                Dim textAdjust As Integer = relativeSizing.Value

            Else
                Dim textHeight As Integer = lines.Count * rowHeight
                Dim extraSpace As Integer = CInt(QuotientRound(IconBounds.Height - textHeight, 2))
                If extraSpace > 0 Then
                    Dim newBounds As New Dictionary(Of Rectangle, String)
                    For Each textBound In TextBounds
                        With textBound.Key
                            newBounds.Add(New Rectangle(.Left, .Top + extraSpace, .Width, .Height), textBound.Value)
                        End With
                    Next
                    _TextBounds = newBounds
                End If
            End If
#End Region
        End If

        Width = 3 + (SideBorderWidths * 2) + {TextBounds.Max(Function(x) x.Key.Right), proposedGridSize.Width}.Max + 3 'Exterior width
        Dim clientWidth As Integer = ClientSize.Width 'Interior / available width

        Dim textBottom As Integer = IconPadding + {TextBounds.Keys.Last.Bottom, IconBounds.Bottom}.Max
        Dim tlpGrid As Control = Controls.Item("gridContainer")
        If tlpGrid Is Nothing Then
            _ButtonBarBounds = New Rectangle(0,
                                             {textBottom, MinimumSize.Height - ButtonBarHeight}.Max,
                                             clientWidth - 1, If(Type = IconOption.TimedMessage, 0,
                                             ButtonBarHeight))
        Else
            Table.Columns.DistibuteWidths()
            tlpGrid.Location = New Point(CInt((clientWidth - tlpGrid.Width) / 2), textBottom)
            _ButtonBarBounds = New Rectangle(0,
                                 {tlpGrid.Bounds.Bottom, MinimumSize.Height - ButtonBarHeight}.Max,
                                 clientWidth - 1, If(Type = IconOption.TimedMessage, 0,
                                 ButtonBarHeight))
        End If

        Height = TitleBarHeight + ButtonBarBounds.Bottom + BottomBorderHeight 'Exterior height

#Region " BUTTON PLACEMENT / VISIBILITY "
        For Each Button In {YES, NO, OK}
            Button.Visible = False
            Button.Top = ButtonBarBounds.Top + Convert.ToInt32((ButtonBarBounds.Height - Button.Height) / 2)
        Next
        Select Case Type
            Case IconOption.Critical, IconOption.OK, IconOption.Warning
                OK.Visible = True
                OK.Left = Convert.ToInt32((clientWidth - OK.Width) / 2)

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
    Private Function TextIconSizing(textHeight As Integer) As KeyValuePair(Of Integer, Integer)

        '/// Key = Icon.Height delta, Value=Text.Height delta ... both must grow only as shrinking the Icon or Text height not ideal
        '/// 4 possible outcomes: a) Neither change, b) Text grows, c) Icon grows or d) both grow
        Dim iconHeight As Integer = IconBounds.Height

        Dim qr = QuotientRemainder(iconHeight, textHeight) 'renders ==> (#Rows of text, #Pixels total remaining)
        Dim rows As Byte = CByte(qr.Key)
        Dim pixels As Byte = CByte(qr.Value)
        '(48, 17)=(2, 14) meaning 2 rows with 14 pixels to split between the 2 rows ( 7 each - too high ) ... additional row is just past the Icon bottom
        '(48, 23)=(2, 2)  meaning 2 rows with 2 pixels to split between the 2 rows ( 1 each - OK ) ... text line height is just short of the icon bottom
        '/// ∴ Low remainder = grow Text while high remainder = grow Icon

        Dim textPixelsGrow = QuotientRemainder(pixels, rows) 'Determines how to distribute pixels...(#Pixels, #Rows) ==> (14 pixels, 2 rows) ==> ( 7 pixels, 0 remainder)

        If textPixelsGrow.Key <= 4 Then
            'OK to use a hard value of 4 since padding 2 above text and 2 below text is ok, more than that is noticeable
            Return New KeyValuePair(Of Integer, Integer)(Convert.ToInt32(textPixelsGrow.Value), Convert.ToInt32(textPixelsGrow.Key))

        Else
            Dim iconDelta As Integer = textHeight - pixels 'If textHeight=17 and pixels=14 then only 3 change
            'Try evenly splitting pixels among the Icon and Rows
            Dim pixelSplit = QuotientRemainder(iconDelta, rows + 1) '...say 4 delta amoung 2 rows and Icon
            Dim textGrowMax As Long = {pixelSplit.Key, 4}.Max
            Dim iconGrowValue As Long = textGrowMax - pixelSplit.Key
            Return New KeyValuePair(Of Integer, Integer)(Convert.ToInt32(iconGrowValue), Convert.ToInt32(textGrowMax))

        End If

    End Function
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
    Protected Overrides Sub OnShown(ByVal e As EventArgs)
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
Public Class InvisibleForm
    Inherits Form
    Private Const WM_NCACTIVATE As Integer = &H86
    Private Const WM_NCPAINT As Integer = &H85

    Public Property BorderColor As Color = Color.Black
    Public Property TitleImage As Image
    Public Property ImageAlign As HorizontalAlignment = HorizontalAlignment.Left
    Private ReadOnly Property TitleFont As New Font("Segoe UI", 9)
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

    Public Sub New()

        ControlBox = False
        BackColor = Color.Lime
        TransparencyKey = Color.Lime
        FormBorderStyle = FormBorderStyle.None
        BackgroundImageLayout = ImageLayout.Center
        ShowInTaskbar = False
        TitleImage = My.Resources.Plus

    End Sub

    Protected Overrides Sub OnTextChanged(e As EventArgs)

        If Text Is Nothing Then
            FormBorderStyle = FormBorderStyle.None
        Else
            FormBorderStyle = FormBorderStyle.FixedSingle
        End If

    End Sub
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            'cp.ExStyle = cp.ExStyle Or 33554432
            cp.ClassStyle = cp.ClassStyle Or &H200
            Return cp
        End Get
    End Property

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
            If DrawForm Or Not DrawForm Then
                Using BC As New SolidBrush(Color.WhiteSmoke)
                    g.FillRectangle(BC, New RectangleF(0, 0, Width, Height))
                End Using
                Using BC As New SolidBrush(BorderColor)
                    g.FillRectangle(BC, TitleBarBounds)
                End Using
            End If
            If If(Text, String.Empty).Any Then
                Dim horizontalPadding As Integer = 2
                Dim MaxImageHeight As Integer = TitleBarBounds.Height - 4
                Dim ImageHeight As Integer = If(TitleImage.Height > MaxImageHeight, MaxImageHeight, TitleImage.Height)
                Dim ImageWidth As Integer = ImageHeight     'Default SQUARE

                If TitleImage.Width = TitleImage.Height Then
                    'Square, so no fancy calcs

                Else
                    Dim TextSize As Size = MeasureText(Text, TitleFont)
                    Dim MaxImageWidth As Integer = TitleBarBounds.Width - (horizontalPadding + TextSize.Width + horizontalPadding)
                    ImageWidth = If(TitleImage.Width > MaxImageWidth, MaxImageWidth, TitleImage.Width)

                End If

                Dim yOffset As Integer = Convert.ToInt32((TitleBarBounds.Height - ImageHeight) / 2)
                Dim ImageBounds As Rectangle
                Dim TextBounds As Rectangle

                If ImageAlign = HorizontalAlignment.Left Then
                    ImageBounds = New Rectangle(horizontalPadding, yOffset, ImageWidth, ImageHeight)
                    TextBounds = New Rectangle(ImageWidth + horizontalPadding, 0, Width - (ImageWidth + horizontalPadding), TitleBarBounds.Height)
                Else
                    TextBounds = New Rectangle(horizontalPadding, 0, Width - (ImageWidth + horizontalPadding), TitleBarBounds.Height)
                    ImageBounds = New Rectangle(TextBounds.Width + horizontalPadding, yOffset, ImageWidth, ImageHeight)
                End If
                g.DrawImage(TitleImage, ImageBounds)
                TextRenderer.DrawText(g, Text, TitleFont, TextBounds, Color.White, BorderColor, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)
            End If

        End Using
        Dim Result = NativeMethods.ReleaseDC(Handle, hdc)

    End Sub
#End Region
End Class

Public NotInheritable Class WaitTimer
    Inherits Control
    Private ReadOnly BaseForm As Form
    Private ReadOnly BaseControl As Control
    Private ReadOnly HideLocation As New Point(-500, -500)
    Private WithEvents TickTimer As New Timer With {.Interval = 100}
    Private ReadOnly TickForm As New InvisibleForm With {
        .Size = New Size(150, 150),
        .Location = HideLocation
    }
    Private ReadOnly Property DelegateImage As Image
    Public Enum ImageType
        Spin
        Circle
    End Enum

    Public Sub New(baseControl As Control, baseForm As Form)

        If baseControl IsNot Nothing Then
            TickColor = Color.Red
            Me.BaseControl = baseControl
            Me.BaseForm = baseForm
            If baseForm IsNot Nothing Then TickForm.Show(baseForm)
        End If

    End Sub

    Private Picture_ As ImageType = ImageType.Circle
    Public Property Picture As ImageType
        Get
            Return Picture_
        End Get
        Set(value As ImageType)
            If value <> Picture_ Then
                TickTimer.Interval = If(value = ImageType.Circle, 100, If(value = ImageType.Spin, 250, 300))
                Picture_ = value
                If value = ImageType.Circle Then SetSafeControlPropertyValue(TickForm, "BackgroundImage", DrawProgress(0, TickColor))
                If value = ImageType.Spin Then SetSafeControlPropertyValue(TickForm, "BackgroundImage", My.Resources.Spin1)
            End If
        End Set
    End Property
    Public Property TickColor As Color
    Public Property Offset As New Point(0, 0)
    Private FormText_ As String
    Public Property FormText As String
        Get
            Return FormText_
        End Get
        Set(value As String)
            If FormText_ <> value Then
                FormText_ = value
                SetSafeControlPropertyValue(TickForm, "Text", value)
            End If
        End Set
    End Property
    Public Property Limit As Integer = 0
    Public Property RunningIcon As Boolean
    Private TitleImage_ As Image
    Public Property TitleImage As Image
        Get
            Return TitleImage_
        End Get
        Set(value As Image)
            If Not SameImage(value, TitleImage_) Then
                TitleImage_ = value
                TickForm.TitleImage = value
                SetSafeControlPropertyValue(TickForm, "TitleImage", value)
            End If
        End Set
    End Property
    Private TimerTicks As Integer
    Public Property TickValue As Integer
        Get
            Return TimerTicks
        End Get
        Set(value As Integer)
            TickTimer.Stop()
            TimerTicks = value
            Dim centerLocation As Point = CenterItem(TickForm.Size)
            centerLocation.Offset(Offset)
            SetSafeControlPropertyValue(TickForm, "Location", centerLocation)
            SetSafeControlPropertyValue(TickForm, "Visible", True)
            _DelegateImage = DrawProgress(TimerTicks, TickColor)
            SetSafeControlPropertyValue(TickForm, "BackgroundImage", DelegateImage)
        End Set
    End Property
    Public Sub StartTicking(Optional tickColor As Color = Nothing)

        If Not tickColor.IsEmpty Then Me.TickColor = tickColor
        TimerTicks = 0
        If Limit = 0 Then TickTimer.Start()
        Dim centerLocation As Point = CenterItem(TickForm.Size)
        centerLocation.Offset(Offset)
        SetSafeControlPropertyValue(TickForm, "Location", centerLocation)
        SetSafeControlPropertyValue(TickForm, "Visible", True)
        _DelegateImage = DrawProgress(TimerTicks, tickColor)
        SetSafeControlPropertyValue(TickForm, "BackgroundImage", DelegateImage)

    End Sub
    Private Sub TickTimer_Tick(sender As Object, e As EventArgs) Handles TickTimer.Tick

        If Picture = ImageType.Circle Then
            TickForm.BackgroundImage = DrawProgress(TimerTicks, TickColor)
        Else
            Dim mod8 As Integer = TimerTicks Mod 8
            If mod8 = 0 Then TickForm.BackgroundImage = My.Resources.Spin1
            If mod8 = 1 Then TickForm.BackgroundImage = My.Resources.Spin2
            If mod8 = 2 Then TickForm.BackgroundImage = My.Resources.Spin3
            If mod8 = 3 Then TickForm.BackgroundImage = My.Resources.Spin4
            If mod8 = 4 Then TickForm.BackgroundImage = My.Resources.Spin5
            If mod8 = 5 Then TickForm.BackgroundImage = My.Resources.Spin6
            If mod8 = 6 Then TickForm.BackgroundImage = My.Resources.Spin7
            If mod8 = 7 Then TickForm.BackgroundImage = My.Resources.Spin8
        End If
        If RunningIcon And BaseForm IsNot Nothing Then BaseForm.Icon = RunIcon(TimerTicks)
        TimerTicks += 1

    End Sub
    Public Sub Increment()

        TimerTicks += CInt(100 / {Limit, 1}.Max)
        TimerTicks = {100, TimerTicks}.Min
        If TimerTicks = Limit Then TimerTicks = 0
        '1 / 50 ... 2
        Dim centerLocation As Point = CenterItem(TickForm.Size)
        centerLocation.Offset(Offset)
        SetSafeControlPropertyValue(TickForm, "Location", centerLocation)
        SetSafeControlPropertyValue(TickForm, "Visible", True)
        _DelegateImage = DrawProgress(TimerTicks, TickColor)
        SetSafeControlPropertyValue(TickForm, "BackgroundImage", DelegateImage)

    End Sub
    Public Sub StopTicking()

        TimerTicks = 0
        TickTimer.Stop()
        SetSafeControlPropertyValue(TickForm, "Location", HideLocation)
        SetSafeControlPropertyValue(TickForm, "BackgroundImage", Nothing)

    End Sub
End Class

Public Enum BarZone
    None
    Image
    Text
    Minimize
    Maximize
    Close
End Enum
Public Class BarEventArgs
    Inherits EventArgs
    Public ReadOnly Property ClickedZone As BarZone
    Public Sub New(zoneClicked As BarZone)
        ClickedZone = zoneClicked
    End Sub
End Class
Public Class TopBar
    Inherits Control
    Private ReadOnly GlossyDictionary As Dictionary(Of Theme, Image) = GlossyImages
    Public Property BarStyle As Theme = Theme.Black
    Public Property MouseStyle As Theme = Theme.Yellow
    Private TextAlignment_ As HorizontalAlignment = HorizontalAlignment.Left
    Public Property TextAlignment As HorizontalAlignment
        Get
            Return TextAlignment_
        End Get
        Set(value As HorizontalAlignment)
            If value <> TextAlignment_ Then
                TextAlignment_ = value
                Invalidate()
            End If
        End Set
    End Property
    Public Property Image As Image
    Private ReadOnly Property BoundsClose As Rectangle
        Get
            Return New Rectangle(Width - Height, 0, Height, Height)
        End Get
    End Property
    Private ReadOnly Property BoundsMaximize As Rectangle
        Get
            Return New Rectangle(BoundsClose.Left - Height, 0, Height, Height)
        End Get
    End Property
    Private ReadOnly Property BoundsMinimize As Rectangle
        Get
            Return New Rectangle(BoundsMaximize.Left - Height, 0, Height, Height)
        End Get
    End Property
    Private ReadOnly Property BoundsImage As Rectangle
        Get
            Return New Rectangle(0, 0, If(Image Is Nothing, 0, Height), Height)
        End Get
    End Property
    Private ReadOnly Property BoundsText As Rectangle
        Get
            Return New Rectangle(BoundsImage.Right, 0, BoundsMinimize.Left - BoundsImage.Right, Height)
        End Get
    End Property
    Private ReadOnly Property MouseZone As BarZone
    Private ReadOnly Property MousePoint As Point
    Private ReadOnly Property InBounds As Boolean
    Private ReadOnly Property IsDragging As Boolean

    Public Event ZoneClicked(sender As Object, e As BarEventArgs)
    Public Event BarMoved(sender As Object, e As BarEventArgs)
    Public Event BarReleased(sender As Object, e As BarEventArgs)

    Public Sub New()

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)
        BackColor = SystemColors.Window
        Dock = DockStyle.Fill

    End Sub
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If e IsNot Nothing Then
            If BarStyle = Theme.None Then
                Using backBrush As New SolidBrush(BackColor)
                    e.Graphics.FillRectangle(backBrush, Bounds)
                End Using
            Else
                e.Graphics.DrawImage(GlossyDictionary(BarStyle), Bounds)
            End If

            If Image IsNot Nothing Then
                Dim imageBounds As New Rectangle(CInt((BoundsImage.Width - Image.Width) / 2), CInt((BoundsImage.Height - Image.Height) / 2), Image.Width, Image.Height)
                e.Graphics.DrawImage(Image, imageBounds)
            End If

            Using buttonAlignment As StringFormat = New StringFormat With {
                    .LineAlignment = StringAlignment.Center,
                    .Alignment = If(TextAlignment = HorizontalAlignment.Center, StringAlignment.Center, If(TextAlignment = HorizontalAlignment.Left, StringAlignment.Near, StringAlignment.Far))
                }
                Dim glossyFore As Color = GlossyForecolor(BarStyle)
                Using glossyBrush As New SolidBrush(glossyFore)
                    e.Graphics.DrawString(Replace(Text, "&", "&&"),
                                                                              Font,
                                                                              glossyBrush,
                                                                              BoundsText,
                                                                              buttonAlignment)
                End Using
            End Using

#Region " M I N I M I Z E "
            Dim bMin As Rectangle = BoundsMinimize
            bMin.Inflate(-3, -3)
            Using gPath As GraphicsPath = DrawRoundedRectangle(bMin, 4)
                e.Graphics.DrawPath(Pens.White, gPath)
            End Using
            bMin.Inflate(-8, -8)
            Using minPen As New Pen(Brushes.White, 3)
                e.Graphics.DrawLine(minPen, New Point(bMin.Left, bMin.Bottom), New Point(bMin.Right, bMin.Bottom))
            End Using
#End Region
#Region " M A X I M I Z E "
            Dim bMax As Rectangle = BoundsMaximize
            bMax.Inflate(-3, -3)
            Using gPath As GraphicsPath = DrawRoundedRectangle(bMax, 4)
                e.Graphics.DrawPath(Pens.White, gPath)
            End Using
            bMax.Inflate(-8, -8)
            Using maxPen As New Pen(Brushes.White, 3)
                e.Graphics.DrawRectangle(maxPen, bMax)
            End Using
#End Region
#Region " C L O S E "
            Dim closeBackcolor As Color = If(MouseZone = BarZone.Close, Color.Gainsboro, BackColor)
            Dim closeLinecolor As Color = If(MouseZone = BarZone.Close, Color.Red, If(BarStyle = Theme.None, Color.White, GlossyForecolor(If(MouseZone = BarZone.Close, MouseStyle, BarStyle))))
            Dim bc As Rectangle = BoundsClose
            bc.Inflate(-3, -3)
            Using gPath As GraphicsPath = DrawRoundedRectangle(bc, 4)
                e.Graphics.DrawPath(Pens.White, gPath)
            End Using
            bc.Inflate(-8, -8)
            Using closePen As New Pen(closeLinecolor, 3)
                e.Graphics.DrawLine(closePen, New Point(bc.Left, bc.Top), New Point(bc.Right, bc.Bottom))
                e.Graphics.DrawLine(closePen, New Point(bc.Left, bc.Bottom), New Point(bc.Right, bc.Top))
            End Using
#End Region
            If InBounds Then
                Dim highlightBounds As Rectangle = If(MouseZone = BarZone.Minimize, BoundsMinimize, If(MouseZone = BarZone.Maximize, BoundsMaximize, If(MouseZone = BarZone.Close, BoundsClose, New Rectangle())))
                Using highBrush As New SolidBrush(Color.FromArgb(64, Color.Yellow))
                    e.Graphics.FillRectangle(highBrush, highlightBounds)
                End Using
            End If
        End If

    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)

        _InBounds = True
        Invalidate()
        MyBase.OnMouseEnter(e)

    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)

        _InBounds = False
        _MouseZone = BarZone.None
        Invalidate()
        MyBase.OnMouseLeave(e)

    End Sub
    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            _IsDragging = False
            If MousePoint <> e.Location Then
                Dim lastZone As BarZone = MouseZone
                If BoundsImage.Contains(e.Location) Then
                    _MouseZone = BarZone.Image

                ElseIf BoundsText.Contains(e.Location) Then
                    _MouseZone = BarZone.Text
                    If e.Button = MouseButtons.Left Then
                        _IsDragging = True
                        RaiseEvent BarMoved(Me, New BarEventArgs(MouseZone))
                    End If

                ElseIf BoundsMinimize.Contains(e.Location) Then
                    _MouseZone = BarZone.Minimize

                ElseIf BoundsMaximize.Contains(e.Location) Then
                    _MouseZone = BarZone.Maximize

                ElseIf BoundsClose.Contains(e.Location) Then
                    _MouseZone = BarZone.Close

                End If
                Cursor = If(IsDragging, Cursors.NoMove2D, Cursors.Default)
                _MousePoint = e.Location
                If lastZone <> MouseZone Then Invalidate()
            End If
        End If
        MyBase.OnMouseMove(e)

    End Sub
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)

        If MouseZone <> BarZone.None Then RaiseEvent ZoneClicked(Me, New BarEventArgs(MouseZone))
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)

        If IsDragging Then RaiseEvent BarReleased(Me, New BarEventArgs(MouseZone))
        _IsDragging = False
        Cursor = Cursors.Default
        MyBase.OnMouseUp(e)

    End Sub
    Protected Overrides Sub OnTextChanged(ByVal e As EventArgs)
        Invalidate()
        MyBase.OnTextChanged(e)
    End Sub
End Class