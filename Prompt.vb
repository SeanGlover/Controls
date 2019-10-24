Option Strict On
Option Explicit On
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms
Public Class Prompt
    Inherits Form
    Private Const WM_NCACTIVATE As Integer = &H86
    Private Const WM_NCPAINT As Integer = &H85
    Private Const ButtonBarHeight As Integer = 36
    Private Const HeaderHeight As Integer = 26
    Private Const RowHeight As Integer = 21
    Private ReadOnly Segoe As New Font("Segoe UI", 9)
    Private WithEvents Table As New DataGridView With {.Font = Segoe, .Visible = True, .Size = New Size(600, 400), .AllowUserToAddRows = False, .RowHeadersVisible = False}
    Private WithEvents OK As New Button With {.Text = "OK", .Font = Segoe, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.OK}
    Private WithEvents YES As New Button With {.Text = "Yes", .Font = Segoe, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.OK}
    Private WithEvents NO As New Button With {.Text = "No", .Font = Segoe, .Margin = New Padding(0), .Size = New Size(100, ButtonBarHeight - 6), .ImageAlign = ContentAlignment.MiddleLeft, .Image = My.Resources.Hand.ToBitmap}
    Private WithEvents PromptTimer As New Timer With {.Interval = 5000}
    Private WorkingSpace As Rectangle = Screen.PrimaryScreen.Bounds
    Private ReadOnly TextBounds As New Dictionary(Of Rectangle, String)
    Public Enum IconOption

        Critical
        OK
        TimedMessage
        Warning
        YesNo

    End Enum
    Public Enum StyleOption

        UseNew
        BlueTones
        Bright
        Dark
        Earth
        Psychedelic

    End Enum
    Public Sub New(Optional Style As StyleOption = StyleOption.BlueTones)

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
        Font = Segoe
        TopMost = True
        Me.Style = Style
        MinimizeBox = False
        MaximizeBox = False
        MinimumSize = New Size(256, 128)
        MaximumSize = New Size(Convert.ToInt32(0.7 * WorkingSpace.Width), Convert.ToInt32(0.7 * WorkingSpace.Height))
        Controls.Add(Table)
        Controls.AddRange({OK, YES, NO})

    End Sub
#Region " PROPERTIES "
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
    Public Property Style As StyleOption
    Public Property TitleMessage As String
    Public Property TitleBarImageLeft As Boolean = True
    Public Property BodyMessage As String
    Public Property BorderColor As Color = Color.Black
    Public Property BorderForeColor As Color = Color.White
    Private PathColors As New List(Of Color)({Color.Chocolate, Color.SaddleBrown, Color.Peru})
    Public Property Image As Image = SystemIcons.Information.ToBitmap
    Public Overloads Property Icon As Icon = SystemIcons.Question
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
    Private ReadOnly Property IconBounds As New Rectangle(10, 10, Icon.Width, Icon.Height)
    Private ReadOnly Property GridBounds As Rectangle
        Get
            Dim GridTop As Integer = {If(TextBounds.Keys.Any, TextBounds.Keys.Last.Bottom, 10), IconBounds.Bottom + 16}.Max
            If IsNothing(Table.DataSource) Then
                Table.Visible = False
                Table.Size = New Size(1, 1)
                Return New Rectangle(0, GridTop, 0, 0)
            Else
                Dim GridColumns As New List(Of DataGridViewColumn)(From C In Table.Columns Select DirectCast(C, DataGridViewColumn))
                Table.Width = GridColumns.Sum(Function(x) x.Width + 1)
                Table.Visible = True
                Table.Height = {3 + HeaderHeight + (RowHeight * {1, Table.Rows.Count}.Max), 360}.Min
                Return New Rectangle(0, 4 + GridTop, Table.Width, Table.Height)
            End If
        End Get
    End Property
    Private Property ButtonBarBounds As Rectangle
#End Region

#Region " PAINT "
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

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
                Using BC As New SolidBrush(Color.Black)
                    g.FillRectangle(BC, New Rectangle(0, 0, Width, Height))
                End Using
                Using bm As New Bitmap(Width, Height)
                    DrawToBitmap(bm, New Rectangle(0, 0, Width, Height))
                    g.DrawImage(bm, 0, 0, Width, Height)
                End Using
            End If
            Using BC As New SolidBrush(Color.FromArgb(32, 32, 32))
                g.FillRectangle(BC, New Rectangle(0, 0, Width, Height))
            End Using

            Dim BarIcon As Image = My.Resources.Info_White

            Dim ImageWidth As Integer = Convert.ToInt32(BarIcon.Width)
            Dim ImageHeight As Integer = Convert.ToInt32(BarIcon.Height)

            Dim Padding As Integer = Convert.ToInt32((TitleBarBounds.Height - ImageHeight) / 2)

            Dim ImageBounds As Rectangle = Nothing
            Dim TextBounds As Rectangle = Nothing

            If TitleBarImageLeft Then
                ImageBounds = New Rectangle(Padding, Padding, ImageWidth, ImageHeight)
                TextBounds = New Rectangle(ImageWidth + Padding, 0, Width - (ImageWidth + Padding), TitleBarBounds.Height)
            Else
                TextBounds = New Rectangle(Padding, 0, Width - (ImageWidth + Padding), TitleBarBounds.Height)
                ImageBounds = New Rectangle(TextBounds.Width + Padding, Padding, ImageWidth, ImageHeight)
            End If
            g.DrawImage(BarIcon, ImageBounds)
            TextRenderer.DrawText(g, TitleMessage, Segoe, TextBounds, Color.White, Color.Black, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)

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

            Dim Testing As Boolean = False

            REM /// DRAW ICON IN THE UPPER LEFT CORNER
            e.Graphics.DrawIcon(Icon, IconBounds)

            REM /// DRAW TEXT IN EACH RECTANGLE
            For Each TextBound In TextBounds.Keys
                If Testing Then
                    Dim TextRectangle As Rectangle = New Rectangle(TextBound.Left, TextBound.Top, TextBound.Width - 1, TextBound.Height)
                    Using Pen As New Pen(Brushes.White, 1)
                        Pen.DashStyle = DashStyle.DashDot
                        e.Graphics.DrawRectangle(Pen, TextRectangle)
                    End Using
                End If
                TextRenderer.DrawText(e.Graphics, TextBounds(TextBound), Segoe, TextBound, ForeColor, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
            Next
            If Testing Then
                Using Pen As New Pen(Brushes.White, 1)
                    Pen.DashStyle = DashStyle.DashDot
                    e.Graphics.DrawRectangle(Pen, GridBounds)
                End Using
                Using Pen As New Pen(Brushes.White, 1)
                    Pen.DashStyle = DashStyle.DashDot
                    e.Graphics.DrawRectangle(Pen, ButtonBarBounds)
                End Using
            End If

            If Type = IconOption.TimedMessage Then
            Else
                e.Graphics.FillRectangle(Brushes.GhostWhite, ButtonBarBounds)
                ControlPaint.DrawBorder3D(e.Graphics, ButtonBarBounds)
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
        Size = MinimumSize
        TitleMessage = String.Empty
        BodyMessage = String.Empty
        Size = New Size(200, 200)

    End Sub
#End Region
    Public Overloads Function Show(TitleMessage As String, BodyMessage As List(Of String), Optional Type As IconOption = IconOption.OK, Optional Style As StyleOption = StyleOption.UseNew, Optional WaitTime As Integer = 3) As DialogResult

        If BodyMessage Is Nothing Then
            Return Show(TitleMessage, String.Empty, Type, Style, WaitTime)
        Else
            Return Show(TitleMessage, Join(BodyMessage.ToArray, vbNewLine), Type, Style, WaitTime)
        End If

    End Function
    Public Overloads Function Show(TitleMessage As String, BodyMessage As String(), Optional Type As IconOption = IconOption.OK, Optional Style As StyleOption = StyleOption.UseNew, Optional WaitTime As Integer = 3) As DialogResult
        Return Show(TitleMessage, Join(BodyMessage, vbNewLine), Type, Style, WaitTime)
    End Function
    Public Overloads Function Show(TitleMessage As String, BodyMessage As String, Optional Type As IconOption = IconOption.OK, Optional Style As StyleOption = StyleOption.UseNew, Optional WaitTime As Integer = 3) As DialogResult

        ControlBox = False
        Table.ColumnHeadersHeight = HeaderHeight
        Table.RowTemplate.Height = RowHeight
        Me.TitleMessage = TitleMessage
        Text = TitleMessage

        BodyMessage = If(BodyMessage, String.Empty)
        Me.BodyMessage = Regex.Replace(Regex.Replace(BodyMessage, "[\n\r]{1,}", " "), "[\s]{2,}", " ")
        If BodyMessage.Length = 0 Then Me.BodyMessage = "No Message"

        Select Case Type
            Case IconOption.Critical
                Icon = SystemIcons.Error

            Case IconOption.OK
                Icon = My.Resources.Check

            Case IconOption.TimedMessage
                Icon = My.Resources.Clock
                PromptTimer.Interval = (1000 * WaitTime)

            Case IconOption.Warning
                Icon = SystemIcons.Warning

            Case IconOption.YesNo
                Icon = SystemIcons.Question

        End Select

        Me.Type = Type
        REM /// IF NEW WAS SETUP WITH STYLE OTHER THAN THE DEFAULT SHOW (BLUETONES) THEN 
        Dim SelectedStyle As StyleOption = Style
        If Style = StyleOption.UseNew Then SelectedStyle = Me.Style
        Dim AlternatingRowColor As Color, BackColor As Color, ForeColor As Color, ShadeColor As Color, AccentColor As Color
        Select Case SelectedStyle
            Case StyleOption.BlueTones
                AlternatingRowColor = Color.LightSkyBlue
                BackColor = Color.CornflowerBlue
                ForeColor = Color.White
                ShadeColor = Color.DarkBlue
                AccentColor = Color.DarkSlateBlue

            Case StyleOption.Bright
                AlternatingRowColor = Color.Gold
                BackColor = Color.HotPink
                ForeColor = Color.White
                ShadeColor = Color.Fuchsia
                AccentColor = Color.DarkOrchid

            Case StyleOption.Dark
                AlternatingRowColor = Color.DarkGray
                BackColor = Color.Gainsboro
                ForeColor = Color.White
                ShadeColor = Color.DarkSlateGray
                AccentColor = Color.Black

            Case StyleOption.Earth
                AlternatingRowColor = Color.Beige
                BackColor = Color.Green
                ForeColor = Color.White
                ShadeColor = Color.DarkGreen
                AccentColor = Color.DarkOliveGreen

            Case StyleOption.Psychedelic
                AlternatingRowColor = Color.Lavender
                BackgroundImageLayout = ImageLayout.Stretch
                ForeColor = Color.White
                ShadeColor = Color.Gainsboro
                AccentColor = Color.DarkOrange

        End Select

        With Table
            .AlternatingRowsDefaultCellStyle = New DataGridViewCellStyle With {.BackColor = AlternatingRowColor, .ForeColor = Color.Black}
            .RowsDefaultCellStyle = New DataGridViewCellStyle With {.BackColor = Color.GhostWhite, .ForeColor = Color.Black}
            With .RowHeadersDefaultCellStyle
                .BackColor = BackColor
                .ForeColor = ForeColor
            End With
        End With
        Me.ForeColor = ForeColor
        PathColors = {Color.Gray, BackColor, ShadeColor, AccentColor}.ToList
        For Each Button In (From B In Controls Where Not IsNothing(TryCast(B, Button)) Select DirectCast(B, Button))
            Button.BackColor = Color.Black
            Button.ForeColor = Color.White
        Next
        ResizeMe()
        Hide()
        ShowDialog()

        Return DialogResult

    End Function
    Public Shadows Sub Closing()
        Datasource = Nothing
    End Sub
    Private Sub ResizeMe()

        REM //////////// DRAW ICON IN UPPER LEFT CORNER OF BOX, OFFSET(10,10)
        Dim Words As String() = Split(BodyMessage, " ")
        Dim Padding As Integer = 4
        Dim LinearTextSize As Size = TextRenderer.MeasureText(BodyMessage, Segoe)
        Dim LinearHeaderTextSize As Size = TextRenderer.MeasureText(TitleMessage, Segoe)
        Dim WindowRatio As Double = (WorkingSpace.Width / WorkingSpace.Height)
        Dim RelativeHeight As Integer = Convert.ToInt32(Math.Sqrt((LinearTextSize.Width * LinearTextSize.Height) / WindowRatio))
        Dim RelativeWidth As Integer = {{Convert.ToInt32(RelativeHeight * WindowRatio), MinimumSize.Width + (2 * SideBorderWidths), (2 * SideBorderWidths) + LinearHeaderTextSize.Width, GridBounds.Width}.Max, MaximumSize.Width + (2 * SideBorderWidths)}.Min
        If Words.Count = 1 Then RelativeWidth = LinearTextSize.Width + 60
        Dim ImageTextRowCount As Integer = Convert.ToInt32(Math.Ceiling(IconBounds.Bottom / {LinearTextSize.Height, 16}.Max))

        REM ////////////
        Me.TextBounds.Clear()
        Dim TextBounds As New Dictionary(Of Rectangle, String)
        Dim LineBuilder As New StringBuilder
        Dim LineWidth As Integer = 0
        Dim ImageRow As Boolean = False
        For Each Word As String In Words
            ImageRow = (TextBounds.Count < ImageTextRowCount)
            LineWidth = (RelativeWidth - (Padding + If(ImageRow, IconBounds.Right, 0)))
            Dim NewLineWidth As Integer = TextRenderer.MeasureText(LineBuilder.ToString & " " & Word, Segoe).Width
            If NewLineWidth > LineWidth Then
                TextBounds.Add(New Rectangle(Padding + If(ImageRow, IconBounds.Right, 0), (TextBounds.Count * LinearTextSize.Height), LineWidth, LinearTextSize.Height), LineBuilder.ToString)
                LineBuilder.Clear()
            Else
            End If
            LineBuilder.Append(Word & " ")
        Next
        REM /// ADD THE LAST WRAPPED ROW- SELDOM NO LAST ROW
        If LineBuilder.Length = 0 Then
        Else
            ImageRow = (TextBounds.Count < ImageTextRowCount)
            TextBounds.Add(New Rectangle(Padding + If(ImageRow, IconBounds.Right, 0), (TextBounds.Count * LinearTextSize.Height), LineWidth, LinearTextSize.Height), LineBuilder.ToString)
        End If

        If TextBounds.Count < ImageTextRowCount Then
            For Each TextBound In TextBounds.Keys
                Dim RowHeight As Integer = Convert.ToInt32(Icon.Height / TextBounds.Count)
                Me.TextBounds.Add(New Rectangle(TextBound.Left, IconBounds.Top + (RowHeight * Me.TextBounds.Count), TextBound.Width, RowHeight), TextBounds(TextBound))
            Next
        Else
            For Each TextBound In TextBounds.Keys
                Me.TextBounds.Add(New Rectangle(TextBound.Left, TextBound.Top, TextBounds.Max(Function(x) x.Key.Right) - TextBound.Left, TextBound.Height), TextBounds(TextBound))
            Next
        End If

        Width = (SideBorderWidths * 2) + {Me.TextBounds.Max(Function(x) x.Key.Right), GridBounds.Right}.Max

        REM /// GRID WIDTH, HEIGHT, PLACEMENT
        Table.Top = GridBounds.Top
        Table.Left = Convert.ToInt32((ClientSize.Width - GridBounds.Width) / 2)
        Table.Invalidate()

        If Type = IconOption.TimedMessage Then
            ButtonBarBounds = New Rectangle(0, Padding + {TextBounds.Keys.Last.Bottom, GridBounds.Bottom}.Max, ClientSize.Width - 1, 0)
        Else
            ButtonBarBounds = New Rectangle(0, Padding + {TextBounds.Keys.Last.Bottom, GridBounds.Bottom}.Max, ClientSize.Width - 1, ButtonBarHeight)
        End If

        Height = TitleBarHeight + ButtonBarBounds.Bottom + 1 + BottomBorderHeight

        REM /// BUTTON PLACEMENT / VISIBILITY
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
        Invalidate()
        ''If TitleMessage = "Are you sure?" Then Stop
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