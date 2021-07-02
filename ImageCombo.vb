Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.ComponentModel
Public NotInheritable Class ImageComboEventArgs
    Inherits EventArgs
    Public Property ComboItem As ComboItem
    Public Sub New(TheComboItem As ComboItem)
        ComboItem = TheComboItem
    End Sub
    Public Sub New()
    End Sub
End Class
Public Enum ImageComboMode
    Combobox
    Button
    ColorPicker
    FontPicker
    RegEx
    Searchbox
    Linkbox
End Enum 'Leave Public since other controls need access
Public Enum MathSymbol
    GreaterThan
    GreaterThanEquals
    LessThan
    LessThanEquals
    Equals
    NotEquals
    Between
End Enum 'Leave Public since DatePicker also uses
Public Enum CheckStyle
    None
    Slide
    Check
End Enum
Public NotInheritable Class ImageCombo
    Inherits Control
    'Fixes:     Screen Scaling of 125, 150 distorts the CopyFromScreen in DropDown.Protected Overrides Sub OnVisibleChanged(e As EventArgs)
    Private ReadOnly ErrorTip As New ToolTip With {
        .BackColor = Color.GhostWhite,
        .ForeColor = Color.Black,
        .ShowAlways = False
    }
    Friend Toolstrip As New ToolStripDropDown With {.AutoClose = False, .AutoSize = False, .Padding = New Padding(0), .DropShadowEnabled = False, .BackColor = Color.Transparent}
    Private PaddedBounds As New Rectangle
    '[0 Image]       [1 Search]      [2 Text]     [3 Eye]       [4 Clear]       [5 DropDown]
    Private ImageBounds As New Rectangle
    Private ImageClickBounds As New Rectangle 'Full height
    Private SearchBounds As New Rectangle
    Friend TextBounds As New Rectangle
    Private TextMouseBounds As New Rectangle
    Private LinkBounds As New Rectangle
    Private ReadOnly EyeImage As Image
    Private EyeBounds As New Rectangle
    Private EyeClickBounds As New Rectangle 'Full height
    Private ReadOnly ClearTextImage As Image
    Private ClearTextBounds As New Rectangle
    Private ClearTextClickBounds As New Rectangle 'Full height
    Private ReadOnly DropImage As Image
    Private DropBounds As New Rectangle
    Private DropClickBounds As New Rectangle 'Full height
    Private CursorBounds As New Rectangle
    Private SelectionBounds As New Rectangle
    Private ProtectedBounds As New Rectangle
    Private WithEvents CursorBlinkTimer As New Timer With {.Interval = 600}
    Private CursorShouldBeVisible As Boolean = True
    Private WithEvents TextTimer As New Timer With {.Interval = 250}
    Private InBounds As Boolean
    Private TextIsVisible As Boolean = True
    Private Const Spacing As Byte = 2
    Private KeyedValue As String
    Private LastValue As String
    Private MouseLeftDown As New KeyValuePair(Of Boolean, Integer)
    Private MouseXY As Point
    Private ReadOnly GlossyDictionary As Dictionary(Of Theme, Image) = GlossyImages

    Friend Enum MouseRegion
        None
        Image
        Search
        Text
        Link
        Eye
        ClearText
        DropDown
    End Enum
    <Flags> Public Enum ValueTypes
        Any
        Integers
        Decimals
    End Enum
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
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
        BackColor = SystemColors.Window
        DropDown = New ImageComboDropDown(Me)
        Items = New ItemCollection(Me)
        EyeImage = Base64ToImage(EyeString)
        ClearTextImage = Base64ToImage(ClearTextString)
        DropImage = Base64ToImage(DropString)
        AddHandler BindingSource.DataSourceChanged, AddressOf BindingSourceChanged
        AddHandler PreviewKeyDown, AddressOf On_PreviewKeyDown

    End Sub
    Protected Overrides Sub InitLayout()

        Toolstrip.Items.Add(New ToolStripControlHost(DropDown))
        MyBase.InitLayout()

    End Sub
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Private Sub CursorBlinkTimer_Tick() Handles CursorBlinkTimer.Tick

        If If(Text, String.Empty).Any Then
            CursorShouldBeVisible = Not CursorShouldBeVisible
            Invalidate()
        End If

    End Sub
    Private Sub TextTimer_Tick() Handles TextTimer.Tick

        TextTimer.Stop()
        RaiseEvent TextPaused(Me, New ImageComboEventArgs)

    End Sub
#Region " DRAWING "

    Protected Overrides Sub OnPaint(e As PaintEventArgs)

        If e IsNot Nothing Then
            Bounds_Set()
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
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
            If Mode = ImageComboMode.Button Then
#Region " BUTTON PROPERTIES "
                e.Graphics.DrawImage(If(InBounds, GlossyDictionary(If(ButtonMouseTheme = Theme.None, Theme.Gray, ButtonMouseTheme)), GlossyDictionary(If(ButtonTheme = Theme.None, Theme.Gray, ButtonTheme))), ClientRectangle)
                e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
                Using buttonAlignment = New StringFormat With {
                    .LineAlignment = StringAlignment.Center,
                    .Alignment = If(HorizontalAlignment = HorizontalAlignment.Center, StringAlignment.Center, If(HorizontalAlignment = HorizontalAlignment.Left, StringAlignment.Near, StringAlignment.Far))
                }
                    Dim glossyFore As Color = GlossyForecolor(If(InBounds, ButtonMouseTheme, ButtonTheme))
                    Using glossyBrush As New SolidBrush(glossyFore)
                        e.Graphics.DrawString(Replace(Text, "&", "&&"),
                                                                              Font,
                                                                              glossyBrush,
                                                                              TextBounds,
                                                                              buttonAlignment)
                    End Using
                End Using
#End Region
            ElseIf Mode = ImageComboMode.Linkbox Then
                Using linkFormat As StringFormat = New StringFormat With {
                    .LineAlignment = StringAlignment.Center,
                    .Alignment = If(HorizontalAlignment = HorizontalAlignment.Center, StringAlignment.Center, If(HorizontalAlignment = HorizontalAlignment.Left, StringAlignment.Near, StringAlignment.Far))
                }
                    Using foreBrush As New SolidBrush(ForeColor)
                        e.Graphics.DrawString(Replace(Text, "&", "&&"),
                                                                                  Font,
                                                                                  foreBrush,
                                                                                  PaddedBounds,
                                                                                  linkFormat)
                        Dim range = New CharacterRange(0, Text.Length)
                        linkFormat.SetMeasurableCharacterRanges({range})
                        Dim regions = e.Graphics.MeasureCharacterRanges(Text, Font, PaddedBounds, linkFormat)
                        If regions.Any Then
                            Dim accurateBoundings As RectangleF = regions.First.GetBounds(e.Graphics)
                            With LinkBounds
                                .X = CInt(accurateBoundings.X)
                                .Y = CInt(accurateBoundings.Y)
                                .Width = CInt(accurateBoundings.Width)
                                .Height = CInt(accurateBoundings.Height)
                                If .Contains(MouseXY) Then
                                    Using linkBrush As New SolidBrush(LinkColor)
                                        Using linkPen As New Pen(linkBrush, 1)
                                            e.Graphics.DrawLine(linkPen, New Point(.Left, .Bottom), New Point(.Right, .Bottom))
                                        End Using
                                    End Using
                                End If
                            End With
                        End If
                    End Using
                End Using
            Else
#Region " REGULAR PROPERTIES "
                If Text.Any Then
                    If PasswordProtected And Not TextIsVisible Then
                        Using Brush As New HatchBrush(HatchStyle.LightUpwardDiagonal, Color.Gray, Color.WhiteSmoke)
                            e.Graphics.FillRectangle(Brush, ProtectedBounds)
                        End Using

                    Else
                        If Mouse_Region = MouseRegion.Image And Mode = ImageComboMode.ColorPicker Then
                            Dim borderColor As Color = DirectCast(Image, Bitmap).GetPixel(10, 10) ' ... Black doesn't work!
                            'Dim foreColor As Color = BackColorToForeColor(backColor)
                            Using backBrush As New SolidBrush(Color.WhiteSmoke)
                                e.Graphics.FillRectangle(backBrush, Bounds)
                            End Using
                            Dim borderRectangle As Rectangle = Bounds
                            borderRectangle.Inflate(-3, -3) : borderRectangle.Offset(-2, -2)
                            Using borderBrush As New SolidBrush(borderColor)
                                Using borderPen As New Pen(borderBrush, 3)
                                    e.Graphics.DrawRectangle(borderPen, borderRectangle)
                                End Using
                            End Using
                            TextRenderer.DrawText(e.Graphics, If(Text.Any, HintText, Text), Font, TextBounds, Color.Black, TextFormatFlags.VerticalCenter)

                        ElseIf Mouse_Region = MouseRegion.Image And Mode = ImageComboMode.FontPicker Then

                        Else
                            Dim alignFlags As TextFormatFlags = If(HorizontalAlignment = HorizontalAlignment.Center, TextFormatFlags.HorizontalCenter, If(HorizontalAlignment = HorizontalAlignment.Left, TextFormatFlags.Left, TextFormatFlags.Right))
                            Dim TextFlags As TextFormatFlags = TextFormatFlags.NoPadding Or alignFlags Or TextFormatFlags.VerticalCenter
                            If WrapText Then TextFlags = TextFlags Or TextFormatFlags.WordBreak
                            TextRenderer.DrawText(e.Graphics, Replace(Text, "&", "&&"), Font, TextBounds, ForeColor, TextFlags)
                            If Enabled Then
                                Using selectedTextBrush As New SolidBrush(Color.FromArgb(60, SelectionColor))
                                    e.Graphics.FillRectangle(selectedTextBrush, SelectionBounds)
                                End Using
                            End If
                        End If
                    End If
                Else
                    TextRenderer.DrawText(e.Graphics, HintText, Font, TextBounds, Color.DarkGray, TextFormatFlags.VerticalCenter)
                End If
                If Enabled Then
                    If Mode = ImageComboMode.Searchbox Then
                        Using searchBrush As New SolidBrush(Color.Transparent)
                            e.Graphics.FillRectangle(searchBrush, SearchBounds)
                            Using sf As New StringFormat With
                                        {
                                            .Alignment = StringAlignment.Center,
                                            .LineAlignment = StringAlignment.Center
                                        }
                                Using searchFont As New Font("Tahoma", 16)
                                    e.Graphics.DrawString(MathSymbols(SearchItem).First, searchFont, Brushes.Black, SearchBounds, sf)
                                End Using
                            End Using
                        End Using
                    End If
                    e.Graphics.DrawImage(EyeImage, EyeBounds)
                    e.Graphics.DrawImage(ClearTextImage, ClearTextBounds)
                    e.Graphics.DrawImage(DropImage, DropBounds)
                    Using mouseOverBrush As New SolidBrush(Color.FromArgb(60, SelectionColor))
                        Dim highlightBounds As Rectangle = If(Mouse_Region = MouseRegion.Image, ImageClickBounds, If(Mouse_Region = MouseRegion.Search, SearchBounds, If(Mouse_Region = MouseRegion.Eye, EyeBounds, If(Mouse_Region = MouseRegion.ClearText, ClearTextBounds, If(Mouse_Region = MouseRegion.DropDown, DropClickBounds, New Rectangle)))))
                        e.Graphics.FillRectangle(mouseOverBrush, highlightBounds)
                    End Using
                    If CursorShouldBeVisible Then
                        Using Pen As New Pen(SelectionColor)
                            e.Graphics.DrawLine(Pen, CursorBounds.X, CursorBounds.Y, CursorBounds.X, CursorBounds.Bottom)
                        End Using
                    End If
                End If
#End Region
            End If
            Dim colorBorder As Color = If(HighlightBorderOnFocus And InBounds, HighlightBorderColor, If(BorderColor = Color.Transparent, BackColor, BorderColor))
            borderBounds = New Rectangle(PaddedBounds.X, PaddedBounds.Y, PaddedBounds.Width - 1, PaddedBounds.Height - 1)
            For i = 0 To BorderWidth - 1
                Using Pen As New Pen(colorBorder, 1)
                    e.Graphics.DrawRectangle(Pen, borderBounds)
                End Using
                borderBounds.Inflate(-1, -1)
            Next
            If Image IsNot Nothing Then e.Graphics.DrawImage(Image, ImageBounds)
        End If

    End Sub
    Public ReadOnly Property IdealSize As Size
        Get
            Dim hasImage As Boolean = Image IsNot Nothing
            Dim hasText As Boolean = Text.Any
            Dim hasDrop As Boolean = Not Mode = ImageComboMode.Button And Items.Any
            Dim hasClear As Boolean = Not Mode = ImageComboMode.Button And hasText
            Dim hasEye As Boolean = PasswordProtected And hasClear
            '===========================
            Dim imageSize As Size = If(hasImage, Image.Size, New Size)
            Dim mathSize As Size = If(Mode = ImageComboMode.Searchbox, New Size(16, 16), New Size)
            Dim dropSize As Size = If(hasDrop, DropImage.Size, New Size)
            Dim clearSize As Size = If(hasClear, ClearTextImage.Size, New Size)
            Dim eyeSize As Size = If(hasEye, EyeImage.Size, New Size)
            '===========================
            Dim sizes As New List(Of Size) From {imageSize, mathSize, TextSize, dropSize, clearSize, eyeSize}
            Dim widths As New List(Of Integer)(From s In sizes Where Not s.Width = 0 Select s.Width)
            Dim heights As New List(Of Integer)(From s In sizes Where Not s.Height = 0 Select s.Height)
            '===========================
            Dim minSize As Size = If(MinimumSize.IsEmpty, New Size(60, 24), MinimumSize)
            Dim maxSize As Size = If(MaximumSize.IsEmpty, WorkingArea.Size, MaximumSize)
            Dim minmaxWidth As Integer = {{If(widths.Any, widths.Sum + Spacing * (widths.Count + 1), minSize.Width), minSize.Width}.Max, maxSize.Width}.Min
            Dim minmaxHeight As Integer = {{If(heights.Any, Spacing + heights.Max + Spacing, minSize.Height), minSize.Height}.Max, maxSize.Height}.Min
            Return New Size(minmaxWidth, minmaxHeight)
        End Get
    End Property
    Private Sub Bounds_Set()

        With Margin
            PaddedBounds = New Rectangle(.Left, .Top, Width - (.Left + .Right), Height - (.Top + .Bottom))
        End With
        Dim hasImage As Boolean = Image IsNot Nothing
        If Mode = ImageComboMode.Button Then
            With ImageBounds
                If hasImage Then
                    .X = Spacing * 2
                    .Y = {PaddedBounds.Y, CInt((PaddedBounds.Height - Image.Height) / 2)}.Max     'Might be negative if Image.Height > PaddedBounds.Height
                    .Width = Image.Width
                    .Height = {PaddedBounds.Height, Image.Height}.Min
                Else
                    .X = Spacing
                    .Y = 0
                    .Width = 0
                    .Height = PaddedBounds.Height
                End If
                ImageClickBounds.X = .X : ImageClickBounds.Y = 0 : ImageClickBounds.Width = .Width : ImageClickBounds.Height = PaddedBounds.Height
            End With
            With TextBounds
                .X = ImageBounds.Right + Spacing          'LOOKS BETTER OFFSET BY A FEW PIXELS
                .Y = 0
                .Width = PaddedBounds.Right - ({ImageBounds.Width, DropBounds.Width, ClearTextBounds.Width, EyeBounds.Width}.Sum + Spacing + Spacing + Spacing)
                .Height = PaddedBounds.Height
                TextMouseBounds.X = .X : TextMouseBounds.Y = 0 : TextMouseBounds.Width = .Width : TextMouseBounds.Height = PaddedBounds.Height
            End With
        Else
            Dim hasText As Boolean = Text.Any
            Dim hasDrop As Boolean = Not Mode = ImageComboMode.Button And Items.Any
            Dim hasClear As Boolean = Not Mode = ImageComboMode.Button And hasText
            Dim hasEye As Boolean = PasswordProtected And hasClear
            '===========================
            If AutoSize Then
                Dim bestSize As Size = IdealSize
                If bestSize <> Size Then Size = bestSize 'Never assign a Control's size unless it's <>. Setting the same size on a Control will ALWAYS fire the SizeChanged.Event
            End If
            With ImageBounds
                If hasImage Then
                    .X = PaddedBounds.X + Spacing
                    .Y = PaddedBounds.Top + CInt((PaddedBounds.Height - Image.Height) / 2)
                    .Width = Image.Width
                    .Height = {PaddedBounds.Height, Image.Height}.Min
                Else
                    .X = PaddedBounds.X
                    .Y = 0
                    .Width = 0
                    .Height = PaddedBounds.Height
                End If
                ImageClickBounds.X = .X : ImageClickBounds.Y = PaddedBounds.Top : ImageClickBounds.Width = .Width : ImageClickBounds.Height = PaddedBounds.Height
            End With
            With SearchBounds
                .X = ImageBounds.Right
                .Y = PaddedBounds.Top
                .Width = If(Mode = ImageComboMode.Searchbox, 16, 0)
                .Height = PaddedBounds.Height
            End With
            With DropBounds
                If hasDrop Then
                    'V LOOKS BETTER WHEN NOT RESIZED
                    .X = PaddedBounds.Right - (DropImage.Width + Spacing)
                    .Y = PaddedBounds.Top + CInt((PaddedBounds.Height - DropImage.Height) / 2)
                    .Width = DropImage.Width
                    .Height = {PaddedBounds.Height, DropImage.Height}.Min
                    DropClickBounds.X = .X : DropClickBounds.Y = PaddedBounds.Top : DropClickBounds.Width = .Width : DropClickBounds.Height = PaddedBounds.Height
                Else
                    .X = Width
                    .Y = 0
                    .Width = 0
                    .Height = PaddedBounds.Height
                    DropClickBounds = DropBounds
                End If
            End With
            With ClearTextBounds
                If hasClear Then
                    'X LOOKS BETTER WHEN NOT RESIZED
                    .X = PaddedBounds.Right - ({DropBounds.Width, ClearTextImage.Width}.Sum + Spacing)
                    .Y = PaddedBounds.Top + CInt((PaddedBounds.Height - ClearTextImage.Height) / 2)
                    .Width = ClearTextImage.Width
                    .Height = {PaddedBounds.Height, ClearTextImage.Height}.Min
                    ClearTextClickBounds.X = .X : ClearTextClickBounds.Y = .Y : ClearTextClickBounds.Width = .Width : ClearTextClickBounds.Height = PaddedBounds.Height
                Else
                    .X = DropBounds.Left
                    .Y = 0
                    .Width = 0
                    .Height = PaddedBounds.Height
                    ClearTextClickBounds = ClearTextBounds
                End If
            End With
            With EyeBounds
                If hasEye Then
                    .X = PaddedBounds.Right - ({DropBounds.Width, ClearTextBounds.Width, EyeImage.Width}.Sum + Spacing)
                    .Y = PaddedBounds.Top + CInt((PaddedBounds.Height - EyeImage.Height) / 2)
                    .Width = EyeImage.Width
                    .Height = {PaddedBounds.Height, EyeImage.Height}.Min
                    EyeClickBounds.X = .X : EyeClickBounds.Y = PaddedBounds.Top : EyeClickBounds.Width = .Width : EyeClickBounds.Height = PaddedBounds.Height
                Else
                    .X = ClearTextBounds.Left
                    .Y = 0
                    .Width = 0
                    .Height = PaddedBounds.Height
                    EyeClickBounds = EyeBounds
                End If
            End With
            With TextBounds
                .X = SearchBounds.Right + Spacing          'LOOKS BETTER OFFSET BY A FEW PIXELS
                .Y = PaddedBounds.Y
                .Width = PaddedBounds.Right - ({ImageBounds.Width, SearchBounds.Width, DropBounds.Width, ClearTextBounds.Width, EyeBounds.Width}.Sum + Spacing + Spacing + Spacing)
                .Height = PaddedBounds.Height
                TextMouseBounds.X = .X : TextMouseBounds.Y = PaddedBounds.Top : TextMouseBounds.Width = .Width : TextMouseBounds.Height = PaddedBounds.Height
            End With
            With CursorBounds
                .X = {Spacing, Get_LetterBoundsLeft(CursorIndex)}.Max
                .Y = PaddedBounds.Y + Spacing
                .Width = 1
                .Height = PaddedBounds.Height - Spacing * 2
            End With
            With SelectionBounds
                .X = Get_LetterBoundsLeft(SelectionStart)
                .Y = PaddedBounds.Y
                .Width = Math.Abs(Get_LetterBoundsLeft(SelectionStart) - Get_LetterBoundsLeft(SelectionEnd))
                .Height = PaddedBounds.Height
            End With
            With ProtectedBounds
                .X = TextBounds.Left
                .Y = PaddedBounds.Y
                .Width = Math.Abs(Get_LetterBoundsLeft(0) - Get_LetterBoundsLeft(LetterWidths.Keys.Last))
                .Height = PaddedBounds.Height
            End With
        End If

    End Sub
    Private Function Get_LetterBoundsLeft(indexLetter As Integer) As Integer

        '// Example    A B C D E
        '// LetterWidths dictionary is an incremental left value based on the letter index. Has a minimum count of 1 where [0]=TextBounds.X
        '// A, Index[0], Left = 2 (spacing)
        '// B, Index[1], Left = 2 + A.Width
        '// C, Index[2], Left = 2 + AB.Width
        '// D, Index[3], Left = 2 + ABC.Width
        '// E, Index[4], Left = 2 + ABCD.Width
        '// , Index[5], Left = 2 + ABCDE.Width
        '// *** 6 entries in a 5-letter word. [0] is left of letter[0] and [n] is right of letter[n]

        indexLetter = {0, indexLetter}.Max '// ensures Index not less than 0
        indexLetter = {1 + Get_LastLetterIndex(), indexLetter}.Min '// ensures not greater than 1+length
        Return LetterWidths(indexLetter)

    End Function
    Public ReadOnly Property LetterWidths As Dictionary(Of Integer, Integer)
        Get
            '// Dictionary(Of Integer, KeyValuePair(Of String, Integer))
            '// Letter Index, {Start, Width}
            Dim widths As New Dictionary(Of Integer, Integer) From
            {
                {0, TextBounds.X}
            }
            If If(Text, String.Empty).Any Then
                '// Example    A B C D E
                '// LetterWidths dictionary is an incremental left value based on the letter index. Has a minimum count of 1 where [0]=TextBounds.X
                '// A, Index[0], Left = 2 (spacing)
                '// B, Index[1], Left = 2 + A.Width
                '// C, Index[2], Left = 2 + AB.Width
                '// D, Index[3], Left = 2 + ABC.Width
                '// E, Index[4], Left = 2 + ABCD.Width
                '// , Index[5], Left = 2 + ABCDE.Width
                '// *** 6 entries in a 5-letter word. [0] is left of letter[0] and [n] is right of letter[n]

                '// letter padding can be determined by measuring the width of A * 2 vs measuring the same letter as AA
                Dim wordPadding As Integer = TextRenderer.MeasureText("A", Font).Width * 2 - TextRenderer.MeasureText("AA", Font).Width
                'Dim words As New List(Of KeyValuePair(Of String, Integer))
                'Dim letters As New List(Of KeyValuePair(Of String, Integer))
                For i As Integer = 1 To Text.Length
                    Dim growingWord As String = Text.Substring(0, i)
                    Dim wordLeft As Integer = TextBounds.Left + TextRenderer.MeasureText(growingWord, Font).Width - wordPadding
                    widths.Add(i, wordLeft)
                    'words.Add(New KeyValuePair(Of String, Integer)(growingWord, wordLeft))
                    'Dim letter As String = Text.Substring(i - 1, 1)
                    'Dim letterWidth As Integer = TextRenderer.MeasureText(letter, Font).Width - wordPadding
                    'letters.Add(New KeyValuePair(Of String, Integer)(letter, letterWidth))
                Next
            End If
            Return widths
        End Get
    End Property
    Private Function Get_LastLetterIndex() As Integer
        Return {If(Text, String.Empty).Length - 1, 0}.Max
    End Function
    Private Function Get_LetterIndex(xCoordinate As Integer) As Integer

        '// avoid IndexOutOfRangeException by bounding input value
        xCoordinate = {0, xCoordinate}.Max

        If xCoordinate > LetterWidths.Last.Value Then
            Return LetterWidths.Last.Key

        Else
            Dim lettersBefore As New List(Of KeyValuePair(Of Integer, Integer))(LetterWidths.Where(Function(lw) lw.Value <= {xCoordinate, TextBounds.X}.Max))
            Dim xBefore = If(lettersBefore.Any, lettersBefore.Last, LetterWidths.First)

            Dim lettersAfter As New List(Of KeyValuePair(Of Integer, Integer))(LetterWidths.Where(Function(lw) lw.Value >= {xCoordinate, TextBounds.X}.Max))
            Dim xAfter = If(lettersAfter.Any, lettersAfter.First, LetterWidths.Last)

            Return If(xCoordinate - xBefore.Value <= xAfter.Value - xCoordinate, xBefore.Key, xAfter.Key)
        End If

    End Function

#End Region
#Region " PROPERTIES "
    Private CursorBlinks_ As Boolean
    Public Property CursorBlinks As Boolean
        Get
            Return CursorBlinks_
        End Get
        Set(value As Boolean)
            If CursorBlinks_ <> value Then
                CursorBlinks_ = value
                If value Then
                    CursorBlinkTimer.Start()
                Else
                    CursorBlinkTimer.Stop()
                    CursorShouldBeVisible = True
                End If
            End If
        End Set
    End Property
    Friend Property Mouse_Region As MouseRegion
    Public Property CheckOnSelect As Boolean = False
    Public Property CheckboxStyle As CheckStyle = CheckStyle.Slide
    Public Property MultiSelect As Boolean
    Private ButtonTheme_ As Theme = Theme.Gray
    Public Property ButtonTheme As Theme
        Get
            Return ButtonTheme_
        End Get
        Set(value As Theme)
            If value <> ButtonTheme_ Then
                ButtonTheme_ = value
                Invalidate()
            End If
        End Set
    End Property
    Private ButtonMouseTheme_ As Theme = Theme.Yellow
    Public Property ButtonMouseTheme As Theme
        Get
            Return ButtonMouseTheme_
        End Get
        Set(value As Theme)
            If value <> ButtonMouseTheme_ Then
                ButtonMouseTheme_ = value
                Invalidate()
            End If
        End Set
    End Property
    Private AutoSize_ As Boolean = False
    Public Overrides Property AutoSize As Boolean
        Get
            Return AutoSize_
        End Get
        Set(value As Boolean)
            AutoSize_ = value
            Invalidate()
        End Set
    End Property
    Private ReadOnly Property MathSymbols As Dictionary(Of MathSymbol, String())
        Get
            If AcceptValues = ValueTypes.Any Then
                Return New Dictionary(Of MathSymbol, String()) From
        {
{MathSymbol.Equals, {"=", "="}},
{MathSymbol.NotEquals, {"≠", "<>"}}
        }
            Else
                Return New Dictionary(Of MathSymbol, String()) From
        {
            {MathSymbol.Equals, {"=", "="}},
            {MathSymbol.GreaterThanEquals, {"≥", ">="}},
            {MathSymbol.GreaterThan, {">", ">"}},
            {MathSymbol.LessThan, {"<", "<"}},
            {MathSymbol.LessThanEquals, {"≤", "<="}},
            {MathSymbol.NotEquals, {"≠", "<>"}},
            {MathSymbol.Between, {"↹", "Between"}}
        }
            End If
        End Get
    End Property
    Public Property SearchItem As MathSymbol
    Public ReadOnly Property SearchString As String
        Get
            Return MathSymbols(SearchItem).Last
        End Get
    End Property
    Public ReadOnly Property ErrorText As String
        Get
            Dim errorMessage As String = Nothing
            If If(Text, String.Empty).Any Then
                If AcceptValues = ValueTypes.Decimals Then
                    Dim amountErrors As New List(Of String)
                    For Each amount As Double In Amounts
                        If amount = Double.MaxValue Then
                            amountErrors.Add($"{amount:C2} is not recognized as a decimal")

                        ElseIf amount > MaxAcceptValue Then
                            amountErrors.Add($"{amount:C2} exceeds the maximum value of {MaxAcceptValue:C2}")

                        ElseIf amount < MinAcceptValue Then
                            amountErrors.Add($"{amount:C2} does not meet the minimum value of {MinAcceptValue:C2}")

                        Else
                        End If
                    Next
                    If amountErrors.Any Then errorMessage = String.Join(Environment.NewLine, amountErrors)

                ElseIf AcceptValues = ValueTypes.Integers Then
                    Dim amountErrors As New List(Of String)
                    For Each number As Long In Numbers
                        If number = Long.MaxValue Then
                            amountErrors.Add($"{number:N0} is not recognized as a whole number")

                        ElseIf number > MaxAcceptValue Then
                            amountErrors.Add($"{number:N0} exceeds the maximum value of {MaxAcceptValue:N}")

                        ElseIf number < MinAcceptValue Then
                            amountErrors.Add($"{number:N0} does not meet the minimum value of {MinAcceptValue:N}")

                        Else
                        End If
                    Next
                    If amountErrors.Any Then errorMessage = String.Join(Environment.NewLine, amountErrors)

                Else 'Any ... duhh
                    errorMessage = If(Text.Length < MinAcceptValue, $"{Text} does not meet the minimum length of {MinAcceptValue:N0} characters", If(Text.Length > MaxAcceptValue, $"{Text} exceeds the minimum length of {MaxAcceptValue:N0} characters", Nothing))

                End If
            End If
            Return errorMessage
        End Get
    End Property
    Public ReadOnly Property ValueError As Boolean
        Get
            Return ErrorText IsNot Nothing
        End Get
    End Property
    Public Property AcceptValues As ValueTypes
    Public Property MaxAcceptValue As Double = Double.MaxValue
    Public Property MinAcceptValue As Double = -Double.MaxValue
    Public ReadOnly Property Amounts As Double()
        Get
            Dim amountList As New List(Of Double)
            For Each amount As String In Split(Text, ";")
                Dim amountText As Double = 0
                Dim canParse As Boolean = Double.TryParse(amount, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), amountText)
                amountList.Add(If(canParse, amountText, Double.MaxValue))
            Next
            Return amountList.ToArray
        End Get
    End Property
    Public ReadOnly Property Amount As Double
        Get
            Dim amountText As Double = 0
            Dim canParse As Boolean = Double.TryParse(Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), amountText)
            Return amountText
        End Get
    End Property
    Public ReadOnly Property Numbers As Long()
        Get
            Dim longList As New List(Of Long)
            For Each amount As String In Split(Text, ";")
                Dim amountText As Long = 0
                Dim canParse As Boolean = Long.TryParse(amount, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), amountText)
                longList.Add(If(canParse, amountText, Long.MaxValue))
            Next
            Return longList.ToArray
        End Get
    End Property
    Public ReadOnly Property Number As Long
        Get
            Dim longText As Long = 0
            Dim canParse As Boolean = Long.TryParse(Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), longText)
            Return longText
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return "ImageCombo.Text=""" & Text & """"
    End Function
    Private ReadOnly BindingSource As New BindingSource
    Private _DataSource As Object
    Public Property DataSource As Object
        Get
            Return _DataSource
        End Get
        Set(value As Object)
            _DataSource = value
            BindingSource.DataSource = value
        End Set
    End Property
    Public ReadOnly Property DataType As Type
    Private Mode_ As ImageComboMode = ImageComboMode.Combobox
    Public Property Mode As ImageComboMode
        Get
            Return Mode_
        End Get
        Set(value As ImageComboMode)
            If Mode_ <> value Then
                Items.Clear()
                If value = ImageComboMode.ColorPicker Then
                    CheckboxStyle = CheckStyle.None
                    For Each colorItem In ColorImages()
                        Dim item As ComboItem = Items.Add(colorItem.Key.Name, colorItem.Value)
                        item.Color_ = colorItem.Key
                    Next

                ElseIf value = ImageComboMode.FontPicker Then
                    CheckboxStyle = CheckStyle.None
                    For Each fontItem In FontImages()
                        Dim item As ComboItem = Items.Add(fontItem.Key.Name, fontItem.Value)
                        item.Font_ = fontItem.Key
                    Next

                ElseIf value = ImageComboMode.Button Then
                    HighlightBorderOnFocus = True

                ElseIf value = ImageComboMode.Searchbox Then
                    _SearchItem = MathSymbol.Equals

                End If
                Mode_ = value
                Invalidate()
            End If
        End Set
    End Property
    Public Property IsReadOnly As Boolean
    Public ReadOnly Property DropDown As ImageComboDropDown
    Public ReadOnly Property Items As ItemCollection
    Private WithEvents DropItems_ As New ItemCollection(Me)
    Public ReadOnly Property DropItems As ItemCollection
        Get
            If Not DropItems_.Any Then DropItems_ = Items
            Return DropItems_
        End Get
    End Property
    Public Property BorderWidth As Byte = 2
    Public Property BorderColor As Color = Color.Gainsboro
    Public Property HighlightBorderColor As Color = Color.LimeGreen
    Public Property HighlightBorderOnFocus As Boolean
    Public Property SelectionColor As Color = Color.Black
    Public Property LinkColor As Color = Color.Blue
    Public Property LinkAddress As String = String.Empty
    Public Property MaxItems As Integer = 15
    Private _HorizontalAlignment As HorizontalAlignment = HorizontalAlignment.Left
    Public Property HorizontalAlignment As HorizontalAlignment
        Get
            Return _HorizontalAlignment
        End Get
        Set(value As HorizontalAlignment)
            If _HorizontalAlignment <> value Then
                _HorizontalAlignment = value
                Invalidate()
            End If
        End Set
    End Property
    Private ReadOnly Property ErrorImage As Image
        Get
            Return If(ValueError, My.Resources.lineerror, Nothing)
        End Get
    End Property
    Private ReadOnly Property ComboItemImage As Image
        Get
            Return SelectedItem?.Image
        End Get
    End Property
    Private _Image As Image
    Public Property Image As Image
        Get
            Dim overrideImage As Image = If(ErrorImage, ComboItemImage)
            Return If(overrideImage, _Image)
        End Get
        Set(value As Image)
            _Image = value
            Invalidate()
        End Set
    End Property
    Private _PasswordProtected As Boolean = False
    Public Property PasswordProtected As Boolean
        Get
            Return _PasswordProtected
        End Get
        Set(value As Boolean)
            If Not _PasswordProtected = value Then
                _PasswordProtected = value
                TextIsVisible = Not value
                Invalidate()
            End If
        End Set
    End Property
    Public Property WrapText As Boolean = False
    Public ReadOnly Property TextIndex(Optional matchText As String = Nothing) As Integer
        Get
            Dim strings As New List(Of String)(From i In Items Select i.Text)
            Return strings.IndexOf(If(matchText, Text))
        End Get
    End Property
    Public Property HintText As String
    '=======================================================
    Public ReadOnly Property SelectedItem As ComboItem
        Get
            Return If(Items.Any And SelectedIndex >= 0 And SelectedIndex < Items.Count, Items(SelectedIndex), Nothing)
        End Get
    End Property
    Private SelectedIndex_ As Integer = -1
    Public Property SelectedIndex As Integer
        Get
            Return SelectedIndex_
        End Get
        Set(value As Integer)
            If Not (value < 0 Or value >= Items.Count) Then
                OnItemSelected(Items(value), False)
                If Not MultiSelect Then
                    Dim LastSelected As New List(Of ComboItem)(From CI In Items Where CI.Selected)
                    If LastSelected.Any Then LastSelected.First._Selected = False
                End If
                SelectedIndex_ = value
                Items(value)._Selected = True
                Invalidate()
            End If
        End Set
    End Property
    Public Sub SelectAll()

        Mouse_Region = MouseRegion.Text
        CursorIndex = 0
        SelectionStart = 0
        SelectionEnd = LetterWidths.Keys.Last
        Invalidate()

    End Sub
    Private SelectionStart_ As Integer
    Public Property SelectionStart As Integer
        Get
            Return SelectionStart_
        End Get
        Set(value As Integer)
            If value <> SelectionStart_ Then
                value = {0, value}.Max
                SelectionStart_ = {value, If(Text, String.Empty).Length}.Min
                Invalidate()
            End If
        End Set
    End Property
    Private SelectionEnd_ As Integer
    Public Property SelectionEnd As Integer
        Get
            Return SelectionEnd_
        End Get
        Set(value As Integer)
            If value <> SelectionEnd_ Then
                value = {0, value}.Max
                SelectionEnd_ = {value, If(Text, String.Empty).Length}.Min
                Invalidate()
            End If
        End Set
    End Property
    Public ReadOnly Property Selection As String
        Get
            Return If(Text, String.Empty).Substring({SelectionStart, SelectionEnd}.Min, SelectionLength)
        End Get
    End Property
    Private SelectionLength_ As Integer
    Public Property SelectionLength As Integer
        Get
            SelectionLength_ = Math.Abs(SelectionEnd - SelectionStart)
            Return SelectionLength_
        End Get
        Set(value As Integer)
            value = {0, value}.Max '// ensures > 0
            SelectionLength_ = {1 + Get_LastLetterIndex() - SelectionStart, value}.Min '// ensures not greater than available length ( from start )
            SelectionEnd = SelectionStart + SelectionLength_
        End Set
    End Property
    Private CursorIndex_ As Integer
    Private Property CursorIndex As Integer
        Get
            Return CursorIndex_
        End Get
        Set(value As Integer)
            If value <> CursorIndex_ Then
                value = {0, value}.Max
                CursorIndex_ = {value, If(Text, String.Empty).Length}.Min
                Invalidate()
            End If
        End Set
    End Property

    Private SelectedColor_ As Color
    Public Property SelectedColor As Color
        Get
            SelectedColor_ = If(Mode = ImageComboMode.ColorPicker And SelectionStart >= 0, SelectedItem.Color, Color.Transparent)
            Return SelectedColor_
        End Get
        Set(value As Color)
            If value <> SelectedColor_ Then
                SelectedColor_ = value
                Dim indexItem As Integer = 0
                For Each item In Items
                    If item.Color = value Then
                        SelectedIndex_ = indexItem
                        Items(SelectedIndex)._Selected = True
                        Text = value.Name
                        Exit For
                    End If
                    indexItem += 1
                Next
                Invalidate()
            End If
        End Set
    End Property
    Private SelectedFont_ As Font
    Public Property SelectedFont As Font
        Get
            If Mode = ImageComboMode.FontPicker And SelectionStart >= 0 Then
                Return SelectedItem.Font
            Else
                Return Nothing
            End If
        End Get
        Set(value As Font)
            If value IsNot Nothing Then
                Dim indexItem As Integer = 0
                For Each item In Items
                    If item.Font.FontFamily.Name = value.FontFamily.Name Then Exit For
                    indexItem += 1
                Next
                SelectionStart = indexItem
            End If
            SelectedFont_ = value
        End Set
    End Property
    Public ReadOnly Property TextSize As Size
#End Region
#Region " EVENTS "
    Private Sub BindingSourceChanged(sender As Object, e As EventArgs)

        Items.Clear()
        If DataSource IsNot Nothing Then
            If TypeOf DataSource Is Dictionary(Of String, String) Then
                _DataType = GetType(Dictionary(Of String, String))
                Dim DictionaryStringString As Dictionary(Of String, String) = DirectCast(DataSource, Dictionary(Of String, String))
                For Each kvp In DictionaryStringString.OrderBy(Function(v) v.Value)
                    Items.Add(New ComboItem With {.Name = kvp.Value, .Value = kvp.Key})
                Next

            ElseIf TypeOf DataSource Is IEnumerable Then
                Dim Types As New List(Of List(Of Object))((From O In DirectCast(DataSource, IEnumerable).AsQueryable Where Not (IsDBNull(O) Or IsNothing(O)) Group O By Type = O.GetType Into TypeGroup = Group Select TypeGroup.ToList).ToList)
                If Types.Any Then
                    Dim Data_Type As Type = Types.First.First.GetType
                    _DataType = Data_Type
                    For Each Type In Types
                        Dim Decimals As Decimal = 0
                        If (From D In Type Where Decimal.TryParse(D.ToString, Decimals)).Count = Type.Count Then
                            Data_Type = GetType(Decimal)
                            Dim Integers As Integer = 0
                            If (From I In Type Where Integer.TryParse(I.ToString, Integers)).Count = Type.Count Then
                                Data_Type = GetType(Integer)
                            End If
                        End If
                        _DataType = Data_Type
                        Select Case Data_Type
                            Case GetType(String)
                                Dim List As New List(Of String)(From Element In Type Select CStr(Element))
                                List.Sort(Function(x, y) String.Compare(x, y, StringComparison.InvariantCulture))
                                For Each Item As String In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Date)
                                Dim List As New List(Of Date)(From Element In Type Select CDate(Element).Date)
                                List.Sort(Function(x, y) y.CompareTo(x))
                                For Each Item As Date In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Boolean)
                                Dim List As New List(Of Boolean)(From Element In Type Select CBool(Element))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Boolean In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Decimal), GetType(Double)
                                Dim List As New List(Of Decimal)(From Element In Type Select CDec(Element))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Decimal In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Long)
                                Dim List As New List(Of Long)(From Element In Type Select CLng(Element))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Long In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Integer)
                                Dim List As New List(Of Integer)(From Element In Type Select CInt(Element))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Integer In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Short)
                                Dim List As New List(Of Short)(From Element In Type Select CShort(Element))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Short In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Byte)
                                Dim List As New List(Of Byte)(From Element In Type Select CByte(Element))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Byte In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Object)

                        End Select
                    Next
                End If

            ElseIf DataSource.GetType Is GetType(DataColumn) Then
                _DataType = GetType(DataColumn)
                Dim DataColumn As DataColumn = DirectCast(DataSource, DataColumn)
                Dim List As New List(Of Object)(From C In DataColumn.Table.Rows Where Not (IsDBNull(DirectCast(C, DataRow)(DataColumn)) Or IsNothing(DirectCast(C, DataRow)(DataColumn))) Select DirectCast(C, DataRow)(DataColumn))
                DataSource = List.ToArray

            End If
        End If
        BindingContext = New BindingContext
        BindingSource.DataSource = DataSource
        Invalidate()

    End Sub
    Public Event ImageClicked(sender As Object, e As ImageComboEventArgs)
    Public Event SearchCriterionChanged(sender As Object, e As ImageComboEventArgs)
    Public Event ClearTextClicked(sender As Object, e As ImageComboEventArgs)
    Public Event ValueSubmitted(sender As Object, e As ImageComboEventArgs)
    Public Event TextPaused(sender As Object, e As ImageComboEventArgs)
    Public Event ValueChanged(sender As Object, e As ImageComboEventArgs)
    Public Event SelectionChanged(sender As Object, e As ImageComboEventArgs)
    Public Event ItemSelected(sender As Object, e As ImageComboEventArgs)
    Public Event TextPasted(sender As Object, e As EventArgs)
    Public Event TextCopied(sender As Object, e As EventArgs)
    Friend Sub OnItemSelected(ComboItem As ComboItem, DropDownVisible As Boolean)

        If ComboItem.Index <> SelectedIndex Then
            Text = ComboItem.Text
            SelectedIndex_ = ComboItem.Index
            RaiseEvent SelectionChanged(Me, New ImageComboEventArgs(ComboItem))
            RaiseEvent ValueChanged(Me, New ImageComboEventArgs(ComboItem))
        End If
        DropDown.Visible = DropDownVisible
        RaiseEvent ItemSelected(Me, New ImageComboEventArgs(ComboItem))

    End Sub
    Public Event ItemChecked(sender As Object, e As ImageComboEventArgs)
    Friend Sub OnItemChecked(ComboItem As ComboItem)
        RaiseEvent ItemChecked(Me, New ImageComboEventArgs(ComboItem))
    End Sub
    Private Sub On_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs)

        If e IsNot Nothing Then
            Select Case e.KeyCode
                Case Keys.Up, Keys.Down, Keys.Left, Keys.Right, Keys.Tab
                    e.IsInputKey = True
            End Select
        End If

    End Sub
#Region " PROTECTED OVERRIDES "
    Protected Overrides Sub OnPaddingChanged(e As EventArgs)
        Bounds_Set()
        MyBase.OnPaddingChanged(e)
    End Sub
    Protected Overrides Sub OnLeave(e As EventArgs)

        If Not LastValue = KeyedValue Then RaiseEvent ValueChanged(Me, New ImageComboEventArgs)
        LastValue = KeyedValue
        MyBase.OnLeave(e)

    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)

        InBounds = True
        Invalidate()
        MyBase.OnMouseEnter(e)

    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)

        Mouse_Region = MouseRegion.None
        InBounds = False
        Invalidate()
        MyBase.OnMouseLeave(e)

    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

        Mouse_Region = MouseRegion.Text
        TextTimer.Start()

        If e IsNot Nothing Then
            Try
                Dim S As Integer = CursorIndex
                If e.KeyCode = Keys.Left Or e.KeyCode = Keys.Right Then
#Region " MOVE LEFT Or RIGHT "
                    Dim movingLeft As Boolean = e.KeyCode = Keys.Left
                    Dim selecting As Boolean = Control.ModifierKeys = Keys.Shift
                    If selecting Then
                        If movingLeft Then
                            CursorIndex -= 1
                            SelectionStart = CursorIndex

                        Else
                            CursorIndex += 1
                            SelectionEnd = CursorIndex
                        End If
                    Else
                        If movingLeft Then
                            If SelectionStart = SelectionEnd Then
                                CursorIndex -= 1
                            Else
                                CursorIndex = SelectionStart
                            End If

                        Else
                            If SelectionStart = SelectionEnd Then
                                CursorIndex += 1
                            Else
                                CursorIndex = SelectionEnd
                            End If
                        End If
                        SelectionStart_ = CursorIndex
                        SelectionEnd = SelectionStart
                    End If
                    'If Control.ModifierKeys = Keys.Shift Then
                    '    SelectionStart += Value
                    '    CursorIndex = SelectionStart
                    'Else
                    '    CursorIndex += Value
                    'End If
#End Region
                ElseIf e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete And Not IsReadOnly Then
#Region " REMOVE BACK Or AHEAD "
                    If If(Text, String.Empty).Any Then
                        If Selection.Any Then
                            '// for either delete or backspace the selected text is removed
                            Text = Text.Remove(SelectionStart, SelectionLength)
                            SelectionEnd_ = SelectionStart
                            CursorIndex_ = SelectionStart
                        Else
                            If e.KeyCode = Keys.Back And CursorIndex >= 1 Then
                                CursorIndex -= 1
                                Text = Text.Remove(CursorIndex, 1)
                                SelectionStart_ = CursorIndex
                                SelectionEnd_ = CursorIndex
                            ElseIf e.KeyCode = Keys.Delete And CursorIndex <= Get_LastLetterIndex Then
                                Text = Text.Remove(CursorIndex, 1)
                                SelectionStart_ = CursorIndex
                                SelectionEnd_ = CursorIndex
                            End If
                        End If
                    End If
#End Region
                ElseIf e.KeyCode = Keys.A AndAlso Control.ModifierKeys = Keys.Control Then
#Region " SELECT ALL "
                    SelectAll()
#End Region
                ElseIf e.KeyCode = Keys.X AndAlso Control.ModifierKeys = Keys.Control And Not IsReadOnly Then
#Region " CUT "
                    If If(Text, String.Empty).Any Then
                        If Selection.Any Then Clipboard.SetText(Selection)
                        CursorIndex = SelectionStart
                        Text = Text.Remove(SelectionStart, Selection.Length)
                        SelectionEnd = SelectionStart
                        RaiseEvent TextCopied(Me, Nothing)
                    End If
#End Region
                ElseIf e.KeyCode = Keys.C AndAlso Control.ModifierKeys = Keys.Control Then
#Region " COPY "
                    If If(Text, String.Empty).Any And Selection.Any Then
                        Clipboard.Clear()
                        Clipboard.SetText(Selection)
                    End If
                    RaiseEvent TextCopied(Me, Nothing)
#End Region
                ElseIf e.KeyCode = Keys.V AndAlso Control.ModifierKeys = Keys.Control And Not IsReadOnly Then
#Region " PASTE "
                    Dim ClipboardText As String = String.Empty
                    Try
                        ClipboardText = Clipboard.GetText()
                    Catch ex As Runtime.InteropServices.ExternalException
                    End Try
                    If ClipboardText.Any Then
                        If Selection.Any Then
                            '// replacing text
                            Text = Text.Remove(SelectionStart, SelectionLength)
                            Text = Text.Insert(SelectionStart, ClipboardText)
                            CursorIndex = SelectionStart
                            SelectionLength = ClipboardText.Length
                        Else
                            '// inserting text
                            Text = Text.Insert(SelectionStart, ClipboardText)
                            SelectionStart += ClipboardText.Length
                            CursorIndex = SelectionStart
                            SelectionEnd = SelectionStart
                        End If
                    End If
#End Region
                ElseIf e.KeyCode = Keys.Enter Then
#Region " SUBMIT "
                    RaiseEvent ValueSubmitted(Me, New ImageComboEventArgs)
#End Region
                ElseIf e.KeyCode = Keys.Tab Then
#Region " TAB FOCUS "
                    Dim ControlCollection = (From CC In Parent.Controls Where DirectCast(CC, Control).TabStop = True And DirectCast(CC, Control).TabIndex > TabIndex)
                    If Not ControlCollection.Any Then
                        ControlCollection = (From CC In Parent.Controls Where DirectCast(CC, Control).TabStop = True)
                    End If
                    DirectCast(ControlCollection.First, Control).Focus()
#End Region
                ElseIf e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then
#Region " MOVE UP Or DOWN "
                    Dim Value As Integer = If(e.KeyCode = Keys.Up, -1, 1)
                    SelectedIndex += Value
#End Region
                ElseIf e.KeyCode = Keys.OemSemicolon And Mode = ImageComboMode.Searchbox And AcceptValues <> ValueTypes.Any Then
                    SearchItem = MathSymbol.Between
                    RaiseEvent SearchCriterionChanged(Me, New ImageComboEventArgs)
                End If
                KeyedValue = Text
                CursorShouldBeVisible = True
                CursorBlinkTimer.Stop()
                CursorBlinkTimer.Start()
                Invalidate()

            Catch ex As IndexOutOfRangeException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace)

            End Try
        End If
        MyBase.OnKeyDown(e)

    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)

        If IsReadOnly Then Exit Sub
        TextTimer.Start()

        If e IsNot Nothing Then
            Try
                If Asc(e.KeyChar) > 31 AndAlso Asc(e.KeyChar) < 127 Then
                    REM /// Printable characters
                    Dim ProposedText As String = Text
                    If Selection.Any Then ProposedText = ProposedText.Remove(SelectionStart, Selection.Length)
                    Try
                        ProposedText = ProposedText.Insert(SelectionStart, e.KeyChar)
                    Catch ex As IndexOutOfRangeException
                        Stop
                    End Try
                    CursorIndex_ = SelectionStart + 1
                    SelectionStart_ = CursorIndex
                    SelectionEnd_ = SelectionStart
                    Text = ProposedText
                    ShowMatches(Text)
                    CursorShouldBeVisible = True
                    CursorBlinkTimer.Stop()
                    CursorBlinkTimer.Start()
                    Invalidate()
                End If

            Catch ex As IndexOutOfRangeException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace)

            End Try
        End If
        MyBase.OnKeyPress(e)

    End Sub
    Friend Sub DelegateKeyPress(e As KeyPressEventArgs)
        OnKeyPress(e)
    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            MouseXY = e.Location
            Bounds_Set()
            Dim redraw As Boolean
            Dim xy As Point = e.Location
            Dim lastMouseRegion As MouseRegion = Mouse_Region
            Mouse_Region = If(ImageClickBounds.Contains(xy), MouseRegion.Image, If(SearchBounds.Contains(xy), MouseRegion.Search, If(TextMouseBounds.Contains(xy), If(LinkBounds.Contains(xy), MouseRegion.Link, MouseRegion.Text), If(EyeClickBounds.Contains(xy), MouseRegion.Eye, If(ClearTextClickBounds.Contains(xy), MouseRegion.ClearText, If(DropClickBounds.Contains(xy), MouseRegion.DropDown, MouseRegion.None))))))
            redraw = Mouse_Region <> lastMouseRegion
            If MouseLeftDown.Key And e.Button = MouseButtons.Left Then
                Dim indexMouseDownLetter = MouseLeftDown.Value
                Dim indexMouseLetter = Get_LetterIndex(e.X)
                If indexMouseDownLetter = indexMouseLetter Then
                    SelectionStart = indexMouseDownLetter
                    SelectionEnd = SelectionStart

                ElseIf indexMouseDownLetter < indexMouseLetter Then
                    '// selecting right
                    SelectionStart = indexMouseDownLetter
                    SelectionEnd = indexMouseLetter

                ElseIf indexMouseDownLetter > indexMouseLetter Then
                    '// selecting left
                    SelectionStart = indexMouseLetter
                    SelectionEnd = indexMouseDownLetter

                End If
                redraw = True
            End If
            If redraw Then Invalidate()
        End If
        MyBase.OnMouseMove(e)

    End Sub
    Private StopMe As Boolean
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

        If e IsNot Nothing Then
            If Mouse_Region = MouseRegion.Image Then
                If ValueError Then
                    Text = String.Empty
                    SelectedIndex = -1
                End If
                RaiseEvent ImageClicked(Me, New ImageComboEventArgs)

            ElseIf Mouse_Region = MouseRegion.Search Then
                '0 1 2 3
                '= > < ≠
                Dim nextIndex As Integer = MathSymbols.Keys.ToList.IndexOf(SearchItem)
                nextIndex = If(nextIndex + 1 = MathSymbols.Count, 0, nextIndex + 1)
                SearchItem = MathSymbols.Keys.ToList(nextIndex)
                RaiseEvent SearchCriterionChanged(Me, New ImageComboEventArgs)

            ElseIf Mouse_Region = MouseRegion.Text Then
                CursorShouldBeVisible = True
                CursorIndex = Get_LetterIndex(e.X)
                SelectionStart = CursorIndex
                SelectionEnd = SelectionStart
                If e.Button = MouseButtons.Left Then MouseLeftDown = New KeyValuePair(Of Boolean, Integer)(True, SelectionStart)
                StopMe = True
                CursorShouldBeVisible = True
                CursorBlinkTimer.Stop()
                CursorBlinkTimer.Start()
                If Mode = ImageComboMode.Linkbox And LinkBounds.Contains(e.Location) And LinkAddress.Any Then
                    '// open a url
                    Dim kvp = GetPreferredBrowser()
                    Process.Start(kvp.Key, LinkAddress)
                End If

            ElseIf Mouse_Region = MouseRegion.Eye Then
                TextIsVisible = Not TextIsVisible

            ElseIf Mouse_Region = MouseRegion.ClearText Then
                RaiseEvent ClearTextClicked(Me, New ImageComboEventArgs)
                If IsReadOnly Then Exit Sub
                Text = String.Empty
                SelectedIndex_ = -1
                SelectionStart_ = 0
                SelectionEnd_ = 0
                SelectionLength_ = 0

            ElseIf Mouse_Region = MouseRegion.DropDown Then
                DropItems_ = Items
                'DropDown.Visible = Not DropDown.Visible
                ShowDropDown()
                'Stop

            End If
            Invalidate()
        End If
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MouseLeftDown = Nothing
    End Sub
    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)

        If IsReadOnly Then Exit Sub
        If e IsNot Nothing And Mouse_Region = MouseRegion.Text Then SelectAll()
        MyBase.OnMouseDoubleClick(e)

    End Sub
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)

        If Items.Any Then
            'mSelectionPixelStart = TextBounds.Left
            'mSelectionPixelEnd = TextLength(Text)
            REM Always False since if Visible just became True, then DropDown should not be visible. If Visible=False, then should be False
            DropDown.Visible = False
        End If
        MyBase.OnVisibleChanged(e)

    End Sub
    Protected Overrides Sub OnParentVisibleChanged(e As EventArgs)
        DropDown.Visible = False
        MyBase.OnParentVisibleChanged(e)
    End Sub
    Protected Overrides Sub OnTextChanged(e As EventArgs)

        If Text Is Nothing Then Text = String.Empty
        _TextSize = TextRenderer.MeasureText(Text, Font) 'MeasureText(Text, Font)
        If Not Text.Any Then
            CursorIndex_ = 0
            SelectionStart_ = 0
            SelectionEnd_ = 0
            SelectionLength_ = 0
        End If
        If ValueError Then
            'ErrorTip.Show(ErrorText.ToString(InvariantCulture), Me, New Point(Width, 0))
        Else
            'ErrorTip.Hide(Me)
        End If
        Invalidate()
        MyBase.OnTextChanged(e)

    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        Bounds_Set()
        MyBase.OnSizeChanged(e)
    End Sub
    Protected Overrides Sub OnFontChanged(e As EventArgs)
        Bounds_Set()
        MyBase.OnFontChanged(e)
    End Sub
#End Region
#End Region
#Region " FUNCTIONS + METHODS "
    Private Sub DropDownItemSelected() Handles Me.ItemSelected
        SelectAll()
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub Font_Changed() Handles Me.FontChanged
        DropDown.Font = Font
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub ShowMatches(MatchText As String)

        Dim dropMatches As New List(Of ComboItem)(Items.Where(Function(ci) ci.Text.ToUpperInvariant.StartsWith(MatchText.ToUpperInvariant, StringComparison.InvariantCulture)))
        With DropDown
            .Visible = False
            If dropMatches.Any Then
                REM Show Matches
                DropItems_.Clear()
                DropItems_.AddRange(dropMatches)
                ShowDropDown()
            End If
            .VScroll.Value = 0
        End With
    End Sub
    Public Sub ShowDropDown()

        If DropItems.Any Then
            Dim Coordinates As Point
            Coordinates = PointToScreen(New Point(0, 0))
            Toolstrip.Show(Coordinates.X, If(Coordinates.Y + DropDown.Height > My.Computer.Screen.WorkingArea.Height, Coordinates.Y - DropDown.Height, Coordinates.Y + Height))
            DropDown.ResizeMe()
            DropDown.Visible = True
            'If Me.Name = "quickSearch" Then Stop
        End If

    End Sub
    Private Sub Items_Changed() Handles DropItems_.Changed
        Bounds_Set()
    End Sub
#End Region
End Class
REM ////////////////////////////////////////////////////////////////////////////////////////////////////////// DROPDOWN
Public Class ImageComboDropDown
    Inherits Control
    Public WithEvents VScroll As New VerticalScrollBar(Me) With {.LargeChange = 0}
    Private WithEvents ToolTip As New ToolTip
    Private _LastMouseOverCombo As ComboItem
    Private BMP_Shadow As Bitmap
    Private Const ShadowDepth As Integer = 12
    Private Const CheckWH As Integer = 14
    Private ReadOnly Property ComboParent As ImageCombo
    Public Sub New(Parent As ImageCombo)

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, False)
        SetStyle(ControlStyles.UserMouse, True)
        Margin = New Padding(0)
        Padding = New Padding(0)
        BackColor = Color.GhostWhite
        ForeColor = Color.DarkSlateGray
        ComboParent = Parent
        _ItemHeight = 1 + TextRenderer.MeasureText("ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ".ToString(InvariantCulture), ComboParent?.Font).Height + 1

    End Sub
    Protected Overrides Sub InitLayout()

        ResizeMe()
        MyBase.OnFontChanged(Nothing)
        MyBase.InitLayout()

    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)

        If e IsNot Nothing Then
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Dim boundsScreenShot As New Rectangle(0, 0, Width, Height)
            e.Graphics.DrawImage(BMP_Shadow, boundsScreenShot)
            boundsScreenShot.Inflate(-1, -1)
            boundsScreenShot.Offset(-2, -2)
            'e.Graphics.DrawRectangle(Pens.Red, boundsScreenShot)

            For Each ComboItem In VisibleComboItems
                With ComboItem
                    Dim boundsItem As Rectangle = .Bounds
                    boundsItem.Offset(0, -VScroll.Value)
                    Dim boundsCheck As Rectangle = .CheckBounds
                    boundsCheck.Offset(0, -VScroll.Value)
                    Dim boundsImage As Rectangle = .ImageBounds
                    boundsImage.Offset(0, -VScroll.Value)
                    Dim boundsText As Rectangle = .TextBounds
                    boundsText.Offset(0, -VScroll.Value)

                    If boundsItem.Bottom > Height - ShadowDepth Then Exit For
                    If ComboParent.CheckboxStyle = CheckStyle.Check Then
                        ControlPaint.DrawCheckBox(e.Graphics, boundsCheck, If(.Checked, ButtonState.Checked, ButtonState.Normal))

                    ElseIf ComboParent.CheckboxStyle = CheckStyle.Slide Then
                        e.Graphics.DrawImage(If(.Checked, My.Resources.slideStateOn, My.Resources.slideStateOff), boundsCheck)

                    End If
                    If Not IsNothing(.Image) Then e.Graphics.DrawImage(.Image, boundsImage)
                    If .Selected Then
                        Using Brush As New LinearGradientBrush(boundsItem, Color.FromArgb(20, SelectionColor), Color.FromArgb(60, SelectionColor), linearGradientMode:=LinearGradientMode.Vertical)
                            e.Graphics.FillRectangle(Brush, Brush.Rectangle)
                        End Using
                        Using Pen As New Pen(SelectionColor)
                            e.Graphics.DrawRectangle(Pen, boundsItem)
                        End Using
                    End If

                    TextRenderer.DrawText(e.Graphics, Replace(.Text, "&", "&&"), Font, boundsText, Color.Black, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                    If ComboItem.Separator Then e.Graphics.DrawLine(Pens.Black, New Point(0, boundsText.Bottom), New Point(boundsText.Right, boundsText.Bottom))

                    If ComboItem.Index = MouseRowIndex Then
                        Using Brush As New LinearGradientBrush(boundsItem, Color.FromArgb(20, Color.DarkSlateGray), Color.FromArgb(60, Color.DarkSlateGray), linearGradientMode:=LinearGradientMode.Vertical)
                            e.Graphics.FillRectangle(Brush, Brush.Rectangle)
                        End Using
                        Using Pen As New Pen(Color.DarkSlateGray)
                            e.Graphics.DrawRectangle(Pen, boundsItem)
                        End Using
                    End If

                End With
            Next ComboItem
            With VScroll
                If .Visible Then
                    e.Graphics.FillRectangle(Brushes.GhostWhite, .Bounds)
                    ControlPaint.DrawBorder3D(e.Graphics, .Bounds, Border3DStyle.RaisedInner)
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                    If .Lines Then
                        Using Pen As New Pen(.Color, 1)
                            For Each Page As Integer In .Pages
                                e.Graphics.DrawLine(Pen, .TrackBounds.Left, Page, .TrackBounds.Right - 2, Page)
                            Next
                        End Using
                    End If
                    Using Pen As New Pen(Brushes.Black, 1)
                        e.Graphics.DrawLine(Pen, .UpBounds.Left, .UpBounds.Bottom, .UpBounds.Right - 2, .UpBounds.Bottom)
                        e.Graphics.DrawLine(Pen, .DownBounds.Left, .DownBounds.Top, .DownBounds.Right - 2, .DownBounds.Top)
                    End Using
                    Using Brush As New SolidBrush(Color.FromArgb(.UpAlpha, .Color))
                        e.Graphics.FillRectangle(Brush, .UpBounds)
                    End Using
                    Using Brush As New SolidBrush(Color.FromArgb(.DownAlpha, .Color))
                        e.Graphics.FillRectangle(Brush, .DownBounds)
                    End Using
                    Dim ArrowWidth As Integer = 8, ArrowHeight As Integer = 4, ArrowCenter As Integer = CInt((.UpBounds.Height - ArrowHeight) / 2)
                    Dim TriangeTop As Integer = .UpBounds.Top + ArrowCenter
                    Dim TriangleLeft As Integer = .Bounds.Left + 1, TRight As Integer = TriangleLeft + ArrowWidth, TMid As Integer = TriangleLeft + CInt(ArrowWidth / 2)
                    Dim Triangle As Point() = {New Point(TMid, TriangeTop), New Point(TRight, TriangeTop + ArrowHeight), New Point(TriangleLeft, TriangeTop + ArrowHeight)}
                    Using Brush As New SolidBrush(Color.FromArgb(255, .Color))
                        e.Graphics.FillPolygon(Brush, Triangle)
                    End Using
                    Triangle = {New Point(TriangleLeft, .DownBounds.Top + ArrowCenter), New Point(TRight, .DownBounds.Top + ArrowCenter), New Point(TMid, .DownBounds.Top + ArrowCenter + ArrowHeight)}
                    Using Brush As New SolidBrush(Color.FromArgb(255, .Color))
                        e.Graphics.FillPolygon(Brush, Triangle)
                    End Using
                    Using Brush As New SolidBrush(Color.FromArgb(.Alpha, .Color))
                        e.Graphics.FillRectangle(Brush, .BarBounds)
                    End Using
                End If
            End With
        End If

    End Sub

    Private ReadOnly Property MatchedItems As List(Of ComboItem)
        Get
            Return ComboParent.DropItems
        End Get
    End Property
    Private ReadOnly Property VisibleComboItems As List(Of ComboItem)
        Get
            Return MatchedItems
        End Get
    End Property
    Private ReadOnly Property TotalHeight As Integer
        Get
            If VisibleComboItems.Any Then
                Return VisibleComboItems.Count * ItemHeight
            Else
                ComboParent.Toolstrip.Size = New Size(0, 0)
                Return 0
            End If
        End Get
    End Property
    Private _ItemHeight As Integer
    Private ReadOnly Property ItemHeight As Integer
        Get
            Return _ItemHeight
        End Get
    End Property
    Private _MouseRowIndex As Integer
    Private ReadOnly Property MouseRowIndex As Integer
        Get
            Return _MouseRowIndex
        End Get
    End Property
    Private _MouseOverCombo As ComboItem
    Private ReadOnly Property MouseOverCombo As ComboItem
        Get
            Return _MouseOverCombo
        End Get
    End Property
    Public Property ShadeColor As Color = Color.WhiteSmoke
    Public Property SelectionColor As Color = Color.Transparent
    Public Property DropShadowColor As Color = Color.Gainsboro
    Private _ForceCapture As Boolean
    Protected Property ForceCapture() As Boolean
        Get
            Return _ForceCapture
        End Get
        Set(value As Boolean)
            _ForceCapture = value
            Capture = value
        End Set
    End Property
    Protected Shadows Sub OnPreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs)

        If e IsNot Nothing Then
            Select Case e.KeyCode
                Case Keys.Up, Keys.Down
                    e.IsInputKey = True
                Case Keys.Left, Keys.Right
                    e.IsInputKey = True
            End Select
        End If

    End Sub
    Protected Overrides Sub OnMouseCaptureChanged(e As EventArgs)
        MyBase.OnMouseCaptureChanged(e)
        Capture = _ForceCapture And Visible
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

        If e IsNot Nothing Then
            Select Case e.KeyCode
                Case Keys.Up
                    If Not MouseRowIndex = 0 Then
                        If VisibleComboItems.IndexOf(MatchedItems(MouseRowIndex)) = 0 Then
                            VScroll.Value -= If(ItemHeight > VScroll.Value, VScroll.Value, ItemHeight)
                        End If
                        _MouseRowIndex -= 1
                    End If

                Case Keys.Down
                    If Not MouseRowIndex = MatchedItems.Count - 1 Then
                        If VisibleComboItems.IndexOf(ComboParent.Items(MouseRowIndex)) = ComboParent.MaxItems - 1 Then
                            VScroll.Value += ItemHeight
                        End If
                        _MouseRowIndex += 1
                    End If

                Case Keys.Return
                    ComboParent.OnItemSelected(VisibleComboItems(MouseRowIndex), False)

            End Select
            Invalidate()
        End If
        MyBase.OnKeyDown(e)

    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

        If e IsNot Nothing Then
            With ComboParent
                If VScroll.Bounds.Contains(e.Location) Then
                ElseIf Bounds.Contains(e.Location) Then
                    Dim VScrollOffset As Point = e.Location
                    VScrollOffset.Offset(0, VScroll.Value)
                    Dim Checked As New List(Of ComboItem)(From CI In VisibleComboItems Where CI.CheckBounds.Contains(VScrollOffset))
                    If Checked.Any Then
                        Checked.First.Checked = Not Checked.First.Checked
                        .OnItemChecked(Checked.First)
                        Invalidate()

                    Else
                        Dim Selected As New List(Of ComboItem)(From CI In VisibleComboItems Where CI.Bounds.Contains(VScrollOffset))
                        If Selected.Any Then
                            If Not .MultiSelect Then
                                'Dim LastSelected As New List(Of ComboItem)(From CI In Items Where CI.Selected)
                                'If LastSelected.Any Then LastSelected.First._Selected = False
                            End If
                            Selected.First._Selected = Not Selected.First.Selected
                            If .CheckOnSelect Then Selected.First.Checked = Not (Selected.First.Checked)
                            .OnItemSelected(Selected.First, Control.ModifierKeys = Keys.Shift)
                        End If
                    End If
                ElseIf .Bounds.Contains(e.Location) Then
                    .Focus()
                Else
                    Visible = False
                End If
            End With
        End If
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            If VScroll.Bounds.Contains(e.Location) Or VScroll.Scrolling Then
            ElseIf Bounds.Contains(e.Location) Then
                Dim Location As New Point(e.Location.X, (e.Location.Y + VScroll.Value))
                Dim MouseOverComboItems As New List(Of ComboItem)(From I In MatchedItems Where I.Bounds.Contains(Location) Select I)
                If MouseOverComboItems.Any Then
                    _MouseOverCombo = MouseOverComboItems.First
                    _MouseRowIndex = _MouseOverCombo.Index
                    If Not (If(MouseOverCombo.TipText, String.Empty).Length = 0 Or _LastMouseOverCombo Is _MouseOverCombo) Then
                        ToolTip.Hide(ComboParent)
                        ToolTip.Show(MouseOverCombo.TipText, ComboParent, 0, 0, 2000)
                        _LastMouseOverCombo = _MouseOverCombo
                    End If
                    ForceCapture = True
                    Invalidate()
                End If
            ElseIf ComboParent.Bounds.Contains(e.Location) Then
            End If
        End If
        MyBase.OnMouseMove(e)

    End Sub
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)

        If Visible Then
            ResizeMe()
            Top = 0
            Dim DisplayFactor = DisplayScale()
            Dim myLocation As Point = PointToScreen(New Point(0, 0))
            Dim widthScale As Integer = CInt(Width * DisplayFactor)
            Dim heightScale As Integer = CInt(Height * DisplayFactor)
            Dim bmp As New Bitmap(widthScale, heightScale)
            Try
                Using Graphics As Graphics = Graphics.FromImage(bmp)
                    Graphics.CopyFromScreen(
                            CInt(myLocation.X * DisplayFactor),
                            CInt(myLocation.Y * DisplayFactor),
                            0,
                            0,
                            bmp.Size,
                            CopyPixelOperation.SourceCopy)
                    Const shrinkFactor As Integer = -1
                    Dim rectangleShade As New Rectangle(0, 0, bmp.Width, bmp.Height)
                    For i = 0 To 23
                        Using brushShade As New SolidBrush(Color.FromArgb({16 + i * 4, 255}.Min, DropShadowColor))
                            Using pathShade As GraphicsPath = DrawRoundedRectangle(rectangleShade, 30)
                                Graphics.FillPath(brushShade, pathShade)
                            End Using
                        End Using
                        rectangleShade.Inflate(shrinkFactor, shrinkFactor)
                        rectangleShade.Offset(shrinkFactor * 2, shrinkFactor * 2)
                    Next
                End Using
                BMP_Shadow = bmp
                Dim SV As IEnumerable(Of ComboItem) = From S In MatchedItems Where S.Index = ComboParent.SelectionStart
                If SV.Any Then
                    Dim ScrollValue As Integer = MatchedItems.IndexOf(SV.First)
                    VScroll.Value = CInt(Split((ScrollValue / ComboParent.MaxItems).ToString(InvariantCulture), ".")(0)) * ComboParent.MaxItems * ItemHeight
                End If
                Invalidate()
            Catch ex As system.ComponentModel.Win32Exception
            End Try
            ForceCapture = True
        Else
            ComboParent.Toolstrip.Size = New Size(0, 0)
            ForceCapture = False
        End If
        MyBase.OnVisibleChanged(e)

    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)

        If e IsNot Nothing Then
            If MatchedItems.Any And Asc(e.KeyChar) > 31 AndAlso Asc(e.KeyChar) < 127 Then
                REM Printable characters
                ComboParent.DelegateKeyPress(e)
            End If
        End If
        MyBase.OnKeyPress(e)

    End Sub
    Protected Overrides Sub OnFontChanged(e As EventArgs)

        _ItemHeight = 1 + TextRenderer.MeasureText("ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ".ToString(InvariantCulture), ComboParent.Font).Height + 1
        ResizeMe()
        MyBase.OnFontChanged(e)

    End Sub
    Friend Sub ResizeMe()

        With ComboParent
            If MatchedItems.Any Then
                Dim indexItem As Integer = 0
                Dim widthProposed As Integer = 0
                Dim widths As New List(Of Integer)
                Dim widthCheck As Integer = If(.CheckboxStyle = CheckStyle.None, 0, If(.CheckboxStyle = CheckStyle.Slide, My.Resources.slideStateOn.Width, CheckWH))
                Dim heightCheck As Integer = If(.CheckboxStyle = CheckStyle.None, 0, If(.CheckboxStyle = CheckStyle.Slide, My.Resources.slideStateOn.Height, CheckWH))
                Dim widthScroll As Integer = If(VisibleComboItems.Count > ComboParent.MaxItems, VScroll.Bounds.Width, 0)
                'If ComboParent.Name = "quickSearch" Then Stop
                VisibleComboItems.ForEach(Sub(ci)
                                              Dim widthRow As Integer = 2 'Pad Left
                                              widthRow += If(ci.Image Is Nothing, 0, ci.Image.Width + 1) 'Item-level property
                                              widthRow += widthCheck 'ImageCombo-level property ( all items same )
                                              widthRow += TextRenderer.MeasureText(ci.Text, Font).Width 'DropDown-level property ( all items same font )
                                              widthRow += widthScroll
                                              widthRow += ShadowDepth
                                              widthRow += 5 'Pad Right
                                              If widthProposed < widthRow Then widthProposed = widthRow
                                          End Sub)
                Dim heightProposed As Integer = ItemHeight * {VisibleComboItems.Count, .MaxItems}.Min + ShadowDepth
                Size = New Size(widthProposed, heightProposed)
                VisibleComboItems.ForEach(Sub(ci)
                                              With ci
                                                  ._Index = indexItem
                                                  ._Bounds.X = 0
                                                  ._Bounds.Y = ItemHeight * .Index
                                                  ._Bounds.Width = Width - ShadowDepth
                                                  ._Bounds.Height = ItemHeight
                                                  If ComboParent.CheckboxStyle = CheckStyle.None Then
                                                      ._CheckBounds.X = 0
                                                      ._CheckBounds.Y = 0
                                                      ._CheckBounds.Width = 0
                                                      ._CheckBounds.Height = 0
                                                  Else
                                                      ._CheckBounds.X = 1
                                                      ._CheckBounds.Y = 1 + ._Bounds.Y + CInt((ItemHeight - heightCheck) / 2)
                                                      ._CheckBounds.Width = widthCheck
                                                      ._CheckBounds.Height = heightCheck
                                                  End If
                                                  If IsNothing(.Image) Then
                                                      ._ImageBounds.X = ._CheckBounds.Right
                                                      ._ImageBounds.Y = ._Bounds.Y
                                                      ._ImageBounds.Width = 0
                                                      ._ImageBounds.Height = 0
                                                  Else
                                                      Dim ImageWidth As Integer = {ItemHeight, .Image.Height}.Min
                                                      ._ImageBounds.X = ._CheckBounds.Right + 2
                                                      ._ImageBounds.Y = ._Bounds.Y + CInt((ItemHeight - ImageWidth) / 2)
                                                      ._ImageBounds.Width = ImageWidth
                                                      ._ImageBounds.Height = ImageWidth
                                                  End If
                                                  ._TextBounds.X = ._ImageBounds.Right + 2
                                                  ._TextBounds.Y = ._Bounds.Y
                                                  ._TextBounds.Width = ._Bounds.Width - ._TextBounds.X
                                                  ._TextBounds.Height = ItemHeight
                                              End With
                                              indexItem += 1
                                          End Sub)
                VScroll.Height = Height - ShadowDepth
                VScroll.Maximum = TotalHeight
                VScroll.SmallChange = ItemHeight
                VScroll.LargeChange = Height - ShadowDepth
                .Toolstrip.Size = Size
            Else
                .Toolstrip.Size = New Size(0, 0)
            End If
        End With
        Invalidate()
    End Sub
End Class
REM ////////////////////////////////////////////////////////////////////////////////////////////////////////// DROPDOWN COLLECTION
Public NotInheritable Class ItemCollection
    Inherits List(Of ComboItem)
    Public Event Changed(sender As Object)
    Public Sub New(Parent As ImageCombo)
        ImageCombo = Parent
    End Sub
    Public ReadOnly Property ImageCombo As ImageCombo
    Private HoldEvents As Boolean = False
    Public Shadows Function Item(TheName As String) As ComboItem

        Dim Items As New List(Of ComboItem)(From CI In Me Where CI.Name = TheName)
        If Items.Any Then
            Return Items.First
        Else
            Return Nothing
        End If

    End Function
    Public Overloads Function Add(Text As String) As ComboItem

        Dim ComboItem As New ComboItem With {.Value = Text}
        Add(ComboItem)
        Return ComboItem

    End Function
    Public Overloads Function Add(Text As String, Image As Image) As ComboItem

        Dim ComboItem As New ComboItem With {.Value = Text, .Image = Image}
        Add(ComboItem)
        Return ComboItem

    End Function
    Public Overloads Function AddRange(items As ItemCollection) As List(Of ComboItem)

        Dim newItems As New List(Of ComboItem)
        If items IsNot Nothing Then
            HoldEvents = True
            items.ForEach(Sub(item)
                              newItems.Add(item)
                          End Sub)
            HoldEvents = False
            RaiseEvent Changed(Me)
        End If
        Return newItems

    End Function
    Public Overloads Function Add(ComboItem As ComboItem) As ComboItem

        If ComboItem IsNot Nothing Then
            If 0 = 1 Then
                Const CheckWH As Integer = 14
                With ComboItem
                    ._ItemCollection = Me
                    Dim ItemHeight As Integer = TextRenderer.MeasureText("ZZZZZZZZZZZZZZZZZZZZ".ToString(InvariantCulture), ImageCombo.Font).Height
                    ._Index = Count
                    ._Bounds.X = 0
                    ._Bounds.Y = (ItemHeight * .Index)
                    ._Bounds.Width = ImageCombo.DropDown.Width - 3
                    ._Bounds.Height = ItemHeight
                    If ImageCombo.CheckboxStyle = CheckStyle.None Then
                        ._CheckBounds.X = 0
                        ._CheckBounds.Y = 0
                        ._CheckBounds.Width = 0
                        ._CheckBounds.Height = 0
                    Else
                        ._CheckBounds.X = 2
                        ._CheckBounds.Y = ._Bounds.Y + CInt((ItemHeight - CheckWH) / 2)
                        ._CheckBounds.Width = CheckWH
                        ._CheckBounds.Height = CheckWH
                    End If
                    If IsNothing(.Image) Then
                        ._ImageBounds.X = ._CheckBounds.Right
                        ._ImageBounds.Y = ._Bounds.Y
                        ._ImageBounds.Width = 0
                        ._ImageBounds.Height = 0
                    Else
                        Dim ImageWidth As Integer = {ItemHeight, .Image.Height}.Min
                        ._ImageBounds.X = ._CheckBounds.Right + 2
                        ._ImageBounds.Y = ._Bounds.Y + CInt((ItemHeight - ImageWidth) / 2)
                        ._ImageBounds.Width = ImageWidth
                        ._ImageBounds.Height = ImageWidth
                    End If
                    ._TextBounds.X = ._ImageBounds.Right + If(Not IsNothing(.Image), 2, 0)
                    ._TextBounds.Y = ._Bounds.Y
                    ._TextBounds.Width = ._Bounds.Width - ._TextBounds.X
                    ._TextBounds.Height = ItemHeight
                End With
            End If 'Don't think this is necessary
            MyBase.Add(ComboItem)
            If Not HoldEvents Then RaiseEvent Changed(Me)
        End If
        Return ComboItem

    End Function
End Class
REM ////////////////////////////////////////////////////////////////////////////////////////////////////////// DROPDOWN COMBO ITEM
<Serializable> <TypeConverter(GetType(PropertyConverter))> Public Class ComboItem
    Implements IDisposable
#Region " DISPOSE "
    Dim disposed As Boolean = False
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposed Then Return
        If disposing Then
            ' Free any other managed objects here.
            _Image?.Dispose()
        End If
        disposed = True
    End Sub
#End Region

    Public Sub New()
    End Sub

    Public ReadOnly Property ImageCombo As ImageCombo
        Get
            If IsNothing(_ItemCollection) Then
                Return Nothing
            Else
                Return _ItemCollection.ImageCombo
            End If
        End Get
    End Property
    <NonSerializedAttribute> Friend _ItemCollection As ItemCollection
    Public ReadOnly Property ItemCollection As ItemCollection
        Get
            Return _ItemCollection
        End Get
    End Property
    Private _Image As Image
    Public Property Image As Image
        Get
            Return _Image
        End Get
        Set(value As Image)
            If Not value Is _Image Then
                _Image = value
            End If
        End Set
    End Property
    Friend _Bounds As New Rectangle
    Public ReadOnly Property Bounds As Rectangle
        Get
            Return _Bounds
        End Get
    End Property
    Friend _CheckBounds As New Rectangle
    Public ReadOnly Property CheckBounds As Rectangle
        Get
            Return _CheckBounds
        End Get
    End Property
    Friend _ImageBounds As New Rectangle
    Public ReadOnly Property ImageBounds As Rectangle
        Get
            Return _ImageBounds
        End Get
    End Property
    Friend _TextBounds As New Rectangle
    Public ReadOnly Property TextBounds As Rectangle
        Get
            Return _TextBounds
        End Get
    End Property
    Public Property Format As String
    Public Property Value As Object
    Public Property Tag As Object
    Friend Color_ As Color = Color.Transparent
    Public ReadOnly Property Color As Color
        Get
            Return Color_
        End Get
    End Property
    Friend Font_ As Font = Nothing
    Public ReadOnly Property Font As Font
        Get
            Return Font_
        End Get
    End Property
    Public Property Name As String
    Public ReadOnly Property Text As String
        Get
            Return Microsoft.VisualBasic.Format(Value, Format)
        End Get
    End Property
    Friend _Index As Integer
    Public ReadOnly Property Index As Integer
        Get
            Return _Index
        End Get
    End Property
    Public Property Checked As Boolean
    Friend _Selected As Boolean
    Public ReadOnly Property Selected As Boolean
        Get
            Return _Selected
        End Get
    End Property
    Public Property Separator As Boolean
    Public Property TipText As String

    Public Overrides Function ToString() As String
        Return Join({Text, Index}, BlackOut)
    End Function
End Class
REM ////////////////////////////////////////////////////////////////////////////////////////////////////////// DROPDOWN SCROLLBAR
Public Class VerticalScrollBar
    Friend WithEvents Timer As New Timer With {.Interval = 250}
    Friend Alpha As Byte = 128
    Friend UpAlpha As Byte
    Friend DownAlpha As Byte
    Private Const Width As Integer = 12
    Private Const ShadowDepth As Integer = 8
    Private Const ArrowsHeight As Integer = Width + 2

    Public Sub New(Control As Control)

        Me.Control = Control
        If Control IsNot Nothing Then
            AddHandler Control.SizeChanged, AddressOf ControlSizeChanged
            AddHandler Control.MouseDown, AddressOf MouseDown
            AddHandler Control.MouseMove, AddressOf MouseMove
            AddHandler Control.MouseUp, AddressOf MouseUp
            AddHandler Control.MouseHover, AddressOf MouseHeld
        End If

    End Sub

    Private mScrolling As Boolean
    Friend ReadOnly Property Scrolling As Boolean
        Get
            Return mScrolling
        End Get
    End Property
    Private mReference As New Point
    Friend Property Reference As Point
        Get
            Return mReference
        End Get
        Set(value As Point)
            If Not mReference = value Then
                mReference = value
            End If
        End Set
    End Property
    Public Property Lines As Boolean
    Public Property Color As Color = Color.CornflowerBlue
    Public Property SmallChange As Integer = 1
    Public Property LargeChange As Integer
    Public ReadOnly Property Control As Control
    Friend ReadOnly Property Pages As List(Of Double)
        Get
            If Bounds.Height = 0 Then
                Return New List(Of Double)
            Else
                Dim PageCount As Double = ScrollHeight / Bounds.Height
                Return Enumerable.Range(0, CInt(Math.Floor(PageCount))).Select(Function(x) ArrowsHeight + (x * Bounds.Height) / 2).ToList
            End If
        End Get
    End Property
    Private _Value As Integer
    Public Property Value As Integer
        Get
            Return _Value
        End Get
        Set(value As Integer)
            If Not (value = _Value) Then
                If value < 0 Then
                    _Value = 0
                ElseIf (value) > ScrollHeight Then
                    _Value = ScrollHeight
                Else
                    _Value = value
                End If
                RaiseEvent ValueChanged(Me, Nothing)
            End If
        End Set
    End Property
    Private _Height As Integer
    Public Property Height As Integer
        Get
            Return _Height
        End Get
        Set(value As Integer)
            _Height = value
            UpdateBounds()
        End Set
    End Property
    Private _Maximum As Integer
    Public Property Maximum As Integer
        Get
            Return _Maximum
        End Get
        Set(value As Integer)
            UpdateBounds()
            _Maximum = value
        End Set
    End Property
    Friend ReadOnly Property ScrollHeight As Integer
        Get
            Return Maximum - Height
        End Get
    End Property
    Private _Bounds As New Rectangle(0, 0, Width, 0)
    Public ReadOnly Property Bounds As Rectangle
        Get
            Return _Bounds
        End Get
    End Property
    Private _TrackBounds As New Rectangle(0, ArrowsHeight, Width, 0)
    Public ReadOnly Property TrackBounds As Rectangle
        Get
            Return _TrackBounds
        End Get
    End Property
    Private _UpBounds As New Rectangle(0, -1, Width, ArrowsHeight)
    Friend ReadOnly Property UpBounds As Rectangle
        Get
            Return _UpBounds
        End Get
    End Property
    Private _BarBounds As New Rectangle(0, ArrowsHeight, Width, 0)
    Friend ReadOnly Property BarBounds As Rectangle
        Get
            If _BarBounds.Top <= UpBounds.Bottom Then
                _BarBounds.Y = UpBounds.Bottom
            ElseIf _BarBounds.Bottom >= DownBounds.Top Then
                _BarBounds.Y = (DownBounds.Top - _BarBounds.Height)
            End If
            Return _BarBounds
        End Get
    End Property
    Private _DownBounds As New Rectangle(0, 0, Width, ArrowsHeight)
    Friend ReadOnly Property DownBounds As Rectangle
        Get
            Return _DownBounds
        End Get
    End Property
    Friend ReadOnly Property Visible As Boolean
        Get
            Return ScrollHeight > Height
        End Get
    End Property

    Public Event ValueChanged(sender As Object, e As EventArgs)
    Private Sub ControlSizeChanged(sender As Object, e As EventArgs)
        UpdateBounds()
    End Sub
    Private Sub MouseDown(sender As Object, e As MouseEventArgs)
        If UpBounds.Contains(e.Location) Then
            Reference = e.Location
            Timer.Start()
            Value -= SmallChange
        ElseIf DownBounds.Contains(e.Location) Then
            Reference = e.Location
            Timer.Start()
            Value += SmallChange
        ElseIf TrackBounds.Contains(e.Location) Then
            If Not BarBounds.Contains(e.Location) Then
                Dim TrackValue As Double = ((e.Y - TrackBounds.Top) / (TrackBounds.Height - BarBounds.Height) * ScrollHeight)
                Value = Convert.ToInt32(Math.Floor(TrackValue / SmallChange) * SmallChange)
                _BarBounds.Y = e.Y
            End If
            Reference = e.Location
            Alpha = 255
            Control.Invalidate()
        End If
        Reference = e.Location
    End Sub
    Private Sub MouseHeld(sender As Object, e As EventArgs) Handles Timer.Tick
        If UpBounds.Contains(Reference) Then
            Value -= LargeChange
            _BarBounds.Y = Convert.ToInt32(Value * (TrackBounds.Height - BarBounds.Height) / ScrollHeight) + TrackBounds.Top
        ElseIf DownBounds.Contains(Reference) Then
            Value += LargeChange
            _BarBounds.Y = Convert.ToInt32(Value * (TrackBounds.Height - BarBounds.Height) / ScrollHeight) + TrackBounds.Top
        End If
        Control.Invalidate()
    End Sub
    Private Sub MouseMove(sender As Object, e As MouseEventArgs)
        Alpha = 60
        UpAlpha = 0
        DownAlpha = 0
        If Bounds.Contains(e.Location) Or Scrolling Then
            If e.Y < TrackBounds.Top Or e.Y > TrackBounds.Bottom Then
                mScrolling = False
            End If
            If TrackBounds.Contains(e.Location) Or Scrolling Then
                Timer.Stop()
                If e.Button = MouseButtons.Left Then
                    Alpha = 255
                    mScrolling = True
                    Dim Change As Integer = (e.Y - Reference.Y)
                    _BarBounds.Y += Change
                    Dim TrackValue As Double = ((BarBounds.Top - TrackBounds.Top) / (TrackBounds.Height - BarBounds.Height) * ScrollHeight)
                    Value = Convert.ToInt32(Math.Floor(TrackValue / SmallChange) * SmallChange)
                    Reference = e.Location
                Else
                    mScrolling = False
                    If BarBounds.Contains(e.Location) Then Alpha = 128
                End If
            ElseIf UpBounds.Contains(e.Location) Then
                UpAlpha = 64
            ElseIf DownBounds.Contains(e.Location) Then
                DownAlpha = 64
            End If
            Control.Invalidate()
        End If
    End Sub
    Private Sub MouseUp(sender As Object, e As MouseEventArgs)
        If Bounds.Contains(e.Location) Then
            Reference = Nothing
        End If
        mScrolling = False
        Control.Invalidate()
    End Sub
    Private Sub UpdateBounds()

        With _Bounds
            .X = Control.Width - Width - ShadowDepth
            .Height = Control.Height - ShadowDepth
            .Width = If(Visible, Width, 0)
        End With
        With _TrackBounds
            .X = _Bounds.X
            .Height = _Bounds.Height - (ArrowsHeight * 2)
            .Width = _Bounds.Width
        End With
        With _BarBounds
            .X = _Bounds.X - 1
            .Width = _Bounds.Width
            .Height = If(Visible, {Convert.ToInt32((Height / Maximum) * TrackBounds.Height), 20}.Max, 0)
        End With
        With _UpBounds
            .X = _Bounds.X - 1
            .Width = _Bounds.Width
        End With
        With _DownBounds
            .X = _Bounds.X - 1
            .Y = _Bounds.Height - ArrowsHeight
            .Width = _Bounds.Width
        End With

    End Sub
End Class