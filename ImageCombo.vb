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
    '[0 Image]       [1 Search]      [2 Text]     [3 Eye]       [4 Clear]       [5 DropDown]
    Private ImageBounds As New Rectangle
    Private ImageClickBounds As New Rectangle 'Full height
    Private SearchBounds As New Rectangle
    Friend TextBounds As New Rectangle
    Private TextMouseBounds As New Rectangle
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
    Private WithEvents CursorBlinkTimer As New Timer With {.Interval = 600}
    Private CursorShouldBeVisible As Boolean = True
    Private WithEvents TextTimer As New Timer With {.Interval = 250}
    Private InBounds As Boolean
    Private TextIsVisible As Boolean = True
    Private Const Spacing As Byte = 2
    Private KeyedValue As String
    Private LastValue As String
    Private ReadOnly GlossyDictionary As Dictionary(Of Theme, Image) = GlossyImages

    Friend Enum MouseRegion
        None
        Image
        Search
        Text
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
        CursorBlinkTimer.Start()

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
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            'e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
            If Mode = ImageComboMode.Button Then
#Region " BUTTON PROPERTIES "
                e.Graphics.DrawImage(If(InBounds, GlossyDictionary(If(ButtonMouseTheme = Theme.None, Theme.Gray, ButtonMouseTheme)), GlossyDictionary(If(ButtonTheme = Theme.None, Theme.Gray, ButtonTheme))), ClientRectangle)
                Dim penTangle As Rectangle = ClientRectangle
                penTangle.Inflate(-2, -2)
                penTangle.Offset(-1, -1)
                Using borderBrush As New SolidBrush(HighlightColor)
                    Using borderPen As New Pen(If(InBounds, borderBrush, Brushes.DarkGray), 3)
                        e.Graphics.DrawRectangle(borderPen, penTangle)
                    End Using
                End Using
                Using buttonAlignment As StringFormat = New StringFormat With {
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
            Else
#Region " REGULAR PROPERTIES "
                Using backBrush As New SolidBrush(BackColor)
                    e.Graphics.FillRectangle(backBrush, ClientRectangle)
                End Using
                If Text.Any Then
                    If PasswordProtected And Not TextIsVisible Then
                        Dim lettersRight As Integer = LetterWidths.Values.Last.Value
                        Using Brush As New HatchBrush(HatchStyle.LightUpwardDiagonal, SystemColors.WindowText, BackColor)
                            e.Graphics.FillRectangle(Brush, New Rectangle(ImageBounds.Width, 0, lettersRight - ImageBounds.Width, Height))
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
                    If HasFocus And CursorShouldBeVisible Then
                        Using Pen As New Pen(SelectionColor)
                            e.Graphics.DrawLine(Pen, CursorBounds.X, CursorBounds.Y, CursorBounds.X, CursorBounds.Bottom)
                        End Using
                    End If
                End If
#End Region
            End If

            If HighlightOnFocus And HasFocus Then
                Using Pen As New Pen(FocusColor, BorderWidth)
                    Dim BorderRectangle As Rectangle = ClientRectangle
                    BorderRectangle.Inflate(-BorderWidth, -BorderWidth)
                    e.Graphics.DrawRectangle(Pen, ClientRectangle)
                End Using

            Else
                Dim drawBorder As Boolean = Not BorderColor = Color.Transparent
                Dim penColor As Color = If(drawBorder, BorderColor, BackColor)
                Dim penWidth As Integer = If(drawBorder, BorderWidth, 4)
                Using Pen As New Pen(penColor, penWidth)
                    Dim BorderRectangle As Rectangle = ClientRectangle
                    BorderRectangle.Inflate(-penWidth, -penWidth)
                    e.Graphics.DrawRectangle(Pen, ClientRectangle)
                End Using

            End If
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

        Dim hasImage As Boolean = Image IsNot Nothing
        If Mode = ImageComboMode.Button Then
            With ImageBounds
                If hasImage Then
                    Dim Padding As Integer = {0, CInt((Height - Image.Height) / 2)}.Max     'Might be negative if Image.Height > Height
                    .X = Spacing
                    .Y = Padding
                    .Width = Image.Width
                    .Height = {Height, Image.Height}.Min
                Else
                    .X = Spacing
                    .Y = 0
                    .Width = 0
                    .Height = Height
                End If
                ImageClickBounds.X = .X : ImageClickBounds.Y = 0 : ImageClickBounds.Width = .Width : ImageClickBounds.Height = Height
            End With
            With TextBounds
                .X = ImageBounds.Right + Spacing          'LOOKS BETTER OFFSET BY A FEW PIXELS
                .Y = 0
                .Width = Width - ({ImageBounds.Width, DropBounds.Width, ClearTextBounds.Width, EyeBounds.Width}.Sum + Spacing + Spacing + Spacing)
                .Height = Height
                TextMouseBounds.X = .X : TextMouseBounds.Y = 0 : TextMouseBounds.Width = .Width : TextMouseBounds.Height = Height
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
                    Dim Padding As Integer = {0, CInt((Height - Image.Height) / 2)}.Max     'Might be negative if Image.Height > Height
                    .X = Spacing
                    .Y = Padding
                    .Width = Image.Width
                    .Height = {Height, Image.Height}.Min
                Else
                    .X = Spacing
                    .Y = 0
                    .Width = 0
                    .Height = Height
                End If
                ImageClickBounds.X = .X : ImageClickBounds.Y = 0 : ImageClickBounds.Width = .Width : ImageClickBounds.Height = Height
            End With
            With SearchBounds
                .X = ImageBounds.Right
                .Y = 0
                .Width = If(Mode = ImageComboMode.Searchbox, 16, 0)
                .Height = Height
            End With
            With DropBounds
                If hasDrop Then
                    'V LOOKS BETTER WHEN NOT RESIZED
                    Dim Padding As Integer = {0, CInt((Height - DropImage.Height) / 2)}.Max     'Might be negative if DropImage.Height > Height
                    .X = Width - (DropImage.Width + Spacing)
                    .Y = Padding
                    .Width = DropImage.Width
                    .Height = {Height, DropImage.Height}.Min
                    DropClickBounds.X = .X : DropClickBounds.Y = 0 : DropClickBounds.Width = .Width : DropClickBounds.Height = Height
                Else
                    .X = Width
                    .Y = 0
                    .Width = 0
                    .Height = Height
                    DropClickBounds = DropBounds
                End If
            End With
            With ClearTextBounds
                If hasClear Then
                    'X LOOKS BETTER WHEN NOT RESIZED
                    Dim Padding As Integer = {0, CInt((Height - ClearTextImage.Height) / 2)}.Max     'Might be negative if ClearTextImage.Height > Height
                    .X = Width - ({DropBounds.Width, ClearTextImage.Width}.Sum + Spacing)
                    .Y = Padding
                    .Width = ClearTextImage.Width
                    .Height = {Height, ClearTextImage.Height}.Min
                    ClearTextClickBounds.X = .X : ClearTextClickBounds.Y = 0 : ClearTextClickBounds.Width = .Width : ClearTextClickBounds.Height = Height
                Else
                    .X = DropBounds.Left
                    .Y = 0
                    .Width = 0
                    .Height = Height
                    ClearTextClickBounds = ClearTextBounds
                End If
            End With
            With EyeBounds
                If hasEye Then
                    Dim Padding As Integer = {0, CInt((Height - EyeImage.Height) / 2)}.Max     'Might be negative if EyeImage.Height > Height
                    .X = Width - ({DropBounds.Width, ClearTextBounds.Width, EyeImage.Width}.Sum + Spacing)
                    .Y = Padding
                    .Width = EyeImage.Width
                    .Height = {Height, EyeImage.Height}.Min
                    EyeClickBounds.X = .X : EyeClickBounds.Y = 0 : EyeClickBounds.Width = .Width : EyeClickBounds.Height = Height
                Else
                    .X = ClearTextBounds.Left
                    .Y = 0
                    .Width = 0
                    .Height = Height
                    EyeClickBounds = EyeBounds
                End If
            End With
            With TextBounds
                .X = SearchBounds.Right + Spacing          'LOOKS BETTER OFFSET BY A FEW PIXELS
                .Y = 0
                .Width = Width - ({ImageBounds.Width, SearchBounds.Width, DropBounds.Width, ClearTextBounds.Width, EyeBounds.Width}.Sum + Spacing + Spacing + Spacing)
                .Height = Height
                TextMouseBounds.X = .X : TextMouseBounds.Y = 0 : TextMouseBounds.Width = .Width : TextMouseBounds.Height = Height
            End With
            With CursorBounds
                .X = {Spacing, GetxPos(CursorIndex)}.Max
                .Y = Spacing
                .Width = 1
                .Height = Height - Spacing * 2
            End With
            With SelectionBounds
                .X = {GetxPos(SelectionIndex), CursorBounds.X}.Min
                .Y = Spacing
                .Width = Math.Abs(GetxPos(CursorIndex) - GetxPos(SelectionIndex))
                .Height = CursorBounds.Height
            End With
        End If

    End Sub

#End Region
#Region " PROPERTIES "
    Friend Property Mouse_Region As MouseRegion
    Public Property CheckOnSelect As Boolean = False
    Public Property CheckboxStyle As CheckStyle = CheckStyle.Slide
    Public Property BorderColor As Color = Color.Gainsboro
    Public Property BorderWidth As Byte = 2
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
    Private _DataType As Type
    Public ReadOnly Property DataType As Type
        Get
            Return GetDataType(Text)
        End Get
    End Property
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
                    HighlightOnFocus = True

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
    Public Property HighlightColor As Color = Color.LimeGreen
    Public Property HighlightOnFocus As Boolean
    Private ReadOnly Property HasFocus As Boolean = False
    Public Property FocusColor As Color = Color.DarkBlue
    Public Property SelectionColor As Color = Color.Black
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
    Private _SelectedIndex As Integer = -1
    Public Property SelectedIndex As Integer
        Get
            Return _SelectedIndex
        End Get
        Set(value As Integer)
            If Not (value < 0 Or value >= Items.Count) Then
                OnItemSelected(Items(value), False)
                If Not MultiSelect Then
                    Dim LastSelected As New List(Of ComboItem)(From CI In Items Where CI.Selected)
                    If LastSelected.Any Then LastSelected.First._Selected = False
                End If
                _SelectedIndex = value
                Items(value)._Selected = True
                Invalidate()
            End If
        End Set
    End Property
    Public ReadOnly Property LetterWidths As Dictionary(Of Integer, KeyValuePair(Of String, Integer))
        Get
            Dim widths As New Dictionary(Of Integer, KeyValuePair(Of String, Integer)) From {
                {0, New KeyValuePair(Of String, Integer)("Spacing", TextBounds.X)}
            }
            If Text.Any Then
                For i As Integer = 1 To Text.Length
                    Dim letter As String = Text.Substring(0, i)
                    widths.Add(i, New KeyValuePair(Of String, Integer)(letter, TextLength(letter)))
                Next
            End If
            Return widths
        End Get
    End Property
    Public Property SelectionStart As Integer
        Get
            Return {CursorIndex, SelectionIndex}.Min
        End Get
        Set(value As Integer)
            CursorIndex = value
        End Set
    End Property
    Public Property SelectionLength As Integer
        Get
            Return Math.Abs(CursorIndex - SelectionIndex)
        End Get
        Set(value As Integer)
            SelectionIndex = SelectionStart + value
        End Set
    End Property
    Public Property SelectionEnd As Integer
        Get
            Return {CursorIndex, SelectionIndex}.Min
        End Get
        Set(value As Integer)
            SelectionIndex = value
        End Set
    End Property
    Private _CursorIndex As Integer
    Private Property CursorIndex As Integer
        Get
            Return _CursorIndex
        End Get
        Set(value As Integer)
            If value <> _CursorIndex And value >= 0 And value < LetterWidths.Count Then
                _CursorIndex = value
                Invalidate()
            End If
        End Set
    End Property
    Private _SelectionIndex As Integer
    Public Property SelectionIndex As Integer
        Get
            Return _SelectionIndex
        End Get
        Set(value As Integer)
            If value <> _SelectionIndex And value >= 0 And value < LetterWidths.Count Then
                _SelectionIndex = value
                Invalidate()
            End If
        End Set
    End Property
    Public ReadOnly Property Selection As String
        Get
            If CursorIndex = SelectionIndex Then
                Return String.Empty
            Else
                Return Text.Substring(SelectionStart, SelectionLength)
            End If
        End Get
    End Property
    Private SelectedColor_ As Color
    Public Property SelectedColor As Color
        Get
            SelectedColor_ = If(Mode = ImageComboMode.ColorPicker And SelectionIndex >= 0, SelectedItem.Color, Color.Transparent)
            Return SelectedColor_
        End Get
        Set(value As Color)
            If value <> SelectedColor_ Then
                SelectedColor_ = value
                Dim indexItem As Integer = 0
                For Each item In Items
                    If item.Color = value Then
                        _SelectedIndex = indexItem
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
            If Mode = ImageComboMode.FontPicker And SelectionIndex >= 0 Then
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
                SelectionIndex = indexItem
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
            If TypeOf DataSource Is IEnumerable Then
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

        If Not ComboItem.Index = SelectedIndex Then
            _SelectedIndex = ComboItem.Index
            Text = ComboItem.Text
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
    Private Sub BorderDraw(sender As Object, e As EventArgs) Handles Me.GotFocus
        _HasFocus = True
        Invalidate()
    End Sub
    Private Sub BorderNoDraw(sender As Object, e As EventArgs) Handles Me.LostFocus
        _HasFocus = False
        Invalidate()
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
        CursorShouldBeVisible = True
        CursorBlinkTimer.Start()
        TextTimer.Start()

        If e IsNot Nothing Then
            Try
                Dim S As Integer = CursorIndex
                If e.KeyCode = Keys.Left Or e.KeyCode = Keys.Right Then
#Region " MOVE LEFT Or RIGHT "
                    Dim Value As Integer = If(e.KeyCode = Keys.Left, -1, 1)
                    If Control.ModifierKeys = Keys.Shift Then
                        SelectionIndex += Value
                    Else
                        CursorIndex += Value
                        SelectionIndex = CursorIndex
                    End If
#End Region
                ElseIf e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete And Not IsReadOnly Then
#Region " REMOVE BACK Or AHEAD "
                    If CursorIndex = SelectionIndex Then
                        If e.KeyCode = Keys.Back Then
                            If Not S = 0 Then
                                CursorIndex -= 1
                                SelectionIndex = CursorIndex
                                Text = Text.Remove(S - 1, 1)
                            End If
                        ElseIf e.KeyCode = Keys.Delete Then
                            If Not S = Text.Length Then
                                Text = Text.Remove(S, 1)
                            End If
                        End If
                    Else
                        Dim TextLength As Integer = SelectionLength
                        CursorIndex = SelectionStart
                        SelectionIndex = CursorIndex
                        Text = Text.Remove(SelectionStart, TextLength)
                    End If
#End Region
                ElseIf e.KeyCode = Keys.A AndAlso Control.ModifierKeys = Keys.Control Then
#Region " SELECT ALL "
                    SelectAll()
#End Region
                ElseIf e.KeyCode = Keys.X AndAlso Control.ModifierKeys = Keys.Control And Not IsReadOnly Then
#Region " CUT "
                    Dim TextSelection As String = Selection
                    CursorIndex = SelectionStart
                    SelectionIndex = CursorIndex
                    Clipboard.SetText(TextSelection)
                    Text = Text.Remove(SelectionStart, TextSelection.Length)
                    RaiseEvent TextCopied(Me, Nothing)
#End Region
                ElseIf e.KeyCode = Keys.C AndAlso Control.ModifierKeys = Keys.Control Then
#Region " COPY "
                    Clipboard.Clear()
                    If Selection.Any Then
                        Clipboard.SetText(Selection)
                    Else
                        Clipboard.Clear()
                    End If
                    RaiseEvent TextCopied(Me, Nothing)
#End Region
                ElseIf e.KeyCode = Keys.V AndAlso Control.ModifierKeys = Keys.Control And Not IsReadOnly Then
#Region " PASTE "
                    S = SelectionStart
                    Text = Text.Remove(SelectionStart, SelectionLength)
                    Dim ClipboardText As String = Nothing
                    Try
                        ClipboardText = Clipboard.GetText()
                        Text = Text.Insert(S, ClipboardText)
                        CursorIndex = S + ClipboardText.Length
                        SelectionIndex = CursorIndex
                        RaiseEvent TextPasted(Me, Nothing)
                    Catch ex As Runtime.InteropServices.ExternalException
                    End Try
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
                End If
                KeyedValue = Text
                Invalidate()

            Catch ex As IndexOutOfRangeException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace)

            End Try
        End If
        MyBase.OnKeyDown(e)

    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)

        If IsReadOnly Then Exit Sub
        CursorShouldBeVisible = True
        CursorBlinkTimer.Start()
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
                    _CursorIndex = SelectionStart + 1
                    _SelectionIndex = CursorIndex
                    Text = ProposedText
                    ShowMatches(Text)
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
            Bounds_Set()
            Dim redraw As Boolean = False
            Dim xy As Point = e.Location
            Dim lastMouseRegion As MouseRegion = Mouse_Region
            Mouse_Region = If(ImageClickBounds.Contains(xy), MouseRegion.Image, If(SearchBounds.Contains(xy), MouseRegion.Search, If(TextMouseBounds.Contains(xy), MouseRegion.Text, If(EyeClickBounds.Contains(xy), MouseRegion.Eye, If(ClearTextClickBounds.Contains(xy), MouseRegion.ClearText, If(DropClickBounds.Contains(xy), MouseRegion.DropDown, MouseRegion.None))))))
            redraw = Mouse_Region <> lastMouseRegion
            If MouseLeftDown.Key And e.Button = MouseButtons.Left Then
                Dim startEnd As Integer() = {CursorIndex, GetLetterIndex(e.X)}
                If MouseLeftDown.Value <> startEnd.Last Then
                    'Moved to left or right
                    Dim leftMost As Integer = startEnd.Min
                    Dim rightMost As Integer = startEnd.Max
                    _CursorIndex = leftMost
                    SelectionIndex = rightMost
                    redraw = True
                End If
            End If
            If redraw Then Invalidate()
        End If
        MyBase.OnMouseMove(e)

    End Sub
    Private MouseLeftDown As New KeyValuePair(Of Boolean, Integer)
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
                CursorBlinkTimer.Stop()
                CursorBlinkTimer.Start()
                CursorIndex = GetLetterIndex(e.X)
                SelectionIndex = CursorIndex
                If e.Button = MouseButtons.Left Then MouseLeftDown = New KeyValuePair(Of Boolean, Integer)(True, SelectionIndex)

            ElseIf Mouse_Region = MouseRegion.Eye Then
                TextIsVisible = Not TextIsVisible

            ElseIf Mouse_Region = MouseRegion.ClearText Then
                RaiseEvent ClearTextClicked(Me, New ImageComboEventArgs)
                If IsReadOnly Then Exit Sub
                Text = String.Empty
                SelectedIndex = -1

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
        If e IsNot Nothing Then
            If Not Mouse_Region = MouseRegion.DropDown And LetterWidths.Any Then
                Dim CurrentIndex As Integer = GetLetterIndex(e.X)
                Dim Index As Integer = {CurrentIndex, Text.Length - 1}.Min
                '// look back
                Do While (Index >= 0 AndAlso Text.Substring(Index, 1) <> " ")
                    Index -= 1
                Loop
                CursorIndex = Index + 1
                '// look ahead
                Index = CurrentIndex
                Do While Index < Text.Length AndAlso Text.Substring(Index, 1) <> " "
                    Index += 1
                Loop
                SelectionIndex = Index
                MoveMouse(Cursor.Position)
                Invalidate()
            End If
        End If
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
            CursorIndex = 0
            SelectionIndex = 0
        End If
        Bounds_Set()
        Try
            If ErrorTip IsNot Nothing Then
                If ValueError Then
                    'ErrorTip.Show(ErrorText.ToString(InvariantCulture), Me, New Point(Width, 0))
                Else
                    'ErrorTip.Hide(Me)
                End If
            End If
        Catch ex As ObjectDisposedException
        End Try
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
    Private Function GetxPos(Index As Integer) As Integer
        Return LetterWidths({0, {Text.Length, Index}.Min}.Max).Value
    End Function
    Private Function GetLetterIndex(X As Integer) As Integer
        Return (From lw In LetterWidths.Keys Where LetterWidths(lw).Value <= {X, TextBounds.X}.Max Select lw).Max
    End Function
    Private Function TextLength(T As String) As Integer

        Dim Padding As Integer = If(T.Length = 0, 0, (2 * TextRenderer.MeasureText(T.First, Font).Width) - TextRenderer.MeasureText(T.First & T.First, Font).Width)
        Return TextBounds.Left + TextRenderer.MeasureText(T, Font).Width - Padding

    End Function
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
    Public Sub SelectAll()

        CursorIndex = 0
        Mouse_Region = MouseRegion.Text
        SelectionStart = 0
        SelectionIndex = LetterWidths.Keys.Last
        MoveMouse(Cursor.Position)

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
                ElseIf Bounds.Contains(e.Location) And Not VScroll.Bounds.Contains(e.Location) Then
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
                            Selected.First._Selected = Not (Selected.First.Selected)
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
                Dim SV As IEnumerable(Of ComboItem) = From S In MatchedItems Where S.Index = ComboParent.SelectionIndex
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