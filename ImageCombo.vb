Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.ComponentModel
Public NotInheritable Class ImageComboEventArgs
    Inherits EventArgs
    Public Property ComboItem As ComboItem
    Public Sub New(ByVal TheComboItem As ComboItem)
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
End Enum
Public Enum OperandSign
    GreaterThan
    LessThan
    Equals
    NotEquals
End Enum
Public NotInheritable Class ImageCombo
    Inherits Control
    'Fixes:     Screen Scaling of 125, 150 distorts the CopyFromScreen in DropDown.Protected Overrides Sub OnVisibleChanged(e As EventArgs)
    Private ReadOnly ErrorTip As New ToolTip With {
        .BackColor = Color.GhostWhite,
        .ForeColor = Color.Black,
        .ShowAlways = False}
    Friend Toolstrip As New ToolStripDropDown With {.AutoClose = False, .AutoSize = False, .Padding = New Padding(0), .DropShadowEnabled = False, .BackColor = Color.Transparent}
    Friend Mouse_Region As New MouseRegion
    Private ImageBounds As New Rectangle
    Friend TextBounds As New Rectangle
    Private ReadOnly EyeImage As Image
    Private EyeBounds As New Rectangle
    Private EyeDrawBounds As New Rectangle
    Private ReadOnly ClearTextImage As Image
    Private ClearTextBounds As New Rectangle
    Private ClearTextDrawBounds As New Rectangle
    Private ReadOnly DropImage As Image
    Private DropBounds As New Rectangle
    Private DropDrawBounds As New Rectangle
    Private CursorBounds As New Rectangle
    Private SelectionBounds As New Rectangle
    Private WithEvents CursorBlinkTimer As New Timer With {.Interval = 600}
    Private CursorShouldBeVisible As Boolean = True
    Private WithEvents TextTimer As New Timer With {.Interval = 250}
    Private InBounds As Boolean
    Private TextIsVisible As Boolean = True
    Private Const Spacing As Byte = 2
    Private ReadOnly GlossyDictionary As Dictionary(Of Theme, Image) = GlossyImages
#Region " STRUCTURES + ENUMS "
    Enum MouseRegion
        None
        Image
        Text
        Eye
        ClearText
        DropDown
    End Enum
    <Flags> Enum ValueTypes
        Any
        Integers
        Decimals
    End Enum
#End Region
#Region " INITIALIZE "
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
        BorderStyle = Border3DStyle.Flat
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
#End Region
    Private Sub CursorBlinkTimer_Tick() Handles CursorBlinkTimer.Tick

        CursorShouldBeVisible = Not CursorShouldBeVisible
        Invalidate()

    End Sub
    Private Sub TextTimer_Tick() Handles TextTimer.Tick

        TextTimer.Stop()
        RaiseEvent TextPaused(Me, New ImageComboEventArgs)

    End Sub
#Region " DRAWING "
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If e IsNot Nothing Then
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

            Dim hasImage As Boolean = Image IsNot Nothing

            If Mode = ImageComboMode.Button Then
#Region " BUTTON STYLE - ADD COLORS "
                e.Graphics.DrawImage(If(InBounds, GlossyDictionary(If(ButtonMouseTheme = Theme.None, Theme.Gray, ButtonMouseTheme)), GlossyDictionary(If(ButtonTheme = Theme.None, Theme.Gray, ButtonTheme))), ClientRectangle)
                Dim penTangle As Rectangle = ClientRectangle
                penTangle.Inflate(-2, -2)
                penTangle.Offset(-1, -1)
                Using borderBrush As New SolidBrush(HighlightColor)
                    Using borderPen As New Pen(If(InBounds, borderBrush, Brushes.DarkGray), 3)
                        e.Graphics.DrawRectangle(borderPen, penTangle)
                    End Using
                End Using
#Region " SET BOUNDS "
                With ImageBounds
                    If hasImage Then
                        Dim Padding As Integer = {0, Convert.ToInt32((Height - Image.Height) / 2)}.Max     'Might be negative if Image.Height > Height
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
                End With
                With TextBounds
                    .X = ImageBounds.Right + Spacing          'LOOKS BETTER OFFSET BY A FEW PIXELS
                    .Y = 0
                    .Width = Width - ({ImageBounds.Width, DropBounds.Width, ClearTextBounds.Width, EyeBounds.Width}.Sum + Spacing + Spacing + Spacing)
                    .Height = Height
                End With
#End Region
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
#Region " SET BOUNDS "
                Dim hasText As Boolean = Text.Any
                Dim hasDrop As Boolean = Not Mode = ImageComboMode.Button And Items.Any
                Dim hasClear As Boolean = Not Mode = ImageComboMode.Button And hasText
                Dim hasEye As Boolean = PasswordProtected And hasClear
                '===========================
                If AutoSize Then
                    Dim textSize As Size = If(hasText, MeasureText(Text, Font), New Size)
                    Dim imageSize As Size = If(hasImage, Image.Size, New Size)
                    Dim dropSize As Size = If(hasDrop, DropImage.Size, New Size)
                    Dim clearSize As Size = If(hasClear, ClearTextImage.Size, New Size)
                    Dim eyeSize As Size = If(hasEye, EyeImage.Size, New Size)
                    '===========================
                    Dim sizes As New List(Of Size) From {textSize, imageSize, dropSize, clearSize, eyeSize}
                    Dim widths As New List(Of Integer)(From s In sizes Where Not s.Width = 0 Select s.Width)
                    Dim heights As New List(Of Integer)(From s In sizes Where Not s.Height = 0 Select s.Height)
                    '===========================
                    Dim minSize As Size = If(MinimumSize.IsEmpty, New Size(60, 24), MinimumSize)
                    Dim maxSize As Size = If(MaximumSize.IsEmpty, WorkingArea.Size, MaximumSize)
                    Dim minmaxWidth As Integer = {{If(widths.Any, widths.Sum + Spacing * (widths.Count + 1), minSize.Width), minSize.Width}.Max, maxSize.Width}.Min
                    Dim minmaxHeight As Integer = {{If(heights.Any, Spacing + heights.Max + Spacing, minSize.Height), minSize.Height}.Max, maxSize.Height}.Min
                    Dim newSize As New Size(minmaxWidth, minmaxHeight)
                    Size = newSize
                End If
                If Not hasText Then
                    CursorIndex = 0
                    SelectionIndex = 0
                End If
                With ImageBounds
                    If hasImage Then
                        Dim Padding As Integer = {0, Convert.ToInt32((Height - Image.Height) / 2)}.Max     'Might be negative if Image.Height > Height
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
                End With
                With DropBounds
                    If hasDrop Then
                        'V LOOKS BETTER WHEN NOT RESIZED
                        Dim Padding As Integer = {0, Convert.ToInt32((Height - DropImage.Height) / 2)}.Max     'Might be negative if DropImage.Height > Height
                        .X = Width - (DropImage.Width + Spacing)
                        .Y = Padding
                        .Width = DropImage.Width
                        .Height = {Height, DropImage.Height}.Min
                        DropDrawBounds.X = .X : DropDrawBounds.Y = 0 : DropDrawBounds.Width = .Width : DropDrawBounds.Height = Height
                    Else
                        .X = Width
                        .Y = 0
                        .Width = 0
                        .Height = Height
                        DropDrawBounds = DropBounds
                    End If
                End With
                With ClearTextBounds
                    If hasClear Then
                        'X LOOKS BETTER WHEN NOT RESIZED
                        Dim Padding As Integer = {0, Convert.ToInt32((Height - ClearTextImage.Height) / 2)}.Max     'Might be negative if ClearTextImage.Height > Height
                        .X = Width - ({DropBounds.Width, ClearTextImage.Width}.Sum + Spacing)
                        .Y = Padding
                        .Width = ClearTextImage.Width
                        .Height = {Height, ClearTextImage.Height}.Min
                        ClearTextDrawBounds.X = .X : ClearTextDrawBounds.Y = 0 : ClearTextDrawBounds.Width = .Width : ClearTextDrawBounds.Height = Height
                    Else
                        .X = DropBounds.Left
                        .Y = 0
                        .Width = 0
                        .Height = Height
                        ClearTextDrawBounds = ClearTextBounds
                    End If
                End With
                With EyeBounds
                    If hasEye Then
                        Dim Padding As Integer = {0, Convert.ToInt32((Height - EyeImage.Height) / 2)}.Max     'Might be negative if EyeImage.Height > Height
                        .X = Width - ({DropBounds.Width, ClearTextBounds.Width, EyeImage.Width}.Sum + Spacing)
                        .Y = Padding
                        .Width = EyeImage.Width
                        .Height = {Height, EyeImage.Height}.Min
                        EyeDrawBounds.X = .X : EyeDrawBounds.Y = 0 : EyeDrawBounds.Width = .Width : EyeDrawBounds.Height = Height
                    Else
                        .X = ClearTextBounds.Left
                        .Y = 0
                        .Width = 0
                        .Height = Height
                        EyeDrawBounds = EyeBounds
                    End If
                End With
                With TextBounds
                    .X = ImageBounds.Right + Spacing          'LOOKS BETTER OFFSET BY A FEW PIXELS
                    .Y = 0
                    .Width = Width - ({ImageBounds.Width, DropBounds.Width, ClearTextBounds.Width, EyeBounds.Width}.Sum + Spacing + Spacing + Spacing)
                    .Height = Height
                End With
                With CursorBounds
                    .X = {Spacing, GetxPos(CursorIndex)}.Max
                    .Y = Spacing
                    .Width = 1
                    .Height = Height - Spacing * 2
                End With
                With SelectionBounds
                    .X = {GetxPos(SelectionIndex), CursorBounds.X}.Min
                    .Y = spacing
                    .Width = Math.Abs(GetxPos(CursorIndex) - GetxPos(SelectionIndex))
                    .Height = CursorBounds.Height
                End With
#End Region
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
                    e.Graphics.DrawImage(EyeImage, EyeBounds)
                    e.Graphics.DrawImage(ClearTextImage, ClearTextBounds)
                    e.Graphics.DrawImage(DropImage, DropBounds)
                    Dim HighlightRectangle As Rectangle = Nothing
                    If Mouse_Region = MouseRegion.Image Then HighlightRectangle = New Rectangle(ImageBounds.X, 0, ImageBounds.Width, Height)
                    If Mouse_Region = MouseRegion.Eye Then HighlightRectangle = EyeDrawBounds
                    If Mouse_Region = MouseRegion.ClearText Then HighlightRectangle = ClearTextDrawBounds
                    If Mouse_Region = MouseRegion.DropDown Then HighlightRectangle = DropDrawBounds
                    Using mouseOverBrush As New SolidBrush(Color.FromArgb(60, SelectionColor))
                        e.Graphics.FillRectangle(mouseOverBrush, HighlightRectangle)
                    End Using
                    If _HasFocus And CursorShouldBeVisible Then
                        Using Pen As New Pen(SelectionColor)
                            e.Graphics.DrawLine(Pen, CursorBounds.X, CursorBounds.Y, CursorBounds.X, CursorBounds.Bottom)
                        End Using
                    End If
                End If
#End Region
            End If
            If Image IsNot Nothing Then e.Graphics.DrawImage(Image, ImageBounds)
            ControlPaint.DrawBorder3D(e.Graphics, ClientRectangle, BorderStyle)
            If HighlightOnFocus And _HasFocus Then
                Dim PenWidth As Integer = 2
                Using Pen As New Pen(Color.FromArgb(64, FocusColor), PenWidth)
                    Dim BorderRectangle As Rectangle = ClientRectangle
                    BorderRectangle.Inflate(-PenWidth, -PenWidth)
                    e.Graphics.DrawRectangle(Pen, BorderRectangle)
                End Using
            End If
        End If

    End Sub
#End Region
#Region " PROPERTIES "
    Public Property CheckBoxes As Boolean = True
    Public Property CheckOnSelect As Boolean = False
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
    Private ReadOnly OperandDictionary As New Dictionary(Of OperandSign, Bitmap)
    Private ReadOnly Operands As New Dictionary(Of String, OperandSign) From
                    {
            {"≥", OperandSign.GreaterThan},
            {"≤", OperandSign.LessThan},
            {"=", OperandSign.Equals},
            {"≠", OperandSign.NotEquals}
            }
    Public ReadOnly Property OperandItem As OperandSign
    Public ReadOnly Property OperandString As String
        Get
            Select Case OperandItem
                Case OperandSign.Equals
                    Return "="
                Case OperandSign.NotEquals
                    Return "¬="
                Case OperandSign.GreaterThan
                    Return ">="
                Case OperandSign.LessThan
                    Return "<="
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property
    Public ReadOnly Property ErrorText As String
        Get
            If ValueError Then
                If AcceptValues = ValueTypes.Decimals Then
                    If Amount > MaxAcceptValue Then
                        Return Join({"Amount exceeds maximum value of", MaxAcceptValue})

                    ElseIf Amount < MinAcceptValue Then
                        Return Join({"Amount is below the minimum value of", MinAcceptValue})

                    Else
                        Return Join({"Typed value is not recognized as a decimal"})
                    End If

                ElseIf AcceptValues = ValueTypes.Integers Then
                    If Number > MaxAcceptValue Then
                        Return Join({"Number exceeds maximum value of", CInt(MaxAcceptValue)})

                    ElseIf Number < MinAcceptValue Then
                        Return Join({"Number is below the minimum value of", CInt(MaxAcceptValue)})

                    Else
                        Return Join({"Typed value is not recognized as a whole number"})
                    End If

                Else 'Any ... duhh
                    If Text.Length < MinAcceptValue Then
                        Return Join({"Length of value is below", MinAcceptValue, "characters"})
                    Else
                        Return Join({"Length of value exceeds", MaxAcceptValue, "characters"})
                    End If
                End If
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property ValueError As Boolean
        Get
            If Text.Any Then
                Dim canParse As Boolean
                If AcceptValues = ValueTypes.Decimals Then
                    Dim amount As Double
                    canParse = Double.TryParse(Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), amount)
                    Return If(canParse, amount > MaxAcceptValue Or amount < MinAcceptValue, True)

                ElseIf AcceptValues = ValueTypes.Integers Then
                    Dim number As Long
                    canParse = Long.TryParse(Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), number)
                    Return If(canParse, number > MaxAcceptValue Or number < MinAcceptValue, True)

                Else
                    Return Text.Length > MaxAcceptValue Or Text.Length < MinAcceptValue

                End If
            Else
                Return False
            End If
        End Get
    End Property
    Public Property AcceptValues As ValueTypes
    Public Property MaxAcceptValue As Double = Double.MaxValue
    Public Property MinAcceptValue As Double = -Double.MaxValue
    Public ReadOnly Property Amount As Double
        Get
            Dim _Amount As Double
            Dim canParse As Boolean = Double.TryParse(Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), _Amount)
            If canParse Then
                Return _Amount
            Else
                Return 0
            End If
        End Get
    End Property
    Public ReadOnly Property Number As Long
        Get
            Dim _Number As Long
            Dim canParse As Boolean = Long.TryParse(Text, Globalization.NumberStyles.Any, Globalization.CultureInfo.CreateSpecificCulture("en-US"), _Number)
            If canParse Then
                Return _Number
            Else
                Return 0
            End If
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return "ImageCombo.Text=""" & Text & """"
    End Function
    Private KeyedValue As String, LastValue As String
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
                    CheckBoxes = False
                    For Each colorItem In ColorImages()
                        Dim item As ComboItem = Items.Add(colorItem.Key.Name, colorItem.Value)
                        item.Tag = colorItem.Key
                    Next

                ElseIf value = ImageComboMode.FontPicker Then
                    CheckBoxes = False
                    For Each fontItem In FontImages()
                        Dim item As ComboItem = Items.Add(fontItem.Key.Name, fontItem.Value)
                        item.Tag = fontItem.Key
                    Next

                ElseIf value = ImageComboMode.Button Then
                    HighlightOnFocus = True

                ElseIf value = ImageComboMode.Searchbox Then
                    OperandDictionary.Clear()
                    For Each item In Operands
                        Dim bmpOperand As Bitmap = New Bitmap(20, 20)
                        Using g As Graphics = Graphics.FromImage(bmpOperand)
                            With g
                                Using backBrush As New SolidBrush(Color.Transparent)
                                    Dim bmpBounds As New Rectangle(0, 0, bmpOperand.Width, bmpOperand.Height)
                                    g.FillRectangle(backBrush, bmpBounds)
                                    Using sf As New StringFormat With {
                                        .Alignment = StringAlignment.Center,
                                        .LineAlignment = StringAlignment.Center
                                        }
                                        g.DrawString(item.Key, New Font("IBM Plex Mono Medium", 16), Brushes.Black, bmpBounds, sf)
                                    End Using
                                End Using
                            End With
                        End Using
                        OperandDictionary.Add(item.Value, bmpOperand)
                    Next
                    _OperandItem = OperandSign.Equals
                End If
                Mode_ = value
                Invalidate()
            End If
        End Set
    End Property
    Public Property IsReadOnly As Boolean
    Public ReadOnly Property DropDown As ImageComboDropDown
    Public ReadOnly Property Items As ItemCollection
    Private _DropItems As New List(Of ComboItem)
    Public ReadOnly Property DropItems As List(Of ComboItem)
        Get
            If Not _DropItems.Any Then
                _DropItems = Items
            End If
            Return _DropItems
        End Get
    End Property
    Public Property HighlightColor As Color = Color.LimeGreen
    Public Property HighlightOnFocus As Boolean
    Private _HasFocus As Boolean = False
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
            Return If(overrideImage, If(Mode = ImageComboMode.Searchbox, OperandDictionary(OperandItem), _Image))
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
            If SelectedIndex < 0 Then
                Return Nothing
            Else
                If Items.Any Then
                    If SelectedIndex < Items.Count Then
                        Return Items(SelectedIndex)
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End If
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
            If (value <> _CursorIndex And value >= 0 And value < LetterWidths.Count) Then
                _CursorIndex = value
                Invalidate()
            End If
        End Set
    End Property
    Private _SelectionIndex As Integer
    Friend Property SelectionIndex As Integer
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
    Public Property BorderStyle As Border3DStyle = Border3DStyle.Adjust
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
                                Dim List As New List(Of String)(From Element In Type Select Convert.ToString(Element, InvariantCulture))
                                List.Sort(Function(x, y) String.Compare(x, y, StringComparison.InvariantCulture))
                                For Each Item As String In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Date)
                                Dim List As New List(Of Date)(From Element In Type Select Convert.ToDateTime(Element, InvariantCulture).Date)
                                List.Sort(Function(x, y) y.CompareTo(x))
                                For Each Item As Date In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Boolean)
                                Dim List As New List(Of Boolean)(From Element In Type Select Convert.ToBoolean(Element, InvariantCulture))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Boolean In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Decimal), GetType(Double)
                                Dim List As New List(Of Decimal)(From Element In Type Select Convert.ToDecimal(Element, InvariantCulture))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Decimal In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Long)
                                Dim List As New List(Of Long)(From Element In Type Select Convert.ToInt64(Element, InvariantCulture))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Long In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Integer)
                                Dim List As New List(Of Integer)(From Element In Type Select Convert.ToInt32(Element, InvariantCulture))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Integer In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Short)
                                Dim List As New List(Of Short)(From Element In Type Select Convert.ToInt16(Element, InvariantCulture))
                                List.Sort(Function(x, y) x.CompareTo(y))
                                For Each Item As Short In List.Distinct
                                    Items.Add(New ComboItem With {.Value = Item})
                                Next

                            Case GetType(Byte)
                                Dim List As New List(Of Byte)(From Element In Type Select Convert.ToByte(Element, InvariantCulture))
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
    Public Event ImageClicked(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Public Event ClearTextClicked(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Public Event ValueSubmitted(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Public Event TextPaused(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Public Event ValueChanged(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Public Event SelectionChanged(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Public Event ItemSelected(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Public Event DropDownOpened(ByVal sender As Object, ByVal e As EventArgs)
    Public Event DropDownClosed(ByVal sender As Object, ByVal e As EventArgs)
    Public Event TextPasted(ByVal sender As Object, e As EventArgs)
    Public Event TextCopied(ByVal sender As Object, e As EventArgs)
    Friend Sub OnItemSelected(ByVal ComboItem As ComboItem, DropDownVisible As Boolean)

        If Not ComboItem.Index = SelectedIndex Then
            _SelectedIndex = ComboItem.Index
            Text = ComboItem.Text
            RaiseEvent SelectionChanged(Me, New ImageComboEventArgs(ComboItem))
            RaiseEvent ValueChanged(Me, New ImageComboEventArgs(ComboItem))
        End If
        DropDown.Visible = DropDownVisible
        RaiseEvent ItemSelected(Me, New ImageComboEventArgs(ComboItem))

    End Sub
    Public Event ItemChecked(ByVal sender As Object, ByVal e As ImageComboEventArgs)
    Friend Sub OnItemChecked(ByVal ComboItem As ComboItem)
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
    Private Sub On_PreviewKeyDown(ByVal sender As Object, ByVal e As PreviewKeyDownEventArgs)

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
    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)

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
                    Dim ClipboardText As String = Clipboard.GetText
                    Text = Text.Insert(S, ClipboardText)
                    CursorIndex = S + ClipboardText.Length
                    SelectionIndex = CursorIndex
                    RaiseEvent TextPasted(Me, Nothing)
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
    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)

        If IsReadOnly Then Exit Sub
        If e IsNot Nothing Then
            If Not Mouse_Region = MouseRegion.DropDown And LetterWidths.Any Then
                Dim CurrentIndex As Integer = GetLetterIndex(e.X)
                Dim Index As Integer = {CurrentIndex, Text.Length - 1}.Min
#Region " LOOK BACK "
                Do While (Index >= 0 AndAlso Text.Substring(Index, 1) <> " ")
                    Index -= 1
                Loop
                CursorIndex = Index + 1
#End Region
                Index = CurrentIndex
#Region " LOOK AHEAD "
                Do While (Index < Text.Length AndAlso Text.Substring(Index, 1) <> " ")
                    Index += 1
                Loop
#End Region
                SelectionIndex = Index
                Invalidate()
            End If
        End If
        MyBase.OnMouseDoubleClick(e)

    End Sub
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            If Mouse_Region = MouseRegion.Eye Then
                TextIsVisible = Not (TextIsVisible)

            ElseIf Mouse_Region = MouseRegion.Image Then
                If ValueError Then
                    Text = String.Empty
                    SelectedIndex = -1
                    Invalidate()
                Else
                    RaiseEvent ImageClicked(Me, New ImageComboEventArgs)
                    If Mode = ImageComboMode.Searchbox Then
                        Dim currentIndex As Integer = Operands.Values.ToList.IndexOf(OperandItem)
                        Dim nextIndex As Integer = (currentIndex + 1) Mod 3
                        _OperandItem = Operands.Values.ToList(nextIndex)
                    End If
                End If

            ElseIf Mouse_Region = MouseRegion.ClearText Then
                RaiseEvent ClearTextClicked(Me, New ImageComboEventArgs)
                If IsReadOnly Then Exit Sub
                Text = String.Empty
                SelectedIndex = -1
                Invalidate()

            ElseIf Mouse_Region = MouseRegion.DropDown Then
                _DropItems = Items
                If Not DropDown.Visible Then
                    DropDown.ResizeMe()
                    Dim Coordinates As Point
                    Coordinates = PointToScreen(New Point(0, 0))
                    Toolstrip.Show(Coordinates.X, If(Coordinates.Y + DropDown.Height > My.Computer.Screen.WorkingArea.Height, Coordinates.Y - DropDown.Height, Coordinates.Y + Height))
                End If
                DropDown.Visible = Not (DropDown.Visible)
                If DropDown.Visible Then
                    RaiseEvent DropDownOpened(Me, e)
                Else
                    RaiseEvent DropDownClosed(Me, e)
                End If

            Else
                REM Add
                Mouse_Region = MouseRegion.Text
                CursorShouldBeVisible = True
                CursorBlinkTimer.Start()
                CursorIndex = GetLetterIndex(e.X)
                SelectionIndex = CursorIndex

            End If
            Invalidate()
        End If
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            If EyeBounds.Width > 0 And e.X > EyeBounds.Left And e.X < EyeBounds.Right Then
                Mouse_Region = MouseRegion.Eye

            ElseIf ClearTextBounds.Width > 0 And e.X > ClearTextBounds.Left And e.X < ClearTextBounds.Right Then
                Mouse_Region = MouseRegion.ClearText

            ElseIf Items.Any And e.X >= DropBounds.Left Then
                Mouse_Region = MouseRegion.DropDown

            ElseIf e.X < ImageBounds.Width Then
                Mouse_Region = MouseRegion.Image

            Else
                Mouse_Region = MouseRegion.None

            End If
            If e.Button = MouseButtons.Left Then
                SelectionIndex = GetLetterIndex(e.X)
            End If
            Invalidate()
        End If
        MyBase.OnMouseMove(e)

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
    Private Function TextLength(ByVal T As String) As Integer

        Dim Padding As Integer = If(T.Length = 0, 0, (2 * TextRenderer.MeasureText(T.First, Font).Width) - TextRenderer.MeasureText(T.First & T.First, Font).Width)
        Return TextBounds.Left + TextRenderer.MeasureText(T, Font).Width - Padding

    End Function
    Private Sub ShowMatches(MatchText As String)

        Dim Matches As New List(Of ComboItem)(From CI In Items Where CI.Text.ToUpperInvariant.StartsWith(MatchText.ToUpperInvariant, StringComparison.InvariantCulture))
        With DropDown
            .Visible = False
            Select Case Matches.Count
                Case 0
                Case Else
                    REM Show Matches
                    _DropItems = Matches
                    .Invalidate()
                    .Visible = True
                    .ResizeMe()
                    Dim Coordinates As Point
                    Coordinates = PointToScreen(New Point(0, 0))
                    Toolstrip.Show(Coordinates.X, If(Coordinates.Y + .Height > My.Computer.Screen.WorkingArea.Height, Coordinates.Y - .Height, Coordinates.Y + Height))
            End Select
            .VScroll.Value = 0
        End With
    End Sub
    Public Sub SelectAll()

        CursorIndex = 0
        _SelectionIndex = LetterWidths.Keys.Last
        Invalidate()

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
    Private Const ShadowDepth As Integer = 8
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
        ImageCombo = Parent

    End Sub
    Protected Overrides Sub InitLayout()

        ResizeMe()
        MyBase.OnFontChanged(Nothing)
        MyBase.InitLayout()

    End Sub
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        If e IsNot Nothing Then
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            e.Graphics.FillRectangle(New SolidBrush(BackColor), ComboItemRectangle)
            Using GP As New GraphicsPath
                Dim Colors() As Color = {BackColor, ShadeColor}
                GP.AddRectangle(ComboItemRectangle)
                Using PathBrush As New PathGradientBrush(GP)
                    PathBrush.SurroundColors = Colors
                    PathBrush.CenterColor = Color.FromArgb(128, Color.WhiteSmoke)
                    e.Graphics.FillPath(PathBrush, GP)
                End Using
            End Using
            For Each ComboItem In VisibleComboItems
                With ComboItem
                    Dim Bounds As Rectangle = .Bounds
                    Bounds.Offset(0, -VScroll.Value)
                    e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(64, BackColor)), .Bounds)
                    If ImageCombo.CheckBoxes Then
                        Dim CheckBounds As Rectangle = .CheckBounds
                        CheckBounds.Offset(0, -VScroll.Value)
                        ControlPaint.DrawCheckBox(e.Graphics, CheckBounds, If(.Checked, ButtonState.Checked, ButtonState.Normal))
                    End If
                    If Not IsNothing(.Image) Then
                        Dim ImageBounds As Rectangle = .ImageBounds
                        ImageBounds.Offset(0, -VScroll.Value)
                        e.Graphics.DrawImage(.Image, ImageBounds)
                    End If
                    If .Selected Then
                        Using Brush As New LinearGradientBrush(Bounds, Color.FromArgb(20, SelectionColor), Color.FromArgb(60, SelectionColor), linearGradientMode:=LinearGradientMode.Vertical)
                            e.Graphics.FillRectangle(Brush, Brush.Rectangle)
                        End Using
                        Using Pen As New Pen(SelectionColor)
                            e.Graphics.DrawRectangle(Pen, Bounds)
                        End Using
                    End If
                    Dim TextBounds As Rectangle = .TextBounds
                    TextBounds.Offset(0, -VScroll.Value)

                    TextRenderer.DrawText(e.Graphics, Replace(.Text, "&", "&&"), Font, TextBounds, Color.Black, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                    If ComboItem.Separator Then e.Graphics.DrawLine(Pens.Black, New Point(0, TextBounds.Bottom), New Point(TextBounds.Right, TextBounds.Bottom))

                    If ComboItem.Index = MouseRowIndex Then
                        Using Brush As New LinearGradientBrush(Bounds, Color.FromArgb(20, Color.DarkSlateGray), Color.FromArgb(60, Color.DarkSlateGray), linearGradientMode:=LinearGradientMode.Vertical)
                            e.Graphics.FillRectangle(Brush, Brush.Rectangle)
                        End Using
                        Using Pen As New Pen(Color.DarkSlateGray)
                            e.Graphics.DrawRectangle(Pen, Bounds)
                        End Using
                    End If

                End With
            Next ComboItem

            Using BMP_Right As New Bitmap(ShadowRight.Width, ShadowRight.Height)
                e.Graphics.DrawImage(BMP_Shadow, ShadowRight.Left, 0, ShadowRight, GraphicsUnit.Pixel)
            End Using
            Using BMP_Bottom As New Bitmap(ShadowBottom.Width, ShadowBottom.Height)
                e.Graphics.DrawImage(BMP_Shadow, 0, ShadowBottom.Top, ShadowBottom, GraphicsUnit.Pixel)
            End Using
            If VScroll.Visible Then
                e.Graphics.FillRectangle(Brushes.GhostWhite, VScroll.Bounds)
                ControlPaint.DrawBorder3D(e.Graphics, VScroll.Bounds, Border3DStyle.RaisedInner)
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                If VScroll.Lines Then
                    Using Pen As New Pen(VScroll.Color, 1)
                        For Each Page As Integer In VScroll.Pages
                            e.Graphics.DrawLine(Pen, VScroll.TrackBounds.Left, Page, VScroll.TrackBounds.Right - 2, Page)
                        Next
                    End Using
                End If
                Using Pen As New Pen(Brushes.Black, 1)
                    e.Graphics.DrawLine(Pen, VScroll.UpBounds.Left, VScroll.UpBounds.Bottom, VScroll.UpBounds.Right - 2, VScroll.UpBounds.Bottom)
                    e.Graphics.DrawLine(Pen, VScroll.DownBounds.Left, VScroll.DownBounds.Top, VScroll.DownBounds.Right - 2, VScroll.DownBounds.Top)
                End Using
                Using Brush As New SolidBrush(Color.FromArgb(VScroll.UpAlpha, VScroll.Color))
                    e.Graphics.FillRectangle(Brush, VScroll.UpBounds)
                End Using
                Using Brush As New SolidBrush(Color.FromArgb(VScroll.DownAlpha, VScroll.Color))
                    e.Graphics.FillRectangle(Brush, VScroll.DownBounds)
                End Using
                Dim ArrowWidth As Integer = 8, ArrowHeight As Integer = 4, ArrowCenter As Integer = Convert.ToInt32((VScroll.UpBounds.Height - ArrowHeight) / 2)
                Dim TriangeTop As Integer = VScroll.UpBounds.Top + ArrowCenter
                Dim TriangleLeft As Integer = VScroll.Bounds.Left + 1, TRight As Integer = TriangleLeft + ArrowWidth, TMid As Integer = TriangleLeft + Convert.ToInt32(ArrowWidth / 2)
                Dim Triangle As Point() = {New Point(TMid, TriangeTop), New Point(TRight, TriangeTop + ArrowHeight), New Point(TriangleLeft, TriangeTop + ArrowHeight)}
                Using Brush As New SolidBrush(Color.FromArgb(255, VScroll.Color))
                    e.Graphics.FillPolygon(Brush, Triangle)
                End Using
                Triangle = {New Point(TriangleLeft, VScroll.DownBounds.Top + ArrowCenter), New Point(TRight, VScroll.DownBounds.Top + ArrowCenter), New Point(TMid, VScroll.DownBounds.Top + ArrowCenter + ArrowHeight)}
                Using Brush As New SolidBrush(Color.FromArgb(255, VScroll.Color))
                    e.Graphics.FillPolygon(Brush, Triangle)
                End Using
                Using Brush As New SolidBrush(Color.FromArgb(VScroll.Alpha, VScroll.Color))
                    e.Graphics.FillRectangle(Brush, VScroll.BarBounds)
                End Using
            End If
        End If

    End Sub
#Region " Properties & Fields "
    Private ReadOnly Property ImageCombo As ImageCombo
    Private ReadOnly Property MatchedItems As List(Of ComboItem)
        Get
            Return ImageCombo.DropItems
        End Get
    End Property
    Private ReadOnly Property VisibleComboItems As List(Of ComboItem)
        Get
            Return MatchedItems
        End Get
    End Property
    Private _ComboItemRectangle As New Rectangle
    Private ReadOnly Property ComboItemRectangle As Rectangle
        Get
            _ComboItemRectangle.Width = Width - ShadowDepth
            _ComboItemRectangle.Height = Height - ShadowDepth
            Return _ComboItemRectangle
        End Get
    End Property
    Private _TotalHeight As Integer
    Private ReadOnly Property TotalHeight As Integer
        Get
            Return _TotalHeight
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
    Private _ShadowRight As New Rectangle
    Private ReadOnly Property ShadowRight As Rectangle
        Get
            _ShadowRight.X = ComboItemRectangle.Right
            _ShadowRight.Y = 0
            _ShadowRight.Width = ShadowDepth
            _ShadowRight.Height = ComboItemRectangle.Height
            Return _ShadowRight
        End Get
    End Property
    Private _ShadowBottom As New Rectangle
    Private ReadOnly Property ShadowBottom As Rectangle
        Get
            _ShadowBottom.X = 0
            _ShadowBottom.Y = ComboItemRectangle.Bottom
            _ShadowBottom.Width = Width
            _ShadowBottom.Height = ShadowDepth
            Return _ShadowBottom
        End Get
    End Property
    Public Property ShadeColor As Color = Color.WhiteSmoke
    Public Property SelectionColor As Color = Color.Transparent
    Public Property DropShadowColor As Color = Color.CornflowerBlue
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
#End Region
    Protected Shadows Sub OnPreviewKeyDown(ByVal sender As Object, ByVal e As PreviewKeyDownEventArgs)

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
                        If VisibleComboItems.IndexOf(ImageCombo.Items(MouseRowIndex)) = ImageCombo.MaxItems - 1 Then
                            VScroll.Value += ItemHeight
                        End If
                        _MouseRowIndex += 1
                    End If

                Case Keys.Return
                    ImageCombo.OnItemSelected(VisibleComboItems(MouseRowIndex), False)

            End Select
            Invalidate()
        End If
        MyBase.OnKeyDown(e)

    End Sub
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            If VScroll.Bounds.Contains(e.Location) Then
            ElseIf Bounds.Contains(e.Location) And Not VScroll.Bounds.Contains(e.Location) Then
                With ImageCombo
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
                            If Not ImageCombo.MultiSelect Then
                                'Dim LastSelected As New List(Of ComboItem)(From CI In Items Where CI.Selected)
                                'If LastSelected.Any Then LastSelected.First._Selected = False
                            End If
                            Selected.First._Selected = Not (Selected.First.Selected)
                            If ImageCombo.CheckOnSelect Then Selected.First.Checked = Not (Selected.First.Checked)
                            .OnItemSelected(Selected.First, Control.ModifierKeys = Keys.Shift)
                        End If
                    End If
                End With
            ElseIf ImageCombo.Bounds.Contains(e.Location) Then
                ImageCombo.Focus()
            Else
                Visible = False
            End If
        End If
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            If VScroll.Bounds.Contains(e.Location) Or VScroll.Scrolling Then
            ElseIf Bounds.Contains(e.Location) Then
                Dim Location As New Point(e.Location.X, (e.Location.Y + VScroll.Value))
                Dim MouseOverComboItems As New List(Of ComboItem)(From I In MatchedItems Where I.Bounds.Contains(Location) Select I)
                If MouseOverComboItems.Any Then
                    _MouseOverCombo = MouseOverComboItems.First
                    _MouseRowIndex = _MouseOverCombo.Index
                    If Not (If(MouseOverCombo.TipText, String.Empty).Length = 0 Or _LastMouseOverCombo Is _MouseOverCombo) Then
                        ToolTip.Hide(ImageCombo)
                        ToolTip.Show(MouseOverCombo.TipText, ImageCombo, 0, 0, 2000)
                        _LastMouseOverCombo = _MouseOverCombo
                    End If
                    ForceCapture = True
                    Invalidate()
                End If
            ElseIf ImageCombo.Bounds.Contains(e.Location) Then
            End If
        End If
        MyBase.OnMouseMove(e)

    End Sub
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)

        If Visible Then
            ResizeMe()
            ImageCombo.Toolstrip.Size = Size
            Top = -1
            Top = 0

            'NativeMethods.SetProcessDPIAware

            Dim BitMap As New Bitmap(Width, Height)
            Using Graphics As Graphics = Graphics.FromImage(BitMap)
                'Graphics.PageScale = Math.Min(fSize.Width / Graphics.DpiX / 1000, fSize.Height / Graphics.DpiY / 1000)
                Dim Point As Point = PointToScreen(New Point(0, 0))
                Graphics.CopyFromScreen(Point.X, Point.Y, 0, 0, BitMap.Size, CopyPixelOperation.SourceCopy)
                For P = 0 To ShadowDepth - 1
                    Using sb As New SolidBrush(Color.FromArgb(16 + (P * 5), DropShadowColor))
                        Graphics.FillRectangle(sb, New Rectangle(ShadowDepth + P, ShadowDepth + P, Width - ShadowDepth - P * 2, Height - ShadowDepth - P * 2))
                    End Using
                Next
            End Using
            BMP_Shadow = BitMap
            Dim SV As IEnumerable(Of ComboItem) = From S In MatchedItems Where S.Index = ImageCombo.SelectionIndex
            If SV.Any Then
                Dim ScrollValue As Integer = MatchedItems.IndexOf(SV.First)
                VScroll.Value = Convert.ToInt32(Split((ScrollValue / ImageCombo.MaxItems).ToString(InvariantCulture), ".")(0), InvariantCulture) * ImageCombo.MaxItems * ItemHeight
                Invalidate()
            End If
            ForceCapture = True
        Else
            ImageCombo.Toolstrip.Size = New Size(0, 0)
            ForceCapture = False
            ImageCombo.Focus()
        End If
        MyBase.OnVisibleChanged(e)

    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)

        If e IsNot Nothing Then
            If MatchedItems.Any And Asc(e.KeyChar) > 31 AndAlso Asc(e.KeyChar) < 127 Then
                REM Printable characters
                ImageCombo.DelegateKeyPress(e)
            End If
        End If
        MyBase.OnKeyPress(e)

    End Sub
    Protected Overrides Sub OnFontChanged(e As EventArgs)

        ResizeMe()
        MyBase.OnFontChanged(e)

    End Sub
    Friend Sub ResizeMe()

        _ItemHeight = (1 + TextRenderer.MeasureText("ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ".ToString(InvariantCulture), ImageCombo.Font).Height + 1)
        With ImageCombo
            If MatchedItems.Any Then
                REM //////////////////////////// UPDATE COMBOITEM BOUNDS
                For Each ComboItem As ComboItem In VisibleComboItems
                    With ComboItem
                        ._Index = VisibleComboItems.IndexOf(ComboItem)
                        ._Bounds.X = 0
                        ._Bounds.Y = (ItemHeight * .Index)
                        ._Bounds.Width = Width - ShadowDepth
                        ._Bounds.Height = ItemHeight
                        If ImageCombo.CheckBoxes Then
                            ._CheckBounds.X = 2
                            ._CheckBounds.Y = ._Bounds.Y + Convert.ToInt32((ItemHeight - 14) / 2)
                            ._CheckBounds.Width = 14
                            ._CheckBounds.Height = 14
                            ._CheckBounds.Offset(0, 1)
                        Else
                            ._CheckBounds.X = 0
                            ._CheckBounds.Y = 0
                            ._CheckBounds.Width = 0
                            ._CheckBounds.Height = 0
                        End If
                        If IsNothing(.Image) Then
                            ._ImageBounds.X = ._CheckBounds.Right
                            ._ImageBounds.Y = ._Bounds.Y
                            ._ImageBounds.Width = 0
                            ._ImageBounds.Height = 0
                        Else
                            Dim ImageWidth As Integer = {ItemHeight, .Image.Height}.Min
                            ._ImageBounds.X = ._CheckBounds.Right + 2
                            ._ImageBounds.Y = ._Bounds.Y + Convert.ToInt32((ItemHeight - ImageWidth) / 2)
                            ._ImageBounds.Width = ImageWidth
                            ._ImageBounds.Height = ImageWidth
                        End If
                        ._TextBounds.X = ._ImageBounds.Right + If(Not IsNothing(.Image), 2, 0)
                        ._TextBounds.Y = ._Bounds.Y
                        ._TextBounds.Width = ._Bounds.Width - ._TextBounds.X
                        ._TextBounds.Height = ItemHeight
                    End With
                Next

                Width = (From R In VisibleComboItems Select 3 + If(IsNothing(R.Image), 0, 1 + R.Image.Width) + R.CheckBounds.Width + TextRenderer.MeasureText(R.Text, Font).Width).Union({ .TextBounds.Left, .Width}).Max + ShadowDepth + If(VScroll.Visible, VScroll.Bounds.Width, 0)
                _TotalHeight = VisibleComboItems.Count * ItemHeight
                Height = { .MaxItems * ItemHeight, TotalHeight}.Min + ShadowDepth
                VScroll.Height = ComboItemRectangle.Height
                VScroll.Maximum = TotalHeight
                VScroll.SmallChange = ItemHeight
                VScroll.LargeChange = ComboItemRectangle.Height
                .Toolstrip.Size = Size
            Else
                _TotalHeight = 0
                .Toolstrip.Size = New Size(0, 0)
            End If
        End With
        Invalidate()
    End Sub
End Class
REM ////////////////////////////////////////////////////////////////////////////////////////////////////////// DROPDOWN COLLECTION
Public NotInheritable Class ItemCollection
    Inherits List(Of ComboItem)
    Public Sub New(Parent As ImageCombo)
        ImageCombo = Parent
    End Sub
    Public ReadOnly Property ImageCombo As ImageCombo
    Public Shadows Function Item(TheName As String) As ComboItem

        Dim Items As New List(Of ComboItem)(From CI In Me Where CI.Name = TheName)
        If Items.Any Then
            Return Items.First
        Else
            Return Nothing
        End If

    End Function
    Public Overloads Function Add(ByVal Text As String) As ComboItem

        Dim ComboItem As New ComboItem With {.Value = Text}
        Add(ComboItem)
        Return ComboItem

    End Function
    Public Overloads Function Add(ByVal Text As String, Image As Image) As ComboItem

        Dim ComboItem As New ComboItem With {.Value = Text, .Image = Image}
        Add(ComboItem)
        Return ComboItem

    End Function
    Public Overloads Function Add(ByVal ComboItem As ComboItem) As ComboItem

        If ComboItem IsNot Nothing Then
            With ComboItem
                ._ItemCollection = Me
                Dim ItemHeight As Integer = TextRenderer.MeasureText("ZZZZZZZZZZZZZZZZZZZZ".ToString(InvariantCulture), ImageCombo.Font).Height
                ._Index = Count
                ._Bounds.X = 0
                ._Bounds.Y = (ItemHeight * .Index)
                ._Bounds.Width = ImageCombo.DropDown.Width - 3
                ._Bounds.Height = ItemHeight
                If ImageCombo.CheckBoxes Then
                    ._CheckBounds.X = 2
                    ._CheckBounds.Y = ._Bounds.Y + Convert.ToInt32((ItemHeight - 14) / 2)
                    ._CheckBounds.Width = 14
                    ._CheckBounds.Height = 14
                Else
                    ._CheckBounds.X = 0
                    ._CheckBounds.Y = 0
                    ._CheckBounds.Width = 0
                    ._CheckBounds.Height = 0
                End If
                If IsNothing(.Image) Then
                    ._ImageBounds.X = ._CheckBounds.Right
                    ._ImageBounds.Y = ._Bounds.Y
                    ._ImageBounds.Width = 0
                    ._ImageBounds.Height = 0
                Else
                    Dim ImageWidth As Integer = {ItemHeight, .Image.Height}.Min
                    ._ImageBounds.X = ._CheckBounds.Right + 2
                    ._ImageBounds.Y = ._Bounds.Y + Convert.ToInt32((ItemHeight - ImageWidth) / 2)
                    ._ImageBounds.Width = ImageWidth
                    ._ImageBounds.Height = ImageWidth
                End If
                ._TextBounds.X = ._ImageBounds.Right + If(Not IsNothing(.Image), 2, 0)
                ._TextBounds.Y = ._Bounds.Y
                ._TextBounds.Width = ._Bounds.Width - ._TextBounds.X
                ._TextBounds.Height = ItemHeight
            End With
            MyBase.Add(ComboItem)
            ImageCombo.Invalidate()
        End If
        Return ComboItem

    End Function
End Class
REM ////////////////////////////////////////////////////////////////////////////////////////////////////////// DROPDOWN COMBO ITEM
<Serializable> <TypeConverter(GetType(PropertyConverter))> Public Class ComboItem
    Public Sub New()
    End Sub
#Region " Properties & Fields "
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
#End Region
End Class