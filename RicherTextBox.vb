Option Explicit On
Option Strict On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Public NotInheritable Class RicherEventArgs
    Inherits EventArgs
    Public ReadOnly Property VScrollValue As Integer
    Public Sub New(VScrollValue As Integer)
        Me.VScrollValue = VScrollValue
    End Sub
End Class
''' <summary>
''' Provides the start, length, bounds and value of the cursors location
''' </summary>
''' 
<StructLayout(LayoutKind.Sequential)>
Public Structure MouseData
    Implements IEquatable(Of MouseData)
    Public Property WordRectangle As Rectangle
    Public Property WordStart As Integer
    Public Property WordLength As Integer
    Public ReadOnly Property WordEnd As Integer
        Get
            Return WordStart + WordLength
        End Get
    End Property
    Public Property Word As String
    Public Property Intersects As Boolean
    Public Overrides Function ToString() As String
        Return Join({Word, WordStart, WordLength, WordRectangle.ToString, Intersects.ToString(InvariantCulture)}, "*")
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return WordRectangle.GetHashCode Xor WordStart.GetHashCode Xor WordLength.GetHashCode Xor Word.GetHashCode Xor Intersects.GetHashCode
    End Function
    Public Overloads Function Equals(other As MouseData) As Boolean Implements IEquatable(Of MouseData).Equals
        Return WordRectangle = other.WordRectangle AndAlso WordStart = other.WordStart AndAlso WordLength = other.WordLength AndAlso Word = other.Word
    End Function
    Public Shared Operator =(value1 As MouseData, value2 As MouseData) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(value1 As MouseData, value2 As MouseData) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(obj As Object) As Boolean
        If TypeOf obj Is MouseData Then
            Return CType(obj, MouseData) = Me
        Else
            Return False
        End If
    End Function
End Structure
Public NotInheritable Class RicherTextBox
    Inherits RichTextBox
    Public Sub New()

        'SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        'SetStyle(ControlStyles.ContainerControl, True)
        'SetStyle(ControlStyles.DoubleBuffer, True)
        'SetStyle(ControlStyles.UserPaint, True)
        'SetStyle(ControlStyles.ResizeRedraw, True)
        'SetStyle(ControlStyles.Selectable, True)
        'SetStyle(ControlStyles.Opaque, True)
        'SetStyle(ControlStyles.UserMouse, True)
        InnerSizeWidth = Size.Width - ClientSize.Width
        InnerSizeHeight = Size.Height - ClientSize.Height
        FindReplace_add()

    End Sub
    'Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
    '    Get
    '        Dim cp As CreateParams = MyBase.CreateParams
    '        cp.ExStyle = cp.ExStyle Or 33554432
    '        Return cp
    '    End Get
    'End Property

    Public Event ScrolledVertical(sender As Object, e As RicherEventArgs)
    Public Event DragStart(sender As Object, e As DragEventArgs)
    Public Event PageChanged(sender As Object, e As Integer)

    Private ReadOnly Dragging As New Dictionary(Of Point, Boolean)
    Private Const SIF_RANGE As Integer = &H1
    Private Const SIF_PAGE As Integer = &H2
    Private Const SIF_POS As Integer = &H4
    Private Const SB_HORZ As Integer = &H0
    Private Const SB_VERT As Integer = &H1
    Private Const WM_VSCROLL = &H115
    Private Const WM_HSCROLL = &H114
    Private Const SB_THUMBPOSITION = 4

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            If IsNothing(Text) OrElse Text.Length = 0 Then
                With _MouseWord
                    .Word = String.Empty
                    .WordStart = 0
                    .WordLength = 0
                    .WordRectangle = New Rectangle(0, 0, 0, 0)
                    .Intersects = False
                End With
            Else
                Dim CharacterIndex As Integer = GetCharIndexFromPosition(e.Location)
                Dim LetterIndex As Integer = 0
                Dim Word As String = Nothing
                REM /// GO BACKWARDS
                Do While (CharacterIndex + LetterIndex) > 0
                    Word = Text.Substring(CharacterIndex + LetterIndex, Math.Abs(LetterIndex))
                    If Regex.Match(Word, "[\s]", RegexOptions.None).Success Then Exit Do
                    LetterIndex -= 1
                Loop
                Dim BackWord As Integer = LetterIndex
                REM /// GO FORWARDS
                LetterIndex = 1
                Do While (CharacterIndex + LetterIndex) < TextLength
                    Word = Text.Substring(CharacterIndex, LetterIndex)
                    If Regex.Match(Word, "[\s]", RegexOptions.None).Success Then Exit Do
                    LetterIndex += 1
                Loop
                Dim ForeWord As Integer = LetterIndex
                Dim WordLength As Integer = ForeWord + Math.Abs(BackWord)
                Dim WordStart As Integer = CharacterIndex + BackWord
                Word = Text.Substring(WordStart, WordLength)
                '-------------------------------------------------------------------------------------
                Dim TextPoint As Point = GetPositionFromCharIndex(WordStart)
                Dim WordWidth As Integer = TextRenderer.MeasureText(Word, Font).Width
                With _MouseWord
                    .Word = Word
                    .WordStart = WordStart
                    .WordLength = WordLength
                    .WordRectangle = New Rectangle(TextPoint.X, TextPoint.Y, WordWidth, LineHeight)
                    .Intersects = (.WordRectangle.Contains(e.Location))
                End With
                If Not Dragging.ContainsKey(e.Location) Then Dragging.Add(e.Location, e.Button = MouseButtons.Left)
                If SelectionLength > 0 And (From d In Dragging.Skip(Dragging.Count - 10) Where d.Value).Count = 10 Then
                    Dim Data As New DataObject
                    Data.SetData(GetType(String), SelectedText)
                    AllowDrop = True
                    RaiseEvent DragStart(Me, New DragEventArgs(Data, e.Button, e.X, e.Y, DragDropEffects.All, DragDropEffects.All))
                End If
            End If
            MyBase.OnMouseMove(e)
        End If

    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        Dragging.Clear()
        MyBase.OnMouseUp(e)
    End Sub
    Public ReadOnly Property MouseWord As MouseData

#Region " FindReplace "
    Private WithEvents FindReplace_ As FindReplace
    Public ReadOnly Property FindReplace As FindReplace
        Get
            Return FindReplace_
        End Get
    End Property
    Private HasFindReplace_ As Boolean = True
    Public Property HasFindReplace As Boolean
        Get
            Return HasFindReplace_
        End Get
        Set(value As Boolean)
            If value <> HasFindReplace_ Then
                HasFindReplace_ = value
                If value Then
                    FindReplace_remove()
                    FindReplace_add()

                Else
                    FindReplace_remove()

                End If
            End If
        End Set
    End Property
    Private Sub FindReplace_add()

        FindReplace_ = New FindReplace With {.Margin = New Padding(0), .Visible = False}
        Controls.Add(FindReplace)
        AddHandler KeyDown, AddressOf KeyDown_FindReplace
        AddHandler FindReplace.FindChanged, AddressOf FindRequest
        AddHandler FindReplace.ZoneClicked, AddressOf FindReplace_zoneClicked
        FindReplace.DataBindings.Add("DataSource", Me, "Text")

    End Sub
    Private Sub FindReplace_remove()

        If FindReplace IsNot Nothing Then
            Controls.Remove(FindReplace)
            RemoveHandler FindReplace.KeyDown, AddressOf KeyDown_FindReplace
            RemoveHandler FindReplace.FindChanged, AddressOf FindRequest
            RemoveHandler FindReplace.ZoneClicked, AddressOf FindReplace_zoneClicked
            FindReplace_.Dispose()
        End If

    End Sub
    Private Sub KeyDown_FindReplace(sender As Object, e As KeyEventArgs)

        If ModifierKeys = Keys.Control Then
            If e.KeyCode = Keys.F Then
                FindReplace.Show()
                If SelectedText.Any Then FindReplace.FindControl.Text = SelectedText
            End If
        End If

    End Sub
    Private Sub FindReplace_zoneClicked(ByVal sender As Object, ByVal e As ZoneEventArgs)

        Dim textSearch As String = Text
        If e.Zone.Name = Zone.Identifier.MatchCase Or e.Zone.Name = Zone.Identifier.MatchWord Or e.Zone.Name = Zone.Identifier.RegEx Then
            FindRequest(Nothing, New FindEventArgs(Nothing))

        ElseIf e.Zone.Name = Zone.Identifier.Close Then
            Dim selectStart As Integer = SelectionStart
            SelectAll()
            SelectionBackColor = Color.Transparent
            SelectionColor = Color.Black
            SelectionStart = selectStart
            SelectionLength = 0

        ElseIf e.Zone.Name = Zone.Identifier.GotoNext Then
            If FindReplace.CurrentMatch.Key >= 0 Then
                Dim frPoint As Point = FindReplace.PointToClient(PointToScreen(New Point()))
                FindRequest(Nothing, New FindEventArgs(Nothing))
                Dim Match = FindReplace.CurrentMatch
                Dim rtfTemp = Rtf
                Using RTB As New RichTextBox With
                    {
                .Rtf = rtfTemp,
                .SelectionStart = Match.Key,
                .SelectionLength = Match.Value.Length,
                .SelectionBackColor = Color.DarkBlue,
                .SelectionColor = Color.White
                    }
                    rtfTemp = RTB.Rtf
                End Using
                Rtf = rtfTemp
                SelectionStart = Match.Key
                Dim currentPoint As Point = GetPositionFromCharIndex(SelectionStart)
                If Not ClientRectangle.Contains(currentPoint) Then ScrollToCaret()
                FindReplace.Location = New Point(Math.Abs(frPoint.X), Math.Abs(frPoint.Y))
            End If

        ElseIf e.Zone.Name = Zone.Identifier.ReplaceOne Then
            If FindReplace.CurrentMatch.Key >= 0 Then
                textSearch = textSearch.Remove(FindReplace.CurrentMatch.Key, FindReplace.CurrentMatch.Value.Length)
                textSearch = textSearch.Insert(FindReplace.CurrentMatch.Key, FindReplace.ReplaceControl.Text)
                Text = textSearch
                'FindReplace.DataSource = textSearch
                FindRequest(Nothing, New FindEventArgs(Nothing))
            End If

        ElseIf e.Zone.Name = Zone.Identifier.ReplaceAll Then
            If FindReplace.CurrentMatch.Key >= 0 Then
                Dim ReverseOrderMatches = FindReplace.Matches.OrderByDescending(Function(x) x.Key)
                For Each match In ReverseOrderMatches
                    textSearch = textSearch.Remove(match.Key, match.Value.Length)
                    textSearch = textSearch.Insert(match.Key, FindReplace.ReplaceControl.Text)
                Next
                If FindReplace.CurrentMatch.Key >= 0 Then
                    Text = textSearch
                    'FindReplace.DataSource = textSearch
                    FindRequest(Nothing, New FindEventArgs(Nothing))
                End If
            End If

        End If
    End Sub
    Private Sub FindRequest(ByVal sender As Object, ByVal e As FindEventArgs)

        Dim selectionStartTemp As Integer = SelectionStart
        SelectAll()
        SelectionBackColor = Color.Transparent
        SelectionColor = Color.Black
        SelectionStart = selectionStartTemp
        SelectionLength = 0

        If (If(FindReplace.FindControl.Text, String.Empty)).Any() Then
            Dim rtfTemp As String = Rtf
            Using RTB As New RichTextBox With {.Rtf = rtfTemp}
                For Each match In FindReplace.Matches
                    SelectionStart = match.Key
                    SelectionLength = match.Value.Length
                    SelectionBackColor = Color.Yellow
                    SelectionColor = Color.Black
                Next
                rtfTemp = Rtf
            End Using
            Rtf = rtfTemp
            SelectionStart = selectionStartTemp
        End If

    End Sub
#End Region

    Public Property VScrollPos() As Integer
        Get
            Return NativeMethods.GetScrollPos(Handle, SB_VERT)
        End Get
        Set(value As Integer)
            Dim result = NativeMethods.SetScrollPos(Handle, SB_VERT, value, True)
            NativeMethods.PostMessageA(Handle, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * value, Nothing)
        End Set
    End Property
    ''' <summary>
    ''' Gets and Sets the Horizontal Scroll position of the control.
    ''' </summary>
    Public Property HScrollPos() As Integer
        Get
            Return NativeMethods.GetScrollPos(Handle, SB_HORZ)
        End Get
        Set(value As Integer)
            Dim result = NativeMethods.SetScrollPos(Handle, SB_HORZ, value, True)
            NativeMethods.PostMessageA(Handle, WM_HSCROLL, SB_THUMBPOSITION + &H10000 * value, Nothing)
        End Set
    End Property
    Public ReadOnly Property ScrollData As SCROLLINFO
        Get
            Dim _ScrollInfo As New SCROLLINFO With {
                .CbSize = Marshal.SizeOf(GetType(SCROLLINFO)),
                .FMask = SIF_RANGE Or SIF_PAGE Or SIF_POS
            }
            NativeMethods.GetScrollInfo(Handle, SB_VERT, _ScrollInfo)
            Return _ScrollInfo
        End Get
    End Property
    Private ReadOnly InnerSizeWidth As Integer
    Public ReadOnly Property LineHeight As Integer
        Get
            Return TextRenderer.MeasureText("XXXXXXXXXXXXXXXXXXXXXXXX".ToString(InvariantCulture), Font).Height
        End Get
    End Property
    Public ReadOnly Property MaxCharacterWidth As Integer
        Get
            Return Convert.ToInt32(TextRenderer.MeasureText(StrDup(100, "X"), Font).Width / 100)
        End Get
    End Property
    Public ReadOnly Property VScrollVisible As Boolean
        Get
            Return VScrollWidth > 0
        End Get
    End Property
    Public ReadOnly Property VScrollWidth As Integer
        Get
            Return Size.Width - ClientSize.Width - InnerSizeWidth
        End Get
    End Property
    Public ReadOnly Property VScrollBounds As Rectangle
        Get
            Dim Location = PointToScreen(New Point(0, 0))
            Return New Rectangle(Location.X + ClientSize.Width, Location.Y, VScrollWidth, Height)
        End Get
    End Property
    Public ReadOnly Property HScrollWidth As Integer
        Get
            Return Size.Width - ClientSize.Width - InnerSizeWidth
        End Get
    End Property
    Private ReadOnly InnerSizeHeight As Integer
    Public ReadOnly Property HScrollVisible As Boolean
        Get
            Return (Size.Height - ClientSize.Height) > InnerSizeHeight
        End Get
    End Property
    Public ReadOnly Property HasText As Boolean
        Get
            If IsNothing(Text) Then
                Return False
            Else
                Return Text.Length > 0
            End If
        End Get
    End Property
    Public ReadOnly Property IsWrapped As Boolean
        Get
            If HasText Then
                Dim LineIsWrapped As Boolean = False
                For Line = 0 To Lines.Count - 1
                    Dim LineStartPosition = GetPositionFromCharIndex(GetFirstCharIndexFromLine(Line))
                    Dim LineEndPosition = GetPositionFromCharIndex(Lines(Line).Length + GetFirstCharIndexFromLine(Line))
                    LineIsWrapped = LineStartPosition.Y <> LineEndPosition.Y
                    If LineIsWrapped Then Exit For
                Next
                Return LineIsWrapped
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property IdealWidth(Optional ReturnZero As Boolean = False) As Integer
        Get
            If Lines.Any Then
#Region " DETERMINE WIDTH "
                'Dim RowWidths = From L In Lines Select TextRenderer.MeasureText(L, Font).Width
                Dim MaxWidth As Integer = 0                         'RowWidths.Max
                Dim Lefts As New List(Of Integer)({0})
                Dim Rights As New List(Of Integer)({WorkingArea.Width})
                Dim Attempts As New List(Of Integer)
                Using InvisibleTextBox As New RicherTextBox With {.Font = Font, .Text = Text, .Width = 0}
                    With InvisibleTextBox
                        Do
                            Dim Delta As Integer = Rights.Min - Lefts.Max
                            Dim Mid As Integer = Lefts.Max + Convert.ToInt32(Delta / 2)
                            Attempts.Add(Mid)
                            If Rights.Min - Mid <= 1 Then
                                'Found it!
                                Exit Do
                            Else
                                .Width = Mid
                                If .IsWrapped Then
                                    Lefts.Add(Mid)
                                Else
                                    Rights.Add(Mid)
                                End If
                            End If
                            If Attempts.Count > Rights.Max Then Exit Do

                        Loop
                        MaxWidth = Attempts.Last
                    End With
                End Using
                If MaximumSize.Width > 0 Then
                    MaxWidth = {MaximumSize.Width, MaxWidth}.Min
                End If
                If MinimumSize.Width > 0 Then
                    MaxWidth = {MinimumSize.Width, MaxWidth}.Max
                End If
#End Region
                Return MaxWidth
            Else
                If ReturnZero Then
                    Return 0
                Else
                    Return Width
                End If
            End If
        End Get
    End Property
    Public ReadOnly Property IdealHeight(Optional ReturnZero As Boolean = False) As Integer
        Get
            Using InvisibleTextBox As New RicherTextBox With {.Font = Font, .Text = Text, .Width = IdealWidth, .Height = 10000, .Margin = New Padding(0)}
                With InvisibleTextBox
                    Dim R = .Rows
                    Return .Rows.Count * .LineHeight
                End With
            End Using
        End Get
    End Property
    Public ReadOnly Property LinesPerPage As Integer
        Get
            Return CInt(Height / LineHeight)
        End Get
    End Property
    Public ReadOnly Property PageNumber As Integer = 1
    Public ReadOnly Property Rows As Dictionary(Of Integer, String)
        Get
            Dim Ys As New Dictionary(Of Integer, List(Of String))
            For l = 0 To Text.Length - 1
                Dim Y = GetPositionFromCharIndex(l).Y
                If Not Ys.ContainsKey(Y) Then Ys.Add(Y, New List(Of String))
                Try
                    Ys(Y).Add(Text.Substring(l, 1))
                Catch ex As ArgumentOutOfRangeException
                End Try
            Next
            Return Ys.ToDictionary(Function(k) k.Key, Function(v) Join(v.Value.ToArray, String.Empty))
        End Get
    End Property
    Private _AutoSize As Boolean
    Public Overrides Property AutoSize As Boolean
        Get
            Return _AutoSize
        End Get
        Set(value As Boolean)
            If _AutoSize <> value Then _AutoSize = value
            If AutoSize Then Size = New Size(IdealWidth, IdealHeight)
        End Set
    End Property
    Public ReadOnly Property VerticalScrollLocation As Point
    Private Sub VerticalScrolled() Handles Me.VScroll

        Dim topLeftCharacter As Integer = GetCharIndexFromPosition(New Point(0, 0))
        Dim topLine As Integer = GetLineFromCharIndex(topLeftCharacter)
        Dim currentPageNbr As Integer = 1 + CInt(Math.Floor(topLine / LinesPerPage))
        If currentPageNbr <> PageNumber Then
            'DrawPageNumber(currentPageNbr)
            _PageNumber = currentPageNbr
            RaiseEvent PageChanged(Me, PageNumber)
        End If
        If VScrollBounds.Contains(Cursor.Position) Then
            _VerticalScrollLocation = Cursor.Position
            RaiseEvent ScrolledVertical(Me, New RicherEventArgs(VScrollPos))
        End If

    End Sub
End Class