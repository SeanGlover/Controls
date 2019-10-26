Option Explicit On
Option Strict On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Public NotInheritable Class TabsEventArgs
    Inherits EventArgs
    Public Property InTab As Tab
    Public Property OutTab As Tab
    Public Property InZone As Tabs.Zone
    Public Property OutZone As Tabs.Zone
    Public Property BeforeBounds As Rectangle
    Public Property AfterBounds As Rectangle
    Public Sub New(Tab As Tab, Before As Rectangle, After As Rectangle)
        InTab = Tab
        BeforeBounds = Before
        AfterBounds = After
    End Sub
    Public Sub New(ClickTab As Tab, ClickZone As Tabs.Zone)
        InTab = ClickTab
        InZone = ClickZone
    End Sub
    Public Sub New(EnterTab As Tab, ExitTab As Tab, EnterZone As Tabs.Zone, ExitZone As Tabs.Zone)
        InTab = EnterTab
        OutTab = ExitTab
        InZone = EnterZone
        OutZone = ExitZone
    End Sub
End Class

Public Class Tabs
    Inherits TabControl
    Public Enum Zone
        None
        Add
        Image
        Text
        Close
    End Enum
    Friend Shared ReadOnly Property Myself As Tabs
    Private ReadOnly Data As New DataObject
    Private IsDragging As Boolean
    Private DragXY As Point
    Private DragTab As Tab

    Public Event TabClicked(sender As Object, e As TabsEventArgs)
    Public Event TabWidthChanged(sender As Object, e As TabsEventArgs)
    Public Event TabMouseChange(sender As Object, e As TabsEventArgs)
    Public Event ZoneMouseChange(sender As Object, e As TabsEventArgs)

    Private MouseXY As Point
    Public Property ZoneColor As Color = Color.White
    Public Property AddNewTabColor As Color = Color.DarkGray
    Public Property SelectedTabColor As Color = Color.Yellow
    Public Property RedrawZoneChange As Boolean
    Public ReadOnly Property MouseTab As Tab
    Public ReadOnly Property MouseZone As Zone

    Public Sub New()

        Dock = DockStyle.Fill
        DrawMode = TabDrawMode.OwnerDrawFixed
        SizeMode = TabSizeMode.Normal
        ItemSize = New Size(40, 26)
        Alignment = TabAlignment.Top
        Multiline = True
        AllowDrop = True
        ImageList = New ImageList
        For Each Image In MyImages()
            ImageList.Images.Add(Image.Key, Image.Value)
        Next
        _Myself = Me
        TabPages_ = New TabCollection

    End Sub

    Private Sub Tabs_DrawHeaders(ByVal sender As Object, ByVal e As DrawItemEventArgs) Handles Me.DrawItem

        '//////////////////////////////////////////////////////////////////////////////
        '/////////////// SETTING THE IMAGE AT RUNTIME CAUSES THE IMAGE TO CORRECTLY PLACE ITSELF, ie) DON'T PAD THE TEXT RIGHT
        '//////////////////////////////////////////////////////////////////////////////

        If e.Index >= 0 And e.Index < TabPages.Count Then
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Dim TabItem As Tab = TabPages.Item(e.Index)
            With TabItem
                Dim FatPenThickness As Integer = If(.Selected, 4, 0)
                Dim TabBounds As Rectangle = e.Bounds
                Dim ImageRegionBounds As Rectangle, ImageBounds As Rectangle
                If TabItem Is AddTab Then
                    ImageRegionBounds = New Rectangle(TabBounds.X, TabBounds.Y, .Image.Width, TabBounds.Height)
                    ImageBounds = New Rectangle(TabBounds.X + Convert.ToInt32((TabBounds.Width - .Image.Width) / 2),
                                          TabBounds.Y + Convert.ToInt32((TabBounds.Height - .Image.Height) / 2),
                                          .Image.Width,
                                          .Image.Height)
                    If TabBounds.Contains(MouseXY) Then _MouseZone = Zone.Add
                    Using AddTabBrush As New LinearGradientBrush(TabBounds, AddNewTabColor, Color.WhiteSmoke, LinearGradientMode.BackwardDiagonal)
                        e.Graphics.FillRectangle(AddTabBrush, TabBounds)
                    End Using
                    If TabItem Is MouseTab Then
                        Using AddTabMouseBrush As New SolidBrush(AddNewTabColor)
                            e.Graphics.FillRectangle(AddTabMouseBrush, TabBounds)
                        End Using
                    End If
                    e.Graphics.DrawImage(.Image, ImageBounds)
                Else
                    Using BoundsBrush As New LinearGradientBrush(TabBounds, .HeaderBackColor, Color.WhiteSmoke, LinearGradientMode.BackwardDiagonal)
                        e.Graphics.FillRectangle(BoundsBrush, TabBounds)
                    End Using
                    Dim SpaceSize As Size = TextRenderer.MeasureText(StrDup(100, " "), .Font)
                    Dim SpaceWidth As Double = SpaceSize.Width / 100
                    Dim Image_LeftSpaceCount As Integer = 0 ' Convert.ToInt32(If(Image Is Nothing, 0, Image.Width) / SpaceWidth)
                    Dim Image_RightSpaceCount As Integer = If(.CanClose And TabItem Is MouseTab, 1 + Convert.ToInt32(My.Resources.Close.Width / SpaceWidth), 0)
                    Dim SpacedText As String = Join({StrDup(Image_LeftSpaceCount, " "), .ItemText, StrDup(8 + Image_RightSpaceCount, " ")}, String.Empty)
                    If SpacedText <> .Text Then
                        'Forces redraw
                        .BeforeBounds = TabBounds
                        .Text = SpacedText
                    Else
                        If .BeforeBounds.Width > 0 Then
                            .AfterBounds = e.Bounds
                            RaiseEvent TabWidthChanged(Me, New TabsEventArgs(TabItem, .BeforeBounds, .AfterBounds))
                            .BeforeBounds = Nothing
                        End If
#Region " BOUNDS "
                        If .Image Is Nothing Then
                            ImageRegionBounds = New Rectangle(TabBounds.X, TabBounds.Y, 0, TabBounds.Height)
                            ImageBounds = ImageRegionBounds
                        Else
                            ImageRegionBounds = New Rectangle(TabBounds.X, TabBounds.Y, .Image.Width, TabBounds.Height)
                            ImageBounds = New Rectangle(TabBounds.X + 3,
                                          TabBounds.Y + Convert.ToInt32((TabBounds.Height - FatPenThickness - .Image.Height) / 2),
                                          .Image.Width,
                                          .Image.Height)
                        End If
                        If ImageRegionBounds.Contains(MouseXY) Then _MouseZone = Zone.Image
#Region " TEXT BOUNDS "
                        Dim TextBounds As Rectangle
                        If If(.ItemText, String.Empty).Any Then
                            Dim TextSize As Size = TextRenderer.MeasureText(.ItemText, Font)
                            TextBounds = New Rectangle(ImageBounds.Right, TabBounds.Top, TextSize.Width, TabBounds.Height - FatPenThickness)
                        Else
                            TextBounds = New Rectangle(ImageBounds.Right, TabBounds.Top, 0, TabBounds.Height - FatPenThickness)
                        End If
#End Region
                        If TextBounds.Contains(Location) Then _MouseZone = Zone.Text
#Region " CLOSE BOUNDS "
                        Dim CloseRegionBounds As Rectangle
                        Dim CloseBounds As Rectangle
                        If .CanClose Then
                            CloseRegionBounds = New Rectangle(TextBounds.Right, TabBounds.Y, TabBounds.Right - TextBounds.Right, TabBounds.Height)
                            CloseBounds = New Rectangle(TabBounds.Right - (My.Resources.Close.Width + 4),
                                             TabBounds.Y + Convert.ToInt32((TabBounds.Height - 4 - My.Resources.Close.Height) / 2),        '4=FatPen.Width ( actually height ) at bottom
                                             My.Resources.Close.Width,
                                             My.Resources.Close.Height)
                        Else
                            CloseBounds = New Rectangle(TabBounds.Right, TabBounds.Top, 0, TabBounds.Height)
                        End If
#End Region
                        If CloseRegionBounds.Contains(Location) Then _MouseZone = Zone.Close

                        If .Selected Then
                            ImageBounds.Offset(6, 0)
                            TextBounds.Offset(4, -1)
                            CloseBounds.Offset(0, 0)
                        End If
#End Region
                        If TabItem Is MouseTab Then
                            'Make solid when mouse is over tab
                            Using MouseBrush As New SolidBrush(.HeaderBackColor)
                                e.Graphics.FillRectangle(MouseBrush, TabBounds)
                            End Using
                            'Only draw X when mouse is over tab
                            If .CanClose Then
                                If MouseZone = Zone.Close Then
                                    Using FadedCloseZoneBrush As New SolidBrush(Color.FromArgb(128, ZoneColor))
                                        Dim CircleBounds As Rectangle = CloseBounds
                                        CircleBounds.Offset(-1, 0)
                                        CircleBounds.Inflate(2, 2)
                                        e.Graphics.FillEllipse(FadedCloseZoneBrush, CircleBounds)
                                    End Using
                                End If
                                e.Graphics.DrawIcon(My.Resources.Close, CloseBounds)
                            End If
                            If IsDragging And MouseTab IsNot DragTab Then
                                Using DragZoneBrush As New SolidBrush(ZoneColor)
                                    e.Graphics.FillRectangle(DragZoneBrush, TabBounds)
                                End Using
                            End If
                        End If
                        If .Image IsNot Nothing Then
                            e.Graphics.DrawImage(.Image, ImageBounds)
                            If TabItem Is MouseTab And MouseZone = Zone.Image Then
                                Using FadedBrush As New SolidBrush(Color.FromArgb(64, ZoneColor))
                                    e.Graphics.FillRectangle(FadedBrush, ImageBounds)
                                End Using
                            End If
                        End If
                        TextRenderer.DrawText(e.Graphics, .Text, e.Font, TextBounds, .HeaderForeColor, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                        If TabItem Is MouseTab And MouseZone = Zone.Text Then
                            Using FadedTextZoneBrush As New SolidBrush(Color.FromArgb(64, ZoneColor))
                                e.Graphics.FillRectangle(FadedTextZoneBrush, TextBounds)
                            End Using
                        End If
                    End If
                    If .Selected Then
                        Using FatPen As New Pen(SelectedTabColor, FatPenThickness)
                            Dim PenTop As Integer = TabBounds.Bottom - CInt(FatPen.Width)
                            e.Graphics.DrawLine(FatPen, New Point(TabBounds.X, PenTop), New Point(TabBounds.Right, PenTop))
                        End Using
                    End If
                End If
            End With
        End If

    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property
    Private Sub Tabs_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter

        MouseOver(PointToClient(New Point(Cursor.Position.X, Cursor.Position.Y)))
        Dim tea = New TabsEventArgs(MouseTab, Nothing, MouseZone, Zone.None)
        RaiseEvent TabMouseChange(Me, tea)
        RaiseEvent ZoneMouseChange(Me, tea)
        Invalidate()

    End Sub
    Private Sub Tabs_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave

        'MouseOver(PointToClient(New Point(Cursor.Position.X, Cursor.Position.Y)))
        RaiseEvent TabMouseChange(Me, New TabsEventArgs(Nothing, MouseTab, Zone.None, MouseZone))
        RaiseEvent ZoneMouseChange(Me, New TabsEventArgs(Nothing, MouseTab, Zone.None, MouseZone))
        _MouseTab = Nothing
        _MouseZone = Zone.None
        Invalidate()

    End Sub
    Private Sub Tabs_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        MouseXY = e.Location
        IsDragging = MouseXY <> DragXY And DragXY.X > 0 And DragXY.Y > 0 And e.Button = MouseButtons.Left
        If IsDragging And UserCanReorder And e.Button = MouseButtons.Left And MouseTab IsNot AddTab Then
            DragTab = MouseTab
            Data.SetData(GetType(Tab), MouseTab)
            MyBase.OnDragOver(New DragEventArgs(Data, 0, e.X, e.Y, DragDropEffects.Copy Or DragDropEffects.Move, DragDropEffects.All))
            DoDragDrop(Data, DragDropEffects.Copy Or DragDropEffects.Move)
        End If
        MouseOver(e.Location)

    End Sub
#Region " REORDER (DRAG/DROP) "
    Private Sub Tabs_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        IsDragging = False
        DragTab = Nothing
    End Sub
    Private Sub Tabs_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

        DragXY = e.Location
        If MouseTab IsNot Nothing Then
            If UserCanAdd And MouseTab Is AddTab Then
                RaiseEvent TabClicked(Me, New TabsEventArgs(MouseTab, MouseZone))

            ElseIf MouseTab.CanClose And MouseZone = Zone.Close Then
                RaiseEvent TabClicked(Me, New TabsEventArgs(MouseTab, MouseZone))

            Else
                RaiseEvent TabClicked(Me, New TabsEventArgs(MouseTab, MouseZone))

            End If
        End If

    End Sub
    Private Sub Tab_DragOver(sender As Object, e As DragEventArgs) Handles Me.DragOver
        MouseOver(PointToClient(New Point(e.X, e.Y)))
        IsDragging = True
    End Sub
    Private Sub Tab_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If UserCanReorder Then e.Effect = DragDropEffects.All
    End Sub
    Private Sub Tab_Dropped(sender As Object, e As DragEventArgs) Handles Me.DragDrop

        If UserCanReorder Then
            MouseOver(PointToClient(New Point(e.X, e.Y)))
            Dim DragTab As Tab = DirectCast(Data.GetData(GetType(Tab)), Tab)
            If DragTab IsNot MouseTab And DragTab IsNot Nothing And MouseTab IsNot Nothing AndAlso MouseTab.Index >= 0 Then
                TabPages.Remove(DragTab)
                TabPages.Insert(MouseTab.Index, DragTab)
            End If
        End If

    End Sub

    Private Sub MouseOver(Location As Point)

        Dim OldTab As Tab = MouseTab
        Dim OldZone As Zone = MouseZone
        Dim OldString As String = Join({If(OldTab Is Nothing, String.Empty, OldTab.ItemText), OldZone.ToString}, "|")

        Dim TabsBound As New Dictionary(Of Tab, Dictionary(Of Zone, Rectangle))
        For t = 0 To TabPages.Count - 1
            Dim TabItem As Tab = DirectCast(TabPages(t), Tab)
            TabsBound.Add(TabItem, New Dictionary(Of Zone, Rectangle))
            Dim TabBounds As Rectangle = GetTabRect(t)

            With TabItem
                .Bounds = TabBounds

                If TabItem Is AddTab Then
                    TabsBound(TabItem).Add(Zone.Add, TabBounds)
                    If TabBounds.Contains(Location) Then _MouseZone = Zone.Add

                Else
                    TabsBound(TabItem).Add(Zone.None, TabBounds)
#Region " IMAGE BOUNDS "
                    Dim ImageRegionBounds As Rectangle
                    Dim ImageBounds As Rectangle
                    If .Image Is Nothing Then
                        ImageRegionBounds = New Rectangle(TabBounds.X, TabBounds.Y, 0, TabBounds.Height)
                        ImageBounds = ImageRegionBounds
                    Else
                        ImageRegionBounds = New Rectangle(TabBounds.X, TabBounds.Y, .Image.Width, TabBounds.Height)
                        ImageBounds = New Rectangle(TabBounds.X + 3,
                                          TabBounds.Y + Convert.ToInt32((TabBounds.Height - .Image.Height) / 2),
                                          .Image.Width,
                                          .Image.Height)
                        TabsBound(TabItem).Add(Zone.Image, ImageBounds)
                    End If
#End Region
                    If ImageBounds.Contains(Location) Then _MouseZone = Zone.Image
#Region " TEXT BOUNDS "
                    Dim TextBounds As Rectangle
                    If If(.ItemText, String.Empty).Any Then
                        Dim TextSize As Size = TextRenderer.MeasureText(.ItemText, Font)
                        TextBounds = New Rectangle(ImageBounds.Right, TabBounds.Top, TextSize.Width, TabBounds.Height)
                    Else
                        TextBounds = New Rectangle(ImageBounds.Right, TabBounds.Top, 0, TabBounds.Height)
                        TabsBound(TabItem).Add(Zone.Text, TextBounds)
                    End If
#End Region
                    If TextBounds.Contains(Location) Then _MouseZone = Zone.Text
#Region " CLOSE BOUNDS "
                    Dim CloseRegionBounds As Rectangle
                    Dim CloseBounds As Rectangle
                    If .CanClose Then
                        CloseRegionBounds = New Rectangle(TextBounds.Right, TabBounds.Y, TabBounds.Right - TextBounds.Right, TabBounds.Height)
                        CloseBounds = New Rectangle(TabBounds.Right - (My.Resources.Close.Width + 4),
                                             TabBounds.Y + Convert.ToInt32((TabBounds.Height - My.Resources.Close.Height) / 2),
                                             My.Resources.Close.Width,
                                             My.Resources.Close.Height)
                        TabsBound(TabItem).Add(Zone.Close, CloseBounds)
                    Else
                        CloseBounds = New Rectangle(TabBounds.Right, TabBounds.Top, 0, TabBounds.Height)
                    End If
#End Region
                    If CloseRegionBounds.Contains(Location) Then _MouseZone = Zone.Close
                End If
                If TabBounds.Contains(Location) Then
                    _MouseTab = TabItem
                    If MouseOverSelection And TabItem IsNot AddTab Then SelectedIndex = .Index
                End If
            End With
        Next

        Dim Redraw As Boolean
        Dim NewString As String = Join({If(MouseTab Is Nothing, String.Empty, MouseTab.ItemText), MouseZone.ToString}, "|")
        If NewString <> OldString Then
            Dim ZoneChanged = OldZone <> MouseZone
            Dim TabChanged = OldTab IsNot MouseTab
            'Definitely draw if Tab changed
            Redraw = If(TabChanged, True, If(ZoneChanged, MouseTab.CanClose Or RedrawZoneChange, False))
            If Redraw Then Invalidate()
            If OldTab IsNot MouseTab Then RaiseEvent TabMouseChange(Me, New TabsEventArgs(MouseTab, OldTab, MouseZone, OldZone))
            If OldZone <> MouseZone Then RaiseEvent ZoneMouseChange(Me, New TabsEventArgs(MouseTab, OldTab, MouseZone, OldZone))
        End If

        If Parent IsNot Nothing Then
            Dim ParentText As String = Join({Location.ToString, OldString, NewString, Redraw}, " *** ")
            'Parent.Text = ParentText
            If Parent.Parent IsNot Nothing Then
                'Parent.Parent.Text = ParentText
                If Parent.Parent.Parent IsNot Nothing Then
                    'Parent.Parent.Parent.Text = ParentText
                End If
            End If
        End If

    End Sub

#End Region
    Private WithEvents TabPages_ As TabCollection
    Public Shadows ReadOnly Property TabPages As TabCollection
        Get
            Return TabPages_
        End Get
    End Property
    Private _UserCanAdd As Boolean = False
    Public Property UserCanAdd As Boolean
        Get
            Return _UserCanAdd
        End Get
        Set(value As Boolean)
            If value <> _UserCanAdd Then
                _UserCanAdd = value
                If value Then
                    _AddTab = TabPages.Add("aDdTaB", String.Empty, "PlusWhite")
                    AddTab.Image = ImageList.Images("PlusWhite")
                Else
                    If TabPages.ContainsKey("aDdTaB") Then TabPages.RemoveByKey("aDdTaB")
                End If
            End If
        End Set
    End Property
    Public Property UserCanReorder As Boolean
    Public Property UserCanName As Boolean
    Public Property MouseOverSelection As Boolean
    Public ReadOnly Property IdealWidth As Integer
        Get
            If TabPages.Pages.Any Then
                Return PointToScreen(New Point(0, 0)).X + TabPages.Pages.Max(Function(tp) tp.Bounds.Right) + 8
            Else
                Return Width
            End If
        End Get
    End Property
    Public ReadOnly Property RightBounds As Rectangle
        Get
            If TabPages.Pages.Any Then
                'Not sure index is 0-n, left to right considering drag and drop
                Return (From tp In TabPages.Pages Where tp.Bounds.Right = TabPages.Pages.Max(Function(p) p.Bounds.Right) Select tp.Bounds).First
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property AddTab As Tab
    Public Overloads ReadOnly Property SelectedTab As Tab
        Get
            Return TabPages.Item(SelectedIndex)
        End Get
    End Property
End Class

Public NotInheritable Class TabCollection
    Inherits TabControl.TabPageCollection
    'Implements IList(Of Tab)
    Friend ReadOnly Property Parent As Tabs
    Friend Shared ReadOnly Property Myself As TabCollection
    Public Event HeaderEnterExit(sender As Object, e As MouseEventArgs)
    Public Sub New()
        MyBase.New(Tabs.Myself)
        _Parent = Tabs.Myself
        _Myself = Me
    End Sub
    Friend ReadOnly Property Pages As List(Of Tab)
        Get
            Return (From p In Me Select DirectCast(p, Tab)).ToList
        End Get
    End Property
    Public Shadows Function Add(Text As String) As Tab
        Return Add(New Tab With {.ItemText = Text})
    End Function
    Public Shadows Function Add(Name As String, Text As String) As Tab
        Return Add(New Tab With {.Name = Name, .ItemText = Text})
    End Function
    Public Shadows Function Add(Name As String, Text As String, Image As String) As Tab
        Return Add(New Tab With {.Name = Name, .ItemText = Text, .ImageKey = Image})
    End Function
    Public Shadows Function Add(Name As String, Text As String, Image As Image) As Tab
        Return Add(New Tab With {.Name = Name, .ItemText = Text, .Image = Image})
    End Function
    Public Shadows Function Add(TabItem As Tab) As Tab

        If Parent IsNot Nothing And TabItem IsNot Nothing Then
            MyBase.Add(TabItem)
            Dim AddTab As Tab = Parent.AddTab
            If Parent.UserCanAdd And AddTab IsNot Nothing And AddTab IsNot TabItem Then
                Remove(AddTab)
                Parent.TabPages.Add(AddTab)
                'Insert(Count, AddTab)
            End If
        End If
        Return TabItem

    End Function
    '------------------------------------------------------------------------------------------
    Public Shadows Function Remove(Index As Integer) As Tab
        Return Remove(Item(Index))
    End Function
    Public Shadows Function Remove(Name As String) As Tab
        Return Remove(Item(Name))
    End Function
    Public Shadows Function Remove(TabItem As Tab) As Tab

        If TabItem IsNot Nothing Then
            MyBase.Remove(TabItem)
        End If
        Return TabItem

    End Function
    '------------------------------------------------------------------------------------------
    Public Shadows Function Item(Name As String) As Tab

        If Name Is Nothing Then
            Return Nothing
        Else
            Dim PagesByName = Pages.Where(Function(p) p.Name = Name)
            If PagesByName.Any Then
                Return PagesByName.First
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Shadows Function Item(Index As Integer) As Tab

        If Index >= 0 And Index < Count Then
            Dim PagesByIndex = Pages.Where(Function(p) p.Index = Index)
            If PagesByIndex.Any Then
                Return PagesByIndex.First
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function
End Class

Public Class Tab
    Inherits TabPage
    Private WithEvents TabControl As Tabs
    Private Overloads ReadOnly Property Parent As TabCollection
    Public Event HeaderClicked(sender As Object, e As MouseEventArgs)
    Public Event HeaderEntered(sender As Object, e As MouseEventArgs)
    Public Event HeadersLeft(sender As Object, e As EventArgs)
    Public Sub New()

        Parent = TabCollection.Myself
        TabControl = Parent.Parent
        _Font = TabControl.Font
        HeaderBackColor = BackColor
        HeaderForeColor = ForeColor

    End Sub
    Friend Overloads Property Bounds As Rectangle
    Friend Property BeforeBounds As Rectangle
    Friend Property AfterBounds As Rectangle
    Public ReadOnly Property Index As Integer
        Get
            Return Parent.IndexOf(Me)
        End Get
    End Property
    Private _CanClose As Boolean = True
    Public Property CanClose As Boolean
        Get
            If TabControl.AddTab Is Me Then _CanClose = False
            Return _CanClose
        End Get
        Set(value As Boolean)
            If value <> _CanClose Then
                _CanClose = value
                'Rework bounds
            End If
        End Set
    End Property
    Public ReadOnly Property Selected As Boolean
        Get
            Return TabControl.SelectedTab Is Me
        End Get
    End Property
    Public Property HeaderBackColor As Color = Color.Gray
    Public Property HeaderForeColor As Color = Color.Black
    Private _Font As Font = Nothing
    Public Overloads Property Font As Font
        Get
            Return _Font
        End Get
        Set(value As Font)
            If _Font Is Nothing Then
                _Font = value
            Else
                If value IsNot Nothing Then
                    If _Font.Name <> value.Name Or _Font.Size <> value.Size Or _Font.Bold <> value.Bold Or _Font.Italic <> value.Italic Then
                        _Font = value
                        'Rework bounds
                    End If
                End If
            End If
        End Set
    End Property
    Private _Image As Image
    Public Property Image As Image
        Get
            Return _Image
        End Get
        Set(value As Image)
            If Not SameImage(value, _Image) Then
                _Image = value
                'Rework bounds
            End If
        End Set
    End Property
    Private _ItemText As String
    Public Property ItemText As String
        Get
            Return _ItemText
        End Get
        Set(value As String)
            If value <> _ItemText Then
                _ItemText = value
                SetSafeControlPropertyValue(Me, "Text", value)
                'Rework bounds
            End If
        End Set
    End Property
End Class