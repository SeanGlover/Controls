Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class NodeEventArgs
    Inherits EventArgs
    Public ReadOnly Property Node As Node
    Public ReadOnly Property Nodes As List(Of Node)
    Public ReadOnly Property ProposedText As String = String.Empty
    Public Sub New(Value As Node)
        Node = Value
    End Sub
    Public Sub New(Values As List(Of Node))
        Nodes = Values
    End Sub
    Public Sub New(Value As Node, NewText As String)
        Node = Value
        ProposedText = NewText
    End Sub
End Class
Public Class TreeViewer
    Inherits Control

    Public Enum MouseRegion
        None
        Expander
        Favorite
        CheckBox
        Image
        Node
    End Enum
    Public Class HitRegion
        Implements IEquatable(Of HitRegion)
        Public Property Region As MouseRegion
        Public Property Node As Node
        Public Overrides Function GetHashCode() As Integer
            Return Region.GetHashCode Xor Node.GetHashCode
        End Function
        Public Overloads Function Equals(other As HitRegion) As Boolean Implements IEquatable(Of HitRegion).Equals
            If other Is Nothing Then
                Return Me Is Nothing
            Else
                Return Region = other.Region AndAlso Node Is other.Node
            End If
        End Function
        Public Shared Operator =(value1 As HitRegion, value2 As HitRegion) As Boolean
            If value1 Is Nothing Then
                Return value2 Is Nothing
            Else
                Return value1.Equals(value2)
            End If
        End Operator
        Public Shared Operator <>(value1 As HitRegion, value2 As HitRegion) As Boolean
            Return Not value1 = value2
        End Operator
        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is HitRegion Then
                Return CType(obj, HitRegion) = Me
            Else
                Return False
            End If
        End Function
    End Class
    Private Structure DragInfo
        Friend CursorSize As Size
        Friend MousePoints As List(Of Point)
        Friend IsDragging As Boolean
        Friend DragNode As Node
        Friend DropHighlightNode As Node
    End Structure

    Public WithEvents VScroll As New VScrollBar
    Public WithEvents HScroll As New HScrollBar
    Private WithEvents NodeTimer As New Timer With {.Interval = 200}
    Private WithEvents CursorTimer As New Timer With {.Interval = 300}
    Private WithEvents ScrollTimer As New Timer With {.Interval = 50}
#Region " TREEVIEW GLOBAL FUNCTIONS (CMS) "
    Private WithEvents IC_NodeAdd As New ImageCombo With {
        .Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "Add Child Node"
    }
    Private WithEvents IC_NodeEdit As New ImageCombo With {
        .Size = New Size(32, 100),
        .Margin = New Padding(0),
        .Visible = False,
        .Image = My.Resources.recycle
    }
#End Region
    Private Const CheckHeight As Integer = 14
    Private Const VScrollWidth As Integer = 14
    Private Const HScrollHeight As Integer = 12
    Private ExpandHeight As Integer = 10
    Private DragData As DragInfo
    Private _Cursor As Cursor
    Private VisibleIndex As Integer, RollingHeight As Integer, RollingWidth As Integer
    Private ExpandImage As Image, CollapseImage As Image
    Public Event Alert(sender As Object, e As AlertEventArgs)

#Region " STRUCTURES / ENUMS "
    Public Enum CheckState
        None
        All
        Mixed
    End Enum
    Private CheckboxStyle_ As CheckStyle = CheckStyle.Slide
    Public Property CheckboxStyle As CheckStyle
        Get
            Return CheckboxStyle_
        End Get
        Set(value As CheckStyle)
            CheckboxStyle_ = value
            Dim nodeStyleCheck = Nodes.Item("Options").Nodes.Find(Function(n) n.Name = "checkStyle")
            nodeStyleCheck.Image = If(value = CheckStyle.Check, My.Resources.checkbox, My.Resources.slideStateOn)
            RefreshNodesBounds_Lines(Nodes)
        End Set
    End Property
    Public Enum ExpandStyle
        PlusMinus
        Arrow
        Book
        LightBulb
    End Enum
    Private _ExpanderStyle As ExpandStyle = ExpandStyle.PlusMinus
    Public Property ExpanderStyle As ExpandStyle
        Get
            Return _ExpanderStyle
        End Get
        Set(value As ExpandStyle)
            _ExpanderStyle = value
            Select Case value
                Case ExpandStyle.Arrow
                    ExpandImage = Base64ToImage(ArrowCollapsed, True)
                    CollapseImage = Base64ToImage(ArrowExpanded, True)

                Case ExpandStyle.Book
                    ExpandImage = Base64ToImage(BookClosed, True)
                    CollapseImage = Base64ToImage(BookOpen, True)

                Case ExpandStyle.LightBulb
                    ExpandImage = Base64ToImage(LightOff, True)
                    CollapseImage = Base64ToImage(LightOn, True)

                Case ExpandStyle.PlusMinus
                    ExpandImage = Base64ToImage(DefaultCollapsed, True)
                    CollapseImage = Base64ToImage(DefaultExpanded, True)

            End Select
            Dim nodeOptions As Node = Nodes.Item("Options")
            If nodeOptions IsNot Nothing Then
                Dim nodeStyleExpand = (From n In Nodes.All Where n.Name = "expandStyle").First
                nodeStyleExpand.Image = ExpandImage
                nodeStyleExpand.Text = value.ToString
            End If
            ExpandHeight = ExpandImage.Height
            REM /// IF EXPAND/COLLAPSE HEIGHT CHANGES, THEN NODE.BOUNDS WILL BE AFFECTED 
            RequiresRepaint()
        End Set
    End Property
#End Region
    Public Sub New()

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)

        BackColor = Color.GhostWhite
        Controls.AddRange({VScroll, HScroll, IC_NodeEdit})
        ExpanderStyle = ExpandStyle.PlusMinus

#Region " GLOBAL OPTIONS SET-UP "
        Dim nodeOptions As Node = Nodes.Add("Options", "Options", My.Resources.tree)
        With nodeOptions
            With .Nodes.Add("sortChildren", "Sort children", My.Resources.sortNone)
                .Tag = SortOrder.None
            End With
            With .Nodes.Add("showIndex", "Show node index", My.Resources.slideStateOff)
                .Tag = False
            End With

            Dim nodeStyleCheck As Node = .Nodes.Add("checkStyle", "Checkbox style", My.Resources.slideStateOn)

            Dim nodeStyleShowHide As Node = .Nodes.Add("expandStyle", "Expander style", ExpandImage)

            Dim nodeExpandAll As Node = .Nodes.Add("expandAll", "Expand all", My.Resources.slideStateMixed)

            If 0 = 1 Then
                Dim nodeCanCheck As Node = .Nodes.Add("showCheckboxes", "Checkboxes", My.Resources.slideStateOn)

                Dim nodeCanMulti As Node = .Nodes.Add("allowMultiSelect", "Multi-select", My.Resources.slideStateOn)
            End If 'Not sure necessary

            .Nodes.ForEach(Sub(optionNode)
                               With optionNode
                                   .CanAdd = False
                                   .CanDragDrop = False
                                   .CanEdit = False
                                   .CanFavorite = False
                                   .CanRemove = False
                               End With
                           End Sub)
        End With
#End Region

    End Sub
    Protected Overrides Sub InitLayout()
        REM /// FIRES AFTER BEING ADDED TO ANOTHER CONTROL...ADD TREEVIEW AFTER LOADING NODES
        RequiresRepaint()
        MyBase.InitLayout()
    End Sub
    Private Sub WhenParentChanges() Handles Me.ParentChanged
        RequiresRepaint()
    End Sub
#Region " GLOBAL OPTIONS "
    Private Sub NodeOption_Clicked(optionClicked As Node)

        Select Case optionClicked.Name
            Case "checkStyle"
                CheckboxStyle = If(CheckboxStyle = CheckStyle.Check, CheckStyle.Slide, CheckStyle.Check)

            Case "expandStyle"
                ExpanderStyle = If(ExpanderStyle = ExpandStyle.Arrow, ExpandStyle.Book, If(ExpanderStyle = ExpandStyle.Book, ExpandStyle.LightBulb, If(ExpanderStyle = ExpandStyle.LightBulb, ExpandStyle.PlusMinus, ExpandStyle.Arrow)))

            Case "sortChildren"
                Dim sortUp As String = "iVBORw0KGgoAAAANSUhEUgAAABEAAAALCAYAAACZIGYHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVChTY2SAgAYoTQ4A6z0PxP8pwPep4hLG//9BhlEGyHFJAzaLSQ2T+yBDkDFVXEL3MMEaFjBAbJhghAUMU8ElDAwAvNhdwMSXsO4AAAAASUVORK5CYII="
                Dim sortDown As String = "iVBORw0KGgoAAAANSUhEUgAAABEAAAALCAYAAACZIGYHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABLSURBVChTY2SAgAYojQ80/P//H8rEBOeBGCRLCN8HGYINU8UljLgkSAGkuAQGsLqI2DCBYYywoYpLBixM0AFYL6lhgo7vU8ElDA0AaFFdwFj1ubQAAAAASUVORK5CYII="
                Dim nodeOrder As SortOrder = DirectCast(optionClicked.Tag, SortOrder)
                nodeOrder = If(nodeOrder = SortOrder.None, SortOrder.Ascending, If(nodeOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.None))
                optionClicked.Image = If(nodeOrder = SortOrder.None, My.Resources.sortNone, Base64ToImage(If(nodeOrder = SortOrder.Ascending, sortDown, sortUp)))
                SelectedNodes.ForEach(Sub(nodeSelected)
                                          nodeSelected.SortChildren(nodeOrder)
                                      End Sub)
                optionClicked.Tag = nodeOrder

            Case "showIndex"
                Dim showIndex As Boolean = Not DirectCast(optionClicked.Tag, Boolean)
                optionClicked.Image = If(showIndex, My.Resources.slideStateOn, My.Resources.slideStateOff)
                SelectedNodes.ForEach(Sub(nodeSelected)
                                          nodeSelected.ShowIndex = showIndex
                                          nodeSelected.Children.ForEach(Sub(nodeChild)
                                                                            nodeChild.ShowIndex = showIndex
                                                                        End Sub)
                                      End Sub)
                optionClicked.Tag = showIndex

            Case "expandAll"
                Dim expandAll As Boolean = Not SameImage(optionClicked.Image, My.Resources.slideStateOn)
                If expandAll Then ExpandNodes()
                If Not expandAll Then CollapseNodes()
                optionClicked.Image = If(expandAll, My.Resources.slideStateOn, My.Resources.slideStateOff)

                'Dim nodeCanCheck As Node = .Nodes.Add("showCheckboxes", "Checkboxes", My.Resources.slideStateOn)
                'AddHandler NodeClicked, AddressOf NodeOption_Clicked

                'Dim nodeCanMulti As Node = .Nodes.Add("allowMultiSelect", "Multi-select", My.Resources.slideStateOn)
                'AddHandler NodeClicked, AddressOf NodeOption_Clicked

        End Select
#Region " MULTI-SELECT / CHECKBOX FUNCTIONS - NOT SURE NEEDED "
        '        If sender Is TSMI_MultiSelect Then MultiSelect = TSMI_MultiSelect.Checked
        '        ToggleSelect()

        '        If sender Is TSMI_SelectAll Then
        '            If MultiSelect Then
        '                For Each Node In Nodes.All
        '                    Node._Selected = True
        '                Next
        '                Invalidate()
        '            End If
        '        ElseIf sender Is TSMI_SelectNone Then
        '            For Each Node In Nodes.All
        '                Node._Selected = False
        '            Next
        '            Invalidate()
        '        End If
        '#End Region
        '#Region " CHECKBOX FUNCTIONS "
        '        TSMI_CheckUncheckAll.Visible = TSMI_Checkboxes.Checked
        '        If Not TSMI_Checkboxes.Checked Then TSMI_Checkboxes.HideDropDown()
        '        If sender Is TSMI_Checkboxes Then
        '            If TSMI_Checkboxes.Checked Then
        '                CheckBoxes = CheckState.All
        '            Else
        '                CheckBoxes = CheckState.None
        '            End If
        '        End If
        '        If sender Is TSMI_CheckUncheckAll Then
        '            TSMI_CheckUncheckAll.Text = If(TSMI_CheckUncheckAll.Checked, "UnCheck All", "Check All").ToString(InvariantCulture)
        '            CheckAll = TSMI_CheckUncheckAll.Checked
        '        End If
        '#End Region
        'TSMI_SelectAll.Visible = MultiSelect
        'If MultiSelect Then
        '    TSMI_MultiSelect.Text = "Multi-Select".ToString(InvariantCulture)
        'Else
        '    TSMI_MultiSelect.Text = "Single-Select".ToString(InvariantCulture)
        'End If
#End Region

    End Sub
    Private Sub NodeAdd_Submitted() Handles IC_NodeAdd.ValueSubmitted, IC_NodeAdd.ItemSelected

        Dim nodeParentCollection As NodeCollection = If(SelectedNodes.Any, SelectedNodes.First.Nodes, Nodes)
        If nodeParentCollection.Any Then
            '// [1] Add a new Node from the Text of the ImageCombo
            Dim Items As New List(Of Node)({New Node With {.Text = IC_NodeAdd.Text, .BackColor = Color.Lavender}})

            '// [2] Potentially add multiple nodes from any selected items in the ImageCombo.DropDown
            Items.AddRange(From I In IC_NodeAdd.Items Where Not I.Text = IC_NodeAdd.Text And I.Checked Select New Node With {.Text = I.Text, .BackColor = Color.Lavender})

            If Items.Count = 1 Then '// Only adding from [1]
                Dim Item As Node = Items.First
                RaiseEvent NodeBeforeAdded(Me, New NodeEventArgs(Item, IC_NodeAdd.Text))

            Else
                '// Coming from [2] but multiple adds doesn't allow for NodeBeforeAdded
                nodeParentCollection.AddRange(Items)
                RaiseEvent NodeAfterAdded(Me, New NodeEventArgs(Items))
            End If
        End If

    End Sub
    Private AddNode_OK_ As Boolean
    Public Property AddNode_OK(Optional nodeAdd As Node = Nothing) As Boolean
        Get
            Return AddNode_OK_
        End Get
        Set(value As Boolean)
            AddNode_OK_ = value
            If nodeAdd IsNot Nothing Then 'If nodeAdd Is Nothing Then it is a reset
                If value Then
                    Dim nodeParentCollection As NodeCollection = If(SelectedNodes.Any, SelectedNodes.First.Nodes, Nodes)
                    nodeParentCollection.Add(nodeAdd)
                    RaiseEvent NodeAfterAdded(Me, New NodeEventArgs(nodeAdd, IC_NodeAdd.Text))
                Else
                    RaiseEvent NodeAfterAdded(Me, New NodeEventArgs(nodeAdd, IC_NodeAdd.Text))
                    nodeAdd.Dispose()
                End If
            End If
        End Set
    End Property
    Private Sub NodeEdit_Submitted() Handles IC_NodeEdit.ValueSubmitted

        Dim editNode As Node = DirectCast(IC_NodeEdit.Tag, Node)
        IC_NodeEdit.Visible = False
        If IC_NodeEdit.Text <> editNode.Text Then RaiseEvent NodeBeforeEdited(Me, New NodeEventArgs(editNode, IC_NodeEdit.Text))

    End Sub
    Private EditNode_OK_ As Boolean
    Public Property EditNode_OK(Optional nodeEdit As Node = Nothing) As Boolean
        Get
            Return EditNode_OK_
        End Get
        Set(value As Boolean)
            EditNode_OK_ = value
            If nodeEdit IsNot Nothing Then 'If nodeEdit Is Nothing Then it is a reset
                If value Then nodeEdit.Text = IC_NodeEdit.Text
                RaiseEvent NodeAfterEdited(Me, New NodeEventArgs(nodeEdit, IC_NodeEdit.Text))
                IC_NodeEdit.Text = String.Empty
            End If
        End Set
    End Property
    Private Sub NodeRemove_Submitted(sender As Object, e As EventArgs) Handles IC_NodeEdit.ImageClicked

        Dim nodeRemove As Node = DirectCast(IC_NodeEdit.Tag, Node)
        RaiseEvent NodeBeforeRemoved(Me, New NodeEventArgs(nodeRemove))

    End Sub
    Private RemoveNode_OK_ As Boolean
    Public Property RemoveNode_OK(Optional nodeRemove As Node = Nothing) As Boolean
        Get
            Return RemoveNode_OK_
        End Get
        Set(value As Boolean)
            RemoveNode_OK_ = value
            If nodeRemove IsNot Nothing Then 'If nodeRemove Is Nothing Then it is a reset
                If value Then
                    nodeRemove.RemoveMe()
                    IC_NodeEdit.Visible = False
                    RaiseEvent NodeAfterRemoved(Me, New NodeEventArgs(nodeRemove))
                End If
            End If
        End Set
    End Property
#End Region
#Region " PROPERTIES "
    Public Property FavoriteImage As Image = Base64ToImage(StarString)
    Public ReadOnly Property OptionsOpen As Boolean
    Public Property FavoritesFirst As Boolean = True
    Public ReadOnly Property DropHighlightNode As Node
        Get
            Return DragData.DropHighlightNode
        End Get
    End Property
    Public Property MouseOverExpandsNode As Boolean = False
    Public Property CanAdd As Boolean = True
    Private _MultiSelect As Boolean
    Public Property MultiSelect As Boolean
        Get
            Return _MultiSelect
        End Get
        Set(value As Boolean)
            If _MultiSelect <> value Then
                If Not value Then
                    For Each Node In Nodes.All
                        Node._Selected = False
                    Next
                End If
                _MultiSelect = value
                Invalidate()
            End If
        End Set
    End Property
    Public ReadOnly Property UnRestrictedSize As Size
        Get
            Return New Size(VScrollWidth + RollingWidth + Offset.X, RollingHeight + Offset.Y)
        End Get
    End Property
    Public ReadOnly Property NodeHeight As Integer
        Get
            Return CInt(Nodes.All.Average(Function(a) a.Height))
        End Get
    End Property
    Public ReadOnly Property SelectedNodes As List(Of Node)
        Get
            Return Nodes.All.Where(Function(n) n.Selected).ToList
        End Get
    End Property
    Public ReadOnly Property Nodes As New NodeCollection(Me)
    Public Property MaxNodes As Integer
    Public Property LineStyle As DashStyle = DashStyle.Dot
    Public Property LineColor As Color = Color.Blue
    Public Property RootLines As Boolean = True
    Public Property DropHighlightColor As Color = Color.Gainsboro
    Public Property MouseOverColor As Color = Color.Gainsboro
    Public Property SelectionColor As Color = Color.Gainsboro
    Public Property Offset As New Point(5, 3)
    Public Overrides Property AutoSize As Boolean = True
    Public Property StopMe As Boolean
    Friend ReadOnly Property Hit As HitRegion
    Private _CheckBoxes As CheckState = CheckState.Mixed
    Public Property CheckBoxes As CheckState
        Get
            Return _CheckBoxes
        End Get
        Set(value As CheckState)
            If _CheckBoxes <> value Then
                If value = CheckState.All Then
                    For Each Node In Nodes.All
                        Node.CheckBox = True
                    Next
                ElseIf value = CheckState.None Then
                    For Each Node In Nodes.All
                        Node.CheckBox = False
                    Next
                End If
                _CheckBoxes = value
            End If
        End Set
    End Property
    Private _CheckAll As Boolean = False
    Public Property CheckAll As Boolean
        Get
            Return _CheckAll
        End Get
        Set(value As Boolean)
            If _CheckAll <> value Then
                Dim CheckNodes = Nodes.All.Where(Function(c) c.CheckBox)
                If _CheckAll And Not CheckNodes.Any Then CheckBoxes = CheckState.All
                For Each Node In Nodes.All.Where(Function(c) c.CheckBox)
                    Node.Checked = value
                Next
                _CheckAll = value
            End If
        End Set
    End Property
#End Region
#Region " PUBLIC EVENTS "
    Public Event NodesChanged(sender As Object, e As NodeEventArgs)
    Public Event NodeBeforeAdded(sender As Object, e As NodeEventArgs)
    Public Event NodeAfterAdded(sender As Object, e As NodeEventArgs)
    Public Event NodeBeforeRemoved(sender As Object, e As NodeEventArgs)
    Public Event NodeAfterRemoved(sender As Object, e As NodeEventArgs)
    Public Event NodeBeforeEdited(sender As Object, e As NodeEventArgs)
    Public Event NodeAfterEdited(sender As Object, e As NodeEventArgs)
    Public Event NodeDragStart(sender As Object, e As NodeEventArgs)
    Public Event NodeDragOver(sender As Object, e As NodeEventArgs)
    Public Event NodeDropped(sender As Object, e As NodeEventArgs)
    Public Event NodeChecked(sender As Object, e As NodeEventArgs)
    Public Event NodeExpanded(sender As Object, e As NodeEventArgs)
    Public Event NodeCollapsed(sender As Object, e As NodeEventArgs)
    Public Event NodeClicked(sender As Object, e As NodeEventArgs)
    Public Event NodeRightClicked(sender As Object, e As NodeEventArgs)
    Public Event NodeFavorited(sender As Object, e As NodeEventArgs)
    Public Event NodeDoubleClicked(sender As Object, e As NodeEventArgs)
#End Region
#Region " MOUSE EVENTS "
    Private LastMouseNode As Node = Nothing
    Private MousePoint As Point
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)

        If e IsNot Nothing Then
            Dim HitRegion As HitRegion = HitTest(e.Location)
            Dim HitNode As Node = HitRegion.Node
            IC_NodeEdit.Visible = False
            DragData = New DragInfo With {
                .DragNode = HitNode,
                .IsDragging = False,
                .MousePoints = New List(Of Point)
            }

            If e.Button = MouseButtons.Right Then
                If HitNode Is Nothing Then

                ElseIf HitNode.CanEdit Then
                    With IC_NodeEdit
                        .Text = HitNode.Text
                        .Image = If(HitNode.CanRemove, My.Resources.recycle, Nothing)
                        .HintText = HitNode.Text
                        Dim nodeXY As Point = HitNode.Bounds_Text.Location
                        nodeXY.Offset(-1, -1)
                        .Location = nodeXY
                        .Size = HitNode.Bounds_Text.Size
                        .MaximumSize = New Size(1000, .Size.Height)
                        .MinimumSize = New Size(10, .Size.Height)
                        .MinimumSize = New Size(.IdealSize.Width, .Size.Height)
                        .AutoSize = True
                        .Visible = True
                        .Tag = HitNode
                    End With
                End If

            Else
                If HitNode IsNot Nothing Then
                    With HitNode
                        If .Parent?.Name = "Options" Then
                            NodeOption_Clicked(HitNode)
                        Else
                            Select Case HitRegion.Region
                                Case MouseRegion.Expander
                                    ._Clicked = True
                                    If .HasChildren Then
                                        ._Expanded = Not .Expanded
                                        If .Expanded Then
                                            .Expand()
                                            RaiseEvent NodeExpanded(Me, New NodeEventArgs(HitNode))
                                        Else
                                            .Collapse()
                                            RaiseEvent NodeCollapsed(Me, New NodeEventArgs(HitNode))
                                        End If
                                    End If

                                Case MouseRegion.Favorite
                                    .Favorite = Not .Favorite
                                    RaiseEvent NodeFavorited(Me, New NodeEventArgs(HitNode))

                                Case MouseRegion.CheckBox
                                    ._Clicked = True
                                    .Checked = Not .Checked
                                    RaiseEvent NodeChecked(Me, New NodeEventArgs(HitNode))

                                Case MouseRegion.Image, MouseRegion.Node
                                    ._Clicked = True
                                    If Not MultiSelect Then
                                        For Each Node In SelectedNodes.Except({HitNode})
                                            Node._Selected = False
                                        Next
                                    End If
                                    HitNode._Selected = Not HitNode.Selected
                                    If e.Button = MouseButtons.Left Then
                                        RaiseEvent NodeClicked(Me, New NodeEventArgs(HitNode))
                                    ElseIf e.Button = MouseButtons.Right Then
                                        RaiseEvent NodeRightClicked(Me, New NodeEventArgs(HitNode))
                                    End If

                            End Select
                        End If
                    End With
                    Invalidate()
                End If
            End If
        End If
        MyBase.OnMouseDown(e)

    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)

        If e IsNot Nothing Then
            MousePoint = e.Location
            If e.Button = MouseButtons.None Then
                Dim hitInfo = HitTest(e.Location)
                _Hit = hitInfo
                If hitInfo.Node IsNot LastMouseNode Then
                    If MouseOverExpandsNode Then
                        If hitInfo.Node Is Nothing Then
                            LastMouseNode.Collapse()
                        Else
                            hitInfo.Node.Expand()
                        End If
                    End If
                    LastMouseNode = hitInfo.Node
                    Invalidate()
                End If

            ElseIf e.Button = MouseButtons.Left Then
                With DragData
                    If Not .MousePoints.Contains(e.Location) Then .MousePoints.Add(e.Location)
                    .IsDragging = .DragNode IsNot Nothing AndAlso (.MousePoints.Count >= 5 And Not .DragNode.Bounds_Text.Contains(.MousePoints.Last))
                    If .IsDragging Then
                        OnDragStart()
                        Dim Data As New DataObject
                        Data.SetData(GetType(Node), .DragNode)
                        MyBase.OnDragOver(New DragEventArgs(Data, 0, e.X, e.Y, DragDropEffects.Copy Or DragDropEffects.Move, DragDropEffects.All))
                        DoDragDrop(Data, DragDropEffects.Copy Or DragDropEffects.Move)
                    End If
                End With

            End If
        End If
        MyBase.OnMouseMove(e)

    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)

        Cursor = Cursors.Default
        DragData.IsDragging = False
        DragData.MousePoints.Clear()
        ScrollTimer.Stop()
        MyBase.OnMouseUp(e)

    End Sub
    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)

        If e IsNot Nothing AndAlso e.Button = MouseButtons.Left Then
            Dim HitRegion As HitRegion = HitTest(e.Location)
            Dim HitNode As Node = HitRegion.Node
            If Not IsNothing(HitNode) Then
                RaiseEvent NodeDoubleClicked(Me, New NodeEventArgs(HitNode))
            End If
        End If
        MyBase.OnMouseDoubleClick(e)

    End Sub
#End Region
#Region " KEYPRESS EVENTS "
    Private Sub On_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs)

        If e IsNot Nothing Then
            Select Case e.KeyCode
                Case Keys.Up, Keys.Down, Keys.Left, Keys.Right, Keys.Tab, Keys.ControlKey
                    e.IsInputKey = True
            End Select
        End If

    End Sub
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
        If e IsNot Nothing AndAlso e.Handled Then
        End If
        MyBase.OnKeyPress(e)
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

        If e IsNot Nothing AndAlso e.Modifiers = Keys.Control AndAlso e.KeyCode = Keys.C Then
            Dim Nodes2Clipboard As New List(Of String)(From sn In SelectedNodes Select sn.Text)
            Clipboard.SetText(Join(Nodes2Clipboard.ToArray, vbNewLine))
        End If
        MyBase.OnKeyDown(e)

    End Sub
#End Region
#Region " DRAG & DROP "
    Private Sub OnDragStart()

        With DragData
            If .DragNode.CanDragDrop Then
#Region " CUSTOM CURSOR WITH NODE.TEXT "
                Using nodeFont As New Font(.DragNode.Font.FontFamily, .DragNode.Font.Size + 4, FontStyle.Bold)
                    Dim textSize As Size = TextRenderer.MeasureText(.DragNode.Text, nodeFont)
                    Dim cursorBounds As New Rectangle(New Point(0, 0), New Size(3 + textSize.Width + 3, 2 + textSize.Height + 2))
                    .CursorSize = cursorBounds.Size
                    Dim shadowDepth As Integer = 16
                    cursorBounds.Inflate(shadowDepth, shadowDepth)
                    Using bmp As New Bitmap(cursorBounds.Width, cursorBounds.Height)
                        Using Graphics As Graphics = Graphics.FromImage(bmp)
                            Graphics.SmoothingMode = SmoothingMode.AntiAlias
                            Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
                            Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
                            Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
                            Dim fadingRectangle As New Rectangle(0, 0, bmp.Width, bmp.Height)
                            Dim fadeFactor As Integer
                            For P = 0 To shadowDepth - 1
                                fadeFactor = 16 + (P * 5)
                                fadingRectangle.Inflate(-1, -1)
                                Using fadingBrush As New SolidBrush(Color.FromArgb(fadeFactor, .DragNode.CursorGlowColor)) '16, 21, 26, 31, 36
                                    Using fadingPen As New Pen(fadingBrush, 1)
                                        Graphics.DrawRectangle(fadingPen, fadingRectangle)
                                    End Using
                                End Using
                            Next
                            Using fadedBrush As New SolidBrush(Color.FromArgb(fadeFactor, Color.Gainsboro))
                                Graphics.FillRectangle(fadedBrush, fadingRectangle)
                            End Using
                            Graphics.DrawRectangle(Pens.Silver, fadingRectangle)
                            Using Format As New StringFormat With {
        .Alignment = StringAlignment.Center,
        .LineAlignment = StringAlignment.Center
    }
                                Graphics.DrawString(.DragNode.Text, nodeFont, Brushes.Black, fadingRectangle, Format)
                            End Using
                        End Using
                        _Cursor = CursorHelper.CreateCursor(bmp, 0, Convert.ToInt32(bmp.Height / 2))
                    End Using
                End Using

#End Region
                REM /// NOW SET DATA
                Dim Data As New DataObject
                Data.SetData(GetType(Node), .DragNode)
                AllowDrop = True
                RaiseEvent NodeDragStart(Me, New NodeEventArgs(.DragNode))
            Else
                CursorTimer.Start()
            End If

        End With

    End Sub
    Private Sub CursorTimerTick() Handles CursorTimer.Tick

        CursorTimer.Tag = If(CursorTimer.Tag, 0)
        Dim Index As Integer = DirectCast(If(CursorTimer.Tag, 0), Integer)
        Dim im As Integer = Index Mod 3
        Using CursorImage As Bitmap = If(im = 0, My.Resources.FingerLeft, If(im = 1, My.Resources.Finger, My.Resources.FingerRight))
            Index += 1
            If Index >= 30 Then
                Index = 0
                CursorTimer.Stop()
                _Cursor = Cursors.Default
            Else
                _Cursor = CursorHelper.CreateCursor(CursorImage, 0, Convert.ToInt32(CursorImage.Height / 2))
            End If
            CursorTimer.Tag = Index
        End Using

    End Sub
    Protected Overrides Sub OnDragLeave(e As EventArgs)

        DragData.DropHighlightNode = Nothing
        Invalidate()
        ScrollTimer.Tag = Nothing

        Dim MouseLocation As Point = PointToClient(New Point(0, 0))
        MouseLocation.Offset(Cursor.Position)
        If VScroll.Visible Then
            If MouseLocation.Y <= 0 Then
                REM /// EXITED TOP
                ScrollTimer.Tag = "Up"
                ScrollTimer.Start()

            ElseIf MouseLocation.Y >= Height Then
                REM /// EXITED BOTTOM
                ScrollTimer.Tag = "Down"
                ScrollTimer.Start()
            End If
        End If
        MyBase.OnDragLeave(e)

    End Sub
    Protected Overrides Sub OnDragEnter(e As DragEventArgs)

        ScrollTimer.Stop()
        Invalidate()
        MyBase.OnDragEnter(e)

    End Sub
    Protected Overrides Sub OnDragOver(e As DragEventArgs)

        If e IsNot Nothing Then
            e.Effect = DragDropEffects.All
            Dim regionHit As HitRegion = DragHit(e)
            Dim nodeHit As Node = regionHit?.Node

            If nodeHit IsNot DragData.DropHighlightNode Then
                DragData.DropHighlightNode = nodeHit
                Invalidate()
                If nodeHit IsNot Nothing Then
                    e.Data.SetData(GetType(Object), nodeHit)
                    With nodeHit
                        If regionHit.Region = MouseRegion.Expander And .HasChildren Then
                            If .Expanded Then
                                .Collapse()
                            Else
                                .Expand()
                            End If
                        End If
                    End With
                End If
            End If
        End If
        MyBase.OnDragOver(e)

    End Sub
    Private Function DragHit(e As DragEventArgs) As HitRegion

        Dim cursorLocation As Point = PointToClient(New Point(e.X, e.Y))
        Dim cursorBounds As New Rectangle(cursorLocation, DragData.CursorSize)
        cursorLocation.Offset(0, -CInt(cursorBounds.Height / 2)) 'CInt(cursorBounds.Width / 2)CInt(cursorBounds.Height / 2)
        cursorBounds.Offset(0, -CInt(cursorBounds.Height / 2))
        Return HitTest(cursorBounds)

    End Function
    Protected Overrides Sub OnDragDrop(e As DragEventArgs)

        If e IsNot Nothing Then
            CursorTimer.Stop()
            Dim nodeBeingDragged As Node = TryCast(e.Data.GetData(GetType(Node)), Node)
            Dim regionHit As HitRegion = DragHit(e)
            Dim nodeHit As Node = regionHit?.Node
            If nodeBeingDragged?.CanDragDrop And nodeHit?.CanDragDrop Then
                e.Data.SetData(GetType(Node), nodeHit)
                RaiseEvent NodeDropped(Me, New NodeEventArgs(nodeHit))
            End If
            DragData.DropHighlightNode = Nothing
            RequiresRepaint()
            MyBase.OnDragDrop(e)
        End If

    End Sub
    Protected Overrides Sub OnGiveFeedback(e As GiveFeedbackEventArgs)

        If e IsNot Nothing Then
            e.UseDefaultCursors = False
            Cursor.Current = _Cursor
        End If
        MyBase.OnGiveFeedback(e)

    End Sub
    Private Sub DragScroll() Handles ScrollTimer.Tick

        If MouseButtons.HasFlag(MouseButtons.Left) And Not IsNothing(ScrollTimer.Tag) Then
            UpdateDrawNodes()
            Dim VScrollValue As Integer = VScroll.Value
            Dim Delta As Integer = 0
            If ScrollTimer.Tag.ToString = "Up" Then
                REM /// EXITED TOP
                VScrollValue -= VScroll.SmallChange
                VScrollValue = {VScroll.Minimum, VScrollValue}.Max
                Delta = VScroll.Value - VScrollValue
                VScroll.Value = VScrollValue
                VScrollUpDown(Delta)

            ElseIf ScrollTimer.Tag.ToString = "Down" Then
                REM /// EXITED TOP
                VScrollValue += VScroll.SmallChange
                VScrollValue = {VScroll.Maximum - Height, VScrollValue}.Min
                Delta = VScroll.Value - VScrollValue
                VScroll.Value = VScrollValue
                VScrollUpDown(Delta)

            End If
            If Delta = 0 Then ScrollTimer.Stop()
        Else
            ScrollTimer.Stop()

        End If

    End Sub
#End Region
#Region " INVALIDATION "
    Dim RR As Boolean
    Protected Overrides Sub OnPaint(e As PaintEventArgs)

        If e IsNot Nothing Then
            With e.Graphics
                .SmoothingMode = SmoothingMode.AntiAlias
                .TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
                Using backBrush As New SolidBrush(BackColor)
                    .FillRectangle(backBrush, ClientRectangle)
                End Using
                If BackgroundImage IsNot Nothing Then
                    Dim xOffset As Integer = {ClientRectangle.Width, Math.Abs(Convert.ToInt32((ClientRectangle.Width - BackgroundImage.Width) / 2))}.Min
                    Dim yOffset As Integer = {ClientRectangle.Height, Math.Abs(Convert.ToInt32((ClientRectangle.Height - BackgroundImage.Height) / 2))}.Min
                    Dim Bounds_Image As New Rectangle(xOffset, yOffset, {ClientRectangle.Width, BackgroundImage.Width}.Min, {ClientRectangle.Height, BackgroundImage.Height}.Min)
                    .DrawImage(BackgroundImage, Bounds_Image)
                End If
                If Nodes.Any Then
                    Dim drawRootLines As Boolean = RootLines And Nodes.Count > 1 'Doesn't make sense if only one root, when User wants RootLines
                    Dim firstRootNode As Node = Nodes.First
                    Dim verticalRootLine_x As Integer = CInt({
                                            firstRootNode.Bounds_Check.Left,
                                            firstRootNode.Bounds_Favorite.Left,
                                            firstRootNode.Bounds_Image.Left,
                                            firstRootNode.Bounds_ShowHide.Left
                                                                          }.Min / 2)
                    Using linePen As New Pen(LineColor) With {.DashStyle = LineStyle}
                        For Each Node As Node In Nodes.Draw
                            With Node
                                Dim mouseInTip As Boolean = False
                                If .Separator = Node.SeparatorPosition.Above And Node IsNot firstRootNode Then
                                    Using Pen As New Pen(Color.Blue, 1)
                                        Pen.DashStyle = DashStyle.DashDot
                                        e.Graphics.DrawLine(Pen, New Point(0, .Bounds_Text.Top), New Point(ClientRectangle.Right, .Bounds_Text.Top))
                                    End Using

                                ElseIf .Separator = Node.SeparatorPosition.Below And Node IsNot Nodes.Last Then
                                    Using Pen As New Pen(Color.Blue, 1)
                                        Pen.DashStyle = DashStyle.DashDot
                                        e.Graphics.DrawLine(Pen, New Point(0, .Bounds_Text.Bottom), New Point(ClientRectangle.Right, .Bounds_Text.Bottom))
                                    End Using

                                End If
                                If .CanFavorite Then e.Graphics.DrawImage(If(.Favorite, My.Resources.star, My.Resources.starEmpty), .Bounds_Favorite)

                                If .TipText IsNot Nothing Then
                                    Dim triangleHeight As Single = 8
                                    Dim trianglePoints As New List(Of PointF) From {
                                            New PointF(.Bounds_Text.Left, .Bounds_Text.Top),
                                            New PointF(.Bounds_Text.Left + triangleHeight, .Bounds_Text.Top),
                                            New PointF(.Bounds_Text.Left, .Bounds_Text.Top + triangleHeight)
                                    }
                                    e.Graphics.FillPolygon(Brushes.DarkOrange, trianglePoints.ToArray)
                                    mouseInTip = InTriangle(MousePoint, trianglePoints.ToArray)
                                End If
                                If .Image IsNot Nothing Then
                                    e.Graphics.DrawImage(.Image, .Bounds_Image)
                                    If Hit?.Node Is Node And Hit?.Region = MouseRegion.Image Then
                                        Using imageBrush As New SolidBrush(Color.FromArgb(96, MouseOverColor))
                                            e.Graphics.FillRectangle(imageBrush, .Bounds_Image)
                                        End Using
                                    End If
                                End If

                                ._Bounds_Text.Width = 2 + MeasureText(If(mouseInTip, .TipText, .Text), .Font, e.Graphics).Width + 2
                                Dim boundsNode As Rectangle = .Bounds_Text
                                Using sf As New StringFormat With {
                                            .Alignment = StringAlignment.Near,
                                            .LineAlignment = StringAlignment.Center,
                                            .FormatFlags = StringFormatFlags.NoWrap Or StringFormatFlags.NoClip
                                        }
                                    boundsNode.Inflate(3, 0)
                                    boundsNode.Offset(3, 0)
                                    Using backBrush As New SolidBrush(If(DragData.DropHighlightNode Is Node, DropHighlightColor, .BackColor))
                                        e.Graphics.FillRectangle(backBrush, boundsNode)
                                    End Using
                                    Using textBrush As New SolidBrush(.ForeColor)
                                        e.Graphics.DrawString(If(mouseInTip, .TipText, .Text),
                                                                               .Font,
                                                                               textBrush,
                                                                               boundsNode,
                                                                               sf)
                                    End Using
                                    boundsNode = .Bounds_Text
                                End Using
                                If Hit?.Node Is Node And .TipText Is Nothing Then
                                    Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, MouseOverColor))
                                        e.Graphics.FillRectangle(SemiTransparentBrush, boundsNode)
                                    End Using
                                End If
                                If .Selected Then
                                    boundsNode.Inflate(0, -1)
                                    boundsNode.Offset(0, -1)
                                    Using SemiTransparentBrush As New SolidBrush(Color.FromArgb(128, SelectionColor))
                                        e.Graphics.FillRectangle(SemiTransparentBrush, boundsNode)
                                    End Using
                                    Using dottedPen As New Pen(SystemBrushes.ControlText, 1)
                                        dottedPen.DashStyle = DashStyle.DashDot
                                        e.Graphics.DrawRectangle(dottedPen, boundsNode)
                                    End Using
                                End If
                                If .CheckBox Then
                                    If CheckboxStyle = CheckStyle.Check Then
                                        '/// Check background as White or Gray
                                        Using checkBrush As New SolidBrush(If(.PartialChecked, Color.FromArgb(192, Color.LightGray), Color.White))
                                            e.Graphics.FillRectangle(checkBrush, .Bounds_Check)
                                        End Using
                                        ''/// Draw the checkmark ( only if .Checked or .PartialChecked )
                                        If .Checked Or .PartialChecked Then
                                            Using CheckFont As New Font("Marlett", 10)
                                                TextRenderer.DrawText(e.Graphics, "a".ToString(InvariantCulture), CheckFont, .Bounds_Check, If(.Checked, Color.Blue, Color.DarkGray), TextFormatFlags.NoPadding Or TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
                                            End Using
                                        End If
                                        '/// Draw the surrounding Check square
                                        Using Pen As New Pen(Color.Blue, 1)
                                            e.Graphics.DrawRectangle(Pen, .Bounds_Check)
                                        End Using

                                    ElseIf CheckboxStyle = CheckStyle.Slide Then
                                        e.Graphics.DrawImage(If(.Checked, My.Resources.slideStateOn, If(.PartialChecked, My.Resources.slideStateMixed, My.Resources.slideStateOff)), .Bounds_Check)
                                        If Hit?.Node Is Node And Hit?.Region = MouseRegion.CheckBox Then
                                            Using checkBrush As New SolidBrush(Color.FromArgb(96, MouseOverColor))
                                                e.Graphics.FillRectangle(checkBrush, .Bounds_Check)
                                            End Using
                                        End If

                                    End If
                                End If

                                Dim objectBounds As Rectangle = .Bounds_ShowHide
                                Dim objectCenter As Integer = CInt(objectBounds.Height / 2)

                                '/// Vertical line between this node and child nodes
                                Dim VerticalCenter As Integer = objectBounds.Top + objectCenter
                                If .HasChildren And .Expanded Then
                                    Dim verticalNodeLineLeft As Integer = objectBounds.Left + objectCenter - 1
                                    Dim verticalNodeLineTop_xy As New Point(verticalNodeLineLeft, VerticalCenter)
                                    Dim childLast = .Nodes.Last
                                    Dim verticalNodeLineBottom_xy As New Point(verticalNodeLineLeft, {childLast.Bounds_Text.Top + CInt(childLast.Bounds_Text.Height / 2), ClientRectangle.Height}.Min)
                                    e.Graphics.DrawLine(linePen, verticalNodeLineTop_xy, verticalNodeLineBottom_xy)
                                End If
                                If .HasChildren Then e.Graphics.DrawImage(If(.Expanded, CollapseImage, ExpandImage), .Bounds_ShowHide)

                                Dim horizontalNodeLine_x As Integer = {
                                                .Bounds_Check.Left,
                                                .Bounds_Favorite.Left,
                                                .Bounds_ShowHide.Left
                                                                              }.Min

                                If IsNothing(.Parent) Then
                                    If drawRootLines Then e.Graphics.DrawLine(linePen, New Point(verticalRootLine_x, VerticalCenter), New Point(horizontalNodeLine_x, VerticalCenter))
                                Else
                                    REM /// HORIZONTAL LINES LEFT OF EXPAND/COLLAPSE
                                    Dim NodeHorizontalLeftPoint As New Point(.Parent.Bounds_ShowHide.Left + objectCenter, VerticalCenter)
                                    Dim NodeHorizontalRightPoint As New Point(horizontalNodeLine_x, VerticalCenter)
                                    e.Graphics.DrawLine(linePen, NodeHorizontalLeftPoint, NodeHorizontalRightPoint)
                                End If

                            End With
                        Next
                        '/// Vertical root line between top ( first ) node and bottom ( last ) node ... but don't draw if the top IS the bottom too ( 1 node only )
                        If drawRootLines Then
                            Dim lastNode As Node = Nodes.Last
                            Dim LineTop As Integer = 2 + firstRootNode.Bounds_ShowHide.Top + CInt(firstRootNode.Bounds_ShowHide.Height / 2)
                            Dim TopPoint As New Point(verticalRootLine_x, {0, LineTop}.Max)
                            Dim LineBottom As Integer = lastNode.Bounds_Text.Top + Convert.ToInt32(lastNode.Height / 2)
                            Dim BottomPoint As New Point(verticalRootLine_x, {LineBottom, Height}.Min)
                            e.Graphics.DrawLine(linePen, TopPoint, BottomPoint)
                        End If
                    End Using
                Else
                    HScroll.Hide()
                    VScroll.Hide()
                End If
            End With
            ControlPaint.DrawBorder3D(e.Graphics, ClientRectangle, Border3DStyle.Sunken)
        End If

    End Sub
    Friend Sub RequiresRepaint()

        RR = True
        REM /// RESET INDEX / HEIGHT
        VisibleIndex = 0
        RollingWidth = Offset.X
        RollingHeight = Offset.Y

        REM /// ITERATE ALL NODES CHANGING BOUNDS
        RefreshNodesBounds_Lines(Nodes)

        REM /// TOTAL SIZE + RESIZE THE CONTROL IF AUTOSIZE
#Region " DETERMINE THE MAXIMUM POSSIBLE SIZE OF THE CONTROL AND COMPARE TO THE UNRESTRICTED SIZE "
        Dim screenLocation As Point = PointToScreen(New Point(0, 0))
        Dim wa As Size = WorkingArea.Size
        Dim maxScreenSize As New Size(wa.Width - screenLocation.X, wa.Height - screenLocation.Y)
        Dim maxParentSize As New Size
        Dim maxUserSize As Size = MaximumSize
        Dim unboundedSize As Size = UnRestrictedSize 'This is strictly the space size required to fit all node text ( does NOT include ScrollBars )
#Region " DETERMINE IF A PARENT RESTRICTS THE SIZE OF THE TREEVIEWER - LOOK FOR <.AutoSize> IN PARENT CONTROL PROPERTIES "
        If Parent IsNot Nothing Then
            Dim controlType As Type = Parent.GetType
            Dim properties As Reflection.PropertyInfo() = controlType.GetProperties
            Dim growParent As Boolean = False
            For Each controlProperty In properties
                If controlProperty.Name = "AutoSize" Then
                    Dim propertyValue As Boolean = DirectCast(controlProperty.GetValue(Parent), Boolean)
                    If propertyValue Then growParent = True
                    Exit For
                End If
            Next
            If Not growParent Then maxParentSize = New Size(Parent.Width, Parent.Height)
            If Not Parent.MaximumSize.IsEmpty Then maxParentSize = Parent.MaximumSize
        End If
#End Region
        Dim maxSize As New Size
        Dim sizes As New List(Of Size) From {maxScreenSize, maxParentSize, maxUserSize} ', New Size(IC_NodeEdit.Location.X + IC_NodeEdit.Width, IC_NodeEdit.Height)
        Dim nonZeroWidths As New List(Of Integer)(From s In sizes Where s.Width > 0 Select s.Width)
        Dim nonZeroHeights As New List(Of Integer)(From s In sizes Where s.Height > 0 Select s.Height)
        Dim maxWidth As Integer = nonZeroWidths.Min
        Dim maxHeight As Integer = nonZeroHeights.Min
        If AutoSize Then 'Can resize
            Dim proposedWidth = {UnRestrictedSize.Width, maxWidth}.Min
            Dim proposedHeight = {UnRestrictedSize.Height, maxHeight}.Min

            Dim hScrollVisible As Boolean = UnRestrictedSize.Width > maxWidth
            If hScrollVisible Then proposedHeight = {proposedHeight + HScrollHeight, maxHeight}.Min

            Dim vscrollVisible As Boolean = UnRestrictedSize.Height > maxHeight
            If vscrollVisible Then proposedWidth = {proposedWidth + VScrollWidth, maxWidth}.Min
            Width = proposedWidth
            Height = proposedHeight
            maxWidth = Width
            maxHeight = Height
        Else
            maxWidth = {Width, maxWidth}.Min
            maxHeight = {Height, maxHeight}.Min
        End If
        With HScroll
            .Minimum = 0
            .Maximum = {0, UnRestrictedSize.Width - 1}.Max
            .Visible = UnRestrictedSize.Width > maxWidth
            .Left = 0
            .Width = maxWidth
            .Top = Height - .Height
            .LargeChange = maxWidth
            If .Visible Then
                If .Value > maxWidth - Width Then .Value = {maxWidth - Width, .Minimum}.Max
                .Show()
            Else
                .Value = 0
                .Hide()
            End If
        End With
        With VScroll
            .Minimum = 0
            .Maximum = {0, UnRestrictedSize.Height - 1}.Max
            .Visible = UnRestrictedSize.Height > maxHeight
            .Left = maxWidth - VScrollWidth
            .Height = maxHeight
            .Top = 0
            .LargeChange = maxHeight
            If .Visible Then
                If .Value > maxHeight - Height Then .Value = {maxHeight - Height, .Minimum}.Max
                .Show()
            Else
                .Value = 0
                .Hide()
            End If
        End With
        If Nodes.Any And StopMe Then Stop
#End Region
        UpdateDrawNodes()

        REM /// FINALLY- REPAINT
        Invalidate()
        RR = False

    End Sub
    Private Sub RefreshNodesBounds_Lines(Nodes As NodeCollection)

        Dim NodeIndex As Integer = 0
        For Each Node As Node In Nodes
            With Node
                ._Index = NodeIndex
                ._Visible = .Parent Is Nothing OrElse .Parent.Expanded
                If .Visible Then
                    RefreshNodeBounds_Lines(Node, True)
                    ._VisibleIndex = VisibleIndex
                    VisibleIndex += 1
                    If .Bounds_Text.Right > RollingWidth Then RollingWidth = .Bounds_Text.Right
                    RollingHeight += .Height
                    If .HasChildren Then RefreshNodesBounds_Lines(.Nodes)
                End If
                NodeIndex += 1
            End With
            If FavoritesFirst Then Node.Nodes.Sort(Function(x, y)
                                                       Dim level1 As Integer = y.Favorite.CompareTo(x.Favorite)
                                                       If level1 = 0 Then
                                                           Return x.Index.CompareTo(y.Index)  'String.Compare(x.Text, y.Text, StringComparison.Ordinal)
                                                       Else
                                                           Return level1
                                                       End If
                                                   End Function)
        Next

    End Sub
    Private Sub RefreshNodeBounds_Lines(Node As Node, Optional ExpandBeforeText As Boolean = True)

        Dim y As Integer = RollingHeight - VScroll.Value
        Const HorizontalSpacing As Integer = 3

        With Node
#Region " S E T   B O U N D S "
            Dim leftMost As Integer = Offset.X + HorizontalSpacing + If(.Parent Is Nothing, If(RootLines, 6, 0), .Parent.Bounds_ShowHide.Right) - HScroll.Value

            If ExpandBeforeText Then
                '■■■■■■■■■■■■■ P r e f e r
#Region " +- Icon precedes Text "
                'Bounds.X cascades from [1] Favorite, [2] Checkbox, [3] Image, [4] ShowHide, [5] Text
                REM FAVORITE
                ._Bounds_Favorite.X = leftMost
                ._Bounds_Favorite.Y = y + CInt((.Height - FavoriteImage.Height) / 2)
                ._Bounds_Favorite.Width = If(.CanFavorite, FavoriteImage.Width, 0)
                ._Bounds_Favorite.Height = If(.CanFavorite, FavoriteImage.Height, .Height)

                REM CHECKBOX
                ._Bounds_Check.X = ._Bounds_Favorite.Right + If(._Bounds_Favorite.Width = 0, 0, HorizontalSpacing)
                If CheckboxStyle = CheckStyle.Slide Then
                    ._Bounds_Check.Width = If(.CheckBox, My.Resources.slideStateOff.Width, 0)
                    ._Bounds_Check.Height = My.Resources.slideStateOff.Height
                    ._Bounds_Check.Y = y + CInt((.Height - ._Bounds_Check.Height) / 2)
                Else
                    ._Bounds_Check.Width = If(.CheckBox, CheckHeight, 0)
                    ._Bounds_Check.Height = CheckHeight
                    ._Bounds_Check.Y = y + CInt((.Height - CheckHeight) / 2)
                End If

                REM IMAGE
                ._Bounds_Image.X = ._Bounds_Check.Right + If(._Bounds_Check.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Image.Height = If(IsNothing(.Image), 0, If(.ImageScaling, .Height, .Image.Height))
                'MAKE IMAGE SQUARE IF SCALING
                ._Bounds_Image.Width = If(IsNothing(.Image), 0, If(.ImageScaling, ._Bounds_Image.Height, .Image.Width))
                ._Bounds_Image.Y = y + CInt((.Height - ._Bounds_Image.Height) / 2)

                REM EXPAND/COLLAPSE
                ._Bounds_ShowHide.X = ._Bounds_Image.Right + HorizontalSpacing '+ If(IsNothing(.Parent), If(RootLines, 6, 0), .Parent._Bounds_Image.Right + HorizontalSpacing)
                ._Bounds_ShowHide.Y = y + CInt((.Height - ExpandHeight) / 2)
                ._Bounds_ShowHide.Width = If(.HasChildren, ExpandHeight, 0)
                ._Bounds_ShowHide.Height = ExpandHeight

                REM TEXT
                ._Bounds_Text.X = ._Bounds_ShowHide.Right + If(._Bounds_ShowHide.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Text.Y = y
                ._Bounds_Text.Height = .Height
                '._Bounds_Text.Width = .TextWidth 'Not necessary- set when the Text changes at the Node level
#End Region
            Else
#Region " +- Icon follows Text "
                REM EXPAND/COLLAPSE
                ._Bounds_ShowHide.X = leftMost
                ._Bounds_ShowHide.Y = y + CInt((.Height - ExpandHeight) / 2)
                ._Bounds_ShowHide.Width = If(.HasChildren, ExpandHeight, 0)
                ._Bounds_ShowHide.Height = ExpandHeight

                REM FAVORITE
                ._Bounds_Favorite.X = ._Bounds_ShowHide.Right + If(._Bounds_ShowHide.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Favorite.Y = y + CInt((.Height - FavoriteImage.Height) / 2)
                ._Bounds_Favorite.Width = If(.CanFavorite, FavoriteImage.Width, 0)
                ._Bounds_Favorite.Height = If(.CanFavorite, FavoriteImage.Height, .Height)

                REM CHECKBOX
                ._Bounds_Check.X = ._Bounds_Favorite.Right + If(._Bounds_Favorite.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Check.Width = If(.CheckBox, CheckHeight, 0)
                ._Bounds_Check.Height = CheckHeight
                ._Bounds_Check.Y = y + CInt((.Height - ._Bounds_Check.Height) / 2)

                REM IMAGE
                ._Bounds_Image.X = ._Bounds_Check.Right + If(._Bounds_Check.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Image.Height = If(IsNothing(.Image), 0, If(.ImageScaling, .Height, .Image.Height))
                'MAKE IMAGE SQUARE IF SCALING
                ._Bounds_Image.Width = If(IsNothing(.Image), 0, If(.ImageScaling, ._Bounds_Image.Height, .Image.Width))
                ._Bounds_Image.Y = y + CInt((.Height - ._Bounds_Image.Height) / 2)

                REM TEXT
                ._Bounds_Text.X = ._Bounds_Image.Right + If(._Bounds_Image.Width = 0, 0, HorizontalSpacing)
                ._Bounds_Text.Y = y
                ._Bounds_Text.Height = .Height
#End Region
            End If
#End Region
        End With

    End Sub
#Region " SCROLLING, SCROLLING, SCROLLING "
    Private Sub OnScrolled(sender As Object, e As ScrollEventArgs) Handles HScroll.Scroll, VScroll.Scroll

        If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
            HScrollLeftRight(e.OldValue - e.NewValue)

        ElseIf e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            VScrollUpDown(e.OldValue - e.NewValue)

        End If
        UpdateDrawNodes()

    End Sub
    Private Sub HScrollLeftRight(changeX As Integer)

        If changeX = 0 Then Exit Sub
        IC_NodeEdit.Location = New Point(IC_NodeEdit.Location.X + changeX, IC_NodeEdit.Location.Y)
        Nodes.Visible.ForEach(Sub(node)
                                  node._Bounds_ShowHide.X += changeX
                                  node._Bounds_Favorite.X += changeX
                                  node._Bounds_Check.X += changeX
                                  node._Bounds_Image.X += changeX
                                  node._Bounds_Text.X += changeX
                              End Sub)
        Invalidate()

    End Sub
    Private Sub VScrollUpDown(changeY As Integer)

        If changeY = 0 Then Exit Sub
        IC_NodeEdit.Location = New Point(IC_NodeEdit.Location.X, IC_NodeEdit.Location.Y + changeY)
        Nodes.Visible.ForEach(Sub(node)
                                  node._Bounds_ShowHide.Y += changeY
                                  node._Bounds_Favorite.Y += changeY
                                  node._Bounds_Check.Y += changeY
                                  node._Bounds_Image.Y += changeY
                                  node._Bounds_Text.Y += changeY
                              End Sub)
        Invalidate()

    End Sub
#End Region
    Protected Overrides Sub OnFontChanged(e As EventArgs)

        For Each node In Nodes.All
            node.Font = Font
        Next
        RequiresRepaint()
        MyBase.OnFontChanged(e)
    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)

        If Not RR Then RequiresRepaint()
        Invalidate()
        MyBase.OnSizeChanged(e)

    End Sub
    Friend Sub NodeTimer_Start(TickNode As Node)

        With NodeTimer
            .Tag = TickNode
            .Start()
        End With

    End Sub
    Private Sub NodeTimer_Tick() Handles NodeTimer.Tick

        With NodeTimer
            .Stop()
            Dim TickNode As Node = DirectCast(.Tag, Node)
            RaiseEvent NodesChanged(Me, New NodeEventArgs(TickNode))
            'If Name = "Scripts" Then Stop
            RequiresRepaint()
        End With

    End Sub
#End Region
#Region " METHODS / FUNCTIONS "
    Public Function HitTest(Location As Point) As HitRegion

        Dim Region As New HitRegion
        Dim expandBounds As New List(Of Node)(From N In Nodes.Draw Where N.Bounds_ShowHide.Contains(Location))
        Dim Bounds_Favorite As New List(Of Node)(From N In Nodes.Draw Where N.Bounds_Favorite.Contains(Location))
        Dim Bounds_Check As New List(Of Node)(From N In Nodes.Draw Where N.Bounds_Check.Contains(Location))
        Dim Bounds_Image As New List(Of Node)(From N In Nodes.Draw Where N.Bounds_Image.Contains(Location))
        Dim nodeBounds As New List(Of Node)(From N In Nodes.Draw Where N.Bounds_Text.Contains(Location))

        Dim HitBounds As New List(Of Node)(expandBounds.Union(Bounds_Favorite).Union(Bounds_Check).Union(nodeBounds))
        If HitBounds.Any Then
            With Region
                .Node = HitBounds.First
                If expandBounds.Any Then .Region = MouseRegion.Expander
                If Bounds_Favorite.Any Then .Region = MouseRegion.Favorite
                If Bounds_Check.Any Then .Region = MouseRegion.CheckBox
                If Bounds_Image.Any Then .Region = MouseRegion.Image
                If nodeBounds.Any Then .Region = MouseRegion.Node
            End With
        End If
        Return Region

    End Function
    Public Function HitTest(dragBounds As Rectangle) As HitRegion

        Dim Region As New HitRegion
        Dim expandBounds As New List(Of Node)(From N In Nodes.Draw Where dragBounds.Contains(N.Bounds_ShowHide))
        Dim Bounds_Favorite As New List(Of Node)(From N In Nodes.Draw Where dragBounds.Contains(N.Bounds_Favorite))
        Dim Bounds_Check As New List(Of Node)(From N In Nodes.Draw Where dragBounds.Contains(N.Bounds_Check))
        Dim Bounds_Image As New List(Of Node)(From N In Nodes.Draw Where dragBounds.Contains(N.Bounds_Image))
        Dim nodeBounds As New List(Of Node)(From N In Nodes.Draw Where dragBounds.Contains(N.Bounds_Text))

        Dim HitBounds As New List(Of Node)(expandBounds.Union(Bounds_Favorite).Union(Bounds_Check).Union(nodeBounds))
        If HitBounds.Any Then
            With Region
                .Node = HitBounds.First
                If expandBounds.Any Then .Region = MouseRegion.Expander
                If Bounds_Favorite.Any Then .Region = MouseRegion.Favorite
                If Bounds_Check.Any Then .Region = MouseRegion.CheckBox
                If Bounds_Image.Any Then .Region = MouseRegion.Image
                If nodeBounds.Any Then .Region = MouseRegion.Node
            End With
        End If
        Return Region

    End Function
    Public Sub ExpandNodes()

        If Nodes.Any Then
            ExpandCollapseNodes(Nodes, True)
            Nodes.First.Expand()
        End If
        VScroll.Value = 0
        HScroll.Value = 0
        RequiresRepaint()

    End Sub
    Public Sub CollapseNodes()

        If Nodes.Any Then
            ExpandCollapseNodes(Nodes, False)
            Nodes.First.Collapse()
        End If
        VScroll.Value = 0
        HScroll.Value = 0
        RequiresRepaint()

    End Sub
    Private Sub ExpandCollapseNodes(Nodes As NodeCollection, State As Boolean)

        For Each Node As Node In Nodes
            If State Then
                Node.Expand()
            Else
                Node.Collapse()
            End If
            If Node.HasChildren Then ExpandCollapseNodes(Node.Nodes, State)
        Next

    End Sub
    Private Sub UpdateDrawNodes()

        Nodes.Client.Clear()
        For Each Node In Nodes.Visible
            If ClientRectangle.Contains(New Point(1, Node.Bounds_Text.Top)) Or ClientRectangle.Contains(New Point(1, Node.Bounds_Text.Bottom)) Then
                Nodes.Client.Add(Node)
            End If
        Next

        Nodes.Draw.Clear()
        Nodes.Draw.AddRange(Nodes.Client)
        Dim NodesDraw = (From N In Nodes.Visible Group N By N.Parent Into Group Select Parent, Children =
                ((From C In Group Where C._Bounds_Text.Bottom < 0 Order By C._Bounds_Text.Y Descending).Take(1)).Union _
                ((From C In Group Where C._Bounds_Text.Top > ClientRectangle.Height Order By C._Bounds_Text.Y Ascending).Take(1)))
        For Each ParentGroup In NodesDraw
            Nodes.Draw.AddRange(ParentGroup.Children)
        Next

    End Sub
    Public Sub AutoWidth()

        If Nodes.Any Then
            Width = Nodes.Max(Function(n) {n.Bounds_Text.Right, If(IC_NodeEdit.Visible, IC_NodeEdit.Location.X + IC_NodeEdit.Width, 0)}.Max)
            RequiresRepaint()
        End If

    End Sub
#End Region
    Public Overrides Function ToString() As String
        Return $"{If(Name, "Treeviewer")} [{Nodes.Count}]"
    End Function
End Class

REM //////////////// NODE COLLECTION \\\\\\\\\\\\\\\\
Public NotInheritable Class NodeCollection
    Inherits List(Of Node)
    Private SortOrder_ As SortOrder
    Public Property SortOrder As SortOrder
        Get
            Return SortOrder_
        End Get
        Set(value As SortOrder)
            If Not SortOrder_ = value Then
                SortOrder_ = value
                If value = SortOrder.None Then SortOriginalIndex()
                If value = SortOrder.Ascending Then SortAscending()
                If value = SortOrder.Descending Then SortDescending()
            End If
        End Set
    End Property
    Friend Sub New(TreeViewer As TreeViewer)
        _TreeViewer = TreeViewer
    End Sub
    Private _TreeViewer As TreeViewer
    Public ReadOnly Property TreeViewer As TreeViewer
        Get
            Return _TreeViewer
        End Get
    End Property
    Friend _Parent As Node
    Public ReadOnly Property Parent As Node
        Get
            Return _Parent
        End Get
    End Property
    Public ReadOnly Property All As List(Of Node)
        Get
            Dim Nodes As New List(Of Node)
            Dim Queue As New Queue(Of Node)
            Dim TopNode As Node, Node As Node
            For Each TopNode In Me
                Queue.Enqueue(TopNode)
            Next
            While Queue.Any
                TopNode = Queue.Dequeue
                Nodes.Add(TopNode)
                For Each Node In TopNode.Nodes
                    Queue.Enqueue(Node)
                Next
            End While
            Nodes.Sort(Function(x, y) x.VisibleIndex.CompareTo(y.VisibleIndex))
            Return Nodes
        End Get
    End Property
    Public ReadOnly Property Roots As List(Of Node)
        Get
            Return All.Where(Function(x) x.Level = 0).ToList
        End Get
    End Property
    Public ReadOnly Property Level(LevelIndex As Integer) As List(Of Node)
        Get
            Return All.Where(Function(x) x.Level = LevelIndex).ToList
        End Get
    End Property
    Public ReadOnly Property Visible As List(Of Node)
        Get
            Return All.Where(Function(m) m.Visible).ToList
        End Get
    End Property
    Public ReadOnly Property Parents As List(Of Node)
        Get
            Return All.Where(Function(m) m.Visible And m.HasChildren).ToList
        End Get
    End Property
    Public ReadOnly Property Client As New List(Of Node)
    Public ReadOnly Property Draw As New List(Of Node)
    Public ReadOnly Property CollectionDataType As Type
        Get
            Dim Types = From n In Me Select n.DataType
            Return GetDataType(Types)
        End Get
    End Property

#Region " Methods "
    Public Overloads Function Contains(Name As String) As Boolean

        If Count = 0 Then
            Return False
        Else
            Dim Found As Boolean = False
            For Each Node In All
                If Node.Name = Name Then
                    Found = True
                    Exit For
                End If
            Next
            Return Found
        End If

    End Function
    Public Overloads Function Add(Name As String, Text As String, Checked As Boolean, Image As Image) As Node
        Return Add(New Node With {.Name = Name, .Text = Text, .Checked = Checked, .Image = Image})
    End Function
    Public Overloads Function Add(Name As String, Text As String, Checked As Boolean) As Node
        Return Add(New Node With {.Name = Name, .Text = Text, .Checked = Checked})
    End Function
    Public Overloads Function Add(Name As String, Text As String, Image As Image) As Node
        Return Add(New Node With {.Name = Name, .Text = Text, .Image = Image})
    End Function
    Public Overloads Function Add(Name As String, Text As String) As Node
        Return Add(New Node With {.Name = Name, .Text = Text})
    End Function
    Public Overloads Function Add(Text As String, Checked As Boolean, Image As Image) As Node
        Return Add(New Node With {.Text = Text, .Checked = Checked, .Image = Image})
    End Function
    Public Overloads Function Add(Text As String, Checked As Boolean) As Node
        Return Add(New Node With {.Text = Text, .Checked = Checked})
    End Function
    Public Overloads Function Add(Text As String, Image As Image) As Node
        Return Add(New Node With {.Text = Text, .Image = Image})
    End Function
    Public Overloads Function Add(Text As String) As Node
        Return Add(New Node With {.Text = Text})
    End Function
    Public Overloads Function Add(AddNode As Node) As Node

        If AddNode IsNot Nothing Then
            With AddNode
                ._Index = Count
                .AddTime = Now
                If Parent Is Nothing Then   ' *** ROOT NODE
                    'mTreeViewer was set when TreeViewer was created with New NodeCollection
                    ._TreeViewer = TreeViewer
                    ._Visible = True
                    ._Level = 0

                Else                        ' *** CHILD NODE
                    REM /// Get TreeViewer value from Parent. A Node Collection shares some Node properties
                    _TreeViewer = Parent.TreeViewer
                    ._TreeViewer = Parent.TreeViewer
                    ._Parent = Parent
                    ._Level = Parent.Level + 1
                    If Parent.Expanded Then
                        ._Visible = True
                        If Count = 0 Then
                            ._VisibleIndex = 0
                        Else
                            ._VisibleIndex = Last.VisibleIndex + 1
                        End If
                    Else
                        ._VisibleIndex = -1
                    End If

                End If
                If TreeViewer IsNot Nothing Then
                    .Font = TreeViewer.Font
                    .TextWidth_Set()
                    TreeViewer.NodeTimer_Start(AddNode)
                End If
            End With
            MyBase.Add(AddNode)
        End If
        Return AddNode

    End Function
    Public Overloads Function AddRange(Nodes As List(Of Node)) As List(Of Node)

        If Nodes IsNot Nothing Then
            For Each Node As Node In Nodes
                Add(Node)
            Next
            If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        End If
        Return Nodes

    End Function
    Public Overloads Function AddRange(Nodes As Node()) As Node()

        If Nodes IsNot Nothing Then
            For Each Node As Node In Nodes
                Add(Node)
            Next
            If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        End If
        Return Nodes

    End Function
    Public Overloads Function AddRange(Nodes As String()) As Node()

        Dim NewNodes As New List(Of Node)
        If Nodes IsNot Nothing Then
            For Each NewNode As String In Nodes
                Add(NewNode)
                NewNodes.Add(New Node With {.Text = NewNode})
            Next
            If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        End If
        Return NewNodes.ToArray

    End Function
    Public Overloads Function Clear(Nodes As NodeCollection) As NodeCollection

        Clear()
        If Not IsNothing(TreeViewer) Then TreeViewer.RequiresRepaint()
        Return Nodes

    End Function
    Public Overloads Function Insert(Index As Integer, InsertNode As Node) As Node

        If InsertNode IsNot Nothing Then
            If IsNothing(TreeViewer) Then
                MyBase.Insert(Index, InsertNode)
            Else
                If TreeViewer.Nodes.All.Contains(InsertNode) Then
                    'Throw New ArgumentException("This node already exists In the Treeviewer. Try Removing the Node")

                Else
                    MyBase.Insert(Index, InsertNode)
                    InsertNode._TreeViewer = TreeViewer
                    InsertNode._Parent = Parent
                    TreeViewer.NodeTimer_Start(InsertNode)
                End If
            End If
        End If
        Return InsertNode

    End Function
    Public Overloads Function Remove(RemoveNode As Node) As Node

        If RemoveNode IsNot Nothing Then
            MyBase.Remove(RemoveNode)
            TreeViewer.NodeTimer_Start(RemoveNode)
        End If
        Return RemoveNode

    End Function
    Public Shadows Function Item(Name As String) As Node
        Dim Nodes As New List(Of Node)((From N In Me Where N.Name = Name).ToArray)
        Return If(Nodes.Any, Nodes.First, Nothing)
    End Function
    Public Shadows Function ItemByTag(TagObject As Object) As Node
        Dim Nodes As New List(Of Node)((From N In All Where N.Tag Is TagObject).ToArray)
        Return If(Nodes.Any, Nodes.First, Nothing)
    End Function
    Friend Sub SortOriginalIndex(Optional repaint As Boolean = True)

        If TreeViewer?.FavoritesFirst Then
            Sort(Function(x, y)
                     Dim Level1 = y.Favorite.CompareTo(x.Favorite) 'False=0, True=1 
                     If Level1 = 0 Then
                         Dim Level2 As Integer = x.AddTime.CompareTo(y.AddTime)
                         Return Level2
                     Else
                         Return Level1
                     End If
                 End Function)
        Else
            Sort(Function(x, y)
                     Return x.AddTime.CompareTo(y.AddTime)
                 End Function)
        End If
        'If repaint Then TreeViewer?.RequiresRepaint()

    End Sub
    Friend Sub SortAscending(Optional repaint As Boolean = True)

        If TreeViewer?.FavoritesFirst Then
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite) 'False=0, True=1 
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = String.Compare(x.Text, y.Text, StringComparison.InvariantCulture)
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Boolean)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Date)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(x, y)
                             Dim Level1 = y.Favorite.CompareTo(x.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)
            End Select
        Else
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(x, y) String.Compare(x.Text, y.Text, StringComparison.Ordinal))

                Case GetType(Boolean)
                    Sort(Function(x, y) Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture)))

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(x, y) Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture)))

                Case GetType(Date)
                    Sort(Function(x, y) Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture)))

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(x, y) Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture)))

            End Select
        End If
        'If repaint Then TreeViewer?.RequiresRepaint()

    End Sub
    Friend Sub SortDescending(Optional repaint As Boolean = True)

        If TreeViewer?.FavoritesFirst Then
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = String.Compare(x.Text, y.Text, StringComparison.InvariantCulture)
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Boolean)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Date)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(y, x)
                             Dim Level1 = x.Favorite.CompareTo(y.Favorite)
                             If Level1 <> 0 Then
                                 Return Level1
                             Else
                                 Dim Level2 = Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture))
                                 Return Level2
                             End If
                         End Function)
            End Select
        Else
            Select Case CollectionDataType
                Case GetType(String)
                    Sort(Function(y, x) String.Compare(x.Text, y.Text, StringComparison.Ordinal))

                Case GetType(Boolean)
                    Sort(Function(y, x) Convert.ToBoolean(x.Text, InvariantCulture).CompareTo(Convert.ToBoolean(y.Text, InvariantCulture)))

                Case GetType(Decimal), GetType(Double)
                    Sort(Function(y, x) Convert.ToDecimal(x.Text, InvariantCulture).CompareTo(Convert.ToDecimal(y.Text, InvariantCulture)))

                Case GetType(Date)
                    Sort(Function(y, x) Convert.ToDateTime(x.Text, InvariantCulture).CompareTo(Convert.ToDateTime(y.Text, InvariantCulture)))

                Case GetType(Long), GetType(Integer), GetType(Short), GetType(Byte)
                    Sort(Function(y, x) Convert.ToInt64(x.Text, InvariantCulture).CompareTo(Convert.ToInt64(y.Text, InvariantCulture)))

            End Select
        End If
        'If repaint Then TreeViewer?.RequiresRepaint()

    End Sub
#End Region
End Class

REM //////////////// NODE \\\\\\\\\\\\\\\\
Public Class Node
    Implements IDisposable
    Public Sub New()
        Nodes = New NodeCollection(_TreeViewer) With {._Parent = Me}
    End Sub
    Public Enum SeparatorPosition
        None
        Above
        Below
    End Enum
#Region " Properties And Fields "
    Friend _TreeViewer As TreeViewer
    Public ReadOnly Property TreeViewer As TreeViewer
        Get
            Return _TreeViewer
        End Get
    End Property
    Friend _Parent As Node
    Public ReadOnly Property Parent As Node
        Get
            Return _Parent
        End Get
    End Property
    Public ReadOnly Property Nodes As NodeCollection
    Public Property CursorGlowColor As Color = Color.LimeGreen
    Public ReadOnly Property HasChildren As Boolean
        Get
            Return Nodes.Any
        End Get
    End Property
    Friend _Expanded As Boolean
    Public ReadOnly Property Expanded As Boolean
        Get
            Return _Expanded
        End Get
    End Property
    Public ReadOnly Property Collapsed As Boolean
        Get
            Return Not Expanded
        End Get
    End Property
    Private _CheckBox As Boolean = False
    Public Property CheckBox As Boolean
        Get
            If TreeViewer Is Nothing Then
                Return _CheckBox
            Else
                If TreeViewer.CheckBoxes = TreeViewer.CheckState.All Then
                    _CheckBox = True
                ElseIf TreeViewer.CheckBoxes = TreeViewer.CheckState.None Then
                    _CheckBox = False
                ElseIf TreeViewer.CheckBoxes = TreeViewer.CheckState.Mixed Then

                End If
                Return _CheckBox
            End If
        End Get
        Set(value As Boolean)
            If value <> _CheckBox Then
                _CheckBox = value
                If TreeViewer Is Nothing Then
                Else
                    If value Then
                        If TreeViewer.CheckBoxes = TreeViewer.CheckState.None Then
                            TreeViewer.CheckBoxes = TreeViewer.CheckState.Mixed
                        End If
                    End If
                End If
                RequiresRepaint()
            End If
        End Set
    End Property
    Private _Checked As Boolean = False
    Public Property Checked As Boolean
        Get
            If Children.Any Then
                _Checked = (From C In Children Where C.Checked).Count = Children.Count
            End If
            Return _Checked
        End Get
        Set(value As Boolean)
            If Not value = _Checked Then
                If value Then CheckBox = True
                _Checked = value
                For Each Child In Children
                    Child.Checked = value
                Next
                RequiresRepaint()
            End If
        End Set
    End Property
    Public ReadOnly Property PartialChecked As Boolean
        Get
            Dim _Children = Children
            Dim ChildrenCheckCount As Integer = (From C In _Children Where C.Checked).Count
            Return ChildrenCheckCount > 0 And ChildrenCheckCount < _Children.Count
        End Get
    End Property
    Private _Image As Image
    Public Property Image As Image
        Get
            Return _Image
        End Get
        Set(value As Image)
            _Image = value
            RequiresRepaint()
        End Set
    End Property
    Public Property CanEdit As Boolean = True
    Public Property CanAdd As Boolean = True
    Public Property CanRemove As Boolean = True
    Private _ImageScaling As Boolean = False
    Public Property ImageScaling As Boolean
        Get
            Return _ImageScaling
        End Get
        Set(value As Boolean)
            _ImageScaling = value
            RequiresRepaint()
        End Set
    End Property
    Public ReadOnly Property Root As Node
        Get
            Dim RootNode As Node = Me
            Do While Not IsNothing(RootNode.Parent)
                RootNode = RootNode.Parent
            Loop
            Return RootNode
        End Get
    End Property
    Public ReadOnly Property IsRoot As Boolean
        Get
            Return Parent Is Nothing
        End Get
    End Property
    Public Property Name As String
    Public Property Tag As Object
    Friend Sub TextWidth_Set()

        'Affected by: 1) Node added ( Treeviewer Property Set ) 2) Text changed 3) Font changed
        If TreeViewer Is Nothing Then
            _Bounds_Text.Width = MeasureText(Text, Font).Width

        ElseIf If(text, String.Empty).Any Then
            Using g As Graphics = TreeViewer.CreateGraphics
                Dim characterRanges As CharacterRange() = {New CharacterRange(0, Text.Length), New CharacterRange(0, 0)}
                Dim width As Single = 1000.0F
                Dim height As Single = 36.0F
                Dim layoutRect As RectangleF = New RectangleF(0.0F, 0.0F, width, height)
                Using sf As StringFormat = New StringFormat With {
                    .FormatFlags = StringFormatFlags.NoWrap,
                    .Alignment = StringAlignment.Near,
                    .LineAlignment = StringAlignment.Center
                    }
                    sf.SetMeasurableCharacterRanges(characterRanges)
                    g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
                    Dim stringRegions As Region() = g.MeasureCharacterRanges(Text, Font, layoutRect, sf)
                    Dim measureRect1 As RectangleF = stringRegions(0).GetBounds(g)
                    Dim textTangle As Rectangle = Rectangle.Round(measureRect1)
                    _Bounds_Text.Width = textTangle.Width
                End Using
            End Using
        Else
            _Bounds_Text.Width = 0
        End If
        RequiresRepaint()

    End Sub 'Only changes to Font & Text affect the width
    Private _Text As String
    Public Property Text As String
        Get
            Return $"{_Text}{If(ShowIndex, $" [{Index}]", String.Empty)}"
        End Get
        Set(value As String)
            If Not _Text = value Then
                _Text = Replace(value, "&", "&&")
                TextWidth_Set()
            End If
        End Set
    End Property
    Public Property TipText As String
    Public ReadOnly Property Options As New List(Of Object)
    Public ReadOnly Property ChildOptions As New List(Of Object)
    Private _Font As New Font("Calibri", 9)
    Public Property Font As Font
        Get
            Return _Font
        End Get
        Set(value As Font)
            _Font = value
            RequiresRepaint()
        End Set
    End Property
    Private _ForeColor As Color = Color.Black
    Public Property ForeColor As Color
        Get
            Return _ForeColor
        End Get
        Set(value As Color)
            _ForeColor = value
            RequiresRepaint()
        End Set
    End Property
    Private _TextBackColor As Color = Color.Transparent
    Public Property TextBackColor As Color
        Get
            Return _TextBackColor
        End Get
        Set(value As Color)
            _TextBackColor = value
            RequiresRepaint()
        End Set
    End Property
    Private _BackColor As Color = Color.Transparent
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(value As Color)
            _BackColor = value
            RequiresRepaint()
        End Set
    End Property
    Friend _Selected As Boolean
    Public ReadOnly Property Selected As Boolean
        Get
            Return _Selected
        End Get
    End Property
    Friend _Index As Integer
    Public ReadOnly Property Index As Integer
        Get
            Return _Index
        End Get
    End Property
    Friend AddTime As Date
    Public ReadOnly Property Height As Integer
        Get
            Dim ImageHeight As Integer = 0
            If Not IsNothing(Image) And Not ImageScaling Then
                ImageHeight = Image.Height
            End If
            Return If(Separator = SeparatorPosition.None, 0, 1) + Convert.ToInt32({1 + Font.GetHeight + 1, ImageHeight}.Max)
        End Get
    End Property
    Friend _Bounds_ShowHide As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_ShowHide As Rectangle
        Get
            Return _Bounds_ShowHide
        End Get
    End Property
    Friend _Bounds_Favorite As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_Favorite As Rectangle
        Get
            Return _Bounds_Favorite
        End Get
    End Property
    Friend _Bounds_Check As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_Check As Rectangle
        Get
            Return _Bounds_Check
        End Get
    End Property
    Friend _Bounds_Image As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_Image As Rectangle
        Get
            Return _Bounds_Image
        End Get
    End Property
    Friend _Bounds_Text As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property Bounds_Text As Rectangle
        Get
            Return _Bounds_Text
        End Get
    End Property
    Friend _Level As Integer
    Public ReadOnly Property Level As Integer
        Get
            Return _Level
        End Get
    End Property
    Private ReadOnly _Path As New List(Of KeyValuePair(Of Integer, String))
    Public ReadOnly Property NamePath() As String
        Get
            _Path.Clear()
            GetNamePath(Me)
            Return Join(_Path.OrderBy(Function(x) x.Key).Select(Function(y) y.Value).ToArray, "±")
        End Get
    End Property
    Private Sub GetNamePath(_Node As Node)
        _Path.Add(New KeyValuePair(Of Integer, String)(_Node.Level, _Node.Name))
        If Not IsNothing(_Node.Parent) Then GetNamePath(_Node.Parent)
    End Sub
    Public ReadOnly Property TextPath() As String
        Get
            _Path.Clear()
            GetTextPath(Me)
            Return Join(_Path.OrderBy(Function(x) x.Key).Select(Function(y) y.Value).ToArray, "±")
        End Get
    End Property
    Private Sub GetTextPath(_Node As Node)
        _Path.Add(New KeyValuePair(Of Integer, String)(_Node.Level, _Node.Text))
        If Not IsNothing(_Node.Parent) Then GetTextPath(_Node.Parent)
    End Sub
    Public ReadOnly Property Parents As List(Of Node)
        Get
            Dim _Parents As New List(Of Node)
            Dim ParentNode As Node = Parent
            Do While ParentNode IsNot Nothing
                _Parents.Add(ParentNode)
                ParentNode = ParentNode.Parent
            Loop
            Return _Parents
        End Get
    End Property
    Private Sub GetChildren(_Node As Node)

        For Each Child In _Node.Nodes
            _Children.Add(Child)
            If Child.HasChildren Then GetChildren(Child)
        Next

    End Sub
    Private ReadOnly _Children As New List(Of Node)
    Public ReadOnly Property Children As List(Of Node)
        Get
            _Children.Clear()
            GetChildren(Me)
            Return _Children
        End Get
    End Property
    Private _Separator As SeparatorPosition = SeparatorPosition.None
    Public Property Separator As SeparatorPosition
        Get
            Return _Separator
        End Get
        Set(value As SeparatorPosition)
            If _Separator <> value Then
                _Separator = value
                RequiresRepaint()
            End If
        End Set
    End Property
    Private ShowIndex_ As Boolean = False
    Public Property ShowIndex As Boolean
        Get
            Return ShowIndex_
        End Get
        Set(value As Boolean)
            ShowIndex_ = value
            RequiresRepaint()
        End Set
    End Property
    Friend _Visible As Boolean = False
    Public ReadOnly Property Visible As Boolean
        Get
            Return _Visible
        End Get
    End Property
    Friend _VisibleIndex As Integer
    Public ReadOnly Property VisibleIndex As Integer
        Get
            Return _VisibleIndex
        End Get
    End Property
    Public ReadOnly Property Siblings As List(Of Node)
        Get
            Return New List(Of Node)(From n In If(IsRoot, TreeViewer?.Nodes, Parent?.Nodes) Where n IsNot Me)
        End Get
    End Property
    Public ReadOnly Property FirstSibling As Node
        Get
            If Siblings.Any Then
                Return Siblings.First
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property IsFirstSibling As Boolean
        Get
            Return FirstSibling Is Me
        End Get
    End Property
    Public ReadOnly Property NextSibling As Node
        Get
            If Siblings.Any Then
                If Me Is LastSibling Then
                    Return Nothing
                Else
                    Return Siblings(Index + 1)
                End If
            Else
                Return Nothing
            End If

        End Get
    End Property
    Public ReadOnly Property LastSibling As Node
        Get
            If Siblings.Any Then
                Return Siblings.Last
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property IsLastSibling As Boolean
        Get
            Return LastSibling Is Me
        End Get
    End Property
    Public ReadOnly Property DataType As Type
        Get
            Return GetDataType(SortValue)
        End Get
    End Property
    Public Property CanDragDrop As Boolean = True
    Public Property CanFavorite As Boolean = False
    Private Favorite_ As Boolean = False
    Public Property Favorite As Boolean
        Get
            Return Favorite_
        End Get
        Set(value As Boolean)
            If value <> Favorite_ Then
                Favorite_ = value
                If TreeViewer?.FavoritesFirst Then Parent?.Nodes.SortAscending()
                RequiresRepaint()
            End If
        End Set
    End Property
    Private SortValue_ As String = String.Empty
    Public Property SortValue As String
        Get
            If Not If(SortValue_, String.Empty).Any Then SortValue_ = Text
            Return SortValue_
        End Get
        Set(value As String)
            If value <> SortValue_ Then
                SortValue_ = value
                'Refresh?
            End If
        End Set
    End Property
#End Region
#Region " Methods "
    Friend _Clicked As Boolean
    Private Sub RequiresRepaint()

        If Not IsNothing(TreeViewer) Then
            If _Clicked Then
                _Clicked = False
                TreeViewer.RequiresRepaint()
            Else
                TreeViewer.NodeTimer_Start(Me)
            End If
        End If

    End Sub
    Public Sub Expand()

        If Nodes.Any Then
            _Expanded = True
            ShowHide(Nodes, True)
            RequiresRepaint()
        End If

    End Sub
    Public Sub Collapse()

        If Nodes.Any Then
            _Expanded = False
            ShowHide(Nodes, False)
            RequiresRepaint()
        End If

    End Sub
    Public Sub Click()
        _Selected = True
        RequiresRepaint()
    End Sub
    Public Sub RemoveMe()
        Try
            Parent?.Nodes.Remove(Me)
        Catch ex As InvalidOperationException
        End Try
    End Sub
    Private Sub ShowHide(Nodes As List(Of Node), Optional Flag As Boolean = True)

        For Each Node As Node In Nodes
            If Node.Parent Is Nothing Then
                Node._Visible = True

            Else
                If Node.Parent.Expanded Then
                    Node._Visible = Flag

                Else
                    Node._Visible = False

                End If
            End If
            If Node.HasChildren Then ShowHide(Node.Nodes, Node._Visible)
        Next

    End Sub
    Public ReadOnly Property SortType As Type
        Get
            Return GetDataType((From n In Nodes Select n.SortValue).ToList)
        End Get
    End Property
    Public Sub SortChildren(Optional SortOrder As SortOrder = SortOrder.Ascending)

        Nodes.SortOrder = SortOrder

    End Sub
    Public Overrides Function ToString() As String
        Return Join({Text, If(Nodes.Any, "( " & Nodes.Count & " )", String.Empty)})
    End Function
#End Region
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                Me._Font.Dispose()
                Me.Font.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class