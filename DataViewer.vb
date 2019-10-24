Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel
Imports System.Drawing.Drawing2D
#Region " STRUCTURES + ENUMERATIONS "
Public Structure PointStruct
    'https://www.infoq.com/articles/Equality-Overloading-DotNET/
    Implements IEquatable(Of PointStruct)
    Public ReadOnly Property X() As Integer
    Public ReadOnly Property Y() As Integer
    Public Sub New(ByVal x As Integer, ByVal y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
    Public Overrides Function GetHashCode() As Integer
        Return X.GetHashCode Xor Y.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As PointStruct) As Boolean Implements IEquatable(Of PointStruct).Equals
        Return X = other.X AndAlso Y = other.Y
    End Function
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is PointStruct Then
            Return CType(obj, PointStruct) = Me
        Else
            Return False
        End If
    End Function
    Public Shared Operator =(ByVal point1 As PointStruct, ByVal point2 As PointStruct) As Boolean
        Return point1.Equals(point2)
    End Operator
    Public Shared Operator <>(ByVal point1 As PointStruct,
      ByVal point2 As PointStruct) As Boolean
        Return Not (point1 = point2)
    End Operator
End Structure
Public Structure MouseInfo
    Implements IEquatable(Of MouseInfo)
    Public Property Column As Column
    Public Property MouseOver As Column
    Public Property Row As Row
    Public Property Bounds As Rectangle
    Public Property Point As Point
    Public Property CurrentAction As Action
    Public Enum Action
        None
        MouseOverHead
        MouseOverGrid
        MouseOverHeadEdge
        HeaderEdgeClicked
        HeaderClicked
        ColumnDragging
        ColumnSizing
        Filtering
        CellClicked
        CellSelecting
    End Enum
    Public Overrides Function GetHashCode() As Integer
        Return Point.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As MouseInfo) As Boolean Implements IEquatable(Of MouseInfo).Equals
        Return Point = other.Point
    End Function
    Public Shared Operator =(ByVal Object1 As MouseInfo, ByVal Object2 As MouseInfo) As Boolean
        Return Object1.Equals(Object2)
    End Operator
    Public Shared Operator <>(ByVal Object1 As MouseInfo, ByVal Object2 As MouseInfo) As Boolean
        Return Not Object1 = Object2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is MouseInfo Then
            Return CType(obj, MouseInfo) = Me
        Else
            Return False
        End If
    End Function
End Structure
Public Enum Scaling
    GrowParent
    ShrinkChild
End Enum
#End Region
Public Class ViewerEventArgs
    Inherits EventArgs
    Public ReadOnly Property MouseData As MouseInfo
    Public Sub New(MI As MouseInfo)
        MouseData = MI
    End Sub
End Class
Public Class DataViewer
    Inherits Control
#Region " GENERAL DECLARATIONS "
    Private ReadOnly BindingSource As New BindingSource
    Private WithEvents SpinTimer As New Timer With {.Interval = 150, .Tag = 0}
    Private WithEvents TSDD_Spin As New ToolStripDropDown With {.AutoClose = False, .AutoSize = False, .Padding = New Padding(0), .DropShadowEnabled = False, .BackColor = Color.Transparent}
    Private WithEvents PB_Spin As New PictureBox With {.Size = My.Resources.Spin1.Size, .Margin = New Padding(0), .BackColor = Color.Transparent, .BorderStyle = BorderStyle.None}
    Private WithEvents Bar_Spin As New ProgressBar With {.Size = New Size(PB_Spin.Width, 20), .Minimum = 0, .Style = ProgressBarStyle.Continuous, .ForeColor = Color.Tomato}
    Public WithEvents VScroll As New VScrollBar With {.Minimum = 0}
    Public WithEvents HScroll As New HScrollBar With {.Minimum = 0}
    Private ReadOnly ColumnHeadTip As ToolTip = New ToolTip With {.BackColor = Color.Black, .ForeColor = Color.White}
#End Region

#Region " INITIALIZE "
    Public Sub New()
        Controls.AddRange({VScroll, HScroll})
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.ContainerControl, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)
        SetStyle(ControlStyles.Selectable, True)
        SetStyle(ControlStyles.Opaque, True)
        SetStyle(ControlStyles.UserMouse, True)
        BackColor = SystemColors.Window
        Size = New Size(450, 350)
        Margin = New Padding(1)
        MaximumSize = WorkingArea.Size
        Application.EnableVisualStyles()
        With TSDD_Spin
            .Items.Add(New ToolStripControlHost(PB_Spin) With {.BackColor = Color.Transparent})
            .Items.Add(New ToolStripControlHost(Bar_Spin) With {.BackColor = Color.GhostWhite})
            'Using rgn As Region = New Region(New Rectangle(0, 0, 100, 100))
            '    rgn.Union(New RectangleF(0, 100 + 2, 150, 200))
            '    .Region = rgn
            'End Using
        End With

    End Sub
#End Region
#Region " DRAWING "
    'HEADER / ROW PROPERTIES...USE A TEMPLATE LIKE MS AND APPLY TO CELL, ROW, HEADER...{V/H ALIGNMENT, FONT, FORCOLOR, BACKCOLOR, ETC}
    Private PaintCount As Integer = 0
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        SetupScrolls()
        If e IsNot Nothing Then
            e.Graphics.FillRectangle(New SolidBrush(BackColor), ClientRectangle)
            PaintCount += 1
#Region " DRAW HEADERS "
            With Columns
                Dim HeadFullBounds As New Rectangle(0, 0, {1, .HeadBounds.Width}.Max, .HeadBounds.Height)
                HeadFullBounds.Offset(-HScroll.Value, 0)
                Using LinearBrush As New LinearGradientBrush(HeadFullBounds, .HeaderStyle.BackColor, .HeaderStyle.ShadeColor, LinearGradientMode.Vertical)
                    e.Graphics.FillRectangle(LinearBrush, HeadFullBounds)
                End Using
                If Not .Any Then
                    TextRenderer.DrawText(e.Graphics, Name, Font, HeadFullBounds, .HeaderStyle.ForeColor, Color.Transparent, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                End If
                Dim ColumnStart As Integer = ColumnIndex(HScroll.Value)
                VisibleColumns.Clear()
                For ColumnIndex As Integer = ColumnStart To {Columns.Count - 1, ColumnStart + 20}.Min
                    Dim Column As Column = Columns(ColumnIndex)
                    With Column
                        If .Visible Then
                            Dim HeadBounds As Rectangle = .HeadBounds
                            HeadBounds.Offset(-HScroll.Value, 0)
                            VisibleColumns.Add(Column, HeadBounds)
                            If HeadBounds.Left >= Width Then Exit For
                            Using LinearBrush As New LinearGradientBrush(HeadBounds, .HeaderStyle.BackColor, .HeaderStyle.ShadeColor, LinearGradientMode.Vertical)
                                e.Graphics.FillRectangle(LinearBrush, HeadBounds)
                            End Using
                            e.Graphics.DrawRectangle(Pens.Silver, HeadBounds)
#Region " DRAW HEADER IMAGE "
                            Dim ImageBounds As Rectangle = .ImageBounds
                            ImageBounds.Offset(-HScroll.Value, 0)
                            If Not IsNothing(.Image) Then e.Graphics.DrawImage(.Image, ImageBounds)
#End Region
#Region " DRAW HEADER TEXT "
                            Dim TextBounds As Rectangle = .TextBounds
                            TextBounds.Offset(-HScroll.Value, 0)
                            TextRenderer.DrawText(e.Graphics, .Text, .HeaderStyle.Font, TextBounds, .HeaderStyle.ForeColor, Color.Transparent, TextFormatFlags.VerticalCenter Or TextFormatFlags.HorizontalCenter)
#End Region
#Region " DRAW SORT TRIANGLE "
                            If Not .SortOrder = SortOrder.None Then
                                Dim SortPoints As New List(Of Point)
                                Dim SortArrowW As Integer = 12
                                Dim SortArrowH As Integer = 7
                                Dim X_Offset As Integer = Convert.ToInt32((.SortBounds.Width - SortArrowW) / 2)
                                Dim Y_Offset As Integer = Convert.ToInt32((.Height - SortArrowH) / 2)
                                Dim SortX As Integer = If(.Filtered, (From X In SortPoints Select X.X).Min, HeadBounds.Right - X_Offset) - X_Offset - SortArrowW
                                Dim SortR As Integer = SortX + SortArrowW
                                Dim MidPoint As Integer = SortX + Convert.ToInt32(SortArrowW / 2)
                                If .SortOrder = SortOrder.Ascending Then
                                    SortPoints.AddRange({New Point(SortX, Y_Offset), New Point(SortR, Y_Offset), New Point(MidPoint, Y_Offset + SortArrowH)})
                                Else
                                    SortPoints.AddRange({New Point(SortX, Y_Offset + SortArrowH), New Point(SortR, Y_Offset + SortArrowH), New Point(MidPoint, Y_Offset)})
                                End If
                                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
                                Using GP As New GraphicsPath
                                    GP.AddPolygon(SortPoints.ToArray)
                                    Using PathBrush As New PathGradientBrush(GP)
                                        PathBrush.CenterColor = Color.Gray
                                        PathBrush.SurroundColors = {Color.Gainsboro}
                                        PathBrush.CenterPoint = SortPoints(0)
                                        e.Graphics.FillPolygon(PathBrush, SortPoints.ToArray)
                                    End Using
                                End Using
                                e.Graphics.DrawPolygon(Pens.DimGray, SortPoints.ToArray)
                            End If
#End Region
                            e.Graphics.DrawRectangle(Pens.Silver, HeadBounds)
                            If Column Is _MouseData.Column Then
                                If _MouseData.CurrentAction = MouseInfo.Action.MouseOverHead Then
                                    Using HighlightBrush As New SolidBrush(Color.FromArgb(128, Color.Yellow))
                                        e.Graphics.FillRectangle(HighlightBrush, HeadBounds)
                                    End Using
                                End If
                            End If
                        End If
                    End With
                Next
#Region " DRAW HEADER EDGE "
                With _MouseData
                    If .CurrentAction = MouseInfo.Action.MouseOverHeadEdge Then
                        Dim EdgeBounds As Rectangle = .Column.EdgeBounds
                        EdgeBounds.Offset(-HScroll.Value, 0)
                        Using HighlightBrush As New SolidBrush(Color.FromArgb(128, Color.LimeGreen))
                            e.Graphics.FillRectangle(HighlightBrush, EdgeBounds)
                        End Using
                    End If
                End With
#End Region
#End Region
#Region " DRAW ROWS "
                If Rows.Any Then
                    Dim Top As Integer = HeaderHeight
                    Dim IsOddRow As Boolean
                    Dim RowStart As Integer = RowIndex(VScroll.Value)
                    VisibleRows.Clear()
                    For RowIndex As Integer = RowStart To {RowStart + 40, Rows.Count - 1}.Min
                        Dim Row = Rows(RowIndex)
                        IsOddRow = RowIndex Mod 2 = 1
                        Dim MouseOverRow As Boolean = _MouseData.Row Is Row And _MouseData.CurrentAction = MouseInfo.Action.MouseOverGrid
                        With Row
                            Dim RowBounds = New Rectangle(0, Top, HeadFullBounds.Width, RowHeight)
                            RowBounds.Offset(-HScroll.Value, 0)
                            VisibleRows.Add(Row, RowBounds)
                            If RowBounds.Top >= Bottom Then Exit For
                            If Row.Selected Then
                                Using LinearBrush As New LinearGradientBrush(RowBounds, Rows.SelectionRowStyle.BackColor, Rows.SelectionRowStyle.ShadeColor, LinearGradientMode.Vertical)
                                    e.Graphics.FillRectangle(LinearBrush, RowBounds)
                                End Using
                            Else
                                If IsOddRow Then
                                    Using LinearBrush As New LinearGradientBrush(RowBounds, Rows.AlternatingRowStyle.BackColor, Rows.AlternatingRowStyle.ShadeColor, LinearGradientMode.Vertical)
                                        e.Graphics.FillRectangle(LinearBrush, RowBounds)
                                    End Using
                                Else
                                    Using LinearBrush As New LinearGradientBrush(RowBounds, Rows.RowStyle.BackColor, Rows.RowStyle.ShadeColor, LinearGradientMode.Vertical)
                                        e.Graphics.FillRectangle(LinearBrush, RowBounds)
                                    End Using
                                End If
                            End If
#Region " DRAW CELLS "
                            For Each Column In VisibleColumns.Keys
                                With Column
                                    Dim CellBounds As New Rectangle(.HeadBounds.Left, RowBounds.Top, .HeadBounds.Width, RowBounds.Height)
                                    CellBounds.Offset(-HScroll.Value, 0)
                                    Dim CellValue As Object = Row.Cell(.Name)
                                    Dim CellString = Format(CellValue, .Format.Value)
                                    Dim CellIsDBNull As Boolean = IsDBNull(Row.DataRow(.Name))
                                    'Boolean, Image, Text, Decimal, Date, Time, Integer
                                    If CellIsDBNull Then
                                        CellString = "Null"

                                    ElseIf .Format.Key = "Date" Then
                                        Dim CellDate As Date
                                        Dim DateParse As Boolean = Date.TryParse(CellValue.ToString, CellDate)
                                        CellString = CellDate.ToString(.Format.Value, InvariantCulture)

                                    ElseIf .Format.Key = "Time" Then
                                        Dim CellDate As Date
                                        Dim TimeParse As Boolean = Date.TryParse(CellValue.ToString, CellDate)
                                        CellString = CellDate.ToString(.Format.Value, InvariantCulture)

                                    ElseIf .Format.Key = "Decimal" Then
                                        Dim CellNumber As Double
                                        Dim DoubleParse As Boolean = Double.TryParse(CellValue.ToString, CellNumber)
                                        CellString = CellNumber.ToString(.Format.Value, InvariantCulture)

                                    End If
                                    If .Format.Key = "Boolean" Then
                                        Dim CheckBounds As New Rectangle(CellBounds.X + 2, CellBounds.Top + Convert.ToInt32((CellBounds.Height - 14) / 2), 14, 14)
                                        e.Graphics.DrawImage(Base64ToImage(If(Convert.ToBoolean(CellString, InvariantCulture), CheckString, UnCheckString)), CheckBounds)
                                        If MouseOverRow Then
                                            Using CheckBrush As New SolidBrush(Color.FromArgb(128, Color.Yellow))
                                                e.Graphics.FillRectangle(CheckBrush, CheckBounds)
                                            End Using
                                        End If
                                    ElseIf .Format.Key = "Image" Then
                                        Dim CellImage As Image = Base64ToImage(CellValue.ToString)
                                        Dim xOffset As Integer = Convert.ToInt32((CellBounds.Width - CellImage.Width) / 2)
                                        Dim yOffset As Integer = Convert.ToInt32((CellBounds.Height - CellImage.Height) / 2)
                                        Dim ImageBounds As New Rectangle(CellBounds.X + xOffset, CellBounds.Y + yOffset, CellImage.Width, CellImage.Height)
                                        e.Graphics.DrawImage(CellImage, ImageBounds)

                                    Else
                                        Dim RowFont As Font = If(IsOddRow, Rows.AlternatingRowStyle.Font, Rows.RowStyle.Font)
                                        If MouseOverRow Then RowFont = New Font(RowFont, FontStyle.Underline)
                                        If Row.Selected Then
                                            TextRenderer.DrawText(e.Graphics, CellString, RowFont, CellBounds, Rows.SelectionRowStyle.ForeColor, Color.Transparent, .Alignment)
                                        Else
                                            If IsOddRow Then
                                                TextRenderer.DrawText(e.Graphics, CellString, RowFont, CellBounds, Rows.AlternatingRowStyle.ForeColor, Color.Transparent, .Alignment)
                                            Else
                                                TextRenderer.DrawText(e.Graphics, CellString, RowFont, CellBounds, Rows.RowStyle.ForeColor, Color.Transparent, .Alignment)
                                            End If
                                        End If

                                    End If
                                    If CellIsDBNull Then
                                        Using NullBrush As New SolidBrush(Color.FromArgb(128, Color.Gainsboro))
                                            e.Graphics.FillRectangle(NullBrush, CellBounds)
                                        End Using
                                    End If
                                    ControlPaint.DrawBorder3D(e.Graphics, CellBounds, Border3DStyle.SunkenOuter)
                                End With
                            Next
#End Region
                            Top += RowHeight
                        End With
                    Next
                    If Rows.Count < 20 And VisibleRows.Count <> Rows.Count Then Stop
                End If
#End Region
#Region " xxx "
                '                With Column.Head
                '#Region " HEADER BACKGROUND + MOUSEOVER "
                '                    Using LinearBrush As New LinearGradientBrush(.Bounds, TheBackColor, TheShadeColor, LinearGradientMode.Vertical)
                '                        Graphics.FillRectangle(LinearBrush, .Bounds)
                '                    End Using
                '                    Dim BottomHalf As New Rectangle(.Bounds.Left, Convert.ToInt32(.Bounds.Height / 2), .Bounds.Width, Convert.ToInt32(.Bounds.Height / 2))
                '                    Using LinearBrush As New LinearGradientBrush(.Bounds, TheBackColor, TheShadeColor, LinearGradientMode.Vertical)
                '                        Graphics.FillRectangle(LinearBrush, BottomHalf)
                '                    End Using
                '                    '============= MOUSE IS OVER THIS COLUMN
                '                    If _MouseData.MouseOver Is Column Then
                '                        Using FadedBrush As New SolidBrush(Color.FromArgb(128, TheShadeColor))
                '                            Graphics.FillRectangle(FadedBrush, .Bounds)
                '                        End Using
                '                    End If
                '#End Region
                '#Region " DRAW IMAGE "
                '                    If Not IsNothing(.Image) Then Graphics.DrawImage(.Image, .ImageBounds)
                '#End Region
                '#Region " DRAW TEXT "
                '                    TextRenderer.DrawText(Graphics, .Text, TheFont, .TextBounds, TheForeColor, Color.Transparent, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                '#End Region
                '#Region " DRAW FILTER "
                '                    Graphics.DrawRectangle(Pens.Red, .FilterBounds)
                '#End Region
                '#Region " DRAW SORT TRIANGLE "
                '                    If Not Column.Sort = Column.SortOrder.None Then
                '                        Dim SortPoints As New List(Of Point)
                '                        Dim SortArrowW As Integer = 12
                '                        Dim SortArrowH As Integer = 7
                '                        Dim X_Offset As Integer = Convert.ToInt32((.SortBounds.Width - SortArrowW) / 2)
                '                        Dim Y_Offset As Integer = Convert.ToInt32((.Height - SortArrowH) / 2)
                '                        Dim SortX As Integer = If(.Filtered, (From X In SortPoints Select X.X).Min, .Bounds.Right - X_Offset) - X_Offset - SortArrowW
                '                        Dim SortR As Integer = SortX + SortArrowW
                '                        Dim MidPoint As Integer = SortX + Convert.ToInt32(SortArrowW / 2)
                '                        If Column.Sort = Column.SortOrder.Ascending Then
                '                            SortPoints.AddRange({New Point(SortX, Y_Offset), New Point(SortR, Y_Offset), New Point(MidPoint, Y_Offset + SortArrowH)})
                '                        Else
                '                            SortPoints.AddRange({New Point(SortX, Y_Offset + SortArrowH), New Point(SortR, Y_Offset + SortArrowH), New Point(MidPoint, Y_Offset)})
                '                        End If
                '                        Graphics.SmoothingMode = SmoothingMode.AntiAlias
                '                        Using GP As New GraphicsPath
                '                            GP.AddPolygon(SortPoints.ToArray)
                '                            Using PathBrush As New PathGradientBrush(GP)
                '                                PathBrush.CenterColor = Color.WhiteSmoke
                '                                PathBrush.SurroundColors = {Color.GhostWhite}
                '                                PathBrush.CenterPoint = SortPoints(0)
                '                                Graphics.FillPolygon(PathBrush, SortPoints.ToArray)
                '                            End Using
                '                        End Using
                '                        Graphics.DrawPolygon(Pens.Silver, SortPoints.ToArray)
                '                    End If
                '#End Region
                '#Region " DRAW HEADER BORDER "
                '                    Using BorderPen As New Pen(TheShadeColor)
                '                        Graphics.DrawRectangle(BorderPen, .Bounds)
                '                    End Using
                '#End Region
                '                End With
#End Region
                If VisibleRows.Any Then
                    Dim BottomRow As Rectangle = VisibleRows.Last.Value
#Region " VERTICAL BOUNDARY "
                    With HeadFullBounds
                        If .Right < Width Then
                            Using VerticalPen As New Pen(Color.Silver, 1) With {.DashStyle = DashStyle.DashDot}
                                e.Graphics.DrawLine(VerticalPen, New Point(.Right, .Bottom), New Point(.Right, {BottomRow.Bottom, ClientSize.Height}.Min))
                            End Using
                        End If
                    End With
#End Region
#Region " HORIZONTAL BOUNDARY "
                    With BottomRow
                        If .Bottom < Height Then
                            Using HorizontalPen As New Pen(Color.Silver, 1) With {.DashStyle = DashStyle.DashDot}
                                e.Graphics.DrawLine(HorizontalPen, New Point(0, .Bottom), New Point({ .Right, ClientSize.Width}.Min, .Bottom))
                            End Using
                        End If
                    End With
#End Region
                End If
                ControlPaint.DrawBorder3D(e.Graphics, HeadFullBounds, Border3DStyle.Sunken)
            End With
            ControlPaint.DrawBorder3D(e.Graphics, ClientRectangle, Border3DStyle.Sunken)
        End If

    End Sub
#End Region
    Private Sub SetupScrolls()

        VScroll.Maximum = {HeaderHeight + RowHeight + TotalSize.Height - 1, 0}.Max
        HScroll.Maximum = {TotalSize.Width - 1, 0}.Max
        VScroll.Visible = VScrollVisible
        If VScrollVisible Then
            With VScroll
                .Top = 2
                .Left = {TotalSize.Width, ClientSize.Width - .Width}.Min
                .Height = ClientRectangle.Height - 2
                .SmallChange = Rows.RowHeight
                .LargeChange = ClientRectangle.Height
            End With
        End If
        HScroll.Visible = HScrollVisible
        If HScrollVisible Then
            With HScroll
                .Top = (ClientRectangle.Bottom - HScroll.Height)
                .Left = 0
                .Width = If(VScroll.Visible, ClientRectangle.Width - VScroll.Width, ClientRectangle.Width)
                .LargeChange = ClientRectangle.Width
            End With
        End If

    End Sub
#Region " EVENTS "
    Public Event ColumnsSized(sender As Object, e As ViewerEventArgs)
    Public Event RowsLoading(sender As Object, e As ViewerEventArgs)
    Public Event RowsLoaded(sender As Object, e As ViewerEventArgs)
    Public Event RowClicked(sender As Object, e As ViewerEventArgs)
    Public Event Alert(sender As Object, e As AlertEventArgs)
#End Region
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Public ReadOnly Property Table As DataTable
    Private WithEvents Columns_ As New ColumnCollection(Me)
    Public ReadOnly Property Columns As ColumnCollection
        Get
            Return Columns_
        End Get
    End Property
    Public ReadOnly Property VisibleColumns As New Dictionary(Of Column, Rectangle)
    Private WithEvents Rows_ As New RowCollection(Me)
    Public ReadOnly Property Rows As RowCollection
        Get
            Return Rows_
        End Get
    End Property
    Public ReadOnly Property VisibleRows As New Dictionary(Of Row, Rectangle)
    Public ReadOnly Property TotalSize As Size
        Get
            Return New Size(Columns.Select(Function(c) c.Width).Sum, Rows.Count * Rows.RowHeight)
        End Get
    End Property
    Private ReadOnly Property HeaderHeight As Integer
        Get
            Return Columns.HeadBounds.Height
        End Get
    End Property
    Private ReadOnly Property RowHeight As Integer
        Get
            Return Rows.RowHeight
        End Get
    End Property
    Private Function ColumnIndex(X As Integer) As Integer

        Dim Widths As Integer = 0
        Dim Index As Integer = -1
        For Each Column In Columns
            If X <= Widths Then Exit For
            Widths += Column.Width
            Index += 1
        Next
        Return {0, Index}.Max

    End Function
    Private Function RowIndex(Y As Integer) As Integer
        Return Convert.ToInt32(Split((Y / RowHeight).ToString(InvariantCulture), ".")(0), InvariantCulture)
    End Function
    Private ReadOnly Property VScrollVisible As Boolean
        Get
            Return TotalSize.Height > ClientRectangle.Height
        End Get
    End Property
    Private ReadOnly Property HScrollVisible As Boolean
        Get
            Return TotalSize.Width > ClientRectangle.Width
        End Get
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub WaitTimer_Tick() Handles SpinTimer.Tick

        Dim PB_SpinIndex As Integer = DirectCast(SpinTimer.Tag, Integer) Mod 8
        PB_Spin.Image = Nothing
        If PB_SpinIndex = 0 Then PB_Spin.Image = My.Resources.Spin1
        If PB_SpinIndex = 1 Then PB_Spin.Image = My.Resources.Spin2
        If PB_SpinIndex = 2 Then PB_Spin.Image = My.Resources.Spin3
        If PB_SpinIndex = 3 Then PB_Spin.Image = My.Resources.Spin4
        If PB_SpinIndex = 4 Then PB_Spin.Image = My.Resources.Spin5
        If PB_SpinIndex = 5 Then PB_Spin.Image = My.Resources.Spin6
        If PB_SpinIndex = 6 Then PB_Spin.Image = My.Resources.Spin7
        If PB_SpinIndex = 7 Then PB_Spin.Image = My.Resources.Spin8
        SpinTimer.Tag = PB_SpinIndex + 1

    End Sub
    Private _Waiting As Boolean = False
    Public Property Waiting(Optional VisibleColor As Color = Nothing) As Boolean
        Get
            Return _Waiting
        End Get
        Set(value As Boolean)
            _Waiting = value
            With TSDD_Spin
                Bar_Spin.Value = Bar_Spin.Minimum
                If _Waiting Then
                    Bar_Spin.Maximum = Columns.Count
                    .BackColor = VisibleColor
                    .Size = New Size(PB_Spin.Width, PB_Spin.Height + Bar_Spin.Height)
                    .Show(CenterItem(PB_Spin.Size))
                    SpinTimer.Start()
                Else
                    .BackColor = Color.Transparent
                    .Size = New Size(0, 0)
                    .Hide()
                    SpinTimer.Stop()
                End If
            End With
        End Set
    End Property
    Private Sub RowsLoadingStart() Handles Me.RowsLoading
        Waiting(Color.Tomato) = True
    End Sub
    Private Sub RowsLoadingEnd() Handles Me.RowsLoaded
        Waiting = False
    End Sub
    Private Sub ColumnsSizingStart() Handles Columns_.CollectionSizingStart
        Waiting(Color.LimeGreen) = True
    End Sub
    Private Sub ColumnsSizingEnd() Handles Columns_.CollectionSizingEnd
        Waiting = False
        RaiseEvent ColumnsSized(Me, Nothing)    'Public
    End Sub
    Private Sub ColumnSized(sender As Object, e As EventArgs) Handles Columns_.ColumnSized

        With DirectCast(sender, Column)
            Dim MaxValue As Integer = {Bar_Spin.Value + 1, Bar_Spin.Maximum}.Min
            Bar_Spin.Value = MaxValue
            RaiseEvent Alert(sender, New AlertEventArgs(Join({"Column", .Name, "Index", .ViewIndex, "resized"})))
            Cursor = Cursors.Default
            Invalidate()
        End With

    End Sub
    Private Sub Spin_Clicked() Handles PB_Spin.Click

        If TSDD_Spin.BackColor = Color.Tomato Then
            'Busy with Database
        Else
            Columns.CancelWorkers()
        End If

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private _DataSource As Object
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <Category("Data")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property DataSource As Object
        Get
            Return _DataSource
        End Get
        Set(value As Object)
            If value IsNot _DataSource Then
                Clear()
                _Table = New DataTable
                _DataSource = value
                BindingSource.DataSource = value
#Region " FILL TABLE "
                If DataSource Is Nothing Then
                    Exit Property
                ElseIf TypeOf DataSource Is String Then
                    Exit Property
                ElseIf TypeOf DataSource Is IEnumerable Then
#Region " UNSTRUCTURED "
                    Select Case DataSource.GetType
                        Case GetType(List(Of String()))
#Region " LIST OF STRING() - LIST ITEMS=ROWS...STRING()=COLUMNS "
                            Dim Rows As List(Of String()) = DirectCast(_DataSource, List(Of String()))
                            If Rows IsNot Nothing Then
                                Dim ColumnCount As Integer = (From C In Rows Select C.Count).Max
                                For Column As Integer = 1 To ColumnCount
                                    Table.Columns.Add(New DataColumn With {.ColumnName = "Column" & Column, .DataType = GetType(String)})
                                Next
                                For Each Row As String() In Rows
                                    Table.Rows.Add(Row)
                                Next
                            End If
#End Region
                        Case GetType(List(Of Object()))
#Region " LIST OF OBJECT() - LIST ITEMS=ROWS...STRING()=COLUMNS "
                            Dim Rows As List(Of Object()) = DirectCast(_DataSource, List(Of Object()))
                            If Rows IsNot Nothing Then
                                Dim ColumnCount As Integer = (From C In Rows Select C.Count).Max
                                For Column As Integer = 1 To ColumnCount
                                    Table.Columns.Add(New DataColumn With {.ColumnName = "Column" & Column, .DataType = GetType(String)})
                                Next
                                For Each Row As String() In Rows
                                    Table.Rows.Add(Row)
                                Next
                            End If
#End Region
                        Case Else

                    End Select
#End Region
                ElseIf DataSource.GetType Is GetType(DataTable) Then
                    _Table = DirectCast(DataSource, DataTable)
                End If
#End Region
                For Each DataColumn As DataColumn In Table.Columns
                    Dim NewColumn = Columns.Add(New Column(DataColumn))
                Next
                If Table.AsEnumerable.Any Then
                    Bar_Spin.Value = 0
                    RaiseEvent RowsLoading(Me, Nothing)
                    Dim RowIndex As Integer = 0
                    For Each _DataRow As DataRow In Table.Rows
                        Rows.Add(New Row(_DataRow))
                        For Column = 0 To Table.Columns.Count - 1
                            Dim CellValue = _DataRow(Column)
                            'VALUES DICTIONARY IS USED TO DETERMINE COLUMN TYPE AND FORMATTING
                            'TRIAL LIMITING TO 250 NON-EMPTY VALUES FOR SPEED
                            If IsDBNull(CellValue) Then
                                'Columns(Column).Values.Add(RowIndex, Nothing)
                            Else
                                If Columns(Column).Values.Count <= 250 Then Columns(Column).Values.Add(RowIndex, CellValue)
                            End If
                        Next
                        RowIndex += 1
                    Next
                    RaiseEvent RowsLoaded(Me, Nothing)
                    Invalidate()
                    Columns.AutoSize()
                End If
            End If

        End Set
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Sub Clear()
        'Columns_.Dispose()
        'Rows_.Dispose()
        'Columns_ = New ColumnCollection(Me)
        'Rows_ = New RowCollection(Me)
        Columns.CancelWorkers()
        Columns.Clear()
        Rows.Clear()
        VisibleColumns.Clear()
        VisibleRows.Clear()
        VScroll.Value = 0
        HScroll.Value = 0
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public ReadOnly Property MouseData As New MouseInfo
    Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
        _MouseData = Nothing
        Invalidate()
        MyBase.OnMouseLeave(e)
    End Sub
    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            With _MouseData
                If .CurrentAction = MouseInfo.Action.HeaderEdgeClicked And e.Location <> .Point Then
                    .CurrentAction = MouseInfo.Action.ColumnSizing
                End If
                If .CurrentAction = MouseInfo.Action.ColumnSizing Then
                    If e.Button = MouseButtons.None Then
                        .CurrentAction = MouseInfo.Action.None

                    Else
                        Dim Delta = e.X - .Point.X
                        With .Column
                            .Width += Delta
                            Dim Bullets As New Dictionary(Of String, List(Of String)) From {
                            { .Text, {"Type is " & .Format.Key, "Width=" & .HeadBounds.Width, "Content Width=" & .ContentWidth, "Row Count=" & Rows.Count, "Sort Order is " & .SortOrder.ToString}.ToList}
                        }
                            ColumnHeadTip.SetToolTip(Me, Bulletize(Bullets))
                        End With
                        Cursor = Cursors.VSplit

                    End If
                    .Point = e.Location
                    Invalidate()

                Else
                    Cursor = Cursors.Default
                    .Point = e.Location
                    Dim _LastMouseColumn As Column = .Column
                    Dim _LastMouseRow As Row = .Row
                    Dim MouseColumns = VisibleColumns.Where(Function(r) New Rectangle(r.Value.X, r.Value.Y, r.Value.Width, RowHeight * Rows.Count).Contains(e.Location)).Select(Function(c) c.Key)
                    If MouseColumns.Any Then
                        .Column = MouseColumns.First
                    End If
                    If Columns.HeadBounds.Contains(e.Location) Then
#Region " HEADER REGION "
                        Dim VisibleEdges As New Dictionary(Of Column, Rectangle)
                        Dim ColumnEdge As Column = Nothing
                        For Each Item In VisibleColumns
                            VisibleEdges.Add(Item.Key, New Rectangle(Item.Key.EdgeBounds.X - HScroll.Value, 0, 10, Item.Key.EdgeBounds.Height))
                        Next
                        Dim Edges = VisibleEdges.Where(Function(x) x.Value.Contains(e.Location)).Select(Function(c) c.Key)
                        If Edges.Any Then
                            ColumnEdge = Edges.First
                            .CurrentAction = MouseInfo.Action.MouseOverHeadEdge
                            .Column = ColumnEdge
                            Cursor = Cursors.VSplit
                        Else
                            .CurrentAction = MouseInfo.Action.MouseOverHead
                        End If
#End Region
                    Else
#Region " GRID REGION "
                        Dim MouseRows = VisibleRows.Where(Function(r) e.Y >= r.Value.Top And e.Y <= r.Value.Bottom)
                        If MouseRows.Any Then
                            .Row = MouseRows.First.Key
                            .CurrentAction = MouseInfo.Action.MouseOverGrid
                            If Not Rows.SingleSelect And _LastMouseRow IsNot .Row And ControlKeyDown Then .Row.Selected = e.Button = MouseButtons.Left
                            .Bounds = VisibleRows(.Row)
                        End If
#End Region
                    End If
                    If Columns.HeadBounds.Contains(e.Location) Or .Column IsNot _LastMouseColumn Or .Row IsNot _LastMouseRow Then
                        Invalidate()
                    End If
                End If
            End With
            MyBase.OnMouseMove(e)
        End If

    End Sub
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            With _MouseData
                If e IsNot Nothing And .CurrentAction = MouseInfo.Action.MouseOverHead And .Column IsNot Nothing Then
                    If e.Button = MouseButtons.Left Then
                        With .Column
                            If .SortOrder = SortOrder.Ascending Then
                                .SortOrder = SortOrder.Descending
                            Else
                                .SortOrder = SortOrder.Ascending
                            End If
                            Select Case .Format.Key
                                Case "Text", "Image"
                                    If .SortOrder = SortOrder.Ascending Then Rows.Sort(Function(x, y) String.Compare(Convert.ToString(x.Cell(.Name), InvariantCulture), Convert.ToString(y.Cell(.Name), InvariantCulture), StringComparison.Ordinal))
                                    If .SortOrder = SortOrder.Descending Then Rows.Sort(Function(y, x) String.Compare(Convert.ToString(x.Cell(.Name), InvariantCulture), Convert.ToString(y.Cell(.Name), InvariantCulture), StringComparison.Ordinal))

                                Case "Integer"
                                    If .SortOrder = SortOrder.Ascending Then Rows.Sort(Function(x, y) Convert.ToInt64(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToInt64(y.Cell(.Name), InvariantCulture)))
                                    If .SortOrder = SortOrder.Descending Then Rows.Sort(Function(y, x) Convert.ToInt64(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToInt64(y.Cell(.Name), InvariantCulture)))


                                Case "Decimal"
                                    If .SortOrder = SortOrder.Ascending Then Rows.Sort(Function(x, y) Convert.ToDecimal(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToDecimal(y.Cell(.Name), InvariantCulture)))
                                    If .SortOrder = SortOrder.Descending Then Rows.Sort(Function(y, x) Convert.ToDecimal(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToDecimal(y.Cell(.Name), InvariantCulture)))

                                Case "Date", "Time"
                                    If .SortOrder = SortOrder.Ascending Then Rows.Sort(Function(x, y) Convert.ToDateTime(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToDateTime(y.Cell(.Name), InvariantCulture)))
                                    If .SortOrder = SortOrder.Descending Then Rows.Sort(Function(y, x) Convert.ToDateTime(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToDateTime(y.Cell(.Name), InvariantCulture)))

                                Case "Boolean"
                                    If .SortOrder = SortOrder.Ascending Then Rows.Sort(Function(x, y) Convert.ToBoolean(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToBoolean(y.Cell(.Name), InvariantCulture)))
                                    If .SortOrder = SortOrder.Descending Then Rows.Sort(Function(y, x) Convert.ToBoolean(x.Cell(.Name), InvariantCulture).CompareTo(Convert.ToBoolean(y.Cell(.Name), InvariantCulture)))

                            End Select
                        End With
                        Invalidate()
                    Else
                        Dim Bullets As New List(Of String) From {"Paint Count=" & PaintCount,
                        "Rows Count=" & Rows.Count,
                        "Bounds=" & VisibleRows.First.Value.ToString}
                        ColumnHeadTip.SetToolTip(Me, Bulletize(Bullets.ToArray))
                    End If

                ElseIf .CurrentAction = MouseInfo.Action.MouseOverHeadEdge Then
                    .Point = e.Location
                    .CurrentAction = MouseInfo.Action.HeaderEdgeClicked

                ElseIf .CurrentAction = MouseInfo.Action.MouseOverGrid And e.Button = MouseButtons.Left Then
                    MouseData.Row.Selected = Not MouseData.Row.Selected
                    RaiseEvent RowClicked(Me, New ViewerEventArgs(MouseData))

                End If
            End With
            MyBase.OnMouseDown(e)
        End If

    End Sub
    Protected Overrides Sub OnMouseDoubleClick(ByVal e As MouseEventArgs)

        With _MouseData
            If .CurrentAction = MouseInfo.Action.HeaderEdgeClicked Or .CurrentAction = MouseInfo.Action.MouseOverHeadEdge Then
                .CurrentAction = MouseInfo.Action.MouseOverHeadEdge
                'RemoveHandler .Column.Sized, AddressOf ColumnResized
                'AddHandler .Column.Sized, AddressOf ColumnResized
                Cursor = Cursors.WaitCursor
                .Column.AutoSize()
                Invalidate()
            End If
        End With
        MyBase.OnMouseDoubleClick(e)

    End Sub
    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
        With _MouseData
            .CurrentAction = MouseInfo.Action.None
            Invalidate()
        End With
        MyBase.OnMouseUp(e)
    End Sub
    Private ControlKeyDown As Boolean
    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)

        If e IsNot Nothing Then
            If e.KeyCode = Keys.ControlKey Then
                ControlKeyDown = True
            End If
            MyBase.OnKeyDown(e)
        End If

    End Sub
    Protected Overrides Sub OnKeyUp(ByVal e As KeyEventArgs)

        If e IsNot Nothing Then
            If e.KeyCode = Keys.ControlKey Then
                ControlKeyDown = False
                If Not KeyIsDown(Keys.LButton) Then
                    For Each Row In Rows
                        Row.Selected = False
                    Next
                End If
            End If
            MyBase.OnKeyUp(e)
        End If

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub Scrolled(sender As Object, e As ScrollEventArgs) Handles VScroll.Scroll, HScroll.Scroll
        Invalidate()
    End Sub
#End Region
End Class
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
Public Class ColumnCollection
    Inherits List(Of Column)
    Implements IDisposable
    Friend Event CollectionSizingStart(sender As Object, e As EventArgs)
    Friend Event CollectionSizingEnd(sender As Object, e As EventArgs)
    Friend Event ColumnSized(sender As Object, e As EventArgs)
    Public ReadOnly Property IsBusy As Boolean
    Private WithEvents ReOrderTimer As New Timer With {.Interval = 100}
    Private ReadOnly MoveColumns As New Dictionary(Of Column, Integer)
    Public Sub New(Viewer As DataViewer)
        Parent = Viewer
    End Sub
    Public ReadOnly Property Parent As DataViewer
    Public ReadOnly Property Width As Integer
        Get
            Return Sum(Function(c) c.Width)
        End Get
    End Property
    Public ReadOnly Property HeadBounds As Rectangle
        Get
            If Count = 0 Then
                Dim HeadSize As New Size(Parent.Width, 3 + TextRenderer.MeasureText("XXXXXXXXXXX".ToString(InvariantCulture), HeaderStyle.Font).Height + 3)
                Return New Rectangle(0, 0, HeadSize.Width, HeadSize.Height)
            Else
                Return New Rectangle(0, 0, Max(Function(c) c.HeadBounds.Right), Max(Function(c) c.Height))
            End If
        End Get
    End Property
    Private HeaderStyle_ As New CellStyle With {.BackColor = Color.Black, .ShadeColor = Color.Purple, .ForeColor = Color.White, .Font = New Font("Century Gothic", 9)}
    Public Property HeaderStyle As CellStyle
        Get
            Return HeaderStyle_
        End Get
        Set(value As CellStyle)
            If value IsNot HeaderStyle_ Then
                HeaderStyle_ = value
                For Each Column In Me
                    REM /// DO NOT SET HeaderStyle=HeaderStyle_...that binds each Column to Columns.HeaderStyle
                    With Column
                        If value IsNot Nothing Then
                            .HeaderStyle = New CellStyle With {
                            .Alignment = value.Alignment,
                            .BackColor = value.BackColor,
                            .Font = value.Font,
                            .ForeColor = value.ForeColor,
                            .Padding = value.Padding,
                            .ShadeColor = value.ShadeColor}
                        End If
                    End With
                Next
            End If
        End Set
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Shadows Function Add(ByVal AddColumn As Column) As Column

        If AddColumn IsNot Nothing Then
            With AddColumn
                .Parent_ = Me
                .HeaderStyle = HeaderStyle
                ._Index = Count
                .Left = Where(Function(c) c.ViewIndex < .ViewIndex).Select(Function(c) c.Width).Sum
            End With
            MyBase.Add(AddColumn)
        End If
        Return AddColumn

    End Function
    Public Shadows Function Insert(MoveIndex As Integer, ByVal MoveColumn As Column) As Column

        If MoveColumn IsNot Nothing Then
            With MoveColumn
                .Parent_ = Me
                ._Index = MoveIndex
                .Left = Where(Function(c) c.ViewIndex < .ViewIndex).Select(Function(c) c.Width).Sum
                MyBase.Insert(MoveIndex, MoveColumn)
            End With
        End If
        Return MoveColumn

    End Function
    Public Shadows Function Remove(ByVal RemoveColumn As Column) As Column

        If RemoveColumn IsNot Nothing Then
            With RemoveColumn
                .Parent_ = Nothing
                ._Index = -1
                .Left = -1
            End With
            MyBase.Remove(RemoveColumn)
        End If
        Return RemoveColumn

    End Function
    Public Shadows Function Contains(ByVal ColumnName As String) As Boolean
        Return Item(ColumnName) IsNot Nothing
    End Function
    Public Shadows Function Item(ByVal ColumnName As String) As Column

        Dim Columns = Where(Function(c) c.Name.ToUpperInvariant = ColumnName.ToUpperInvariant)
        If Columns.Any Then
            Return Columns.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Sub Clear()
        'AllSized = False
        'Elapsed.Clear()
        MyBase.Clear()
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    'Private AllSized As Boolean = False
    'Private ReadOnly Elapsed As New Dictionary(Of Integer, KeyValuePair(Of Date, Date))
    'Private Sub ColumnSized(sender As Object, e As EventArgs)

    '    Dim Column As Column = DirectCast(sender, Column)
    '    If Not AllSized Then
    '        If Contains(Column) Then
    '            With Column
    '                Dim KVP = New KeyValuePair(Of Date, Date)(.SizeStart, .SizeEnd)
    '                If Not Elapsed.ContainsKey(.Index) Then Elapsed.Add(.Index, KVP)
    '            End With
    '            If Elapsed.Count = Count Then
    '                Parent.ColumnsResized()
    '                AllSized = True
    '            End If
    '        Else
    '            'Handle!
    '            'Parent.ColumnsResized()
    '            'AllSized = True
    '        End If
    '    End If

    'End Sub
    Friend Sub ColumnsLeft()

        For Each Column In Me
            Dim Lefts = Where(Function(c) c.ViewIndex < Column.ViewIndex).Select(Function(c) c.Width)
            If Lefts.Any Then
                Column.Left = Lefts.Sum
            End If
        Next

    End Sub
    Friend Sub ColumnsHeight()

        Dim MaxHeight As Integer = Max(Function(c) c.ImageScaledHeight)
        For Each Column In Me
            Column.Height = MaxHeight
        Next

    End Sub
    Friend Sub Reorder(Column As Column, ViewIndex As Integer)

        ReOrderTimer.Start()
        MoveColumns.Add(Column, ViewIndex)

    End Sub
    Private Sub ReOrderTimer_Tick() Handles ReOrderTimer.Tick

        ReOrderTimer.Stop()
        If IsBusy Then
            AddHandler CollectionSizingEnd, AddressOf CanReorder
        Else
            CanReorder(Nothing, Nothing)
        End If

    End Sub
    Private Sub CanReorder(sender As Object, e As EventArgs)

        RemoveHandler CollectionSizingEnd, AddressOf CanReorder
        First.Left = 0
        _IsBusy = True
        For Each Column In MoveColumns
            Remove(Column.Key)
            Insert(Column.Value, Column.Key)
        Next
        MoveColumns.Clear()
        _IsBusy = False
        Dim Columns As New List(Of Column) From {First}
        For Each Column In Skip(1)
            Column.Left = Columns.Sum(Function(c) c.Width)
            Columns.Add(Column)
        Next
        Parent.Refresh()

    End Sub
    Private WithEvents ColumnsWorker As New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
    Public Sub AutoSize()
        RaiseEvent CollectionSizingStart(Me, Nothing)
        ColumnsWorker.RunWorkerAsync()
    End Sub
    Private Sub FormatSizeWorker_Start(sender As Object, e As DoWorkEventArgs) Handles ColumnsWorker.DoWork

        If Not IsBusy Then
            _IsBusy = True
            For Each Column In Where(Function(c) c.Visible)
                With Column
                    ._DataType = GetDataType(.Values.Values.ToList)
                    .DataStart = Now
                    .Formatted_ = False
#Region " ALIGNMENT "
                    Select Case .DataType
                        Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long), GetType(Date)
                            .GridStyle.Alignment = ContentAlignment.MiddleCenter

                        Case GetType(Decimal), GetType(Double)
                            .GridStyle.Alignment = ContentAlignment.MiddleRight

                        Case GetType(String), GetType(Boolean)
                            .GridStyle.Alignment = ContentAlignment.MiddleLeft

                    End Select
#End Region
#Region " FORMAT "
                    Select Case .DataType
                        Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                            .Format = New KeyValuePair(Of String, String)("Integer", String.Empty)
                            .DefaultValue = 0

                        Case GetType(Date)
                            Dim DateValue As Date
                            Dim Dates As New List(Of Date)(From D In .Values Where Date.TryParse(Convert.ToString(D.Value, InvariantCulture), DateValue) Select Date.Parse(D.Value.ToString, InvariantCulture))
                            REM /// TEST IF ALWAYS MIDNIGHT
                            Dim CultureInfo = Threading.Thread.CurrentThread.CurrentCulture
                            If Dates.Count = (From D In Dates Where D.TimeOfDay.Ticks = 0 Select D).Count Then
                                .Format = New KeyValuePair(Of String, String)("Date", CultureInfo.DateTimeFormat.ShortDatePattern)
                            Else
                                .Format = New KeyValuePair(Of String, String)("Time", Microsoft.VisualBasic.Join({CultureInfo.DateTimeFormat.ShortDatePattern, "@", CultureInfo.DateTimeFormat.ShortTimePattern}))
                            End If
                            .DefaultValue = New Date

                        Case GetType(Decimal), GetType(Double)
                            .Format = New KeyValuePair(Of String, String)("Decimal", "C2")
                            .DefaultValue = 0

                        Case GetType(String)
                            .Format = New KeyValuePair(Of String, String)("Text", String.Empty)
                            Dim BooleanValue As Boolean
                            Dim Booleans As New List(Of Boolean)(From N In .Values Select Boolean.TryParse(Convert.ToString(N.Value, InvariantCulture), BooleanValue))
                            Booleans = Booleans.Distinct.ToList
                            Dim Images = From N In .Values Where Convert.ToString(N.Value, InvariantCulture).StartsWith("iVBORw0KGgo", StringComparison.InvariantCulture)
                            If Booleans.Count = 1 AndAlso Booleans.First = True Then
                                .Format = New KeyValuePair(Of String, String)("Boolean", String.Empty)
                            ElseIf Images.Any Then
                                .Format = New KeyValuePair(Of String, String)("Image", String.Empty)
                            Else
                                .Format = New KeyValuePair(Of String, String)("Text", String.Empty)
                            End If
                            .DefaultValue = String.Empty

                        Case GetType(Boolean)
                            .Format = New KeyValuePair(Of String, String)("Boolean", String.Empty)
                            .DefaultValue = False

                    End Select
#End Region
                    .DataEnd = Now
                    .Formatted_ = True
                End With
                SizeColumn(Column, True)
                If ColumnsWorker.CancellationPending Then Exit For
            Next
        End If

    End Sub
    Friend Sub SizeColumn(ColumnItem As Column, Optional BackgroundProcess As Boolean = False)

        With ColumnItem
            .SizeStart = Now
            If .Values.Any Then
                If .Format.Key = "Image" Then
                    .ContentWidth = (From V In .Values Select Base64ToImage(V.Value.ToString).Width).Max
                Else
                    .ContentWidth = (From V In .Values Select TextRenderer.MeasureText(Format(V.Value, .Format.Value), If(V.Key Mod 2 = 0, Parent.Rows.RowStyle.Font, Parent.Rows.AlternatingRowStyle.Font)).Width).Max
                End If
            End If
            .Width = { .HeadWidth, .ContentWidth, .MinimumWidth}.Max
            ColumnsLeft()
            .SizeEnd = Now
        End With
        If BackgroundProcess Then ColumnsWorker.ReportProgress({0, ColumnItem.Index}.Max)

    End Sub
    Private Sub FormatSizeColumn_Progress(sender As Object, e As ProgressChangedEventArgs) Handles ColumnsWorker.ProgressChanged
        RaiseEvent ColumnSized(Me(e.ProgressPercentage), Nothing)
    End Sub
    Private Sub FormatSizeWorker_End(sender As Object, e As RunWorkerCompletedEventArgs) Handles ColumnsWorker.RunWorkerCompleted
        _IsBusy = False
        RaiseEvent CollectionSizingEnd(Me, Nothing)
    End Sub
    Friend Sub CancelWorkers()
        ColumnsWorker.CancelAsync()
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                ColumnsWorker.Dispose()
                HeaderStyle_.Dispose()
                ReOrderTimer.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
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
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
<Serializable()> Public Class Column
    Implements IDisposable
    Public Sub New(NewColumn As DataColumn)

        If NewColumn IsNot Nothing Then
            With NewColumn
                Name = .ColumnName
                Table = .Table
            End With
        End If

    End Sub
    Public Property DefaultValue As Object
    <NonSerialized> Friend Parent_ As ColumnCollection
    Public ReadOnly Property Parent As ColumnCollection
        Get
            Return Parent_
        End Get
    End Property
    Public ReadOnly Property Table As DataTable
    Public ReadOnly Property Alignment As TextFormatFlags
        Get
            Select Case GridStyle.Alignment
                Case ContentAlignment.BottomCenter
                    Return TextFormatFlags.Bottom Or TextFormatFlags.HorizontalCenter

                Case ContentAlignment.BottomLeft
                    Return TextFormatFlags.Bottom Or TextFormatFlags.Left

                Case ContentAlignment.BottomRight
                    Return TextFormatFlags.Bottom Or TextFormatFlags.Right

                Case ContentAlignment.MiddleCenter
                    Return TextFormatFlags.VerticalCenter Or TextFormatFlags.HorizontalCenter

                Case ContentAlignment.MiddleLeft
                    Return TextFormatFlags.VerticalCenter Or TextFormatFlags.Left

                Case ContentAlignment.MiddleRight
                    Return TextFormatFlags.VerticalCenter Or TextFormatFlags.Right

                Case ContentAlignment.TopCenter
                    Return TextFormatFlags.Top Or TextFormatFlags.HorizontalCenter

                Case ContentAlignment.TopLeft
                    Return TextFormatFlags.Top Or TextFormatFlags.Left

                Case ContentAlignment.TopRight
                    Return TextFormatFlags.Top Or TextFormatFlags.Right

                Case Else
                    Return TextFormatFlags.VerticalCenter Or TextFormatFlags.HorizontalCenter

            End Select
        End Get
    End Property
    Friend _Index As Integer = 0
    Public ReadOnly Property Index As Integer
        Get
            Return _Index
        End Get
    End Property
    Public Property ViewIndex As Integer
        Get
            If Parent Is Nothing Then
                Return -1
            Else
                Return Parent.IndexOf(Me)
            End If
        End Get
        Set(value As Integer)
            If Parent IsNot Nothing Then Parent.Reorder(Me, value)
        End Set
    End Property
    Private _Name As String
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            If value <> _Name Then
                _Name = value
                Text = value
            End If
        End Set
    End Property
    Private WithEvents HeaderStyle_ As New CellStyle With {.BackColor = Color.Black, .ShadeColor = Color.Purple, .ForeColor = Color.White, .Font = New Font("Century Gothic", 9), .Alignment = ContentAlignment.MiddleCenter}
    Public Property HeaderStyle As CellStyle
        Get
            Return HeaderStyle_
        End Get
        Set(value As CellStyle)
            If value IsNot HeaderStyle_ Then
                HeaderStyle_ = value
                ViewerInvalidate()
            End If
        End Set
    End Property
    Private WithEvents GridStyle_ As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.Transparent, .ForeColor = Color.Transparent, .Font = New Font("Century Gothic", 8)}
    Public Property GridStyle As CellStyle
        Get
            Return GridStyle_
        End Get
        Set(value As CellStyle)
            If value IsNot GridStyle_ Then
                GridStyle_ = value
            End If
        End Set
    End Property
    Friend _Left As Integer = 0
    Friend Property Left As Integer
        Get
            Return _Left
        End Get
        Set(value As Integer)
            'LEFT Property is managed by Parent ColumnCollection...which means RowCollection is accesible
            _Left = value
            ReBounds()
        End Set
    End Property
    Private ReadOnly Property DefaultHeight As Integer
        Get
            Return 4 + TextSize.Height + 4
        End Get
    End Property
    Private _Height As Integer = 24
    Public Property Height As Integer
        Get
            Return _Height
        End Get
        Set(value As Integer)
            If _Height <> value Then
                _Height = value
                ReBounds()
                Parent.ColumnsHeight()
            End If
        End Set
    End Property
    Friend ReadOnly Property ImageScaledHeight As Integer
        Get
            'Padding=2
            'Image=16
            'Padding=2
            'Height = 20
            If Image IsNot Nothing Then
                If Image.Height > DefaultHeight And ImageScaling = Scaling.GrowParent Then
                    Return Image.Height
                Else
                    Return {Image.Height, DefaultHeight}.Min
                End If
            Else
                Return DefaultHeight
            End If
        End Get
    End Property
    Private _MinimumWidth As Integer = 60
    Public Property MinimumWidth As Integer
        Get
            Return _MinimumWidth
        End Get
        Set(value As Integer)
            If _MinimumWidth <> value Then
                _MinimumWidth = value
                ReBounds()
            End If
        End Set
    End Property
    Private _Width As Integer = 1
    Public Property Width As Integer
        Get
            If Visible Then
                Return _Width
            Else
                Return 0
            End If
        End Get
        Set(value As Integer)
            If _Width <> value Then
                If value < 2 Then value = 2
                'value = {value, MinimumWidth}.Max
                _Width = value
                ReBounds()
                If Parent IsNot Nothing Then Parent.ColumnsLeft()
            End If
        End Set
    End Property
    Public Property HeadWidth As Integer = 0
    Public Property ContentWidth As Integer = 0
    Private _Scaling As New Scaling
    Public Property ImageScaling As Scaling
        Get
            Return _Scaling
        End Get
        Set(value As Scaling)
            If _Scaling <> value Then
                _Scaling = value
                Height = DefaultHeight
                Width = 1
            End If
        End Set
    End Property
    Public ReadOnly Property ImageBounds As New Rectangle(0, 0, 0, 0)
    Private _Image As Image = Nothing
    Property Image() As Image
        Get
            Return _Image
        End Get
        Set(ByVal value As Image)
            If Not SameImage(value, _Image) Then
                _Image = value
                'If value IsNot Nothing Then
                '    Dim BM As New Bitmap(value)
                '    BM.MakeTransparent(BM.GetPixel(0, 0))
                '    _Image = BM
                'End If
                ReBounds()
            End If
        End Set
    End Property
    Public ReadOnly Property TextBounds As New Rectangle(0, 0, 0, 0)
    Private _Text As String = Nothing
    Property Text() As String
        Get
            Return _Text
        End Get
        Set(ByVal value As String)
            If _Text <> value Then
                _Text = value
                ViewerInvalidate()
            End If
        End Set
    End Property
    Private _TextSize As Size
    Private ReadOnly Property TextSize As Size
        Get
            Return _TextSize
        End Get
    End Property
    Public ReadOnly Property FilterBounds As New Rectangle(0, 0, 0, 0)
    Private _Filtered As Boolean = False
    Friend Property Filtered As Boolean
        Get
            Return _Filtered
        End Get
        Set(value As Boolean)
            If _Filtered <> value Then
                _Filtered = value
                ReBounds()
            End If
        End Set
    End Property
    Public ReadOnly Property SortBounds As New Rectangle(0, 0, 0, 0)
    Private _SortOrder As SortOrder = SortOrder.None
    Public Property SortOrder As SortOrder
        Get
            Return _SortOrder
        End Get
        Set(value As SortOrder)
            If Not value = _SortOrder Then
                Dim Initialize = _SortOrder = SortOrder.None
                _SortOrder = value
                If Initialize Then ReBounds()
            End If
        End Set
    End Property
    ReadOnly Property HeadBounds As New Rectangle(0, 0, 0, 0)
    Public ReadOnly Property GridBounds As New Rectangle(0, 0, 0, 0)
    ReadOnly Property EdgeBounds As New Rectangle(0, 0, 0, 0)
    Private _Format As New KeyValuePair(Of String, String)(String.Empty, String.Empty)
    Property Format() As KeyValuePair(Of String, String)
        Get
            Return _Format
        End Get
        Set(ByVal value As KeyValuePair(Of String, String))
            If _Format.Key <> value.Key Then
                _Format = value
            End If
        End Set
    End Property
    Public ReadOnly Property Values As New Dictionary(Of Integer, Object)
    Public Property Visible As Boolean = True
    Private Sub ReBounds()

        Dim _SpaceWidth As Integer = 2
        With _ImageBounds
            Dim ImageHeight As Integer = ImageScaledHeight
            If IsNothing(Image) Then
                .X = Left
                .Width = 0
                .Y = 0
                .Height = Height
            Else
                .X = Left + _SpaceWidth
                .Width = ImageScaledHeight
                .Y = Convert.ToInt32((Height - ImageHeight) / 2)
                .Height = .Width
            End If
        End With
        With _TextBounds
            .X = _ImageBounds.Right + If(.Width = 0, 0, _SpaceWidth)
            .Width = TextSize.Width
            .Y = 0
            .Height = Height
        End With
        With _FilterBounds
            .Width = If(Filtered, 16, 0)
            .X = _TextBounds.Right + If(.Width = 0, 0, _SpaceWidth)
            .Y = 0
            .Height = Height
        End With
        With _SortBounds
            .Width = If(SortOrder = SortOrder.None, 0, 16)
            .X = _FilterBounds.Right + If(.Width = 0, 0, _SpaceWidth)
            .Y = 0
            .Height = Height
        End With

        Dim CombinedWidth As Integer = SortBounds.Right - Left
        If CombinedWidth < MinimumWidth Then
            'Pad TextBounds
            _TextBounds.Width += MinimumWidth - CombinedWidth
            'ReBounds()
        End If
        HeadWidth = SortBounds.Right - Left

        If Not Formatted Then _Width = HeadWidth

        If ImageBounds.Height > Height Then
            Height = ImageBounds.Height
        End If
        With _HeadBounds
            .X = Left
            .Y = 0
            .Width = Width
            .Height = Height
        End With
        With _EdgeBounds
            .X = Left + Width - 5
            .Y = 0
            .Width = 10
            .Height = Height
        End With
        With _GridBounds
            .X = Left
            .Y = Height
            .Width = Width
            .Height = Height
        End With

    End Sub
    Private Sub ViewerInvalidate() Handles HeaderStyle_.PropertyChanged

        _TextSize = TextRenderer.MeasureText(Text, HeaderStyle.Font)
        If Parent IsNot Nothing Then
            ReBounds()
            Parent.Parent.Invalidate()
        End If

    End Sub
    Friend Formatted_ As Boolean = False
    Friend ReadOnly Property Formatted As Boolean
        Get
            Return Formatted_
        End Get
    End Property
    Friend Property DataStart As Date
    Friend Property DataEnd As Date
    Friend Property SizeStart As Date
    Friend Property SizeEnd As Date
    Public Sub AutoSize()
        If Parent IsNot Nothing Then Parent.SizeColumn(Me)
    End Sub
    Friend _DataType As Type
    Public ReadOnly Property DataType As Type
        Get
            Return _DataType
        End Get
    End Property
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                _Image.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
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
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
Public Class RowCollection
        Inherits List(Of Row)
        Implements IDisposable
        Public Sub New(Viewer As DataViewer)
            _Parent = Viewer
        End Sub
        Friend _Parent As DataViewer
        Public ReadOnly Property Parent As DataViewer
            Get
                Return _Parent
            End Get
        End Property
        Friend Event Loaded()
        Private WithEvents AddTimer As New Timer With {.Interval = 200}
        Private Sub AddTimerTick() Handles AddTimer.Tick
            AddTimer.Stop()
            RaiseEvent Loaded()
        End Sub
        Private WithEvents HeaderStyle_ As New CellStyle With {.BackColor = Color.Silver, .ShadeColor = Color.Gainsboro, .ForeColor = Color.White, .Font = New Font("Century Gothic", 9)}
        Public Property HeaderStyle As CellStyle
            Get
                Return HeaderStyle_
            End Get
            Set(value As CellStyle)
                If HeaderStyle_ IsNot value Then
                    HeaderStyle_ = value
                    ReBoundsHead(Nothing, Nothing)
                End If
            End Set
        End Property
        Private WithEvents RowStyle_ As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.White, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
        Public Property RowStyle As CellStyle
            Get
                Return RowStyle_
            End Get
            Set(value As CellStyle)
                If RowStyle_ IsNot value Then
                    RowStyle_ = value
                    ReBoundsEven(Nothing, Nothing)
                End If
            End Set
        End Property
        Private WithEvents AlternatingRowStyle_ As New CellStyle With {.BackColor = Color.Silver, .ShadeColor = Color.Lavender, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
        Public Property AlternatingRowStyle As CellStyle
            Get
                Return AlternatingRowStyle_
            End Get
            Set(value As CellStyle)
                If AlternatingRowStyle_ IsNot value Then
                    AlternatingRowStyle_ = value
                    ReBoundsOdd(Nothing, Nothing)
                End If
            End Set
        End Property
        Private WithEvents SelectionRowStyle_ As New CellStyle With {.BackColor = Color.DarkSlateGray, .ShadeColor = Color.Gray, .ForeColor = Color.White, .Font = New Font("Century Gothic", 8)}
        Public Property SelectionRowStyle As CellStyle
            Get
                Return SelectionRowStyle_
            End Get
            Set(value As CellStyle)
                If SelectionRowStyle_ IsNot value Then
                    SelectionRowStyle_ = value
                End If
            End Set
        End Property
        Public ReadOnly Property EvenRowHeight As Integer = 16
        Public ReadOnly Property OddRowHeight As Integer = 16
        Public ReadOnly Property RowHeight As Integer
            Get
                Return {_EvenRowHeight, _OddRowHeight}.Max
            End Get
        End Property
    Public Property SingleSelect As Boolean = True
    Public ReadOnly Property Selected As List(Of Row)
        Get
            Return Where(Function(r) r.Selected).ToList
        End Get
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Shadows Function Add(ByVal NewRow As Row) As Row

            If NewRow IsNot Nothing Then
                NewRow._Parent = Me
                If Count = 0 Then _EvenRowHeight = TextRenderer.MeasureText("XXXXXXXXXXX".ToString(InvariantCulture), RowStyle.Font).Height
                If Count = 1 Then _OddRowHeight = TextRenderer.MeasureText("XXXXXXXXXXX".ToString(InvariantCulture), AlternatingRowStyle.Font).Height
                MyBase.Add(NewRow)
                AddTimer.Start()
            End If
            Return NewRow

        End Function
        Private Sub ReBoundsHead(sender As Object, e As StyleEventArgs) Handles HeaderStyle_.PropertyChanged

        End Sub
        Private Sub ReBoundsEven(sender As Object, e As StyleEventArgs) Handles RowStyle_.PropertyChanged
            _EvenRowHeight = TextRenderer.MeasureText("XXXXXXXXXXX".ToString(InvariantCulture), RowStyle.Font).Height
            Parent.Invalidate()
        End Sub
        Private Sub ReBoundsOdd(sender As Object, e As StyleEventArgs) Handles AlternatingRowStyle_.PropertyChanged
            _OddRowHeight = TextRenderer.MeasureText("XXXXXXXXXXX".ToString(InvariantCulture), AlternatingRowStyle.Font).Height
            Parent.Invalidate()
        End Sub
#Region "IDisposable Support"
        Private DisposedValue As Boolean ' To detect redundant calls IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not DisposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).

                End If
                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            DisposedValue = True
        End Sub
        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        Protected Overrides Sub Finalize()
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
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
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
<Serializable()> Public Class Row
    Implements IDisposable
    Public Sub New(NewDataRow As DataRow)

        If NewDataRow IsNot Nothing Then
            _DataRow = NewDataRow
            Table = NewDataRow.Table
        End If

    End Sub
    <NonSerialized> Friend _Parent As RowCollection
    Public ReadOnly Property Parent As RowCollection
        Get
            Return _Parent
        End Get
    End Property
    Public ReadOnly Property Table As DataTable
    <NonSerialized> Private ReadOnly _DataRow As DataRow
    Public ReadOnly Property DataRow As DataRow
        Get
            Return _DataRow
        End Get
    End Property
    Public Function Cell(Column As Column) As Object

        If Column IsNot Nothing Then
            Try
                If IsDBNull(DataRow(Column.Name)) Then
                    Return Nothing
                Else
                    Return DataRow(Column.Name)
                End If
            Catch ex As KeyNotFoundException
                Return Nothing
            End Try
        Else
            Return Nothing
        End If

    End Function
    Public Function Cell(Column As String) As Object

        If DataRow.Table.Columns.Contains(Column) Then
            If IsDBNull(DataRow(Column)) Then
                Return Nothing
            Else
                Return DataRow(Column)
            End If
        Else
            Return Nothing
        End If

    End Function
    Public Function Cell(ColumnIndex As Integer) As Object

        If ColumnIndex < Table.Columns.Count Then
            Return Cell(Table.Columns(ColumnIndex).ColumnName)
        Else
            Return Nothing
        End If

    End Function
    Private _Selected As Boolean
    Public Property Selected As Boolean
        Get
            Return _Selected
        End Get
        Set(value As Boolean)
            If _Selected <> value Then
                If value And Parent.SingleSelect And Parent.Where(Function(r) r.Selected).Any Then
                    For Each Row In Parent.Except({Me}).Where(Function(r) r.Selected)
                        Row.Selected = False
                    Next
                End If
                _Selected = value
                Parent.Parent.Invalidate()
            End If
        End Set
    End Property
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                Table.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
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
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
Public NotInheritable Class StyleEventArgs
    Inherits EventArgs
    Public ReadOnly Property PropertyName As String
    Public Sub New(value As String)
        PropertyName = value
    End Sub
End Class
<Serializable()> <TypeConverter(GetType(PropertyConverter))> Public Class CellStyle
    Implements IDisposable
    Public Event PropertyChanged(sender As Object, e As StyleEventArgs)
    Public Sub New()
    End Sub
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Color")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Private _BackColor As Color = Color.Gainsboro
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(ByVal value As Color)
            If value <> _BackColor Then
                _BackColor = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs("BackColor"))
            End If
        End Set
    End Property
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Appearance")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Private _Font As New Font("Century Gothic", 9)
    Public Property Font As Font
        Get
            Return _Font
        End Get
        Set(ByVal value As Font)
            If value IsNot _Font Then
                _Font = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs("Font"))
            End If
        End Set
    End Property
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Color")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Private _ForeColor As Color = Color.Black
    Public Property ForeColor As Color
        Get
            Return _ForeColor
        End Get
        Set(ByVal value As Color)
            If value <> _ForeColor Then
                _ForeColor = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs("ForeColor"))
            End If
        End Set
    End Property
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Color")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Private _ShadeColor As Color = Color.WhiteSmoke
    Public Property ShadeColor As Color
        Get
            Return _ShadeColor
        End Get
        Set(ByVal value As Color)
            If value <> _ShadeColor Then
                _ShadeColor = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs("ShadeColor"))
            End If
        End Set
    End Property
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Layout")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Private _Alignment As ContentAlignment = ContentAlignment.MiddleCenter
    Public Property Alignment As ContentAlignment
        Get
            Return _Alignment
        End Get
        Set(ByVal value As ContentAlignment)
            If value <> _Alignment Then
                _Alignment = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs("Alignment"))
            End If
        End Set
    End Property
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Layout")>
    <Description("Specifies the object Type")>
    <RefreshProperties(RefreshProperties.All)>
    Private _Padding As Padding = New Padding(0)
    Public Property Padding As Padding
        Get
            Return _Padding
        End Get
        Set(ByVal value As Padding)
            If value <> _Padding Then
                _Padding = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs("Padding"))
            End If
        End Set
    End Property
    Public Shadows Function ToString() As String
        Return Join({BackColor.ToString, ForeColor.ToString, Font.ToString, ShadeColor.ToString}, Delimiter)
    End Function
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                _Font.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        DisposedValue = True
    End Sub
    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
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