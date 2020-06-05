Option Strict On
Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel
Imports System.Drawing.Drawing2D
#Region " STRUCTURES + ENUMERATIONS "
Public Enum MouseRegion
    Header
    Grid
End Enum
Public Structure MouseInfo
    Implements IEquatable(Of MouseInfo)
    Public Property Column As Column
    Public Property Row As Row
    Public Property RowBounds As Rectangle
    Public Property Cell As Cell
    Public Property CellBounds As Rectangle
    Public Property Point As Point
    Public Property SelectPointA As Point
    Public Property SelectPointB As Point
    Public Property CurrentAction As Action
    Public Property CurrentRegion As MouseRegion
    Public Enum Action
        None
        MouseOverHead
        MouseOverGrid
        GridSelecting
        MouseOverHeadEdge
        HeaderEdgeClicked
        HeaderClicked
        ColumnDragging
        ColumnSizing
        Filtering
        CellClicked
        CellDoubleClicked
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
    Public ReadOnly Property Table As DataTable
    Public Sub New(MI As MouseInfo)
        MouseData = MI
    End Sub
    Public Sub New(Table As DataTable)
        Me.Table = Table
    End Sub
End Class
Public Class DataViewer
    Inherits Control
#Region " GENERAL DECLARATIONS "
    Private WithEvents BindingSource As New BindingSource
    Private ReadOnly GothicFont As Font = My.Settings.applicationFont
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ C O N T R O L S
    Private WithEvents SaveFile As New SaveFileDialog
    Private WithEvents CopyTimer As New Timer With {.Interval = 3000, .Tag = 0}
    Private WithEvents RowTimer As New Timer With {.Interval = 250}
    Public WithEvents VScroll As New VScrollBar With {.Minimum = 0}
    Public WithEvents HScroll As New HScrollBar With {.Minimum = 0}
    Private WithEvents HeaderOptions As New ToolStripDropDown With {.AutoClose = False}
    Public WithEvents GridOptions As New ContextMenuStrip With {.AutoClose = False}
    Private WithEvents HeaderGridAlignment As New ImageCombo With {
        .Margin = New Padding(0),
        .DataSource = EnumNames(GetType(ContentAlignment)),
        .Dock = DockStyle.Fill
    }
    Private WithEvents HeaderDistinctItems As New ImageCombo With {
        .Margin = New Padding(0),
        .Dock = DockStyle.Fill
    }
    Private WithEvents GridBackColor As New ImageCombo With {.Mode = ImageComboMode.ColorPicker,
        .Size = New Size(200, 24)}
    Private WithEvents GridForeColor As New ImageCombo With {.Mode = ImageComboMode.ColorPicker,
        .Size = New Size(200, 24)}
    Private ReadOnly ColumnHeadTip As ToolTip = New ToolTip With {
        .BackColor = Color.Black,
        .ForeColor = Color.White
    }
    Private ReadOnly Grid_FileExport As New ToolStripMenuItem With {.Text = "File",
        .Image = My.Resources.Folder,
        .Font = GothicFont}
    Private ReadOnly Grid_csvExport As New ToolStripMenuItem With {.Text = ".csv",
        .Image = My.Resources.csv,
        .Font = GothicFont}
    Private ReadOnly Grid_txtExport As New ToolStripMenuItem With {.Text = ".txt",
        .Image = My.Resources.txt,
        .Font = GothicFont}
    Private ReadOnly Grid_ExcelExport As New ToolStripMenuItem With {.Text = "Excel",
        .Image = My.Resources.Excel,
        .Font = GothicFont}
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ E L S E
    Private ReadOnly GroupedProperties As New Dictionary(Of String, List(Of System.Configuration.SettingsPropertyValue))
    Private ControlKeyDown As Boolean
#End Region
#Region " EVENTS "
    Public Event ColumnsSized(sender As Object, e As ViewerEventArgs)
    Public Event RowsLoading(sender As Object, e As ViewerEventArgs)
    Public Event RowsLoaded(sender As Object, e As ViewerEventArgs)
    Public Event RowMouseChanged(sender As Object, e As ViewerEventArgs)
    Public Event RowClicked(sender As Object, e As ViewerEventArgs)
    Public Event CellMouseChanged(sender As Object, e As ViewerEventArgs)
    Public Event CellClicked(sender As Object, e As ViewerEventArgs)
    Public Event CellDoubleClicked(sender As Object, e As ViewerEventArgs)
    Public Event Alert(sender As Object, e As AlertEventArgs)
#End Region
#Region "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ INITIALIZE "
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

        Dim colorFont As New Font("Century Gothic", 9)
        HeaderGridAlignment.Font = colorFont
        HeaderDistinctItems.Font = colorFont

        For Each tsmiExport As ToolStripMenuItem In {Grid_ExcelExport, Grid_csvExport, Grid_txtExport}
            tsmiExport.ImageScaling = ToolStripItemImageScaling.None
            Grid_FileExport.DropDownItems.Add(tsmiExport)
            AddHandler tsmiExport.Click, AddressOf ExportToFile
        Next
        GridOptions.Items.Add(Grid_FileExport)

        For Each setting In Settings
            Dim settingMatches = RegexMatches(setting.Name, "^[a-z]{1,}(?=[A-Z])", System.Text.RegularExpressions.RegexOptions.None)
            If settingMatches.Any Then
                Dim controlName As String = settingMatches.First.Value
                If Not GroupedProperties.ContainsKey(controlName) Then GroupedProperties.Add(controlName, New List(Of System.Configuration.SettingsPropertyValue))
                GroupedProperties(controlName).Add(setting)
            End If
        Next
        If GroupedProperties.Any Then
            HeaderOptions.Items.Add(FontsColorsToTSMI(MouseRegion.Header))
            GridOptions.Items.Add(FontsColorsToTSMI(MouseRegion.Grid))
        End If

    End Sub
    Private Function FontsColorsToTSMI(viewerRegion As MouseRegion) As ToolStripMenuItem

        Dim viewerProperties = GroupedProperties("grid")
        Dim properties As New List(Of System.Configuration.SettingsPropertyValue)
        Dim objectProperties As New Dictionary(Of String, List(Of System.Configuration.SettingsPropertyValue))
        For Each viewerProperty In viewerProperties
            Dim subPropertyName As String = viewerProperty.Name.Remove(0, "grid".Length)
            Dim startsWith As String = If(viewerRegion = MouseRegion.Grid, "row", If(viewerRegion = MouseRegion.Header, "header", "XXX"))
            If subPropertyName.StartsWith(startsWith, StringComparison.InvariantCultureIgnoreCase) Then
                properties.Add(viewerProperty)
                subPropertyName = subPropertyName.Remove(0, startsWith.Length)
                Dim colorProperty As String() = System.Text.RegularExpressions.Regex.Split(subPropertyName, "(?=ForeColor|BackColor|ShadeColor)", System.Text.RegularExpressions.RegexOptions.None)
                Dim colorGroupKey As String = colorProperty.First
                Dim colorGroupValue As String = colorProperty.Last
                If Not objectProperties.ContainsKey(colorGroupKey) Then objectProperties.Add(colorGroupKey, New List(Of System.Configuration.SettingsPropertyValue))
                objectProperties(colorGroupKey).Add(viewerProperty)
            End If
        Next
        Dim outDictionary As New Dictionary(Of String, Integer) From {
            {"", 0},
            {"Alternating", 1},
            {"Selection", 2}
        }
        Dim inDictionary As New Dictionary(Of String, Integer) From {
            {"BackColor", 0},
            {"ShadeColor", 1},
            {"ForeColor", 2}
        }

        Dim tlpOutside As New TableLayoutPanel With {
                .Font = New Font("Century Gothic", 9),
                .Margin = New Padding(0),
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
                .ColumnCount = 2,
                .RowCount = objectProperties.Count
                }
        With tlpOutside
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 200})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 200})
        End With
        Dim rowIndexOutside As Integer
        For Each objectProperty In objectProperties.OrderBy(Function(op) outDictionary(op.Key)) '... Header="", Row={"", Alternating, Selection}
            Dim groupName As String = If(viewerRegion = MouseRegion.Header, "Header", If(objectProperty.Key.Any, objectProperty.Key, "Row"))
            Dim tlpInside As New TableLayoutPanel With {
                .Font = New Font("Century Gothic", 9),
                .Margin = New Padding(0),
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                .ColumnCount = 1,
                .RowCount = objectProperties.Count
                }
            tlpInside.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 200})
            Dim rowIndexInside As Integer = 0
            For Each subProperty In objectProperty.Value.OrderBy(Function(sp) inDictionary(System.Text.RegularExpressions.Regex.Match(sp.Name, "(Back|Shade|Fore)Color", System.Text.RegularExpressions.RegexOptions.None).Value))
                Dim rowColor As Color
                If subProperty.PropertyValue Is Nothing Then
#Region " IDEAL DATAVIEWER CELLSTYLE PROPERTIES "
                    Dim defaultHeaderStyle As New CellStyle With {.BackColor = Color.Silver, .ShadeColor = Color.Gainsboro, .ForeColor = Color.White, .Font = New Font("Century Gothic", 9)}
                    Dim defaultRowStyle As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.White, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
                    Dim defaultAlternatingStyle As New CellStyle With {.BackColor = Color.Silver, .ShadeColor = Color.Lavender, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
                    Dim defaultSelectionStyle As New CellStyle With {.BackColor = Color.DarkSlateGray, .ShadeColor = Color.Gray, .ForeColor = Color.White, .Font = New Font("Century Gothic", 8)}
#End Region
                    Dim defaultStyle As CellStyle = If(viewerRegion = MouseRegion.Header, defaultHeaderStyle,
                        If(groupName = "Row", defaultRowStyle, If(groupName = "Selection", defaultSelectionStyle, defaultAlternatingStyle)))
                    Dim colorIndex As Integer = If(subProperty.Name.Contains("BackColor"), 0, If(subProperty.Name.Contains("ForeColor"), 1, 2))
                    rowColor = If(colorIndex = 0, defaultStyle.BackColor, If(colorIndex = 1, defaultStyle.ForeColor, defaultStyle.ShadeColor))
                Else
                    rowColor = DirectCast(subProperty.PropertyValue, Color)
                End If
                Dim colorControl As New ImageCombo With {
                    .Margin = New Padding(0),
                    .Dock = DockStyle.Fill,
                    .Mode = ImageComboMode.ColorPicker,
                    .HintText = If(subProperty.Name.Contains("BackColor"), "Background", If(subProperty.Name.Contains("ForeColor"), "Text", "Gradient")) & " color",
                    .Text = rowColor.Name,
                    .SelectedIndex = .TextIndex,
                    .Tag = subProperty
                    }
                AddHandler colorControl.SelectionChanged, AddressOf CellStyleProperty_SelectionChanged
                With tlpInside
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 26})
                    .Controls.Add(colorControl, 0, rowIndexInside)
                End With
                rowIndexInside += 1
            Next
            If viewerRegion = MouseRegion.Header Then
                With tlpInside
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 26})
                    .Controls.Add(HeaderGridAlignment, 0, rowIndexInside)
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 26})
                    .Controls.Add(HeaderDistinctItems, 0, rowIndexInside + 1)
                End With
            End If
            TLP.SetSize(tlpInside)
            Dim subButton As New ImageCombo With {
                .Dock = DockStyle.Fill,
                .Mode = ImageComboMode.Button,
                .Text = groupName
            }
            With tlpOutside
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = tlpInside.Height})
                .Controls.Add(subButton, 0, rowIndexOutside)
                .Controls.Add(tlpInside, 1, rowIndexOutside)
            End With
            TLP.SetSize(tlpOutside)
            rowIndexOutside += 1
        Next
        Dim tsmiFontsColors As New ToolStripMenuItem(viewerRegion.ToString & " fonts and colors".ToString(InvariantCulture)) With {
            .Font = Gothic,
            .AutoSize = True,
            .Name = "properties",
            .Image = My.Resources.circles
        }
        tsmiFontsColors.DropDownItems.Add(New ToolStripControlHost(tlpOutside) With {.AutoSize = True, .Name = "properties"})
        Return tsmiFontsColors

    End Function
#End Region
#Region "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ DRAWING "
    'HEADER / ROW PROPERTIES...USE A TEMPLATE LIKE MS AND APPLY TO CELL, ROW, HEADER...{V/H ALIGNMENT, FONT, FORCOLOR, BACKCOLOR, ETC}
    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)

        SetupScrolls()
        If e IsNot Nothing Then
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            e.Graphics.FillRectangle(New SolidBrush(BackColor), ClientRectangle)
            If BackgroundImage IsNot Nothing Then e.Graphics.DrawImage(BackgroundImage, CenterItem(BackgroundImage.Size))
#Region " DRAW HEADERS "
            With Columns
                Dim HeadFullBounds As New Rectangle(-HScroll.Value, 0, {1, .HeadBounds.Width}.Max, .HeadBounds.Height)
                If .HeaderStyle.Theme = Theme.None Then
                    Using LinearBrush As New LinearGradientBrush(HeadFullBounds, .HeaderStyle.BackColor, .HeaderStyle.ShadeColor, LinearGradientMode.Vertical)
                        e.Graphics.FillRectangle(LinearBrush, HeadFullBounds)
                    End Using
                Else
                    e.Graphics.DrawImage(GlossyImages(.HeaderStyle.Theme), HeadFullBounds)
                End If
                If .Any Then
                    ControlPaint.DrawBorder3D(e.Graphics, HeadFullBounds, Border3DStyle.Sunken)
                    Dim ColumnStart As Integer = ColumnIndex(HScroll.Value)
                    VisibleColumns.Clear()
                    Try
                        For ColumnIndex As Integer = ColumnStart To {Columns.Count - 1, ColumnStart + 30}.Min
                            Dim Column As Column = Columns(ColumnIndex)
                            With Column
                                If .Visible Then
                                    Dim HeadBounds As Rectangle = .HeadBounds
                                    HeadBounds.Offset(-HScroll.Value, 0)
                                    If HeadBounds.Left >= Width Then Exit For
                                    VisibleColumns.Add(Column, HeadBounds)
                                    If .HeaderStyle.Theme = Theme.None Then
                                        Using LinearBrush As New LinearGradientBrush(HeadBounds, .HeaderStyle.BackColor, .HeaderStyle.ShadeColor, LinearGradientMode.Vertical)
                                            e.Graphics.FillRectangle(LinearBrush, HeadBounds)
                                        End Using
                                    Else
                                        e.Graphics.DrawImage(GlossyImages(.HeaderStyle.Theme), HeadBounds)
                                    End If
                                    e.Graphics.DrawRectangle(Pens.Silver, HeadBounds)
#Region " [0] DRAW HEADER IMAGE "
                                    Dim imageSize As Size = .SizeImage
                                    Dim imageTop As Integer = CInt((HeadBounds.Height - imageSize.Height) / 2)
                                    Dim ImageBounds As New Rectangle(New Point(HeadBounds.Left + If(imageSize.Width = 0, 0, 2), imageTop), imageSize)
                                    If .Image IsNot Nothing Then
                                        e.Graphics.DrawImage(.Image, ImageBounds)
                                        'e.Graphics.DrawRectangle(Pens.Yellow, ImageBounds)
                                    End If
#End Region
#Region " [3] DRAW SORT TRIANGLE "
                                    Dim sortSize As Size = .SizeSort
                                    Dim sortTop As Integer = CInt((HeadBounds.Height - sortSize.Height) / 2)
                                    Dim sortBounds As New Rectangle(HeadBounds.Right - (.SizeSort.Width + If(sortSize.Width = 0, 0, 4)), sortTop, .SizeSort.Width, .SizeSort.Height) '4 is 2+Sort+2
                                    If Not .SortOrder = SortOrder.None Then e.Graphics.DrawImage(If(.SortOrder = SortOrder.Ascending, My.Resources.SortDown, My.Resources.SortUp), sortBounds)
#End Region
#Region " [2] DRAW FILTER "
                                    Dim filterSize As Size = .SizeFilter
                                    Dim filterTop As Integer = CInt((HeadBounds.Height - filterSize.Height) / 2)
                                    Dim filterBounds As New Rectangle(sortBounds.Left - filterSize.Width, filterTop, filterSize.Width, filterSize.Height)
                                    If .Filtered Then e.Graphics.DrawImage(My.Resources.FilterCancel, filterBounds)
#End Region
#Region " [1] DRAW HEADER TEXT "
                                    Dim textLeft As Integer = ImageBounds.Right + If(ImageBounds.Width = 0, 0, 2)
                                    Dim textTop As Integer = CInt((HeadBounds.Height - .SizeText.Height) / 2)
                                    Dim TextBounds As Rectangle = New Rectangle(textLeft, textTop, filterBounds.Left - textLeft, .SizeText.Height)
                                    TextRenderer.DrawText(e.Graphics, .Text, .HeaderStyle.Font, TextBounds, .HeaderStyle.ForeColor, Color.Transparent, TextFormatFlags.VerticalCenter Or TextFormatFlags.HorizontalCenter)
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
                    Catch ex As IndexOutOfRangeException
                    End Try
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
                    Dim tipCells As New Dictionary(Of Rectangle, String)
#Region " DRAW ROWS "
                    If Rows.Any Then
                        Dim Top As Integer = HeaderHeight
                        Dim RowStart As Integer = RowIndex(VScroll.Value)
                        Dim drawBounds As Rectangle = PointsToRectangle(MouseData.SelectPointA, MouseData.SelectPointB)
                        VisibleRows.Clear()
                        For RowIndex As Integer = RowStart To {RowStart + VisibleRowCount, Rows.Count - 1}.Min
                            Dim Row = Rows(RowIndex)
                            Dim MouseOverRow As Boolean = _MouseData.Row Is Row And _MouseData.CurrentAction = MouseInfo.Action.MouseOverGrid
                            With Row
                                Dim RowBounds = New Rectangle(0, Top, HeadFullBounds.Width, RowHeight)
                                RowBounds.Offset(-HScroll.Value, 0)
                                VisibleRows.Add(Row, RowBounds)
                                If RowBounds.Top >= Bottom Then Exit For
                                Dim rowStyle As CellStyle = If(.StyleChanged, .Style, If(.Selected, Rows.SelectionRowStyle, If(RowIndex Mod 2 = 0, Rows.RowStyle, Rows.AlternatingRowStyle)))
                                If rowStyle.Theme = Theme.None Then
                                    Using LinearBrush As New LinearGradientBrush(RowBounds, rowStyle.BackColor, rowStyle.ShadeColor, LinearGradientMode.Vertical)
                                        e.Graphics.FillRectangle(LinearBrush, RowBounds)
                                    End Using
                                Else
                                    e.Graphics.DrawImage(GlossyImages(rowStyle.Theme), RowBounds)
                                End If
#Region " DRAW CELLS "
                                For Each Column In VisibleColumns.Keys
                                    With Column
                                        Dim CellBounds As New Rectangle(.HeadBounds.Left, RowBounds.Top, .HeadBounds.Width, RowBounds.Height)
                                        CellBounds.Offset(-HScroll.Value, 0)
                                        Dim rowCell As Cell = Row.Cells(Column.Name)
                                        Dim MouseOverCell As Boolean = _MouseData.Cell Is rowCell And _MouseData.CurrentAction = MouseInfo.Action.MouseOverGrid
                                        With rowCell
                                            Dim cellValue As Object = .Value
                                            Dim drawCellAsSelected As Boolean = If(FullRowSelect, Row.Selected, .Selected)
                                            Dim cellStyle = If(drawCellAsSelected, Row.SelectionStyle, rowStyle)
                                            If MouseData.CurrentAction = MouseInfo.Action.GridSelecting Then .Selected = drawBounds.IntersectsWith(CellBounds)
#Region " C E L L   B A C K G R O U N D "
                                            If drawCellAsSelected Then   'Already drew the entire row before "DRAW CELLS" Region
                                                If cellStyle.Theme = Theme.None Then
                                                    Using LinearBrush As New LinearGradientBrush(CellBounds, cellStyle.BackColor, cellStyle.ShadeColor, LinearGradientMode.Vertical)
                                                        e.Graphics.FillRectangle(LinearBrush, CellBounds)
                                                    End Using
                                                Else
                                                    e.Graphics.DrawImage(GlossyImages(Row.SelectionStyle.Theme), CellBounds)
                                                End If
                                            End If
#End Region
                                            If cellValue Is Nothing Or .Text = Date.MinValue.ToShortDateString Then
#Region " C E L L   N U L L "
                                                Using NullBrush As New SolidBrush(Color.FromArgb(128, Color.Gainsboro))
                                                    e.Graphics.FillRectangle(NullBrush, CellBounds)
                                                End Using
                                                Using textBrush As New SolidBrush(cellStyle.ForeColor)
                                                    e.Graphics.DrawString("(null)",
                                                                          If(MouseOverRow, New Font(cellStyle.Font, FontStyle.Underline), cellStyle.Font),
                                                                          textBrush,
                                                                          CellBounds,
                                                                          Column.GridStyle.Alignment)
                                                End Using
#End Region
                                            Else
                                                '///////////  KNOWN ISSUES
                                                '///////////  CHANGING A NON-IMAGE VALUE TO AN IMAGE WILL SWITCH THE COLUMN FORMAT TO IMAGES WHICH THROWS AN ERROR @ Dim ImageWidth As Integer
                                                '///////////  ALSO FILLING A DATATABLE WITH IMAGES BEFORE SETTING TO THE DATAVIEWER.DATASOURCE WON'T SET THE ROWHEIGHT CORRECTLY

                                                Dim cellText As String = .Text 'Getting .Text also Sets the .Image based on the Column.DataType
                                                If .Column.Format.Key = Column.TypeGroup.Images Or .Column.Format.Key = Column.TypeGroup.Booleans Then
                                                    Dim EdgePadding As Integer = 1 'all sides to ensure Image doesn't touch the edge of the Cell Rectangle
                                                    Dim MaxImageWidth As Integer = CellBounds.Width - EdgePadding * 2
                                                    Dim MaxImageHeight As Integer = CellBounds.Height - EdgePadding * 2
                                                    Dim ImageWidth As Integer = { .ValueImage.Width, MaxImageWidth}.Min
                                                    Dim ImageHeight As Integer = { .ValueImage.Height, MaxImageHeight}.Min
                                                    Dim xOffset As Integer = CInt((CellBounds.Width - ImageWidth) / 2)
                                                    Dim yOffset As Integer = CInt((CellBounds.Height - ImageHeight) / 2)
                                                    Dim imageBounds As New Rectangle(CellBounds.X + xOffset, CellBounds.Y + yOffset, ImageWidth, ImageHeight)
                                                    If .Column.Format.Key = Column.TypeGroup.Booleans Then
                                                        Dim useWhite As Boolean = BackColorToForeColor(cellStyle.BackColor) = Color.White
                                                        Dim imageName As String = If(.ValueBoolean, String.Empty, "un") & "checked" & If(useWhite, "White", "Black")
                                                        e.Graphics.DrawImage(Base64ToImage(CheckImages(imageName)), imageBounds)
                                                    Else
                                                        e.Graphics.DrawImage(.ValueImage, imageBounds)
                                                    End If
                                                    If MouseOverRow Then
                                                        Using yellowBrush As New SolidBrush(Color.FromArgb(128, Color.Yellow))
                                                            e.Graphics.FillRectangle(yellowBrush, imageBounds)
                                                        End Using
                                                    End If
                                                Else
                                                    Using textBrush As New SolidBrush(cellStyle.ForeColor)
                                                        e.Graphics.DrawString(.Text,
                                                                              If(MouseOverRow, New Font(cellStyle.Font, FontStyle.Underline), cellStyle.Font),
                                                                              textBrush,
                                                                              CellBounds,
                                                                              Column.GridStyle.Alignment)
                                                    End Using
                                                End If
                                            End If
                                            If .TipText IsNot Nothing Then
                                                Dim triangleHeight As Single = 8
                                                Dim trianglePoints As New List(Of PointF) From {New PointF(CellBounds.Right - triangleHeight, CellBounds.Top),
                                        New PointF(CellBounds.Right, CellBounds.Top),
                                        New PointF(CellBounds.Right, CellBounds.Top + triangleHeight)}
                                                e.Graphics.FillPolygon(Brushes.DarkOrange, trianglePoints.ToArray)
                                                If MouseOverCell Then tipCells.Add(CellBounds, .TipText)
                                            End If
                                        End With
                                        ControlPaint.DrawBorder3D(e.Graphics, CellBounds, Border3DStyle.SunkenOuter)
                                    End With
                                Next
#End Region
                                Top += RowHeight
                            End With
                        Next
                    End If
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
#Region " OVERLAYS "
                        For Each tipCell In tipCells
                            Dim tipMessage As String = tipCell.Value
                            Using tipFont = New Font(Font.FontFamily, 15, FontStyle.Regular)
                                Dim tipSize As SizeF = e.Graphics.MeasureString(tipMessage, tipFont, StringTrimming.None)
                                Dim tipRectangle As New Rectangle(New Point(tipCell.Key.Right + 28, tipCell.Key.Top - 20), tipSize.ToSize)
                                tipRectangle.Inflate(8, 8)
                                Using copyPath As GraphicsPath = DrawSpeechBubble(tipRectangle)
                                    Using backBrush As New SolidBrush(Color.FromArgb(200, Color.GhostWhite))
                                        e.Graphics.FillPath(backBrush, copyPath)
                                        Using copyPen As New Pen(Brushes.DarkGray, 2)
                                            e.Graphics.DrawPath(copyPen, copyPath)
                                        End Using
                                    End Using
                                End Using
                                Dim textAlignment As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center, .FormatFlags = StringFormatFlags.NoWrap}
                                e.Graphics.DrawString(tipMessage, tipFont, Brushes.Black, tipRectangle, textAlignment)
                            End Using
                        Next
                        Dim copyValue As Integer = DirectCast(CopyTimer.Tag, Integer)
                        If Math.Abs(copyValue) > 0 Then
                            Dim imageOffsetXY As Integer = 3
                            Dim imageWH As Integer = My.Resources.Copied.Width
                            Dim bannerWidth As Integer
                            Dim copyMessage As String = "Copied"
                            If copyValue < 0 Then
                                copyMessage = Join({copyMessage, "value", Clipboard.GetText})
                            Else
                                copyMessage = Join({copyMessage, copyValue, "row" & If(copyValue = 1, String.Empty, "s")})
                            End If
                            If copyMessage.Length >= 30 Then copyMessage = copyMessage.Substring(0, 30) & "..."
                            Using messageFont = New Font(Font.FontFamily, 15, FontStyle.Bold)
                                bannerWidth = imageOffsetXY + imageWH + CInt(e.Graphics.MeasureString(copyMessage, messageFont, StringTrimming.None).Width) + imageOffsetXY
                                Dim bannerSize As New Size(bannerWidth, imageOffsetXY + imageWH + imageOffsetXY)
                                Dim bannerRectangle As New Rectangle(New Point(imageWH, imageWH), bannerSize)
                                Dim imageRectangle As New Rectangle(bannerRectangle.Left + imageOffsetXY, bannerRectangle.Top + imageOffsetXY, imageWH, imageWH)
                                Dim textRectangle As New Rectangle(imageRectangle.Right,
                                                                   bannerRectangle.Top,
                                                                   bannerRectangle.Width - imageOffsetXY - imageRectangle.Width,
                                                                   bannerRectangle.Height)
                                Using copyPath As GraphicsPath = DrawRoundedRectangle(bannerRectangle, 22)
                                    Using copyPen As New Pen(Brushes.DarkGray, 2)
                                        e.Graphics.DrawPath(copyPen, copyPath)
                                    End Using
                                    Using copyBrush As New SolidBrush(Color.FromArgb(208, Color.GhostWhite))
                                        e.Graphics.FillPath(copyBrush, copyPath)
                                    End Using
                                End Using
                                e.Graphics.DrawImage(My.Resources.Copied, imageRectangle)
                                Dim textAlignment As New StringFormat With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Center, .FormatFlags = StringFormatFlags.NoWrap}
                                e.Graphics.DrawString(copyMessage, messageFont, Brushes.Black, textRectangle, textAlignment)
                            End Using
                        End If
#End Region
                    End If
                Else
                    TextRenderer.DrawText(e.Graphics,
                                          Name,
                                          Font,
                                          HeadFullBounds,
                                          .HeaderStyle.ForeColor,
                                          Color.Transparent,
                                          TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                    ControlPaint.DrawBorder3D(e.Graphics, HeadFullBounds, Border3DStyle.Sunken)
                End If
            End With
            ControlPaint.DrawBorder3D(e.Graphics, ClientRectangle, Border3DStyle.Sunken)
        End If

    End Sub
    Private Sub CopyTimer_Tick(sender As Object, e As EventArgs) Handles CopyTimer.Tick
        CopyTimer.Stop()
        CopyTimer.Tag = 0
        Invalidate()
    End Sub
#End Region
    Private Sub SetupScrolls()

        VScroll.Maximum = {HeaderHeight + RowHeight + TotalSize.Height - 1, 0}.Max
        HScroll.Maximum = {TotalSize.Width - 1, 0}.Max
        VScroll.Visible = VScrollVisible
        If VScrollVisible Then
            With VScroll
                .Top = 2
                .Left = {TotalSize.Width, ClientSize.Width - .Width, Columns.HeadBounds.Right}.Min
                .Height = ClientRectangle.Height - 2
                .SmallChange = Rows.RowHeight
                .LargeChange = ClientRectangle.Height
            End With
        End If
        HScroll.Visible = HScrollVisible
        If HScrollVisible Then
            With HScroll
                .Top = ClientRectangle.Bottom - HScroll.Height
                .Left = 0
                .Width = If(VScroll.Visible, ClientRectangle.Width - VScroll.Width, ClientRectangle.Width)
                .LargeChange = ClientRectangle.Width
            End With
        End If

    End Sub
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Private AutoSize_ As Boolean = False
    Public Overloads Property AutoSize As Boolean
        Get
            Return AutoSize_
        End Get
        Set(value As Boolean)
            If value <> AutoSize_ Then
                AutoSize_ = value
                If value Then
                    Columns.ColumnWidths()
                    Columns.DistibuteWidths()
                    Size = IdealSize
                End If
            End If
        End Set
    End Property
    Public ReadOnly Property IdealSize As Size
        Get
            Dim auto_Size As Size = TotalSize
            Return New Size(auto_Size.Width, auto_Size.Height + HeaderHeight)
        End Get
    End Property
    Public ReadOnly Property TotalSize As Size
        Get
            Dim totalWidth As Integer = Columns.Select(Function(c) c.Width).Sum
            Dim totalHeight As Integer = Rows.Count * Rows.RowHeight
            If totalWidth > ClientRectangle.Width Then totalHeight += HScroll.Height 'Make room for Horizontal Scroll Bar to fully show rows
            If totalHeight > ClientRectangle.Height Then totalWidth += VScroll.Width 'Make room for Vertical Scroll Bar to fully show columns
            Return New Size(totalWidth, totalHeight)
        End Get
    End Property
    Private Timer_ As WaitTimer
    Public ReadOnly Property Timer As WaitTimer
        Get
            Return Timer_
        End Get
    End Property
    Private BaseForm_ As Form
    Public Property BaseForm As Form
        Get
            Return BaseForm_
        End Get
        Set(value As Form)
            BaseForm_ = value
            Timer_ = New WaitTimer(Me, value)
        End Set
    End Property
    Public ReadOnly Property MouseData As New MouseInfo
    Public Property FullRowSelect As Boolean
    Public ReadOnly Property LoadTime As TimeSpan
    Private WithEvents Table_ As DataTable
    Public ReadOnly Property Table As DataTable
        Get
            Return Table_
        End Get
    End Property
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
    Public ReadOnly Property VisibleRowCount As Integer
        Get
            Return CInt(Math.Ceiling(Height - HeaderHeight) / RowHeight)
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
    Friend ReadOnly Property HScrollVisible As Boolean
        Get
            Return TotalSize.Width > ClientRectangle.Width
        End Get
    End Property
    Friend ReadOnly Property DistinctValues As New Dictionary(Of String, Dictionary(Of String, List(Of Cell)))
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
                Table_ = New DataTable
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
                                    Table_.Columns.Add(New DataColumn With {.ColumnName = "Column" & Column, .DataType = GetType(String)})
                                Next
                                For Each Row As String() In Rows
                                    Table_.Rows.Add(Row)
                                Next
                            End If
#End Region
                        Case GetType(List(Of Object()))
#Region " LIST OF OBJECT() - LIST ITEMS=ROWS...STRING()=COLUMNS "
                            Dim Rows As List(Of Object()) = DirectCast(_DataSource, List(Of Object()))
                            If Rows IsNot Nothing Then
                                Dim ColumnCount As Integer = (From C In Rows Select C.Count).Max
                                For Column As Integer = 1 To ColumnCount
                                    Table_.Columns.Add(New DataColumn With {.ColumnName = "Column" & Column, .DataType = GetType(String)})
                                Next
                                For Each Row As String() In Rows
                                    Table_.Rows.Add(Row)
                                Next
                            End If
#End Region
                        Case Else

                    End Select
#End Region
                ElseIf DataSource.GetType Is GetType(DataTable) Then
                    Table_ = DirectCast(DataSource, DataTable)

                End If
#End Region
                Dim startLoad As Date = Now
                For Each DataColumn As DataColumn In Table_.Columns
                    DistinctValues.Add(DataColumn.ColumnName, New Dictionary(Of String, List(Of Cell)))
                    Dim NewColumn = Columns.Add(New Column(DataColumn))
                    Columns.ColumnWidth(NewColumn)
                Next
                If DistinctValues.Any Then
                    If Table_.AsEnumerable.Any Then
                        RaiseEvent RowsLoading(Me, New ViewerEventArgs(Table_))
                        For Each row As DataRow In Table.Rows
                            Rows.Add(New Row(DistinctValues.Keys.ToList, row.ItemArray))
                        Next
                        _LoadTime = Now.Subtract(startLoad)
                        RaiseEvent RowsLoaded(Me, New ViewerEventArgs(Table_))
                        Columns.FormatSize()
                    End If
                Else
                    _LoadTime = Now.Subtract(startLoad)
                End If
            End If
        End Set
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub TableRowAdd(sender As Object, e As DataRowChangeEventArgs) Handles Table_.RowChanged

        If e.Action = DataRowAction.Add Then
            Dim columnNames As New List(Of String)(From c In e.Row.Table.Columns Select DirectCast(c, DataColumn).ColumnName)
            Rows.Add(New Row(columnNames, e.Row.ItemArray))
            RowTimer.Start()
        End If

    End Sub
    Private Sub RowTimer_Tick() Handles RowTimer.Tick
        RowTimer.Stop()
        Columns.ColumnWidths()
        Invalidate()
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub ColumnsSizingStart() Handles Columns_.CollectionSizingStart
        Timer?.StartTicking(Color.LimeGreen)
    End Sub
    Private Sub ColumnSized(sender As Object, e As EventArgs) Handles Columns_.ColumnSized

        With DirectCast(sender, Column)
            RaiseEvent Alert(sender, New AlertEventArgs(Join({"Column", .Name, "Index", .ViewIndex, "resized"})))
        End With

    End Sub
    Private Sub ColumnsSizingEnd() Handles Columns_.CollectionSizingEnd

        RaiseEvent ColumnsSized(Me, Nothing)
        Invalidate()
        Timer?.StopTicking()

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ C L E A R ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Sub Clear()

        With Columns
            .CancelWorkers()
            .Clear()
        End With
        With Rows
            .Clear()
        End With
        _MouseData = Nothing
        DistinctValues.Clear()
        VisibleColumns.Clear()
        VisibleRows.Clear()
        VScroll.Value = 0
        HScroll.Value = 0

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ M O U S E ▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)

        _MouseData = Nothing
        Invalidate()
        MyBase.OnMouseLeave(e)

    End Sub
    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)

        If e IsNot Nothing Then
            With _MouseData
                .CurrentRegion = If(Columns.HeadBounds.Contains(e.Location), MouseRegion.Header, MouseRegion.Grid)
                Dim lastPoint As Point = .Point
                Dim newPoint As Point = e.Location
                If lastPoint <> newPoint Then
                    If .CurrentAction = MouseInfo.Action.HeaderEdgeClicked Or .CurrentAction = MouseInfo.Action.ColumnSizing Then
#Region " HEADER - SIZING ... MUST BE 1st SO MouseMove OUTSIDE HEADBOUNDS INTO GRID CONTINUES TO SIZE "
                        .CurrentAction = MouseInfo.Action.ColumnSizing
                        If e.Button = MouseButtons.Left Then
                            RaiseEvent Alert(True, New AlertEventArgs(If(.Column Is Nothing, "No column", .Column.Name)))
                            Dim Delta = e.X - .Point.X
                            With .Column
                                .Width += Delta
                                Dim formatName As String = .Format.Key.ToString
                                Dim formatType As String = .DataType.ToString
                                Dim headWidth As String = .HeadBounds.Width.ToString(InvariantCulture)
                                Dim contentWidth As String = .ContentWidth.ToString(InvariantCulture)
                                Dim rowCount As String = Rows.Count.ToString(InvariantCulture)
                                Dim sortOrder As String = .SortOrder.ToString
                                Dim alignment As String = StringFormatToContentAlignString(.GridStyle.Alignment)
                                Dim Bullets As New Dictionary(Of String, List(Of String)) From {
                                { .Text, {"Type is " & formatName,
                                "Datatype is " & formatType,
                                "Width=" & headWidth,
                                "Content Width=" & contentWidth,
                                "Index=" & .ViewIndex,
                                "Row Count=" & rowCount,
                                "Row Height=" & RowHeight,
                                "Sort Order is " & sortOrder,
                                "Alignment is " & alignment}.ToList}
                            }
                                ColumnHeadTip.SetToolTip(Me, Bulletize(Bullets))
                            End With
                            'Cursor = Cursors.VSplit
                            Cursor = CursorDirection(New Point(e.X, 0), New Point(.Point.X, 0))
                            'If .Column.Name = "InvoiceNbr" Then Stop
                        Else
                            .CurrentAction = MouseInfo.Action.None

                        End If
                        Invalidate()
#End Region
                    Else
                        Dim lastMouseColumn As Column = .Column
                        Dim lastMouseRow As Row = .Row
                        Dim Redraw As Boolean = False
                        If .CurrentRegion = MouseRegion.Header Then
#Region " HEADER REGION "
                            Dim VisibleEdges As New Dictionary(Of Column, Rectangle)
                            For Each Item In VisibleColumns
                                VisibleEdges.Add(Item.Key, New Rectangle(Item.Key.EdgeBounds.X - HScroll.Value, 0, 10, Item.Key.EdgeBounds.Height))
                            Next
                            Dim Edges = VisibleEdges.Where(Function(x) x.Value.Contains(newPoint)).Select(Function(c) c.Key)
                            If Edges.Any Then
                                .CurrentAction = MouseInfo.Action.MouseOverHeadEdge
                                .Column = Edges.First
                                Cursor = Cursors.VSplit

                            Else
                                Dim MouseColumns = VisibleColumns.Where(Function(vc) vc.Value.Contains(New Point(newPoint.X, 0))).Select(Function(c) c.Key)
                                If MouseColumns.Any Then .Column = MouseColumns.First
                                .CurrentAction = MouseInfo.Action.MouseOverHead
                                Cursor = Cursors.Default
                            End If
#End Region
                        Else
#Region " GRID REGION "
                            Cursor = Cursors.Default
                            Dim MouseRows = VisibleRows.Where(Function(r) e.Y >= r.Value.Top And e.Y <= r.Value.Bottom)
                            Dim lastMouseCell As Cell = .Cell
                            If MouseRows.Any Then
                                .Row = MouseRows.First.Key
                                .RowBounds = VisibleRows(.Row)
                                If .CurrentAction = MouseInfo.Action.CellClicked And newPoint <> lastPoint Then
                                    .CurrentAction = MouseInfo.Action.GridSelecting
                                    .SelectPointA = lastPoint

                                ElseIf .CurrentAction = MouseInfo.Action.GridSelecting Then
                                    .SelectPointB = newPoint
                                    Redraw = True
                                    ' NEEDS WORK
                                    If Width - newPoint.X < 10 Then HScroll.Value = {HScroll.Value + 20, HScroll.Maximum}.Min
                                    If Height - newPoint.Y < 10 Then VScroll.Value = {VScroll.Value + RowHeight, VScroll.Maximum}.Min
                                    ' NEEDS WORK
                                    RaiseEvent Alert({ .SelectPointA, .SelectPointB}, New AlertEventArgs("Grid selecting"))

                                Else
                                    .CurrentAction = If(.Column Is Nothing, MouseInfo.Action.None, MouseInfo.Action.MouseOverGrid)
                                    If Not Rows.SingleSelect And ControlKeyDown And .Row IsNot lastMouseRow Then
                                        .Row.Selected = e.Button = MouseButtons.Left 'Row.Selected may not take the value if Me.FullRowSelect=False
                                        Redraw = True
                                    End If
                                End If
                                If .Column IsNot Nothing Then
                                    .Cell = .Row.Cells(.Column.Name)
                                    .CellBounds = New Rectangle(.Column.HeadBounds.Left, .RowBounds.Top, .Column.HeadBounds.Width, .RowBounds.Height)
                                End If
                            Else
                                .CurrentAction = MouseInfo.Action.None
                                .Row = Nothing
                                .Cell = Nothing
                            End If
                            If lastMouseCell IsNot .Cell Then RaiseEvent CellMouseChanged(Me, New ViewerEventArgs(MouseData))
                            If lastMouseRow IsNot .Row Then RaiseEvent RowMouseChanged(Me, New ViewerEventArgs(MouseData))
#End Region
                        End If
                        If Redraw Or Columns.HeadBounds.Contains(newPoint) Or .Column IsNot lastMouseColumn Or .Row IsNot lastMouseRow Then Invalidate()
                    End If
                    .Point = newPoint
                End If
            End With
            MyBase.OnMouseMove(e)
        End If

    End Sub
    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)

        'MouseOver feeds MouseInfo.Action

        If e IsNot Nothing Then
            With HeaderOptions
                .Tag = Nothing
                .AutoClose = True
                .Hide()
            End With
            With _MouseData
                Dim clickPoint As Point = e.Location
                Dim mouseColumns = VisibleColumns.Where(Function(vc) vc.Value.Contains(New Point(clickPoint.X, 0))).Select(Function(c) c.Key) 'Use for all EXCEPT Edge ( left button )
                .CurrentRegion = If(Columns.HeadBounds.Contains(clickPoint), MouseRegion.Header, MouseRegion.Grid)
                If .CurrentRegion = MouseRegion.Header Then
#Region " HEADER "
                    If e.Button = MouseButtons.Left Then
#Region " L E F T - Sizing ( Edge ) / Sort "
                        'Check if Edge 1st since .Column calc is different ... 5 pixels into the column to the right counts as the left column
                        Dim VisibleEdges As New Dictionary(Of Column, Rectangle)
                        For Each Item In VisibleColumns
                            VisibleEdges.Add(Item.Key, New Rectangle(Item.Key.EdgeBounds.X - HScroll.Value, 0, 10, Item.Key.EdgeBounds.Height))
                        Next
                        Dim Edges = VisibleEdges.Where(Function(x) x.Value.Contains(clickPoint)).Select(Function(c) c.Key)
                        If Edges.Any Then
#Region " HEADER - EDGE "
                            .CurrentAction = MouseInfo.Action.HeaderEdgeClicked
                            .Column = Edges.First
                            Cursor = Cursors.VSplit
#End Region
                        Else
#Region " HEADER - NOT EDGE ( SORT ) "
                            .CurrentAction = MouseInfo.Action.HeaderClicked
                            .Column = If(mouseColumns.Any, mouseColumns.First, Nothing) 'Ensure .Column has current value
                            Cursor = Cursors.Default
                            'Change the sort order
                            If .Column IsNot Nothing Then
                                If .Column.SortOrder = SortOrder.Ascending Then
                                    .Column.SortOrder = SortOrder.Descending

                                Else
                                    Dim formerSortOrder = .Column.SortOrder
                                    .Column.SortOrder = SortOrder.Ascending
                                    If formerSortOrder = SortOrder.None Then .Column.AutoWidth()

                                End If
                                Rows.SortBy(.Column)
                            End If
#End Region
                        End If
#End Region
                    ElseIf e.Button = MouseButtons.Right Then
#Region " R I G H T - View / Change HeaderStyle {backcolor, shadecolor, alignment} + View items+counts / Filter items "
                        .Column = If(mouseColumns.Any, mouseColumns.First, Nothing) 'Ensure .Column has current value
                        Dim headerProperties As ToolStripMenuItem = DirectCast(HeaderOptions.Items("properties"), ToolStripMenuItem)
                        If headerProperties IsNot Nothing Then
                            Dim tschProperties As ToolStripControlHost = DirectCast(headerProperties.DropDownItems(0), ToolStripControlHost)
                            Dim tlpProperties As TableLayoutPanel = DirectCast(tschProperties.Control, TableLayoutPanel)
                            tlpProperties.ColumnStyles(0).Width = 0 'Don't need a big button
                            TLP.SetSize(tlpProperties)
                            Dim subButton As ImageCombo = DirectCast(tlpProperties.GetControlFromPosition(0, 0), ImageCombo)
                            subButton.Text = If(.Column Is Nothing, "All columns".ToString(InvariantCulture), .Column.Name)

                            HeaderGridAlignment.Text = StringFormatToContentAlignString(If(.Column Is Nothing, Columns.HeaderStyle.Alignment, .Column.GridStyle.Alignment))
                            HeaderGridAlignment.SelectedIndex = HeaderGridAlignment.TextIndex

#Region " Adhoc pull on Distinct values - too problematic using Columns.ColumnWidth @#$%^ "
                            If .Column IsNot Nothing Then
                                Dim columnValues As Dictionary(Of String, List(Of Cell)) = DistinctValues(.Column.Name)
                                If Not columnValues.Any Then 'Get them, otherwise don't
                                    For Each row In Rows
                                        Dim rowCell As Cell = row.Cells(.Column.Name)
                                        Dim rowCellText As String = If(rowCell.Text, "(null)")
                                        If Not columnValues.ContainsKey(rowCellText) Then columnValues.Add(rowCellText, New List(Of Cell))
                                        columnValues(rowCellText).Add(rowCell)
                                    Next
                                End If
                                With HeaderDistinctItems
                                    With .Items
                                        .Clear()
                                        Dim values = columnValues.OrderByDescending(Function(c) c.Value.Count).ThenBy(Function(t) t.Key)
                                        For Each distinctValue In values
                                            .Add(distinctValue.Key & " (" & distinctValue.Value.Count & ")")
                                        Next
                                    End With
                                    .HintText = "Distinct count = " & .Items.Count
                                End With
                            End If
#End Region
                            Dim relativePoint As Point = If(.Column Is Nothing, e.Location, New Point(.Column.HeadBounds.Right - HeaderOptions.Width, HeaderHeight))
                            With HeaderOptions
                                .AutoClose = False
                                .Tag = _MouseData.Column
                                .Location = PointToScreen(relativePoint)
                            End With
                            headerProperties.ShowDropDown()
                        End If
#End Region
                    End If
#End Region
                Else
#Region " GRID "
                    Dim mouseRows = VisibleRows.Where(Function(r) e.Y >= r.Value.Top And e.Y <= r.Value.Bottom)
                    .Row = If(mouseRows.Any, mouseRows.First.Key, Nothing)
                    .Column = If(mouseColumns.Any, mouseColumns.First, Nothing) 'Ensure .Column has current value
                    .Cell = If(.Row Is Nothing Or .Column Is Nothing, Nothing, .Row.Cells(.Column.Name))
                    If e.Button = MouseButtons.Left Then
                        With GridOptions
                            .Tag = Nothing
                            .AutoClose = True
                            .Hide()
                        End With
#Region " L E F T - Cell / Row selection "
                        If .Cell Is Nothing Then
                            .CurrentAction = MouseInfo.Action.None
                        Else
                            .CurrentAction = MouseInfo.Action.CellClicked
                            .Row.Selected = Not .Row.Selected  'Row.Selected may not take the value if Me.FullRowSelect=False
                            Dim cellSelectedCounter As Integer
                            Rows.ForEach(Function(row) As Row
                                             For Each cell In row.Cells.Values.Except({ .Cell}).Where(Function(c) c.Selected)
                                                 cell.Selected = False
                                                 cellSelectedCounter += 1
                                             Next
                                             Return Nothing
                                         End Function)
                            .Cell.Selected = If(cellSelectedCounter = 0, Not .Cell.Selected, True)
                            RaiseEvent CellClicked(Me, New ViewerEventArgs(MouseData))
                        End If
#End Region
                    ElseIf e.Button = MouseButtons.Right Then
#Region " R I G H T - Show Cell properties "
                        GridOptions.AutoClose = False
                        Dim relativePoint As Point = If(.Cell Is Nothing, e.Location, New Point(.CellBounds.Right, .CellBounds.Top))
                        GridOptions.Show(PointToScreen(relativePoint))
#End Region
                    End If
#End Region
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
                .Column.AutoWidth()
                Invalidate()

            ElseIf .CurrentAction = MouseInfo.Action.CellClicked Then
                .CurrentAction = MouseInfo.Action.CellDoubleClicked
                RaiseEvent CellDoubleClicked(Me, New ViewerEventArgs(MouseData))

            End If
        End With
        MyBase.OnMouseDoubleClick(e)

    End Sub
    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)

        With _MouseData
            .CurrentAction = MouseInfo.Action.None
            .SelectPointA = Nothing
            .SelectPointB = Nothing
            Invalidate()
        End With
        MyBase.OnMouseUp(e)

    End Sub
    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)

        If e IsNot Nothing Then

            ControlKeyDown = e.KeyCode = Keys.ControlKey

            Dim CursorIndex As Integer
            Dim IsReadOnly As Boolean = True

            Try
                Dim S As Integer = CursorIndex
                If e.KeyCode = Keys.Left Or e.KeyCode = Keys.Right Then
#Region " MOVE LEFT Or RIGHT "
                    'Dim Value As Integer = If(e.KeyCode = Keys.Left, -1, 1)
                    'If Control.ModifierKeys = Keys.Shift Then
                    '    SelectionIndex += Value
                    'Else
                    '    CursorIndex += Value
                    '    SelectionIndex = CursorIndex
                    'End If
#End Region
                ElseIf e.KeyCode = Keys.Back Or e.KeyCode = Keys.Delete And Not IsReadOnly Then
#Region " REMOVE BACK Or AHEAD "
                    'If CursorIndex = SelectionIndex Then
                    '    If e.KeyCode = Keys.Back Then
                    '        If Not S = 0 Then
                    '            CursorIndex -= 1
                    '            SelectionIndex = CursorIndex
                    '            Text = Text.Remove(S - 1, 1)
                    '        End If
                    '    ElseIf e.KeyCode = Keys.Delete Then
                    '        If Not S = Text.Length Then
                    '            Text = Text.Remove(S, 1)
                    '        End If
                    '    End If
                    'Else
                    '    Dim TextLength As Integer = SelectionLength
                    '    CursorIndex = SelectionStart
                    '    SelectionIndex = CursorIndex
                    '    Text = Text.Remove(SelectionStart, TextLength)
                    'End If
#End Region
                ElseIf e.KeyCode = Keys.A AndAlso Control.ModifierKeys = Keys.Control Then
#Region " SELECT ALL "
                    Rows.ForEach(Function(row) As Row
                                     For Each cell In row.Cells.Values
                                         cell.Selected = True
                                     Next
                                     Return Nothing
                                 End Function)
#End Region
                ElseIf e.KeyCode = Keys.X AndAlso Control.ModifierKeys = Keys.Control And Not IsReadOnly Then
#Region " CUT "
                    'Dim TextSelection As String = Selection
                    'CursorIndex = SelectionStart
                    'SelectionIndex = CursorIndex
                    'Clipboard.SetText(TextSelection)
                    'Text = Text.Remove(SelectionStart, TextSelection.Length)
#End Region
                ElseIf e.KeyCode = Keys.C AndAlso Control.ModifierKeys = Keys.Control Then
#Region " COPY "
                    Dim selectedCells As New Dictionary(Of Cell, Integer)
                    Dim selectedHeaders As New List(Of String)
                    Rows.ForEach(Function(row) As Row
                                     For Each cell In row.Cells.Values.Where(Function(c) c.Selected)
                                         If Not selectedHeaders.Contains(cell.Name) Then selectedHeaders.Add(cell.Name) 'Don't use the Cell.DataType - Need an Aggregate DataType
                                         selectedCells.Add(cell, row.Index)
                                     Next
                                     Return Nothing
                                 End Function)
                    If selectedCells.Any Then
                        Clipboard.Clear()
                        If selectedCells.Count = 1 Then
                            Clipboard.SetText(If(selectedCells.First.Key.Text, String.Empty))
                            CopyTimer.Tag = -1
                        Else
                            Using copyTable As New DataTable
                                With copyTable
                                    For Each columnName In selectedHeaders
                                        .Columns.Add(columnName, Columns.Item(columnName).DataType)
                                    Next
                                    Dim selectedRows = (From sc In selectedCells Group sc By rowIndex = sc.Value Into rowGroup = Group
                                                        Select New With {.Index = rowIndex, .rowValues = (From c In rowGroup Order By c.Key.Index Select c.Key.Value).ToArray}).ToDictionary(Function(k) k.Index, Function(v) v.rowValues)
                                    For Each row In selectedRows.Values
                                        .Rows.Add(row)
                                    Next
                                End With
                                Dim htmlTable As String = DataTableToHtml(copyTable, Columns.HeaderStyle.BackColor, Columns.HeaderStyle.ForeColor)
                                ClipboardHelper.CopyToClipboard(htmlTable)
                                CopyTimer.Tag = copyTable.Rows.Count
                            End Using
                        End If
                    End If
                    Invalidate()
                    CopyTimer.Start()
#End Region
                ElseIf e.KeyCode = Keys.V AndAlso Control.ModifierKeys = Keys.Control And Not IsReadOnly Then
#Region " PASTE "
                    'S = SelectionStart
                    'Text = Text.Remove(SelectionStart, SelectionLength)
                    'Dim ClipboardText As String = Clipboard.GetText
                    'Text = Text.Insert(S, ClipboardText)
                    'CursorIndex = S + ClipboardText.Length
                    'SelectionIndex = CursorIndex
#End Region
                ElseIf e.KeyCode = Keys.Enter Then
#Region " SUBMIT "
                    'RaiseEvent ValueSubmitted(Me, New ImageComboEventArgs(Nothing))
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
                    'SelectedIndex += Value
#End Region
                End If
                'KeyedValue = Text
                Invalidate()

            Catch ex As IndexOutOfRangeException
                MsgBox(ex.Message & vbCrLf & ex.StackTrace)

            End Try
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
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ GRID PROPERTIES INPUT
    Private Sub Scrolled(sender As Object, e As ScrollEventArgs) Handles VScroll.Scroll, HScroll.Scroll
        Invalidate()
    End Sub
    Private Sub CellStyleProperty_SelectionChanged(sender As Object, e As ImageComboEventArgs)

        Dim propertyCombo As ImageCombo = DirectCast(sender, ImageCombo)
        Dim propertyColor As Color = DirectCast(propertyCombo.SelectedItem.Tag, Color)
        Dim propertySetting As System.Configuration.SettingsPropertyValue = DirectCast(propertyCombo.Tag, System.Configuration.SettingsPropertyValue)

        Select Case True
            Case propertySetting.Name.Contains("HeaderBackColor")
                Columns.HeaderStyle.BackColor = propertyColor

            Case propertySetting.Name.Contains("HeaderShadeColor")
                Columns.HeaderStyle.ShadeColor = propertyColor

            Case propertySetting.Name.Contains("HeaderForeColor")
                Columns.HeaderStyle.ForeColor = propertyColor

            Case propertySetting.Name.Contains("AlternatingBackColor")
                Rows.AlternatingRowStyle.BackColor = propertyColor

            Case propertySetting.Name.Contains("AlternatingShadeColor")
                Rows.AlternatingRowStyle.ShadeColor = propertyColor

            Case propertySetting.Name.Contains("AlternatingForeColor")
                Rows.AlternatingRowStyle.ForeColor = propertyColor

            Case propertySetting.Name.Contains("SelectionBackColor")
                Rows.SelectionRowStyle.BackColor = propertyColor

            Case propertySetting.Name.Contains("SelectionShadeColor")
                Rows.SelectionRowStyle.ShadeColor = propertyColor

            Case propertySetting.Name.Contains("SelectionForeColor")
                Rows.SelectionRowStyle.ForeColor = propertyColor

            Case propertySetting.Name.Contains("RowBackColor")
                Rows.RowStyle.BackColor = propertyColor

            Case propertySetting.Name.Contains("RowShadeColor")
                Rows.RowStyle.ShadeColor = propertyColor

            Case propertySetting.Name.Contains("RowForeColor")
                Rows.RowStyle.ForeColor = propertyColor

        End Select
        Invalidate()

    End Sub
    Private Sub HeaderGridAlignment_SelectionChanged(sender As Object, e As ImageComboEventArgs) Handles HeaderGridAlignment.SelectionChanged

        If HeaderOptions.Tag Is Nothing Then
            Columns.HeaderStyle.Alignment = ContentAlignToStringFormat(HeaderGridAlignment.Text)
        Else
            Dim mouseColumn As Column = DirectCast(HeaderOptions.Tag, Column)
            If sender Is HeaderGridAlignment Then
                mouseColumn.GridStyle.Alignment = ContentAlignToStringFormat(HeaderGridAlignment.Text)
                RaiseEvent Alert(mouseColumn.GridStyle, New AlertEventArgs(mouseColumn.GridStyle.Alignment.ToString))
            End If
        End If

    End Sub
    Private Sub Filter_Enter(sender As Object, e As EventArgs)
        HeaderDistinctItems.DropDown.Show()
    End Sub
    Private Sub Filter_Click(sender As Object, e As EventArgs)
        With DirectCast(sender, ToolStripMenuItem)
            .Tag = Not DirectCast(.Tag, Boolean)
            If DirectCast(.Tag, Boolean) Then
                .HideDropDown()
            Else
                .ShowDropDown()
            End If
        End With
    End Sub
    Private Sub Filter_Leave(sender As Object, e As EventArgs)
        With DirectCast(sender, ToolStripMenuItem)
            If DirectCast(.Tag, Boolean) Then .HideDropDown()
        End With
    End Sub
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)
        With HeaderOptions
            .Tag = Nothing
            .AutoClose = True
            .Hide()
        End With
        With GridOptions
            .Tag = Nothing
            .AutoClose = True
            .Hide()
        End With
        MyBase.OnVisibleChanged(e)
    End Sub

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ T O   F I L E
    Public Sub Export(filePath As String)
        DataTableToExcel(Table, filePath, True, False, False, True, True)
    End Sub
    Private Sub ExportToFile(sender As Object, e As EventArgs)

        If Rows.Any Then
            With SaveFile
                .InitialDirectory = Desktop
                Dim ExportObject As ToolStripDropDownItem = DirectCast(sender, ToolStripDropDownItem)
                Select Case ExportObject.Text
                    Case "Excel"
                        .FileName = "FileName.xlsx"
                        .Filter = "Excel Files|*.xls,*.xlsx".ToString(InvariantCulture)

                    Case ".csv"
                        .FileName = "FileName.csv"
                        .Filter = "CSV|*.csv".ToString(InvariantCulture)

                    Case ".txt"
                        .FileName = "FileName.txt"
                        .Filter = "TXT Files (*.txt*)|*.txt".ToString(InvariantCulture)

                    Case Else

                End Select
                .ShowDialog()
            End With
        Else

        End If

    End Sub
    Private Sub SaveFileClosed(sender As Object, e As EventArgs) Handles SaveFile.FileOk

        Select Case GetFileNameExtension(SaveFile.FileName).Value
            Case ExtensionNames.Excel
                AddHandler Alerts, AddressOf FileSaved
                DataTableToExcel(Table, SaveFile.FileName, True, False, False, True, True)

            Case ExtensionNames.Text
                DataTableToTextFile(Table, SaveFile.FileName)

            Case ExtensionNames.CommaSeparated
            Case ExtensionNames.SQL

        End Select

    End Sub
    Private Sub FileSaved(sender As Object, e As AlertEventArgs)
        RaiseEvent Alert(sender, e)
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
    Private IsBusy_ As Boolean = False
    Friend ReadOnly Property IsBusy As Boolean
        Get
            IsBusy_ = ColumnsWorker.IsBusy Or IsBusy_
            Return IsBusy_
        End Get
    End Property
    Private WithEvents AddRemoveTimer As New Timer With {.Interval = 100}
    Private WithEvents ReOrderTimer As New Timer With {.Interval = 100}
    Private ReadOnly MoveColumns As New Dictionary(Of Column, Integer)
    Public Sub New(Viewer As DataViewer)
        Parent = Viewer
    End Sub
    Public ReadOnly Property Parent As DataViewer
    Public ReadOnly Property Names As Dictionary(Of String, Integer)
        Get
            Dim columnNames As New Dictionary(Of String, Integer)
            For Each column In Me
                columnNames.Add(column.Name, column.ViewIndex)
            Next
            Return columnNames
        End Get
    End Property
    Public ReadOnly Property Width As Integer
        Get
            Return Sum(Function(c) c.Width)
        End Get
    End Property
    Public ReadOnly Property HeadBounds As Rectangle
        Get
            Dim border3D As Integer = 3
            If Count = 0 Then
                Dim HeadSize As New Size(Parent.Width, 3 + TextRenderer.MeasureText("XXXXXXXXXXX".ToString(InvariantCulture), HeaderStyle.Font).Height + 3 + border3D)
                Return New Rectangle(0, 0, HeadSize.Width, {HeadSize.Height, HeaderStyle.Height}.Max)
            Else
                Return New Rectangle(0, 0, Max(Function(c) c.HeadBounds.Right), Max(Function(c) {c.HeadBounds.Height, HeaderStyle.Height}.Max) + border3D)
            End If
        End Get
    End Property
    Private WithEvents HeaderStyle_ As New CellStyle With {.BackColor = Color.Black, .ShadeColor = Color.LimeGreen, .ForeColor = Color.White, .Font = New Font("Century Gothic", 9, FontStyle.Bold), .Height = 24, .ImageScaling = Scaling.GrowParent, .Padding = New Padding(2)}
    Public ReadOnly Property HeaderStyle As CellStyle
        Get
            Return HeaderStyle_
        End Get
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Shadows Function Add(ByVal AddColumn As Column) As Column

        If AddColumn IsNot Nothing Then
            With AddColumn
                .Parent_ = Me
                .HeaderStyle = HeaderStyle
                ._Index = Count
                ColumnsXH()
            End With
            MyBase.Add(AddColumn)
        End If
        Return AddColumn

    End Function
    Public Shadows Function Insert(MoveIndex As Integer, ByVal MoveColumn As Column) As Column

        If MoveColumn IsNot Nothing Then
            With MoveColumn
                .Parent_ = Me
                If MoveIndex >= 0 And MoveIndex < Count Then
                    ._Index = MoveIndex
                    ColumnsXH()
                    MyBase.Insert(MoveIndex, MoveColumn)
                Else
                    'Do nothing ... would like to Throw an Exception so the error is caught by the end-user
                End If
            End With
        End If
        Return MoveColumn

    End Function
    Public Shadows Function Remove(ByVal RemoveColumn As Column) As Column

        If RemoveColumn IsNot Nothing Then
            AddRemoveTimer.Start()
            With RemoveColumn
                .Parent_ = Nothing
                ._Index = -1
            End With
            MyBase.Remove(RemoveColumn)
        End If
        Return RemoveColumn

    End Function
    Public Shadows Function Remove(columnName As String) As Column

        If columnName Is Nothing Then
            Return Nothing
        Else
            Dim columnItem As Column = Item(columnName)
            If columnItem Is Nothing Then
                Return Nothing
            Else
                Return Remove(columnItem)
            End If
        End If

    End Function
    Private Sub AddRemoveTimer_Tick() Handles AddRemoveTimer.Tick
        If Not IsBusy Then AddRemoveTimer.Stop()
        ColumnsXH()
    End Sub
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
    Friend Sub ColumnsXH()

        Dim columnHeight As Integer = 0
        Try
            For Each column In Me
                columnHeight = {column.HeadSize.Height, column.HeaderStyle.Height, columnHeight}.Max
            Next
        Catch ex As InvalidOperationException

        End Try
        Dim columnLeft As Integer = 0
        Try
            For Each column In Me
                column.HeadBounds = New Rectangle(columnLeft, 0, column.Width, columnHeight)
                column.HeaderStyle.Height = columnHeight
                columnLeft += column.Width
            Next
            Parent?.Invalidate()
        Catch ex As InvalidOperationException
        End Try

    End Sub
    Friend Sub Reorder(Column As Column, ViewIndex As Integer)

        If Not MoveColumns.ContainsKey(Column) Then MoveColumns.Add(Column, ViewIndex)
        If IsBusy Then
            AddHandler CollectionSizingEnd, AddressOf CanReorder
        Else
            CanReorder(Nothing, Nothing)
        End If

    End Sub
    Private Sub CanReorder(sender As Object, e As EventArgs)

        RemoveHandler CollectionSizingEnd, AddressOf CanReorder
        IsBusy_ = True
        For Each Column In MoveColumns
            Remove(Column.Key)
            Insert(Column.Value, Column.Key)
        Next
        MoveColumns.Clear()
        IsBusy_ = False
        ColumnsXH()

    End Sub
    Private WithEvents ColumnsWorker As New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
    Friend Sub FormatSize()             ' I N I T I A L  F O R M A T + S I Z I N G
        RaiseEvent CollectionSizingStart(Me, Nothing)
        If Not ColumnsWorker.IsBusy Then ColumnsWorker.RunWorkerAsync()
    End Sub
    Public Sub ColumnWidths()
        For Each Column In Me
            ColumnWidth(Column)
        Next
        RaiseEvent CollectionSizingEnd(Me, Nothing)
    End Sub
    Public Sub DistibuteWidths(Optional testing As Boolean = False)

        'If Viewer.Width>Columns.Width ... Then share extra space among columns
        Dim VisibleColumns As New List(Of Column)(From c In Me Where c.Visible)
        If VisibleColumns.Any Then
            Dim parentControl As Control = Parent
            Do While parentControl.Parent IsNot Nothing AndAlso parentControl.Dock = DockStyle.Fill 'Or <> DockStyle.None?
                parentControl = parentControl.Parent
            Loop
            Dim ExtraWidth = CInt((parentControl.Width - HeadBounds.Width) / Count)
            If ExtraWidth >= 1 Then
                Dim rollingWidth As Integer = 0
                Dim maxWidth As Integer = parentControl.Width
                'Space to spare
                For Each visibleColumn In VisibleColumns
                    rollingWidth += ExtraWidth
                    visibleColumn.Width += ExtraWidth
                    If rollingWidth > maxWidth Then Exit For
                Next
                If testing Then Stop
                Do While Parent.HScrollVisible
                    For Each Column In Me
                        Column.Width -= 1
                        Parent.Invalidate()
                    Next
                Loop
            End If
        End If

    End Sub
    Private Sub FormatSizeWorker_Start(sender As Object, e As DoWorkEventArgs) Handles ColumnsWorker.DoWork

        For Each Column In Where(Function(c) c.Visible)
            ColumnWidth(Column, True)
            If ColumnsWorker.CancellationPending Then Exit For
        Next

    End Sub
    Friend Sub ColumnWidth(ColumnItem As Column, Optional BackgroundProcess As Boolean = False)

        With ColumnItem
            Dim cellTypes As New List(Of Type)
            Dim cellValues As New List(Of Object)
            .ContentWidth = .MinimumWidth
            For Each row In Parent.Rows
                Dim rowCell As Cell = row.Cells(.Name)
                cellValues.Add(rowCell.Value)
                If rowCell IsNot Nothing Then
                    cellTypes.Add(rowCell.DataType)
                    Dim cellText As String = If(rowCell.Text, String.Empty)
                    If rowCell.ValueImage Is Nothing Then
                        Dim rowStyle As CellStyle = row.Style
                        .ContentWidth = { .ContentWidth, TextRenderer.MeasureText(cellText, rowStyle.Font).Width}.Max
                    Else
                        Try
                            .ContentWidth = { .ContentWidth, rowCell.ValueImage.Width}.Max
                        Catch ex As InvalidOperationException
                        End Try
                    End If
                End If
            Next
            .Width = { .HeadSize.Width, .ContentWidth}.Max
            If BackgroundProcess Then
                'Dim aggregateDataType As Type = GetDataType(cellTypes, .Name = "COMPANY")
                Dim aggregateValueType As Type = GetDataType(cellValues)
                'If .Name = "IBMDIV" Then Stop
                'If .Name = "COMPANY" Then Stop
                ColumnsWorker.ReportProgress({0, .Index}.Max, New KeyValuePair(Of Column, Type)(ColumnItem, aggregateValueType))
            End If
        End With

    End Sub
    Private Sub FormatSizeColumn_Progress(sender As Object, e As ProgressChangedEventArgs) Handles ColumnsWorker.ProgressChanged

        'Can not change the .DataType in the Background Thread *** New *** Null DataType = DON'T CHANGE EXISTING
        Dim kvp = DirectCast(e.UserState, KeyValuePair(Of Column, Type))
        With kvp
            If .Value IsNot Nothing Then .Key.DataType = .Value
        End With
        RaiseEvent ColumnSized(Me(e.ProgressPercentage), Nothing)

    End Sub
    Private Sub FormatSizeWorker_End(sender As Object, e As RunWorkerCompletedEventArgs) Handles ColumnsWorker.RunWorkerCompleted
        RaiseEvent CollectionSizingEnd(Me, Nothing)
    End Sub
    Friend Sub CancelWorkers()
        ColumnsWorker.CancelAsync()
    End Sub
    Private Sub HeadersStyle_PropertyChanged(sender As Object, e As StyleEventArgs) Handles HeaderStyle_.PropertyChanged

        Dim value As CellStyle = DirectCast(sender, CellStyle)
        For Each Column In Me
            With Column
                If value IsNot Nothing Then
                    If e.ChangedProperty = CellStyle.Properties.Alignment Then .HeaderStyle.Alignment = HeaderStyle.Alignment
                    If e.ChangedProperty = CellStyle.Properties.BackColor Then .HeaderStyle.BackColor = HeaderStyle.BackColor
                    If e.ChangedProperty = CellStyle.Properties.Font Then .HeaderStyle.Font = HeaderStyle.Font
                    If e.ChangedProperty = CellStyle.Properties.ForeColor Then .HeaderStyle.ForeColor = HeaderStyle.ForeColor
                    If e.ChangedProperty = CellStyle.Properties.Height Then .HeaderStyle.Height = HeaderStyle.Height
                    If e.ChangedProperty = CellStyle.Properties.ImageScaling Then .HeaderStyle.ImageScaling = HeaderStyle.ImageScaling
                    If e.ChangedProperty = CellStyle.Properties.Padding Then .HeaderStyle.Padding = HeaderStyle.Padding
                    If e.ChangedProperty = CellStyle.Properties.ShadeColor Then .HeaderStyle.ShadeColor = HeaderStyle.ShadeColor
                    If e.ChangedProperty = CellStyle.Properties.Theme Then .HeaderStyle.Theme = HeaderStyle.Theme
                End If
            End With
        Next

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
                AddRemoveTimer.Dispose()
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
    Public Event WidthChanged(sender As Object, e As EventArgs)
    Friend Enum TypeGroup
        None
        Booleans
        Decimals
        Integers
        Dates
        Times
        Images
        Strings
    End Enum
    Public Sub New(NewColumn As DataColumn)

        If NewColumn IsNot Nothing Then
            With NewColumn
                Name = .ColumnName
                DTable = .Table
                _DataType = NewColumn.DataType
                DColumn = NewColumn
            End With
        End If

    End Sub
    <NonSerialized> Private ReadOnly DTable As DataTable
    <NonSerialized> Private DColumn As DataColumn
    <NonSerialized> Friend Parent_ As ColumnCollection
    Public ReadOnly Property Parent As ColumnCollection
        Get
            Return Parent_
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
            Parent?.Reorder(Me, value)
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
    Private WithEvents HeaderStyle_ As New CellStyle With {.BackColor = Color.Black, .ShadeColor = Color.Purple, .ForeColor = Color.White, .Font = New Font("Century Gothic", 9), .Height = 24, .ImageScaling = Scaling.GrowParent, .Padding = New Padding(2)}
    Public Property HeaderStyle As CellStyle
        Get
            Return HeaderStyle_
        End Get
        Set(value As CellStyle)
            If value <> HeaderStyle_ Then HeaderStyle_ = value
        End Set
    End Property
    Private WithEvents GridStyle_ As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.Transparent, .ForeColor = Color.Transparent, .Font = New Font("Century Gothic", 8), .Height = 22, .ImageScaling = Scaling.GrowParent, .Padding = New Padding(2)}
    Public Property GridStyle As CellStyle
        Get
            Return GridStyle_
        End Get
        Set(value As CellStyle)
            If value <> GridStyle_ Then GridStyle_ = value
        End Set
    End Property
    Private _MinimumWidth As Integer = 60
    Public Property MinimumWidth As Integer
        Get
            Return _MinimumWidth
        End Get
        Set(value As Integer)
            If _MinimumWidth <> value Then
                _MinimumWidth = value
                If value > Width Then Parent?.Parent?.Invalidate()
            End If
        End Set
    End Property
    Public Property HeadBounds As Rectangle
    Friend ReadOnly Property SizeImage As Size
        Get
            Try
                Return If(Image Is Nothing, New Size(0, 0),
                                If(HeaderStyle.ImageScaling = Scaling.GrowParent,
                                        New Size(Image.Width, Image.Height),
                                        New Size(HeaderStyle.Height - (HeaderStyle.Padding.Top + HeaderStyle.Padding.Bottom), HeaderStyle.Height)))
            Catch ex As InvalidOperationException
                Return Nothing
            End Try
        End Get
    End Property
    Friend ReadOnly Property SizeText As Size
        Get
            Return TextRenderer.MeasureText(Text, HeaderStyle.Font)
        End Get
    End Property
    Friend ReadOnly Property SizeFilter As Size
        Get
            Return If(Filtered, My.Resources.FilterCancel.Size, New Size(0, 0))
        End Get
    End Property
    Friend ReadOnly Property SizeSort As Size
        Get
            Return If(SortOrder = SortOrder.None, New Size(0, 0), My.Resources.SortUp.Size) 'Same size, up or down
        End Get
    End Property
    Friend ReadOnly Property HeadSize As Size
        Get
            'IMAGE      TEXT    FILTER      SORT
            '[I]       [ABC]     [F]        [S]

            Dim imageSize As Size = SizeImage,
                textSize As Size = SizeText,
                filterSize As Size = SizeFilter,
                sortSize As Size = SizeSort

            Return New Size({HeaderStyle.Padding.Left + imageSize.Width + SizeWidth(imageSize) + textSize.Width + SizeWidth(textSize) + filterSize.Width + SizeWidth(filterSize) + sortSize.Width + SizeWidth(sortSize) + HeaderStyle.Padding.Right, MinimumWidth}.Max,
                                     HeaderStyle.Padding.Top + {imageSize.Height, textSize.Height, filterSize.Height, sortSize.Height}.Max + HeaderStyle.Padding.Bottom)

        End Get
    End Property
    Private _Width As Integer = 2
    Public Property Width As Integer
        Get
            If Visible Then
                Return _Width
            Else
                Return 0
            End If
        End Get
        Set(value As Integer)
            If _Width <> value And Visible Then
                If value < 2 Then value = 2
                _Width = {value, MinimumWidth}.Max
                _HeadBounds.Width = _Width
                Parent?.ColumnsXH()
                RaiseEvent WidthChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Public Property ContentWidth As Integer = 0
    Private _Image As Image = Nothing
    Property Image() As Image
        Get
            Return _Image
        End Get
        Set(ByVal value As Image)
            If Not SameImage(value, _Image) Then
                _Image = value
                Parent?.ColumnWidth(Me)
            End If
        End Set
    End Property
    Private _Text As String = Nothing
    Property Text() As String
        Get
            Return _Text
        End Get
        Set(ByVal value As String)
            If _Text <> value Then
                _Text = value
                Parent?.ColumnsXH()
            End If
        End Set
    End Property
    Private _Filtered As Boolean = False
    Friend Property Filtered As Boolean
        Get
            Return _Filtered
        End Get
        Set(value As Boolean)
            If Not value = _Filtered Then
                _Filtered = value
                Parent?.ColumnsXH()
            End If
        End Set
    End Property
    Private _SortOrder As SortOrder = SortOrder.None
    Public Property SortOrder As SortOrder
        Get
            Return _SortOrder
        End Get
        Set(value As SortOrder)
            If Not value = _SortOrder Then
                _SortOrder = value
                Parent?.Parent?.Rows.SortBy(Me)
                Parent?.ColumnWidth(Me)
            End If
        End Set
    End Property
    Public ReadOnly Property EdgeBounds As Rectangle
        Get
            Return New Rectangle(HeadBounds.Right - 5, 0, 10, HeadBounds.Height)
        End Get
    End Property
    Public Property Visible As Boolean = True
    Private Shared Function SizeWidth(sz As Size) As Integer
        Return If(sz.Width = 0, 0, 2)
    End Function
    Private Sub HeaderStyle_PropertyChanged(sender As Object, e As StyleEventArgs) Handles HeaderStyle_.PropertyChanged

        Select Case e.ChangedProperty
            Case CellStyle.Properties.Height
                Parent?.ColumnsXH()
        End Select
        Parent?.ColumnsXH()

    End Sub
    Private Sub GridStyle_PropertyChanged(sender As Object, e As StyleEventArgs) Handles GridStyle_.PropertyChanged
        Parent?.Parent?.Invalidate()
    End Sub
    Private Sub DrawTimer_Tick(sender As Object, e As EventArgs)
        With DirectCast(sender, Timer)
            RemoveHandler .Tick, AddressOf DrawTimer_Tick
            .Stop()
        End With
        Parent?.Parent?.Invalidate()
    End Sub
    Public Sub AutoWidth()
        Parent?.ColumnWidth(Me)
    End Sub
    Friend _DataType As Type
    Public Property DataType As Type
        Get
            Return _DataType
        End Get
        Set(value As Type)
            If value IsNot Nothing And DataType <> value Then
                Dim existingFormat = Get_kvpFormat(_DataType)
                _DataType = value
                If existingFormat.Key <> Format.Key And Not {TypeGroup.Dates, TypeGroup.Times}.Intersect({existingFormat.Key, Format.Key}).Count = 2 Then
#Region " CHANGE/REORDER DATATABLE COLUMNS - REMOVE OLD DATATYPE, INSERT NEW - DO NOT DO THIS FOR Dates Or Times AS SYSTEM.DateAndTime IS NOT A REAL TYPE "
                    Dim columnValues As New List(Of Object)(DataColumnToList(DColumn))
                    Dim ColumnOridinal As Integer = DColumn.Ordinal
                    DTable.Columns.Remove(DColumn)
                    Dim NewColumn As DataColumn = New DataColumn With {.DataType = value, .ColumnName = DColumn.ColumnName}
                    DTable.Columns.Add(NewColumn)
                    NewColumn.SetOrdinal(ColumnOridinal)
                    DColumn = NewColumn
                    Dim rowCounter As Integer
                    For Each row In DTable.AsEnumerable
                        Try
                            row(DColumn) = columnValues(rowCounter)
                        Catch ex As ArgumentException
                        End Try
                        rowCounter += 1
                    Next
#End Region
                    With New Timer With {.Interval = 100}
                        AddHandler .Tick, AddressOf DrawTimer_Tick
                        .Start()
                    End With
                End If
            End If
            Select Case value
                Case GetType(Boolean), GetType(Byte), GetType(Short), GetType(Integer), GetType(Long), GetType(Date), GetType(DateAndTime), GetType(Image), GetType(Bitmap), GetType(Icon)
                    GridStyle_.Alignment = New StringFormat With {
        .Alignment = StringAlignment.Center,
        .LineAlignment = StringAlignment.Center}

                Case GetType(Decimal), GetType(Double)
                    GridStyle_.Alignment = New StringFormat With {
        .Alignment = StringAlignment.Far,
        .LineAlignment = StringAlignment.Center}

                Case Else
                    GridStyle_.Alignment = New StringFormat With {
        .Alignment = StringAlignment.Near,
        .LineAlignment = StringAlignment.Center}

            End Select
        End Set
    End Property
    Friend Shared Function Get_kvpFormat(DataType As Type) As KeyValuePair(Of TypeGroup, String)

        Select Case DataType
            Case GetType(Boolean)
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.Booleans, String.Empty)

            Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.Integers, String.Empty)

            Case GetType(Date)
                Dim CultureInfo = Threading.Thread.CurrentThread.CurrentCulture
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.Dates, CultureInfo.DateTimeFormat.ShortDatePattern)

            Case GetType(DateAndTime)
                Dim CultureInfo = Threading.Thread.CurrentThread.CurrentCulture
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.Times, CultureInfo.DateTimeFormat.FullDateTimePattern)

            Case GetType(Decimal), GetType(Double)
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.Decimals, "C2")

            Case GetType(Image), GetType(Bitmap), GetType(Icon)
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.Images, String.Empty)

            Case GetType(String)
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.Strings, String.Empty)

            Case Else
                Return New KeyValuePair(Of TypeGroup, String)(TypeGroup.None, String.Empty)

        End Select

    End Function
    Friend ReadOnly Property Format() As KeyValuePair(Of TypeGroup, String)
        Get
            Return Get_kvpFormat(DataType)
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return Join({Name, DataType.ToString, Format.Key}, ", ")
    End Function
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                _Image.Dispose()
                HeaderStyle_.Dispose()
                GridStyle_.Dispose()
                DColumn.Dispose()
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
    Private WithEvents DrawTimer As New Timer With {.Interval = 100}
    Private Sub DrawTimer_Tick() Handles DrawTimer.Tick
        DrawTimer.Stop()
        Parent.Invalidate()
    End Sub
    Private WithEvents HeaderStyle_ As New CellStyle With {.BackColor = Color.Silver, .ShadeColor = Color.Gainsboro, .ForeColor = Color.White, .Font = New Font("Century Gothic", 9)}
    Public ReadOnly Property HeaderStyle As CellStyle
        Get
            Return HeaderStyle_
        End Get
    End Property
    Private WithEvents RowStyle_ As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.White, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
    Public ReadOnly Property RowStyle As CellStyle
        Get
            Return RowStyle_
        End Get
    End Property
    Private WithEvents AlternatingRowStyle_ As New CellStyle With {.BackColor = Color.Silver, .ShadeColor = Color.Lavender, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
    Public ReadOnly Property AlternatingRowStyle As CellStyle
        Get
            Return AlternatingRowStyle_
        End Get
    End Property
    Private WithEvents SelectionRowStyle_ As New CellStyle With {.BackColor = Color.DarkSlateGray, .ShadeColor = Color.Gray, .ForeColor = Color.White, .Font = New Font("Century Gothic", 8)}
    Public ReadOnly Property SelectionRowStyle As CellStyle
        Get
            Return SelectionRowStyle_
        End Get
    End Property
    Public ReadOnly Property RowHeight As Integer
        Get
            Return RowStyle.Height
        End Get
    End Property
    Public ReadOnly Property Selected As List(Of Row)
        Get
            Return Where(Function(r) r.Selected).ToList
        End Get
    End Property
    Private SingleSelect_ As Boolean = True
    Public Property SingleSelect As Boolean
        Get
            Return SingleSelect_
        End Get
        Set(value As Boolean)
            If value <> SingleSelect_ Then
                SingleSelect_ = value
                If value Then
                    For Each SelectedRow In Selected.Skip(1)
                        SelectedRow.Selected = False
                    Next
                End If
                Parent.Invalidate()
            End If
        End Set
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Shadows Function Add(ByVal newRow As Row) As Row

        If newRow IsNot Nothing Then
            newRow._Parent = Me
            MyBase.Add(newRow)
            DrawTimer.Start()
        End If
        Return newRow

    End Function
    Public Sub SortBy(ByVal Column As Column)

        If Column IsNot Nothing Then
            With Column
                Select Case .Format.Key
                    Case Column.TypeGroup.Strings, Column.TypeGroup.Images
                        If .SortOrder = SortOrder.Ascending Then Sort(Function(x, y) String.Compare(x.Cells(.Name).Text, y.Cells(.Name).Text, StringComparison.Ordinal))
                        If .SortOrder = SortOrder.Descending Then Sort(Function(y, x) String.Compare(x.Cells(.Name).Text, y.Cells(.Name).Text, StringComparison.Ordinal))

                    Case Column.TypeGroup.Integers
                        If .SortOrder = SortOrder.Ascending Then Sort(Function(x, y) x.Cells(.Name).ValueWhole.CompareTo(y.Cells(.Name).ValueWhole))
                        If .SortOrder = SortOrder.Descending Then Sort(Function(y, x) x.Cells(.Name).ValueWhole.CompareTo(y.Cells(.Name).ValueWhole))


                    Case Column.TypeGroup.Decimals
                        If .SortOrder = SortOrder.Ascending Then Sort(Function(x, y) x.Cells(.Name).ValueDecimal.CompareTo(y.Cells(.Name).ValueDecimal))
                        If .SortOrder = SortOrder.Descending Then Sort(Function(y, x) x.Cells(.Name).ValueDecimal.CompareTo(y.Cells(.Name).ValueDecimal))

                    Case Column.TypeGroup.Dates, Column.TypeGroup.Times
                        If .SortOrder = SortOrder.Ascending Then Sort(Function(x, y) x.Cells(.Name).ValueDate.CompareTo(y.Cells(.Name).ValueDate))
                        If .SortOrder = SortOrder.Descending Then Sort(Function(y, x) x.Cells(.Name).ValueDate.CompareTo(y.Cells(.Name).ValueDate))

                    Case Column.TypeGroup.Booleans
                        If .SortOrder = SortOrder.Ascending Then Sort(Function(x, y) x.Cells(.Name).ValueBoolean.CompareTo(y.Cells(.Name).ValueBoolean))
                        If .SortOrder = SortOrder.Descending Then Sort(Function(y, x) x.Cells(.Name).ValueBoolean.CompareTo(y.Cells(.Name).ValueBoolean))

                End Select
            End With
            Parent.Invalidate()
        End If

    End Sub
    Private Sub RowStyle_PropertyChanged(sender As Object, e As StyleEventArgs) Handles RowStyle_.PropertyChanged, AlternatingRowStyle_.PropertyChanged, SelectionRowStyle_.PropertyChanged, HeaderStyle_.PropertyChanged

        If Not (sender Is Nothing Or e Is Nothing) Then
            If e.ChangedProperty = CellStyle.Properties.Height Then
                If sender Is RowStyle Then
                    AlternatingRowStyle_.Height = RowStyle.Height
                    SelectionRowStyle_.Height = RowStyle.Height

                ElseIf sender Is AlternatingRowStyle Then
                    RowStyle_.Height = AlternatingRowStyle.Height
                    SelectionRowStyle_.Height = AlternatingRowStyle.Height

                ElseIf sender Is SelectionRowStyle Then
                    AlternatingRowStyle_.Height = SelectionRowStyle.Height
                    RowStyle_.Height = SelectionRowStyle.Height

                End If
            ElseIf e.ChangedProperty = CellStyle.Properties.ImageScaling Then
                If sender Is RowStyle Then
                    AlternatingRowStyle_.ImageScaling = RowStyle.ImageScaling
                    SelectionRowStyle_.ImageScaling = RowStyle.ImageScaling

                ElseIf sender Is AlternatingRowStyle Then
                    RowStyle_.ImageScaling = AlternatingRowStyle.ImageScaling
                    SelectionRowStyle_.ImageScaling = AlternatingRowStyle.ImageScaling

                ElseIf sender Is SelectionRowStyle Then
                    AlternatingRowStyle_.ImageScaling = SelectionRowStyle.ImageScaling
                    RowStyle_.ImageScaling = SelectionRowStyle.ImageScaling

                End If
            End If
        End If
        Parent?.Invalidate()

    End Sub
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                RowStyle_.Dispose()
                AlternatingRowStyle_.Dispose()
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
    Public Sub New(columnNames As List(Of String), rowValues As Object())

        If columnNames IsNot Nothing And rowValues IsNot Nothing Then
            For column = 0 To columnNames.Count - 1
                Dim columnName As String = columnNames(column)
                Cells.Add(columnName, New Cell(Me, columnName, column, rowValues(column)))
            Next
        End If

    End Sub
    <NonSerialized> Friend _Parent As RowCollection
    Public ReadOnly Property Parent As RowCollection
        Get
            Return _Parent
        End Get
    End Property
    Public ReadOnly Property Cells As New Dictionary(Of String, Cell)
    Public ReadOnly Property Index As Integer
        Get
            Return If(Parent Is Nothing, -1, Parent.IndexOf(Me))
        End Get
    End Property
    Public Property Tag As Object
    Private WithEvents Style_ As New CellStyle With {
        .BackColor = Color.Transparent,
        .ShadeColor = Color.White,
        .ForeColor = Color.Black,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.None
    }
    Public ReadOnly Property Style As CellStyle
        Get
            Return Style_
        End Get
    End Property
    Private WithEvents SelectionStyle_ As New CellStyle With {
        .BackColor = Color.DarkSlateGray,
        .ShadeColor = Color.Gray,
        .ForeColor = Color.White,
        .Font = New Font("Century Gothic", 8),
        .Theme = Theme.None
    }
    Public ReadOnly Property SelectionStyle As CellStyle
        Get
            Return SelectionStyle_
        End Get
    End Property
    Friend ReadOnly Property StyleChanged As Boolean
    Private _Selected As Boolean
    Public Property Selected As Boolean
        Get
            If Not Parent.Parent.FullRowSelect Then _Selected = False
            Return _Selected
        End Get
        Set(value As Boolean)
            With Parent?.Parent
                If .FullRowSelect And _Selected <> value Then
                    _Selected = value
                    If value Then
                        If Parent.SingleSelect Then
                            For Each Row In Parent.Except({Me}).Where(Function(r) r.Selected)
                                Row.Selected = False
                            Next
                        End If
                    End If
                    .Invalidate()
                End If
            End With

        End Set
    End Property
    Private Sub RowStyle_PropertyChanged(sender As Object, e As StyleEventArgs) Handles Style_.PropertyChanged
        Using defaultStyle As New CellStyle With {.BackColor = Color.Transparent, .ShadeColor = Color.White, .ForeColor = Color.Black, .Font = New Font("Century Gothic", 8)}
            _StyleChanged = Style_ <> defaultStyle
        End Using
        Parent?.Parent?.Invalidate()
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
'- WORK IN PROGRESS ... MUST MIRROR DataRow.ItemArray
<Serializable()> Public Class Cell
    Implements IDisposable
    Private IsNew As Boolean = True
    Public Sub New(cellParent As Row, columnName As String, columnIndex As Integer, cellValue As Object)
        Parent = cellParent
        Name = columnName
        Index = columnIndex
        Value = If(IsDBNull(cellValue), Nothing, cellValue)
    End Sub
    Public ReadOnly Property Parent As Row
    Friend ReadOnly Property DataType As Type
    Friend ReadOnly Property FormatData As KeyValuePair(Of Column.TypeGroup, String)
    Public Property TipText As String
    Private Value_ As Object
    Public Property Value As Object
        Get
            Return Value_
        End Get
        Set(newValue As Object)
            If Value_ IsNot newValue Then

                _DataType = GetDataType(newValue)

                Select Case _DataType
                    Case GetType(String)

                    Case GetType(Double), GetType(Decimal)
                        _ValueDecimal = CType(newValue, Double)
                        Dim CultureInfo = New Globalization.CultureInfo("en-US")
                        With CultureInfo.NumberFormat
                            .CurrencyGroupSeparator = ","
                            .NumberDecimalDigits = 2
                        End With

                    Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                        _ValueWhole = CType(newValue, Long)

                    Case GetType(Boolean)
                        _ValueBoolean = CType(newValue, Boolean)
                        Dim checkImages As New Dictionary(Of String, String) From {
                         {"checkedBlack", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAABwSURBVDhPlZABDoAgCEWh+x+0WxifxIHDgrcxBXxiMRENiTYqCljLMPO45r5NWcQUxExrogn+k37FTAJL3J8CThJYojXt8JcEwlOdfG+5hgeZ9OOtmOZrJklNV/TTn2NSNslANdxe4Tixgk58tx2IHtIlOgxG8FAIAAAAAElFTkSuQmCC"},
                         {"uncheckedBlack", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAAAtSURBVDhPY2RgYPgPxCQDsEYgANFEA0ZGxv9MUDbJYFQjHjCqEQ8gM5EzMAAAoBMHFwfr1LQAAAAASUVORK5CYII="},
                         {"checkedWhite", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAABnSURBVDhPxZJRDoAgDEM3739nXNEmbCko/vi+XOhLYdFbYB+g6Nf4mnbcH9ukRvdcXF7BAaHciCDDC6kjr/okgVFEIBmBlMAo8pDhqQTqVavcZyytLk69kQnZRORygmkT+e/P2cTsBCdlLwZDKAEtAAAAAElFTkSuQmCC"},
                         {"uncheckedWhite", "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsEAAA7BAbiRa+0AAAAnSURBVDhPY/wPBAxkAJhGRgiXaPCfCcogGYxqxANGNeIBZCZyBgYAk5cNDhG2VLEAAAAASUVORK5CYII="}
                         }
                        Dim useWhite As Boolean = False 'If(Style Is Nothing, False, BackColorToForeColor(Style.BackColor) = Color.White)
                        Dim imageName As String = If(ValueBoolean, String.Empty, "un") & "checked" & If(useWhite, "White", "Black")
                        _ValueImage = Base64ToImage(checkImages(imageName))

                    Case GetType(Date), GetType(DateAndTime)
                        _ValueDate = CType(newValue, Date)

                    Case GetType(Image), GetType(Icon)
                        If newValue IsNot Nothing Then
                            _ValueImage = If(DataType = GetType(Icon), CType(newValue, Icon).ToBitmap, CType(newValue, Bitmap))
                            If Parent?.Parent?.RowStyle IsNot Nothing Then
                                With Parent.Parent.RowStyle
                                    'RowStyle is the master - SelectionRowStyle and AlternatingRowStyle must follow Scaling and Height 
                                    If .ImageScaling = Scaling.GrowParent Then .Height = { .Height, ValueImage.Height}.Max
                                End With
                            End If
                        End If

                End Select
                Value_ = newValue
                Dim existingFormat = FormatData.Key
                _FormatData = Column.Get_kvpFormat(_DataType)
                If Not IsNew And existingFormat <> _FormatData.Key Then Column.DataType = DataType
                Parent?.Parent?.Parent.Invalidate()
            End If
            IsNew = False
        End Set
    End Property
    Public ReadOnly Property ValueDecimal As Double
    Public ReadOnly Property ValueWhole As Long
    Public ReadOnly Property ValueImage As Image
    Public ReadOnly Property ValueBoolean As Boolean
    Public ReadOnly Property ValueDate As Date
    Public ReadOnly Property Column As Column
        Get
            Return Parent?.Parent?.Parent?.Columns.Item(Name)
        End Get
    End Property
    Public ReadOnly Property Name As String
    Public ReadOnly Property Text As String
        Get
            If Value Is Nothing Then
                Return Nothing
            Else
                Select Case Column.DataType
                    Case GetType(String)
                        Return Value.ToString

                    Case GetType(Double), GetType(Decimal)
                        _ValueDecimal = CType(Value, Double)
                        Dim CultureInfo = New Globalization.CultureInfo("en-US")
                        With CultureInfo.NumberFormat
                            .CurrencyGroupSeparator = ","
                            .NumberDecimalDigits = 2
                        End With
                        Return CType(Value, Double).ToString("N", CultureInfo)

                    Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                        _ValueWhole = CType(Value, Long)
                        Return Format(Value, FormatData.Value)

                    Case GetType(Boolean)
                        'Value could be either "Y"|"N"|"True"|"False"|0|1|True|False ... If Column.DataType is Boolean with String Or Integer values, likely from Columns.ColumnWidth
                        _ValueBoolean = {"Y", "TRUE", "1"}.Contains(Value.ToString.ToUpperInvariant)
                        _ValueImage = Base64ToImage(If(ValueBoolean, CheckString, UnCheckString))
                        Return Value.ToString

                    Case GetType(Date), GetType(DateAndTime)
                        _ValueDate = CType(Value, Date)
                        Return Format(Value, FormatData.Value)

                    Case GetType(Image)
                        'Column.DataType is Image ... Cell.FormatData might not be
                        If FormatData.Key = Column.TypeGroup.Images Then
                            'Cell.FormatData = Column.Format - compatible
                            _ValueImage = CType(Value, Bitmap)
                            Return ImageToBase64(ValueImage)
                        Else
                            'Cell.FormatData <> Column.Format - incompatible
                            _ValueImage = Nothing
                            Return Nothing
                        End If

                    Case GetType(Icon)
                        'Column.DataType is Icon ... Cell.FormatData might not be
                        If FormatData.Key = Column.TypeGroup.Images Then
                            'Cell.FormatData = Column.Format - compatible
                            _ValueImage = CType(Value, Icon)?.ToBitmap
                            Return ImageToBase64(ValueImage)
                        Else
                            'Cell.FormatData <> Column.Format - incompatible
                            _ValueImage = Nothing
                            Return Nothing
                        End If

                    Case Else
                        Return Value.ToString

                End Select
            End If
        End Get
    End Property
    Public ReadOnly Property Index As Integer
    Private _Selected As Boolean
    Public Property Selected As Boolean
        Get
            Return _Selected
        End Get
        Set(value As Boolean)
            If _Selected <> value Then
                _Selected = value
                If value Then
                    With Parent

                    End With
                End If
            End If
        End Set
    End Property
    Public Shadows ReadOnly Property ToString As String
        Get
            Return Join({Name, Text, DataType.ToString}, ", ")
        End Get
    End Property
    Public ReadOnly Property Style As CellStyle
        Get
            If Parent.Parent Is Nothing Then
                Return Nothing
            Else
                Return If(Selected, Parent.Parent.SelectionRowStyle, If(Index Mod 2 = 0, Parent.Parent.RowStyle, Parent.Parent.AlternatingRowStyle))
            End If
        End Get
    End Property
#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                _Parent.Dispose()
                _ValueImage?.Dispose()

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
    Public ReadOnly Property ChangedProperty As CellStyle.Properties
    Public ReadOnly Property PropertyValue As Object
    Public Sub New(changedProperty As CellStyle.Properties, changedValue As Object)
        Me.ChangedProperty = changedProperty
        PropertyName = changedProperty.ToString
        PropertyValue = changedValue
    End Sub
End Class
<Serializable()> <TypeConverter(GetType(PropertyConverter))> Public Class CellStyle
    Implements IEquatable(Of CellStyle)
    Implements IDisposable
    Public Event PropertyChanged(sender As Object, e As StyleEventArgs)
    <Flags> Public Enum Properties
        Font
        BackColor
        ForeColor
        ShadeColor
        Theme
        Alignment
        ImageScaling
        Height
        Padding
    End Enum
    Public Sub New()
        Height = Padding.Top + CInt(Font.GetHeight) + Padding.Bottom
    End Sub

    Private _Font As New Font("Century Gothic", 9)
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Appearance")>
    <Description("Specifies the object Font")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Font As Font
        Get
            Return _Font
        End Get
        Set(ByVal value As Font)
            If value IsNot _Font Then
                _Font = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.Font, value))
            End If
        End Set
    End Property

    Private _BackColor As Color = Color.Gainsboro
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Color")>
    <Description("Specifies the object BackColor")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property BackColor As Color
        Get
            Return _BackColor
        End Get
        Set(ByVal value As Color)
            If value <> _BackColor Then
                _BackColor = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.BackColor, value))
            End If
        End Set
    End Property

    Private _ForeColor As Color = Color.Black
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Color")>
    <Description("Specifies the object ForeColor")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property ForeColor As Color
        Get
            If Theme = Theme.None Then
                Return _ForeColor
            Else
                Return GlossyForecolor(Theme)
            End If
        End Get
        Set(ByVal value As Color)
            If value <> _ForeColor Then
                _ForeColor = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.ForeColor, value))
            End If
        End Set
    End Property

    Private _ShadeColor As Color = Color.WhiteSmoke
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Color")>
    <Description("Specifies the object Shading color")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property ShadeColor As Color
        Get
            Return _ShadeColor
        End Get
        Set(ByVal value As Color)
            If value <> _ShadeColor Then
                _ShadeColor = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.ShadeColor, value))
            End If
        End Set
    End Property
    Private _Theme As Theme = Theme.None
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Color")>
    <Description("Specifies the object Shading color")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Theme As Theme
        Get
            Return _Theme
        End Get
        Set(ByVal value As Theme)
            If value <> _Theme Then
                _Theme = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.Theme, value))
            End If
        End Set
    End Property

    <NonSerialized> Private _Alignment As StringFormat = New StringFormat With {.LineAlignment = StringAlignment.Center,
        .Alignment = StringAlignment.Center}
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Layout")>
    <Description("Specifies the object Alignment")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Alignment As StringFormat
        Get
            Return _Alignment
        End Get
        Set(ByVal value As StringFormat)
            If value IsNot _Alignment Then
                _Alignment = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.Alignment, value))
            End If
        End Set
    End Property

    Private _ImageScaling As Scaling = Scaling.ShrinkChild
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Layout")>
    <Description("Specifies the object Grow style")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property ImageScaling As Scaling
        Get
            Return _ImageScaling
        End Get
        Set(ByVal value As Scaling)
            If value <> _ImageScaling Then
                _ImageScaling = value
                If value = Scaling.GrowParent Then _Height = {Padding.Top + CInt(Font.GetHeight) + Padding.Bottom, Height}.Max
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.ImageScaling, value))
            End If
        End Set
    End Property

    Private _Height As Integer
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Layout")>
    <Description("Specifies the object Height")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Height As Integer
        Get
            Return _Height
        End Get
        Set(ByVal value As Integer)
            If value <> _Height Then
                _Height = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.Height, value))
            End If
        End Set
    End Property

    Private _Padding As Padding = New Padding(2)
    <Browsable(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    <Category("Layout")>
    <Description("Specifies the padding")>
    <RefreshProperties(RefreshProperties.All)>
    Public Property Padding As Padding
        Get
            Return _Padding
        End Get
        Set(ByVal value As Padding)
            If value <> _Padding Then
                _Padding = value
                RaiseEvent PropertyChanged(Me, New StyleEventArgs(Properties.Padding, value))
            End If
        End Set
    End Property

    Public Overrides Function GetHashCode() As Integer
        Return Font.GetHashCode Xor BackColor.GetHashCode Xor ForeColor.GetHashCode Xor ShadeColor.GetHashCode Xor Alignment.GetHashCode Xor ImageScaling.GetHashCode Xor Height.GetHashCode Xor Padding.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As CellStyle) As Boolean Implements IEquatable(Of CellStyle).Equals
        If other Is Nothing Then
            Return False
        Else
            Return BackColor = other.BackColor And Font.FontFamily.Name = other.Font.FontFamily.Name And Font.Size = other.Font.Size And Font.Style = other.Font.Style And ForeColor = other.ForeColor And ShadeColor = other.ShadeColor And Theme = other.Theme And Alignment.Alignment = other.Alignment.Alignment And Alignment.LineAlignment = other.Alignment.LineAlignment And ImageScaling = other.ImageScaling And Padding = other.Padding
        End If
    End Function
    Public Shared Operator =(ByVal Object1 As CellStyle, ByVal Object2 As CellStyle) As Boolean
        If Object1 Is Nothing Then
            Return Object2 Is Nothing
        ElseIf Object2 Is Nothing Then
            Return Object1 Is Nothing
        Else
            Return Object1.Equals(Object2)
        End If
    End Operator
    Public Shared Operator <>(ByVal Object1 As CellStyle, ByVal Object2 As CellStyle) As Boolean
        Return Not Object1 = Object2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is CellStyle Then
            Return CType(obj, CellStyle) = Me
        Else
            Return False
        End If
    End Function

#Region "IDisposable Support"
    Private DisposedValue As Boolean ' To detect redundant calls IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not DisposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                _Font.Dispose()
                _Alignment.Dispose()
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
Public NotInheritable  Class WaitTimer
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
    Public Property Limit As Integer
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
    Public Sub Increment()

        TimerTicks += CInt(100 / {Limit, 1}.Max)
        TimerTicks = {100, TimerTicks}.Min
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