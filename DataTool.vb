Option Explicit On
Option Strict On
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.ComponentModel
Imports System.Runtime.InteropServices

#Region " IMPROVEMENTS "
'SPEED - UPDATE ONLY WHEN NECESSARY AND USE THREADING

'[0] MOVING DATA - DRAG n DROP NOT WORKING RIGHT

'[1] PASTE LIST TO ACTIVEPANE
'[2] ADD PARAMETERS AS: ?INPUT
'[4] SEARCH DATABASE FOR TABLES Or COLUMNS WITH COLUMN / TABLE NAME
'[5] CANCEL QUERY
'[6] MODIFY, ADD, DELETE CONNECTIONS ... RIGHT CLICK PANE, ADD TO TSMI
'[7] "DRIVER={IBM DB2 ODBC DRIVER};Database=DSNA1;Hostname=sbrysa1.somers.hqregion.ibm.com;Port=5000;Protocol=TCPIP;UID=C085365;PWD=Y0Y0Y0Y0"
#End Region
Public Class ScriptsEventArgs
    Inherits EventArgs
    Public ReadOnly Property Item As Script
    Public ReadOnly Property State As CollectionChangeAction
    Public Sub New(Item As Script, State As CollectionChangeAction)
        Me.Item = Item
        Me.State = State
    End Sub
End Class
<ComVisible(False)> Public Class ScriptCollection
    Inherits List(Of Script)
#Region " DECLARATIONS "
    Private Initialized As Boolean
    Private ReadOnly Scripts_DirectoryInfo As DirectoryInfo = Directory.CreateDirectory(MyDocuments & "\DataManager\Scripts")
    Private WithEvents ChangeTimer As New Timer With {.Interval = 500}
#End Region
    Public Event Alert(sender As Object, e As AlertEventArgs)
    Public Event CollectionChanged(sender As Object, e As ScriptsEventArgs)
    Public Sub New(Parent As DataTool)

        Me.Parent = Parent
        'If Parent IsNot Nothing Then Add(New Script With {._Tabs = Parent.Script_Tabs, .State = Script.ViewState.OpenDraft})
        With New BackgroundWorker
            AddHandler .DoWork, AddressOf NewWorkStart
            AddHandler .RunWorkerCompleted, AddressOf NewWorkEnd
            .RunWorkerAsync()
        End With

    End Sub
    Private Sub NewWorkStart(sender As Object, e As DoWorkEventArgs)

        With DirectCast(sender, BackgroundWorker)
            RemoveHandler .DoWork, AddressOf NewWorkStart
        End With
        Dim fileScripts = ReadFiles(Scripts_DirectoryInfo.FullName)
        Dim fileConnections = (From fs In fileScripts Group fs By dsn = Split(fs.Value, vbNewLine).First Into dsnGroup = Group
                               Select New With {.server = dsn, .items = dsnGroup.ToList}).ToDictionary(Function(k) k.server, Function(v) v.items)

        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
        Dim Testing As Boolean = Parent.TestMode
        Dim takeCount As Integer = If(Testing, 25, 10000)
        Dim eachCount As Integer = CInt(Math.Ceiling(takeCount / fileConnections.Count))
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

        Dim smallestCountToLargest = fileConnections.OrderBy(Function(fc) fc.Value.Count).ToList
        For Each fileConnection In smallestCountToLargest
            If fileConnection.Key Is smallestCountToLargest.Last.Key Then eachCount += {takeCount - eachCount - Count, 0}.Max
            For Each fileScript In fileConnection.Value.Take(eachCount)
                Dim NewScript As New Script(fileScript.Key)
                RaiseEvent Alert(Me, New AlertEventArgs(Strings.Join({"Loading", NewScript.Name})))
                Do While NewScript.Body.IsBusy
                Loop
                Add(NewScript)
            Next
        Next

    End Sub
    Private Sub NewWorkEnd(sender As Object, e As RunWorkerCompletedEventArgs)
        With DirectCast(sender, BackgroundWorker)
            Initialized = True
            RemoveHandler .RunWorkerCompleted, AddressOf NewWorkEnd
            RaiseEvent Alert(Me, New AlertEventArgs(Count & " scripts loaded"))
            RaiseEvent CollectionChanged(Me, New ScriptsEventArgs(Nothing, CollectionChangeAction.Refresh))
        End With
    End Sub
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Public ReadOnly Property Parent As DataTool
    Private ReadOnly Connections As ConnectionCollection = DataTool.Connections
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Shadows Function Add(Item As Script) As Script

        ChangeTimer.Stop()
        ChangeTimer.Start()
        MyBase.Add(Item)
        If Item IsNot Nothing Then
            With Item
                .Parent = Me
                ._Tabs = Parent.Script_Tabs
                If .ToString.Contains("DSN=") Then .Connection = Connections.Item(.Connection.ToString)
            End With
            If Initialized Then SortCollection()
            RaiseEvent CollectionChanged(Me, New ScriptsEventArgs(Item, CollectionChangeAction.Add))
        End If
        Return Item

    End Function
    Public Shadows Function Remove(Item As Script) As Script

        If Item IsNot Nothing Then
            ChangeTimer.Stop()
            ChangeTimer.Start()
            MyBase.Remove(Item)
            Item.Parent = Nothing
            If Initialized Then SortCollection()
            RaiseEvent CollectionChanged(Me, New ScriptsEventArgs(Item, CollectionChangeAction.Remove))
        End If
        Return Item

    End Function
    Public Shadows Function Item(DataSource As String, Name As String) As Script

        Dim _Items As New List(Of Script)(Where(Function(x) x.DataSourceName = DataSource And x.Name = Name))
        _Items.Sort(Function(x, y) x.Created.CompareTo(y.Created))      'Sort oldest first
        If _Items.Any Then
            Return _Items.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Item(Created As Date) As Script

        Dim DateString As String = DateTimeToString(Created)
        Dim Items As New List(Of Script)(From m In Me Where m.CreatedString = DateString)
        Items.Sort(Function(x, y) x.Created.CompareTo(y.Created))      'Sort oldest first
        If Items.Any Then
            Return Items.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Item(Value As String) As Script

        If Value Is Nothing Then
            Return Nothing
        Else
            Dim _Name As String = Value
            If Value.Contains(Delimiter) Then
                REM /// COMES FROM MY.SETTINGS.SCRIPTS...EXTRACT NAME
                _Name = Split(Value, Delimiter)(1)
            End If
            Dim _Items As New List(Of Script)(Where(Function(x) x.Name = _Name))
            _Items.Sort(Function(x, y) x.Created.CompareTo(y.Created))      'Sort oldest first
            If _Items.Any Then
                Return _Items.First
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Shadows Function Item(ScriptItem As Script) As Script

        Dim _Items As New List(Of Script)(Where(Function(x) x.Created = ScriptItem.Created))
        _Items.Sort(Function(x, y) x.Created.CompareTo(y.Created))      'Sort oldest first
        If _Items.Any Then
            Return _Items.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Item(Tab As Tab) As Script

        Dim _Tabs = From SI In Me Where SI.Tab Is Tab
        If _Tabs.Any Then
            Return _Tabs.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Item(State As Script.ViewState) As Script

        Dim _Items As New List(Of Script)(Where(Function(x) x.State = State))
        _Items.Sort(Function(x, y) x.Created.CompareTo(y.Created))      'Sort oldest first
        If _Items.Any Then
            Return _Items.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Contains(DataSource As String, Name As String) As Boolean
        Return Not IsNothing(Item(DataSource, Name))
    End Function
    Public Shadows Function Contains(ItemX As String) As Boolean
        Return Not IsNothing(Item(ItemX))
    End Function
    Public Shadows Function Contains(ItemX As Script) As Boolean
        Return Not IsNothing(Item(ItemX))
    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub ChangeTimerTick() Handles ChangeTimer.Tick
        ChangeTimer.Stop()
        SortCollection()
        Initialized = True
    End Sub
    Public Sub SortCollection()

        Sort(Function(f1, f2)
                 Dim Level1 = String.Compare(f1.DataSourceName, f2.DataSourceName, StringComparison.InvariantCulture)
                 If Level1 <> 0 Then
                     Return Level1
                 Else
                     Dim Level2 = String.Compare(f1.Name, f2.Name, StringComparison.InvariantCulture)
                     Return Level2
                 End If
             End Function)

    End Sub
    Public Sub View()

        SortCollection()
        Using Message As New Prompt
            Using DT As New DataTable
                With DT
                    .Columns.Add(New DataColumn With {.ColumnName = "DSN", .DataType = GetType(String)})
                    .Columns.Add(New DataColumn With {.ColumnName = "Name", .DataType = GetType(String)})
                    .Columns.Add(New DataColumn With {.ColumnName = "Created", .DataType = GetType(Date)})
                    .Columns.Add(New DataColumn With {.ColumnName = "Modified", .DataType = GetType(Date)})
                    .Columns.Add(New DataColumn With {.ColumnName = "Ran", .DataType = GetType(Date)})
                    For Each ScriptItem In Where(Function(s) s.State = Script.ViewState.ClosedSaved)
                        With ScriptItem
                            DT.Rows.Add({ .DataSourceName, .Name, .Created, .Modified, .Ran})
                        End With
                    Next
                End With
                Message.Datasource = DT
            End Using
            Message.Show("Scripts Count=" & Count, "Saved Scripts", Prompt.IconOption.Warning, Prompt.StyleOption.Grey)
        End Using

    End Sub
    Public Function ToStringArray() As String()
        Return (From m In Me Where m.FileCreated Select m.ToString & String.Empty).ToArray
    End Function
    Public Overrides Function ToString() As String
        Return Strings.Join(ToStringArray, vbNewLine)
    End Function
    Public Function ToStringList() As List(Of String)
        Return ToStringArray.ToList
    End Function
    Public Function ToStringCollection() As Specialized.StringCollection
        Dim SSC As New Specialized.StringCollection()
        SSC.AddRange(ToStringArray)
        Return SSC
    End Function
#End Region
End Class
<Serializable> Public Class Script
    Implements IDisposable
#Region " DISPOSE "
    Dim disposed As Boolean = False
    <NonSerialized> Private ReadOnly Handle As Runtime.InteropServices.SafeHandle = New Microsoft.Win32.SafeHandles.SafeFileHandle(IntPtr.Zero, True)
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposed Then Return
        If disposing Then
            Handle.Dispose()
            ' Free any other managed objects here.
            If _Pane IsNot Nothing Then _Pane.Dispose()
            If _Tab IsNot Nothing Then _Tab.Dispose()
        End If
        disposed = True
    End Sub
#End Region
#Region " EVENTS "
    Friend Event ConnectionChanged(sender As Object, e As ConnectionChangedEventArgs)
    Friend Event TypeChanged(sender As Object, e As ScriptTypeChangedEventArgs)
    Friend Event StateChanged(sender As Object, e As ScriptStateChangedEventArgs)
    Friend Event NameChanged(sender As Object, e As ScriptNameChangedEventArgs)
#End Region
#Region " CLASSES - ENUMS - STRUCTURES "
    Public Enum ViewState
        None
        ClosedSaved
        ClosedNotSaved
        OpenDraft
        OpenSaved
    End Enum
    Public Enum SaveAction
        ChangeName
        ChangeContent
        UpdateExecutionTime
    End Enum
#End Region
    Private Sub InstructionTypeChanged(sender As Object, e As ScriptTypeChangedEventArgs) Handles Body.TypeChanged

        RaiseEvent TypeChanged(sender, e)
        If State = ViewState.OpenDraft Or State = ViewState.OpenSaved Then
            Tab.Image = If(e.CurrentType = ExecutionType.DDL, My.Resources.DDL,
                        If(e.CurrentType = ExecutionType.SQL, My.Resources.SQL,
                        My.Resources.QuestionMark))
            If State = ViewState.OpenDraft Then Tab.ItemText = Name
        End If

    End Sub
    Private Sub BodyConnectionChanged(sender As Object, e As ConnectionChangedEventArgs) Handles Body.ConnectionChanged

        If Connection Is Nothing Then
            'Only change the active connection once (initially). Let the User override to pick a connection if needed
            If e.NewConnection IsNot Nothing Then Connection = e.NewConnection
        End If

    End Sub
#Region " NEW "
    'New Instance
    Public Sub New()
        Created = Now
    End Sub
    'From Saved
    Public Sub New(ScriptPath As String)

        _Path = ScriptPath
        _State = ViewState.ClosedSaved
        _Name = GetFileNameExtension(ScriptPath).Key
        _Created = File.GetCreationTime(ScriptPath)
        Dim Elements As String() = FileElements()
        Dim DSN As String = "DSN=" & Elements.First     'File Shows DSN Name only and not a ConnectionString as DSN=...
        _Connection = Connections.Item(DSN)
        Text = Elements.Last

    End Sub
#End Region
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Public ReadOnly Property Root As DataTool
        Get
            Dim _Root As DataTool = Nothing
            If IsNothing(Parent) Then
            Else
                If IsNothing(Parent.Parent) Then
                Else
                    _Root = Parent.Parent
                End If
            End If
            Return _Root
        End Get
    End Property
    Public ReadOnly Property Form As Control
        Get
            Dim _Form As Control = Nothing
            If IsNothing(Root) Then
            Else
                _Form = Root.Parent
            End If
            Return _Form
        End Get
    End Property
    <NonSerialized> Friend Parent As ScriptCollection
    Friend SystemColumns As New List(Of ColumnProperties)
    <NonSerialized> Public WithEvents Body As New BodyElements
    <NonSerialized> Private ReadOnly Connections As ConnectionCollection = DataTool.Connections
    Private _Connection As Connection
    Public Property Connection As Connection
        Get
            Return _Connection
        End Get
        Set(value As Connection)
            If _Connection <> value Then
                Body.Connection = value
                Dim FormerValue As Connection = _Connection
                _Connection = value
                If Tab IsNot Nothing And value IsNot Nothing Then
                    SetSafeControlPropertyValue(Tab, "HeaderBackColor", value.BackColor)
                    SetSafeControlPropertyValue(Tab, "HeaderForeColor", value.ForeColor)
                    Tabs.Invalidate()
                End If
                Save(SaveAction.ChangeContent)
                RaiseEvent ConnectionChanged(Me, New ConnectionChangedEventArgs(FormerValue, value))
            End If
        End Set
    End Property
    Public ReadOnly Property DataSourceName As String
        Get
            REM ONLY ATTEMPT TO CHANGE IF CURRENT VALUE IS NOTHING
            If IsNothing(Connection) Then
                Return Nothing
            Else
                Return Connection.DataSource
            End If
        End Get
    End Property
    Private ReadOnly FileDivider As String = vbNewLine + StrDup(10, EmDash) + vbNewLine
    Public ReadOnly Property FileCreated As Boolean
        Get
            Return File.Exists(Path)
        End Get
    End Property
    Private Function FileElements() As String()

        Using SR As New StreamReader(Path)
            Dim FileContent As String = SR.ReadToEnd
            Return Split(FileContent, FileDivider)
        End Using

    End Function
    Public ReadOnly Property TextWasModified As Boolean
        Get
            Return If(FileCreated, Not FileTextMatchesText, Body.HasText)
        End Get
    End Property
    Public ReadOnly Property FileTextMatchesText As Boolean
        Get
            If FileCreated Then
                'VBNEWLINE IS USED AS CARRIAGE RETURN IN THE .txt FILE BUT vbLf IS USED IN THE PANE.TEXT
                Dim FileText As String = Replace(FileElements.Last, vbNewLine, vbLf)
                Return FileText = Text
            Else
                Return False
            End If
        End Get
    End Property
    Friend ReadOnly Property Created As Date
    Friend ReadOnly Property CreatedString As String
        Get
            Return DateTimeToString(Created)
        End Get
    End Property
    Private _Modified As Date = New Date
    Friend ReadOnly Property Modified As Date
        Get
            If Path Is Nothing Then
                Return _Modified
            Else
                Return File.GetLastWriteTime(Path)
            End If
        End Get
    End Property
    Private _Name As String
    Public Property Name As String
        Get
            If State = ViewState.OpenDraft And Parent IsNot Nothing Then
                Dim OpenDrafts As New List(Of Script)(From SI In Parent Where SI.State = ViewState.OpenDraft And SI.Body.InstructionType = Body.InstructionType)
                Dim SystemGeneratedName As String = Body.InstructionType.ToString & (1 + OpenDrafts.IndexOf(Me))
                _Name = SystemGeneratedName
            End If
            Return _Name
        End Get
        Set(value As String)
            If _Name <> value Then
                Dim NameMatch As Match = Regex.Match(value, "(?<=[012][0-9]:[0-5][0-9]:[0-5][0-9]\.[0-9]{3}§)", RegexOptions.IgnoreCase)
                If value IsNot Nothing AndAlso NameMatch.Success Then
#Region " NAME CHANGE CAME FROM IC_SaveAs (OPEN) -OR- Tree_ClosedScripts (CLOSED) "
                    Dim FormerName As String = _Name
                    _Name = Split(value, Delimiter).Last
                    If FileCreated And _Name = FormerName Then
                        'SIMPLE SAVE TEXT REQUEST... BUT DO ONLY IF TEXT WAS MODIFIED
                        If Not FileTextMatchesText Then Save(SaveAction.ChangeContent)
                        RaiseEvent NameChanged(Me, New ScriptNameChangedEventArgs(FormerName, _Name))
                    Else
#Region " CREATE NEW FILE "
                        'FileCreated + _Name = FormerName OR Not FileCreated
                        Dim NewName As String = _Name & ".txt"
                        Dim SourcePath As String = If(Path, MyDocuments & "\DataManager\Scripts\" & NewName)
                        Dim Directory = IO.Path.GetDirectoryName(SourcePath)
                        Dim DestinationPath As String = IO.Path.Combine(Directory, NewName)
                        Using message As New Prompt
                            If File.Exists(DestinationPath) AndAlso message.Show("File already exists", Join({"Replace", NewName, "with", FormerName, "?"}), Prompt.IconOption.YesNo, Prompt.StyleOption.Blue) = DialogResult.No Then
                                'CANCELLED...UNDO NAME CHANGE
                                _Name = FormerName
                            Else
                                'CREATING A NEW FILE ERASES THE FILE...NOT LIKE A FOLDER
                                _Path = DestinationPath
                                If FileCreated Then
                                    File.Move(SourcePath, DestinationPath)
                                Else
                                    State = ViewState.OpenSaved
                                End If
                                Save(SaveAction.ChangeContent)
                                RaiseEvent NameChanged(Me, New ScriptNameChangedEventArgs(FormerName, _Name))
                            End If
                        End Using
#End Region
                    End If
#End Region
                End If
                If Tab IsNot Nothing AndAlso Tab.Parent IsNot Nothing Then
                    Tab.ItemText = _Name
                End If
            End If
        End Set
    End Property
    Friend ReadOnly Property Path As String
    Private _Ran As Date = New Date
    Friend ReadOnly Property Ran As Date
        Get
            If Path Is Nothing Then
                Return _Ran
            Else
                Return File.GetLastAccessTime(Path)
            End If
        End Get
    End Property
    Private _State As New ViewState
    Friend Property State As ViewState
        Get
            Return _State
        End Get
        Set(value As ViewState)
            If _State <> value Then
                Dim FormerState As ViewState = _State
                Dim NewState As ViewState = value
                _State = value
                'Permutations (None|Dummy|Draft|OpenSaved|ClosedSaved)
                'Existing= 1 of 5 Options, New= 4 remaining
                'ExistingState=None + NewState=(Dummy|Draft|OpenSaved|ClosedSaved)
                'ExistingState=Dummy + NewState=(None|Draft|OpenSaved|ClosedSaved)
                'ExistingState=Draft + NewState=(None|Dummy|OpenSaved|ClosedSaved)
                'ExistingState=OpenSaved + NewState=(None|Dummy|Draft|ClosedSaved)
                'ExistingState=ClosedSaved + NewState=(None|Dummy|Draft|OpenSaved)
                Select Case True
#Region " From None aka IsNew"
                    Case FormerState = ViewState.None And NewState = ViewState.OpenDraft
                        AddControls()

                    Case FormerState = ViewState.None And NewState = ViewState.OpenSaved
                        REM /// NOT LIKELY
                        AddControls()

                    Case FormerState = ViewState.None And NewState = ViewState.ClosedSaved
                            REM /// NEW FROM MY.SETTINGS.SCRIPTS
#End Region
#Region " From Draft "
                    Case FormerState = ViewState.OpenDraft And NewState = ViewState.None
                        REM /// Discard (User clicked X and doesn't want to save work)
                        RemoveControls()

                    Case FormerState = ViewState.OpenDraft And NewState = ViewState.OpenSaved
                            REM /// Handled in Name Set since the only method to change from OpenDraft to OpenSaved is via IC_SaveAs -OR- Tree_ClosedScripts

                    Case FormerState = ViewState.OpenDraft And NewState = ViewState.ClosedSaved
                            REM /// No longer applies- OpenDraft can not become ClosedSaved, only OpenDraft to OpenSaved as immediately above

#End Region
#Region " From OpenSaved "
                    Case FormerState = ViewState.OpenSaved And NewState = ViewState.None
                            REM /// OpenSaved can not become None. Only ClosedSaved can become None

                    Case FormerState = ViewState.OpenSaved And NewState = ViewState.OpenDraft
                            REM /// Not Logical...Unsave?

                    Case FormerState = ViewState.OpenSaved And NewState = ViewState.ClosedSaved
                        REM /// Save Text changes
                        Save(SaveAction.ChangeContent)
                        RemoveControls()

                    Case FormerState = ViewState.OpenSaved And NewState = ViewState.ClosedNotSaved
                        REM /// Discard any Text changes...revert back to FileText
                        Text = FileElements().Last
                        _State = ViewState.ClosedSaved
                        RemoveControls()
#End Region
#Region " From ClosedSaved "
                    Case FormerState = ViewState.ClosedSaved And NewState = ViewState.None
                        REM /// Delete...Tree_ClosedScripts, NodeRemove Clicked
                        File.Delete(Path)
                        Parent.Remove(Me)

                    Case FormerState = ViewState.ClosedSaved And NewState = ViewState.OpenDraft
                            REM /// Not Logical

                    Case FormerState = ViewState.ClosedSaved And NewState = ViewState.OpenSaved
                        REM /// Script to become visible
                        REM /// Drop Dummy Tab
                        REM /// Add at TabIndex
                        AddControls()
#End Region
                End Select
                RaiseEvent StateChanged(Me, New ScriptStateChangedEventArgs(FormerState, NewState))
            End If
        End Set
    End Property
    Private _Text As String
    Friend Property Text() As String
        Get
            If Pane IsNot Nothing Then _Text = Pane.Text
            Return _Text
        End Get
        Set(value As String)
            _Text = value
            Body.Text = value
            If Pane IsNot Nothing Then Pane.Text = value
        End Set
    End Property
    <NonSerialized> Friend _Tabs As Tabs
    Public ReadOnly Property Tabs As Tabs
        Get
            Return _Tabs
        End Get
    End Property
    <NonSerialized> Private _Tab As Tab
    Public ReadOnly Property Tab As Tab
        Get
            If _Tab Is Nothing And _Tabs IsNot Nothing Then
                Dim _Page As Tab = _Tabs.TabPages.Item(Name)
                If _Page IsNot Nothing Then
                    _Tab = _Page
                End If
            End If
            Return _Tab
        End Get
    End Property
    <NonSerialized> Public WithEvents Pane As RicherTextBox
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Friend Shadows ReadOnly Property ToString As String
        Get
            If IsNothing(Connection) Then
                Return Join({String.Empty, Name, CreatedString, Text, DateTimeToString(Modified), DateTimeToString(Ran)}, Delimiter)
            Else
                Return Join({Connection.Key, Name, CreatedString, Text, DateTimeToString(Modified), DateTimeToString(Ran)}, Delimiter)
            End If
        End Get
    End Property
    Public Function Save(Action As SaveAction) As Boolean

        Dim ActionTime As Date = Now
        If Action = SaveAction.ChangeContent Then _Modified = ActionTime
        If Action = SaveAction.UpdateExecutionTime Then _Ran = ActionTime

        If Parent Is Nothing Then
            'STILL INITIALIZING
            Return False

        ElseIf Path Is Nothing Then
            'QUERY Or PROCEDURE SUCCEEDED
            Return False

        Else
            Dim ConnectionText As String = If(Connection Is Nothing, String.Empty, Connection.Properties("DSN"))
            Dim ScriptText As String = Regex.Replace(If(Text, String.Empty), "[\n\r]", vbNewLine)      'vbNewLine MAKES THE .txt FILE MUCH MORE READABLE
            Select Case Action
                Case SaveAction.ChangeContent
                    Using SW As New StreamWriter(Path)
                        SW.Write(Join({ConnectionText, ScriptText}, FileDivider))
                    End Using
                    File.SetLastWriteTime(Path, ActionTime)

                Case SaveAction.UpdateExecutionTime
                    File.SetLastAccessTime(Path, ActionTime)

            End Select
            Return True
        End If

    End Function
    Private Sub AddControls()

        '2 NEW OPEN...ADD TAB + PANE
        _Tab = New Tab With {.HeaderBackColor = If(Connection Is Nothing, Color.Gainsboro, Connection.BackColor),
                    .HeaderForeColor = If(Connection Is Nothing, Color.Black, Connection.ForeColor),
                    .ItemText = Name,
                    .Image = If(Body.InstructionType = ExecutionType.DDL, My.Resources.DDL,
                             If(Body.InstructionType = ExecutionType.SQL, My.Resources.SQL,
                             My.Resources.QuestionMark)),
                    .Tag = Me,
                    .AllowDrop = True}

        Pane = New RicherTextBox With {.Name = "Pane",
                    .Dock = DockStyle.Fill,
                    .Multiline = True,
                    .WordWrap = True,
                    .AllowDrop = True,
                    .AcceptsTab = True,
                    .Font = My.Settings.Font_Pane,
                    .Tag = Me,
                    .EnableAutoDragDrop = True,
                    .Text = Text}
        Pane.AllowDrop = True
        Tab.Controls.Clear()
        Tab.Controls.Add(Pane)

        Tabs.TabPages.Add(Tab)

    End Sub
    Private Sub RemoveControls()

        Pane.Controls.Clear()
        Tab.Controls.Clear()
        With Root.Script_Tabs
            .TabPages.Remove(Tab)
            .Invalidate()
        End With
        _Tab.Dispose()

    End Sub
    '=======================================================================================================================================
    '=========================================================    TIDY TEXT    =============================================================
    '=======================================================================================================================================
#End Region
End Class
Public Class DataTool
    Inherits Control
#Region " DECLARATIONS "
    Private ReadOnly DataDirectory As DirectoryInfo = Directory.CreateDirectory(MyDocuments & "\DataManager")
    Private ReadOnly Path_Columns As String = DataDirectory.FullName & "\Columns.txt"
    Private ReadOnly GothicFont As New Font("Century Gothic", 9, FontStyle.Regular)
    Private ReadOnly Message As New Prompt With {.Font = GothicFont}
    Private WithEvents TLP_PaneGrid As New TableLayoutPanel With {.Dock = DockStyle.Fill,
        .ColumnCount = 3,
        .RowCount = 1,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .AllowDrop = True,
        .Margin = New Padding(0),
        .Font = GothicFont}
    Private WithEvents TLP_Objects As New TableLayoutPanel With {.Dock = DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 3,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .Margin = New Padding(0),
        .Font = GothicFont}
    Friend WithEvents Script_Tabs As New Tabs With {.Dock = DockStyle.Fill,
        .UserCanAdd = True,
        .UserCanReorder = True,
        .MouseOverSelection = True,
        .AddNewTabColor = Color.Black,
        .Font = GothicFont,
        .Alignment = TabAlignment.Top,
        .Multiline = True,
        .Margin = New Padding(0),
        .SelectedTabColor = Color.Black}
    Private WithEvents Script_Grid As New DataViewer With {.Dock = DockStyle.Fill,
        .Font = GothicFont,
        .AllowDrop = True,
        .Margin = New Padding(0)}
    Private WithEvents Button_ObjectsSync As New Button With {.Dock = DockStyle.Fill,
        .Text = String.Empty,
        .TextImageRelation = TextImageRelation.Overlay,
        .Image = My.Resources.Sync,
        .ImageAlign = ContentAlignment.MiddleLeft,
        .Margin = New Padding(0),
        .Font = GothicFont}
    Private WithEvents IC_ObjectsSearch As New ImageCombo With {.Dock = DockStyle.Fill,
        .Text = String.Empty,
        .HintText = "Search Database",
        .Image = My.Resources.View,
        .Margin = New Padding(0),
        .Font = GothicFont}
    Private WithEvents Button_ObjectsClose As New Button With {.Dock = DockStyle.Fill,
        .Text = String.Empty,
        .TextImageRelation = TextImageRelation.Overlay,
        .Image = My.Resources.Close.ToBitmap,
        .ImageAlign = ContentAlignment.MiddleCenter,
        .Margin = New Padding(0),
        .Font = GothicFont}
    Private WithEvents Tree_Objects As New TreeViewer With {.Name = "Database Objects",
        .Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .DropHighlightColor = Color.Gold,
        .CheckBoxes = TreeViewer.CheckState.Mixed,
        .MultiSelect = True,
        .Font = GothicFont}
    Private WithEvents TSDD_SaveAs As New ToolStripDropDown With {.AutoClose = False,
        .Padding = New Padding(0),
        .DropShadowEnabled = True,
        .BackColor = Color.Firebrick,
        .Renderer = New CustomRenderer,
        .Font = GothicFont}
    Private WithEvents IC_SaveAs As New ImageCombo With {.Image = My.Resources.Save,
        .HintText = "Save Or Save As",
        .Size = New Size(200, 28),
        .Font = GothicFont}
    Private ReadOnly SaveAsHost As New ToolStripControlHost(IC_SaveAs)
    Private ReadOnly SaveAsItem As Integer = TSDD_SaveAs.Items.Add(SaveAsHost)
    Private WithEvents TSDD_ClosedScripts As New ToolStripDropDown With {.AutoSize = False,
        .AutoClose = False,
        .Padding = New Padding(0),
        .DropShadowEnabled = True,
        .BackColor = Color.Transparent,
        .Renderer = New CustomRenderer,
        .Font = GothicFont}
    Private WithEvents TLP_ClosedScripts As New TableLayoutPanel With {.Size = New Size(200, 200),
        .ColumnCount = 1,
        .RowCount = 1,
        .Font = GothicFont}
    Private ReadOnly TLPCSCS As Integer = TLP_ClosedScripts.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 300})
    Private ReadOnly TLPCSRS As Integer = TLP_ClosedScripts.RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 600})
    Private WithEvents Tree_ClosedScripts As New TreeViewer With {.Name = "Scripts",
        .AutoSize = True,
        .Margin = New Padding(0),
        .MouseOverExpandsNode = False,
        .Font = GothicFont}
    Private WithEvents TT_Tabs As New ToolTip With {.ToolTipIcon = ToolTipIcon.Info}
    Private WithEvents TT_GridTip As New ToolTip With {.ToolTipIcon = ToolTipIcon.Info}
    Private WithEvents CMS_ExcelSheets As New ContextMenuStrip With {.AutoClose = False,
        .AutoSize = True,
        .Margin = New Padding(0),
        .DropShadowEnabled = False,
        .BackColor = Color.WhiteSmoke,
        .ForeColor = Color.DarkViolet,
        .Font = GothicFont}
    Private ReadOnly OpenFileNode As Node = New Node With {.Text = "Open File",
        .Image = My.Resources.Folder,
        .AllowEdit = False,
        .AllowRemove = False,
        .AllowDragDrop = False,
        .Font = GothicFont}
    Private WithEvents OpenFile As New OpenFileDialog
    Private WithEvents SaveFile As New SaveFileDialog
    '-----------------------------------------
    Private WithEvents DragNode As Node
    Private ScriptsInitialized As Boolean
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private WithEvents TT_PaneTip As New ToolTip With {.ToolTipIcon = ToolTipIcon.Info}
    Private WithEvents CMS_PaneOptions As New ContextMenuStrip With {.AutoClose = False,
        .Padding = New Padding(0),
        .ImageScalingSize = New Size(15, 15),
        .DropShadowEnabled = True,
        .Renderer = New CustomRenderer,
        .BackColor = Color.Gainsboro,
        .Font = GothicFont}
    '-----------------------------------------
    Private WithEvents TSMI_Connections As New ToolStripMenuItem With {.Text = "Connections",
        .Image = My.Resources.Database.ToBitmap,
        .Font = GothicFont}
    Private ReadOnly TT_Submit As New ToolTip With {.ShowAlways = True, .ToolTipTitle = "New connection:"}
    Private WithEvents TSMI_Comment As New ToolStripMenuItem With {.Text = "Comment",
        .Image = My.Resources.Comment,
        .Font = GothicFont}
    '-----------------------------------------
    Private WithEvents TSMI_Copy As New ToolStripMenuItem With {.Text = "Copy",
        .Image = My.Resources.Clipboard,
        .Font = GothicFont}
    Private WithEvents TSMI_CopyPlainText As New ToolStripMenuItem With {.Text = "Without format",
        .Image = My.Resources.txt,
        .Font = GothicFont}
    Private WithEvents TSMI_CopyColorText As New ToolStripMenuItem With {.Text = "With format",
        .Image = My.Resources.Colors,
        .Font = GothicFont}
    '-----------------------------------------
    Private WithEvents TSMI_Divider As New ToolStripMenuItem With {.Text = "Insert divider",
        .Image = My.Resources.InsertBefore,
        .Font = GothicFont}
    Private WithEvents TSMI_DividerSingle As New ToolStripMenuItem With {.Text = "Single line",
        .Image = My.Resources.Zap,
        .Font = GothicFont}
    Private WithEvents TSMI_DividerDouble As New ToolStripMenuItem With {.Text = "Double line",
        .Image = My.Resources.Zap,
        .Font = GothicFont}
    '-----------------------------------------
    Private WithEvents TSMI_Font As New ToolStripMenuItem With {.Text = "Font",
        .Image = My.Resources.Info,
        .Font = GothicFont}
    Private WithEvents Dialogue_Font As New FontDialog With {.Font = My.Settings.Font_Pane}
    '-----------------------------------------
    Private WithEvents TSMI_ObjectType As New ToolStripMenuItem With {.Text = String.Empty,
        .Image = My.Resources.Info,
        .ImageScaling = ToolStripItemImageScaling.None,
        .Font = GothicFont}
    Private ReadOnly TSMI_ObjectValue As New ToolStripMenuItem With {.Text = String.Empty,
        .Image = My.Resources._Property,
        .ImageScaling = ToolStripItemImageScaling.None,
        .Font = GothicFont}
    Private ReadOnly TSMI_TipSwitch As New ToolStripMenuItem With {.Text = "Tips On",
        .Image = My.Resources.LightOn,
        .Font = GothicFont}
    Private WithEvents IC_BackColor As New ImageCombo With {.Size = New Size(160, 28),
        .HintText = "BackColor",
        .Font = GothicFont}
    Private WithEvents IC_ForeColor As New ImageCombo With {.Size = New Size(160, 28),
        .HintText = "ForeColor",
        .Font = GothicFont}
    Private ReadOnly TLP_Type As New TableLayoutPanel With {.ColumnCount = 1,
        .RowCount = 2,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .Font = GothicFont}
    Private ReadOnly TSCH_TypeHost As New ToolStripControlHost(TLP_Type) With {.ImageScaling = ToolStripItemImageScaling.None}
    Private ReadOnly A2 As Integer = TSMI_ObjectType.DropDownItems.Add(TSCH_TypeHost)
    Private WithEvents FindAndReplace As New FindReplace With {.Font = GothicFont}
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private ReadOnly ObjectsSet As New DataSet With {.DataSetName = "Objects"}
    Private WithEvents ObjectsWorker As New BackgroundWorker With {.WorkerReportsProgress = True}
    Private SyncWorkers As Dictionary(Of String, BackgroundWorker)
    Private SyncSet As Dictionary(Of String, DataTable)
    Private WithEvents Stop_Watch As New Stopwatch
    Private ReadOnly Intervals As New Dictionary(Of String, TimeSpan)
    Private ReadOnly Aliases As New Dictionary(Of String, String)
    Private ReadOnly ObjectsDictionary As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of SystemObject.ObjectType, List(Of SystemObject))))
    Private ReadOnly ConnectionsDictionary As New Dictionary(Of String, Boolean)
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
#Region " EXPORT DATA "
    Private WithEvents CMS_GridOptions As ContextMenuStrip
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
    Private ReadOnly Grid_ExcelQueryExport As New ToolStripMenuItem With {.Text = "+ Query",
        .Image = My.Resources.ExcelQuery,
        .Font = GothicFont}
    Private WithEvents Grid_DatabaseExport As New ToolStripMenuItem With {.Text = "Database",
        .Image = My.Resources.Database.ToBitmap,
        .Font = GothicFont}
#End Region
    Private Pane_MouseLocation As Point
    Private Pane_MouseObject As InstructionElement
#End Region
    Private Enum Sizing
        None
        MouseOverOPS
        MouseOverPGS
        MouseDownOPS
        MouseDownPGS
    End Enum

#Region " ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ N E W ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ "
    Public Sub New(Optional TestMode As Boolean = False)

        'Sync populates a Treeview with Checkmarks...those selected are imported. Submit how?
        'Interface to add, change, or remove a connection
        'Export / ETL
        'Casting using Select Min(Length(Trim(Field))), Max(Length(Trim(Field)))...Where Length(Trim(Field))>0

        Dock = DockStyle.Fill
        Me.TestMode = TestMode
        Scripts_ = New ScriptCollection(Me)

#Region " CONNECTIONS "
        Connections.SortCollection()
        'Connections.View()
        For Each Connection In Connections
            RaiseEvent Alert(Me, New AlertEventArgs("Initializing " & Connection.DataSource))
#Region " TOP LEVEL "
            AddHandler Connection.PasswordChanged, AddressOf ConnectionChanged
            Dim ColorKeys = ColorImages()
            Dim ColorImage As Image = ColorKeys(Connection.BackColor.Name)
            Dim ConnectionItem = TSMI_Connections.DropDownItems.Add(New ToolStripMenuItem With {
                                                                        .Text = Connection.DataSource,
                                                                        .Name = Connection.ToString,
                                                                        .Image = ColorImage,
                                                                        .Tag = Connection,
                                                                        .Font = GothicFont})
            AddHandler TSMI_Connections.DropDownItems(ConnectionItem).Click, AddressOf DataSource_Clicked
            AddHandler DirectCast(TSMI_Connections.DropDownItems(ConnectionItem), ToolStripMenuItem).DropDownOpening, AddressOf ConnectionProperties_Showing
            AddHandler DirectCast(TSMI_Connections.DropDownItems(ConnectionItem), ToolStripMenuItem).DropDownClosed, AddressOf ConnectionProperties_Closed
            Dim tsmiExport As ToolStripMenuItem = DirectCast(Grid_DatabaseExport.DropDownItems.Add(Connection.DataSource, ColorImage), ToolStripMenuItem)
            tsmiExport.Tag = Connection
            tsmiExport.Font = GothicFont

            Dim tlpExport As New TableLayoutPanel With {.Width = 305,
                .RowCount = 2,
                .ColumnCount = 1,
                .Margin = New Padding(0),
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                .BorderStyle = BorderStyle.Fixed3D,
                .Tag = Connection,
                .Font = GothicFont}

            Dim imagecomboTableName As New ImageCombo With {.Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .HintText = "Tablename",
                .Tag = Connection,
                .Font = GothicFont,
                .Name = "tableName"}
            Dim checkboxClearTable As New CheckBox With {.CheckState = CheckState.Checked,
                .Dock = DockStyle.Fill,
                .Margin = New Padding(5),
                .TextAlign = ContentAlignment.MiddleLeft,
                .CheckAlign = ContentAlignment.MiddleLeft,
                .TextImageRelation = TextImageRelation.ImageBeforeText,
                .Text = "Clear table".ToString(InvariantCulture),
                .Font = GothicFont,
                .Name = "clearTable"}
            Dim imagecomboTablespaceName As New ImageCombo With {.Dock = DockStyle.Fill,
                .Margin = New Padding(0),
                .HintText = "Table Space name",
                .Tag = Connection,
                .Font = GothicFont,
                .Name = "tableSpace"}

            AddHandler imagecomboTableName.MouseEnter, AddressOf ExportConnection_Enter
            AddHandler imagecomboTableName.ValueSubmitted, AddressOf ExportConnection_Submitted

            With tlpExport
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 300})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 28})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 28})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 0})
                .Controls.Add(imagecomboTableName, 0, 0)
                .Controls.Add(checkboxClearTable, 0, 1)
                .Controls.Add(imagecomboTablespaceName, 0, 2)
            End With
            tsmiExport.DropDownItems.Add(New ToolStripControlHost(tlpExport))
            AddHandler tsmiExport.DropDownOpening, AddressOf ExportConnection_Opening
#End Region
            Dim tlpConnection As New TableLayoutPanel With {.ColumnCount = 1,
                .RowCount = 2,
                .Margin = New Padding(0),
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                .BorderStyle = BorderStyle.None,
                .Tag = Connection,
                .Font = GothicFont}
            Dim buttonSubmit As New Button With {.Margin = New Padding(0),
                .Text = "S U B M I T".ToUpperInvariant,
                .Dock = DockStyle.Fill,
                .Height = 30,
                .Font = GothicFont}
            With tlpConnection
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 300})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 30})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 16})
                .Controls.Add(buttonSubmit, 0, 0)
            End With
            AddHandler buttonSubmit.Click, AddressOf ConnectionProperty_Submitted

            Dim tlpProperties As New TableLayoutPanel With {.ColumnCount = 3,
                .RowCount = 1 + Connection.PropertyIndices.Count,
                .BorderStyle = BorderStyle.None,
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                .Tag = Connection,
                .Font = GothicFont}
            With tlpProperties
                .Tag = Connection
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 1})
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 1})
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 1})
#Region " Add New Property row "
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 1})
                Dim addControl As New Button With {.Dock = DockStyle.Fill,
                    .Image = My.Resources.Plus,
                    .Margin = New Padding(0),
                    .ImageAlign = ContentAlignment.MiddleCenter,
                    .FlatStyle = FlatStyle.Standard,
                    .BackColor = Color.GhostWhite,
                    .Font = GothicFont}
                Dim addkeyControl As New ImageCombo With {.Dock = DockStyle.Fill,
                    .Text = String.Empty,
                    .Margin = New Padding(0),
                    .HintText = "Name",
                    .Font = GothicFont}
                addkeyControl.DropDown.CheckBoxes = False
                Dim addvalueControl As New ImageCombo With {.Dock = DockStyle.Fill,
                    .Text = String.Empty,
                    .Margin = New Padding(0),
                    .HintText = "Value",
                    .Enabled = False,
                    .Font = GothicFont}
                .Controls.Add(addControl, 0, 0)
                .Controls.Add(addkeyControl, 1, 0)
                .Controls.Add(addvalueControl, 2, 0)
                AddHandler addkeyControl.TextChanged, AddressOf ConnectionProperty_Change
                AddHandler addkeyControl.ValueSubmitted, AddressOf ConnectionProperty_Submitted
                AddHandler addkeyControl.ValueChanged, AddressOf ConnectionNewKeyProperty_Selected
                AddHandler addvalueControl.TextChanged, AddressOf ConnectionProperty_Change
                AddHandler addvalueControl.ValueSubmitted, AddressOf ConnectionProperty_Submitted
#End Region
                Dim rowIndex As Integer = 1
                For Each connectionProperty In Connection.PropertyIndices
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 1})
                    Dim deleteControl As New Button With {.Dock = DockStyle.Fill,
                        .Margin = New Padding(0),
                        .FlatStyle = FlatStyle.Standard,
                        .ImageAlign = ContentAlignment.MiddleCenter,
                        .BackColor = Color.GhostWhite,
                        .Font = GothicFont}
                    Dim keyControl As New ImageCombo With {.Dock = DockStyle.Fill,
                        .Text = connectionProperty.Key,
                        .Margin = New Padding(0),
                        .Enabled = False,
                        .Name = connectionProperty.Key,
                        .Font = GothicFont}
                    Dim valueControl As New ImageCombo With {.Dock = DockStyle.Fill,
                        .Text = String.Empty,
                        .Margin = New Padding(0),
                        .Font = GothicFont}
                    .Controls.Add(deleteControl, 0, rowIndex)
                    .Controls.Add(keyControl, 1, rowIndex)
                    .Controls.Add(valueControl, 2, rowIndex)
                    AddHandler deleteControl.Click, AddressOf ConnectionProperty_Change
                    AddHandler valueControl.TextChanged, AddressOf ConnectionProperty_Change
                    AddHandler valueControl.ValueSubmitted, AddressOf ConnectionProperty_Submitted
                    rowIndex += 1
                Next
            End With
            tlpConnection.Controls.Add(tlpProperties, 0, 1)
            ResizeConnections(tlpConnection, tlpProperties)
            DirectCast(TSMI_Connections.DropDownItems(ConnectionItem), ToolStripMenuItem).DropDownItems.Add(New ToolStripControlHost(tlpConnection) With {.BackColor = Connection.BackColor, .Font = GothicFont})
        Next
        TSMI_Copy.DropDownItems.AddRange({TSMI_CopyPlainText, TSMI_CopyColorText})
        TSMI_Divider.DropDownItems.AddRange({TSMI_DividerSingle, TSMI_DividerDouble})
#End Region

        Jobs.SortCollection()
        'Jobs.View()

        SystemObjects.SortCollection()
        'SystemObjects.View()

        Scripts.SortCollection()
        'Scripts.View()
        'PaneHandlers(HandlerAction.Add)

#Region " INITIALIZE CONTROLS "
        With TLP_Objects
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 100})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 30})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 50})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 50})
            Dim TLP_ObjectsHeader As New TableLayoutPanel With {.ColumnCount = 3,
                .RowCount = 1,
                .Margin = New Padding(0),
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
                .Dock = DockStyle.Fill}
            With TLP_ObjectsHeader
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 32})
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 100})
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 32})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 100})
                .Controls.Add(Button_ObjectsSync, 0, 0)
                .Controls.Add(IC_ObjectsSearch, 1, 0)
                .Controls.Add(Button_ObjectsClose, 2, 0)
            End With
            .Controls.Add(TLP_ObjectsHeader, 0, 0)
            .Controls.Add(Tree_Objects, 0, 1)
            Tree_Objects.AllowDrop = True
        End With
        With TLP_PaneGrid
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 0})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 50})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 50})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 100})
            .Controls.Add(TLP_Objects, 0, 0)
            .Controls.Add(Script_Tabs, 1, 0)
            .Controls.Add(Script_Grid, 2, 0)
            Controls.Add(TLP_PaneGrid)
        End With
        With Script_Grid
            With .Columns.HeaderStyle
                .ShadeColor = Color.Purple
                .BackColor = Color.Black
                .ForeColor = Color.White
            End With
            With .Rows.AlternatingRowStyle
                .BackColor = Color.GhostWhite
                .ForeColor = Color.Black
            End With
            With .Rows.RowStyle
            End With
        End With
        With TLP_ClosedScripts
            .Controls.Add(Tree_ClosedScripts, 0, 0)
            Dim TSCH_ClosedScripts As New ToolStripControlHost(TLP_ClosedScripts)
            TSDD_ClosedScripts.Items.Add(TSCH_ClosedScripts)
        End With
        For Each TextString In {"--", "=="}
            Dim _Image As New Bitmap(16, 16)
            Using G As Graphics = Graphics.FromImage(_Image)
                G.DrawString(TextString, GothicFont, Brushes.Black, 0, 0)
            End Using
            If TextString = "--" Then
                TSMI_DividerSingle.Image = _Image
            Else
                TSMI_DividerDouble.Image = _Image
            End If
        Next
#End Region
        SystemObjects.SortCollection()
        LoadSystemObjects(Nothing, Nothing)
#Region " EXPORT DATA "
        Script_Grid.AllowDrop = True
        Grid_FileExport.ImageScaling = ToolStripItemImageScaling.None
        Grid_ExcelExport.ImageScaling = ToolStripItemImageScaling.None
        Grid_ExcelQueryExport.ImageScaling = ToolStripItemImageScaling.None
        Grid_csvExport.ImageScaling = ToolStripItemImageScaling.None
        Grid_txtExport.ImageScaling = ToolStripItemImageScaling.None
        Grid_DatabaseExport.ImageScaling = ToolStripItemImageScaling.None
#End Region
        ExpandCollapseOnOff(HandlerAction.Add)

        For Each tsmiExport As ToolStripMenuItem In {Grid_ExcelExport, Grid_csvExport, Grid_txtExport}
            Grid_FileExport.DropDownItems.Add(tsmiExport)
            AddHandler tsmiExport.Click, AddressOf ExportToFile
        Next
        Grid_ExcelExport.DropDownItems.Add(Grid_ExcelQueryExport)
        AddHandler Grid_ExcelQueryExport.Click, AddressOf ExportToFile

    End Sub

    Protected Overrides Sub InitLayout()
        UpdateParentIcon_Text()
        MyBase.InitLayout()
    End Sub
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property
    Protected Overrides Sub OnParentChanged(e As EventArgs)
        UpdateParentIcon_Text()
        MyBase.OnParentChanged(e)
    End Sub
    Private Sub UpdateParentIcon_Text()

        Dim ParentForm As Form = TryCast(Parent, Form)
        If ParentForm IsNot Nothing Then
            ParentForm.Icon = My.Resources.Database
            ParentForm.Text = "Data Tool".ToString(InvariantCulture)
        End If

    End Sub
#End Region

    Public Event Alert(sender As Object, e As AlertEventArgs)

#Region " PROPERTIES - FUNCTIONS - METHODS "
    Public Shared ReadOnly Connections As New ConnectionCollection
    Public Shared ReadOnly SystemObjects As New SystemObjectCollection
    Public Shared ReadOnly Jobs As New JobCollection
    Private WithEvents Scripts_ As ScriptCollection
    Public ReadOnly Property Scripts As ScriptCollection
        Get
            Return Scripts_
        End Get
    End Property
    Private WithEvents ActiveTab_ As Tab
    Private ReadOnly Property ActiveTab As Tab
        Get
            ActiveTab_ = Script_Tabs.TabPages.Item({Script_Tabs.SelectedIndex, 0}.Max)
            ActiveTab_.AllowDrop = True
            Return ActiveTab_
        End Get
    End Property
    Private WithEvents ActivePane_ As RicherTextBox
    Private ReadOnly Property ActivePane As RicherTextBox
        Get
            If ActiveTab IsNot Nothing Then
                ActivePane_ = DirectCast(ActiveTab.Controls("Pane"), RicherTextBox)
                FindAndReplace.Parent = ActivePane_
            End If
            Return ActivePane_
        End Get
    End Property
    Private WithEvents ActiveScript_ As Script
    Private ReadOnly Property ActiveScript As Script
        Get
            ActiveTab_ = Script_Tabs.SelectedTab
            ActiveScript_ = DirectCast(ActiveTab.Tag, Script)
            Return ActiveScript_
        End Get
    End Property
    Private ReadOnly Property ActiveBody As BodyElements
        Get
            Return ActiveScript.Body
        End Get
    End Property
    Public ReadOnly Property Viewer As DataViewer
        Get
            Return Script_Grid
        End Get
    End Property
    Public Property TestMode As Boolean = False
    Public ReadOnly Property Pane As RicherTextBox
        Get
            Return ActivePane
        End Get
    End Property
    Public ReadOnly Property Grid As DataViewer
        Get
            Return Script_Grid
        End Get
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ ALERTS
    Private Sub ScriptAlerts(sender As Object, e As AlertEventArgs) Handles Scripts_.Alert
        RaiseEvent Alert(sender, e)
    End Sub
    Private Sub ViewerAlerts(sender As Object, e As AlertEventArgs) Handles Script_Grid.Alert
        RaiseEvent Alert(sender, e)
    End Sub
    Private Sub TreeViewerAlerts(sender As Object, e As AlertEventArgs) Handles Tree_ClosedScripts.Alert, Tree_Objects.Alert
        RaiseEvent Alert(sender, e)
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
#End Region

#Region " CONNECTION MANAGEMENT "
    Private Sub DataSource_Clicked(sender As Object, e As EventArgs)
        With DirectCast(sender, ToolStripMenuItem)
            ActiveScript.Connection = Connections.Item("DSN=" & .Text & ";")
        End With
    End Sub
    Private Sub ConnectionChanged(sender As Object, e As ConnectionChangedEventArgs) Handles ActiveScript_.ConnectionChanged

        If e.NewConnection Is Nothing Then
            'Password change
            Stop
        Else
            With e.NewConnection
                Dim Message As String = "Currently connected to " & .DataSource
                If .Properties.ContainsKey("NICKNAME") Then Message &= Join({String.Empty, "(", .Properties("NICKNAME"), ")"})
                RaiseEvent Alert(e.NewConnection, New AlertEventArgs(Message))
            End With
        End If

    End Sub
    Private Sub ConnectionProperties_Showing(sender As Object, e As EventArgs)

        '///////////////   R E S E T S   K E Y S  +  V A L U E S   ///////////////
        Dim tsmi_Connection As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim openingConnection As Connection = DirectCast(tsmi_Connection.Tag, Connection)
        Dim tlpConnection As TableLayoutPanel = DirectCast(DirectCast(tsmi_Connection.DropDownItems(0), ToolStripControlHost).Control, TableLayoutPanel)
        Dim tlpProperties As TableLayoutPanel = DirectCast(tlpConnection.GetControlFromPosition(0, 1), TableLayoutPanel)
        Dim tlpRows As Dictionary(Of Integer, List(Of Control)) = TLP.GetRows(tlpProperties)
        Dim newKey As ImageCombo = DirectCast(tlpRows(0)(1), ImageCombo)
        With newKey
            .Text = Nothing
            newKey.DataSource = openingConnection.PropertiesEmpty
            .HintText = "Property name"
        End With
        Dim newValue As ImageCombo = DirectCast(tlpRows(0)(2), ImageCombo)
        With newValue
            RemoveHandler .TextChanged, AddressOf ConnectionProperty_Change
            .Text = Nothing
            AddHandler .TextChanged, AddressOf ConnectionProperty_Change
            .HintText = "Property value"
        End With

        Dim buttonSubmit As Button = DirectCast(tlpConnection.GetControlFromPosition(0, 0), Button)
        With buttonSubmit
            .BackgroundImage = Nothing
            .FlatStyle = FlatStyle.System
        End With

        Dim rowIndex As Integer = 1
        For Each connectionProperty As KeyValuePair(Of String, Integer) In openingConnection.PropertyIndices
            Dim propertyIsUsed As Boolean = openingConnection.Properties.Keys.Contains(connectionProperty.Key)
            Dim backColor As Color = If(propertyIsUsed, Color.White, Color.Gainsboro)
            Dim foreColor As Color = If(propertyIsUsed, Color.Black, Color.DarkGray)
            Dim deleteControl As Button = DirectCast(tlpRows(rowIndex)(0), Button)
            Dim keyControl As ImageCombo = DirectCast(tlpRows(rowIndex)(1), ImageCombo)
            With keyControl
                .Text = connectionProperty.Key
                .BackColor = backColor
                .ForeColor = foreColor
            End With
            Dim valueControl As ImageCombo = DirectCast(tlpRows(rowIndex)(2), ImageCombo)
            With valueControl
                RemoveHandler .TextChanged, AddressOf ConnectionProperty_Change
                .Enabled = propertyIsUsed
                .Text = If(propertyIsUsed, openingConnection.Properties(connectionProperty.Key), String.Empty)
                .BackColor = backColor
                .ForeColor = foreColor
                AddHandler .TextChanged, AddressOf ConnectionProperty_Change
                deleteControl.Image = If(.Enabled, My.Resources.Close.ToBitmap, My.Resources.Plus)
            End With
            rowIndex += 1
        Next
        ResizeConnections(tlpConnection, tlpProperties)

    End Sub
    Private Sub ConnectionNewKeyProperty_Selected(sender As Object, e As ImageComboEventArgs)

        With DirectCast(sender, ImageCombo)
            Dim tlpProperties As TableLayoutPanel = DirectCast(.Parent, TableLayoutPanel)
            Dim newRow = TLP.GetRows(tlpProperties)(0)
            newRow(2).Enabled = newRow(1).Text.Any
            newRow(2).BackColor = If(newRow(2).Enabled, Color.White, Color.Gainsboro)
        End With

    End Sub
    Private Sub ConnectionProperty_Change(sender As Object, e As EventArgs)

        Dim senderControl As Control = DirectCast(sender, Control)
        Dim tlpProperties As TableLayoutPanel = DirectCast(senderControl.Parent, TableLayoutPanel)
        Dim tlpConnection As TableLayoutPanel = DirectCast(tlpProperties.Parent, TableLayoutPanel)
        Dim submitButton As Button = DirectCast(tlpConnection.Controls(0), Button) ' S U B M I T   B U T T O N
        Dim existingConnection As Connection = DirectCast(tlpProperties.Tag, Connection)
        Dim tlpRows = TLP.GetRows(tlpProperties)
        Dim valueWidth As Integer = 0
        Dim connectionProperties As New Dictionary(Of Integer, String)

        For Each row In tlpRows
            Dim rowButton As Button = DirectCast(row.Value(0), Button)
            Dim rowKeyCombo As ImageCombo = DirectCast(row.Value(1), ImageCombo)
            Dim rowValueCombo As ImageCombo = DirectCast(row.Value(2), ImageCombo)
            Dim rowButtonClicked = rowButton Is senderControl
            With rowValueCombo
                RemoveHandler .TextChanged, AddressOf ConnectionProperty_Change
                If rowButtonClicked Then
                    .Enabled = Not .Enabled
                    If .Enabled Then .Text = existingConnection.Properties(rowKeyCombo.Text)
                Else
                    .Enabled = rowKeyCombo.Text.Any And .Text.Any
                End If
                If .Enabled Then
                    'Include in consideration ... but reset to default
                    rowButton.Image = My.Resources.Close.ToBitmap
                    .BackColor = Color.White
                    .ForeColor = Color.Black

                    If row.Key = 0 Then
                        rowKeyCombo.DataSource = existingConnection.PropertiesEmpty
                    Else
                        rowKeyCombo.BackColor = Color.White
                    End If

                    Dim rowValueWidth As Integer = MeasureText(.Text, .Font).Width
                    If valueWidth < rowValueWidth Then valueWidth = rowValueWidth
                    connectionProperties.Add(existingConnection.PropertyIndices(rowKeyCombo.Text), Join({rowKeyCombo.Text, .Text}, "="))

                ElseIf row.Key > 0 Then
                    'Remove from consideration
                    rowButton.Image = My.Resources.Plus
                    rowKeyCombo.BackColor = Color.Gainsboro
                    .BackColor = Color.Gainsboro
                    .ForeColor = Color.DarkGray
                End If
                AddHandler .TextChanged, AddressOf ConnectionProperty_Change
            End With
        Next

        If senderControl.GetType Is GetType(ImageCombo) Then
            'ImageCombo text changing ... Modify width
            ResizeConnections(tlpConnection, tlpProperties)
        End If

        Dim orderedProperties As New List(Of String)(connectionProperties.OrderBy(Function(k) k.Key).Select(Function(v) v.Value))
        Dim newConnectionString As String = Join(orderedProperties.ToArray, ";")
        With submitButton
            TT_Submit.Hide(submitButton)
            If newConnectionString = existingConnection.ToString Then
                .BackgroundImage = Nothing
                .FlatStyle = FlatStyle.System

            Else
                .BackgroundImage = My.Resources.Button_Bright
                .BackgroundImageLayout = ImageLayout.Stretch
                .FlatStyle = FlatStyle.Flat
                TT_Submit.Show(newConnectionString, submitButton, New Point(-3, -(5 + submitButton.Height + 5)))
            End If
        End With

    End Sub
    Private Sub ConnectionProperty_Submitted(sender As Object, e As ImageComboEventArgs)

        With DirectCast(sender, ImageCombo)
            Dim tlpProperties As TableLayoutPanel = DirectCast(.Parent, TableLayoutPanel)
            Dim tlpConnection As TableLayoutPanel = DirectCast(tlpProperties.Parent, TableLayoutPanel)
            Dim submitButton As Button = DirectCast(tlpConnection.Controls(0), Button) ' S U B M I T   B U T T O N
            ConnectionProperty_Submitted(submitButton, New EventArgs)
        End With

    End Sub
    Private Sub ConnectionProperty_Submitted(sender As Object, e As EventArgs)

        Dim buttonSubmit As Button = DirectCast(sender, Button)
        Dim tlpConnection As TableLayoutPanel = DirectCast(buttonSubmit.Parent, TableLayoutPanel)
        Dim tlpProperties As TableLayoutPanel = DirectCast(tlpConnection.GetControlFromPosition(0, 1), TableLayoutPanel)
        Dim connectionSubmitted As Connection = DirectCast(tlpProperties.Tag, Connection)
        Dim tsmiConnection As ToolStripMenuItem = DirectCast(TSMI_Connections.DropDownItems(connectionSubmitted.ToString), ToolStripMenuItem)
        Dim tlpRows = TLP.GetRows(tlpProperties).OrderByDescending(Function(r) r.Key)       'Make Descending since New property is position 0 and must override actual hidden property

        For Each row In tlpRows
            Dim keyControl As Control = row.Value(1)
            Dim valueControl As Control = row.Value(2)
            Dim useProperty As Boolean = valueControl.Enabled And keyControl.Text.Any And valueControl.Text.Any
            connectionSubmitted.SetProperty(keyControl.Text, If(useProperty, valueControl.Text, String.Empty))
        Next

        CMS_PaneOptions.AutoClose = True
        CMS_PaneOptions.Hide()
        TT_Submit.Hide(buttonSubmit)
        tsmiConnection.Name = connectionSubmitted.ToString
        connectionSubmitted.Parent.Save()

    End Sub
    Private Sub ConnectionProperties_Closed(sender As Object, e As EventArgs)

        Dim tsmi_Connection As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim tlpConnection As TableLayoutPanel = DirectCast(DirectCast(tsmi_Connection.DropDownItems(0), ToolStripControlHost).Control, TableLayoutPanel)
        Dim buttonSubmit As Button = DirectCast(tlpConnection.GetControlFromPosition(0, 0), Button)
        TT_Submit.Hide(buttonSubmit)

    End Sub

    Private Sub ResizeConnections(tlpConnection As TableLayoutPanel, tlpProperties As TableLayoutPanel)

        Dim connectionResize = DirectCast(tlpConnection.Tag, Connection)
        Dim propertyRows = TLP.GetRows(tlpProperties)
        Dim keyWidth As Integer = MeasureText("Property name".ToUpperInvariant, Font).Width
        Dim valueWidth As Integer = MeasureText("Property value".ToUpperInvariant, Font).Width
        Dim imageWH As Integer = {24, My.Resources.Close.Width, My.Resources.Plus.Width}.Max

        For Each row In propertyRows
            Dim buttonControl As Button = DirectCast(row.Value(0), Button)
            Dim keyControl As Control = row.Value(1)
            Dim valueControl As Control = row.Value(2)
            keyWidth = {keyWidth, MeasureText(keyControl.Text.ToUpperInvariant, keyControl.Font).Width}.Max
            valueWidth = {valueWidth, MeasureText(valueControl.Text.ToUpperInvariant, valueControl.Font).Width}.Max
            Dim isVisible As Boolean = If(row.Key = 0, True, connectionResize.Properties.ContainsKey(keyControl.Text))
            buttonControl.Image = If(isVisible, If(row.Key = 0, My.Resources.Plus, If(valueControl.Enabled, My.Resources.Close.ToBitmap, My.Resources.Plus)), Nothing)
            tlpProperties.RowStyles(row.Key).Height = If(isVisible, imageWH, 0)
        Next

        With tlpProperties
            .ColumnStyles(0).Width = imageWH
            .ColumnStyles(1).Width = keyWidth
            .ColumnStyles(2).Width = valueWidth + 16 ' [ X ] Clear Text Image width
        End With
        Dim sizeProperties As Size = TLP.GetSize(tlpProperties)
        With tlpConnection
            .ColumnStyles(0).Width = sizeProperties.Width
            .RowStyles(0).Height = 30
            .RowStyles(1).Height = sizeProperties.Height
        End With
        TLP.SetSize(tlpConnection)
        tlpProperties.Size = sizeProperties

    End Sub
#End Region

#Region " TABLELAYOUTPANEL SIZING - PANE→|←GRID "
    Private ObjectsWidth As Integer = 200
    Private Sub ObjectsClose() Handles Button_ObjectsClose.Click
        TLP_PaneGrid.ColumnStyles(0).Width = 0
    End Sub
    Private SeparatorSizing As New Sizing
    Private ReadOnly Property ObjectsPaneSeparator As Rectangle
        Get
            Dim OPS As New Rectangle
            If ActivePane_ IsNot Nothing Then
                Dim Pane_Location = ActivePane_.PointToScreen(New Point(0, 0))
                OPS = New Rectangle(Pane_Location.X - 10, Pane_Location.Y, 10, ActivePane_.Height)
            End If
            Return OPS
        End Get
    End Property
    Private ReadOnly Property PaneGridSeparator As Rectangle
        Get
            Dim PGS As New Rectangle
            If ActivePane_ IsNot Nothing Then
                Dim Grid_Location = Script_Grid.PointToScreen(New Point(0, 0))
                Dim Pane_Location = ActivePane_.PointToScreen(New Point(0, 0))
                Dim Pane_Right = Pane_Location.X + ActivePane_.Width
                Dim PGS_Width As Integer = Grid_Location.X - Pane_Right
                PGS = New Rectangle(Pane_Right, Pane_Location.Y, PGS_Width, ActivePane_.Height)
            End If
            Return PGS
        End Get
    End Property
    Private _ForceCapture As Boolean
    Protected Property ForceCapture() As Boolean
        Get
            Return _ForceCapture
        End Get
        Set(value As Boolean)
            _ForceCapture = value
            TLP_PaneGrid.Capture = value
        End Set
    End Property
    Private Sub OnPanelMouseOver(sender As Object, e As MouseEventArgs) Handles TLP_PaneGrid.MouseMove

        If e.Button = MouseButtons.Left Or ForceCapture Then
            If SeparatorSizing = Sizing.MouseDownOPS Then
                TLP_PaneGrid.ColumnStyles(0).SizeType = SizeType.Absolute
                TLP_PaneGrid.ColumnStyles(0).Width = e.X
                ObjectsWidth = e.X

            ElseIf SeparatorSizing = Sizing.MouseDownPGS Then
                TLP_PaneGrid.ColumnStyles(1).SizeType = SizeType.Absolute
                TLP_PaneGrid.ColumnStyles(1).Width = e.X - TLP_PaneGrid.ColumnStyles(0).Width

            End If

        Else
            If ObjectsPaneSeparator.Contains(Cursor.Position) Then
                SeparatorSizing = Sizing.MouseOverOPS

            ElseIf PaneGridSeparator.Contains(Cursor.Position) Then
                SeparatorSizing = Sizing.MouseOverPGS

            Else
                SeparatorSizing = Sizing.None

            End If

        End If
        If SeparatorSizing = Sizing.None Then
            Cursor = Cursors.Default
        Else
            Cursor = Cursors.VSplit
        End If

        If e.X < 10 And TLP_Objects.Width < 2 Then
            If My.Settings.DontShowObjectsMessage Then
                Cursor = Cursors.VSplit
            Else
#Region " CUSTOM CURSOR WITH TEXT "
                Dim BoxSize As New Size(200, 200)
                Dim CursorBounds As New Rectangle(0, 0, Cursor.Size.Width, Cursor.Size.Width)
                Dim CursorText As String = "Double-Click to view Objects. Don't show again (Right-Click)".ToString(InvariantCulture)
                Dim TextSize As Size = TextRenderer.MeasureText(CursorText, GothicFont, BoxSize)
                Dim TextBounds As New Rectangle(CursorBounds.Right, CursorBounds.Top, TextSize.Width, CursorBounds.Height)
                Dim CursorTextBounds As New Rectangle(CursorBounds.X, CursorBounds.Y, CursorBounds.Width + TextSize.Width, {CursorBounds.Height, TextSize.Height}.Max)
                Dim BorderBounds = CursorTextBounds
                Dim bmp As New Bitmap(CursorTextBounds.Width, CursorTextBounds.Height)
                Using Graphics As Graphics = Graphics.FromImage(bmp)
                    With Graphics
                        .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        .FillRectangle(Brushes.White, CursorTextBounds)
                        Cursor.Draw(Graphics, CursorBounds)
                        .DrawString(CursorText, GothicFont, Brushes.Black, TextBounds, StringFormat.GenericDefault)
                        BorderBounds.Inflate(-1, -1)
                        .DrawRectangle(Pens.CornflowerBlue, BorderBounds)
                        bmp.MakeTransparent(Color.White)
                    End With
                End Using
                Cursor = CursorHelper.CreateCursor(bmp, 0, Convert.ToInt32(bmp.Height / 2))
#End Region
            End If
        Else
        End If

    End Sub
    Private Shadows Sub OnMouseCaptureChanged(sender As Object, e As EventArgs) Handles TLP_PaneGrid.MouseCaptureChanged
        TLP_PaneGrid.Capture = _ForceCapture
    End Sub
    Private Sub OnPanelDown(sender As Object, e As MouseEventArgs) Handles TLP_PaneGrid.MouseDown

        If e.Button = MouseButtons.Left Then
            If ObjectsPaneSeparator.Contains(Cursor.Position) Then
                SeparatorSizing = Sizing.MouseDownOPS

            ElseIf PaneGridSeparator.Contains(Cursor.Position) Then
                SeparatorSizing = Sizing.MouseDownPGS

            Else
                SeparatorSizing = Sizing.None

            End If
            ForceCapture = Not SeparatorSizing = Sizing.None

        ElseIf e.Button = MouseButtons.Right Then
            My.Settings.DontShowObjectsMessage = True
            My.Settings.Save()
        End If

    End Sub
    Private Sub OnPanelDoubleClick(sender As Object, e As EventArgs) Handles TLP_PaneGrid.DoubleClick

        If ObjectsPaneSeparator.Contains(Cursor.Position) Then
            TLP_PaneGrid.ColumnStyles(0).Width = ObjectsWidth

        ElseIf PaneGridSeparator.Contains(Cursor.Position) Then
            AutoWidth(ActivePane)

        End If

    End Sub
    Private Sub OnPanelUp(sender As Object, e As MouseEventArgs) Handles TLP_PaneGrid.MouseUp

        ForceCapture = False
        If ObjectsPaneSeparator.Contains(e.Location) Then
            SeparatorSizing = Sizing.MouseOverOPS

        ElseIf PaneGridSeparator.Contains(e.Location) Then
            SeparatorSizing = Sizing.MouseOverPGS

        End If

    End Sub
    Private Sub OnPanelLeave(sender As Object, e As EventArgs) Handles TLP_PaneGrid.Leave, Tree_Objects.Enter, Tree_Objects.MouseMove

        If ForceCapture Then
        Else
            Cursor = Cursors.Default
        End If

    End Sub
#End Region

#Region " Manage Toolstrip Visibility {TSDD_SaveAs ( IC_SaveAs ) + TSDD_ClosedScripts ( Tree_ClosedScripts )} "
    'Mouseover Tab close [X] shows SaveAs, if text has changed. Exit right stays visible. Exit left closes
    'Mouseover AddTab shows ClosedScripts. Exit right stays visible. Exit left closes
    Private Sub TabDirection(TabItem As Tab)

        If TabItem IsNot Nothing Then
            'Inside [X]=Show SaveAs ( if text has changed )
            Dim RelativePosition As RelativeCursor = CursorToControlPosition(Script_Tabs, TabItem.Bounds)
            Dim TabLocation As Point = Script_Tabs.PointToScreen(New Point(TabItem.Bounds.Right, 1))    'TabItem.Bounds.Top

            If RelativePosition = RelativeCursor.RightOf Then
                If TabItem Is Script_Tabs.AddTab Then
                    TSDD_ClosedScripts.Location = TabLocation  'Upper-right corner
                    Show_DropDown(TSDD_ClosedScripts, TabLocation)
                    Hide_DropDown(TSDD_SaveAs)
                    'ClosedScripts changes Size so the Cursor may end up outside the bounds when a Node is collapsed. Allow a 2s window before closing
                    With New Timer With {.Interval = 2000}
                        AddHandler .Tick, AddressOf HideTimer_ClosedScripts
                        .Start()
                    End With

                Else
                    Show_DropDown(TSDD_SaveAs, TabLocation)
                    Hide_DropDown(TSDD_ClosedScripts)

                End If

            ElseIf RelativePosition = RelativeCursor.Inside Then
                If TabItem Is Script_Tabs.AddTab Then
                    If TSDD_SaveAs.Visible Then
                        Return
                    End If
                    Show_DropDown(TSDD_ClosedScripts, TabLocation)

                Else
                    Show_DropDown(TSDD_SaveAs, TabLocation)
                    Hide_DropDown(TSDD_ClosedScripts)

                End If

            Else
                If TabItem Is Script_Tabs.AddTab Then
                    Hide_DropDown(TSDD_ClosedScripts)
                    If RelativePosition = RelativeCursor.LeftOf Then Show_DropDown(TSDD_SaveAs, TabLocation)

                Else
                    Hide_DropDown(TSDD_SaveAs)

                End If
            End If
        End If

    End Sub
    Private Sub Show_DropDown(TSDD As ToolStripDropDown, Optional Location As Point = Nothing)

        If TSDD Is TSDD_SaveAs And ActiveScript IsNot Nothing AndAlso ActiveScript.TextWasModified Or TSDD Is TSDD_ClosedScripts And ScriptsInitialized Then
            'Do not show SaveAs when ScriptText has not changed
            With TSDD
                .AutoClose = False
                .Show(Location)
            End With
        End If

    End Sub
    Private Sub Hide_DropDowns() Handles IC_SaveAs.MouseLeave
        Hide_DropDown(TSDD_SaveAs)
        Hide_DropDown(TSDD_ClosedScripts)
    End Sub
    Private Shared Sub Hide_DropDown(TSDD As ToolStripDropDown)
        With TSDD
            .AutoClose = True
            .Hide()
        End With
    End Sub
    Private Sub HideTimer_ClosedScripts(sender As Object, e As EventArgs)

        With DirectCast(sender, Timer)
            RemoveHandler .Tick, AddressOf HideTimer_ClosedScripts
            .Stop()
            If CursorOverControl(Tree_ClosedScripts) Or Tree_ClosedScripts.OptionsOpen Then
                AddHandler .Tick, AddressOf HideTimer_ClosedScripts
                .Start()
            Else
                If Not CursorToControlPosition(Script_Tabs, Script_Tabs.AddTab.Bounds) = RelativeCursor.Inside Then
                    Hide_DropDown(TSDD_ClosedScripts)
                End If
            End If
        End With

    End Sub
#End Region

#Region " IC_CloseAndSave EVENTS "
    Private Sub SaveAs_Showing() Handles TSDD_SaveAs.Opening
        IC_SaveAs.Text = ActiveScript.Name
    End Sub
    Private Sub SaveAs_ImageClicked() Handles IC_SaveAs.ValueSubmitted, IC_SaveAs.ImageClicked

        Using cb As New CursorBusy
            'USING Now.ToLongTimeString ENSURE NAME<>value AND ACTION IS TAKEN
            Dim ActiveScriptName As String = Join({DateTimeToString(Now), IC_SaveAs.Text}, Delimiter)
            ActiveScript.Name = ActiveScriptName
        End Using

    End Sub
#End Region

#Region " Tree_ClosedScripts EVENTS "
    Private Sub ClosedScripts_SizeChanged(sender As Object, e As EventArgs)

        With TLP_ClosedScripts
            .ColumnStyles(0).Width = {Tree_ClosedScripts.Width, WorkingArea.Width - TSDD_ClosedScripts.Left}.Min
            .RowStyles(0).Height = {Tree_ClosedScripts.Height, WorkingArea.Height - TSDD_ClosedScripts.Top}.Min
            .Width = Convert.ToInt32(.ColumnStyles(0).Width + 3)
            .Height = Convert.ToInt32(.RowStyles(0).Height + 3)
            TSDD_ClosedScripts.Size = .Size
        End With

    End Sub
    Private Sub ClosedScript_NodeDragStart(sender As Object, e As NodeEventArgs) Handles Tree_ClosedScripts.NodeDragStart

        DragNode = e.Node
        If ActivePane IsNot Nothing Then ActivePane.AllowDrop = True
        Script_Grid.AllowDrop = True

    End Sub
    Private Sub ClosedScript_NodeDragOver(sender As Object, e As DragEventArgs) Handles Script_Tabs.DragOver, Script_Grid.DragOver

        Dim DragNode As Node = DirectCast(e.Data.GetData(GetType(Node)), Node)
        If DragNode IsNot Nothing Then
            If DragNode.AllowDragDrop Then
                e.Effect = DragDropEffects.All
            Else
                e.Effect = DragDropEffects.None
            End If
        End If

    End Sub
    Private Sub ClosedScript_NodeDroppedTabs(sender As Object, e As DragEventArgs) Handles Script_Tabs.DragDrop
        Pane_NodeDropped(e)
    End Sub
    Private Sub ClosedScript_NodeDroppedPane(sender As Object, e As DragEventArgs) Handles ActivePane_.DragDrop
        Pane_NodeDropped(e)
    End Sub
    Private Sub ClosedScript_NodeClicked(sender As Object, e As NodeEventArgs) Handles Tree_ClosedScripts.NodeClicked
        If e.Node Is OpenFileNode Then
            OpenFile.Tag = Nothing
            OpenFile.ShowDialog()
        End If
    End Sub
    'Open Closed Script
    Private Sub ClosedScript_NodeDoubleClicked(sender As Object, e As NodeEventArgs) Handles Tree_ClosedScripts.NodeDoubleClicked

        Hide_DropDowns()
        If e.Node.Tag IsNot Nothing Then
            If e.Node.Tag.GetType Is GetType(Connection) Then
                'New Pane with Connection
                Scripts.Add(New Script With {
                            ._Tabs = Script_Tabs,
                            .Connection = DirectCast(e.Node.Tag, Connection),
                            .State = Script.ViewState.OpenDraft})

            ElseIf e.Node.Tag.GetType Is GetType(Script) Then
                'Opening a Closed Script
                Dim NodeScript As Script = DirectCast(e.Node.Tag, Script)
                NodeScript.State = Script.ViewState.OpenSaved

            End If
        End If

    End Sub
    Private Sub ClosedScript_NodeRemoveClicked(sender As Object, e As NodeEventArgs) Handles Tree_ClosedScripts.NodeAfterRemoved

        Dim RemoveScript As Script = DirectCast(e.Node.Tag, Script)
        RemoveScript.State = Script.ViewState.None

    End Sub
    Private Sub ClosedScript_NodeEditClicked(sender As Object, e As NodeEventArgs) Handles Tree_ClosedScripts.NodeEdited

        Using cb As New CursorBusy
            'USING Now.ToLongTimeString ENSURE NAME<>value AND ACTION IS TAKEN
            Dim ClosedScript As Script = DirectCast(e.Node.Tag, Script)
            ClosedScript.Name = Join({DateTimeToString(Now), e.ProposedText}, Delimiter)
            If ClosedScript.Name = e.ProposedText Then e.Node.Text = e.ProposedText     'Script.Name will only change if it can
            e.Node.Parent.SortChildren()
            Tree_ClosedScripts.Refresh()
        End Using

    End Sub
    Private Sub Nodes_Changed(sender As Object, e As NodeEventArgs) Handles Tree_ClosedScripts.NodesChanged

    End Sub
#End Region
#Region " SCRIPT CONTROL EVENTS "
    Private Sub ActivePane_KeyDown(sender As Object, e As KeyEventArgs) Handles ActivePane_.KeyDown

        If Control.ModifierKeys = Keys.Control Then
            Select Case e.KeyCode
                Case Keys.Left
                    InsertDividers(TSMI_DividerSingle, Nothing)

                Case Keys.Right
                    InsertDividers(TSMI_DividerDouble, Nothing)

                Case Keys.C

                Case Keys.V

                Case Keys.F
                    With FindAndReplace
                        .DataSource = ActivePane.Text
                        If ActivePane.SelectedText.Length = 0 Then
                        Else
                            .FindControl.Text = ActivePane.SelectedText
                        End If
                    End With

            End Select
        End If

    End Sub
    Private Sub ActivePane_MouseDown(sender As Object, e As MouseEventArgs) Handles ActivePane_.MouseDown

        With CMS_PaneOptions
            If e.Button = MouseButtons.Left Then
                .AutoClose = True
                .Hide()

            ElseIf e.Button = MouseButtons.Right Then
                With .Items
                    .Clear()
#Region " LINE HAS A COMMENT? "
                    GetCommentMatch()
#End Region
                    If IsNothing(Pane_MouseObject.Highlight) Then
#Region " RIGHTCLICKED UNDEFINED REGION "
                        .AddRange({TSMI_Connections,
                                                   TSMI_Comment,
                                                   TSMI_Copy,
                                                   TSMI_Divider,
                                                   TSMI_Font})
#End Region
                    Else
#Region " RIGHTCLICKED ON OBJECT "
                        .AddRange({TSMI_ObjectType,
                                                    TSMI_ObjectValue,
                                                    TSMI_TipSwitch})
#End Region
                        RemoveHandler IC_BackColor.SelectionChanged, AddressOf ColorSelected
                        RemoveHandler IC_ForeColor.SelectionChanged, AddressOf ColorSelected
                        With Pane_MouseObject
                            If Not TLP_Type.Controls.Count = 0 Then
                                REM /// INITIALIZE THEM
                                With IC_BackColor
                                    .DropDown.CheckBoxes = False
                                    .ColorPicker = True
                                End With
                                With IC_ForeColor
                                    .DropDown.CheckBoxes = False
                                    .ColorPicker = True
                                End With
                                TLP_Type.Controls.Add(IC_BackColor)
                                TLP_Type.Controls.Add(IC_ForeColor)
                            End If
                            Dim MouseWords As New List(Of StringData)(From M In Regex.Matches(.Highlight.Value, "[^\s]{1,}", RegexOptions.IgnoreCase) Select New StringData(M))
                            Dim MouseWord As String = MouseWords.First.Value
                            If MouseWord.Length < .Highlight.Value.Length Then MouseWord += "..."
                            With TSMI_ObjectType
                                .BackColor = Pane_MouseObject.Highlight.BackColor
                                If .BackColor.IsKnownColor AndAlso IC_BackColor.Items.Any Then
                                    Dim BackColors = IC_BackColor.Items.Where(Function(x) x.Text = .BackColor.Name).Select(Function(y) y.Index)
                                    If BackColors.Any Then
                                        IC_BackColor.SelectedIndex = BackColors.Max
                                    End If
                                End If
                                .ForeColor = Pane_MouseObject.Highlight.ForeColor
                                If .ForeColor.IsKnownColor AndAlso IC_ForeColor.Items.Any Then
                                    Dim ForeColors = IC_ForeColor.Items.Where(Function(x) x.Text = .ForeColor.Name).Select(Function(y) y.Index)
                                    If ForeColors.Any Then
                                        IC_ForeColor.SelectedIndex = ForeColors.Max
                                    End If
                                End If
                                .Text = Pane_MouseObject.Source.ToString
                            End With
                            With TSMI_ObjectValue
                                .Text = MouseWord
                                .DropDownItems.Clear()
                                If Pane_MouseObject.Source = InstructionElement.LabelName.SystemTable Then
                                    .Image = My.Resources.Table
                                Else
                                    .Image = Nothing
                                End If
                            End With
                        End With
                        AddHandler IC_BackColor.SelectionChanged, AddressOf ColorSelected
                        AddHandler IC_ForeColor.SelectionChanged, AddressOf ColorSelected
                    End If
                End With
                .Show(Cursor.Position)
            End If
        End With

    End Sub
    Private Sub ActivePane_MouseEnter(sender As Object, e As EventArgs) Handles ActivePane_.MouseEnter
        CMS_PaneOptions.Close()
    End Sub
    Private Sub ActivePane_MouseMove(sender As Object, e As MouseEventArgs) Handles ActivePane_.MouseMove

        With ActivePane
            If Pane_MouseLocation <> e.Location Then
                Pane_MouseLocation = e.Location
                Dim CharacterIndex As Integer = .GetCharIndexFromPosition(e.Location)
                If .MouseWord.Intersects Then
                    Dim Labels As New List(Of InstructionElement)(ActiveBody.Labels)
                    Dim MO As New List(Of InstructionElement)(From l In Labels Where Enumerable.Range(l.Highlight.Start, l.Highlight.Length).Contains(CharacterIndex))
                    With MO
                        If .Any Then
                            Dim MoveObject = .First
                            If MoveObject.Highlight <> Pane_MouseObject.Highlight Then
                                With MoveObject
                                    Pane_MouseObject = MoveObject
                                    Dim TipText As String = Nothing
                                    Dim MouseWords As New List(Of StringData)(From M In Regex.Matches(.Highlight.Value, "[^\s]{1,}", RegexOptions.IgnoreCase) Select New StringData(M))
                                    Dim MouseWord As String = MouseWords.First.Value
                                    If MouseWord.Length < .Highlight.Value.Length Then MouseWord += "..."
                                    If ActiveBody.GroupedLabels.ContainsKey(Pane_MouseObject.Source) Then
                                        Dim ElementObjects = ActiveBody.ElementObjects
                                        If ElementObjects.ContainsKey(Pane_MouseObject) Then
                                            Dim Objects As List(Of SystemObject) = ElementObjects(Pane_MouseObject)
                                            Dim Items As New Dictionary(Of String, List(Of String))
                                            For Each DataSource As SystemObject In Objects
                                                If Not Items.ContainsKey(DataSource.DSN) Then Items.Add(DataSource.DSN, New List(Of String))
                                                With Items(DataSource.DSN)
                                                    .Add(DataSource.Type.ToString)
                                                    .Add(DataSource.DBName & " (Database Name)")
                                                    .Add(DataSource.TSName & " (TableSpace Name)")
                                                End With
                                            Next
                                            TipText = Join({MouseWord, Bulletize(Items)}, "|")
                                            Dim Location As Point = CenterItem(CMS_PaneOptions.Size)
                                            Location.Offset(ActivePane.PointToClient(New Point(0, 0)))
                                            Pane_TipManager(TipText, Location)
                                        End If

                                    Else
                                        Dim Items As New List(Of String) From {Pane_MouseObject.Source.ToString}
                                        TipText = Join({MouseWord, Bulletize(Items.ToArray)}, "|")
                                    End If
                                End With
                            End If
                        End If
                    End With
                Else
                    Pane_MouseObject = Nothing
                    TT_PaneTip.Hide(ActivePane)
                End If
            End If
        End With

    End Sub
    Private Sub ActivePane_TextChanged(sender As Object, e As EventArgs) Handles ActivePane_.TextChanged
        ActiveBody.Text = ActivePane.Text
    End Sub
    Private Sub ActivePane_SelectionChanged(sender As Object, e As EventArgs) Handles ActivePane_.SelectionChanged

        With ActivePane
            FindAndReplace.StartAt = .SelectionStart
            Dim Statement As String = .Text
            Dim Parentheses As New List(Of Match)(From M In Regex.Matches(Statement, "\(|\)", RegexOptions.IgnoreCase) Select DirectCast(M, Match))
            If .SelectedText = "(" Or .SelectedText = ")" Then
                Dim SelectedStartAndLength As New KeyValuePair(Of Integer, Integer)(.SelectionStart, .SelectionLength)
                If .SelectedText = "(" Then
                    REM /// FORWARDS
                    Parentheses = (From P In Parentheses Order By P.Index Ascending Where P.Index >= .SelectionStart).ToList
                Else
                    REM /// BACKWARDS
                    Parentheses = (From P In Parentheses Order By P.Index Descending Where P.Index <= .SelectionStart).ToList
                End If
                Dim LeftCount As Integer = 0, RightCount As Integer = 0
                For Each Parenthese In Parentheses
                    If Parenthese.Value = "(" Then LeftCount += 1
                    If Parenthese.Value = ")" Then RightCount += 1
                    If LeftCount = RightCount Then
                        .SelectionStart = {Parentheses.First.Index, Parenthese.Index}.Min
                        .SelectionLength = 1 + Math.Abs(Parenthese.Index - Parentheses.First.Index)
                        Exit For
                    End If
                Next
            Else
            End If
        End With

    End Sub
    Private Sub ActivePane_ScrolledV(sender As Object, e As RicherEventArgs) Handles ActivePane_.ScrolledVertical

        With New Timer With {.Interval = 250}
            AddHandler .Tick, AddressOf ScrollTimer_Tick
            .Start()
        End With

    End Sub
    Private Sub ScrollTimer_Tick(sender As Object, e As EventArgs)

        With DirectCast(sender, Timer)
            RemoveHandler .Tick, AddressOf ScrollTimer_Tick
            .Stop()
        End With
        FindAndReplace.Top = 0

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub FindRequest(sender As Object, e As ZoneEventArgs) Handles FindAndReplace.ZoneClicked

        Dim Text_Search As String = ActivePane.Text
        Select Case e.Zone.Name
            Case Zone.Identifier.MatchCase, Zone.Identifier.MatchWord, Zone.Identifier.RegEx
                FindRequest()

            Case Zone.Identifier.Close
                'Remove the Highlights
                With ActivePane
                    Dim _SelectionStart As Integer = .SelectionStart
                    .SelectAll()
                    .SelectionBackColor = Color.Transparent
                    .SelectionColor = Color.Black
                    .SelectionStart = _SelectionStart
                    .SelectionLength = 0
                End With

            Case Zone.Identifier.GotoNext
                If FindAndReplace.CurrentMatch.Key >= 0 Then
                    FindRequest()
                    Dim Match = FindAndReplace.CurrentMatch
                    Dim _rtf As String = ActivePane.Rtf
                    Using RTB As New RichTextBox With {.Rtf = _rtf}
                        With RTB
                            .SelectionStart = Match.Key
                            .SelectionLength = Match.Value.Length
                            .SelectionBackColor = Color.DarkBlue
                            .SelectionColor = Color.White
                            _rtf = .Rtf
                        End With
                    End Using
                    With ActivePane
                        .Rtf = _rtf
                        .SelectionStart = Match.Key
                        Dim CurrentPosition As Point = .GetPositionFromCharIndex(.SelectionStart)
                        If Not .ClientRectangle.Contains(CurrentPosition) Then .ScrollToCaret()
                        Dim WordLocation As Point = .GetPositionFromCharIndex(Match.Key + Match.Value.Length)
                        Dim Bounds_FaR As New Rectangle(.Width - FindAndReplace.Width - .VScrollWidth, WordLocation.Y, FindAndReplace.Width, FindAndReplace.Height)
                        If Bounds_FaR.Contains(WordLocation) Then Bounds_FaR.Offset(0, .LineHeight)
                        With FindAndReplace
                            .Location = Bounds_FaR.Location
                            MoveMouse(.PointToScreen(.Zone_GotoClickPoint))
                            .StartAt += Match.Value.Length
                        End With
                    End With
                End If

            Case Zone.Identifier.ReplaceOne
                If FindAndReplace.CurrentMatch.Key >= 0 Then
                    With FindAndReplace.CurrentMatch
                        Text_Search = Text_Search.Remove(.Key, .Value.Length)
                        Text_Search = Text_Search.Insert(.Key, FindAndReplace.ReplaceControl.Text)
                    End With
                    ActivePane.Text = Text_Search
                    FindAndReplace.DataSource = Text_Search
                    FindRequest()
                End If

            Case Zone.Identifier.ReplaceAll
                If FindAndReplace.CurrentMatch.Key >= 0 Then
                    Dim ReverseOrderMatches = FindAndReplace.Matches.OrderByDescending(Function(x) x.Key)
                    For Each Match In ReverseOrderMatches
                        With Match
                            Text_Search = Text_Search.Remove(.Key, .Value.Length)
                            Text_Search = Text_Search.Insert(.Key, FindAndReplace.ReplaceControl.Text)
                        End With
                    Next
                    ActivePane.Text = Text_Search
                    FindAndReplace.DataSource = Text_Search
                    FindRequest()
                End If

        End Select

    End Sub
    Private Sub FindRequest() Handles FindAndReplace.FindChanged

        Dim SelectionStart As Integer = ActivePane.SelectionStart
        Dim _rtf As String = ActivePane.Rtf
        Using RTB As New RichTextBox With {.Rtf = _rtf}
            With RTB
                For Each Match In FindAndReplace.Matches
                    .SelectionStart = Match.Key
                    .SelectionLength = Match.Value.Length
                    .SelectionBackColor = Color.Yellow
                    .SelectionColor = Color.Black
                Next
                _rtf = .Rtf
            End With
        End Using
        ActivePane.Rtf = _rtf
        ActivePane.SelectionStart = SelectionStart

    End Sub
    Private Sub InsertComment(sender As Object, e As EventArgs) Handles TSMI_Comment.Click

        If ActiveBody.HasText Then
            'RemoveHandler ActivePane.SelectionChanged, AddressOf ActivePane_SelectionChanged
            Dim SelectionStart As Integer = ActivePane.SelectionStart
            Dim SelectionLength As Integer = ActivePane.SelectionLength
            Dim OldTextLength As Integer = ActivePane.Text.Length
            Dim NewBodyText As String = ActiveBody.Text
            Dim LineStarts As New List(Of Match)(From M In Regex.Matches(NewBodyText, "^[^\n\r]{1,}", RegexOptions.Multiline) Select DirectCast(M, Match))
            LineStarts = (From LS In LineStarts Order By LS.Index Descending Where (SelectionStart >= LS.Index And SelectionStart <= (LS.Index + LS.Length)) Or (LS.Index >= SelectionStart And LS.Index < (SelectionStart + SelectionLength))).ToList
            For Each LIneStart In LineStarts
                NewBodyText = NewBodyText.Remove(LIneStart.Index, LIneStart.Length)
                If LIneStart.Value.StartsWith("--", StringComparison.InvariantCulture) Then
                    NewBodyText = NewBodyText.Insert(LIneStart.Index, Regex.Replace(LIneStart.Value, "^[-]{2,}", String.Empty))
                Else
                    NewBodyText = NewBodyText.Insert(LIneStart.Index, "--" & LIneStart.Value)
                End If
            Next
            ActivePane.Text = NewBodyText
            ActivePane.SelectionStart = LineStarts.Last.Index
            ActivePane.SelectionLength = {0, (LineStarts.First.Index + LineStarts.First.Length) - (LineStarts.Last.Index) + (NewBodyText.Length - OldTextLength)}.Max
            'AddHandler ActivePane.SelectionChanged, AddressOf ActivePane_SelectionChanged
            TidyText()
        End If

    End Sub
    Private Sub CopyText(sender As Object, e As EventArgs) Handles TSMI_CopyPlainText.Click, TSMI_CopyColorText.Click

        If ActiveBody.HasText Then
            If sender Is TSMI_CopyColorText Then
                ActivePane.SelectAll()
                ActivePane.Copy()

            ElseIf sender Is TSMI_CopyPlainText Then
                Clipboard.SetText(ActiveBody.SystemText)

            End If
        End If

    End Sub
    Private Sub InsertDividers(sender As Object, e As EventArgs) Handles TSMI_DividerDouble.Click, TSMI_DividerSingle.Click

        Dim Separator As String = If(sender Is TSMI_DividerSingle, StrDup(30, "-"), "--" & StrDup(20, "=")) & vbNewLine
        With ActivePane
            Dim CharIndex = .SelectionStart
            Dim LineNbr = .GetLineFromCharIndex(CharIndex)
            Dim LineStart = .GetFirstCharIndexFromLine(LineNbr)
            .Text = .Text.Insert(LineStart, Separator)
            .SelectionStart = CharIndex
        End With

    End Sub
    Private Function GetCommentMatch() As StringData

        With ActivePane
            If .Lines.Any Then
                If ActiveBody.Labels IsNot Nothing Then
                    Dim Comments = ActiveBody.Labels.Where(Function(x) x.Source = InstructionElement.LabelName.Comment)
                    Dim CharIndex = .SelectionStart
                    Dim LineNbr = .GetLineFromCharIndex(CharIndex)
                    Dim LineStart = .GetFirstCharIndexFromLine(LineNbr)
                    Dim LineLength = .Lines({LineNbr, .Lines.Count - 1}.Min).Length
                    Dim LineMatch = New StringData With {.Start = LineStart, .Length = LineLength}
                    Dim LineComments = (From C In Comments Where LineMatch.Contains(C.Highlight))
                    With TSMI_Comment
                        .Tag = LineNbr
                        If LineComments.Any Then
                            .Text = "Remove Comment".ToString(InvariantCulture)
                            .Image = My.Resources.Comment
                            Return Comments.First.Highlight
                        Else
                            .Text = "Comment".ToString(InvariantCulture)
                            .Image = My.Resources.UnComment
                            Return LineMatch
                        End If
                    End With
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End With

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub TSMI_ShowDropDown(sender As Object, e As EventArgs) Handles TSMI_Connections.MouseEnter, TSMI_Copy.MouseEnter, TSMI_Divider.MouseEnter
        With DirectCast(sender, ToolStripMenuItem)
            .ShowDropDown()
        End With
    End Sub
    Private Sub ColorSelected(sender As Object, e As ImageComboEventArgs)

        Dim ChangedBackColor As Boolean = sender Is IC_BackColor
        Dim NewColor As Color = Color.FromName(e.ComboItem.Text)

        With My.Settings
            Select Case Pane_MouseObject.Source
                Case InstructionElement.LabelName.Comment
                Case InstructionElement.LabelName.Constant
                Case InstructionElement.LabelName.FloatingTable
                    If ChangedBackColor Then
                        .TableFloating_Back = NewColor
                    Else
                        .TableFloating_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.RoutineTable
                    If ChangedBackColor Then
                        .TableRoutine_Back = NewColor
                    Else
                        .TableRoutine_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.SystemTable
                    If ChangedBackColor Then
                        .TableSystem_Back = NewColor
                    Else
                        .TableSystem_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.GroupBlock, InstructionElement.LabelName.GroupField
                    If ChangedBackColor Then
                        .Group_Back = NewColor
                    Else
                        .Group_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.Limit
                    If ChangedBackColor Then
                        .Limit_Back = NewColor
                    Else
                        .Limit_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.OrderBlock, InstructionElement.LabelName.OrderField
                    If ChangedBackColor Then
                        .Order_Back = NewColor
                    Else
                        .Order_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.SelectBlock, InstructionElement.LabelName.SelectField
                    If ChangedBackColor Then
                        .Select_Back = NewColor
                    Else
                        .Select_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.WithBlock
                    If ChangedBackColor Then
                    Else
                        .WithBlock_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.Union
                    If ChangedBackColor Then
                        .Union_Back = NewColor
                    Else
                        .Union_Fore = NewColor
                    End If
            End Select
        End With
        TidyText()

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub LightSwitch(sender As Object, e As EventArgs)

        If LightSwitchedOn() Then
            'TSMI_TipSwitch.Image = Base64ToImage(LightOff)
            TSMI_TipSwitch.Image = My.Resources.LightOff
        Else
            'TSMI_TipSwitch.Image = Base64ToImage(LightOn)
            TSMI_TipSwitch.Image = My.Resources.LightOn
        End If

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub ObjectType_MouseOver(sender As Object, e As EventArgs) Handles TSMI_ObjectType.MouseEnter
        TSMI_ObjectType.ShowDropDown()
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub Pane_TipManager(ToolTipText As String, Location As Point)

        Dim TextValues As String() = Split(ToolTipText, "|")
        If IsNothing(ToolTipText) Then
            TT_PaneTip.Hide(ActivePane)
            CMS_PaneOptions.AutoClose = False
            CMS_PaneOptions.Show(Location)
        Else
            If LightSwitchedOn() Then
                TT_PaneTip.ToolTipTitle = TextValues.First
                TT_PaneTip.Show(TextValues.Last, ActivePane, Location)
            End If
            CMS_PaneOptions.AutoClose = True
            CMS_PaneOptions.Hide()
        End If

    End Sub
    Private Function LightSwitchedOn() As Boolean

        Return SameImage(My.Resources.LightOn, TSMI_TipSwitch.Image)
        Return SameImage(Base64ToImage(LightOn), TSMI_TipSwitch.Image)

    End Function
    '=====================================================================================

#Region " TIDY TEXT + SUPPORTING DECLARATIONS/FUNCTIONS "
    Private SelectionIndex As Integer
    Public Sub TidyText()

        With ActivePane
            Dim ShowChangedText As Boolean = False
            If ShowChangedText Then ShowTextChange(Text, Text)
            Dim TextToTidy As String = .Text
            REM /// REMOVES EXTRA SPACES + EVENLY SPACES UNIONS & COMMAS + COLOR CODES KEY WORDS / SECTIONS
            REM /// THE PROBLEM WITH CHANGING THE TEXT IS THAT IT MOVES YOUR CURSOR POSITION...ANCHOR IT WITH A BLACKOUT BEFORE MODIFYING, THEN LOCATE AFTER CHANGES
            SelectionIndex = .SelectionStart
#Region " TEXT TRANSFORMATION "
            'REMOVE EXTRA LINES
            TextToTidy = RegExText(TextToTidy, "(?<=[\n\r])[\n\r]{1,}")
            'REMOVE EXTRA SPACES
            TextToTidy = RegExText(TextToTidy, "(?<=[ ])[ ]{1,}")
            'REMOVE SPACE PRECEDING COMMA
            TextToTidy = RegExText(TextToTidy, "[ ](?=,)")
            'INSERT SPACE FOLLOWING COMMA
            TextToTidy = RegExText(TextToTidy, "(?<=,)(?=[^\s])", Space,, {"{[^}]{1,}"}.ToList)
            'REMOVE ANY LEADING AND TRAILING SPACES
            TextToTidy = RegExText(TextToTidy, "^ +| +(?=[\n\r]|$)")
            'INSERT 2 CARRIAGE RETURNS BEFORE AND AFTER UNIONS...UNION ALL WILL HAVE ONLY 1 SPACE SINCE PRIOR CHANGES HAVE REMOVED ANY EXTRA SPACES
            For Each Union In Split("UNION ALL|UNION|EXCEPT|INTERSECT", "|")
                Dim UnionBefore As String = "[\n\r\s]{1,}(?=\b■\b)"
                TextToTidy = RegExText(TextToTidy, Replace(UnionBefore, "■", Union), NewLine & NewLine, RegexOptions.IgnoreCase)
                Dim UnionAfter As String = "(?<=\b■\b)[\s\n\r]{0,}"
                If Union = "UNION" Then UnionAfter = "(?<=\bUNION\b(?! ALL))[\s\n\r]{0,}"
                TextToTidy = RegExText(TextToTidy, Replace(UnionAfter, "■", Union), NewLine & NewLine, RegexOptions.IgnoreCase)
            Next
            'INSERT 1 CARRIAGE RETURN BEFORE BELOW KEYWORDS
            For Each KeyWord In Split("WHERE|GROUP|ORDER BY|FETCH|LIMIT|GRANT|REVOKE|ALTER|DROP|INSERT|DELETE|WITH", "|")
                TextToTidy = RegExText(TextToTidy, "[\n\r\s]{1,}(?=\b" & KeyWord & "\b)", NewLine, RegexOptions.IgnoreCase)
            Next
            For Each KeyWord In Split("FROM", "|")
                TextToTidy = RegExText(TextToTidy, "(^, .*){1,}[\s]{1,}(?=\b" & KeyWord & "\b)", NewLine, RegexOptions.IgnoreCase Or RegexOptions.Multiline)
            Next
#End Region
            REM /// TEXT CAN NOW BE COLORED
            .Text = TextToTidy
            .SelectionStart = SelectionIndex
            .Focus()
            RemoveHandler ActiveBody.Completed, AddressOf ColorText
            AddHandler ActiveBody.Completed, AddressOf ColorText
            ActiveBody.Text = TextToTidy
        End With

    End Sub
    Private Sub ColorText(sender As Object, e As EventArgs)

        RemoveHandler ActiveBody.Completed, AddressOf ColorText
        Dim TextToColor As String = ActivePane.Text
        Exit Sub
        Dim PreserveIndex As Integer = ActivePane.SelectionStart
        Dim PreserveScrollIndex As Integer = ActivePane.VScrollPos
        Dim PreserveCursorPosition As Point = Cursor.Position

        If TextToColor.Length = 0 Then
        ElseIf IsNothing(ActiveBody.Labels) Then
        ElseIf Not ActiveBody.Labels.Any Then
        ElseIf IsNothing(ActivePane) Then
        Else
            Using InvisibleRicherTextBox As New RicherTextBox With {.Text = TextToColor, .Font = ActivePane.Font, .Visible = False}
                With InvisibleRicherTextBox
                    .SelectAll()
                    .SelectionColor = Color.Black
                    ActiveBody.Labels.Sort(Function(x, y) x.Source.CompareTo(y.Source))
                    For Each Label In ActiveBody.Labels
                        '.Where(Function(c) c.Source = InstructionElement.LabelName.SystemTable)
                        REM /// USING BOTH THE BLOCK + HIGHLIGHT CREATES A LAYERED EFFECT
                        '.SelectionStart = _Object.Block.Start
                        '.SelectionLength = _Object.Block.Length
                        '.SelectionBackColor = _Object.Block.BackColor
                        '.SelectionColor = _Object.Block.ForeColor
                        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
                        .SelectionStart = Label.Highlight.Start
                        .SelectionLength = Label.Highlight.Length
                        .SelectionBackColor = Label.Highlight.BackColor
                        .SelectionColor = Label.Highlight.ForeColor
                    Next
                    ActivePane.Rtf = .Rtf
                End With
            End Using
        End If
        ActivePane.SelectionStart = PreserveIndex
        If ActivePane.VScrollVisible Then
            ActivePane.VScrollPos = PreserveScrollIndex
            ClickLeftMouseButton(ActivePane.VerticalScrollLocation)
            Cursor.Position = PreserveCursorPosition
        End If

    End Sub
    Private Function RegExText(StringIn As String, Pattern As String, Optional InsertString As String = Nothing, Optional Options As RegexOptions = RegexOptions.Multiline Or RegexOptions.IgnoreCase, Optional Exclusions As List(Of String) = Nothing) As String

        Dim ExclusionText As String = StringIn
        If Exclusions Is Nothing Then
            Exclusions = New List(Of String)({CommentPattern})
        Else
            Exclusions.Add(CommentPattern)
        End If
        For Each Exclusion In Exclusions
            Dim Matches As New List(Of StringData)(From SI In Regex.Matches(StringIn, Exclusion, RegexOptions.IgnoreCase) Select New StringData(SI))
            Matches.Sort(Function(x, y) y.Start.CompareTo(y.Start))
            For Each Match In Matches
                ExclusionText = ExclusionText.Remove(Match.Start, Match.Length)
                ExclusionText = ExclusionText.Insert(Match.Start, StrDup(Match.Length, BlackOut))
            Next
        Next
        Dim List As New List(Of StringData)(From M In Regex.Matches(ExclusionText, Pattern, Options) Select New StringData(M))
        List.Sort(Function(x, y) y.Start.CompareTo(x.Start))
        For Each Item As StringData In List
            If Item.Start < SelectionIndex Then SelectionIndex -= Item.Length
            StringIn = StringIn.Remove(Item.Start, Item.Length)
            If Not IsNothing(InsertString) Then
                StringIn = StringIn.Insert(Item.Start, InsertString)
                If Item.Start <= SelectionIndex Then
                    SelectionIndex += InsertString.Length
                End If
            End If
        Next
        Return StringIn

    End Function
    Private Shared Function ShowTextChange(OriginalText As String, ModifiedText As String) As Boolean

        Dim DeltaOT = OriginalText.ToCharArray
        Dim DeltaTT = ModifiedText.ToCharArray
        Dim CharArray As New List(Of String)

        For T = 0 To ({DeltaOT.Length, DeltaTT.Length}.Max - 1)
            Dim LeftSide As String
            Dim RightSide As String
            If (DeltaOT.Length - 1) < T Then
                LeftSide = BlackOut
            Else
                LeftSide = DeltaOT(T)
                If Regex.Match(LeftSide, "[\n\r]", RegexOptions.Multiline).Success Then
                    LeftSide = "New Line"
                End If
            End If
            If (DeltaTT.Length - 1) < T Then
                RightSide = BlackOut
            Else
                RightSide = DeltaTT(T)
                If Regex.Match(RightSide, "[\n\r]", RegexOptions.Multiline).Success Then
                    RightSide = "New Line"
                End If
            End If
            CharArray.Add(Join({LeftSide, " | ", RightSide}))
        Next
        MsgBox(Join(CharArray.ToArray, vbNewLine), MsgBoxStyle.Information, "Changed=" & Not (OriginalText = ModifiedText))
        Return Not (OriginalText = ModifiedText)

    End Function
#End Region

#End Region
#Region " TABS EVENTS "
    Private Sub Tabs_PageAdded(sender As Object, e As ControlEventArgs) Handles Script_Tabs.ControlAdded

        AddHandler e.Control.TextChanged, AddressOf AutoWidth
        With DirectCast(e.Control, Tab)
            If .Tag IsNot Nothing Then
                Dim TabScript As Script = DirectCast(.Tag, Script)
                If TabScript.Connection IsNot Nothing Then
                    .HeaderBackColor = TabScript.Connection.BackColor
                    .HeaderForeColor = TabScript.Connection.ForeColor
                End If
            End If
        End With

    End Sub
    Private Sub Tabs_PageDropped(sender As Object, e As ControlEventArgs) Handles Script_Tabs.ControlRemoved
        REM /// TABMOUSEREGION WILL BE CURRENT TAB SINCE YOU HAVE TO CLICK THE <X> TO CLOSE, IE) MOUSED OVER TAB BEING DROPPED
        RemoveHandler e.Control.TextChanged, AddressOf AutoWidth
    End Sub
    Private Sub Tabs_TabChange(sender As Object, e As TabsEventArgs) Handles Script_Tabs.TabWidthChanged, Script_Tabs.TabMouseChange

        If e.AfterBounds = Nothing And e.AfterBounds = Nothing Then
            If e.InTab IsNot Nothing Then TabDirection(e.InTab) 'Script_Tabs.PointToScreen(New Point(e.InTab.Bounds.Right, 1))
        Else
            e.InTab.Bounds = e.AfterBounds
            TabDirection(If(e.OutTab, e.InTab))
            'TSDD_SaveAs.Location = Script_Tabs.PointToScreen(New Point(e.AfterBounds.Right, 1))
        End If

    End Sub
    Private Sub Tabs_ZoneChange(sender As Object, e As TabsEventArgs) Handles Script_Tabs.ZoneMouseChange

        TabDirection(If(e.OutTab, e.InTab))
        TT_Tabs.Hide(Script_Tabs)
        Dim TipLocation = Script_Tabs.PointToScreen(If(e.InTab, e.OutTab).Bounds.Location)
        TipLocation.Offset(ActiveTab.Bounds.Width, 3)

        Select Case e.InZone
            Case Tabs.Zone.None

            Case Tabs.Zone.Add
                If Not ScriptsInitialized Then Tabs_TipManager("Please wait|Collection initializing", TipLocation)

            Case Tabs.Zone.Image
                Dim TipValues As String = Nothing
                With ActiveScript
                    TipValues = "Run Script|" & Bulletize({"Current datasource is " & If(IsNothing(.Connection), "undetermined", .DataSourceName),
                                            "Type is " & If(.Body.InstructionType = ExecutionType.Null, "undetermined", .Body.InstructionType.ToString),
                                            Join({"Text has", If(.TextWasModified, String.Empty, " not"), " changed"}, String.Empty),
                                            Join({"Last modified", .Modified.ToShortDateString, "@", .Modified.ToShortTimeString}),
                                            Join({"Last successful run", .Ran.ToShortDateString, "@", .Ran.ToShortTimeString}),
                                            "Location=" & If(.Path, "None - not saved")})
                End With
                Tabs_TipManager(TipValues, TipLocation)

            Case Tabs.Zone.Text
                Tabs_TipManager("Reorder tab|Drag tab and drop in new position", TipLocation)

            Case Tabs.Zone.Close
                If Not ActiveScript.Body.HasText Then
                    Tabs_TipManager("Close Tab|Click to close empty tab", TipLocation)

                ElseIf ActiveScript.FileTextMatchesText Then
                    Tabs_TipManager("Close Tab|Click to close saved script", TipLocation)

                Else

                End If
        End Select

    End Sub
    Private Sub Tabs_TipManager(ToolTipText As String, Location As Point)

        If IsNothing(ToolTipText) Then
            TT_Tabs.Hide(Script_Tabs)

        Else
            Dim TextValues As String() = Split(ToolTipText, "|")
            TT_Tabs.ToolTipTitle = TextValues.First
            TT_Tabs.Show(TextValues.Last, Script_Tabs, Location)

        End If

    End Sub
    Private Sub Tab_Clicked(sender As Object, e As TabsEventArgs) Handles Script_Tabs.TabClicked

        Hide_DropDowns()

        Select Case e.InZone
            Case Tabs.Zone.Add
                Scripts.Add(New Script With {._Tabs = Script_Tabs, .State = Script.ViewState.OpenDraft})
                Dim paneActive = ActivePane

            Case Tabs.Zone.Image
#Region " RUN "
                If IsNothing(ActiveScript.Connection) Then
                    REM /// DSN HAS NOT YET BEEN SET. CLICKING THE TAB IS AUTO-SELECT
                    Dim Tables = ActiveScript.Body.TablesFullName   '<======== NOT SURE ABOUT THIS
                    If Tables.Any Then
                        Message.Show("Datasource not found. One must be selected.", "The following tables do not exist in your saved items: " & vbNewLine & Join(Tables.ToArray, ","), Prompt.IconOption.Critical, Prompt.StyleOption.Blue)
                    Else
                        Message.Show("Datasource not found.", "One must be selected.", Prompt.IconOption.Critical, Prompt.StyleOption.Grey)
                    End If
                Else
                    REM /// DSN HAS BEEN SET. USE CURRENT VALUE. USER CAN CHANGE FROM ScriptPane_Run_DSNs (TSMI)
                    RunScript()
                End If
#End Region
            Case Tabs.Zone.Text

            Case Tabs.Zone.Close
#Region " CLOSE "
                With ActiveScript
                    If .FileCreated Then
                        If .FileTextMatchesText Then
                            'NO CHANGES...DO NOTHING
                            .State = Script.ViewState.ClosedNotSaved
                        Else
                            If Message.Show(.Body.InstructionType.ToString & " has changed", "Save your work?", Prompt.IconOption.YesNo, Prompt.StyleOption.Grey) = DialogResult.No Then
                                'DO NOT WANT CHANGES SAVED...TEXT NEEDS TO REVERT TO FILE TEXT
                                .State = Script.ViewState.ClosedNotSaved
                            Else
                                'DO WANT CHANGES SAVED
                                .State = Script.ViewState.ClosedSaved
                            End If
                        End If

                    ElseIf .Body.HasText Then
                        Using message As New Prompt With {.TitleBarImage = My.Resources.Warning_.ToBitmap}
                            If message.Show("You have unsaved work. Continue?",
                                            "[Yes] to discard, [No] to cancel",
                                            Prompt.IconOption.YesNo,
                                            Color.GhostWhite,
                                            Color.DarkGray,
                                            Color.White,
                                            Color.GhostWhite,
                                            Color.GhostWhite,
                                            Color.DarkBlue) = DialogResult.No Then
                                'DO NOTHING AND LEAVE TAB OPEN
                            Else
                                'NO FILE, WITH TEXT...DISCARD EMPTY TAB
                                .State = Script.ViewState.None
                            End If
                        End Using

                    Else
                        'NO FILE, NO TEXT...DISCARD EMPTY TAB
                        .State = Script.ViewState.None
                    End If
                End With
#End Region
        End Select
    End Sub
#End Region

    Public Sub AddPane(Instruction As String, Optional Run As Boolean = False)

        Dim startupScript As Script = Scripts.Add(New Script With {._Tabs = Script_Tabs,
                                              .State = Script.ViewState.OpenDraft,
                                              .Name = "adhoc",
                                              .Text = Instruction})
        Dim paneActive = ActivePane
        If Run Then AddHandler Scripts_.CollectionChanged, AddressOf ScriptsLoaded

    End Sub
    Private Sub ScriptsLoaded(sender As Object, e As ScriptsEventArgs)

        If e.State = CollectionChangeAction.Refresh Then
            RemoveHandler Scripts_.CollectionChanged, AddressOf ScriptsLoaded
            Dim startupScript As Script = Scripts.Item("adhoc")
            RunScript(startupScript)
        End If

    End Sub

#Region " OBJECT EVENTS "
#Region " Tree_Objects POPULATION "
    Private ReadOnly Property SelectedConnections As List(Of Connection)
        Get
            Return (From n In Tree_Objects.SelectedNodes Where n Is n.Root Select DirectCast(n.Tag, Connection)).ToList
        End Get
    End Property
    Private RequestInitiated As Boolean
    Private WithEvents SpinTimer As New Timer With {.Interval = 250, .Tag = 0}
    Private Sub ObjectSyncClicked(sender As Object, e As EventArgs) Handles Button_ObjectsSync.Click

        ' *** Correct any discrepancies between SystemObjects and Database ***
        'SelectedConnections ie) User decides which items to update ( NOT USED )
        If Not RequestInitiated Then
            RequestInitiated = True
            If ObjectsSet.Tables.Count = 0 And Not ObjectsWorker.IsBusy Then
                SpinTimer.Start()
#Region " SQL BARRAGE "
                Using ObjectsTable As DataTable = SystemObjects.ToDataTable
                    Dim OwnersNames = From ot In ObjectsTable.AsEnumerable Group ot By _Server = ot("DataSource").ToString Into SourceGrp = Group
                                      Select New With {
                                     .Server = _Server,
                                     .TablesViews = New List(Of String)(From sg In SourceGrp Where sg("DataSource").ToString = _Server And {"Table", "View"}.Contains(sg("Type").ToString) Select ValueToField(Join({sg("Owner").ToString, sg("Name").ToString}, ".")))
                                     }
                    SyncWorkers = New Dictionary(Of String, BackgroundWorker)
                    SyncSet = New Dictionary(Of String, DataTable)
                    For Each DataSource In OwnersNames
                        Dim SyncWorker = New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = False}
                        Dim DB_Alias As String = Aliases(DataSource.Server)
                        Dim SyncNode As Node = Tree_Objects.Nodes.Item(DB_Alias)
                        Dim Connection As Connection = DirectCast(SyncNode.Tag, Connection)
                        SyncWorkers.Add(DB_Alias, SyncWorker)
                        SyncSet.Add(DB_Alias, Nothing)
                        SyncNode.BackColor = Color.FromArgb(64, Color.LimeGreen)
                        Dim ObjectSQL As String = Replace(My.Resources.SQL_DATAOBJECTS, "--WHERE OWNER='///OWNER///'", "WHERE TRIM(OWNER)||'.'||TRIM(NAME) In (" + Join(DataSource.TablesViews.ToArray, ",") + ")")
                        With New SQL(Connection, ObjectSQL) With {.Name = DB_Alias}
                            AddHandler .Completed, AddressOf SyncSQL_Completed
                            .Execute()
                        End With
                    Next
                End Using
#End Region
            End If
        End If

    End Sub
    Private Sub SyncSQL_Completed(sender As Object, e As ResponseEventArgs)

        With DirectCast(sender, SQL)
            RemoveHandler .Completed, AddressOf SyncSQL_Completed
            Dim SyncNode As Node = Tree_Objects.Nodes.Item(.Name)
            SyncNode.Separator = Node.SeparatorPosition.Above
            If e.Succeeded Then
                Dim NodeConnection As Connection = DirectCast(SyncNode.Tag, Connection)
                Dim DatabaseColor As Color = If(NodeConnection Is Nothing, Color.Blue, NodeConnection.BackColor)
                SyncNode.Image = ChangeImageColor(My.Resources.Sync, Color.FromArgb(255, 64, 64, 64), DatabaseColor)
            Else
                SyncNode.Image = ChangeImageColor(My.Resources.Sync, Color.FromArgb(255, 64, 64, 64), Color.Red)
            End If
            SyncWorkers.Remove(.Name)
            SyncSet.Item(.Name) = .Table
        End With

        If Not SyncWorkers.Any Then
            SpinTimer.Stop()
            Tree_Objects.BackgroundImage = Nothing
            Using ObjectsTable As DataTable = SystemObjects.ToDataTable
                Dim GroupedTables = From ot In ObjectsTable.AsEnumerable Group ot By _Name = ot("DataSource").ToString Into SourceGrp = Group
                                    Select New With {.Name = _Name, .Table = SourceGrp.CopyToDataTable}
                For Each DataSource In GroupedTables
                    If SyncSet.ContainsKey(DataSource.Name) Then
                        Dim SyncTable As DataTable = SyncSet(DataSource.Name)
                        Dim ObjectTable As DataTable = DataSource.Table
                        Dim Objects_Server As New SystemObjectCollection(SyncTable)
                        Dim Objects_Local As New SystemObjectCollection(ObjectTable)
                        Dim Objects_Remove As New List(Of SystemObject)
                        Dim Objects_Modify As New Dictionary(Of SystemObject, SystemObject)
                        For Each Local_Object In Objects_Local
                            Dim Server_Object As SystemObject = Objects_Server.Item(Local_Object.Key)
                            If Server_Object Is Nothing Then
                                'DSN+Owner+Name not in SyncTable ==> Drop from ObjectTable
                                Objects_Remove.Add(Local_Object)

                            ElseIf Server_Object.ToString = Local_Object.ToString Then
                                'OK

                            Else
                                'Mismatch ==> Modify ObjectTable
                                Objects_Modify.Add(Local_Object, Server_Object)
                            End If
                        Next
                        For Each Item In Objects_Remove
                            SystemObjects.Remove(Item.Key)
                        Next
                        For Each Item In Objects_Modify
                            SystemObjects.Remove(Item.Key.Key)
                            SystemObjects.Add(Item.Value)
                        Next
                    End If
                Next
                SystemObjects.RemoveDuplicates()
                For Each Level1Node In Tree_Objects.Nodes
                    Level1Node.Nodes.Clear()
                Next
                RequestInitiated = False
                ObjectsWorker.RunWorkerAsync()
            End Using
        End If

    End Sub
    Private Sub SpinTimer_Tick() Handles SpinTimer.Tick

        Dim SpinImageIndex As Integer = DirectCast(SpinTimer.Tag, Integer) Mod 8
        If SpinImageIndex = 0 Then Tree_Objects.BackgroundImage = My.Resources.Spin1
        If SpinImageIndex = 1 Then Tree_Objects.BackgroundImage = My.Resources.Spin2
        If SpinImageIndex = 2 Then Tree_Objects.BackgroundImage = My.Resources.Spin3
        If SpinImageIndex = 3 Then Tree_Objects.BackgroundImage = My.Resources.Spin4
        If SpinImageIndex = 4 Then Tree_Objects.BackgroundImage = My.Resources.Spin5
        If SpinImageIndex = 5 Then Tree_Objects.BackgroundImage = My.Resources.Spin6
        If SpinImageIndex = 6 Then Tree_Objects.BackgroundImage = My.Resources.Spin7
        If SpinImageIndex = 7 Then Tree_Objects.BackgroundImage = My.Resources.Spin8
        SpinTimer.Tag = SpinImageIndex + 1

    End Sub
    Private Sub LoadSystemObjects(sender As Object, e As EventArgs) Handles ObjectsWorker.DoWork

        Dim LoadFromSettings As Boolean = sender Is Nothing
        Dim ClockLoadTime As Boolean = False
        If ClockLoadTime Then Stop_Watch.Start()
        ExpandCollapseOnOff(HandlerAction.Remove)
#Region " FILL TABLE WITH DATABASE OBJECTS "
        Dim ActiveConnections = Connections.Where(Function(c) c.CanConnect And c.IsDB2).Take(1000)
        If SelectedConnections IsNot Nothing AndAlso SelectedConnections.Any Then ActiveConnections = ActiveConnections.Where(Function(c) SelectedConnections.Contains(c))
        Dim SuccessCount As Integer = 0
        For Each Connection In ActiveConnections
            Dim _Alias As String = Connection.DataSource
            If Connection.DataSource = "CDNIW" Then _Alias = "TORDSNQ"
            If Not Aliases.ContainsKey(_Alias) Then
                Aliases.Add(_Alias, Connection.DataSource)
                ConnectionsDictionary.Add(Connection.DataSource, LoadFromSettings)
                Dim ConnectionTable As New DataTable
                If Not LoadFromSettings Then
                    ConnectionTable = RetrieveData(Connection.ToString, My.Resources.SQL_DATAOBJECTS)
                    ConnectionsDictionary(Connection.DataSource) = ConnectionTable IsNot Nothing AndAlso ConnectionTable.Columns.Count > 0
                End If
                If ConnectionsDictionary(Connection.DataSource) Then
                    If Not LoadFromSettings Then ObjectsSet.Tables.Add(ConnectionTable)
                    SuccessCount += 1
                    Dim DatabaseColor As Color = If(Connection Is Nothing, Color.Blue, Connection.BackColor)
                    Tree_Objects.Nodes.Add(New Node With {.Text = Connection.DataSource,
                                                    .Name = .Text,
                                                    .Image = ChangeImageColor(My.Resources.Sync, Color.FromArgb(255, 64, 64, 64), DatabaseColor),
                                                    .Separator = Node.SeparatorPosition.Above,
                                                    .Tag = Connection,
                                                    .AllowAdd = False,
                                                    .AllowDragDrop = False,
                                                    .AllowEdit = False,
                                                    .AllowRemove = False})
                End If
                If ClockLoadTime Then
                    Intervals.Add(Connection.DataSource, Stop_Watch.Elapsed)
                    Stop_Watch.Restart()
                End If
            End If
        Next
#End Region
#Region "ODBC.txt - ALIAS CDNIW, TargetDatabase = TORDSNQ "
        '        [DB>NDEEFA28TORDSNQ]
        '        Dir_entry_type = REMOTE
        '        Authentication = NOTSPEC
        '        DBName = TORDSNQ

        '        [DB>TORSTL3CDNIW]
        '        Dir_entry_type = DCS
        '        Authentication = SERVER
        '        DBName = CDNIW
        '        Comment = Canadian IW prod
        '        TargetDatabase = TORDSNQ

        '        [CLI_ODBC>CDNIW]
        '        DataSourceName = CDNIW
        '        DataSourceType = System
        '        AsyncEnable = 0
        '        DBALIAS = CDNIW
#End Region
#Region " POPULATE ObjectsDictionary "
        'DataSources/Owners/Types
        Dim _Objects As New List(Of SystemObject)
        If LoadFromSettings Then
            _Objects.AddRange(SystemObjects)
        Else
            For Each ObjectsTable As DataTable In ObjectsSet.Tables
                _Objects.AddRange(ObjectsTable.AsEnumerable.Select(Function(r) New SystemObject(r)))
            Next
        End If
        For Each _SystemObject In _Objects
            With _SystemObject
                If Not ObjectsDictionary.ContainsKey(.DSN) Then
                    ObjectsDictionary.Add(.DSN, New Dictionary(Of String, Dictionary(Of SystemObject.ObjectType, List(Of SystemObject))))
                End If
                If Not ObjectsDictionary(.DSN).ContainsKey(.Owner) Then
                    ObjectsDictionary(.DSN).Add(.Owner, New Dictionary(Of SystemObject.ObjectType, List(Of SystemObject)))
                End If
                If Not ObjectsDictionary(.DSN)(.Owner).ContainsKey(.Type) Then
                    ObjectsDictionary(.DSN)(.Owner).Add(.Type, New List(Of SystemObject))
                End If
                ObjectsDictionary(.DSN)(.Owner)(.Type).Add(_SystemObject)
            End With
        Next
#End Region
        If ClockLoadTime Then
            Intervals.Add("Iterate DataTable", Stop_Watch.Elapsed)
            Stop_Watch.Restart()
        End If
#Region " LOAD OBJECT TREEVIEW - FROM SYSTEM OBJECTS ( NOT DATASOURCE ) "
        With Tree_Objects
            For Each DataSource In ObjectsDictionary.Keys
                Dim SourceNode = .Nodes.Item(Aliases(DataSource))
                SourceNode.Name = Aliases(DataSource)
                Dim _Connection As Connection = DirectCast(SourceNode.Tag, Connection)
                Dim Owners = ObjectsDictionary(DataSource)
                For Each Owner In Owners
                    Dim OwnerNode = SourceNode.Nodes.Add(New Node With {.Text = Owner.Key,
                                            .Name = Owner.Key,
                                            .BackColor = If(Owner.Key = _Connection.UserID, Color.Gainsboro, Color.Transparent),
                                            .AllowAdd = False,
                                            .AllowDragDrop = False,
                                            .AllowEdit = False,
                                            .AllowRemove = False})
                    For Each ObjectType In Owner.Value
                        Dim TypeImage As Image = Nothing
                        If ObjectType.Key = SystemObject.ObjectType.Routine Then TypeImage = My.Resources.Gear
                        If ObjectType.Key = SystemObject.ObjectType.Table Then TypeImage = My.Resources.Table
                        If ObjectType.Key = SystemObject.ObjectType.Trigger Then TypeImage = My.Resources.Zap
                        If ObjectType.Key = SystemObject.ObjectType.View Then TypeImage = My.Resources.Eye
                        For Each Item In ObjectType.Value
                            Dim NameNode As Node = OwnerNode.Nodes.Add(New Node With {.Name = Item.Name,
                                          .Text = Item.Name,
                                          .Image = TypeImage,
                                          .Checked = True,
                                          .Tag = Item,
                                          .AllowAdd = False})
                        Next
                    Next
                Next
            Next
            If ClockLoadTime Then
                Intervals.Add("Populate Treeview", Stop_Watch.Elapsed)
                Stop_Watch.Stop()
                Using SW As New StreamWriter("C:\Users\SEANGlover\Desktop\Intervals.txt")
                    For Each _Split In Intervals
                        SW.WriteLine(Join({_Split.Key, _Split.Value.ToString}, vbTab))
                    Next
                End Using
            End If
            .CheckBoxes = TreeViewer.CheckState.All
            .BackgroundImage = Nothing
        End With
        SpinTimer.Stop()
#End Region

    End Sub
    Private Sub ObjectTreeviewLoaded() Handles ObjectsWorker.RunWorkerCompleted
        ObjectsTreeview_AutoWidth(Nothing, Nothing)
        ExpandCollapseOnOff(HandlerAction.Add)
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
#End Region

#Region " Tree_Objects DRAG ONTO PANE Or GRID [ V I E W   S T R U C T U R E   Or   C O N T E N T ] "
    Private Sub ObjectNode_StartDrag(sender As Object, e As NodeEventArgs) Handles Tree_Objects.NodeDragStart

        'Root=DataSource, Level 1=Owner, Level 2=Name + Image {Trigger, Table, View, Routine}
        If ActivePane IsNot Nothing And e.Node.Level = 2 Then   'Dragging a Table, Routine, Trigger or View
#Region " DRAG UP Or DOWN ON Tree_Objects "

#End Region
#Region " DRAGGED NODE EXITS Tree_Objects "

#End Region
            Dim A_S = ActiveScript
            Dim NodeObject = NodeProperties(e.Node)
            ActivePane.AllowDrop = True
            Script_Grid.AllowDrop = True
            Select Case NodeObject.Type
                Case SystemObject.ObjectType.Table, SystemObject.ObjectType.View
                    'Pane shows Table/View structure while Grid shows Content
                    'Initiate threads for each so when dropped it's done
                    SpinTimer.Start()
                    Dim SQL_Sample As String = Join({"SELECT *", "FROM " & NodeObject.FullName, "FETCH FIRST 50 ROWS ONLY"}, vbNewLine)
                    Dim SQL_Structure As String = ColumnSQL(NodeObject.FullName)
                    With Jobs
                        .Clear()
                        .Add(New Job(New SQL(NodeObject.Connection, SQL_Sample) With {.Name = "50 Row Sample"}) With {.Name = "50 Row Sample"})
                        .Add(New Job(New SQL(NodeObject.Connection, SQL_Structure) With {.Name = "Table Structure"}) With {.Name = "Table Structure"})
                        AddHandler .Completed, AddressOf TableContentStructureRetrieved
                        .Execute()
                    End With

                Case SystemObject.ObjectType.Routine

                Case SystemObject.ObjectType.Trigger

            End Select
        End If

    End Sub

    Private Sub Pane_DragEnter(sender As Object, e As DragEventArgs) Handles ActivePane_.DragEnter
        'Stop
    End Sub
    Private Sub Pane_NodeDropped(e As DragEventArgs)

        Hide_DropDowns()
        Dim DroppedNode As Node = DirectCast(e.Data.GetData(GetType(Node)), Node)
        If DroppedNode IsNot Nothing Then
            With DroppedNode
                If .TreeViewer Is Tree_Objects Then
                    AutoWidth(ActivePane)

                ElseIf .TreeViewer Is Tree_ClosedScripts Then
                    If .Tag.GetType = GetType(Script) Then
                        Dim ClosedScript As Script = DirectCast(.Tag, Script)
                        Dim PanesNoText As New List(Of Script)(From S In Scripts Where S.State = Script.ViewState.OpenDraft And Not S.Body.HasText)
                        If PanesNoText.Any Then
                            REM /// DROP THE TAB *** INSERT AT ITS POSITION
                            PanesNoText.First.State = Script.ViewState.None
                        End If
                        ClosedScript.State = Script.ViewState.OpenSaved
                    End If

                End If
            End With
        End If

    End Sub

    Private Sub Grid_NodeDropped(sender As Object, e As DragEventArgs) Handles Script_Grid.DragDrop

        Dim DroppedNode As Node = DirectCast(e.Data.GetData(GetType(Node)), Node)
        With DroppedNode
            If .TreeViewer Is Tree_Objects Then

            ElseIf .TreeViewer Is Tree_ClosedScripts Then
                Dim _Script As Script = DirectCast(.Tag, Script)
                RunScript(_Script)

            End If
        End With

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Function NodeProperties(_Node As Node) As SystemObject

        'Root=DataSource, Level 1=Owner, Level 2=Name + Image {Trigger, Table, View, Routine}
        With _Node
            If .Level = 2 Then
                Dim Item_Connection As Connection = DirectCast(.Root.Tag, Connection)
                Dim Item_Owner As String = .Parent.Name
                Dim Item_Name As String = .Name
                Dim Item_Type As SystemObject.ObjectType = Image_Type(.Image)
                Return New SystemObject() With {.DSN = Item_Connection.DataSource,
                                .Owner = Item_Owner,
                                .Type = Item_Type,
                                .Name = Item_Name}
            Else
                Return Nothing
            End If
        End With

    End Function
    Private Sub TableContentStructureRetrieved(sender As Object, e As ResponsesEventArgs)

        With DirectCast(sender, JobCollection)
            RemoveHandler .Completed, AddressOf TableContentStructureRetrieved
            Script_Grid.DataSource = .Item("50 Row Sample").SQL.Table
            Dim Columns = DataTableToListOfColumnsProperties(.Item("Table Structure").SQL.Table)
            Dim TableColumns As String = ColumnPropertiesToTableViewProcedure(Columns)
            ActivePane.Text = CreateTableText(TableColumns)
        End With
        SpinTimer.Stop()
        Tree_Objects.BackgroundImage = Nothing

    End Sub
#End Region

#Region " ActivePane.Text Or Script_Tabs.Tab Or Script_Grid DRAG ONTO Tree_Objects [ E T L ] "
    Private ReadOnly Data As New DataObject
    Private Sub Pane_StartDrag(sender As Object, e As DragEventArgs) Handles ActivePane_.DragStart
        Data.SetData(ActivePane_.GetType, ActivePane_)
    End Sub
    Private Sub Tab_StartDrag(sender As Object, e As TabsEventArgs) Handles Script_Tabs.TabDragDrop
        Data.SetData(Script_Tabs.GetType, Script_Tabs)
    End Sub
    Private Sub Pane_DragOver() Handles ActivePane_.DragOver

        Dim Grid = Data.GetData(GetType(DataTool))
        If Grid IsNot Nothing Then
            TLP_PaneGrid.ColumnStyles(0).Width = ObjectsWidth
        End If

    End Sub
    Private Sub TreeObjects_DragDrop(sender As Object, e As DragEventArgs) Handles Tree_Objects.DragDrop

        Dim DropNode As Node = Tree_Objects.HitTest(Tree_Objects.PointToClient(Cursor.Position)).Node
        Dim DragObject As Object = Data.GetData(GetType(Object))

        Dim Grid = Data.GetData(GetType(DataTool))
        Stop
        If Grid IsNot Nothing AndAlso DropNode IsNot Nothing Then
            If DropNode.IsRoot Then
                'Dropping to Database (Root Level)...Locate TableSpace + Create Table?
                Dim ConnectionNode As Connection = DirectCast(DropNode.Tag, Connection)
                With New SQL(ConnectionNode, Replace(My.Resources.SQL_DATAOBJECTS,
                                                     "--WHERE OWNER='///OWNER///'",
                                                     "WHERE TYPE='TABLE' AND OWNER=" & ValueToField(ConnectionNode.UserID)))
                    AddHandler .Completed, AddressOf TableSpaces
                    .Execute()
                End With

            ElseIf DropNode.Level = 1 Then
                'Dropping to Owner Level...Create Table? TableSpace is known

            ElseIf SameImage(DropNode.Image, My.Resources.Table) Then
                'Dropping Data onto an Existing table...Clear Rows Or Add to Existing?

            End If
        End If

    End Sub
    Private Sub ObjectNode_DroppedOnTreeView(sender As Object, e As NodeEventArgs) Handles Tree_Objects.NodeDropped

        Dim DropNode As Node = e.Node
        If DragNode Is Nothing Then

        ElseIf DragNode.TreeViewer Is Tree_Objects Then
            If Not DragNode Is DropNode Then

            End If
        ElseIf DragNode.TreeViewer Is Tree_ClosedScripts Then
            REM /// FUTURE OPTION TO SCHEDULE JOBS
            Dim SourceScript As Script = DirectCast(DragNode.Tag, Script)
            Dim DestinationObject As SystemObject = DirectCast(DropNode.Tag, SystemObject)
            ObjectTreeview_ETL(SourceScript, DestinationObject)
        End If

    End Sub
#End Region

#Region " MAKE CHANGES TO / SEARCH THE DATABASE "
    Private Sub ObjectNodeRemoveRequested(sender As Object, e As NodeEventArgs) Handles Tree_Objects.NodeBeforeRemoved

        Dim _Connection As Connection = DirectCast(e.Node.Root.Tag, Connection)
        Dim NodeObject As SystemObject = DirectCast(e.Node.Tag, SystemObject)

        Dim Remove_Owner As String = NodeObject.Owner
        Dim Remove_Name As String = NodeObject.Name
        Dim Remove_Type As String = NodeObject.Type.ToString
        Dim Remove_Message As String = Join({"You are about to drop the", Remove_Type, Remove_Name})

        Dim Remove_OK As Boolean = False

        If Message.Show("Proceed with removal from " & _Connection.DataSource & "?", Remove_Message, Prompt.IconOption.YesNo, Prompt.StyleOption.Blue) = DialogResult.Yes Then
            Dim SQL_Dependants As String = My.Resources.SQL_DEPENDANTS
            SQL_Dependants = Replace(SQL_Dependants, "//OWNER//", Remove_Owner)
            SQL_Dependants = Replace(SQL_Dependants, "//NAME//", Remove_Name)
            SQL_Dependants = Replace(SQL_Dependants, "//TYPE//", Remove_Type)
            Dim DependantTable As DataTable = RetrieveData(_Connection.ToString, SQL_Dependants)
            If IsNothing(DependantTable) Then
                'CONNECTION FAILED
            ElseIf DependantTable.Rows.Count = 0 Then
                'NO DEPENDANTS...OK TO DROP
                Remove_OK = True
            Else
                Dim Dependant_Message As New List(Of String)
                For Each DependantRow In DependantTable.AsEnumerable
                    Dim ReferenceCount As Integer = Convert.ToInt32(DependantRow("REFERENCES"), InvariantCulture)
                    Dim ReferenceName As String = DependantRow("ITEM").ToString
                    Dim ReferenceType As String = DependantRow("DEPENDANT_TYPE").ToString
                    Dim Statement As String = Join({If(Dependant_Message.Any, "and there", "There"), If(ReferenceCount = 1, "is", "are"), ReferenceCount, If(ReferenceCount = 1, "reference", "references"), "to the", Remove_Type, Remove_Name, "in the", ReferenceType, ReferenceName})
                    Dependant_Message.Add(Statement)
                Next
                Message.Datasource = DependantTable
                If Message.Show("Are you certain? Other dependant objects will be dropped too", Join(Dependant_Message.ToArray, vbNewLine), Prompt.IconOption.YesNo) = DialogResult.Yes Then
                    Remove_OK = True
                Else
                    Message.Show("Operation cancelled", "No change to " & _Connection.DataSource, Prompt.IconOption.TimedMessage, Prompt.StyleOption.Blue)
                End If
            End If
        Else
            Message.Show("Operation cancelled", "No change to " & _Connection.DataSource, Prompt.IconOption.TimedMessage, Prompt.StyleOption.Blue)
        End If
        If Remove_OK Then
            Dim Drop_DDL As String = Join({"DROP", Remove_Type, NodeObject.FullName})
            With New DDL(_Connection.ToString, Drop_DDL, False, False)
                .Execute()
            End With
            With SystemObjects
                .Remove(NodeObject)
                .Save()
            End With
        End If
        e.Node.CancelAction = Remove_OK
        UpdateNodeText(e.Node)

    End Sub
    Private Sub ObjectTreeView_Search(sender As Object, e As ImageComboEventArgs) Handles IC_ObjectsSearch.ValueSubmitted
        MsgBox(IC_ObjectsSearch.Text)
    End Sub
#End Region
#Region " MAKE CHANGES TO SYSTEMOBJECTS FILE - 2 SOURCES {BULK IMPORT Or RunQuery RESULTS} "
    Private Sub ObjectNode_Checked(sender As Object, e As NodeEventArgs) Handles Tree_Objects.NodeChecked

        'Root=DataSource, Level 1=Owner, Level 2=Type {Trigger, Table, View, Routine}, Level 3=Name

        Dim BaseNodes = Tree_Objects.Nodes.All.Where(Function(n) n.Level = 2 And n.Checked)
        Dim CheckedObjects As New List(Of SystemObject)(BaseNodes.Select(Function(n) DirectCast(n.Tag, SystemObject)))
        Dim CheckedStrings As String() = (From CO In CheckedObjects Select CO.ToString & String.Empty).ToArray
        Dim CheckedString As String = Join(CheckedStrings, vbNewLine)

        Dim MySettingsObjects = SystemObjects.ToStringList

        Dim Items_Removed As New List(Of String)(MySettingsObjects.Except(CheckedStrings))
        Dim Items_Added As New List(Of String)(CheckedStrings.Except(MySettingsObjects))

        With SystemObjects
            For Each Item_Removed In Items_Removed
                .Remove(Item_Removed)
            Next
            For Each Item_Added In Items_Added
                .Add(Item_Added)
            Next
            .Save()
        End With
        UpdateNodeText(e.Node)

    End Sub
    Private Sub UpdateNodeText(NodeItem As Node)

        For Each ParentNode In NodeItem.Parents
            Dim Numerator As String = ParentNode.Children.Where(Function(n) n.Checked).Count.ToString(InvariantCulture)
            Dim Denominator As String = If(ObjectsSet.Tables.Count = 0, "---", ParentNode.Children.Count.ToString(InvariantCulture))
            ParentNode.Text = Join({ParentNode.Name, " (", Numerator, "/", Denominator, ")"}, String.Empty)
        Next

    End Sub
#End Region
    Private Shared Function Image_Type(_Image As Image) As SystemObject.ObjectType

        If SameImage(_Image, My.Resources.Gear) Then
            Return SystemObject.ObjectType.Routine
        ElseIf SameImage(_Image, My.Resources.Table) Then
            Return SystemObject.ObjectType.Table
        ElseIf SameImage(_Image, My.Resources.Zap) Then
            Return SystemObject.ObjectType.Trigger
        ElseIf SameImage(_Image, My.Resources.Eye) Then
            Return SystemObject.ObjectType.View
        Else
            Return SystemObject.ObjectType.None
        End If

    End Function
    Private Function Type_Image(_ObjectType As SystemObject.ObjectType) As Image

        If _ObjectType = SystemObject.ObjectType.Routine Then
            Return My.Resources.Gear
        ElseIf _ObjectType = SystemObject.ObjectType.Table Then
            Return My.Resources.Table
        ElseIf _ObjectType = SystemObject.ObjectType.Trigger Then
            Return My.Resources.Zap
        ElseIf _ObjectType = SystemObject.ObjectType.View Then
            Return My.Resources.Eye
        Else
            Return Nothing
        End If

    End Function
    Private Sub ExpandCollapseOnOff(Action As HandlerAction)
        If Action = HandlerAction.Add Then
            AddHandler Tree_Objects.NodeExpanded, AddressOf ObjectsTreeview_AutoWidth
            AddHandler Tree_Objects.NodeExpanded, AddressOf ObjectsTreeview_AutoWidth
        Else
            RemoveHandler Tree_Objects.NodeExpanded, AddressOf ObjectsTreeview_AutoWidth
            RemoveHandler Tree_Objects.NodeExpanded, AddressOf ObjectsTreeview_AutoWidth
        End If
    End Sub
    Private Sub ObjectTreeview_ETL(SourceScript As Script, DestinationObject As SystemObject)

        REM /// FUTURE OPTION TO SCHEDULE JOBS
        If SourceScript.Body.InstructionType = ExecutionType.SQL And DestinationObject.Type = SystemObject.ObjectType.Table Then
            With New ETL
                .Sources.Add(New ETL.Source(SourceScript.Connection, SourceScript.Text))
                .Destinations.Add(New ETL.Destination(DestinationObject.Connection, DestinationObject.FullName) With {.ClearTable = Message.Show("Clear destination table?", "Select YES to clear, NO to append new rows", Prompt.IconOption.YesNo) = DialogResult.Yes})
                AddHandler .Completed, AddressOf ETL_Completed
                .Name = Join({SourceScript.Name, DestinationObject.Name}, " ==> ")
                .Execute()
            End With
        Else
        End If

    End Sub
    Private Sub ETL_Completed(sender As Object, e As ResponsesEventArgs)
        With DirectCast(sender, ETL)
            RemoveHandler .Completed, AddressOf ETL_Completed
            Message.Show("Transfer request completed " & If(.Succeeded, String.Empty, "un") + "successfully", .Name, Prompt.IconOption.TimedMessage)
        End With
    End Sub
    Private Sub ObjectsTreeview_AutoWidth(sender As Object, e As NodeEventArgs)

        With Tree_Objects
            .AutoWidth()
            ObjectsWidth = 3 + .TotalSize.Width + 3
            TLP_PaneGrid.ColumnStyles(0).Width = ObjectsWidth
        End With

    End Sub
#End Region

#Region " RUN SCRIPT "
    Private Sub RunScript(Optional _Script As Script = Nothing)

        Script_Grid.Columns.CancelWorkers()
        With TT_GridTip
            .ToolTipTitle = Nothing
            .Show(Nothing, Me, 1)
            .Hide(Me)
        End With

        If IsNothing(_Script) Then _Script = ActiveScript
        With _Script
            With .Body
                If .HasText And Not IsNothing(_Script.Connection) Then
                    If _Script.Connection.CanConnect Then
                        If .InstructionType = ExecutionType.DDL Then
#Region " D D L "
                            Cursor.Current = Cursors.WaitCursor
                            Dim procedure As New DDL(.Connection, .SystemText, True, True)
                            If procedure.ProceduresOK.Any Then
                                RaiseEvent Alert(_Script, New AlertEventArgs("Running procedure " & _Script.Name))
                                With procedure
                                    AddHandler .Completed, AddressOf Execute_Completed
                                    .Name = _Script.CreatedString
                                    .Tag = _Script
                                    .Execute(True)
                                End With
                            Else
                                RaiseEvent Alert(Me, New AlertEventArgs("Procedure cancelled"))
                            End If
#End Region

                        ElseIf .InstructionType = ExecutionType.SQL Then
#Region " S Q L "
                            RaiseEvent Alert(_Script, New AlertEventArgs("Running query " & _Script.Name))
                            'https://www.ibm.com/support/knowledgecenter/SSEPEK_11.0.0/cattab/src/tpc/db2z_catalogtablesintro.html
                            With _Script
                                If .Connection.IsFile Then
                                    For Each SheetName In SystemObjects
                                        'SQL_Statement = Replace(SQL_Statement, SheetName.Name, "[" & SheetName.Name & "]")
                                    Next

                                Else
                                    Dim TablesNeed As String() = .Body.TablesNeedObject.ToArray
                                    If TablesNeed.Any Then
                                        RaiseEvent Alert(.Body, New AlertEventArgs("Adding to profile: " & Join(TablesNeed, ",") & "-(RunQuery)"))
                                        Dim TableColumnSQL As String = ColumnSQL(TablesNeed)
                                        With New SQL(.Connection, TableColumnSQL)
                                            AddHandler .Completed, AddressOf ColumnsSQL_Completed
                                            .Execute()
                                        End With
                                    End If
                                    Dim BodyText As String = .Body.SystemText
                                    With New SQL(.Connection, BodyText)
                                        AddHandler .Completed, AddressOf Execute_Completed
                                        .Name = _Script.CreatedString
                                        .Execute()
                                    End With
                                End If
                            End With
#End Region
                        ElseIf .InstructionType = ExecutionType.Null Then

                        End If
                    Else
                        Dim Items As New List(Of String)
                        If _Script.Connection.MissingUserID Then Items.Add("userid")
                        If _Script.Connection.MissingPassword Then Items.Add("password")
                        Message.Show("Can not connect", "Connection is missing " & Join(Items.ToArray, " and "), Prompt.IconOption.Critical, Prompt.StyleOption.Blue)
                    End If
                Else
                    Message.Show("No datasource found or selected", "Please set your connection", Prompt.IconOption.Critical, Prompt.StyleOption.Blue)
                End If
            End With
        End With

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub Execute_OKd(sender As Object, e As ResponseEventArgs)

        With DirectCast(sender, DDL)
            Dim ddlScript As Script = DirectCast(.Tag, Script)
            RaiseEvent Alert(ddlScript, New AlertEventArgs("Running procedure " & ddlScript.Name))
        End With

    End Sub
    Private Sub Execute_Completed(sender As Object, e As ResponseEventArgs)

        Cursor.Current = Cursors.Default

        Dim IsQuery As Boolean = sender.GetType Is GetType(SQL)
        Dim ItemName As String

        If IsQuery Then
            With DirectCast(sender, SQL)
                ItemName = .Name
                RemoveHandler .Completed, AddressOf Execute_Completed
            End With

        Else
            With DirectCast(sender, DDL)
                ItemName = .Name
                RemoveHandler .Completed, AddressOf Execute_Completed
            End With
        End If

        With TT_GridTip
            .Show(String.Empty, ActivePane, 1)
            .Hide(Script_Grid)
            Dim BulletMessage As String = e.Message
            Dim FlatMessage As String
            Dim CreatedDate As Date = StringToDateTime(ItemName)
            Dim _Script As Script = Scripts.Item(CreatedDate)

            .ToolTipTitle = Join({If(IsQuery, "Query", "Procedure"), _Script.Name, If(e.Succeeded, "succeeded", "failed")})

            If e.Succeeded Then
                _Script.Save(Script.SaveAction.UpdateExecutionTime)
                Dim ElapsedMessage As String = Nothing
                With e.ElapsedTime
                    If .Seconds < 1 Then
                        ElapsedMessage = Join({ .Milliseconds, "milliseconds"})

                    ElseIf .Minutes < 1 Then
                        ElapsedMessage = Math.Round(.TotalSeconds, 3) & " seconds"

                    Else
                        ElapsedMessage = Join({ .Minutes, "minutes", .Seconds, "seconds"})

                    End If
                End With
                If IsQuery Then
                    'Show message immediately as it can take time to set datasource, etc
                    BulletMessage = Bulletize({e.Columns.ToString(InvariantCulture) & " columns", e.Rows.ToString(InvariantCulture) & " rows", ElapsedMessage})
                    .Show(BulletMessage, ActivePane, 10 * 1000)

                    Script_Grid.DataSource = e.Table
                    Script_Grid.Refresh()
                    TLP_PaneGrid.ColumnStyles(0).Width = 0
                    AutoWidth(Script_Grid)
                Else
                    BulletMessage = Bulletize({ElapsedMessage})
                    .Show(BulletMessage, ActivePane, 10 * 1000)

                End If

            Else
                BulletMessage = Bulletize({e.Message})
                .Show(BulletMessage, ActivePane, 10 * 1000)

            End If
            FlatMessage = Join(Split(BulletMessage, "● ").Skip(1).ToArray, " ● ")
            RaiseEvent Alert(e, New AlertEventArgs(Join({ .ToolTipTitle, ":", FlatMessage})))
        End With

    End Sub
    Private Sub ColumnsSQL_Completed(sender As Object, e As ResponseEventArgs)

        With DirectCast(sender, SQL)
            RemoveHandler .Completed, AddressOf ColumnsSQL_Completed
            .Table.Namespace = "<Retrieved>"
            Dim ColumnData = DataTableToListOfColumnProperties(.Table)
            If ColumnData.Any Then
                Dim Objects_Retrieved As New List(Of String)
                Dim Columns_Retrieved As New List(Of String)
                Dim Nodes = From CD In ColumnData Group CD By Source = CD.SystemInfo.DSN Into SourceGrp = Group Select New With {
                    .DateSource = Source, .Owners = From SG In SourceGrp Group SG By ObjectOwner = SG.SystemInfo.Owner Into OwnerGrp = Group Select New With {
                    .Owner = ObjectOwner, .Names = From OG In OwnerGrp Group OG By ObjectInfo = OG.SystemInfo Into NameGrp = Group Select New With {
                    .Info = ObjectInfo, .Columns = NameGrp}}}

                For Each _SourceNode In Nodes
                    Dim SourceNode As Node = Tree_Objects.Nodes.Item(Aliases(_SourceNode.DateSource))
                    For Each _OwnerNode In _SourceNode.Owners
                        Dim OwnerNode As Node = SourceNode.Nodes.Item(_OwnerNode.Owner)
                        If OwnerNode Is Nothing Then
                            OwnerNode = SourceNode.Nodes.Add(_OwnerNode.Owner, _OwnerNode.Owner)
                        End If
                        For Each _NameNode In _OwnerNode.Names
                            Dim NameNode As Node = OwnerNode.Nodes.Item(_NameNode.Info.Name)
                            If NameNode Is Nothing Then
                                NameNode = OwnerNode.Nodes.Add(_NameNode.Info.Name, _NameNode.Info.Name, Type_Image(_NameNode.Info.Type))
                            End If
                            NameNode.Tag = _NameNode.Info
                            Objects_Retrieved.Add(_NameNode.Info.ToString)
                            For Each _ColumnNode In _NameNode.Columns
                                Dim ColumnNode As Node = NameNode.Nodes.Item(_ColumnNode.Name)
                                If ColumnNode Is Nothing Then
                                    ColumnNode = NameNode.Nodes.Add(New Node With {.Text = _ColumnNode.Name,
                                                                    .Name = _ColumnNode.Name,
                                                                    .AllowAdd = False,
                                                                    .AllowDragDrop = False,
                                                                    .CheckBox = False})
                                End If
                                Columns_Retrieved.Add(_ColumnNode.ToString)
                            Next
                        Next
                    Next
                Next
                Dim Objects_Saved = PathToList(SystemObjects.Path)
                Dim Objects_New = Objects_Retrieved.Except(Objects_Saved)

                If Objects_New.Any Then
                    For Each ObjectString In Objects_New
                        SystemObjects.Add(New SystemObject(ObjectString))
                    Next
                    SystemObjects.SortCollection()
                    SystemObjects.Save()
                End If

                Dim Columns_Saved = PathToList(Path_Columns)
                Dim Columns_New = Columns_Retrieved.Except(Columns_Saved)
                If Columns_New.Any Then
                    Columns_Saved.AddRange(Columns_New)
                    Columns_Saved.Sort()
                    Using SW As New StreamWriter(Path_Columns)
                        SW.Write(Join(Columns_Saved.ToArray, vbNewLine))
                    End Using
                End If
            End If
        End With

    End Sub
#End Region

#Region " OPEN FILE "
    Private Sub OpenFileClosed(sender As Object, e As EventArgs) Handles OpenFile.FileOk

        Dim _FileType As Extensions = GetFileNameExtension(OpenFile.FileName).Value
        Dim SQL_Statement As String = String.Empty
        If _FileType = Extensions.Excel Then
            Dim Sheets As New List(Of String)(ExcelSheetNames(OpenFile.FileName))
            If Sheets.Count = 1 Then
                SQL_Statement = "Select * FROM [" & Sheets.First & "]"
            Else
                CreateSheetList(Sheets)
                Exit Sub
            End If

        ElseIf _FileType = Extensions.Text Then
            SQL_Statement = "Select * FROM [" & Split(OpenFile.FileName, "\").Last & "]"

        End If
        If IsNothing(OpenFile.Tag) Then
            Script_Grid.DataSource = RetrieveData(OpenFile.FileName, SQL_Statement)
        Else
            OpenFile.Tag = Nothing
        End If

    End Sub
    Private Sub CreateSheetList(Sheets As List(Of String))

        With CMS_ExcelSheets
            .AutoClose = False
            .Items.Clear()
            .BringToFront()
            For Each Sheet In Sheets
                Dim SheetOption As ToolStripItem = .Items.Add(Sheet, My.Resources.Table, AddressOf SheetSelected)
            Next
            .Show(CenterItem(CMS_ExcelSheets.Size))
            .Focus()
        End With

    End Sub
    Private Sub SheetSelected(sender As Object, e As EventArgs)

        CMS_ExcelSheets.AutoClose = True
        CMS_ExcelSheets.Hide()
        Dim Sheet As String = DirectCast(sender, ToolStripItem).Text
        Dim SQL_Statement As String = "Select * FROM [" & Sheet & "]"
        Dim FileTable = RetrieveData(OpenFile.FileName, SQL_Statement)
        If IsNothing(OpenFile.Tag) Then
        Else
            OpenFile.Tag = Nothing
        End If

    End Sub
#End Region

#Region " EXPORT "
    Private Sub ExportOptions_Opening(sender As Object, e As MouseEventArgs) Handles Script_Grid.MouseClick

        Dim canExport As Boolean = Script_Grid.Table IsNot Nothing AndAlso Script_Grid.Table.AsEnumerable.Any
        If canExport And e.Button = MouseButtons.Right Then
            CMS_GridOptions = New ContextMenuStrip With {.Name = "Options",
                .Text = "Options".ToString(InvariantCulture),
                .Font = GothicFont}
            With CMS_GridOptions
                AddHandler .Closed, AddressOf ExportOptions_Closing
                .Items.AddRange({Grid_FileExport, Grid_DatabaseExport})
                .Show(Cursor.Position)
            End With
        End If

    End Sub
    Private Sub ExportOptions_Closing(sender As Object, e As EventArgs)
        RemoveHandler CMS_GridOptions.Closed, AddressOf ExportOptions_Closing
        CMS_GridOptions = Nothing
    End Sub
    Private Sub ExportConnection_Opening(sender As Object, e As EventArgs)

        Dim tsmi As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        With tsmi
            Dim tsch As ToolStripControlHost = DirectCast(.DropDownItems(0), ToolStripControlHost)
            With tsch
                Dim tlp As TableLayoutPanel = DirectCast(.Control, TableLayoutPanel)
                With tlp
                    .BackColor = SystemColors.Control
                    With DirectCast(.Controls("tableName"), ImageCombo)
                        .Image = Nothing
                        .Text = Nothing
                        .Enabled = True
                        .ForeColor = Color.Black
                        .BackColor = Color.GhostWhite
                        .DataSource = Nothing
                    End With
                    Dim clearTable As CheckBox = DirectCast(.Controls("clearTable"), CheckBox)
                    With clearTable
                        .Enabled = True
                        .Checked = True
                        .BackColor = SystemColors.Control
                    End With
                    Dim tableSpace As ImageCombo = DirectCast(.Controls("tableSpace"), ImageCombo)
                    With tableSpace
                        .Image = Nothing
                        .Text = Nothing
                        .DataSource = Nothing
                    End With
                    With .Controls
                        .Remove(clearTable)
                        .Remove(tableSpace)
                        .Add(clearTable, 0, 1)
                        .Add(tableSpace, 0, 2)
                    End With
                End With
            End With
        End With

    End Sub
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ T O   D A T A B A S E
    Private Sub ExportConnection_Enter(sender As Object, e As EventArgs)

        Dim exportCombo As ImageCombo = DirectCast(sender, ImageCombo)
        If exportCombo.Enabled Then
            Dim exportConnection As Connection = DirectCast(exportCombo.Tag, Connection)
            Dim gridColumns = Script_Grid.Columns.Names
            Dim Names = ValuesToFields(gridColumns.Keys.ToArray)
            Dim sqlTablename As String = "SELECT TBNAME FROM (SELECT TBNAME, COUNT(*) X
FROM SYSIBM.SYSCOLUMNS
WHERE NAME In " & Names & "
GROUP BY TBNAME) COLUMNS
WHERE CAST(X AS SMALLINT)=" & gridColumns.Count
            Dim sqlExport = New SQL(exportConnection, sqlTablename)
            With sqlExport
                .Execute(False)
                If .Status = TriState.True Then
                    Dim results = From r In .Table.AsEnumerable Select CStr(r("TBNAME"))
                    If results.Any Then
                        exportCombo.DataSource = results
                        exportCombo.SelectedIndex = 0
                        exportCombo.Image = My.Resources.Check.ToBitmap
                        exportCombo.Image.Tag = Join(results.ToArray, BlackOut)
                    Else
                        exportCombo.Image = My.Resources.Info
                    End If
                Else
                    RaiseEvent Alert(Me, New AlertEventArgs(.Response.Message))
                    exportCombo.Text = .Response.Message
                    exportCombo.ForeColor = Color.DarkRed
                    exportCombo.BackColor = Color.Gainsboro
                    exportCombo.Enabled = False
                    Dim tlpConnection As TableLayoutPanel = DirectCast(exportCombo.Parent, TableLayoutPanel)
                    tlpConnection.BackColor = Color.Gainsboro
                    Dim exportCheckbox As CheckBox = DirectCast(tlpConnection.Controls(1), CheckBox)
                    exportCheckbox.CheckState = CheckState.Indeterminate
                    exportCheckbox.BackColor = Color.Gainsboro
                    exportCheckbox.Enabled = False
                End If
            End With
        End If

    End Sub
    Private Sub ExportConnection_Submitted(sender As Object, e As ImageComboEventArgs)

        Dim tableName As ImageCombo = DirectCast(sender, ImageCombo)
        Dim tlpConnection As TableLayoutPanel = DirectCast(tableName.Parent, TableLayoutPanel)
        Dim exportConnection As Connection = DirectCast(tableName.Tag, Connection)
        Dim clearTable As CheckBox = DirectCast(tlpConnection.Controls("clearTable"), CheckBox)
        Dim tableSpace As ImageCombo = DirectCast(tlpConnection.Controls("tableSpace"), ImageCombo)

        If tableSpace.Image Is Nothing Then
            If tableName.Text?.Any Then
                Dim matchingNames As New List(Of String)
                If SameImage(My.Resources.Check.ToBitmap, tableName.Image) Then matchingNames.AddRange(Split(tableName.Image?.Tag?.ToString, BlackOut))
                Dim foundTablename As Boolean = matchingNames.Contains(tableName.Text)
                If Not foundTablename Then 'Results from MouseOver SQL did not return any results ... probably new table but maybe not. Check if TableName exists
                    Dim validTablename As String = DB2TableNamingConvention(tableName.Text)
                    If tableName.Text = validTablename Then 'DB2 would accept the submitted name ... now check if it exists ( Insert into existing Or Create new ) 
                        Dim Instruction As String = "WITH SPACES (SPACE) As (Select
                DISTINCT TRIM(DBNAME)||'.'||TRIM(TSNAME) SPACE
                FROM SYSIBM.SYSTABLES T
                WHERE T.CREATOR='" & exportConnection.UserID & "'
                AND TYPE='T')
                , TABLES (SPACE, COUNT) AS (SELECT SPACE
                , (SELECT COUNT(*)
                FROM SYSIBM.SYSTABLES TT
                WHERE TT.CREATOR='" & exportConnection.UserID & "' AND TT.NAME='" & Trim(tableName.Text?.ToUpperInvariant) & "' AND S.SPACE=TRIM(DBNAME)||'.'||TRIM(TSNAME)) COUNT
                FROM SPACES S)
                SELECT *
                FROM TABLES"
                        Dim tableSQL As New SQL(exportConnection, Instruction)
                        With tableSQL
                            .Execute(False)
                            If .Status = TriState.True Then
                                Dim Spaces As New Dictionary(Of String, Integer)(.Table.AsEnumerable.ToDictionary(Function(x) x("SPACE").ToString, Function(y) DirectCast(y("COUNT"), Integer)))
                                If Spaces.Values.Sum = 0 Then
#Region " CREATE NEW TABLE - CHECK # OF TABLESPACES WHERE THE TABLE IS TO BE CREATED "
                                    With tlpConnection.Controls
                                        .Remove(clearTable)
                                        .Remove(tableSpace)
                                        .Add(tableSpace, 0, 1)
                                        .Add(clearTable, 0, 2)
                                        With tableSpace
                                            .DataSource = Spaces.Keys
                                            .SelectedIndex = 0
                                            .IsReadOnly = True
                                            .Image = If(Spaces.Count = 1, My.Resources.Check.ToBitmap, My.Resources.Info)
                                        End With
                                        If Spaces.Count = 1 Then
                                            REM /// NO NEED FOR USER INPUT. BEGIN EXPORT INTO NEW TABLE
                                            Export_CreateTable(exportConnection, Spaces.First.Key, tableName.Text)
                                            tableSpace.Image = My.Resources.Check.ToBitmap
                                        End If
                                    End With
#End Region
                                Else
                                    REM /// TABLE FOUND ///. REQUIRES USER TO PICK AN ACTION { Clear Or Add to existing rows }
                                    foundTablename = True

                                End If
                            End If
                        End With
                    Else
                        Using message As New Prompt With {.MinimumSize = New Size(600, 300)}
                            Dim validConvention As String = "A Table name must satisfy all below conditions:" & Bulletize({
                               "Length can not exceed 18 characters",
                               "Begin with a letter or one of $, #, @",
                               "Can contain: letters A-Z, any valid letter with an accent, digits 0 through 9, _, $, #, @",
                               "A name cannot be a DB2 Or an SQL reserved word, such as WHERE Or VIEW",
                               "nb) A name enclosed in quotes will be case-sensitive"})
                            message.Show("Invalid Table name", validConvention, Prompt.IconOption.Critical, Prompt.StyleOption.Blue)
                        End Using
                        Exit Sub
                    End If
                End If
                If foundTablename Then
                    With New ETL()
                        .Sources.Add(New ETL.Source(Script_Grid.Table))
                        .Destinations.Add(New ETL.Destination(exportConnection, tableName.Text) With {.ClearTable = clearTable.Checked})
                        AddHandler .Completed, AddressOf ViewerTableExportedToDatabase
                        .Execute()
                    End With
                End If
            End If
        Else
            'User hit Enter after TableSpaces were checked - need to perform above
            Export_CreateTable(exportConnection, tableSpace.Text, tableName.Text)

        End If

    End Sub
    Private Sub Export_CreateTable(exportConnection As Connection, tableSpace As String, tableName As String)

        Dim exportETL As New ETL
        With exportETL
            .Sources.Add(New ETL.Source(Script_Grid.Table))
            .Destinations.Add(New ETL.Destination(exportConnection, tableSpace, tableName))
            AddHandler .Completed, AddressOf ViewerTableExportedToDatabase
            .Name = Join({"Exporting",
                                            Script_Grid.Table.Columns.Count,
                                            "columns and ",
                                            Script_Grid.Table.Rows.Count,
                                            "rows to into a new table in",
                                            exportConnection.DataSource,
                                            ", named", tableName})
            RaiseEvent Alert(exportETL, New AlertEventArgs(.Name))
            .Execute()
        End With

    End Sub
    Private Sub ViewerTableExportedToDatabase(sender As Object, e As ResponsesEventArgs)

        With DirectCast(sender, ETL)
            RemoveHandler .Completed, AddressOf ViewerTableExportedToDatabase
            RaiseEvent Alert(sender, New AlertEventArgs(If(.Succeeded, "Succeeded ", "Failed ") & Replace(.Name, "Exporting", "exporting")))
        End With

    End Sub
    Private Sub TableSpaces(sender As Object, e As ResponseEventArgs)

        Dim Spaces As New Dictionary(Of String, List(Of SystemObject))
        With DirectCast(sender, SQL)
            RemoveHandler .Completed, AddressOf TableSpaces
            Dim Objects As New SystemObjectCollection(.Table)
            For Each TableObject In Objects
                Dim Space As String = TableObject.TSName
                If Not Spaces.ContainsKey(Space) Then Spaces.Add(Space, New List(Of SystemObject))
                Spaces(Space).Add(TableObject)
            Next
        End With

    End Sub

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ T O   F I L E
    Private Sub ExportToFile(sender As Object, e As EventArgs)

        REM /// EVENTUALLY, MOVE DATA BYPASSING SQL_VIEW (QUERY TO DB)
        Dim SourceTable As DataTable = DirectCast(Script_Grid.DataSource, DataTable)
        If SourceTable Is Nothing Then

        Else
            Dim FileName As String = Split(SourceTable.TableName, ".").Last
            Dim ExportObject As ToolStripDropDownItem = DirectCast(sender, ToolStripDropDownItem)
            Select Case ExportObject.Text
                Case "Excel", "+ Query"
                    SaveFile.FileName = Join({Desktop, "\", FileName, ".xlsx"}, String.Empty)
                    SaveFile.Filter = "Excel Files|*.xls,*.xlsx".ToString(InvariantCulture)
                    SaveFile.Title = ExportObject.Text
                    SaveFile.ShowDialog()

                Case ".csv"
                    SaveFile.FileName = Join({Desktop, "\", FileName, ".csv"}, String.Empty)
                    SaveFile.Filter = "CSV|*.csv".ToString(InvariantCulture)
                    SaveFile.ShowDialog()

                Case ".txt"
                    SaveFile.FileName = Join({Desktop, "\", FileName, ".txt"}, String.Empty)
                    SaveFile.Filter = "TXT Files (*.txt*)|*.txt".ToString(InvariantCulture)
                    SaveFile.ShowDialog()

                Case Else
                    REM /// DATABASE. NEED NAME, TABLESPACE, NEW(CHECK), CLEAR OR ADD
                    REM /// ALL EXPORTS TO DSN REQUIRE A NAME
                    REM /// NEW PARAMETER=(OWNER+TABLE.NAME) EXISTS
                    REM /// IF NEW, THEN TABLESPACE MUST BE OBTAINED (COULD BE MULTIPLE)
                    REM /// IF NOT NEW, PARAMETER FOR CLEAR Or INSERT AS RADIO BUTTON
                    Dim QueryTable As DataTable = TryCast(Script_Grid.DataSource, DataTable)
                    If Not IsNothing(QueryTable) AndAlso QueryTable.AsEnumerable.Any Then

                    End If

            End Select
        End If

    End Sub
    Private Sub SaveFileClosed(sender As Object, e As EventArgs) Handles SaveFile.FileOk

        If IsNothing(SaveFile.Tag) Then
            REM /// Saving structured table
            Dim QueryTable As DataTable = TryCast(Script_Grid.DataSource, DataTable)
            If Not IsNothing(QueryTable) AndAlso QueryTable.AsEnumerable.Any Then

                Select Case GetFileNameExtension(SaveFile.FileName).Value
                    Case Extensions.Excel
                        ' DOES NOT WORK !!! USER MUST RUN EXCEL AS ADMINISTRATOR IN Windows10 + ConnectionString/SQL=String.Empty
                        'Dim ConnectionString As String = String.Empty
                        'Dim SQL As String = String.Empty
                        'Dim WithQuery As Boolean = SaveFile.Title = "+ Query"
                        DataTableToExcel(QueryTable, SaveFile.FileName, True, False, False, True, True)

                    Case Extensions.Text
                        DataTableToTextFile(QueryTable, SaveFile.FileName)

                    Case Extensions.CommaSeparated
                    Case Extensions.SQL

                End Select
            Else
                MessageBox.Show("Nothing to Export".ToString(InvariantCulture))
            End If

        Else
            REM /// Saving SQL or DDL

        End If
        SaveFile.Tag = Nothing

    End Sub
#End Region

    Private Sub Scripts_Changed(sender As Object, e As ScriptsEventArgs) Handles Scripts_.CollectionChanged

        If e.State = CollectionChangeAction.Refresh Then
            'RemoveHandler ActivePane_.SelectionChanged, AddressOf ActivePane_SelectionChanged
            'AddHandler ActivePane_.SelectionChanged, AddressOf ActivePane_SelectionChanged
            RemoveHandler Tree_ClosedScripts.SizeChanged, AddressOf ClosedScripts_SizeChanged
            AddHandler Tree_ClosedScripts.SizeChanged, AddressOf ClosedScripts_SizeChanged
            For Each Script In Scripts
                ScriptToNode(Script)
            Next
            With Tree_ClosedScripts
                .Nodes.SortOrder = SortOrder.Ascending
                .Nodes.Insert(0, OpenFileNode)
            End With
            ScriptsInitialized = True

        ElseIf e.State = CollectionChangeAction.Add Then
            If ScriptsInitialized Then ScriptToNode(e.Item)

        ElseIf e.State = CollectionChangeAction.Remove Then
            Dim RemoveNode As Node = Tree_ClosedScripts.Nodes.ItemByTag(e.Item)
            If RemoveNode IsNot Nothing Then
                RemoveNode.Parent.Nodes.Remove(RemoveNode)
            End If
            RemoveHandler e.Item.StateChanged, AddressOf Script_StateChanged
            RemoveHandler e.Item.NameChanged, AddressOf Script_NameChanged

        End If

    End Sub
    Private Sub ScriptToNode(Item As Script)

        AddHandler Item.StateChanged, AddressOf Script_StateChanged
        AddHandler Item.NameChanged, AddressOf Script_NameChanged
        With Tree_ClosedScripts
            Dim ConnectionName As String = If(Item.Connection Is Nothing, "Undetermined", Item.Connection.DataSource)
            Dim DatabaseColor As Color = If(Item.Connection Is Nothing, Color.Blue, Item.Connection.BackColor)
            Dim Database_Image As Image = ChangeImageColor(My.Resources.Sync, Color.FromArgb(255, 64, 64, 64), DatabaseColor)

            If Not .Nodes.Exists(Function(n) n.Name = ConnectionName) Then
                .Nodes.Add(New Node With {
                            .Text = ConnectionName,
                            .Name = ConnectionName,
                            .Image = Database_Image,
                            .AllowAdd = False,
                            .AllowDragDrop = False,
                            .AllowEdit = False,
                            .AllowRemove = False,
                            .Separator = Node.SeparatorPosition.Above,
                            .Tag = Item.Connection})
            End If
            Dim ConnectionNode As Node = .Nodes.Item(ConnectionName)
            If Item.Connection Is Nothing Then ConnectionNode.BackColor = Color.FromArgb(128, Color.Gainsboro)
            ConnectionNode.Nodes.Add(New Node With {.Text = Item.Name,
                                                    .Name = Item.Name,
                                                    .Image = If(Item.Body.InstructionType = ExecutionType.DDL, My.Resources.DDL, My.Resources.SQL),
                                                    .AllowAdd = False,
                                                    .AllowEdit = True,
                                                    .AllowRemove = True,
                                                    .Tag = Item})
            ConnectionNode.Nodes.SortOrder = SortOrder.Ascending
        End With

    End Sub
    Private Sub Script_StateChanged(sender As Object, e As ScriptStateChangedEventArgs)

        Dim ScriptItem As Script = DirectCast(sender, Script)
        Dim ScriptNode As Node = Tree_ClosedScripts.Nodes.ItemByTag(ScriptItem)

        If e.FormerState = Script.ViewState.OpenSaved And (e.CurrentState = Script.ViewState.ClosedSaved Or e.CurrentState = Script.ViewState.ClosedNotSaved) Then
            'Closing
            ScriptNode.Image = If(ScriptItem.Body.IsDDL, My.Resources.DDL, My.Resources.SQL)

        ElseIf e.CurrentState = Script.ViewState.OpenSaved And (e.FormerState = Script.ViewState.ClosedSaved Or e.FormerState = Script.ViewState.ClosedNotSaved) Then
            'Opening
            Dim NodeColor As Color = If(ScriptItem.Connection Is Nothing, Color.Black, ScriptItem.Connection.BackColor)
            Dim bmp As Bitmap = My.Resources.Eye
            If bmp IsNot Nothing Then
                Using g As Graphics = Graphics.FromImage(bmp)
                    Using Attributes As Imaging.ImageAttributes = New Imaging.ImageAttributes()
                        Dim rect As Rectangle = New Rectangle(0, 0, bmp.Width, bmp.Height)
                        g.DrawImage(bmp, rect, 0, 0, rect.Width, rect.Height, GraphicsUnit.Pixel, Attributes)
                        Using SolidLine As New Pen(Color.FromArgb(192, NodeColor), 3)
                            g.DrawLine(SolidLine, New Point(rect.Left + 1, rect.Bottom - 3), New Point(rect.Right - 1, rect.Bottom - 3))
                        End Using
                    End Using
                End Using
            End If
            ScriptNode.Image = bmp

        ElseIf e.FormerState = Script.ViewState.OpenDraft And e.CurrentState = Script.ViewState.None Then
            Scripts.Remove(ScriptItem)

        End If

    End Sub
    Private Sub Script_NameChanged(sender As Object, e As ScriptNameChangedEventArgs)
        Tree_ClosedScripts.Nodes.ItemByTag(sender).Text = e.CurrentName
    End Sub

    Private Sub TSMI_FontClicked() Handles TSMI_Font.Click

        With CMS_PaneOptions
            .AutoClose = True
            .Hide()
        End With
        FindAndReplace.Close()
        With Dialogue_Font
            .ShowApply = True
            .ShowColor = True
            .ShowEffects = True
            .ShowDialog(ActivePane)
        End With

    End Sub
    Private Sub Dialogue_FontApply(sender As Object, e As EventArgs) Handles Dialogue_Font.Apply

        With Dialogue_Font
            If ActivePane.SelectionLength = 0 Then
                ActivePane.Font = .Font
                ActivePane.ForeColor = .Color
            Else
                ActivePane.SelectionColor = .Color
                ActivePane.SelectionFont = .Font
            End If

            My.Settings.Font_Pane = .Font
            My.Settings.Save()
        End With

    End Sub

#Region " FUN-CTIONS "
    Private Function CreateTableText(InputString As String) As String

        Dim Locations As New List(Of String)
        Dim Lines As New List(Of String)(Split(InputString, vbNewLine))
        Dim NewLines As New List(Of String)
        Dim Items As New Dictionary(Of Integer, Integer)
        Dim MaxWidth As Integer = 0
        Using RTB As New RichTextBox With {.Font = My.Settings.Font_Pane, .Width = 2000, .Text = InputString}
            With RTB
                Dim TabWidths As New List(Of Integer)
                For i = 0 To 10
                    .Text = StrDup(i, vbTab) & "."
                    TabWidths.Add(.GetPositionFromCharIndex(i).X)
                Next
                Dim TabWidth As Integer = Convert.ToInt32((TabWidths.Last - TabWidths.First) / (TabWidths.Count - 1))

                For Each Line In Lines
                    Dim ColumnLine As Match = Regex.Match(Line, "^[^\t]{1,}(?=\t)", RegexOptions.IgnoreCase)
                    Dim ColumnValue = ColumnLine.Value
                    If ColumnLine.Success Then
                        Dim TabString = ColumnValue & StrDup(4, vbTab) & "X"
                        .Text = TabString
                        Dim PeriodIndex = TabString.Length - 1
                        Dim TabLocation = .GetPositionFromCharIndex(PeriodIndex)
                        Items.Add(Items.Count, TabLocation.X)
                        If TabLocation.X > MaxWidth Then MaxWidth = TabLocation.X
                    Else
                        Items.Add(Items.Count, -1)
                    End If
                Next
                Dim LineWidths As New Dictionary(Of String, List(Of Integer))
                For Each Item In Items
                    Dim ColumnText As String = Lines(Item.Key)
                    If Item.Value < 0 Then
                        NewLines.Add(ColumnText)
                    Else
                        Dim Values = Split(ColumnText, vbTab)
                        Dim TabCount As Integer = 0
                        Dim LineTabWidth As Integer = 0
                        Dim ColumnName As String = Values.First
                        Dim ColumnFormat As String = Values.Last
                        Dim Widths As New List(Of Integer)
                        Do
                            TabCount += 1
                            Dim LineText As String = Join({ColumnName, StrDup(TabCount, vbTab), ColumnFormat}, String.Empty)
                            .Text = LineText
                            Dim TabRight As Integer = ColumnName.Length + TabCount
                            Dim CF As String = LineText.Substring(TabRight, 1)
                            LineTabWidth = .GetPositionFromCharIndex(TabRight).X
                            Widths.Add(LineTabWidth)
                        Loop While LineTabWidth < MaxWidth
                        NewLines.Add(Join({ColumnName, StrDup(TabCount, vbTab), ColumnFormat}, String.Empty))
                        LineWidths.Add(ColumnName, Widths)
                    End If
                Next
            End With
        End Using
        Return Join(NewLines.ToArray, vbNewLine)

    End Function
    Private Sub AutoWidth(sender As Object, Optional e As Object = Nothing) Handles Script_Grid.ColumnsSized

        If TLP_PaneGrid.ColumnStyles.Count >= 2 Then

            Dim Column_1_Width As Integer = 0
            Dim Column_2_Width As Integer = 0

            Dim Column1Percent As Integer = 50
            Dim ColumnToPercent As Integer = 100 - Column1Percent

            Dim AutoSize As Boolean = False

            If ActivePane Is Nothing Then
                'Only DummyTab showing...Leave Column1Percent @ 50
            Else
                If Not ActivePane.HasText And Script_Grid.DataSource Is Nothing Then
                    'Both Pane and Grid are empty...Leave Column1Percent @ 50
                Else
                    AutoSize = True
                    'Calculate best fit based on what looks best
#Region " RULE 1 - Column1 (Tabs+ActivePane).Width can't be less < Right side of last tab otherwise control wraps and looks like @#$% "
                    Dim TabBounds As New List(Of Rectangle)
                    For TabIndex As Integer = 0 To Script_Tabs.TabPages.Count - 1
                        Dim TabHeaderBounds As Rectangle = Script_Tabs.GetTabRect(TabIndex)
                        TabBounds.Add(TabHeaderBounds)
                    Next
                    Dim Tabs_BestWidth As Integer = TabBounds.Sum(Function(t) t.Width) + 32
#End Region
#Region " RULE 2 - Avoid setting Column1 (Tabs+ActivePane).Width less than ActivePane.IdealWidth as it looks best when text is not wrapped "
                    Dim Pane_BestWidth As Integer = ActivePane.IdealWidth(True)
#End Region
#Region " RULE 3 - Allow the Grid.Width as much space as possible without affecting above 2 Rules "
                    Dim Grid_BestWidth As Integer = Script_Grid.TotalSize.Width
#End Region
                    'Sender is one of 3 Controls {Script_Grid.RetrievedData, ActiveTab.TextChanged, ActivePane.DroppedText}

                    Dim Column_1_MinimumWidth As Integer = {Tabs_BestWidth, Pane_BestWidth}.Max
                    Dim Column_2_IdealWidth As Integer = Grid_BestWidth + PaneGridSeparator.Width

                    Dim Actual_Column0_Width As Integer = Convert.ToInt32(TLP_PaneGrid.ColumnStyles(0).Width)
                    Dim Actual_Column1_Width As Integer = 0
                    Dim Actual_Column2_Width As Integer = Script_Grid.Width

                    Dim AvailableWidth As Integer = TLP.GetContentSpace(TLP_PaneGrid) - Actual_Column0_Width

                    If TLP_PaneGrid.ColumnStyles(1).SizeType = SizeType.Absolute Then
                        Actual_Column1_Width = Convert.ToInt32(TLP_PaneGrid.ColumnStyles(1).Width)
                    Else
                        Actual_Column1_Width = AvailableWidth - Actual_Column2_Width
                    End If

                    If sender Is Script_Grid Then
                        'Potentially chew up Tab+Pane.Width
                        'Close Column 0
                        If {Column_1_MinimumWidth, Column_2_IdealWidth}.Sum > AvailableWidth Then
                            'Doesn't all fit so give it the remaining available
                            Column_1_Width = Column_1_MinimumWidth
                            Column_2_Width = AvailableWidth - Column_1_Width

                        Else
                            Column_2_Width = Column_2_IdealWidth
                            Column_1_Width = AvailableWidth - Column_2_Width

                        End If

                    ElseIf sender Is ActivePane Then
                        Column_1_Width = Column_1_MinimumWidth
                        Column_2_Width = AvailableWidth - Column_1_Width

                    ElseIf sender.GetType Is GetType(Tab) Then
                        If Actual_Column1_Width < Column_1_MinimumWidth Then
                            Column_1_Width = Column_1_MinimumWidth
                        Else
                            Column_1_Width = Actual_Column1_Width
                        End If
                        Column_2_Width = AvailableWidth - Column_1_Width

                    Else
                        Stop
                    End If
                    Column1Percent = Convert.ToInt32(100 * Column_1_Width / AvailableWidth)
                End If
            End If

            TLP_PaneGrid.ColumnStyles(1).SizeType = SizeType.Percent
            TLP_PaneGrid.ColumnStyles(2).SizeType = SizeType.Percent

            Column1Percent = {{Column1Percent, 25}.Max, 75}.Min
            ColumnToPercent = 100 - Column1Percent

            TLP_PaneGrid.ColumnStyles(1).Width = Column1Percent
            TLP_PaneGrid.ColumnStyles(2).Width = ColumnToPercent

        End If

    End Sub
#End Region
End Class