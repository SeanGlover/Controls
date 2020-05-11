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
    Public Sub New(item As Script, state As CollectionChangeAction)
        Me.Item = item
        Me.State = state
    End Sub
End Class
Public Enum ExecutionType
    Null
    DDL
    SQL
End Enum
Public Class ScriptTypeChangedEventArgs
    Inherits EventArgs
    Public ReadOnly Property FormerType As ExecutionType
    Public ReadOnly Property CurrentType As ExecutionType
    Public Sub New(FormerType As ExecutionType, CurrentType As ExecutionType)
        Me.FormerType = FormerType
        Me.CurrentType = CurrentType
    End Sub
End Class
Public Class ScriptStateChangedEventArgs
    Inherits EventArgs
    Public ReadOnly Property FormerState As Script.ViewState
    Public ReadOnly Property CurrentState As Script.ViewState
    Public Sub New(FormerState As Script.ViewState, CurrentState As Script.ViewState)
        Me.FormerState = FormerState
        Me.CurrentState = CurrentState
    End Sub
End Class
Public Class ScriptNameChangedEventArgs
    Inherits EventArgs
    Public ReadOnly Property FormerName As String
    Public ReadOnly Property CurrentName As String
    Public Sub New(FormerName As String, CurrentName As String)
        Me.FormerName = FormerName
        Me.CurrentName = CurrentName
    End Sub
End Class
Public Class ScriptVisibleChangedEventArgs
    Inherits EventArgs
    Public ReadOnly Property Item As Script
    Public ReadOnly Property Visible As Boolean
    Public Sub New(item As Script, visible As Boolean)
        Me.Item = item
        Me.Visible = visible
    End Sub
End Class
<ComVisible(False)> Public Class ScriptCollection
    Inherits List(Of Script)
    Private ReadOnly Scripts_DirectoryInfo As DirectoryInfo = Directory.CreateDirectory(MyDocuments & "\DataManager\Scripts")
    Public Event Alert(sender As Object, e As AlertEventArgs)
    Public Event CollectionChanged(sender As Object, e As ScriptsEventArgs)
    Public Sub New()
    End Sub
    Public Sub Load()

        Dim fileScripts As New List(Of String)(GetFiles(Scripts_DirectoryInfo.FullName, ".ddl").Union(GetFiles(Scripts_DirectoryInfo.FullName, ".sql")).Union(GetFiles(Scripts_DirectoryInfo.FullName, ".txt")))
        For Each fileScript In fileScripts
            Add(New Script(fileScript))
        Next
        SortCollection()
        RaiseEvent Alert(Me, New AlertEventArgs(Count & " scripts loaded"))
        RaiseEvent CollectionChanged(Me, New ScriptsEventArgs(Nothing, CollectionChangeAction.Refresh))

    End Sub
#Region " PROPERTIES - FUNCTIONS - METHODS "
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Shadows Function Add(Item As Script) As Script

        If Item IsNot Nothing Then
            Item.Parent = Me
            MyBase.Add(Item)
            RaiseEvent CollectionChanged(Me, New ScriptsEventArgs(Item, CollectionChangeAction.Add))
        End If
        Return Item

    End Function
    Public Shadows Function Remove(Item As Script) As Script

        If Item IsNot Nothing Then
            MyBase.Remove(Item)
            Item.Parent = Nothing
            SortCollection()
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
    Public Overrides Function ToString() As String
        Dim savedScripts As New List(Of String)(From m In Me Where m.FileCreated Select m.ToString & String.Empty)
        Return Strings.Join(savedScripts.ToArray, vbNewLine)
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

        End If
        disposed = True
    End Sub
#End Region
#Region " EVENTS "
    Friend Event GenericEvent(sender As Object, e As AlertEventArgs)
    Friend Event ConnectionChanged(sender As Object, e As ConnectionChangedEventArgs)
    Friend Event TypeChanged(sender As Object, e As ScriptTypeChangedEventArgs)
    Friend Event StateChanged(sender As Object, e As ScriptStateChangedEventArgs)
    Friend Event NameChanged(sender As Object, e As ScriptNameChangedEventArgs)
    Friend Event VisibleChanged(sender As Object, e As ScriptVisibleChangedEventArgs)
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
        ChangeFavorite
        ChangeContent
        UpdateExecutionTime
    End Enum
#End Region
    Private Const BodyDelimiter As String = vbNewLine + "■■■■■■■■■■■■■■■■■■■■" + vbNewLine
    Private Sub Body_TypeChanged(sender As Object, e As ScriptTypeChangedEventArgs) Handles Body.TypeChanged
        _Type = If(e.CurrentType = Controls.ExecutionType.DDL, ExecutionType.DDL, If(e.CurrentType = Controls.ExecutionType.SQL, ExecutionType.SQL, ExecutionType.Null))
        RaiseEvent TypeChanged(Me, e)
    End Sub
    Private Sub Body_ConnectionChanged(sender As Object, e As ConnectionChangedEventArgs) Handles Body.ConnectionChanged
        If Connection Is Nothing And e.NewConnection IsNot Nothing Then Connection = e.NewConnection 'Only change the active connection once (initially). Let the User override to pick a connection if needed
        RaiseEvent ConnectionChanged(Me, e)
    End Sub
#Region " NEW "
    Public Sub New()
        Created = Now
    End Sub 'New Instance
    Public Sub New(ScriptPath As String) 'From Saved

        _Path = If(ScriptPath, String.Empty)
        _State = ViewState.ClosedSaved
        Dim fileName As String = GetFileNameExtension(ScriptPath).Key
        _Name = Replace(fileName, "♥", String.Empty)
        _Created = File.GetCreationTime(ScriptPath)
        Dim dsnBody As String()
        Using SR As New StreamReader(Path)
            dsnBody = Split(SR.ReadToEnd, BodyDelimiter)
        End Using
        DSN_Body = New KeyValuePair(Of String, String)("DSN=" & dsnBody.First, Regex.Replace(dsnBody.Last, vbCrLf, vbLf, RegexOptions.None))
        Dim dsn = DSN_Body.Key
        _Connection = DataTool.Connections.Item(dsn)
        _Type = If(_Path.EndsWith(".ddl", StringComparison.InvariantCulture), ExecutionType.DDL, If(_Path.EndsWith(".sql", StringComparison.InvariantCulture), ExecutionType.SQL, ExecutionType.Null))
        Favorite_ = fileName.StartsWith("♥", StringComparison.CurrentCulture)
        _Text = DSN_Body.Value

    End Sub
#End Region
#Region " PROPERTIES - FUNCTIONS - METHODS "
    <NonSerialized> Friend Parent As ScriptCollection
    <NonSerialized> Public WithEvents Body As New BodyElements
    <NonSerialized> Friend TabPage_ As Tab
    Friend ReadOnly Property TabPage As Tab
        Get
            Return TabPage_
        End Get
    End Property
    Public ReadOnly Property TabImage As Image
        Get
            Return If(Type = ExecutionType.DDL, My.Resources.DDL, If(Type = ExecutionType.SQL, My.Resources.SQL, My.Resources.QuestionMark))
        End Get
    End Property
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
                Save(SaveAction.ChangeContent)
                RaiseEvent ConnectionChanged(Me, New ConnectionChangedEventArgs(FormerValue, value))
            End If
        End Set
    End Property
    Public ReadOnly Property Type As ExecutionType
    Private ReadOnly Property FileExtension As String
        Get
            Return "." & If(Type = ExecutionType.Null, "txt", If(Type = ExecutionType.DDL, "ddl", "sql"))
        End Get
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
    Public ReadOnly Property FileCreated As Boolean
        Get
            Return File.Exists(Path)
        End Get
    End Property
    Friend ReadOnly Property DSN_Body As KeyValuePair(Of String, String)
    Public ReadOnly Property TextWasModified As Boolean
        Get
            Return If(FileCreated, Not FileTextMatchesText, Body.HasText)
        End Get
    End Property
    Public ReadOnly Property FileTextMatchesText As Boolean
        Get
            If FileCreated Then
                Dim fileText As String = DSN_Body.Value
                fileText = Replace(fileText, vbCrLf, "♀")
                fileText = Replace(fileText, vbLf, "♀")
                Dim paneText As String = Text
                paneText = Replace(paneText, vbCrLf, "♀")
                paneText = Replace(paneText, vbLf, "♀")
                Return fileText = paneText
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
                        If _Name.Any Then
                            'FileCreated + _Name = FormerName OR Not FileCreated
                            Dim NewName As String = _Name & FileExtension
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
                        End If
#End Region
                    End If
#End Region
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
    Private Favorite_ As Boolean
    Friend Property Favorite As Boolean
        Get
            Return Favorite_
        End Get
        Set(value As Boolean)
            If value <> Favorite_ Then
                Favorite_ = value
                Save(SaveAction.ChangeFavorite)
            End If
        End Set
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
                        REM /// ( User clicked [+] AddTab )
                        RaiseEvent VisibleChanged(Me, New ScriptVisibleChangedEventArgs(Me, True))

                    Case FormerState = ViewState.None And NewState = ViewState.OpenSaved
                        REM /// NOT LIKELY

                    Case FormerState = ViewState.None And NewState = ViewState.ClosedSaved
                        REM /// NEW FROM MY.SETTINGS.SCRIPTS
#End Region
#Region " From Draft "
                    Case FormerState = ViewState.OpenDraft And NewState = ViewState.None
                        REM /// Discard (User clicked [x] and doesn't want to save work)
                        RaiseEvent VisibleChanged(Me, New ScriptVisibleChangedEventArgs(Me, False))
                        Parent.Remove(Me)

                    Case FormerState = ViewState.OpenDraft And NewState = ViewState.OpenSaved
                        REM /// Handled in Name Set since the only method to change from OpenDraft to OpenSaved is via IC_SaveAs -OR- Tree_ClosedScripts
                        Parent.Add(Me)

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
                        RaiseEvent VisibleChanged(Me, New ScriptVisibleChangedEventArgs(Me, False))

                    Case FormerState = ViewState.OpenSaved And NewState = ViewState.ClosedNotSaved
                        REM /// Discard any Text changes...revert back to FileText
                        Text = DSN_Body.Value
                        _State = ViewState.ClosedSaved
                        RaiseEvent VisibleChanged(Me, New ScriptVisibleChangedEventArgs(Me, False))
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
                        Body.Text = Text
                        RaiseEvent VisibleChanged(Me, New ScriptVisibleChangedEventArgs(Me, True))
#End Region
                End Select
                RaiseEvent StateChanged(Me, New ScriptStateChangedEventArgs(FormerState, NewState))
            End If
        End Set
    End Property
    Private _Text As String
    Friend Property Text() As String
        Get
            Return _Text
        End Get
        Set(value As String)
            If _Text <> value Then
                _Text = value
                Body.Text = value
            End If
        End Set
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
            'vbCrLf (VBNEWLINE) IS USED AS CARRIAGE RETURN IN THE .txt FILE BUT vbLf IS USED IN THE PANE.TEXT ... vbNewLine MAKES THE .txt FILE MUCH MORE READABLE
            Dim ConnectionText As String = If(Connection Is Nothing, String.Empty, Connection.Properties("DSN"))
            Dim ScriptText As String = Regex.Replace(If(Text, String.Empty), vbLf, vbCrLf)

            Select Case Action
                Case SaveAction.ChangeContent
                    Dim splitPath As String() = Split(Path, ".")
                    Dim extensions As New List(Of String) From {"txt", "sql", "ddl"}
                    extensions.Remove(splitPath.Last)
                    For Each otherExtension In extensions
                        Dim otherPath As String = Join({splitPath.First, otherExtension}, ".")
                        If File.Exists(otherPath) Then
                            Stop
                            File.Delete(otherPath)
                        End If
                    Next
                    Using SW As New StreamWriter(Path)
                        SW.Write(Join({ConnectionText, ScriptText}, BodyDelimiter))
                    End Using
                    File.SetLastWriteTime(Path, ActionTime)

                Case SaveAction.UpdateExecutionTime
                    File.SetLastAccessTime(Path, ActionTime)

                Case SaveAction.ChangeFavorite
                    Dim formerPath As String = Path
                    Dim newPath As String = Replace(Path, "♥", String.Empty)
                    If Favorite Then
                        Dim pathElements As New List(Of String)(Regex.Split(newPath, "[\\\/]", RegexOptions.None))
                        Dim fileName As String = pathElements.Last
                        pathElements.Remove(fileName)
                        pathElements.Add("♥" & fileName)
                        newPath = Join(pathElements.ToArray, "\")
                    End If
                    Try
                        Using SW As New StreamWriter(newPath)
                            SW.Write(Join({ConnectionText, ScriptText}, BodyDelimiter))
                        End Using
                        _Path = newPath
                    Catch ex As IOException
                        Using message As New Prompt
                            message.Show("Failed to save file", ex.Message, Prompt.IconOption.Critical, Prompt.StyleOption.Grey)
                        End Using
                    End Try
                    Try
                        File.SetLastWriteTime(newPath, Modified) 'Keep original
                        File.SetLastAccessTime(newPath, Ran) 'Keep original
                    Catch ex As IOException
                    End Try
                    Try
                        File.Delete(formerPath)
                    Catch ex As IOException
                        Using message As New Prompt
                            message.Show("Failed to delete file", ex.Message, Prompt.IconOption.Critical, Prompt.StyleOption.Grey)
                        End Using
                    End Try
            End Select
            _DSN_Body = New KeyValuePair(Of String, String)("DSN=" & ConnectionText, ScriptText)
            Return True
        End If

    End Function
    Public Overrides Function ToString() As String

        Dim scriptText As String = If(Text, String.Empty)
        Dim abbreviatedText As String = If(scriptText.Any, If(scriptText.Length > 10, scriptText.Substring(0, 10) & "...", scriptText), String.Empty)
        Return Join({Name, DataSourceName, Type.ToString, abbreviatedText}, BlackOut)

    End Function
#End Region
End Class
Public Class BodyElements
    Implements IDisposable
#Region " DISPOSE "
    Dim disposed As Boolean = False
    ReadOnly Handle As SafeHandle = New Microsoft.Win32.SafeHandles.SafeFileHandle(IntPtr.Zero, True)
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposed Then Return
        If disposing Then
            ' Free any other managed objects here.
            Handle.Dispose()
            ElementsWorker.Dispose()
            ChangedTimer.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public Event ConnectionChanged(sender As Object, e As ConnectionChangedEventArgs)
    Public Event TypeChanged(sender As Object, e As ScriptTypeChangedEventArgs)
    Public Event Completed(sender As Object, e As EventArgs)

    Private WithEvents ElementsWorker As New BackgroundWorker With {.WorkerReportsProgress = True}
    Private WithEvents ChangedTimer As New Timer With {.Interval = 400}         'When Connection or Instruction ( Text ) changes

    Private Const NonCharacter As String = "©"
    Private Const SelectPattern As String = "SELECT[^■]{1,}?(?=FROM)"
    Private Const CommentPattern As String = "--[^\r\n]{1,}(?=\r|\n|$)"
    Private Const ObjectPattern As String = "([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})"     'DataSource.Owner.Name
    'Private Const OrderByPattern As String = "ORDER\s+BY\s+" & ObjectPattern & "(,\s+" & ObjectPattern & "){0,}"           * U N U S E D - BUT USEFUL
    Private Const FromJoinCommaPattern As String = "(?<=FROM |JOIN )[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}|(?<=,)[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}"

    Public Sub New()
        IsBusy = True
        Initializing = True
        Connections = New ConnectionCollection()
        Objects = New SystemObjectCollection
    End Sub

    Public ReadOnly Property HasText As Boolean
        Get
            Return If(Text, String.Empty).Any
        End Get
    End Property
    Private ReadOnly Property Initializing As Boolean = False
    Private ConnectionChange As Boolean = False
    Private _Text As String
    Public Property Text() As String
        Get
            Return _Text
        End Get
        Set(value As String)
            If _Text <> value Then
                ConnectionChange = False
                _Text = value
                _IsBusy = False
                'Timer Stop/Start method ensures that pauses in keystroking >=400ms advance to Timer.Tick
                If If(value, String.Empty).Any Then
                    If Initializing Then
                        'Start immediately
                        ElementsWorker.RunWorkerAsync()
                    Else
                        With ChangedTimer
                            .Stop()
                            .Start()
                        End With
                    End If
                End If
            End If
        End Set
    End Property
    Private LastType As ExecutionType
    Public ReadOnly Property InstructionType As ExecutionType = ExecutionType.Null
    Private LastConnection As Connection
    Private Connection_ As Connection
    Public Property Connection As Connection
        Get
            Return Connection_
        End Get
        Set(value As Connection)
            ConnectionChange = value <> Connection_
            If ConnectionChange Then
                Connection_ = value
                _IsBusy = False
                ChangedTimer_Tick()
            End If
        End Set
    End Property
    Private Sub ChangedTimer_Tick() Handles ChangedTimer.Tick

        With ChangedTimer
            .Stop()
            'If Class is currently working on Text, do no restart the process
            If Not (IsBusy Or ElementsWorker.IsBusy) Then
                _IsBusy = True
                ElementsWorker.RunWorkerAsync()
            End If
        End With

    End Sub
    Public ReadOnly Property UncommentedText As String
    Public ReadOnly Property CountOrLimitText(Optional Limit As Integer = 0) As String
        Get
            'EITHER GET A FULL COUNT OF A QUERY Or LIMIT ROW COUNT IN THE FOLLOWING MANNER
            'COUNT:    Select * From ABC ===> "Select COUNT(*)# FROM ("    SELECT * FROM ABC    "  ) WRAP"
            'LIMIT:    Select * From ABC ===> "Select * FROM ("    SELECT * FROM ABC    "  ) WRAP FETCH FIRST n ROWS ONLY"
            'ie) SELECT STATEMENT MUST BE MODIFIED
            'a) SQL has With(s), INSERT TAKES PLACE NEAR END OF SQL
            'b) SQL does not have With(s), INSERT TAKES PLACE AT START OF SQL

            If If(UncommentedText, String.Empty).Any Then
                Select Case InstructionType
                    Case ExecutionType.SQL
                        Dim Lines As New List(Of String)
                        Dim SelectSection As String
                        If Withs.Any Then
                            'WITH STATEMENTS...
                            '     With BASE (X, Y, Z) AS (SELECT * FROM C085365.PROFILES) Select * FROM BASE
                            '===> With BASE (X, Y, Z) AS (SELECT * FROM C085365.PROFILES) Select * FROM (SELECT * FROM BASE) WRAP FETCH FIRST X ROWS ONLY

                            Dim LastWith As StringData = Withs.Last.Block
                            Dim WithSection As String = UncommentedText.Substring(0, LastWith.Finish)
                            Lines.Add(WithSection)
                            SelectSection = UncommentedText.Remove(0, LastWith.Finish)

                        Else
                            REM /// NO WITH STATEMENTS...INSERT AT POSITION 0 EITHER: SELECT * / SELECT COUNT(*) FROM (...) WRAP
                            '     Select * FROM C085365.PROFILES
                            '===> Select * FROM (SELECT * FROM C085365.PROFILES) WRAP FETCH FIRST X ROWS ONLY
                            SelectSection = UncommentedText


                        End If
                        If Limit = 0 Then
                            REM /// COUNT(*) QUERY
                            Lines.Add("Select COUNT(*)# FROM (")
                            Lines.Add(SelectSection)
                            Lines.Add(") WRAP")

                        Else
                            REM /// SELECT QUERY
                            Lines.Add("Select * FROM (")
                            Lines.Add(SelectSection)
                            Lines.Add(") WRAP")
                            Select Case Language
                                Case QueryLanguage.Netezza
                                    Lines.Add("LIMIT " & Limit)
                                Case Else
                                    Lines.Add("FETCH FIRST " & Limit & " ROWS ONLY")
                            End Select
                        End If
                        Return Join(Lines.ToArray, vbNewLine)

                    Case ExecutionType.DDL
                        'DELETE FROM ABC WHERE X=0 ===> SELECT COUNT(*)# FROM ABC WHERE X=0 [REMOVE REPLACE DELETE WITH SELECT COUNT(*)#...]
                        'INSERT INTO ABC WITH ABC...SELECT * FROM XYZ ===> SELECT COUNT(*)# FROM XYZ [REMOVE INSERT ...]
                        'UPDATE ABC SET X=0 WHERE... ===> SELECT COUNT(*)# FROM ABC WHERE [REMOVE UPDATE ...]
                        'Dim DDLs As New List(Of String)(Split)
                        Return Nothing  'For now    ... translate DDL to SQL

                    Case Else
                        Return Nothing

                End Select
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property SystemText As String
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private ReadOnly Property Connections As ConnectionCollection
    Private ReadOnly Property Objects As SystemObjectCollection
    Public ReadOnly Property IsBusy As Boolean
    Public ReadOnly Property IsDDL() As Boolean
    Public ReadOnly Property Labels As New List(Of InstructionElement)
    Public ReadOnly Property GroupedLabels As Dictionary(Of InstructionElement.LabelName, List(Of InstructionElement))
        Get
            Dim gl As New Dictionary(Of InstructionElement.LabelName, List(Of InstructionElement))
            Dim dl As New List(Of InstructionElement)(Labels.Distinct)
            For Each Label In dl
                If Not gl.ContainsKey(Label.Source) Then gl.Add(Label.Source, New List(Of InstructionElement))
                gl(Label.Source).Add(Label)
                gl(Label.Source).Sort(Function(x, y) x.Block.Start.CompareTo(y.Block.Start))
            Next
            Return gl
        End Get
    End Property
    Public ReadOnly Property TablesFullName As List(Of String)
        Get
            If TablesObject.Any Then
                Return TablesObject.Select(Function(t) t.FullName).ToList
            Else
                Return New List(Of String)
            End If
        End Get
    End Property
    Public ReadOnly Property TablesElement As List(Of InstructionElement)
        Get
            If GroupedLabels.ContainsKey(InstructionElement.LabelName.SystemTable) Then
                Return GroupedLabels(InstructionElement.LabelName.SystemTable)
            Else
                Return New List(Of InstructionElement)
            End If
        End Get
    End Property
    Public ReadOnly Property TablesObject As New List(Of SystemObject)
    Public ReadOnly Property Withs As New List(Of InstructionElement)
    Public ReadOnly Property DataSources As New Dictionary(Of String, List(Of SystemObject))
    Public ReadOnly Property DataSource As SystemObject
    Public ReadOnly Property ElementObjects As New Dictionary(Of InstructionElement, List(Of SystemObject))
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public ReadOnly Property Language As QueryLanguage
        Get
            If IsNothing(Connection) Then
                Return QueryLanguage.None
            Else
                Return Connection.Language
            End If
        End Get
    End Property
    '============================================================================================
    Private Sub BeginObjectsWork(sender As Object, e As DoWorkEventArgs) Handles ElementsWorker.DoWork

        If Not ConnectionChange Then
            'Skip this work when only the Connection changed since only SystemText has a dependancy on Connection properties. 
#Region " RESET COLLECTIONS "
            Withs.Clear()
            Labels.Clear()
            ElementObjects.Clear()
            TablesObject.Clear()
            DataSources.Clear()
#End Region

            REM /// BELOW FUNCTIONS ARE ORDERED BY DEPENDANCY

            'COMMENTED TEXT...Commented text MUST be ignored
            _UncommentedText = StripComments(Text)

#Region " D D L   O r   S Q L "
            LastType = InstructionType
            _InstructionType = GetInstructionType()
            If LastType <> InstructionType Then
                ElementsWorker.ReportProgress(0)
                LastType = InstructionType
            End If
#End Region

            '----------------------------CLASSIFY TEXT
            AssignLabels()

            '----------------------------CROSS-REFERENCE TEXT AS AN OBJECT THAT RESIDES IN A DATABASE
            AddSystemObjects()

            '----------------------------DETEREMINE DATASOURCE
            GetDataSources()

            '----------------------------SET THE CONNECTION
            LastConnection = Connection
            Connection = GetConnection()
        End If
        '----------------------------SET THE DATABASE TEXT
        _SystemText = If(GetSystemText(), Text)


    End Sub
    Private Sub MidObjectsWork(sender As Object, e As ProgressChangedEventArgs) Handles ElementsWorker.ProgressChanged
        RaiseEvent TypeChanged(Me, New ScriptTypeChangedEventArgs(LastType, InstructionType))
    End Sub
    Private Sub EndObjectsWork(sender As Object, e As RunWorkerCompletedEventArgs) Handles ElementsWorker.RunWorkerCompleted

        _Initializing = False
        _IsBusy = False

        If LastConnection <> Connection Then
            RaiseEvent ConnectionChanged(Me, New ConnectionChangedEventArgs(LastConnection, Connection))
            LastConnection = Connection
        End If
        ConnectionChange = False
        RaiseEvent Completed(Me, Nothing)

    End Sub
    '============================================================================================
#Region " FILL PROPERTIES AND VALUES - SEQUENTIAL ORDER "
    Private Function StripComments(TextIn As String) As String

        REM /// -- EXEMPTS TEXT FROM CONSIDERATION, BUT NOT IF IT IS IN APOSTROPHES (CONSTANTS)
        REM /// 1] SELECT  '----------------------' = CONSTANT
        REM /// 2] --SELECT 'SPG'                   = GREENOUT

        Dim _UnCommentedText As String = If(IsNothing(TextIn), String.Empty, TextIn) 'RegEx THROWS AN ERROR FROM A NULL INPUT VALUE...

        Dim GreenOuts As New List(Of StringData)(From M In Regex.Matches(_UnCommentedText, CommentPattern, RegexOptions.IgnoreCase) Select New StringData(M))
        Dim Constants As New List(Of StringData)(From M In Regex.Matches(_UnCommentedText, "'[^'\r\n]{0,}'", RegexOptions.IgnoreCase) Select New StringData(M))

        For Each Constant In Constants
            With Constant
                .BackColor = Color.Gainsboro
                .ForeColor = Color.Black
            End With
            Labels.Add(New InstructionElement With {.Source = InstructionElement.LabelName.Constant,
                             .Block = Constant,
                             .Highlight = Constant})
        Next
        GreenOuts = (From G In GreenOuts Where (From C In Constants Where Not C.Contains(G)).Any).ToList
        For Each GreenOut In GreenOuts
            With GreenOut
                .BackColor = Color.White
                .ForeColor = Color.DarkGreen
            End With
            Labels.Add(New InstructionElement With {.Source = InstructionElement.LabelName.Comment,
                         .Block = GreenOut,
                         .Highlight = GreenOut})
            _UnCommentedText = _UnCommentedText.Remove(GreenOut.Start, GreenOut.Length)
            _UnCommentedText = _UnCommentedText.Insert(GreenOut.Start, StrDup(GreenOut.Length, "-"))
        Next
        Return _UnCommentedText

    End Function
    Private Function GetInstructionType() As ExecutionType

        _IsDDL = False
        REM /// _CommentsReplaced REMOVES POTENTIAL MATCHES FROM TEXT INSIDE A COMMENT...WHICH SHOULD NOT BE CONSIDERED
        Dim _CurrentType As ExecutionType = ExecutionType.Null

        If UncommentedText.Any Then
            Dim Match_Comment As Match = Regex.Match(UncommentedText, "COMMENT\s{1,}ON\s{1,}(TABLE|COLUMN|FUNCTION|TRIGGER|DOCUMENT|PROCEDURE|ROLE|TRUSTED|MASK)\s{1,}", RegexOptions.IgnoreCase)
            Dim Match_Drop As Match = Regex.Match(UncommentedText, "DROP[\s]{1,}(TABLE|VIEW|Function|TRIGGER)[\s]{1,}" & ObjectPattern, RegexOptions.IgnoreCase)
            Dim Match_Insert As Match = Regex.Match(UncommentedText, "INSERT[\s]{1,}INTO[\s]{1,}" + ObjectPattern + "([\s]{0,}\([A-Z0-9!%{}^~_@#$]{1,}(,[\s]{0,}[A-Z0-9!%{}^~_@#$]{1,}){0,}\)){0,}", RegexOptions.IgnoreCase)
            Dim Match_Delete As Match = Regex.Match(UncommentedText, "DELETE[\s]{1,}FROM[\s]{1,}" + ObjectPattern, RegexOptions.IgnoreCase)
            Dim Match_Update As Match = Regex.Match(UncommentedText, "UPDATE[\s]{1,}" + ObjectPattern + "([\s]{1,}([A-Z0-9!%{}^~_@#$]{1,})){0,1}[\s]{1,}Set[\s]{1,}", RegexOptions.IgnoreCase)
            Dim Match_CreateAlterDrop As Match = Regex.Match(UncommentedText, "(CREATE|ALTER|DROP)(\s{1,}OR REPLACE){0,1}\s{1,}((AUXILIARY\s+){0,1}TABLE|(BLOB\s+|CLOB\s+|LOB\s+)TABLESPACE|VIEW|Function|TRIGGER)[\s]{1,}" + ObjectPattern, RegexOptions.IgnoreCase)
            Dim Match_GrantRevoke As Match = Regex.Match(UncommentedText, "(GRANT|REVOKE)[\s]{1,}((Select|UPDATE|INSERT|DELETE|ALTER|INDEX|REFERENCES|EXECUTE)[\s]{0,}[,]{0,1}[\s]{0,}){1,8}[\s]{1,}On[\s]{1,}" + ObjectPattern, RegexOptions.IgnoreCase)
            _IsDDL = Match_Comment.Success Or Match_Drop.Success Or Match_Insert.Success Or Match_Delete.Success Or Match_Update.Success Or Match_CreateAlterDrop.Success Or Match_GrantRevoke.Success
            If IsDDL Then
                _CurrentType = ExecutionType.DDL

            ElseIf Regex.Match(UncommentedText, SelectPattern, RegexOptions.IgnoreCase).Success Then
                _CurrentType = ExecutionType.SQL

            End If
        End If
        Return _CurrentType

    End Function
#Region " LABELS "
    Private Sub AssignLabels()

        REM /// REQUIRES KNOWING IF IsDDL + CALLS ParenthesisNodes
        Dim TextIn As String = UncommentedText
        Dim Blackout_Selects As String = TextIn
        Dim BlackOut_Parentheses As String = TextIn
        Dim Blackout_Handled As String = TextIn

        If IsNothing(TextIn) OrElse TextIn.Length = 0 Then
        Else
            REM /// BEGIN BY IDENTIFYING SIMPLE OBJECTS
#Region " UNION - ADD BLOCK "
            Dim Unions As New List(Of StringData)(From M In Regex.Matches(TextIn, "[\s\r\n]{1,}\b(UNION ALL|UNION|EXCEPT|INTERSECT)\b[\s\r\n]{1,}", RegexOptions.IgnoreCase) Select New StringData(M))
            For Each Union In Unions
                With Union
                    '.BackColor = My.Settings.Union_Back
                    '.ForeColor = My.Settings.Union_Fore
                End With
                Labels.Add(New InstructionElement With {.Source = InstructionElement.LabelName.Union,
                         .Block = Union,
                         .Highlight = Union})
            Next
#End Region
#Region " SELECT STATEMENTS "
            Dim Selects As New List(Of StringData)(From M In Regex.Matches(TextIn, SelectPattern, RegexOptions.IgnoreCase) Select New StringData(M))
            For Each SelectStatement In Selects
                With SelectStatement
                    .BackColor = Color.FromArgb(64, Color.Tomato)
                    .ForeColor = Color.Black
                End With
                Blackout_Selects = ChangeText(Blackout_Selects, SelectStatement)
                Labels.Add(New InstructionElement With {.Source = InstructionElement.LabelName.SelectBlock,
                             .Block = SelectStatement,
                             .Highlight = New StringData With {
                                                .Start = SelectStatement.Start,
                                                .Length = "SELECT".Length,
                                                .Value = "SELECT"
                                                }
                            })
                REM /// COMPLICATED TO DETERMINE END OF FIELD...EXAMPLE:    (CASE LEFT(R.SAI, 2) WHEN 'WW' THEN 'Y' ELSE 'N' END) --H.IN
                Labels.AddRange(FieldsFromBlocks(SelectStatement, "SELECT[\s]{1,}"))
            Next
#End Region
#Region " GROUP BY/ORDER BY "
            Dim GroupBys As New List(Of StringData)(From M In Regex.Matches(TextIn, "\bGROUP[\s]{1,}BY\b[\s]{1,}([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})(,[\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2}){0,}", RegexOptions.IgnoreCase) Select New StringData(M))
            For Each GroupBy In GroupBys
                Dim GroupByHighlight = Regex.Match(GroupBy.Value, "\bGROUP[\s]{1,}BY\b", RegexOptions.IgnoreCase).Value
                Labels.Add(New InstructionElement With {.Source = InstructionElement.LabelName.GroupBlock,
                             .Block = GroupBy,
                             .Highlight = New StringData With {.Start = GroupBy.Start,
                             .Length = GroupByHighlight.Length,
                             .Value = GroupByHighlight}})
                Labels.AddRange(FieldsFromBlocks(GroupBy, "GROUP[\s]{1,}BY[\s]{1,}"))
            Next
            Dim OrderBys As New List(Of StringData)(From M In Regex.Matches(TextIn, "\bORDER[\s]{1,}BY\b[\s]{1,}([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})(,[\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2}){0,}", RegexOptions.IgnoreCase) Select New StringData(M))
            For Each OrderBy In OrderBys
                Dim OrderByHighlight = Regex.Match(OrderBy.Value, "\bORDER[\s]{1,}BY\b", RegexOptions.IgnoreCase).Value
                Labels.Add(New InstructionElement With {.Source = InstructionElement.LabelName.OrderBlock,
                             .Block = OrderBy,
                             .Highlight = New StringData With {.Start = OrderBy.Start,
                             .Length = OrderByHighlight.Length,
                             .Value = OrderByHighlight}})
                Labels.AddRange(FieldsFromBlocks(OrderBy, "ORDER[\s]{1,}BY[\s]{1,}"))
            Next
#End Region
#Region " FETCH/LIMIT "
            Dim Limits As New List(Of StringData)(From M In Regex.Matches(TextIn, "(FETCH[\s]{1,}FIRST[\s]{1,}[0-9]{1,}[\s]{1,}ROWS[\s]{1,}ONLY|LIMIT[\s]{1,}[0-9]{1,})", RegexOptions.IgnoreCase) Select New StringData(M))
            For Each Limit In Limits
                With Limit
                    .BackColor = Color.FromArgb(64, Color.MediumVioletRed)
                    .ForeColor = Color.Black
                End With
                Labels.Add(New InstructionElement With {.Source = InstructionElement.LabelName.Limit,
                             .Block = Limit,
                             .Highlight = New StringData With {.Start = Limit.Start,
                             .Length = 5,
                             .Value = Regex.Match(Limit.Value, "FETCH|LIMIT", RegexOptions.IgnoreCase).Value}})
            Next
#End Region
            REM /// STRIP AWAY PARTS OF TEXT IDENTIFIED SO THEY ARE NOT CONSIDERED AGAIN
            REM /// LOOKS FOR ACCEPTABLE OBJECT NAMING CONVENTIONS- CERTAIN CHARACTERS ARE NOT ALLOWED IN TABLE, VIEW, FUNCTION, AND TRIGGER NAMES + CAN BE AS: {1] DB.OWNER.NAME, 2] OWNER.NAME, 3] NAME}
            Dim PotentialObjects As New List(Of StringData)(From M In Regex.Matches(TextIn, ObjectPattern, RegexOptions.IgnoreCase) Select New StringData(M))
#Region " PROCEDURAL ACTIONS ON SYSTEM.TABLES "
            If IsDDL Then
                Dim Patterns As New Dictionary(Of String, String)
#Region " TRIGGER SECTION - CAN CONTAIN TABLES, VIEWS, FUNCTIONS "
                Patterns.Add("TriggerInsertDelete", "(BEFORE|AFTER|INSTEAD[\s]{1,}OF)[\s]{1,}(INSERT|DELETE)[\s]{1,}ON[\s]")
                Patterns.Add("TriggerUpdate", "(BEFORE|AFTER|INSTEAD[\s]{1,}OF)[\s]{1,}(UPDATE[\s]{1,}OF[\s]{1,})([A-Z0-9!%{}^~_@#$]{1,})([\s]{0,}[,][\s]{0,}[A-Z0-9!%{}^~_@#$]{1,}){0,}[\s]{1,}ON[\s]{1,}")
#End Region
#Region " GRANT/REVOKE SECTION "
                REM /// OTHER DDL COMMANDS: GRANT|REVOKE (SELECT|UPDATE|INSERT|DELETE|ALTER|INDEX|REFERENCES|EXECUTE) ON
                Patterns.Add("GrantRevoke", "(GRANT|REVOKE)[\s]{1,}(SELECT|UPDATE|INSERT|DELETE|ALTER|INDEX|REFERENCES|EXECUTE)[\s]{1,}ON[\s]{1,}(FUNCTION[\s]{1,}){0,}")
#End Region
#Region " ALTER/DROP SECTION "
                REM /// OTHER DDL COMMANDS: ALTER|DROP TABLE (CREATE WOULD BE NEW AND THEREFORE SHOULD NOT COUNT)
                Patterns.Add("AlterDrop", "(ALTER|DROP)[\s]{1,}(TABLE|VIEW)[\s]{1,}")
#End Region
#Region " INSERT/UPDATE/DELETE SECTION "
                REM /// OTHER DDL COMMANDS: INSERT INTO, UPDATE, DELETE FROM}
                Patterns.Add("InsertUpdateDelete", "(INSERT[\s]{1,}INTO|DELETE[\s]{1,}FROM|UPDATE)[\s]{1,}")
#End Region
#Region " ITERATE PATTERNS...TextElements.Add(SystemTable) "
                For Each Pattern In Patterns.Keys
                    Dim Statements As New List(Of StringData)(From M In Regex.Matches(TextIn, Patterns(Pattern) + ObjectPattern, RegexOptions.IgnoreCase) Select New StringData(M))
                    Dim KeyWords As New List(Of String)(From PK In Regex.Matches(Patterns(Pattern), "[A-Z]{2,}", RegexOptions.IgnoreCase) Select DirectCast(PK, Match).Value)
                    Dim Tables = (From S In Statements, PO In PotentialObjects Where S.Contains(PO) And Not KeyWords.Intersect({PO.Value}).Any Select PO)
                    For Each Table In Tables
                        With Table
                            '.BackColor = My.Settings.TableSystem_Back
                            '.ForeColor = My.Settings.TableSystem_Fore
                        End With
                        Dim SystemTableElement As New InstructionElement With {.Source = InstructionElement.LabelName.SystemTable,
                                     .Block = Table,
                                     .Highlight = Table}
                        Labels.Add(SystemTableElement)
                    Next
                    For Each Statement In Statements
                        Blackout_Handled = ChangeText(TextIn, Statement)
                    Next
                Next
#End Region
            End If
#End Region
            Dim Root As New StringData(Blackout_Handled)
            ParenthesisNodes(Root, TextIn)
            BlackOut_Parentheses = Blackout_Handled
#Region " WITH BLOCKS - ALWAYS TEXT OUTSIDE PARENTHESES: WITH ABC (X, Y, Z) AS (SELECT ...) "
            REM /// EASIER TO CAPTURE WITH BLOCKS WHEN IGNORING CONTENT INSIDE WITH(ignore me)
            For Each ParenthesesBlock As StringData In Root.All
                BlackOut_Parentheses = ChangeText(BlackOut_Parentheses, ParenthesesBlock)
            Next
            Dim WithColors As New List(Of Color)({Color.SlateBlue, Color.OrangeRed, Color.Peru, Color.YellowGreen, Color.BlueViolet, Color.Olive, Color.DarkOliveGreen, Color.DarkMagenta})
            REM /// 1] WITH DEBITS ■■■■ AS ■■■■ |2] WITH DEBITS AS ■■■■ |3] , FINAL ■■■■ AS ■■■■ |4] , FINAL AS ■■■■
            Dim WithPattern As String = "(?<=WITH |,)[\s]{0,}[A-Z0-9!%{}^~_@#$]{1,}[\s]{0,}[■]{0,}[\s]AS[\s]{0,}[■]{1,}"
            Dim WithBlocks As New List(Of StringData)(From M In Regex.Matches(BlackOut_Parentheses, WithPattern, RegexOptions.IgnoreCase) Select New StringData(M))
            For Each WithBlock In WithBlocks
                REM /// REGEX LOOKBEHIND MUST HAVE A FIXED LENGTH WHICH MEANS HAVING TO ADJUST THE StringData.START ACCOUNTING FOR PRECEDING SPACES
                Dim WithStart As Integer = 0
                For Each Item In WithBlock.Value
                    If Item = " " Then
                        WithStart += 1
                    Else
                        Exit For
                    End If
                Next
                Dim NewStart As Integer = (WithBlock.Start + WithStart)
                Dim NewLength As Integer = (WithBlock.Length - WithStart)
                Dim WithValue As String = Split(WithBlock.Value.Substring(WithStart, NewLength), " ").First
                WithValue = Split(WithValue, BlackOut).First
                Dim WithElement As New InstructionElement With {
                                 .Source = InstructionElement.LabelName.WithBlock,
                                 .Block = New StringData With {
                                             .Start = NewStart,
                                             .Length = WithBlock.Length - WithStart,
                                             .Value = TextIn.Substring(NewStart, NewLength)
                                 },
                                 .Highlight = New StringData With {
                                            .Start = NewStart,
                                            .Length = WithValue.Length,
                                            .Value = WithValue,
                                            .BackColor = WithColors(_Withs.Count Mod WithColors.Count)
                                 }
                                 }
                Withs.Add(WithElement)
            Next
            Labels.AddRange(Withs)
#End Region
#Region " GET TABLES - WHICH ALWAYS FOLLOW <FROM> - 3 STAGES "
            REM /// IT IS BEST TO HANDLE OUTSIDE () AND INSIDE () SEPARATELY
#Region " OUTSIDE () "
            Dim From_OutsideWiths As New List(Of InstructionElement)(FromBlocks(New StringData With {.Value = BlackOut_Parentheses}))
            Labels.AddRange(From_OutsideWiths)
#End Region
#Region " INSIDE (a) - INNER FROM STATEMENTS ( NO NESTED FROM STATEMENT(s) ) "
            REM /// (SELECT...FROM TABLENAME...WHERE)
            REM /// NEED INNERMOST FROM STATEMENTS FIRST SINCE THEY WILL INTERFERE WITH OUTER FROM STATEMENTS...SELECT A, (SELECT B FROM) FROM (SELECT *)
            Dim FromInnerValuePair As KeyValuePair(Of String, List(Of InstructionElement)) = FromWhittle(Blackout_Handled)
            Labels.AddRange(FromInnerValuePair.Value)
#End Region
#Region " INSIDE (b) - OUTER FROM STATEMENTS ( HAS NESTED FROM STATEMENT(s) ) "
            Dim FromOuterValuePair As KeyValuePair(Of String, List(Of InstructionElement)) = FromWhittle(FromInnerValuePair.Key)
            Labels.AddRange(FromOuterValuePair.Value)
#End Region
#End Region
        End If

    End Sub
    Private Shared Function ChangeText(FullString As String, _StringData As StringData, Optional Value As String = BlackOut) As String

        Dim NewValue As String = FullString
        With _StringData
            NewValue = NewValue.Remove(.Start, .Length)
            NewValue = NewValue.Insert(.Start, StrDup(.Length, Value))
        End With
        Return NewValue

    End Function
    Private Function FromWhittle(TextIn As String) As KeyValuePair(Of String, List(Of InstructionElement))

        Dim Root As New StringData With {.Value = TextIn}
        ParenthesisNodes(Root, TextIn)

        Dim FromsAll As New List(Of StringData)(From PT In Root.All Where PT.Value.ToUpperInvariant.Contains("FROM"))
        Dim FromsWithFroms As New List(Of StringData)(From FA In FromsAll Where (From P In FA.Parentheses Where P.Value.ToUpperInvariant.Contains("FROM")).Any)
        Dim FromsNoFroms As New List(Of StringData)(FromsAll.Except(FromsWithFroms))
        Dim FromElements As New List(Of InstructionElement)
        Dim FromText As String = TextIn
        For Each Section In FromsNoFroms
            FromElements.AddRange(FromBlocks(Section))
            FromText = FromText.Remove(Section.Start, Section.Length)
            FromText = FromText.Insert(Section.Start, StrDup(Section.Length, BlackOut))
        Next
        Return New KeyValuePair(Of String, List(Of InstructionElement))(FromText, FromElements)

    End Function
    Private Function FromBlocks(_StringData As StringData) As List(Of InstructionElement)

        Dim From_SectionValue As String = Nothing
        Dim FromElements As New List(Of InstructionElement)
        Dim WithList As New Dictionary(Of String, Color)
        For Each _With In _Withs
            If WithList.ContainsKey(_With.Highlight.Value) Then
            Else
                WithList.Add(_With.Highlight.Value, _With.Highlight.BackColor)
            End If
        Next

        REM /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        REM /// ***** DUPLICATIONS DUE TO NESTED FROM STATEMENTS...ObjectsFromText.FromsInsideBubbles CALLS FOR EACH FROM IN A BUBBLE
        REM /// FUNCTION TAKES A SECTION OF BODY.TEXT AND SEGMENTS TEXT BLOCKS OF FROM...=>|WHERE
        REM /// FromBlockPattern IS NON-GREEDY SO NEED TO ITERATE MULTIPLE FROM's UNTIL ALL ARE GONE (EX. UNIONS)
        REM /// FROM[^©] = FROM+ANYTHING UP TO A KEY WORD OR EOL... DO NOT USE <BlackOut> AS BUBBLES WILL HAVE BLACKED OUT ANY () IN THE FROM BLOCK
        REM /// NonCharacter As String = "©"
        REM /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        Dim FromBlockPattern As String = "FROM[\s]{1,}[^" & NonCharacter & "]{1,}?(?=\bWHERE\b|\bUNION\b|\bEXCEPT\b|\bINTERSECT\b|\bGROUP\b|\bORDER\b|\bLIMIT\b|\bFETCH\b|\z)"

        REM /// GET THE FROM CHUNK FROM START TO END INCLUDING ALL JOINS, ETC UP TO BUT NOT INCLUDING WHERE, UNION, ETC
        Dim From_Sections As New List(Of StringData)(From M In Regex.Matches(_StringData.Value, FromBlockPattern, RegexOptions.IgnoreCase) Select New StringData(M))

        For Each From_Section In From_Sections
            From_SectionValue = From_Section.Value
            Dim Root As New StringData(From_Section.Value)
            ParenthesisNodes(Root, From_SectionValue)
            If Root.Parentheses.Any Then
                Dim Base = Root.All.First
                From_SectionValue = From_SectionValue.Remove(Base.Start, Base.Length)
                From_SectionValue = From_SectionValue.Insert(Base.Start, StrDup(Base.Length, BlackOut))
            End If
            Do
                REM /// FromBlockPattern IS LAZY...DO IS REQUIRED AS Regex.Match SET TO LAZY
                Dim FromBlockMatch = Regex.Match(From_SectionValue, FromJoinCommaPattern, RegexOptions.IgnoreCase)
                If FromBlockMatch.Success Then
                    Dim InnerItems As New List(Of Match)(From R In Regex.Matches(From_SectionValue, FromJoinCommaPattern, RegexOptions.IgnoreCase) Select DirectCast(R, Match))
                    REM /// InnerItems=EACH MATCH OF OTHER REFERENCED TABLES IN THE FROM BLOCK SUCH AS:     C085365.ACTIONS_TODAY AT
                    For Each InnerItem In InnerItems
                        Dim InnerChunk As String = From_SectionValue.Substring(InnerItem.Index, From_SectionValue.Length - InnerItem.Index)
#Region " TESTING "
                        If InnerItem.Value.Contains("Ø") Then
                            Stop
                        ElseIf InnerItem.Value.Contains("Ø") Then
                            Stop
                        End If
#End Region
                        Dim InnerStart As Integer = 0
#Region " CLEANUP - REMOVE PRECEDING SPACES AND MOVE FORWARD THE START POSITION BASED ON COUNT OF PRECEDING SPACES "
                        For Each Item In InnerItem.Value
                            If Item = " " Then
                                InnerStart += 1
                            Else
                                Exit For
                            End If
                        Next
#End Region
                        Dim NewStart As Integer = _StringData.Start + From_Section.Start + InnerItem.Index + InnerStart
                        Dim InnerValue As String = Split(InnerItem.Value.Substring(InnerStart, InnerItem.Length - InnerStart), " ").First
                        Dim SourceType As InstructionElement.LabelName = Nothing
                        Dim HighlightBackColor As Color = Color.DarkBlue
                        Dim HighlightForeColor As Color = Color.White

                        If WithList.ContainsKey(InnerValue) Then
                            REM /// WITH (a,b) AS (SELECT WILL MATCH IsRoutineTable SO CHECK FIRST
                            SourceType = InstructionElement.LabelName.WithTable
                            HighlightBackColor = Color.White
                            HighlightForeColor = WithList(InnerValue)

                        ElseIf InnerValue.ToUpper(Globalization.CultureInfo.InvariantCulture) = "TABLE" And Regex.Match(InnerChunk, "TABLE[■]{2,}", RegexOptions.IgnoreCase).Success Then
                            REM /// TABLE(SELECT... IS NESTED SO CONTENT OF () *IS* BLACKED OUT
                            SourceType = InstructionElement.LabelName.FloatingTable
                            'HighlightBackColor = My.Settings.TableFloating_Back
                            'HighlightForeColor = My.Settings.TableFloating_Fore

                        ElseIf Regex.Match(InnerChunk, InnerValue & "[■]{1,}", RegexOptions.IgnoreCase).Success Then
                            REM /// XMLTABLE( + OTHER ROUTINE TABLES ARE *NOT* NESTED WITH ANOTHER SELECT STATEMENT SO CONTENT OF () IS NOT BLACKED OUT
                            SourceType = InstructionElement.LabelName.RoutineTable
                            'HighlightBackColor = My.Settings.TableRoutine_Back
                            'HighlightForeColor = My.Settings.TableRoutine_Fore

                        Else
                            SourceType = InstructionElement.LabelName.SystemTable
                            'HighlightBackColor = My.Settings.TableSystem_Back
                            'HighlightForeColor = My.Settings.TableSystem_Fore

                        End If
                        FromElements.Add(New InstructionElement With
                                             {.Source = SourceType,
                                              .Block = New StringData With {
                                                    .Start = _StringData.Start + From_Section.Start,
                                                    .Length = From_Section.Length,
                                                    .Value = From_Section.Value,
                                                    .BackColor = Color.FromArgb(64, Color.Black)},
                                             .Highlight = New StringData With {
                                                    .Start = NewStart,
                                                    .Length = InnerValue.Length,
                                                    .Value = InnerValue,
                                                    .BackColor = HighlightBackColor,
                                                    .ForeColor = HighlightForeColor}}
                                                    )
#Region " REPLACE THE FOUND OBJECT WITH A NON-CHARACTER USED IN THE PATTERN SO IT IS REMOVED FROM CONSIDERATION IN THE NEXT ITERATION "
                        REM /// EACH ITERATION REMOVES A FOUND ITEM AND IS NOT CONSIDERED IN NEXT EVALUATION ///
                        From_SectionValue = From_SectionValue.Remove(InnerItem.Index, InnerItem.Length)
                        From_SectionValue = From_SectionValue.Insert(InnerItem.Index, StrDup(InnerItem.Length, NonCharacter))
#End Region
                    Next
                Else
                    REM /// ALL MATCHES HAVE BEEN REPLACED BY <NonCharacter>. NOTHING REMAINS
                    Exit Do
                End If
            Loop
        Next
        FromElements.Sort(Function(x, y) String.Compare(x.Source.ToString.ToUpperInvariant, y.Source.ToString.ToUpperInvariant, StringComparison.Ordinal))
        Labels.AddRange(FromElements)
        Return FromElements

    End Function
    Private Shared Function FieldsFromBlocks(DataString As StringData, Pattern As String) As List(Of InstructionElement)

        Dim Fields As New List(Of InstructionElement)
        Dim SourceType As InstructionElement.LabelName
        Dim FieldPattern As String = Nothing
        Dim ForeColor As Color = Color.Black
        If Pattern.Contains("GROUP") Then
            SourceType = InstructionElement.LabelName.GroupField
            FieldPattern = "\bGROUP[\s]{1,}BY\b[\s]{1,}"
            ForeColor = Color.DarkOrange

        ElseIf Pattern.Contains("ORDER") Then
            SourceType = InstructionElement.LabelName.OrderField
            FieldPattern = "\bORDER[\s]{1,}BY\b[\s]{1,}"
            ForeColor = Color.Blue

        ElseIf Pattern.Contains("SELECT") Then
            SourceType = InstructionElement.LabelName.SelectField
            FieldPattern = "\bSELECT\b[\s]{1,}"

        End If
        Dim FieldSection As String = DataString.Value.Remove(0, Regex.Match(DataString.Value, FieldPattern, RegexOptions.IgnoreCase).Length)
        Dim FieldSectionNoParenthesis As String = FieldSection
        REM /// REMOVE CONTENT INSIDE () SINCE FUNCTIONS, ETC OFTEN CONTAIN COMMAS WHICH IS NEEDED AS A "§" FOR THE FIELD
        Dim Root As New StringData With {.Value = FieldSection}
        ParenthesisNodes(Root, FieldSection)
        For Each Section In Root.Parentheses
            FieldSectionNoParenthesis = FieldSectionNoParenthesis.Remove(Section.Start, Section.Length)
            FieldSectionNoParenthesis = FieldSectionNoParenthesis.Insert(Section.Start, StrDup(Section.Length, BlackOut))
        Next
        Dim DelimitMatches As New List(Of StringData)(From M In Regex.Matches(FieldSectionNoParenthesis, ",[ ]{0,}", RegexOptions.IgnoreCase) Select New StringData(M))
        For Each Section In DelimitMatches
            FieldSectionNoParenthesis = FieldSectionNoParenthesis.Remove(Section.Start, Section.Length)
            FieldSectionNoParenthesis = FieldSectionNoParenthesis.Insert(Section.Start, StrDup(Section.Length, "½"))
        Next
        FieldSectionNoParenthesis = Regex.Replace(FieldSectionNoParenthesis, " ", "¾")
        Dim FieldStart As Integer = (DataString.Value.Length - FieldSection.Length)
        Dim FieldMatches As New List(Of StringData)(From M In Regex.Matches(FieldSectionNoParenthesis, "[^½\s®]{1,}", RegexOptions.IgnoreCase) Select New StringData(M))

        For Each Field As StringData In FieldMatches
            Dim FieldValue As String = FieldSection.Substring(Field.Start, Field.Length)
            FieldValue = Regex.Replace(FieldValue, "[\t\r\n]", BlackOut)
            FieldValue = Regex.Replace(FieldValue, "■$", String.Empty)
            If FieldValue.StartsWith("½", StringComparison.InvariantCulture) Then Stop
            Dim FieldElement As New StringData With {
                                        .Start = DataString.Start + FieldStart + Field.Start,
                                        .Length = Field.Length,
                                        .Value = FieldValue,
                                        .BackColor = Color.White,
                                        .ForeColor = ForeColor
                                        }
            Fields.Add(New InstructionElement With {.Source = SourceType,
                       .Block = FieldElement,
                       .Highlight = FieldElement
                       })
        Next
        Return Fields

    End Function
#End Region
    Private Sub AddSystemObjects()

        REM /// Translates ME.Text into a list of MY.SETTINGS.SystemObjects
        REM /// LOCAL VARIABLE <SETTINGSOBJECTS> IS A REPLICA OF DataTool.SystemObjects (MY.SETTINGS.SystemObjects)
        If TablesElement.Any Then
            REM /// TEXT BUT NOT A QUERY OR PROCEDURE STATEMENT... ie) LIST: C085365.EMAILS, C085365.ADDRESSES, C.OPENACTH3, C.CUSTACTH3, C.CUSTINDH3
            Dim UnstructuredItems As New List(Of StringData)(From M In Regex.Matches(Text, ObjectPattern, RegexOptions.IgnoreCase) Select New StringData(M))
            For Each Item In UnstructuredItems
                REM /// ONLY ADD WORDS THAT EXIST IN MY.SETTINGS.SystemObjects
                TablesObject.AddRange(Objects.Items(Item.Value))
            Next

        Else
            REM /// SCHEMAS SHOWS DEPTH OF DETAIL IN BODY.TEXT...FROM ACTIONS=1, C085365.ACTIONS=2, DSNA1.C085365.ACTIONS=3
            REM /// USE SystemObjectCollection.Items(DataString As String)

            '======================= ENUMERATION ERRORS FOR PSRR UNIVERSE PA
            For Each Item In TablesElement
                If Not ElementObjects.ContainsKey(Item) Then ElementObjects.Add(Item, Objects.Items(Item.Highlight.Value))
            Next
            For Each Item In TablesElement
                TablesObject.AddRange(Objects.Items(Item.Highlight.Value))
            Next
            '======================= ENUMERATION ERRORS FOR PSRR UNIVERSE PA

        End If
        _TablesObject = TablesObject.Distinct.ToList
        TablesObject.Sort(Function(f1, f2)
                              Dim Level1 = String.Compare(f1.DSN, f2.DSN, StringComparison.InvariantCulture)
                              If Level1 <> 0 Then
                                  Return Level1
                              Else
                                  Dim Level2 = String.Compare(f1.Type.ToString, f2.Type.ToString, StringComparison.InvariantCulture)
                                  If Level2 <> 0 Then
                                      Return Level2
                                  Else
                                      Dim Level3 = String.Compare(f1.Owner, f2.Owner, StringComparison.InvariantCulture)
                                      If Level3 <> 0 Then
                                          Return Level3
                                      Else
                                          Dim Level4 = String.Compare(f1.Name, f2.Name, StringComparison.InvariantCulture)
                                          Return Level4
                                      End If
                                  End If
                              End If
                          End Function)

    End Sub
    Private Sub GetDataSources()

        Dim CommonObjects = From o In Objects, t In TablesObject Where t.DSN = o.DSN Group o By _DSN = o.DSN Into DSNGroup = Group Select New With {.DSN = _DSN, .Tables = DSNGroup.ToList}
        Dim OrderedCount = CommonObjects.Where(Function(x) Not If(x.DSN, String.Empty).Length = 0).OrderByDescending(Function(x) x.Tables.Count)

        _DataSources = OrderedCount.ToDictionary(Function(x) x.DSN, Function(y) y.Tables)
        'SELECT * FROM profiles ===> C085365.PROFILES

        If DataSources.Values.Any Then
            Dim ObjectsInDatasource = DataSources.Values.First
            'SYSIBM.SYSDUMMY1...and others exists in all DB2 databases
            Dim CommonTableObjects = From t In TablesObject Group t By _DSN = t.DSN Into DSNGroup = Group Select New With {.DSN = _DSN, .Tables = DSNGroup.ToList}
            Dim TableOrderedCount = CommonTableObjects.Where(Function(x) Not If(x.DSN, String.Empty).Length = 0).OrderByDescending(Function(x) x.Tables.Count)

            Dim TableOrderedTop = TableOrderedCount.First.Tables
            If ObjectsInDatasource.Intersect(TableOrderedTop).Any Then
                _DataSource = TableOrderedTop.First
            Else
                _DataSource = ObjectsInDatasource.First
            End If
            '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■   TESTING
            'If Text.Contains("Q085365.PSRR_UNIVERSES") And Text.Contains("CAST('NY' AS VARCHAR(5))") Then Stop
            '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■   TESTING
        Else
            REM /// COULD NOT MATCH MY.SETTINGS.SYSTEMOBJECTS TO TEXT
            _DataSource = Nothing
        End If

    End Sub
    Private Function GetConnection() As Connection

        If DataSource Is Nothing Then
            'Check if a DSN is provided
            Dim Datasource_Pattern As String = Join((From S In Connections.Sources Select Join({"\b", S, "\b"}, String.Empty)).ToArray, "|")
            '(ABC|DEF|XYZ)
            Dim MatchDatasourceName As Match = Regex.Match(Text, Datasource_Pattern, RegexOptions.IgnoreCase)
            If MatchDatasourceName.Success Then
                Return Connections.Item("DSN=" & MatchDatasourceName.Value)

            Else
                'Check if a UID is provided
                Dim UserId_Pattern As String = Join((From S In Connections Where S.UserID.Length > 0 Select Join({"\b", S.UserID, "\b"}, String.Empty)).ToArray, "|")
                Dim MatchUserId As Match = Regex.Match(Text, UserId_Pattern, RegexOptions.IgnoreCase)
                If MatchUserId.Success Then
                    Dim UIDs = From S In Connections Where S.UserID = MatchUserId.Value
                    If UIDs.Any Then
                        Return UIDs.First
                    Else
                        Return Nothing
                    End If
                Else
                    Return Nothing
                End If
            End If
        Else
            Return _DataSource.Connection
        End If

    End Function
    Private Function GetSystemText() As String

        REM /// GETTING SYSTEMTEXT IS PREDICATED ON A SHORTHAND APPROACH TO SAVE TYPING THE OWNER NAME, FETCH, ETC
        REM /// THIS ROUTINE WILL FILL IN THE BLANKS TO SUBMIT TO THE DATABASE
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        REM /// WHEN SUBMITTING A REQUEST TO DB2 - IF THE OBJECT IS NOT YOUR OWN, THE OWNER VALUE MUST BE PROVIDED
        REM /// SINCE THE CONNECTION IS NOW SET...LOOK FOR SYSTEMOBJECTS IN ONE DATASOURCE

        REM /// ME.SystemObjects - List(Of SystemObject), ME.SystemTables - List(Of Element)
        REM /// OBJECTIVE IS TO CORRELATE Elements (TEXT) TO SystemObjects (DATABASE)
        If Connection Is Nothing Then
            Return Nothing
        Else
            Dim DatabaseText As String = Text
            REM /// STEP 1] GET FULLNAMES (OWNER+NAME) FOR EACH TABLE/VIEW
            Dim ConnectionDictionary As New Dictionary(Of InstructionElement, String)
            For Each Element In ElementObjects.Keys
                Dim FullName As String = Nothing
                Dim ConnectionCollection = ElementObjects(Element).Where(Function(x) x.DSN = Connection.DataSource).ToList
                If ConnectionCollection.Any Then
                    REM /// IF THERE IS A LIST, IT WILL ONLY HAVE 1 ITEM. SYSTEM OBJECTS ARE DISTINCT AS: DSN+OWNER+NAME
                    FullName = ConnectionCollection.First.FullName
                    'If Text IsNot Nothing AndAlso Text.Contains("WITH DEPENDANTS") Then Stop
                Else
                    REM /// EITHER a) ELEMENT (KEY) HAS AN EMPTY LIST (VALUE) ... THEN NEW ITEM
                    REM /// OR b) ELEMENT (KEY) HAS A LIST (VALUE) BUT NOT WITHIN THE DATASOURCE ... THEN NEW ITEM IN DATASOURCE
                    REM /// NOW ENSURE OWNER+NAME
                    Dim TableViewName As String = Element.Highlight.Value

                    Dim XXX = Objects.Items(TableViewName)

                    Dim Levels As String() = Split(TableViewName, ".")
                    REM /// a) DSNA1.C.REALTIMEH3 b) C.REALTIMEH3 c) DSNA1.REALTIMEH3 d) REALTIMEH3
                    Select Case Levels.Count
                        Case 3
                            REM /// a) DSNA1.C.REALTIMEH3
                            FullName = Join({Levels(1), Levels(2)}, ".")
                        Case 2
                            REM /// b) C.REALTIMEH3 Or c) DSNA1.REALTIMEH3
                            Dim SourcePattern As String = "(" & Join(Connections.Sources.ToArray, "|") & ")[\s]{0,}\."
                            If Regex.Match(TableViewName, SourcePattern, RegexOptions.IgnoreCase).Success Then
                                REM /// c) DSNA1.REALTIMEH3...OWNER NOT STATED
                                FullName = Join({Connection.UserID, Levels(1)}, ".")
                                'If Text IsNot Nothing AndAlso Text.Contains("WITH DEPENDANTS") Then Stop
                            Else
                                REM /// b) C.REALTIMEH3
                                FullName = Join({Levels(0), Levels(1)}, ".")
                                'If Text IsNot Nothing AndAlso Text.Contains("WITH DEPENDANTS") Then Stop
                            End If
                        Case 1
                            REM /// d) REALTIMEH3...OWNER NOT STATED
                            If Text.Contains("_v_relation_column VC") Then Stop
                            'FullName = Join({_Connection.UserID, Levels(0)}, ".")
                            FullName = Levels(0)
                    End Select
                    'If Text IsNot Nothing AndAlso Text.Contains("WITH DEPENDANTS") Then Stop
                End If
                ConnectionDictionary.Add(Element, FullName)
            Next
            REM /// STEP 2] UPDATE THE TEXT WITH THE FULLNAMES + CHANGE LIMIT TO FETCH (FOR DB2)
            REM /// MUST SORT ON ALL OBJECTS SINCE CHANGING BOTH SystemTable AND Limit
            Dim ReverseOrder As New List(Of InstructionElement)(Labels)
            ReverseOrder.Sort(Function(y, x) x.Highlight.Start.CompareTo(y.Highlight.Start))
            For Each Element In ReverseOrder
                If Element.Source = InstructionElement.LabelName.SystemTable And ConnectionDictionary.ContainsKey(Element) Then
                    REM /// IS SYSTEM TABLE
                    With Element.Highlight
                        DatabaseText = DatabaseText.Remove(.Start, .Length)
                        DatabaseText = DatabaseText.Insert(.Start, ConnectionDictionary(Element))
                    End With
                Else
                    REM /// LIMIT
                    If Element.Source = InstructionElement.LabelName.Limit And Connection.Language = QueryLanguage.Db2 Then
                        Dim Limit As InstructionElement = Element
                        With Limit
                            Dim RowCount As Integer = Integer.Parse(Regex.Match(.Block.Value, "[0-9]{1,}", RegexOptions.None).Value, Globalization.CultureInfo.InvariantCulture)
                            Try
                                Dim LimitText As String = DatabaseText.Substring(.Block.Start, .Block.Length)
                                If LimitText.ToUpper(Globalization.CultureInfo.InvariantCulture).StartsWith("LIMIT", StringComparison.InvariantCulture) Then
                                    DatabaseText = DatabaseText.Remove(.Block.Start, .Block.Length)
                                    DatabaseText = DatabaseText.Insert(.Block.Start, Join({"FETCH FIRST", RowCount.ToString(Globalization.CultureInfo.InvariantCulture), "ROWS ONLY"}))
                                    If Not Regex.Match(DatabaseText, "FETCH\s+FIRST\s+[0-9]{1,9}\s+ROWS\s+ONLY", RegexOptions.IgnoreCase).Success Then
                                        Stop
                                    End If
                                End If
                            Catch ex As ArgumentOutOfRangeException
                            End Try
                        End With
                    End If
                End If
            Next
            'If Text = "DROP TABLE AXE" Then Stop
            Return DatabaseText
        End If

    End Function
#End Region
End Class
Public Class DataTool
    Inherits Control
#Region " DECLARATIONS "
#Region " organised "
    Private ReadOnly GothicFont As Font = My.Settings.applicationFont
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Private WithEvents FunctionsStripBar As New ToolStrip With {
        .Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .GripStyle = ToolStripGripStyle.Hidden,
        .Font = GothicFont
    }
    Private WithEvents FileTree As New TreeViewer With {.Name = "Scripts",
        .AutoSize = True,
        .Margin = New Padding(0),
        .MouseOverExpandsNode = False,
        .Font = GothicFont,
        .FavoritesFirst = True}
    Private WithEvents SaveAs As New ImageCombo With {.Margin = New Padding(0),
        .Image = My.Resources.Save,
        .Size = New Size(2 + .Image.Width + 2, 2 + .Image.Width + 2),
        .Font = GothicFont,
        .MinimumSize = .Size,
        .AutoSize = True}
    Private WithEvents FilesButton As New ToolStripDropDownButton With {
        .Margin = New Padding(0),
        .Image = My.Resources.Folder,
        .ImageScaling = ToolStripItemImageScaling.None,
        .Font = GothicFont
    }
    Private WithEvents MessageButton As New ToolStripDropDownButton With {
        .Margin = New Padding(0),
        .Image = My.Resources.message,
        .ImageScaling = ToolStripItemImageScaling.None,
        .Font = GothicFont
    }
    Private WithEvents MessageRicherBox As New RicherTextBox With {
        .Margin = New Padding(0),
        .Dock = DockStyle.Fill,
        .AcceptsTab = True
    }
    Private WithEvents SettingsButton As New ToolStripDropDownButton With {
        .Margin = New Padding(0),
        .Image = My.Resources.settings,
        .ImageScaling = ToolStripItemImageScaling.None,
        .Font = GothicFont
    }
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Private WithEvents SettingsTreeA As New TreeViewer With {
        .Margin = New Padding(0),
        .AutoSize = False,
        .MinimumSize = New Size(250, 400),
        .Dock = DockStyle.Fill,
        .Font = GothicFont
    }
    Private WithEvents SettingsTreeB As New TreeViewer With {
        .Margin = New Padding(0),
        .AutoSize = False,
        .MinimumSize = New Size(250, 400),
        .Dock = DockStyle.Fill,
        .Font = GothicFont,
        .FavoritesFirst = True
    }
    Private ReadOnly SettingsTrees_dictionary As New Dictionary(Of Node, NodeCollection)
    Private ReadOnly SettingsDictionary As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String))) From {
        {"Fonts and Colors", New Dictionary(Of String, Dictionary(Of String, String)) From {
                {"Pane", New Dictionary(Of String, String) From {
                {"Font", "paneFont"},
                {"Backcolor", "paneBackColor"},
                {"Forecolor", "paneForeColor"}
        }},
                {"ViewerHeader", New Dictionary(Of String, String) From {
                {"Backcolor", "gridHeaderBackColor"},
                {"Forecolor", "gridHeaderForeColor"},
                {"Shadecolor", "gridHeaderShadeColor"}
        }},
                {"ViewerGrid", New Dictionary(Of String, String) From {
                {"Font", "gridFont"},
                {"RowBackcolor", "gridRowBackColor"},
                {"RowShadecolor", "gridRowShadeColor"},
                {"RowForecolor", "gridRowForeColor"},
                {"AlternatingRowBackcolor", "gridRowAlternatingBackColor"},
                {"AlternatingRowShadecolor", "gridRowAlternatingShadeColor"},
                {"AlternatingRowForecolor", "gridRowAlternatingForeColor"},
                {"SelectionRowBackcolor", "gridRowSelectionBackColor"},
                {"SelectionRowShadecolor", "gridRowSelectionShadeColor"},
                {"SelectionRowForecolor", "gridRowSelectionForeColor"}
        }},
                {"Application", New Dictionary(Of String, String) From {
                {"Font", "applicationFont"},
                {"Backcolor", "applicationBackColor"},
                {"Forecolor", "applicationForeColor"}
        }}
        }
        },
        {"DDL Settings", New Dictionary(Of String, Dictionary(Of String, String)) From {
                {"Dialogue", New Dictionary(Of String, String) From {
                {"Prompt to OK", "ddlPrompt"},
                {"Get row count", "ddlRowCount"}
        }}
        }
        }
        }
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Private WithEvents AddTab As Tab
    Private WithEvents TLP_Objects As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .ColumnCount = 1,
        .RowCount = 3,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .Margin = New Padding(0),
        .Font = GothicFont
    }
    Private WithEvents Tree_Objects As New TreeViewer With {.Name = "Database Objects",
        .Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .DropHighlightColor = Color.Gold,
        .CheckBoxes = TreeViewer.CheckState.Mixed,
        .MultiSelect = True,
        .Font = GothicFont
    }
    Friend WithEvents Script_Tabs As New Tabs With {
        .Dock = DockStyle.Fill,
        .UserCanAdd = True,
        .UserCanReorder = True,
        .MouseOverSelection = True,
        .AddNewTabColor = Color.Black,
        .Font = GothicFont,
        .Alignment = TabAlignment.Top,
        .Multiline = True,
        .Margin = New Padding(0),
        .SelectedTabColor = Color.Black
    }
    Private WithEvents Script_Grid As New DataViewer With {.Dock = DockStyle.Fill,
        .Font = GothicFont,
        .AllowDrop = True,
        .Margin = New Padding(0)
    }
#Region " EXPORT DATA "
    Private WithEvents Grid_DatabaseExport As New ToolStripMenuItem With {.Text = "Database",
        .Image = My.Resources.Database.ToBitmap,
        .ImageScaling = ToolStripItemImageScaling.None,
        .Font = GothicFont}
#End Region
    Private WithEvents CMS_PaneOptions As New ContextMenuStrip With {.AutoClose = False,
        .Padding = New Padding(0),
        .ImageScalingSize = New Size(15, 15),
        .DropShadowEnabled = True,
        .Renderer = New CustomRenderer,
        .BackColor = Color.Gainsboro,
        .Font = GothicFont}
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Collections
    Friend Shared ReadOnly Connections As New ConnectionCollection
    Private ReadOnly SystemObjects As New SystemObjectCollection
    Private ReadOnly Jobs As New JobCollection
    Private WithEvents Scripts As New ScriptCollection
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ Panel Sizing
    Private ObjectsWidth As Integer = 200
    Private SeparatorSizing As New Sizing
    Private WithEvents TLP_PaneGrid As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .ColumnCount = 3,
        .RowCount = 1,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .AllowDrop = True,
        .Margin = New Padding(0),
        .Font = GothicFont
    }
#End Region
#Region " organise "
    Private ReadOnly Message As New Prompt With {.Font = GothicFont}
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
    Private WithEvents CMS_ExcelSheets As New ContextMenuStrip With {.AutoClose = False,
        .AutoSize = True,
        .Margin = New Padding(0),
        .DropShadowEnabled = False,
        .BackColor = Color.WhiteSmoke,
        .ForeColor = Color.DarkViolet,
        .Font = GothicFont}
    Private ReadOnly OpenFileNode As Node = New Node With {.Text = "Open File",
        .Image = My.Resources.Folder,
        .CanEdit = False,
        .CanRemove = False,
        .CanDragDrop = False,
        .Font = GothicFont}
    Private WithEvents OpenFile As New OpenFileDialog
    Private WithEvents SaveFile As New SaveFileDialog
    '-----------------------------------------
    Private WithEvents DragNode As Node
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
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
                .Font = GothicFont}
    '-----------------------------------------
    Private WithEvents TSMI_Divider As New ToolStripMenuItem With {.Text = "Insert divider",
        .Image = My.Resources.InsertBefore,
        .Font = GothicFont}
    Private WithEvents TSMI_Tidy As New ToolStripMenuItem With {.Text = "Tidy",
        .Image = My.Resources.Broom,
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
    Private WithEvents Dialogue_Font As New FontDialog With {.Font = My.Settings.paneFont}
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
    Private WithEvents FindAndReplace As New FindReplace With {.Font = GothicFont}
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private ReadOnly ObjectsSet As New DataSet With {.DataSetName = "Objects"}
    Private WithEvents ObjectsWorker As New BackgroundWorker With {.WorkerReportsProgress = True}
    Private SyncWorkers As Dictionary(Of String, BackgroundWorker)
    Private SyncSet As Dictionary(Of String, DataTable)
    Private WithEvents Stop_Watch As New Stopwatch
    Private ReadOnly Intervals As New Dictionary(Of String, TimeSpan)
    Private ReadOnly ObjectsDictionary As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of SystemObject.ObjectType, List(Of SystemObject))))
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Pane_MouseLocation As Point
    Private Pane_MouseObject As InstructionElement
#End Region
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

        Directory.CreateDirectory(MyDocuments & "\DataManager")

        '*** Before changes=4394
        'Sync populates a Treeview with Checkmarks...those selected are imported. Submit how?
        'Interface to add, change, or remove a connection
        'Export / ETL
        'Casting using Select Min(Length(Trim(Field))), Max(Length(Trim(Field)))...Where Length(Trim(Field))>0

        Dim timeStart As Date = Now
        Dock = DockStyle.Fill
        Me.TestMode = TestMode
        Scripts.Load()
        Dim timeStop As Date = Now
        Dim timeElapsed = timeStop.Subtract(timeStart)

        AddTab = Script_Tabs.AddTab

#Region " FILL COLLECTIONS - { Connections, Jobs, SystemObjects, Scripts } "
#Region " CONNECTIONS "
        Connections.SortCollection()
        'Connections.View()
        For Each Connection In Connections
            RaiseEvent Alert(Me, New AlertEventArgs("Initializing " & Connection.DataSource))
#Region " TOP LEVEL "
            Dim ColorKeys = ColorImages()
            Dim ColorImage As Image = ColorKeys(Connection.BackColor)
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
#End Region

#Region " INITIALIZE CONTROLS "
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ FunctionsStripBar
        With FunctionsStripBar
            .Items.Add(FilesButton)
            .Items.Add(New ToolStripControlHost(SaveAs) With {.AutoSize = True})
            .Items.Add(MessageButton)
            .Items.Add(SettingsButton)
        End With
        Dim tlpMessage As New TableLayoutPanel With {
            .ColumnCount = 1,
            .RowCount = 1,
            .Size = New Size(600, 400),
            .Margin = New Padding(0),
            .Font = New Font("IBM Plex Mono Light", 10, FontStyle.Regular)
        }
        With tlpMessage
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 600})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 400})
            .Controls.Add(MessageRicherBox)
            TLP.SetSize(tlpMessage)
        End With
        MessageButton.DropDownItems.Add(New ToolStripControlHost(tlpMessage) With {.AutoSize = True})
        '===============================================================================
        Dim tlpFileTree As New TableLayoutPanel With {.Size = New Size(200, 200),
        .ColumnCount = 1,
        .RowCount = 1,
        .Font = GothicFont,
        .MinimumSize = New Size(20, 20),
        .AutoSize = True}
        With tlpFileTree
            .Controls.Add(FileTree, 0, 0)
        End With
        FilesButton.DropDownItems.Add(New ToolStripControlHost(tlpFileTree) With {.AutoSize = True})
        '===============================================================================
        Dim tlpSettings As New TableLayoutPanel With {
        .Size = New Size(600, 1000),
        .Margin = New Padding(0),
        .ColumnCount = 2,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.InsetDouble,
        .AutoSize = False
        }
        For Each rootSetting In SettingsDictionary
            Dim rootNode As New Node With {
        .CanAdd = False,
        .CanDragDrop = False,
        .CanEdit = False,
        .CanRemove = False,
        .Text = rootSetting.Key,
        .Font = GothicFont
    }
            SettingsTreeA.Nodes.Add(rootNode)
            Dim nodesB As New NodeCollection(SettingsTreeB)
            SettingsTrees_dictionary.Add(rootNode, nodesB)
            For Each subSetting In rootSetting.Value
                Dim subNode As New Node With {
        .CanAdd = False,
        .CanDragDrop = False,
        .CanEdit = False,
        .CanRemove = False,
        .Text = subSetting.Key,
        .Font = GothicFont
    }
                nodesB.Add(subNode)
                For Each subSubSetting In subSetting.Value
                    Dim subSubNode As New Node With {
        .CanAdd = False,
        .CanDragDrop = False,
        .CanEdit = False,
        .CanRemove = False,
        .Text = subSubSetting.Key,
        .Name = subSubSetting.Value,
        .Font = GothicFont
    }
                    subNode.Nodes.Add(subSubNode)
                Next
            Next
        Next

        With tlpSettings
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 300})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 700})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 600})
            .Controls.Add(SettingsTreeA, 0, 0)
            .Controls.Add(SettingsTreeB, 1, 0)
        End With
        With SettingsButton
            .DropDownItems.Add(New ToolStripControlHost(tlpSettings))
        End With
        '===============================================================================
        '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ DatabaseObjects
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
            Tree_Objects.MultiSelect = False
        End With
        With TLP_PaneGrid
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 0})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 50})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Percent, .Width = 50})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 100})
            .Controls.Add(TLP_Objects, 0, 0)
            .Controls.Add(Script_Tabs, 1, 0)
            .Controls.Add(Script_Grid, 2, 0)
        End With
        Dim tlpBasePanel As New TableLayoutPanel With {.Margin = New Padding(0),
            .ColumnCount = 1,
            .RowCount = 2,
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
            .Dock = DockStyle.Fill}
        With tlpBasePanel
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Percent, .Height = 100})
            .Controls.Add(FunctionsStripBar, 0, 0)
            .Controls.Add(TLP_PaneGrid, 0, 1)
        End With
        Controls.Add(tlpBasePanel)
        With Script_Grid
            With .Columns
                With .HeaderStyle
                    .BackColor = My.Settings.gridHeaderBackColor
                    .ShadeColor = My.Settings.gridHeaderShadeColor
                    .ForeColor = My.Settings.gridHeaderForeColor
                    AddHandler .PropertyChanged, AddressOf Viewer_CellStyleChanged
                End With
            End With
            With .Rows
                With .AlternatingRowStyle
                    .BackColor = My.Settings.gridRowAlternatingBackColor
                    .ShadeColor = My.Settings.gridRowAlternatingShadeColor
                    .ForeColor = My.Settings.gridRowAlternatingForeColor
                    AddHandler .PropertyChanged, AddressOf Viewer_CellStyleChanged
                End With
                With .RowStyle
                    .BackColor = My.Settings.gridRowBackColor
                    .ShadeColor = My.Settings.gridRowShadeColor
                    .ForeColor = My.Settings.gridRowForeColor
                    AddHandler .PropertyChanged, AddressOf Viewer_CellStyleChanged
                End With
                With .SelectionRowStyle
                    .BackColor = My.Settings.gridRowSelectionBackColor
                    .ShadeColor = My.Settings.gridRowSelectionShadeColor
                    .ForeColor = My.Settings.gridRowSelectionForeColor
                    AddHandler .PropertyChanged, AddressOf Viewer_CellStyleChanged
                End With
            End With
            .GridOptions.Items.Add(Grid_DatabaseExport)
            .AllowDrop = True
            .BaseForm = Nothing
        End With

        Dim TSCH_TypeHost As New ToolStripControlHost(TLP_Type) With {.ImageScaling = ToolStripItemImageScaling.None}
        TSMI_ObjectType.DropDownItems.Add(TSCH_TypeHost)

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
        LoadSystemObjects(Nothing, Nothing) ' LOADS FROM Objects.txt
        ExpandCollapseOnOff(HandlerAction.Add)

    End Sub
    Protected Overrides Sub InitLayout()
        UpdateParentIcon_Text()
        MyBase.InitLayout()
    End Sub
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
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ ALERTS
    Public Event Alert(sender As Object, e As AlertEventArgs)
    Private Sub ScriptAlerts(sender As Object, e As AlertEventArgs)
        RaiseEvent Alert(sender, e)
    End Sub
    Private Sub ViewerAlerts(sender As Object, e As AlertEventArgs) Handles Script_Grid.Alert
        RaiseEvent Alert(sender, e)
    End Sub
    Private Sub TreeViewerAlerts(sender As Object, e As AlertEventArgs) Handles FileTree.Alert, Tree_Objects.Alert
        RaiseEvent Alert(sender, e)
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ ALERTS
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Public Property TestMode As Boolean = False
    Public ReadOnly Property Viewer As DataViewer
        Get
            Return Script_Grid
        End Get
    End Property
    Public ReadOnly Property Pane As RicherTextBox
        Get
            Return ActivePane()
        End Get
    End Property
    Public ReadOnly Property Grid As DataViewer
        Get
            Return Script_Grid
        End Get
    End Property
    Private Function ActiveTab() As Tab
        Return Script_Tabs.TabPages.Item({Script_Tabs.SelectedIndex, 0}.Max)
    End Function
    Private Function ActivePane() As RicherTextBox
        If ActiveTab()?.Controls.Count = 0 Then
            Return Nothing
        Else
            Return DirectCast(ActiveTab()?.Controls(0), RicherTextBox)
        End If
    End Function
    Private Function ActiveScript() As Script
        Return DirectCast(ActivePane()?.Tag, Script)
    End Function
    Private Function ActiveBody() As BodyElements
        Return ActiveScript()?.Body
    End Function
#End Region

#Region " CONNECTION MANAGEMENT "
    Private Sub DataSource_Clicked(sender As Object, e As EventArgs)
        With DirectCast(sender, ToolStripMenuItem)
            ActiveScript.Connection = Connections.Item("DSN=" & .Text & ";")
        End With
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
                .BackgroundImage = My.Resources.glossyYellow
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

#Region " FUNCTIONS STRIPBAR "
#Region " SETTINGS "
    Private Sub Viewer_CellStyleChanged(sender As Object, e As StyleEventArgs)

        Dim changedName As String = Nothing
        With Script_Grid
            If sender Is .Columns.HeaderStyle Then
                changedName = "gridHeader"

            ElseIf sender Is .Rows.AlternatingRowStyle Then
                changedName = "gridRowAlternating"

            ElseIf sender Is .Rows.RowStyle Then
                changedName = "gridRow"

            ElseIf sender Is .Rows.SelectionRowStyle Then
                changedName = "gridRowSelection"

            End If
        End With
        Dim changedProperty As System.Configuration.SettingsPropertyValue = NameToProperty(changedName & e.PropertyName)
        If changedProperty IsNot Nothing Then
            changedProperty.PropertyValue = e.PropertyValue
            My.Settings.Save()
            If SettingsTreeB.Nodes.Any Then
                Dim settingNode As Node = SettingsTreeB.Nodes.ItemByTag(changedProperty)
                If settingNode IsNot Nothing Then
                    Dim settingColor As Color = DirectCast(changedProperty.PropertyValue, Color)
                    With settingNode
                        .Image = ColorImages(settingColor)
                        Dim nodeText As String = .Text
                        Dim textElements As String() = Regex.Split(nodeText, " \(", RegexOptions.None)
                        .Text = textElements.First & " (" & settingColor.Name & ")"
                    End With
                End If
            End If
            'If Now.Minute > 1 Then Stop
        End If

    End Sub
    Private Sub SettingA_nodeClick(sender As Object, e As NodeEventArgs) Handles SettingsTreeA.NodeClicked

        With SettingsTreeB.Nodes
            .Clear()
            Dim settingNodes As NodeCollection = SettingsTrees_dictionary(e.Node)
            .AddRange(settingNodes)
            For Each settingNode As Node In settingNodes
                For Each childNode In settingNode.Nodes
                    With childNode
                        Dim settingItem As System.Configuration.SettingsPropertyValue = My.Settings.PropertyValues(.Name)
                        Select Case settingItem.Property.PropertyType
                            Case GetType(Font)
                                .Favorite = True

                            Case GetType(Boolean)
                                .CheckBox = True
                                .Checked = DirectCast(settingItem.PropertyValue, Boolean)

                            Case GetType(Color)
                                Dim settingColor As Color
                                If settingItem.PropertyValue Is Nothing Then
                                    If .Name.ToUpperInvariant.Contains("BACK") Then settingColor = Color.White
                                    If .Name.ToUpperInvariant.Contains("FORE") Then settingColor = Color.Black
                                    If .Name.ToUpperInvariant.Contains("SHADE") Then settingColor = Color.Gainsboro
                                Else
                                    settingColor = DirectCast(settingItem.PropertyValue, Color)
                                End If
                                .Image = ColorImages(settingColor)
                                .Separator = If({"gridRowBackColor", "gridRowSelectionBackColor"}.Contains(.Name), Node.SeparatorPosition.Above, Node.SeparatorPosition.None)
                                Dim nodeText As String = .Text
                                Dim textElements As String() = Regex.Split(nodeText, " \(", RegexOptions.None)
                                .Text = textElements.First & " (" & settingColor.Name & ")"
                        End Select
                        .Tag = settingItem
                    End With
                Next
            Next
        End With
        SettingsTreeB.ExpandNodes()

    End Sub
    Private Sub SettingB_nodeChecked(sender As Object, e As NodeEventArgs) Handles SettingsTreeB.NodeChecked

        With My.Settings
            Dim settingItem As System.Configuration.SettingsPropertyValue = My.Settings.PropertyValues(e.Node.Name)
            settingItem.PropertyValue = e.Node.Checked
            .Save()
        End With

    End Sub
#End Region
#Region " OPEN + SAVE "
    Private Sub FileTree_NodeAfterEdited(sender As Object, e As NodeEventArgs) Handles FileTree.NodeAfterEdited

        Using cb As New CursorBusy
            'USING Now.ToLongTimeString ENSURE NAME<>value AND ACTION IS TAKEN
            Dim ClosedScript As Script = DirectCast(e.Node.Tag, Script)
            ClosedScript.Name = Join({DateTimeToString(Now), e.ProposedText}, Delimiter)
            If ClosedScript.Name = e.ProposedText Then e.Node.Text = e.ProposedText     'Script.Name will only change if it can
            e.Node.Parent.SortChildren()
            FileTree.Refresh()
        End Using

    End Sub
    Private Sub FileTree_NodeAfterRemoved(sender As Object, e As NodeEventArgs) Handles FileTree.NodeAfterRemoved

        Dim RemoveScript As Script = DirectCast(e.Node.Tag, Script)
        RemoveScript.State = Script.ViewState.None

    End Sub
    Private Sub FileTree_NodeDoubleClicked(sender As Object, e As NodeEventArgs) Handles FileTree.NodeDoubleClicked

        If e.Node.Tag IsNot Nothing Then
            If e.Node.Tag.GetType Is GetType(Connection) Then
                '( Root node = Connection ) Adding empty sql pane connected to clicked node Connection
                Dim emptyScript As New Script With {.Connection = DirectCast(e.Node.Tag, Connection)}
                Scripts.Add(emptyScript)
                emptyScript.State = Script.ViewState.OpenDraft

            ElseIf e.Node.Tag.GetType Is GetType(Script) Then
                '( Child node = Script ) Adding populated sql pane from a closed script
                Dim savedScript As Script = DirectCast(e.Node.Tag, Script)
                savedScript.State = Script.ViewState.OpenSaved

            End If
        End If
        FilesButton.HideDropDown()

    End Sub
    Private Sub FileTree_NodeDragStart(sender As Object, e As NodeEventArgs) Handles FileTree.NodeDragStart

        DragNode = e.Node
        Dim dragScript As Script = DirectCast(DragNode.Tag, Script)
        dragScript.Body.Text = dragScript.Text
        If ActivePane() IsNot Nothing Then ActivePane.AllowDrop = True
        Script_Grid.AllowDrop = True

    End Sub
    Private Sub FileTree_NodeClicked(sender As Object, e As NodeEventArgs) Handles FileTree.NodeClicked
        If e.Node Is OpenFileNode Then
            OpenFile.Tag = Nothing
            OpenFile.ShowDialog()
        End If
    End Sub
    Private Sub FileTree_NodeFavorited(sender As Object, e As NodeEventArgs) Handles FileTree.NodeFavorited

        Dim favoriteScript As Script = DirectCast(e.Node.Tag, Script)
        favoriteScript.Favorite = e.Node.Favorite

    End Sub
    '===============================================================================
    Private Sub SaveAs_MouseEnter(sender As Object, e As EventArgs) Handles SaveAs.MouseEnter

        SaveAs.MaximumSize = Nothing
        If ActiveScript() IsNot Nothing Then
            If ActiveScript() IsNot Nothing Then
                SaveAs.Text = ActiveScript.Name
                SaveAs.Image = If(ActiveScript.FileTextMatchesText, My.Resources.saved, My.Resources.savedNot)
            End If
        End If
        SaveAs.AutoSize = True

    End Sub
    Private Sub SaveAs_MouseLeave(sender As Object, e As EventArgs) Handles SaveAs.MouseLeave

        SaveAs.AutoSize = True
        SaveAs.Text = String.Empty
        SaveAs.MaximumSize = New Size(My.Resources.Save.Width + 2, My.Resources.Save.Width + 2)

    End Sub
    Private Sub SaveAs_TextChanged(sender As Object, e As EventArgs)

        If If(SaveAs.Text, String.Empty).Any Then
            SaveAs.AutoSize = True
            RemoveHandler SaveAs.TextChanged, AddressOf SaveAs_TextChanged
        End If

    End Sub
    Private Sub SaveAs_ClearTextClicked(sender As Object, e As EventArgs) Handles SaveAs.ClearTextClicked
        SaveAs.AutoSize = False
        AddHandler SaveAs.TextChanged, AddressOf SaveAs_TextChanged
    End Sub
    Private Sub SaveAs_ImageClicked() Handles SaveAs.ImageClicked, SaveAs.ValueSubmitted

        Dim saveScript As Script = ActiveScript()
        If saveScript IsNot Nothing Then
            Using cb As New CursorBusy
                'USING Now.ToLongTimeString ENSURE NAME<>value AND ACTION IS TAKEN
                Dim ActiveScriptName As String = Join({DateTimeToString(Now), SaveAs.Text}, Delimiter)
                saveScript.Name = ActiveScriptName
                If saveScript.Save(Script.SaveAction.ChangeContent) Then SaveAs.Image = My.Resources.saved
            End Using
        End If

    End Sub
#End Region
    Private Sub ButtonMouseEnter(sender As Object, e As EventArgs) Handles MessageButton.MouseEnter
        If sender Is MessageButton And MessageRicherBox.Text.Any Then MessageButton.ShowDropDown()
    End Sub
#End Region
#Region " PANEL SIZING - OBJECTS→|←PANE→|←GRID "
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
    Private Sub Panel_MouseCaptureChanged(sender As Object, e As EventArgs) Handles TLP_PaneGrid.MouseCaptureChanged
        TLP_PaneGrid.Capture = _ForceCapture
        AddTab.Capture = False
        TLP_Objects.Capture = False
    End Sub
    Private Sub PanelObjects_MouseCaptureChanged(sender As Object, e As EventArgs) Handles TLP_Objects.MouseCaptureChanged
        TLP_Objects.Capture = _ForceCapture
        AddTab.Capture = False
        TLP_PaneGrid.Capture = False
    End Sub
    Private Sub AddTab_MouseCaptureChanged(sender As Object, e As EventArgs) Handles AddTab.MouseCaptureChanged
        AddTab.Capture = _ForceCapture
        TLP_PaneGrid.Capture = False
        TLP_Objects.Capture = False
    End Sub
    Private ReadOnly Property ObjectsPaneSeparator As Rectangle
        Get
            Dim objectsTreePanelPoint As Point = TLP_Objects.PointToScreen(New Point(0, 0))
            Dim gridPoint As Point = Script_Grid.PointToScreen(New Point(0, 0))
            Dim cellBorderCenter As Integer = 5
            Dim cellBorderWidth As Integer = cellBorderCenter * 2
            Dim separatorRectangle As New Rectangle({objectsTreePanelPoint.X + TLP_Objects.Width - cellBorderCenter, 0}.Max, gridPoint.Y, cellBorderWidth, Script_Grid.Height)
            Return separatorRectangle
        End Get
    End Property
    Private ReadOnly Property PaneGridSeparator As Rectangle
        Get
            Dim gridPoint As Point = Script_Grid.PointToScreen(New Point(0, 0))
            Dim cellBorderCenter As Integer = 5
            Dim cellBorderWidth As Integer = cellBorderCenter * 2
            Dim separatorRectangle As New Rectangle(gridPoint.X - cellBorderCenter, gridPoint.Y, cellBorderWidth, Script_Grid.Height)
            Return separatorRectangle
        End Get
    End Property
    Private Sub Panel_MouseMove(sender As Object, e As MouseEventArgs) Handles TLP_PaneGrid.MouseMove, AddTab.MouseMove, TLP_Objects.MouseMove

        If e.Button = MouseButtons.Left Or ForceCapture Then
            If SeparatorSizing = Sizing.MouseDownOPS Then
                TLP_PaneGrid.ColumnStyles(0).SizeType = SizeType.Absolute
                TLP_PaneGrid.ColumnStyles(0).Width = e.X
                ObjectsWidth = e.X
                RaiseEvent Alert(e.Location, New AlertEventArgs("Move: MouseDownOPS" & ObjectsPaneSeparator.ToString))

            ElseIf SeparatorSizing = Sizing.MouseDownPGS Then
                TLP_PaneGrid.ColumnStyles(1).SizeType = SizeType.Absolute
                TLP_PaneGrid.ColumnStyles(1).Width = e.X - TLP_PaneGrid.ColumnStyles(0).Width
                RaiseEvent Alert(e.Location, New AlertEventArgs("Move: MouseDownPGS" & PaneGridSeparator.ToString))

            Else
                RaiseEvent Alert(e.Location, New AlertEventArgs("Move: Nothing"))

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

        If 0 = 1 Then
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
        End If

    End Sub
    Private Sub Panel_MouseDown(sender As Object, e As MouseEventArgs) Handles TLP_PaneGrid.MouseDown, AddTab.MouseDown, TLP_Objects.MouseDown

        If e.Button = MouseButtons.Left Then
            If ObjectsPaneSeparator.Contains(Cursor.Position) Then
                SeparatorSizing = Sizing.MouseDownOPS
                RaiseEvent Alert(e.Location, New AlertEventArgs("Down: MouseDownOPS"))

            ElseIf PaneGridSeparator.Contains(Cursor.Position) Then
                SeparatorSizing = Sizing.MouseDownPGS
                RaiseEvent Alert(e.Location, New AlertEventArgs("Down: MouseDownPGS"))

            Else
                SeparatorSizing = Sizing.None
                RaiseEvent Alert(e.Location, New AlertEventArgs("Down: Sizing.None," & ObjectsPaneSeparator.ToString))

            End If
            ForceCapture = Not SeparatorSizing = Sizing.None

        ElseIf e.Button = MouseButtons.Right Then
            My.Settings.DontShowObjectsMessage = True
            My.Settings.Save()
        End If

    End Sub
    Private Sub Panel_DoubleClick(sender As Object, e As EventArgs) Handles TLP_PaneGrid.DoubleClick, AddTab.DoubleClick, TLP_Objects.DoubleClick

        If ObjectsPaneSeparator.Contains(Cursor.Position) Then
            TLP_PaneGrid.ColumnStyles(0).Width = ObjectsWidth

        ElseIf PaneGridSeparator.Contains(Cursor.Position) Then
            AutoWidth(ActivePane)

        End If

    End Sub
    Private Sub Panel_MouseUp(sender As Object, e As MouseEventArgs) Handles TLP_PaneGrid.MouseUp, AddTab.MouseUp, TLP_Objects.MouseUp

        ForceCapture = False
        If ObjectsPaneSeparator.Contains(e.Location) Then
            SeparatorSizing = Sizing.MouseOverOPS

        ElseIf PaneGridSeparator.Contains(e.Location) Then
            SeparatorSizing = Sizing.MouseOverPGS

        End If

    End Sub
    Private Sub Panel_Leave(sender As Object, e As EventArgs) Handles TLP_PaneGrid.Leave, Tree_Objects.Enter, Tree_Objects.MouseMove, AddTab.MouseLeave, TLP_Objects.MouseLeave

        If ForceCapture Then
        Else
            Cursor = Cursors.Default
        End If

    End Sub
    Private Sub ObjectsClose() Handles Button_ObjectsClose.Click
        TLP_PaneGrid.ColumnStyles(0).Width = 0
    End Sub
#End Region
#Region " Tabs Events "
    Private Sub Tab_Clicked(sender As Object, e As TabsEventArgs) Handles Script_Tabs.TabClicked

        Select Case e.InZone
            Case Tabs.Zone.Add
                Dim newScript As New Script
                Scripts.Add(newScript) 'CollectionChanged ... adds handlers
                newScript.State = Script.ViewState.OpenDraft'VisibleChanged ... adds tab + pane

            Case Tabs.Zone.Image
#Region " RUN "
                If ActiveScript.Connection Is Nothing Then
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
                Dim scriptActive As Script = ActiveScript()
                Dim activeText As String = scriptActive.Text
                Dim fileText As String = scriptActive.DSN_Body.Value
                With scriptActive
                    If .FileCreated Then
                        Dim textA As String = .Text
                        Dim textB As String = .DSN_Body.Value

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
    Private Sub Tabs_ZoneChange(sender As Object, e As TabsEventArgs) Handles Script_Tabs.ZoneMouseChange

        If Not e.InZone = Tabs.Zone.None Then
            Dim TipLocation = Script_Tabs.PointToScreen(If(e.InTab, e.OutTab).Bounds.Location)
            TipLocation.Offset(ActiveTab.Bounds.Width, 3)

            Select Case e.InZone
                Case Tabs.Zone.Add

                Case Tabs.Zone.Image
                    Dim TipValues As String = Nothing
                    With ActiveScript()
                        TipValues = "Run Script|" & Bulletize({"Current datasource is " & If(IsNothing(.Connection), "undetermined", .DataSourceName),
                                            "Type is " & If(.Body.InstructionType = ExecutionType.Null, "undetermined", .Body.InstructionType.ToString),
                                            Join({"Text has", If(.TextWasModified, String.Empty, " not"), " changed"}, String.Empty),
                                            Join({"Last modified", .Modified.ToShortDateString, "@", .Modified.ToShortTimeString}),
                                            Join({"Last successful run", .Ran.ToShortDateString, "@", .Ran.ToShortTimeString}),
                                            "Location=" & If(.Path, "None - not saved")})
                    End With

                Case Tabs.Zone.Text
                    'Tabs_TipManager("Reorder tab|Drag tab and drop in new position", TipLocation)

                Case Tabs.Zone.Close
                    If Not ActiveScript.Body.HasText Then
                        'Tabs_TipManager("Close Tab|Click to close empty tab", TipLocation)

                    ElseIf ActiveScript.FileTextMatchesText Then
                        'Tabs_TipManager("Close Tab|Click to close saved script", TipLocation)

                    Else

                    End If
            End Select
        End If

    End Sub
#End Region
#Region " Active Events "
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ ActiveScript
    Private Sub Scripts_CollectionChanged(sender As Object, e As ScriptsEventArgs) Handles Scripts.CollectionChanged

        If e.State = CollectionChangeAction.Refresh Then
            With FileTree
                .Nodes.SortOrder = SortOrder.Ascending
                .Nodes.Insert(0, OpenFileNode)
            End With

        ElseIf e.State = CollectionChangeAction.Add Then
            Dim addScript As Script = e.Item
            With FileTree
                Dim ConnectionName As String = If(addScript.Connection Is Nothing, "Undetermined", addScript.Connection.DataSource)
                Dim DatabaseColor As Color = If(addScript.Connection Is Nothing, Color.Blue, addScript.Connection.BackColor)
                Dim Database_Image As Image = ChangeImageColor(My.Resources.Sync, Color.FromArgb(255, 64, 64, 64), DatabaseColor)

                If Not .Nodes.Exists(Function(n) n.Name = ConnectionName) Then
                    .Nodes.Add(New Node With {
                            .Text = ConnectionName,
                            .Name = ConnectionName,
                            .Image = Database_Image,
                            .CanAdd = False,
                            .CanDragDrop = False,
                            .CanEdit = False,
                            .CanRemove = False,
                            .Separator = Node.SeparatorPosition.Above,
                            .Tag = addScript.Connection})
                End If
                Dim ConnectionNode As Node = .Nodes.Item(ConnectionName)
                If addScript.Connection Is Nothing Then ConnectionNode.BackColor = Color.FromArgb(128, Color.Gainsboro)
                ConnectionNode.Nodes.Add(New Node With {
                            .Text = addScript.Name,
                            .Name = addScript.Name,
                            .Image = addScript.TabImage,
                            .CanAdd = False,
                            .CanEdit = True,
                            .CanRemove = True,
                            .CanFavorite = True,
                            .Favorite = addScript.Favorite,
                            .Tag = addScript,
                            .CursorGlowColor = DatabaseColor
                                         })
                ConnectionNode.Nodes.SortOrder = SortOrder.Ascending
            End With
            AddHandler addScript.VisibleChanged, AddressOf Script_VisibleChanged
            AddHandler addScript.NameChanged, AddressOf Script_NameChanged
            AddHandler addScript.TypeChanged, AddressOf Script_TypeChanged
            AddHandler addScript.ConnectionChanged, AddressOf Script_ConnectionChanged
            AddHandler addScript.GenericEvent, AddressOf Script_GenericEvent

        ElseIf e.State = CollectionChangeAction.Remove Then
            Dim removeScript As Script = e.Item
            RemoveHandler removeScript.VisibleChanged, AddressOf Script_VisibleChanged
            RemoveHandler removeScript.NameChanged, AddressOf Script_NameChanged
            RemoveHandler removeScript.TypeChanged, AddressOf Script_TypeChanged
            RemoveHandler removeScript.ConnectionChanged, AddressOf Script_ConnectionChanged
            RemoveHandler removeScript.GenericEvent, AddressOf Script_GenericEvent
            Dim RemoveNode As Node = FileTree.Nodes.ItemByTag(removeScript)
            RemoveNode?.RemoveMe()

        End If

    End Sub
    Private Sub Script_VisibleChanged(sender As Object, e As ScriptVisibleChangedEventArgs)

        Dim scriptNode As Node = FileTree.Nodes.ItemByTag(e.Item)
        If e.Visible Then
            Dim visibleScript As Script = e.Item
            Dim newTab As New Tab With {
                .HeaderBackColor = If(visibleScript.Connection Is Nothing, Color.Gainsboro, visibleScript.Connection.BackColor),
                .HeaderForeColor = If(visibleScript.Connection Is Nothing, Color.Black, visibleScript.Connection.ForeColor),
                .ItemText = visibleScript.Name,
                .Image = visibleScript.TabImage,
                .Tag = visibleScript,
                .AllowDrop = True
            }
            Dim newPane As New RicherTextBox With {
                .Name = "Pane",
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .WordWrap = True,
                .AllowDrop = True,
                .AcceptsTab = True,
                .Font = My.Settings.paneFont,
                .Tag = visibleScript,
                .EnableAutoDragDrop = True,
                .Text = visibleScript.Text
            }
            newTab.Controls.Add(newPane)
            Script_Tabs.TabPages.Add(newTab)
            visibleScript.TabPage_ = newTab
            AddHandler newTab.TextChanged, AddressOf AutoWidth
            With newPane
                AddHandler .TextChanged, AddressOf ActivePane_TextChanged
                AddHandler .KeyDown, AddressOf ActivePane_KeyDown
                AddHandler .MouseDown, AddressOf ActivePane_MouseDown
                AddHandler .MouseEnter, AddressOf ActivePane_MouseEnter
                AddHandler .MouseMove, AddressOf ActivePane_MouseMove
                AddHandler .SelectionChanged, AddressOf ActivePane_SelectionChanged
                AddHandler .ScrolledVertical, AddressOf ActivePane_ScrolledVertical
                AddHandler .DragStart, AddressOf ActivePane_DragStart
                AddHandler .DragOver, AddressOf ActivePane_DragOver
                AddHandler .DragDrop, AddressOf ActivePane_DragDrop
            End With
            Dim NodeColor As Color = If(visibleScript.Connection Is Nothing, Color.Black, visibleScript.Connection.BackColor)
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
            scriptNode.Image = bmp

        Else
            Dim invisibleScript As Script = e.Item
            Dim oldTab As Tab = invisibleScript.TabPage
            RemoveHandler oldTab.TextChanged, AddressOf AutoWidth
            Dim oldPane As RicherTextBox = DirectCast(oldTab.Controls.Item("Pane"), RicherTextBox)
            If oldPane IsNot Nothing Then
                With oldPane
                    RemoveHandler .TextChanged, AddressOf ActivePane_TextChanged
                    RemoveHandler .KeyDown, AddressOf ActivePane_KeyDown
                    RemoveHandler .MouseDown, AddressOf ActivePane_MouseDown
                    RemoveHandler .MouseEnter, AddressOf ActivePane_MouseEnter
                    RemoveHandler .MouseMove, AddressOf ActivePane_MouseMove
                    RemoveHandler .SelectionChanged, AddressOf ActivePane_SelectionChanged
                    RemoveHandler .ScrolledVertical, AddressOf ActivePane_ScrolledVertical
                    RemoveHandler .DragStart, AddressOf ActivePane_DragStart
                    RemoveHandler .DragOver, AddressOf ActivePane_DragOver
                    RemoveHandler .DragDrop, AddressOf ActivePane_DragDrop
                End With
                invisibleScript.TabPage.Controls.Remove(oldPane)
                oldPane.Dispose()
            End If
            scriptNode.Image = e.Item.TabImage
            Script_Tabs.TabPages.Remove(oldTab)
            oldTab.Dispose()
        End If

    End Sub
    Private Sub Script_NameChanged(sender As Object, e As ScriptNameChangedEventArgs)

        Dim changedScript As Script = DirectCast(sender, Script)
        Dim changedNode As Node = FileTree.Nodes.ItemByTag(changedScript)
        changedNode.Text = e.CurrentName
        changedScript.TabPage.ItemText = changedScript.Name
        SaveAs.Text = changedScript.Name
        Script_Tabs.Refresh()

    End Sub
    Private Sub Script_TypeChanged(sender As Object, e As ScriptTypeChangedEventArgs)

        With DirectCast(sender, Script)
            If .TabPage IsNot Nothing Then 'If dragging a closed script ( no tab ) - this event can fire
                .TabPage.ItemText = .Name
                .TabPage.Image = .TabImage
            End If
            Script_Tabs.Refresh()
        End With

    End Sub
    Private Sub Script_ConnectionChanged(sender As Object, e As ConnectionChangedEventArgs)

        If e.NewConnection Is Nothing Then
            'Password change
            'AddHandler Connection.PasswordChanged, AddressOf ConnectionChanged *** USE???

        Else
            With e.NewConnection
                Dim scriptConnection As Script = DirectCast(sender, Script)
                If scriptConnection.TabPage IsNot Nothing Then 'If dragging a closed script ( no tab ) - this event can fire
                    scriptConnection.TabPage.HeaderBackColor = e.NewConnection.BackColor
                    scriptConnection.TabPage.HeaderForeColor = e.NewConnection.ForeColor
                End If
                Dim Message As String = "Currently connected to " & .DataSource
                If .Properties.ContainsKey("NICKNAME") Then Message &= Join({String.Empty, "(", .Properties("NICKNAME"), ")"})
                RaiseEvent Alert(e.NewConnection, New AlertEventArgs(Message))
            End With
        End If

    End Sub
    Private Sub Script_GenericEvent(sender As Object, e As AlertEventArgs)
        RaiseEvent Alert(sender, e)
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬ ActivePane
    Private Sub ActivePane_MouseEnter(sender As Object, e As EventArgs)

        FindAndReplace.Parent = DirectCast(sender, RicherTextBox)
        CMS_PaneOptions.Close()

    End Sub
    Private Sub ActivePane_TextChanged(sender As Object, e As EventArgs)

        ActiveScript.Text = ActivePane.Text
        SaveAs.Image = If(ActiveScript.FileTextMatchesText, My.Resources.saved, My.Resources.savedNot)

    End Sub
    Private Sub ActivePane_KeyDown(sender As Object, e As KeyEventArgs)

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
    Private Sub ActivePane_MouseDown(sender As Object, e As MouseEventArgs)

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
                                                   TSMI_Font,
                                                   TSMI_Tidy})
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
                                    .Mode = ImageComboMode.ColorPicker
                                End With
                                With IC_ForeColor
                                    .DropDown.CheckBoxes = False
                                    .Mode = ImageComboMode.ColorPicker
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
    Private Sub ActivePane_MouseMove(sender As Object, e As MouseEventArgs)

        Dim pane As RicherTextBox = ActivePane()
        With pane
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
                                            Location.Offset(pane.PointToClient(New Point(0, 0)))
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
                End If
            End If
        End With

    End Sub
    Private Sub ActivePane_SelectionChanged(sender As Object, e As EventArgs)

        With ActivePane()
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
    Private Sub ActivePane_ScrolledVertical(sender As Object, e As RicherEventArgs)

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
    Private Sub ActivePane_DragStart(sender As Object, e As DragEventArgs)
        Data.SetData(ActivePane.GetType, ActivePane)
    End Sub
    Private Sub ActivePane_DragOver(sender As Object, e As DragEventArgs)

        Dim Grid = Data.GetData(GetType(DataTool))
        If Grid IsNot Nothing Then
            TLP_PaneGrid.ColumnStyles(0).Width = ObjectsWidth
        End If

    End Sub
    Private Sub ActivePane_DragDrop(sender As Object, e As DragEventArgs)
        'A ClosedScript node was dropped on an active pane
        Pane_NodeDropped(e)
    End Sub
#End Region
#Region " ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■ RUN SCRIPT "
    Private Sub RunScript(Optional _Script As Script = Nothing)

        Script_Grid.Columns.CancelWorkers()
        _Script = If(_Script, ActiveScript())

        With _Script
            If .Body.HasText And _Script.Connection IsNot Nothing Then
                If .Connection.CanConnect Then
                    If .Type = ExecutionType.DDL Then
#Region " D D L "
                        Cursor.Current = Cursors.WaitCursor
                        Dim procedure As New DDL(.Connection, .Body.SystemText, My.Settings.ddlPrompt, My.Settings.ddlRowCount)
                        If procedure.ProceduresOK.Any Then
                            RaiseEvent Alert(_Script, New AlertEventArgs("Running procedure " & .Name))
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

                    ElseIf .Type = ExecutionType.SQL Then
#Region " S Q L "
                        RaiseEvent Alert(_Script, New AlertEventArgs("Running query " & .Name))
                        'https://www.ibm.com/support/knowledgecenter/SSEPEK_11.0.0/cattab/src/tpc/db2z_catalogtablesintro.html
                        If .Connection.IsFile Then
                            For Each SheetName In SystemObjects
                                'SQL_Statement = Replace(SQL_Statement, SheetName.Name, "[" & SheetName.Name & "]")
                            Next
                        Else
                            Dim body As BodyElements = .Body
                            Dim bodyNames As New List(Of String)(From te In body.TablesElement Select te.Name.ToUpperInvariant)
                            Dim systemNames As New List(Of String)(From so In SystemObjects Select so.Name.ToUpperInvariant)
                            Dim objectsNeed As New List(Of String)(bodyNames.Except(systemNames))
                            If objectsNeed.Any Then
                                RaiseEvent Alert(.Body, New AlertEventArgs("Adding to profile: " & Join(objectsNeed.ToArray, ",") & "-(RunQuery)"))
                                Dim tableColumnSQL As String = ColumnSQL(objectsNeed.ToArray, .Connection.Language)
                                With New SQL(.Connection, tableColumnSQL)
                                    AddHandler .Completed, AddressOf ColumnsSQL_Completed
                                    .Execute()
                                End With
                            End If
                            Dim whichText As String = If(0 = 0, .Text, .Body.SystemText)
                            With New SQL(.Connection, whichText)
                                AddHandler .Completed, AddressOf Execute_Completed
                                .Name = _Script.CreatedString
                                .Execute()
                            End With
                        End If
#End Region
                    ElseIf .Type = ExecutionType.Null Then

                    End If
                Else
                    Dim Items As New List(Of String)
                    If .Connection.MissingUserID Then Items.Add("userid")
                    If .Connection.MissingPassword Then Items.Add("password")
                    Message.Show("Can not connect", "Connection is missing " & Join(Items.ToArray, " and "), Prompt.IconOption.Critical, Prompt.StyleOption.Blue)
                End If
            Else
                Stop
                Message.Show("No datasource found or selected", "Please set your connection", Prompt.IconOption.Critical, Prompt.StyleOption.Blue)
            End If
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
        Script_Grid?.Timer?.StopTicking()
        Dim pane As RicherTextBox = ActivePane()

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

        Dim BulletMessage As String = e.Message
        Dim FlatMessage As String
        Dim CreatedDate As Date = StringToDateTime(ItemName)
        Dim _Script As Script = Scripts.Item(CreatedDate)

        Dim ToolTipTitle = Join({If(IsQuery, "Query", "Procedure"), _Script.Name, If(e.Succeeded, "succeeded", "failed")})

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
                TLP_PaneGrid.ColumnStyles(0).Width = 0
                Script_Grid?.Timer?.StartTicking(Color.LawnGreen)
                With New Worker
                    .Tag = e.Table
                    AddHandler .DoWork, AddressOf Async_StartDatasourceWidth
                    AddHandler .RunWorkerCompleted, AddressOf Async_EndDatasourceWidth
                    .RunWorkerAsync()
                End With
            Else
                BulletMessage = Bulletize({ElapsedMessage})

            End If

        Else
            Dim errorMatch As Match = Regex.Match(e.Message, "\(at char [0-9]{1,}\)", RegexOptions.None)
            If errorMatch.Success Then
                Dim errorPosition As Integer = CInt(Regex.Match(errorMatch.Value, "[0-9]{1,}", RegexOptions.None).Value) - 1 'at char xx is 1-based
                If pane IsNot Nothing Then
                    With pane
                        .SelectionStart = errorPosition
                        .SelectionLength = 1
                        .SelectionBackColor = Color.Red
                    End With
                End If
            End If
            BulletMessage = Bulletize({e.Message})
        End If
        FlatMessage = Join(Split(BulletMessage, "● ").Skip(1).ToArray, " ● ")
        RaiseEvent Alert(e, New AlertEventArgs(Join({ToolTipTitle, ":", FlatMessage})))

    End Sub
    Private Sub Async_StartDatasourceWidth(sender As Object, e As DoWorkEventArgs)
        With DirectCast(sender, Worker)
            RemoveHandler .DoWork, AddressOf Async_StartDatasourceWidth
            SetSafeControlPropertyValue(Script_Grid, "DataSource", DirectCast(.Tag, DataTable))
        End With
    End Sub
    Private Sub Async_EndDatasourceWidth(sender As Object, e As RunWorkerCompletedEventArgs)
        AutoWidth(Script_Grid)
    End Sub
    Private Sub ColumnsSQL_Completed(sender As Object, e As ResponseEventArgs)

        'Column SQL completed, Update Objects.txt + update ObjectsTree.Nodes???
        With DirectCast(sender, SQL)
            RemoveHandler .Completed, AddressOf ColumnsSQL_Completed
            Dim queryObjects As New List(Of SystemObject)(ColumnTypesToSystemObject(.Table))
            If queryObjects.Any Then
#Region " ADD TO C:\Users\SeanGlover\Documents\DataManager\Objects.txt "
                Dim commonObjects As New List(Of SystemObject)(SystemObjects.Intersect(queryObjects))
                For Each commonObject In commonObjects
                    SystemObjects.Remove(commonObject) 'Kepp info current by using latest pull
                Next
                SystemObjects.AddRange(queryObjects)
                SystemObjects.Save()
#End Region
#Region " ADD TO ObjectsTree - Show new connection + Owner + Table "
                If 0 = 1 Then 'Another day ... only reflects any new adds on ObjectTree
                    '                                           dsn                  owner                  name
                    Dim objectDictionaryX As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, List(Of ColumnProperties))))
                    For Each queryObject As SystemObject In queryObjects
                        Dim dsn As String = queryObject.DSN
                        Dim owner As String = queryObject.Owner
                        Dim objectName As String = queryObject.Name
                        For Each cp As ColumnProperties In queryObject.Columns.Values
                            If Not objectDictionaryX.ContainsKey(dsn) Then objectDictionaryX.Add(dsn, New Dictionary(Of String, Dictionary(Of String, List(Of ColumnProperties))))
                            If Not objectDictionaryX(dsn).ContainsKey(owner) Then objectDictionaryX(dsn).Add(owner, New Dictionary(Of String, List(Of ColumnProperties)))
                            If Not objectDictionaryX(dsn)(owner).ContainsKey(objectName) Then objectDictionaryX(dsn)(owner).Add(objectName, New List(Of ColumnProperties))
                            objectDictionaryX(dsn)(owner)(objectName).Add(cp)
                        Next
                    Next
                    For Each dsn In objectDictionaryX
                        Dim sourceNode As Node = Tree_Objects.Nodes.Item(dsn.Key) 'Assumes exists
                        Dim _Connection As Connection = DirectCast(sourceNode.Tag, Connection)

                        For Each owner In dsn.Value
                            Dim ownerNode As Node = sourceNode.Nodes.Item(owner.Key)
                            If ownerNode Is Nothing Then ownerNode = sourceNode.Nodes.Add(New Node With {
                                        .Text = owner.Key,
                                        .Name = owner.Key,
                                        .BackColor = If(owner.Key = _Connection.UserID, Color.Gainsboro, Color.Transparent),
                                        .CanAdd = False,
                                        .CanDragDrop = False,
                                        .CanEdit = False,
                                        .CanRemove = False
                                        })

                            For Each objectName In owner.Value
                                Dim nameNode As Node = ownerNode.Nodes.Item(objectName.Key)
                                If nameNode Is Nothing Then nameNode = ownerNode.Nodes.Add(objectName.Key, objectName.Key)



                                For Each column As ColumnProperties In objectName.Value
                                    Dim columnNode As Node = nameNode.Nodes.Item(column.Name)
                                    If columnNode Is Nothing Then columnNode = nameNode.Nodes.Add(New Node With {
                                        .Text = column.Name,
                                        .Name = column.Name,
                                        .CanAdd = False,
                                        .CanDragDrop = False,
                                        .CheckBox = False
                                })
                                Next
                            Next
                        Next
                    Next
                End If
#End Region
            Else
            End If
        End With

    End Sub
#End Region
#Region " FileTree EVENTS "
    Private Sub ClosedScript_NodeDragOver(sender As Object, e As DragEventArgs) Handles Script_Tabs.DragOver, Script_Grid.DragOver

        Dim DragNode As Node = DirectCast(e.Data.GetData(GetType(Node)), Node)
        If DragNode IsNot Nothing Then
            If DragNode.CanDragDrop Then
                e.Effect = DragDropEffects.All
            Else
                e.Effect = DragDropEffects.None
            End If
        End If

    End Sub
    Private Sub ClosedScript_NodeDroppedTabs(sender As Object, e As DragEventArgs) Handles Script_Tabs.DragDrop
        Pane_NodeDropped(e)
    End Sub
#End Region
#Region " SCRIPT CONTROL EVENTS "
    Private Sub FindRequest(sender As Object, e As ZoneEventArgs) Handles FindAndReplace.ZoneClicked

        Dim Text_Search As String = ActivePane.Text
        Select Case e.Zone.Name
            Case Zone.Identifier.MatchCase, Zone.Identifier.MatchWord, Zone.Identifier.RegEx
                FindRequest()

            Case Zone.Identifier.Close
                'Remove the Highlights
                With ActivePane()
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
                    With ActivePane()
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

        If FindAndReplace.FindControl?.Text.Any Then
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
        Else
            With ActivePane()
                Dim _SelectionStart As Integer = .SelectionStart
                .SelectAll()
                .SelectionBackColor = Color.Transparent
                .SelectionColor = Color.Black
                .SelectionStart = _SelectionStart
                .SelectionLength = 0
            End With
        End If

    End Sub
    Private Sub InsertComment(sender As Object, e As EventArgs) Handles TSMI_Comment.Click

        If ActiveBody.HasText Then
            'RemoveHandler ActivePane.SelectionChanged, AddressOf ActivePaneSelectionChanged
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
            'AddHandler ActivePane.SelectionChanged, AddressOf ActivePaneSelectionChanged
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
        With ActivePane()
            Dim CharIndex = .SelectionStart
            Dim LineNbr = .GetLineFromCharIndex(CharIndex)
            Dim LineStart = .GetFirstCharIndexFromLine(LineNbr)
            .Text = .Text.Insert(LineStart, Separator)
            .SelectionStart = CharIndex
        End With

    End Sub
    Private Function GetCommentMatch() As StringData

        With ActivePane()

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
                        '.TableFloating_Back = NewColor
                    Else
                        '.TableFloating_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.RoutineTable
                    If ChangedBackColor Then
                        '.TableRoutine_Back = NewColor
                    Else
                        '.TableRoutine_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.SystemTable
                    If ChangedBackColor Then
                        '.TableSystem_Back = NewColor
                    Else
                        '.TableSystem_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.GroupBlock, InstructionElement.LabelName.GroupField
                    If ChangedBackColor Then
                        '.Group_Back = NewColor
                    Else
                        '.Group_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.Limit
                    If ChangedBackColor Then
                        '.Limit_Back = NewColor
                    Else
                        '.Limit_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.OrderBlock, InstructionElement.LabelName.OrderField
                    If ChangedBackColor Then
                        '.Order_Back = NewColor
                    Else
                        '.Order_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.SelectBlock, InstructionElement.LabelName.SelectField
                    If ChangedBackColor Then
                        '.Select_Back = NewColor
                    Else
                        '.Select_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.WithBlock
                    If ChangedBackColor Then
                    Else
                        '.WithBlock_Fore = NewColor
                    End If
                Case InstructionElement.LabelName.Union
                    If ChangedBackColor Then
                        '.Union_Back = NewColor
                    Else
                        '.Union_Fore = NewColor
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
            CMS_PaneOptions.AutoClose = False
            CMS_PaneOptions.Show(Location)
        Else
            CMS_PaneOptions.AutoClose = True
            CMS_PaneOptions.Hide()
        End If

    End Sub
    Private Function LightSwitchedOn() As Boolean

        Return SameImage(My.Resources.LightOn, TSMI_TipSwitch.Image)
        Return SameImage(Base64ToImage(LightOn), TSMI_TipSwitch.Image)

    End Function
    '===============================================================================

#Region " TIDY TEXT + SUPPORTING DECLARATIONS/FUNCTIONS "
    Private SelectionIndex As Integer
    Public Sub TidyText()

        With ActivePane()
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

#Region " OBJECT EVENTS "
#Region " Tree_Objects POPULATION "
    Private ReadOnly Property SelectedConnections As List(Of Connection)
        Get
            Return (From n In Tree_Objects.SelectedNodes Where n Is n.Root Select DirectCast(n.Tag, Connection)).ToList
        End Get
    End Property
    Private RequestInitiated As Boolean
    Private Sub ObjectSyncClicked(sender As Object, e As EventArgs) Handles Button_ObjectsSync.Click

        ' *** Correct any discrepancies between SystemObjects and Database ***
        'SelectedConnections ie) User decides which items to update ( NOT USED )
        If Not RequestInitiated Then
            RequestInitiated = True
            If ObjectsSet.Tables.Count = 0 And Not ObjectsWorker.IsBusy Then
                Script_Grid.Timer.Picture = WaitTimer.ImageType.Spin
                Script_Grid.Timer.StartTicking()
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
                        Dim DB_Alias As String = DataSource.Server
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
            Script_Grid.Timer.StopTicking()
            Script_Grid.Timer.Picture = WaitTimer.ImageType.Circle
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
                SystemObjects.Save()
                For Each Level1Node In Tree_Objects.Nodes
                    Level1Node.Nodes.Clear()
                Next
                RequestInitiated = False
                ObjectsWorker.RunWorkerAsync()
            End Using
        End If

    End Sub
    Private Sub LoadSystemObjects(sender As Object, e As EventArgs) Handles ObjectsWorker.DoWork

        Dim ClockLoadTime As Boolean = False

        Dim LoadFromSettings As Boolean = sender Is Nothing
        Dim ConnectionsDictionary As New Dictionary(Of String, Boolean)

        If ClockLoadTime Then Stop_Watch.Start()
        ExpandCollapseOnOff(HandlerAction.Remove)
#Region " FILL TABLE WITH DATABASE OBJECTS "
        Dim ActiveConnections = Connections.Where(Function(c) c.CanConnect).Take(1000)
        If SelectedConnections IsNot Nothing AndAlso SelectedConnections.Any Then ActiveConnections = ActiveConnections.Where(Function(c) SelectedConnections.Contains(c))
        Dim SuccessCount As Integer = 0
        For Each Connection In ActiveConnections
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
                                                    .CanAdd = False,
                                                    .CanDragDrop = False,
                                                    .CanEdit = False,
                                                    .CanRemove = False})
            End If
            If ClockLoadTime Then
                Intervals.Add(Connection.DataSource, Stop_Watch.Elapsed)
                Stop_Watch.Restart()
            End If
        Next
#End Region
#Region " ODBC.txt - ALIAS CDNIW, TargetDatabase = TORDSNQ "
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
                Dim SourceNode = .Nodes.Item(DataSource)
                SourceNode.Name = DataSource
                Dim _Connection As Connection = DirectCast(SourceNode.Tag, Connection)
                Dim Owners = ObjectsDictionary(DataSource)
                For Each Owner In Owners
                    Dim OwnerNode = SourceNode.Nodes.Add(New Node With {
                                        .Text = Owner.Key,
                                        .Name = Owner.Key,
                                        .BackColor = If(Owner.Key = _Connection.UserID, Color.Gainsboro, Color.Transparent),
                                        .CanAdd = False,
                                        .CanDragDrop = False,
                                        .CanEdit = False,
                                        .CanRemove = False
                                        })
                    For Each ObjectType In Owner.Value
                        Dim TypeImage As Image = Nothing
                        If ObjectType.Key = SystemObject.ObjectType.Routine Then TypeImage = My.Resources.Gear
                        If ObjectType.Key = SystemObject.ObjectType.Table Then TypeImage = My.Resources.Table
                        If ObjectType.Key = SystemObject.ObjectType.Trigger Then TypeImage = My.Resources.Zap
                        If ObjectType.Key = SystemObject.ObjectType.View Then TypeImage = My.Resources.Eye
                        For Each Item In ObjectType.Value
                            Dim NameNode As Node = OwnerNode.Nodes.Add(New Node With {
                                        .Name = Item.Name,
                                        .Text = Item.Name,
                                        .Image = TypeImage,
                                        .Checked = True,
                                        .Tag = Item,
                                        .CanAdd = False,
                                        .CanDragDrop = True,
                                        .CanFavorite = True,
                                        .Favorite = Item.Favorite
                                        })
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
        Script_Grid.Timer.StopTicking()
        Script_Grid.Timer.Picture = WaitTimer.ImageType.Circle
#End Region

    End Sub
    Private Sub ObjectTreeviewLoaded() Handles ObjectsWorker.RunWorkerCompleted
        ObjectsTreeview_AutoWidth(Nothing, Nothing)
        ExpandCollapseOnOff(HandlerAction.Add)
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub ObjectsTreeview_NodeFavorited(sender As Object, e As NodeEventArgs) Handles Tree_Objects.NodeFavorited
        DirectCast(e.Node.Tag, SystemObject).Favorite = e.Node.Favorite
    End Sub
#End Region

#Region " Tree_Objects DRAG ONTO PANE Or GRID [ V I E W   S T R U C T U R E   Or   C O N T E N T ] "
    Private Sub ObjectNode_NodeDragStart(sender As Object, e As NodeEventArgs) Handles Tree_Objects.NodeDragStart

        Dim NodeObject = NodeProperties(e.Node)
        Dim Pane As RicherTextBox = ActivePane()
        If Pane IsNot Nothing Then Pane.AllowDrop = True
        Script_Grid.AllowDrop = True
        Select Case NodeObject.Type
            Case SystemObject.ObjectType.Table, SystemObject.ObjectType.View
                'Pane shows Table/View structure while Grid shows Content
                'Initiate threads for each so when dropped it's done
                Script_Grid.Timer.Picture = WaitTimer.ImageType.Spin
                Script_Grid.Timer.StartTicking()
                Dim SQL_Sample As String = Join({"SELECT *", "FROM " & NodeObject.FullName, "FETCH FIRST 50 ROWS ONLY"}, vbNewLine)
                Dim SQL_Structure As String = ColumnSQL(NodeObject.FullName)
                With Jobs
                    .Clear()
                    .Add(New Job(New SQL(NodeObject.Connection, SQL_Sample) With {
                                 .Name = e.Node.Text,
                                 .Tag = e.Node
                                 }) With {
                                 .Name = "50 Row Sample"})
                    .Add(New Job(New SQL(NodeObject.Connection, SQL_Structure) With {
                                 .Name = e.Node.Text,
                                 .Tag = e.Node
                                 }) With {.Name = "Table Structure"})
                    AddHandler .Completed, AddressOf ContentAndStructure_Completed
                    .Execute()
                End With

            Case SystemObject.ObjectType.Routine

            Case SystemObject.ObjectType.Trigger

        End Select

    End Sub
    Private Sub Pane_NodeDropped(e As DragEventArgs)

        Dim DroppedNode As Node = DirectCast(e.Data.GetData(GetType(Node)), Node)
        If DroppedNode IsNot Nothing Then
            With DroppedNode
                If .TreeViewer Is Tree_Objects Then
                    AutoWidth(ActivePane)

                ElseIf .TreeViewer Is FileTree Then
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
        FilesButton.HideDropDown()

    End Sub
    Private Sub Grid_NodeDropped(sender As Object, e As DragEventArgs) Handles Script_Grid.DragDrop

        Dim DroppedNode As Node = DirectCast(e.Data.GetData(GetType(Node)), Node)
        With DroppedNode
            If .TreeViewer Is Tree_Objects Then

            ElseIf .TreeViewer Is FileTree Then
                Dim _Script As Script = DirectCast(.Tag, Script)
                RunScript(_Script)

            End If
        End With
        FilesButton.HideDropDown()

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
    Private Sub ContentAndStructure_Completed(sender As Object, e As ResponsesEventArgs)

        Dim timeString As String = StrDup(10, BlackOut) & StrDup(5, " ") & Now.ToString("f", InvariantCulture) & StrDup(5, " ") & StrDup(10, BlackOut) & vbNewLine
        With DirectCast(sender, JobCollection)
            RemoveHandler .Completed, AddressOf ContentAndStructure_Completed
            Dim messages As New List(Of String)
            'table content
            Dim contentJob As Job = .Item("50 Row Sample")
            If contentJob.Succeeded Then
                Script_Grid.DataSource = contentJob.SQL.Table
            Else
                'Select * From <TableName> throws an error in DB2 if the requested table does not exist
                messages.Add(timeString & contentJob.SQL.Response.Message)
                'Remove the item from the ObjectsTree + SystemObjects
                Dim errorNode As Node = DirectCast(contentJob.SQL.Tag, Node)
                Dim errorObject As SystemObject = DirectCast(errorNode.Tag, SystemObject)
                errorNode.RemoveMe()
                errorObject.RemoveMe()
                SystemObjects.Save()
            End If
            'table structure
            Dim structureJob As Job = .Item("Table Structure")
            If structureJob.Succeeded Then
                'structureJob results from dragging < a > Table over to the Dataviewer ie) One Table only 
                Dim structureObjects As New List(Of SystemObject)(ColumnTypesToSystemObject(structureJob.SQL.Table))
                If structureObjects.Any Then
                    Dim structureObject = structureObjects.First
                    Dim structureColumns As New List(Of ColumnProperties)(structureObject.Columns.Values)
                    If structureColumns.Any Then
                        Dim dropCreate As New List(Of String) From {
                            Join({"DROP", structureObject.Type.ToString.ToUpperInvariant, structureObject.FullName}),
                            Join({"; CREATE", structureObject.Type.ToString.ToUpperInvariant, structureObject.FullName, "("})
                        }
                        For Each cp In structureColumns
                            Dim Line As String = Join({cp.Name, cp.Format}, StrDup(4, vbTab))
                            If cp.Index = 1 Then
                                dropCreate.Add(Line)
                            Else
                                dropCreate.Add(", " & Line)
                            End If
                        Next
                        dropCreate.Add(") IN " & structureObject.TSName)
                        messages.Add(timeString & CreateTableText(Join(dropCreate.ToArray, vbNewLine)))
                    End If
                End If
            Else
                'Select From SysTables Where Name=<'TableName'> will only throw an error on a timeout or connection issue 
                messages.Add(timeString & contentJob.SQL.Response.Message)
            End If

            If messages.Any Then
                MessageButton.Image = My.Resources.message_unread
                Dim priorText As String = MessageRicherBox.Text
                Dim messageText As String = Join(messages.ToArray, vbNewLine & StrDup(20, "-") & vbNewLine)
                With MessageRicherBox
                    .Text = Join({messageText, priorText}, vbNewLine)
                    .SelectAll()
                    .SelectionFont = New Font("IBM Plex Mono Light", 9, FontStyle.Regular)
                    .SelectionStart = 0
                    .SelectionLength = 0
                End With
            Else
                MessageButton.Image = My.Resources.message
            End If
        End With
        Script_Grid.Timer.StopTicking()
        Script_Grid.Timer.Picture = WaitTimer.ImageType.Circle

    End Sub
#End Region
    Private Sub MessageButton_Opening(sender As Object, e As EventArgs) Handles MessageButton.DropDownOpening
        MessageButton.Image = My.Resources.message
    End Sub
#Region " ActivePane.Text Or Script_Tabs.Tab Or Script_Grid DRAG ONTO Tree_Objects [ E T L ] "
    Private ReadOnly Data As New DataObject
    Private Sub Tab_StartDrag(sender As Object, e As TabsEventArgs) Handles Script_Tabs.TabDragDrop
        Data.SetData(Script_Tabs.GetType, Script_Tabs)
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
        ElseIf DragNode.TreeViewer Is FileTree Then
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
    Private Sub ObjectsSearch_TextCleared(sender As Object, e As EventArgs) Handles IC_ObjectsSearch.ClearTextClicked
        ObjectsTree_TransparentBackColor()
    End Sub
    Private Sub ObjectTreeView_Search(sender As Object, e As ImageComboEventArgs) Handles IC_ObjectsSearch.ValueSubmitted

        With New Worker
            AddHandler .DoWork, AddressOf ObjectSearch_Start
            AddHandler .RunWorkerCompleted, AddressOf ObjectSearch_End
            .RunWorkerAsync()
        End With

    End Sub
    Private Sub ObjectSearch_Start(sender As Object, e As DoWorkEventArgs)

        RemoveHandler DirectCast(sender, Worker).DoWork, AddressOf ObjectSearch_Start
        ObjectsTree_TransparentBackColor()
        For Each node In Tree_Objects.Nodes.All
            If node.Level = 2 And node.Text.ToUpperInvariant.Contains(IC_ObjectsSearch.Text.ToUpperInvariant) Then
                node.BackColor = Color.Yellow
                node.Parent.BackColor = Color.Yellow
                node.Root.BackColor = Color.Yellow
            End If
        Next

    End Sub
    Private Sub ObjectsTree_TransparentBackColor()
        For Each node In Tree_Objects.Nodes.All
            node.Root.BackColor = Color.Transparent
        Next
    End Sub
    Private Sub ObjectSearch_End(sender As Object, e As RunWorkerCompletedEventArgs)
        RemoveHandler DirectCast(sender, Worker).RunWorkerCompleted, AddressOf ObjectSearch_End
        Tree_Objects.Refresh()
    End Sub
#End Region

#Region " MAKE CHANGES TO SYSTEMOBJECTS FILE - 2 SOURCES {BULK IMPORT Or RunQuery RESULTS} "
    Private Sub ObjectNode_Checked(sender As Object, e As NodeEventArgs) Handles Tree_Objects.NodeChecked

        'Root=DataSource, Level 1=Owner, Level 2=Type {Trigger, Table, View, Routine}, Level 3=Name

        Dim BaseNodes = Tree_Objects.Nodes.All.Where(Function(n) n.Level = 2 And n.Checked)
        Dim CheckedObjects As New List(Of SystemObject)(BaseNodes.Select(Function(n) DirectCast(n.Tag, SystemObject)))
        Dim CheckedStrings As String() = (From CO In CheckedObjects Select CO.ToString & String.Empty).ToArray
        Dim CheckedString As String = Join(CheckedStrings, vbNewLine)

        Dim MySettingsObjects = SystemObjects.ToStringArray

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
            ObjectsWidth = 3 + .UnRestrictedSize.Width + 3
            TLP_PaneGrid.ColumnStyles(0).Width = ObjectsWidth
        End With

    End Sub
#End Region

#Region " OPEN FILE "
    Private Sub OpenFileClosed(sender As Object, e As EventArgs) Handles OpenFile.FileOk

        Dim _FileType As ExtensionNames = GetFileNameExtension(OpenFile.FileName).Value
        Dim SQL_Statement As String = String.Empty
        If _FileType = ExtensionNames.Excel Then
            Dim Sheets As New List(Of String)(ExcelSheetNames(OpenFile.FileName))
            If Sheets.Count = 1 Then
                SQL_Statement = "Select * FROM [" & Sheets.First & "]"
            Else
                CreateSheetList(Sheets)
                Exit Sub
            End If

        ElseIf _FileType = ExtensionNames.Text Then
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
#End Region

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
    Private Sub TSMI_TidyClicked() Handles TSMI_Tidy.Click
        TidyText()
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

            My.Settings.paneFont = .Font
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
        Using RTB As New RichTextBox With {.Font = My.Settings.paneFont, .Width = 2000, .Text = InputString}
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
    Private Sub AutoWidth(sender As Object, Optional e As EventArgs = Nothing) Handles Script_Grid.ColumnsSized

        'Optional e As Object = Nothing .... allows for other EventArgs to come here for resizing
        'Must use ActivePane if sender is a Worker, otherwise the Get throws a cross-thread error
        Script_Grid?.Timer?.StopTicking()
        If TLP_PaneGrid.ColumnStyles.Count >= 2 Then

            Dim Column_1_Width As Integer = 0
            Dim Column_2_Width As Integer = 0

            Dim Column1Percent As Integer = 50
            Dim ColumnToPercent As Integer = 100 - Column1Percent

            Dim AutoSize As Boolean = False

            If ActivePane Is Nothing Then
                'Only DummyTab showing...Leave Column1Percent @ 50
            Else
                If Not ActivePane.Text.Any And Script_Grid.DataSource Is Nothing Then
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
                    Dim Grid_BestWidth As Integer = Script_Grid.Width 'Script_Grid.TotalSize.Width
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
                        SaveAs.Text = DirectCast(sender, Tab).Name
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