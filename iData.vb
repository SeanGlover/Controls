Option Explicit On
Option Strict On
Imports System.IO
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Data.SqlClient
Imports Controls

#Region " RESPONSE EVENTS "
Public Structure Errors
    Implements IEquatable(Of Errors)
    Public Enum Type
        None
        Combatibility
        Requirement
        Access
        RunTime
    End Enum
    Public Enum BitVersion
        OK
        ArchitectureMismatch
    End Enum
    Public Enum WindowsPermission
        OK
        NotRunningAsAdministrator
    End Enum
    Public Enum AccessResponse
        OK
        Revoked
        PasswordMissing
        PasswordExpired
        PasswordIncorrect
    End Enum
    Public Enum ErrorSource
        None
        Timeout
        Script
    End Enum
    Public ReadOnly Property ErrorType As Type
    Public ReadOnly Property Compatibility As BitVersion
    Public ReadOnly Property Requirement As WindowsPermission
    Public ReadOnly Property Access As AccessResponse
    Public ReadOnly Property RunTime As ErrorSource
    Public Sub New(ExceptionMessage As String)

        If ExceptionMessage IsNot Nothing Then
            If ExceptionMessage.ToUpperInvariant.Contains("PASSWORD") Or ExceptionMessage.Contains("LDAP authentication failed For user") Then
#Region " ACCESS "
                _ErrorType = Type.Access
                If ExceptionMessage.ToUpperInvariant.Contains("EXPIRED") Then
                    _Access = AccessResponse.PasswordExpired

                ElseIf ExceptionMessage.ToUpperInvariant.Contains("MISSING") Then
                    _Access = AccessResponse.PasswordMissing

                ElseIf ExceptionMessage.ToUpperInvariant.Contains("PASSWORD INVALID") Or ExceptionMessage.Contains("LDAP authentication failed For user") Then
                    '[IBM][CLI DRIVER] SQL30082N  SECURITY PROCESSING FAILED WITH REASON "24" ("USERNAME AND/OR PASSWORD INVALID").  SQLSTATE=08001
                    _Access = AccessResponse.PasswordIncorrect

                ElseIf ExceptionMessage.ToUpperInvariant.Contains("REVOKED") Then
                    _Access = AccessResponse.Revoked

                End If
#End Region
            ElseIf ExceptionMessage.ToUpperInvariant.Contains("ARCHITECTURE") Then
#Region " COMPATIBILITY "
                _ErrorType = Type.Combatibility
                _Compatibility = BitVersion.ArchitectureMismatch
#End Region
            ElseIf ExceptionMessage.Contains("SQLAllocHandle") Then
#Region " REQUIREMENT "
                ErrorType = Type.Requirement
                _Requirement = WindowsPermission.NotRunningAsAdministrator
#End Region
            Else
#Region " RUNTIME "
                ErrorType = Type.RunTime
                If Regex.Match(ExceptionMessage, "time[d]{0,1}[ ]{0,1}out", RegexOptions.IgnoreCase).Success Then
                    _RunTime = ErrorSource.Timeout
                Else
                    _RunTime = ErrorSource.Script
                End If
#End Region
            End If
        End If

    End Sub
    Public Overrides Function GetHashCode() As Integer
        Return ErrorType.GetHashCode Xor Compatibility.GetHashCode Xor Requirement.GetHashCode Xor Access.GetHashCode Xor RunTime.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As Errors) As Boolean Implements IEquatable(Of Errors).Equals
        Return ErrorType = other.ErrorType AndAlso Compatibility = other.Compatibility
    End Function
    Public Shared Operator =(ByVal value1 As Errors, ByVal value2 As Errors) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(ByVal value1 As Errors, ByVal value2 As Errors) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is Errors Then
            Return CType(obj, Errors) = Me
        Else
            Return False
        End If
    End Function
End Structure
Public NotInheritable Class ResponsesEventArgs
    Inherits EventArgs
    Public ReadOnly Property Responses As List(Of ResponseEventArgs)
    Public Sub New(Responses As List(Of ResponseEventArgs))
        Me.Responses = Responses
    End Sub
    Public Sub New(Response As ResponseEventArgs)
        Responses = {Response}.ToList
    End Sub
    Public Overrides Function ToString() As String
        Return Join((From r In Responses Select If(r.Message, "OK")).ToArray, vbNewLine)
    End Function
End Class
Public NotInheritable Class ResponseEventArgs
    Inherits EventArgs
    Public ReadOnly Property Table As DataTable
    Public ReadOnly Property ConnectionString As String
    Public ReadOnly Property Connection As Connection
    Public ReadOnly Property ElapsedTime As TimeSpan
    Public ReadOnly Property Statement As String
    Public ReadOnly Property Message As String
    Public ReadOnly Property Columns As Integer
    Public ReadOnly Property Rows As Integer
    Public ReadOnly Property Value As Object
    Public ReadOnly Property RequestType As InstructionType
    Public ReadOnly Property RunError As Errors
    Public ReadOnly Property Succeeded As Boolean
    Public Sub New(RequestType As InstructionType, ConnectionString As String, Statement As String, ResultTable As DataTable, ElapsedTime As TimeSpan)
        REM /// SUCCEEDED
        Me.RequestType = RequestType
        Table = ResultTable
        If ResultTable Is Nothing Then
            Columns = 0
            Rows = 0
        Else
            Columns = ResultTable.Columns.Count
            Rows = ResultTable.Rows.Count
            Value = If(Columns = 1 And Rows = 1, ResultTable.Rows(0).Item(0), Nothing)
        End If
        Me.ElapsedTime = ElapsedTime
        Dim Connections = New ConnectionCollection
        Connection = Connections.Item(ConnectionString)
        Me.Statement = Statement
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Message = Nothing
        RunError = Nothing
        Succeeded = True
    End Sub
    Public Sub New(RequestType As InstructionType, ConnectionString As String, Statement As String, Message As String, RunError As Errors)
        REM /// FAILED
        Me.RequestType = RequestType
        Me.ConnectionString = ConnectionString
        Dim Connections = New ConnectionCollection
        Connection = Connections.Item(ConnectionString)
        Me.Statement = Statement
        Me.Message = Message
        Me.RunError = RunError
        '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
        Table = Nothing
        ElapsedTime = Nothing
        Columns = 0
        Rows = 0
        Succeeded = False
    End Sub
    Public Overrides Function ToString() As String
        Return If(Message, "OK")
    End Function
End Class
Friend Class ResponseFailure
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
            Handle.Dispose()
            ' Free any other managed objects here.
            Message.Dispose()
            IssuePrompt.Dispose()
            IssueMessage.Dispose()
            Password_New.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Private ReadOnly Message As New Prompt
    Private ReadOnly IssuePrompt As New ToolStripDropDown With {.AutoClose = False, .Margin = New Padding(0), .DropShadowEnabled = False, .BackColor = Color.Transparent}
    Private WithEvents IssueClose As New Button With {.Dock = DockStyle.Fill, .Image = My.Resources.Close.ToBitmap, .ImageAlign = ContentAlignment.MiddleLeft, .Margin = New Padding(0), .TextImageRelation = TextImageRelation.ImageBeforeText, .Text = "Expired Password..."}
    Private WithEvents IssueMessage As New Button With {.Dock = DockStyle.Fill, .Margin = New Padding(0), .Image = My.Resources.Warning.ToBitmap, .ImageAlign = ContentAlignment.MiddleLeft, .TextImageRelation = TextImageRelation.ImageBeforeText, .FlatStyle = FlatStyle.Standard, .Text = "Your password for {Database} has expired. You must create a new password now."}
    Private WithEvents Password_New As New ImageCombo With {.Dock = DockStyle.Fill, .Margin = New Padding(0), .HintText = "Enter your new password", .Text = String.Empty, .Image = My.Resources.Password}
    Friend Sub New()
    End Sub
    Friend Sub New(Response As ResponseEventArgs)
        Query_Procedure_Failed(Response)
    End Sub
    Private Sub Query_Procedure_Failed(e As ResponseEventArgs)

        Dim errorConnection As New Connection(e.ConnectionString)
        Dim errorDatasource As String = errorConnection.DataSource

        If e.RunError.ErrorType = Errors.Type.Access Then

#Region " PASSWORD ISSUE "
            Dim TLP_IP As TableLayoutPanel
            Dim PT As String = String.Empty
            Dim ShowPasswordBox As Boolean = True

            Password_New.Tag = e

            Select Case e.RunError.Access
                Case Errors.AccessResponse.PasswordExpired
                    With IssueClose
                        .Text = "Password expired!".ToString(InvariantCulture)
                        .BackColor = Color.Yellow
                        .ForeColor = Color.Black
                    End With
                    PT = "Your password {Password} for {Database} has expired. You must create a new password now.".ToString(InvariantCulture)

                Case Errors.AccessResponse.PasswordIncorrect
                    With IssueClose
                        .Text = "Password incorrect!".ToString(InvariantCulture)
                        .BackColor = Color.Orange
                        .ForeColor = Color.White
                    End With
                    PT = "Your password {Password} for {Database} is incorrect. You must submit another value.".ToString(InvariantCulture)

                Case Errors.AccessResponse.PasswordMissing
                    With IssueClose
                        .Text = "Password missing!".ToString(InvariantCulture)
                        .BackColor = Color.Orange
                        .ForeColor = Color.White
                    End With
                    PT = "Your password {Password} for {Database} is missing. You must provide a value.".ToString(InvariantCulture)

                Case Errors.AccessResponse.Revoked
                    With IssueClose
                        .Text = "Access revoked!".ToString(InvariantCulture)
                        .BackColor = Color.Red
                        .ForeColor = Color.White
                    End With
                    PT = "Your access to {Database} has been revoked! Submit a request to reinstate your access".ToString(InvariantCulture)
                    ShowPasswordBox = False

            End Select

            Dim SingleRowSize As Size = TextRenderer.MeasureText(PT, Segoe, New Size(300, 600), TextFormatFlags.Left)
            Dim MultiRowSize As Size = TextRenderer.MeasureText(PT, Segoe, New Size(300, 600), TextFormatFlags.WordBreak)
            Dim MessageLineCount As Integer = Convert.ToInt32(MultiRowSize.Height / SingleRowSize.Height)

            TLP_IP = New TableLayoutPanel With {.Width = 300, .ColumnCount = 1, .RowCount = If(ShowPasswordBox, 3, 2), .Font = Segoe, .AutoSize = True, .BorderStyle = BorderStyle.None, .CellBorderStyle = TableLayoutPanelCellBorderStyle.None, .Margin = New Padding(0)}
            With TLP_IP
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = TLP_IP.Width})
                REM /// CLOSE PROMPT ROW
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36})
                REM /// MESSAGE ROW
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36 * MessageLineCount})

                .Controls.Add(IssueClose, 0, 0)
                .Controls.Add(IssueMessage, 0, 1)

                If ShowPasswordBox Then
                    .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36})
                    .Controls.Add(Password_New, 0, 2)
                End If
                TLP.SetSize(TLP_IP)
            End With

            With IssuePrompt
                .Items.Clear()
                .Items.Add(New ToolStripControlHost(TLP_IP))
            End With

            With e.Connection
                If e.Connection IsNot Nothing Then
                    IssueMessage.Text = Replace(PT, "{Database}", .DataSource)
                    IssueMessage.Text = Replace(IssueMessage.Text, "{Password}", .Password)
                    IssuePromptShow()
                End If
            End With

#End Region
        ElseIf e.RunError.ErrorType = Errors.Type.Combatibility Then
#Region " ARCHITECTURE MISMATCH 32 bit vs 64 bit "

            Message.Show("Connection to " & errorDatasource & " is not currently possible", "Architecture Mismatch", Prompt.IconOption.TimedMessage, Prompt.StyleOption.Earth)
#End Region
        ElseIf e.RunError.ErrorType = Errors.Type.Requirement Then
#Region " NOT RUNNING AS ADMINISTRATOR "
            Message.Show("Connection to " & errorDatasource & " is not currently possible", "Administrator error. Reopen application running as administrator", Prompt.IconOption.TimedMessage, Prompt.StyleOption.Earth)
#End Region
        ElseIf e.RunError.ErrorType = Errors.Type.RunTime Then
#Region " RUNTIME (SCRIPT ERROR) "

#End Region
        Else
#Region " UNDEFINED "
            Message.Show("Undefined error", "Undefined error", Prompt.IconOption.TimedMessage, Prompt.StyleOption.Earth)
#End Region
        End If

    End Sub
    Private Sub IssuePromptShow()

        With IssuePrompt
            .AutoClose = False
            .Show(CenterItem(IssuePrompt.Size))
        End With
        Dim PasswordLocation = Password_New.PointToScreen(New Point(0, 0))
        PasswordLocation.Offset(Password_New.Image.Width + 4, Convert.ToInt32(Password_New.Height / 2))

    End Sub
    Private Sub IssuePromptHide() Handles IssueClose.Click, IssueMessage.Click

        With IssuePrompt
            .AutoClose = True
            .Hide()
        End With

    End Sub
    Private Sub PasswordChange(sender As Object, e As ImageComboEventArgs) Handles Password_New.ValueSubmitted

        Dim ResponseArgs = DirectCast(DirectCast(sender, ImageCombo).Tag, ResponseEventArgs)
        With ResponseArgs
            Dim _Connection = ResponseArgs.Connection
            If Password_New.Text = _Connection.Password Then
                IssuePromptShow()

            ElseIf Not Password_New.Text.any Then
                IssuePromptShow()

            ElseIf .RunError.Access = Errors.AccessResponse.PasswordExpired Then
                With _Connection
                    .NewPassword = Password_New.Text
                    RetrieveData(.ToString, "SELECT * FROM SYSIBM.SYSDUMMY1")
                    REM /// SAVING CLEARS THE NEWPWD FIELD
                    .Password = Password_New.Text
                    .Save()
                    IssuePromptHide()
                End With

            ElseIf .RunError.Access = Errors.AccessResponse.PasswordIncorrect Then
                With _Connection
                    .Password = Password_New.Text
                    .Save()
                End With
                IssuePromptHide()

            End If
        End With

    End Sub
End Class
#End Region
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
Public Class StringData
    Implements IEquatable(Of StringData)
    Public Sub New()
    End Sub
    Public Sub New(RM As Match)

        If RM Is Nothing Then
            Start = 0
            Length = 0
            Value = Nothing
        Else
            Start = RM.Index
            Length = RM.Length
            Value = RM.Value
        End If

    End Sub
    Public Sub New(RMO As Object)

        If RMO IsNot Nothing AndAlso RMO.GetType Is GetType(Match) Then
            Dim RM As Match = DirectCast(RMO, Match)
            Start = RM.Index
            Length = RM.Length
            Value = RM.Value
        End If

    End Sub
    Public ReadOnly Property All As List(Of StringData)
        Get
            Dim Nodes As New List(Of StringData)
            Dim Queue As New Queue(Of StringData)
            Dim TopNode As StringData
            Dim Node As StringData
            For Each TopNode In Parentheses
                Queue.Enqueue(TopNode)
            Next
            While Queue.Any
                TopNode = Queue.Dequeue
                Nodes.Add(TopNode)
                For Each Node In TopNode.Parentheses
                    Queue.Enqueue(Node)
                Next
            End While
            Return Nodes.OrderBy(Function(n) Nodes.IndexOf(n)).ToList
        End Get
    End Property
    Public ReadOnly Property Parentheses As New List(Of StringData)
    Public Property Start As Integer
    Public Property Length As Integer
    Public ReadOnly Property Finish As Integer
        Get
            Return Start + Length
        End Get
    End Property
    Public Property Value As String
    Public Property BackColor As Color
    Public Property ForeColor As Color
    Public Function Contains(SM As StringData) As Boolean

        If SM IsNot Nothing Then
            Return SM.Start >= Start And SM.Finish <= Finish
        Else
            Return False
        End If

    End Function
    Public Overrides Function ToString() As String
        Return Join({Value, Start, Length}, "*")
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return Start.GetHashCode Xor Length.GetHashCode Xor Value.GetHashCode Xor BackColor.GetHashCode Xor ForeColor.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As StringData) As Boolean Implements IEquatable(Of StringData).Equals
        If other Is Nothing Then
            Return Me Is Nothing
        Else
            Return Start = other.Start AndAlso Length = other.Length AndAlso Value = other.Value
        End If
    End Function
    Public Shared Operator =(ByVal value1 As StringData, ByVal value2 As StringData) As Boolean
        If value1 Is Nothing Then
            Return value2 Is Nothing
        ElseIf value2 Is Nothing Then
            Return value1 Is Nothing
        Else
            Return value1.Equals(value2)
        End If
    End Operator
    Public Shared Operator <>(ByVal value1 As StringData, ByVal value2 As StringData) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is StringData Then
            Return CType(obj, StringData) = Me
        Else
            Return False
        End If
    End Function
End Class
Public Structure InstructionElement
    Implements IEquatable(Of InstructionElement)
    Public Enum LabelName
        None
        WithTable
        WithBlock
        SystemTable
        FloatingTable
        RoutineTable
        SelectBlock
        SelectField
        GroupBlock
        GroupField
        OrderBlock
        OrderField
        Union
        Comment
        Constant
        Limit
    End Enum
    Public Property Block As StringData
    Public Property Highlight As StringData
    Public Property Source As LabelName
    Public ReadOnly Property Key As String
        Get
            Dim StartString As String = String.Empty
            Dim ValueString As String = String.Empty
            If IsNothing(Highlight) Then
                If IsNothing(Block) Then
                Else
                    StartString = Block.Start.ToString(CultureInfo.InvariantCulture)
                    If IsNothing(Block.Value) Then
                    Else
                        ValueString = Block.Value
                    End If
                End If
            Else
                StartString = Highlight.Start.ToString(CultureInfo.InvariantCulture)
                If IsNothing(Highlight.Value) Then
                Else
                    ValueString = Highlight.Value
                End If
            End If
            Return Join({Source.ToString, StartString, ValueString}, "§")
        End Get
    End Property
    Public ReadOnly Property Owner As String
        Get
            Dim ValueString As String = Split(Key, "§").Last
            Dim ObjectMatch As Match = Regex.Match(ValueString, "([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})", RegexOptions.IgnoreCase)
            If ObjectMatch.Success Then
                Dim Levels As String() = Split(ObjectMatch.Value, ".")
                If Levels.Count = 1 Then
                    Return String.Empty
                Else
                    Return Levels(Levels.Count - 2)
                End If
            Else
                Return String.Empty
            End If
        End Get
    End Property
    Public ReadOnly Property Name As String
        Get
            Dim ValueString As String = Split(Key, "§").Last
            Dim ObjectMatch As Match = Regex.Match(ValueString, "([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})", RegexOptions.IgnoreCase)
            If ObjectMatch.Success Then
                Dim Levels As String() = Split(ObjectMatch.Value, ".")
                Return Levels.Last
            Else
                Return String.Empty
            End If

        End Get
    End Property
    Public ReadOnly Property FullName As String
        Get
            If Owner.Length = 0 Then
                Return Name
            Else
                Return Join({Owner, Name}, ".")
            End If
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return Key
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return Block.GetHashCode Xor Highlight.GetHashCode Xor Source.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As InstructionElement) As Boolean Implements IEquatable(Of InstructionElement).Equals
        Return Block = other.Block AndAlso Highlight = other.Highlight AndAlso Source = other.Source
    End Function
    Public Shared Operator =(ByVal value1 As InstructionElement, ByVal value2 As InstructionElement) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(ByVal value1 As InstructionElement, ByVal value2 As InstructionElement) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is InstructionElement Then
            Return CType(obj, InstructionElement) = Me
        Else
            Return False
        End If
    End Function
End Structure
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
Public NotInheritable Class ConnectionCollection
    Inherits List(Of Connection)
    Private ReadOnly DataManagerDirectory As DirectoryInfo = Directory.CreateDirectory(MyDocuments & "\DataManager")
    Private WithEvents ChangeTimer As New Timer With {.Interval = 500}
    Public Sub New()

        If Not File.Exists(Path) Then
            Using SW As New StreamWriter(Path)
                SW.Write(My.Resources.Base_Connections)
            End Using
        End If
        Dim ConnectionStrings As New List(Of String)(PathToList(Path))
        For Each ConnectionString As String In ConnectionStrings
            Add(New Connection(ConnectionString))
        Next

    End Sub
    Public Sub New(Connections As String)

        For Each ConnectionString In Split(Connections, vbNewLine)
            Add(New Connection(ConnectionString))
        Next

    End Sub
    Public Sub New(Connections As List(Of String))

        If Connections IsNot Nothing Then
            For Each ConnectionString In Connections
                Add(New Connection(ConnectionString))
            Next
        End If

    End Sub
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Public Shadows Function Add(NewConnection As Connection) As Connection

        ChangeTimer.Stop()
        If NewConnection IsNot Nothing Then
            ChangeTimer.Start()
            MyBase.Add(NewConnection)
            NewConnection._Parent = Me
        End If
        Return NewConnection

    End Function
    Public Shadows Function AddRange(NewConnections As List(Of Connection)) As List(Of Connection)

        If NewConnections IsNot Nothing Then
            For Each NewConnection In NewConnections
                Add(NewConnection)
            Next
        End If
        Return NewConnections

    End Function
    Public Shadows Function Remove(OldConnection As Connection) As Connection

        ChangeTimer.Stop()
        If OldConnection IsNot Nothing Then
            ChangeTimer.Start()
            MyBase.Remove(OldConnection)
            OldConnection._Parent = Nothing
        End If
        Return OldConnection

    End Function
    Public Shadows Function Item(ConnectionString As String) As Connection

        Dim newConnection As New Connection(ConnectionString) 'Use the Connection.New Sub to create all the Connection properties
        Dim emptyConnections As New List(Of Connection)
        Dim uidConnections As New List(Of Connection)
        For Each child In Me
            If child.DataSource = newConnection.DataSource Then
                If child.UserID = newConnection.UserID Then
                    uidConnections.Add(child)
                Else
                    emptyConnections.Add(child)
                End If
            End If
        Next
        If uidConnections.Any Then
            Return uidConnections.First
        ElseIf emptyConnections.Any Then
            Return emptyConnections.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Item(ConnectionItem As Connection) As Connection

        REM /// DataSource And UserID IS A KEY...
        If ConnectionItem Is Nothing Then
            Return Nothing
        Else
            Dim Matches As IEnumerable(Of Connection)
            If ConnectionItem.MissingUserID And ConnectionItem.MissingPassword Then
                Matches = Where(Function(x) x.DataSource = ConnectionItem.DataSource)

            ElseIf ConnectionItem.MissingUserID Then
                Matches = Where(Function(x) x.DataSource = ConnectionItem.DataSource And x.Password = ConnectionItem.Password)

            ElseIf ConnectionItem.MissingPassword Then
                Matches = Where(Function(x) x.DataSource = ConnectionItem.DataSource And x.UserID = ConnectionItem.UserID)

            Else
                Matches = Where(Function(x) x.DataSource = ConnectionItem.DataSource And x.UserID = ConnectionItem.UserID And x.Password = ConnectionItem.Password)

            End If
            If Matches.Any Then
                Return Matches.First
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Shadows Function Contains(ConnectionItem As Connection) As Boolean
        Return Not IsNothing(Item(ConnectionItem))
    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private ReadOnly Property Path As String
        Get
            Return DataManagerDirectory.FullName & "\Connections.txt"
        End Get
    End Property
    Public ReadOnly Property Keys As List(Of String)
        Get
            Dim _Keys As New List(Of String)
            For Each ConnectionItem In Me
                _Keys.Add(ConnectionItem.Key)
            Next
            Return _Keys
        End Get
    End Property
    Public ReadOnly Property Sources As List(Of String)
        Get
            Dim _Sources As New List(Of String)
            For Each ConnectionItem In Me
                _Sources.Add(ConnectionItem.DataSource)
            Next
            Return _Sources
        End Get
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub ChangeTimerTick() Handles ChangeTimer.Tick
        ChangeTimer.Stop()
        SortCollection()
    End Sub
    Public Sub SortCollection()

        Sort(Function(f1, f2)
                 Dim Level1 = String.Compare(f1.DataSource, f2.DataSource, StringComparison.InvariantCulture)
                 If Level1 <> 0 Then
                     Return Level1
                 Else
                     Dim Level2 = String.Compare(f1.UserID, f2.UserID, StringComparison.InvariantCulture)
                     Return Level2
                 End If
             End Function)

    End Sub
    Public Sub View()
        Using Message As New Prompt
            Dim DT As New DataTable
            With DT
                .Columns.Add(New DataColumn With {.ColumnName = "DSN", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "UID", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "PWD", .DataType = GetType(String)})
                For Each Connection In Me
                    .Rows.Add({Connection.DataSource, Connection.UserID, Connection.Password})
                Next
            End With
            Message.Datasource = DT
            Message.Show("Connections Count=" & Count, "Current Connections", Prompt.IconOption.Warning, Prompt.StyleOption.Grey)
        End Using
    End Sub
    Public Sub Save()
        Using SW As New StreamWriter(Path)
            SW.Write(ToString)
        End Using
    End Sub
    Public Function ToStringArray() As String()
        Return (From m In Me Select m.ToString & String.Empty).ToArray
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
Public NotInheritable Class ConnectionChangedEventArgs
    Inherits EventArgs
    Public ReadOnly Property FormerConnection As Connection
    Public ReadOnly Property NewConnection As Connection
    Public ReadOnly Property FormerPassword As String
    Public ReadOnly Property NewPassword As String
    Public Sub New(FormerConnection As Connection, NewConnection As Connection)
        Me.FormerConnection = FormerConnection
        Me.NewConnection = NewConnection
    End Sub
    Public Sub New(FormerPassword As String, NewPassword As String)
        Me.FormerPassword = FormerPassword
        Me.NewPassword = NewPassword
    End Sub
End Class
<Serializable> Public NotInheritable Class Connection
    Implements IEquatable(Of Connection)

#Region " DECLARATIONS "
    Private ReadOnly NetezzaString As String = "Driver=;DSN=;UID=;PWD=;NEWPWD=;Database=;Servername=;SchemaName=;Port=;ReadOnly=;SQLBitOneZero=;FastSelect=;LegacySQLTables=;NumericAsChar=;ShowSystemTables=;LoginTimeout=;QueryTimeout=0;DateFormat=;SecurityLevel=;CaCertFile=;Nickname="
    Private ReadOnly DB2String As String = "Driver=;DSN=DSNA1;UID=;PWD=;NEWPWD=;Database=;MODE=;DBALIAS=;ASYNCENABLE=;USESCHEMAQUERIES=;Protocol=;HOSTNAME=;PORT=;QueryTimeout=600;Nickname="
#End Region
    Public Event PasswordChanged(sender As Object, e As ConnectionChangedEventArgs)
    Public Sub New(ConnectionString As String)

        Dim ConnectionElements As New List(Of String)
        IsFile = Regex.Match(ConnectionString, FilePattern, RegexOptions.IgnoreCase).Success
        If IsFile Then
            _Properties.Add(ConnectionString, ConnectionString)
        Else
            PropertyIndices = New Dictionary(Of String, Integer)
            IsNetezza = Regex.Match(ConnectionString, "DRIVER=\{(Netezza|NZ)SQL\}", RegexOptions.IgnoreCase).Success
            If IsNetezza Then
                '********** CANNOT BE FIRST PROPERTY!!! ===> DRIVER={NZSQL}
                ConnectionElements.AddRange(Split(NetezzaString.ToUpperInvariant, ";"))
            Else
                ConnectionElements.AddRange(Split(DB2String.ToUpperInvariant, ";"))
            End If
            IsDB2 = Not IsNetezza
            REM *** SOME CONNECTIONSTRINGS MAY COME THROUGH WITH ONLY A DSN+UID...GET THE PASSWORD FROM PARENT
            'DRIVER={IBM DB2 ODBC DRIVER};DSN=MWNCDSNB;UID=glover;PWD=PEANUT12;MODE=SHARE;ASYNCENABLE=0;USESCHEMAQUERIES=1;
            For Each Element In ConnectionElements
                Dim PropertyName As String = Split(Element, "=").First
                Dim PropertyValue As String = Split(Element, "=").Last
                If Not _Properties.ContainsKey(PropertyName) Then _Properties.Add(PropertyName, PropertyValue)
                PropertyIndices.Add(PropertyName, PropertyIndices.Count)
            Next
            For Each ProvidedElement In Split(ConnectionString, ";")
                Dim KeyValuePair() = Split(ProvidedElement, "=")
                If KeyValuePair.Count = 2 Then
                    Dim Key As String = Trim(KeyValuePair.First).ToUpperInvariant
                    Dim Value As String = Trim(KeyValuePair.Last)
                    If _Properties.ContainsKey(Key) Then _Properties(Key) = Value
                End If
            Next
            If ConnectionString = "DB2B1" Then Stop
        End If

    End Sub
#Region " PROPERTIES - FUNCTIONS - METHODS "
    <NonSerialized> Friend _Parent As ConnectionCollection
    Public ReadOnly Property Parent As ConnectionCollection
        Get
            Return _Parent
        End Get
    End Property
    Public ReadOnly Property DataSource As String
        Get
            Return If(_Properties.ContainsKey("DSN"), _Properties("DSN"), String.Empty)
        End Get
    End Property
    Public Property UserID As String
        Get
            Return If(_Properties.ContainsKey("UID"), _Properties("UID"), String.Empty)
        End Get
        Set(value As String)
            If _Properties.ContainsKey("UID") Then _Properties("UID") = value
        End Set
    End Property
    Public Property Password As String
        Get
            Return If(_Properties.ContainsKey("PWD"), _Properties("PWD"), String.Empty)
        End Get
        Set(value As String)
            If _Properties.ContainsKey("PWD") Then
                RaiseEvent PasswordChanged(Me, New ConnectionChangedEventArgs(_Properties("PWD"), value))
                _Properties("PWD") = value
            End If
        End Set
    End Property
    Public Property NewPassword As String
        Get
            Return _Properties("NEWPWD")
        End Get
        Set(value As String)
            _Properties("NEWPWD") = value
        End Set
    End Property
    Public ReadOnly Property CanConnect As Boolean
        Get
            Return Not (_Properties("DSN").Length = 0 Or _Properties("UID").Length = 0 Or _Properties("PWD").Length = 0)
        End Get
    End Property
    Public ReadOnly Property IsFile As Boolean
    Public ReadOnly Property IsNetezza As Boolean
    Public ReadOnly Property IsDB2 As Boolean
    Public ReadOnly Property Key As String
        Get
            Return Join({"DSN=" & DataSource, "UID=" & UserID}, ";")
        End Get
    End Property
    Public ReadOnly Property MissingUserID As Boolean
        Get
            Return If(IsFile, False, UserID.Length = 0)
        End Get
    End Property
    Public ReadOnly Property MissingPassword As Boolean
        Get
            Return If(IsFile, False, Password.Length = 0)
        End Get
    End Property
    Private ReadOnly _Properties As New Dictionary(Of String, String)
    Public ReadOnly Property Properties As Dictionary(Of String, String)
        Get
            Return _Properties.Where(Function(x) x.Value.Any).ToDictionary(Function(x) x.Key, Function(y) y.Value)
        End Get
    End Property
    Public Sub SetProperty(name As String, value As String)
        If _Properties.ContainsKey(name) Then _Properties(name) = value
    End Sub
    Public ReadOnly Property PropertiesEmpty As List(Of String)
        Get
            Return _Properties.Keys.Except(Properties.Keys).ToList
        End Get
    End Property
    Public ReadOnly Property PropertyIndices As Dictionary(Of String, Integer)
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public ReadOnly Property Index As Integer
        Get
            If Parent Is Nothing Then
                Return 1
            Else
                Return Parent.IndexOf(Me)
            End If
        End Get
    End Property
    Private ReadOnly ConnectionBackColors As New List(Of Color) From {
                    Color.LightBlue,
                    Color.LimeGreen,
                    Color.Peru,
                    Color.CornflowerBlue,
                    Color.BlueViolet,
                    Color.SlateBlue,
                    Color.DeepPink,
                    Color.Gold,
                    Color.Green,
                    Color.IndianRed,
                    Color.Silver}
    Private ReadOnly ConnectionForeColors As New List(Of Color) From {
                    Color.DarkBlue,
                    Color.DarkGreen,
                    Color.White,
                    Color.White,
                    Color.White,
                    Color.White,
                    Color.White,
                    Color.Black,
                    Color.White,
                    Color.White,
                    Color.Black}
    Public ReadOnly Property BackColor As Color
        Get
            Return ConnectionBackColors({Index Mod ConnectionBackColors.Count, 0}.Max)
        End Get
    End Property
    Public ReadOnly Property ForeColor As Color
        Get
            Return ConnectionForeColors({Index Mod ConnectionForeColors.Count, 0}.Max)
        End Get
    End Property
    Public Shared Function FromString(Value As String) As Connection
        Return New Connection(Value)
    End Function
    Public Overrides Function ToString() As String
        If IsFile Then
            Return Properties.Keys.First
        Else
            Return Join((From p In Properties Select Join({p.Key, p.Value}, "=")).ToArray, ";")
        End If
    End Function
    Public Sub Save()
        _Properties("NEWPWD") = String.Empty
        Parent.Save()
    End Sub

    Public Overrides Function GetHashCode() As Integer
        Return DataSource.GetHashCode Xor UserID.GetHashCode Xor Password.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As Connection) As Boolean Implements IEquatable(Of Connection).Equals

        If DataSource Is Nothing Then
            Return other Is Nothing
        ElseIf other Is Nothing Then
            Return DataSource Is Nothing
        Else
            Return DataSource = other.DataSource AndAlso UserID = other.UserID
        End If

    End Function
    Public Shared Operator =(ByVal value1 As Connection, ByVal value2 As Connection) As Boolean

        Dim string1 As String = If(value1 Is Nothing, String.Empty, value1.Key)
        Dim string2 As String = If(value2 Is Nothing, String.Empty, value2.Key)
        Return string1 = string2

    End Operator
    Public Shared Operator <>(ByVal value1 As Connection, ByVal value2 As Connection) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is Connection Then
            Return CType(obj, Connection) = Me
        Else
            Return False
        End If
    End Function
#End Region
End Class
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
Public NotInheritable Class SystemObjectCollection
    Inherits List(Of SystemObject)
    Private ReadOnly Folder_DataManager As String = MyDocuments & "\DataManager"
    Private WithEvents ChangeTimer As New Timer With {.Interval = 500}
#Region " NEW "
    Public Sub New()

        Directory.CreateDirectory(Folder_DataManager)
        If Not File.Exists(Path) Then
            Using SW As New StreamWriter(Path)
                'SW.Write(My.Resources.BASE_OBJECTS)
            End Using
        End If
        For Each ObjectString In PathToList(Path)
            Add(New SystemObject(ObjectString))
        Next

    End Sub
    Public Sub New(Items As List(Of SystemObject))

        Directory.CreateDirectory(Folder_DataManager)
        If Items IsNot Nothing Then
            For Each ObjectItem In Items
                Add(ObjectItem)
            Next
        End If

    End Sub
    Public Sub New(Table As DataTable)

        Directory.CreateDirectory(Folder_DataManager)
        If Table IsNot Nothing Then
            For Each Row In Table.AsEnumerable
                Add(New SystemObject(Row))
            Next
        End If

    End Sub
#End Region
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Public Shadows Function Add(ObjectItem As SystemObject) As SystemObject

        If ObjectItem IsNot Nothing Then
            If ObjectItem.Type = SystemObject.ObjectType.None Then
                Return Nothing
            Else
                ChangeTimer.Stop()
                ChangeTimer.Start()
                MyBase.Add(ObjectItem)
                ObjectItem._Parent = Me
                Return ObjectItem
            End If
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Add(ObjectItem As String) As SystemObject

        Dim NewItem As New SystemObject(ObjectItem)
        If NewItem.Type = SystemObject.ObjectType.None Then
            Return Nothing
        Else
            ChangeTimer.Stop()
            ChangeTimer.Start()
            MyBase.Add(NewItem)
            NewItem._Parent = Me
            Return NewItem
        End If

    End Function
    Public Shadows Function Remove(ObjectItem As String) As SystemObject

        Dim OldItem As SystemObject = Item(ObjectItem)
        If OldItem Is Nothing Then
            Return Nothing
        Else
            If OldItem.Type = SystemObject.ObjectType.None Then
                Return Nothing
            Else
                ChangeTimer.Stop()
                ChangeTimer.Start()
                MyBase.Remove(OldItem)
                Return OldItem
            End If
        End If

    End Function
    Public Shadows Function AddRange(ObjectItems As List(Of SystemObject)) As List(Of SystemObject)

        Dim ObjectItemsAdded As New List(Of SystemObject)
        If ObjectItems IsNot Nothing Then
            For Each ObjectItem In ObjectItems
                If Not Contains(ObjectItem) Then
                    ObjectItemsAdded.Add(ObjectItem)
                    Add(ObjectItem)
                End If
            Next
        End If
        'RaiseEvent ItemsAdded(ObjectItem, New ChangeEventArgs(ObjectItems))
        Return ObjectItemsAdded

    End Function
    Public Shadows Function Remove(ObjectItem As SystemObject) As SystemObject

        ChangeTimer.Stop()
        If ObjectItem IsNot Nothing Then
            ChangeTimer.Start()
            MyBase.Remove(ObjectItem)
            ObjectItem._Parent = Nothing
            'RaiseEvent ItemRemoved(ObjectItem, New ChangeEventArgs(ObjectItem))
        End If
        Return ObjectItem

    End Function
    Public Shadows Function Item(ObjectItem As SystemObject) As SystemObject

        REM /// DSN And Owner And Name IS A KEY...
        If IsNothing(ObjectItem) Then
            Return Nothing
        Else
            Dim _SystemObjects = Where(Function(x) x.DSN = ObjectItem.DSN And x.FullName = ObjectItem.FullName)
            If _SystemObjects.Any Then
                Return _SystemObjects.First
            Else
                Return Nothing
            End If
        End If

    End Function
    Public Shadows Function Item(ObjectItem As String) As SystemObject

        REM /// DSN And Owner And Name IS A KEY...
        If IsNothing(ObjectItem) Then
            Return Nothing
        Else
            Dim Levels = Split(ObjectItem, Delimiter)
            If Levels.Count = 3 Then
                REM /// DSN And Owner And Name IS A KEY...
                Return Item(Levels.First, Strings.Join({Levels(1), Levels(2)}, "."))
            Else
                REM /// CDNIW§Table§QBIMKTS§MKTGTSS§A085365§FACTORING_XYZ
                Dim _SystemObjects = Where(Function(x) x.ToString = ObjectItem)
                If _SystemObjects.Any Then
                    Return _SystemObjects.First
                Else
                    Return Nothing
                End If
            End If
        End If

    End Function
    Public Shadows Function Item(DataSource As String, FullName As String) As SystemObject

        REM /// DSN And Owner And Name IS A KEY...
        Dim _SystemObjects = Where(Function(x) x.DSN = DataSource And x.FullName = FullName)
        If _SystemObjects.Any Then
            Return _SystemObjects.First
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Contains(ObjectItem As SystemObject) As Boolean
        Return Not IsNothing(Item(ObjectItem))
    End Function
    Public Shadows Function Distinct() As List(Of SystemObject)

        Dim Items As New List(Of SystemObject)
        Dim Groups = (From M In Me Group By _Key = M.ToString Into StringGroup = Group Select New With {.Key = _Key, .Value = StringGroup.First})
        For Each ObjectItem In Groups
            Items.Add(ObjectItem.Value)
        Next
        Return Items

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Friend ReadOnly Property Path As String
        Get
            Return Folder_DataManager & "\Objects.txt"
        End Get
    End Property
    Public ReadOnly Property DataSources As Dictionary(Of String, List(Of SystemObject))
        Get
            Return (From D In Me Group D By _DSN = D.DSN Into DataSourceGroup = Group Select New With {.DataSource = _DSN, .Items = DataSourceGroup}).ToDictionary(Function(x) x.DataSource, Function(y) y.Items.ToList)
        End Get
    End Property
    Public ReadOnly Property Owners As Dictionary(Of String, List(Of SystemObject))
        Get
            Return (From O In Me Group O By _Owner = O.Owner Into OwnerGroup = Group Select New With {.Owner = _Owner, .Items = OwnerGroup}).ToDictionary(Function(x) x.Owner, Function(y) y.Items.ToList)
        End Get
    End Property
    Public ReadOnly Property Names As Dictionary(Of String, List(Of SystemObject))
        Get
            Return (From N In Me Group N By _Name = N.Name Into NameGroup = Group Select New With {.Name = _Name, .Items = NameGroup}).ToDictionary(Function(x) x.Name, Function(y) y.Items.ToList)
        End Get
    End Property
    Public ReadOnly Property Objects As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, SystemObject)))
        Get
            Dim Grouped As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, SystemObject)))
            Dim X = From D In Me Group D By _DataSource = D.DSN Into SourceGroup = Group
                    Select New With {.DataSource = _DataSource, .Owners = From S In SourceGroup Group S By _Owner = S.Owner Into OwnerGroup = Group
                                                                          Select New With {.Owner = _Owner, .Items = OwnerGroup}}
            For Each Lvl1 In X
                Grouped.Add(Lvl1.DataSource, New Dictionary(Of String, Dictionary(Of String, SystemObject)))
                For Each Lvl2 In Lvl1.Owners
                    Grouped(Lvl1.DataSource).Add(Lvl2.Owner, New Dictionary(Of String, SystemObject))
                    For Each Lvl3 In Lvl2.Items
                        If Not Grouped(Lvl1.DataSource)(Lvl2.Owner).ContainsKey(Lvl3.Name) Then Grouped(Lvl1.DataSource)(Lvl2.Owner).Add(Lvl3.Name, Lvl3)
                    Next
                Next
            Next
            Return Grouped
        End Get
    End Property
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function Items(DataString As String) As List(Of SystemObject)

        REM /// SPLIT ALWAYS RENDERS A COUNT OF 1. CURRENTLY MY.SETTINGS.SystemObjects=6
        REM /// NOT FROM MY.SETTINGS...SystemObjectCOLLECTION HAS BEEN POPULATED
        REM /// RETURN ONLY ITEMS THAT EXIST IN MY.SETTINGS

        Dim _Objects = Objects
        Dim _DataSources = DataSources
        Dim _Owners = Owners
        Dim _Names = Names

        REM /// LIMIT 3 LEVELS ///
        Dim Levels As String() = Split(DataString, ".").Take(3).ToArray
        Dim ObjectName As String = Levels.Last.ToUpperInvariant
        Dim Collection As New List(Of SystemObject)

        If Names.ContainsKey(ObjectName) Then
            'OBJECTNAME EXISTS IN SystemObjects
            Select Case Levels.Count
                Case 1
#Region " 1 LEVEL=NAME ONLY...ACTIONS "
                    'Return Names(ObjectName)
                    Collection.AddRange(Names(ObjectName))
#End Region
                Case 2
#Region " 2 LEVELS={DATASOURCE.NAME Or OWNER.NAME}...DSNA1.ACTIONS Or C085365.ACTIONS "
                    Dim Level1 As String = Levels.First.ToUpperInvariant
                    If _DataSources.ContainsKey(Level1) Then
#Region " DATASOURCE.NAME - RETURN MULTIPLE OWNERS IN ONE DATASOURCE HAVING THE SAME NAME - NEVER EMPTY "
                        'Return _DataSources(Level1).Where(Function(x) x.Name = ObjectName).ToList
                        Collection.AddRange(_DataSources(Level1).Where(Function(x) x.Name = ObjectName))
#End Region
                    ElseIf _Owners.ContainsKey(Level1) Then
#Region " OWNER.NAME - RETURN MATCHES Or A NEW SystemObject "
                        'C085365.ADDRESSES DOES NOT EXIST...A085365.ADDRESSES DOES...SO MATCHED OK ON <NAME>
                        'HOWEVER NAMES.CONTAINS('ADDRESSES') AND OWNERS.CONTAINS('C085365') BUT THERE IS NO C085365+ADDRESSES AND RETURNED EMPTY LIST
                        Dim Owners = _Owners(Level1).Where(Function(x) x.Name = ObjectName).ToList
                        If Owners.Any Then
                            'Return Owners
                            Collection.AddRange(Owners)
                        Else
                            'Return {New SystemObject With {.Owner = Level1, .Name = ObjectName}}.ToList
                            Collection.AddRange({New SystemObject With {.Owner = Level1, .Name = ObjectName}})
                        End If
                    Else
                        'ASSUME             C085365.ACTIONS
                        'Return {New SystemObject With {.Owner = Level1, .Name = ObjectName}}.ToList
                        Collection.AddRange({New SystemObject With {.Owner = Level1, .Name = ObjectName}})
                    End If
#End Region
#End Region
                Case 3
#Region " 3 LEVELS=DATASOURCE.OWNER.NAME (FULL) "
#Region " EITHER EXISTS Or NOT - RETURN OBJECT Or NEW SystemObject "
                    Dim Level1 As String = Levels.First.ToUpperInvariant
                    Dim Level2 As String = Levels(1).ToUpperInvariant
                    If _Objects.ContainsKey(Level1) AndAlso _Objects(Level1).ContainsKey(Level2) AndAlso _Objects(Level1)(Level2).ContainsKey(ObjectName) Then
                        'Return _DataSources(Level1).Where(Function(x) x.Owner = Level2 And x.Name = ObjectName).ToList
                        Collection.AddRange(_DataSources(Level1).Where(Function(x) x.Owner = Level2 And x.Name = ObjectName))
                    Else
                        'Return {New SystemObject With {.DSN = Level1, .Owner = Level2, .Name = ObjectName}}.ToList
                        Collection.AddRange({New SystemObject With {.DSN = Level1, .Owner = Level2, .Name = ObjectName}})
                    End If
#End Region
#End Region
                Case Else
#Region " 0 Or 4+ LEVELS (WON'T EXIST...CAN'T BE 0 AND IS LIMITED TO 3 BUT MUST HAVE ELSE IN ORDER TO RETURN ON ALL PATHS) "
                    'Return {New SystemObject With {.Name = ObjectName}}.ToList
#End Region
            End Select
        Else
#Region " CREATE A NEW OBJECT - NO MATCH ON NAME=DOESN'T EXIST UNLESS TYPO? "
            REM /// TRY MISSPELLED WORDS???
            Select Case Levels.Count
                Case 1
                        'NAME ONLY:             MONKEY
                        'Return {New SystemObject With {.Name = ObjectName}}.ToList

                Case 2
                    'EITHER OF:             DSNA1.MONKEY Or C085365.MONKEY
                    Dim Level1 As String = Levels.First.ToUpperInvariant
                    If _DataSources.ContainsKey(Level1) Then
                        'Return {New SystemObject With {.DSN = Level1, .Name = ObjectName}}.ToList
                    Else
                        'Return {New SystemObject With {.Owner = Level1, .Name = ObjectName}}.ToList
                    End If

                Case 3
                    'ALL WERE PROVIDED:     DSNA1.C085365.MONKEY
                    Dim Level1 As String = Levels.First.ToUpperInvariant
                    Dim Level2 As String = Levels(1).ToUpperInvariant
                    'Return {New SystemObject With {.DSN = Level1, .Owner = Level2, .Name = ObjectName}}.ToList

                Case Else
                    'WON'T EXIST...CAN'T BE 0 AND IS LIMITED TO 3
                    'Return {New SystemObject With {.Name = ObjectName}}.ToList

            End Select
#End Region
        End If
        Return Collection

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Private Sub ChangeTimerTick() Handles ChangeTimer.Tick
        ChangeTimer.Stop()
        SortCollection()
    End Sub
    Public Sub RemoveDuplicates()

        Dim Duplicates As New Dictionary(Of String, List(Of SystemObject))
        For Each ItemObject In Me
            Dim ItemKey As String = ItemObject.Key
            If Not Duplicates.ContainsKey(ItemKey) Then Duplicates.Add(ItemKey, New List(Of SystemObject))
            Duplicates(ItemKey).Add(ItemObject)
        Next
        For Each ItemObjects In Duplicates.Where(Function(io) io.Value.Count > 1)
            For Each DuplicateItem In ItemObjects.Value.Skip(1)
                Remove(DuplicateItem.Key)
            Next
        Next
        SortCollection()
        Save()

    End Sub
    Public Sub SortCollection()

        Sort(Function(f1, f2)
                 Dim Level1 = String.Compare(f1.DSN, f2.DSN, True, InvariantCulture)
                 If Level1 <> 0 Then
                     Return Level1
                 Else
                     Dim Level2 = String.Compare(f1.Type.ToString, f2.Type.ToString, True, InvariantCulture)
                     If Level2 <> 0 Then
                         Return Level2
                     Else
                         Dim Level3 = String.Compare(f1.Owner, f2.Owner, True, InvariantCulture)
                         If Level3 <> 0 Then
                             Return Level3
                         Else
                             Dim Level4 = String.Compare(f1.Name, f2.Name, True, InvariantCulture)
                             Return Level4
                         End If
                     End If
                 End If
             End Function)

    End Sub
    Public Sub View()
        Using Message As New Prompt
            Dim DT As New DataTable
            With DT
                .Columns.Add(New DataColumn With {.ColumnName = "DSN", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "TYPE", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "FULLNAME", .DataType = GetType(String)})
                For Each SystemObject In Me
                    .Rows.Add({SystemObject.DSN, SystemObject.Type.ToString, SystemObject.FullName})
                Next
            End With
            Message.Datasource = DT
            Message.Show("Objects Count=" & Count, "Current Objects", Prompt.IconOption.Warning, Prompt.StyleOption.Grey)
        End Using
    End Sub
    Public Sub Save()
        Using SW As New StreamWriter(Path)
            SW.Write(ToString)
        End Using
    End Sub
    Public Function ToStringArray() As String()
        Return (From m In Me Select m.ToString & String.Empty).ToArray
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
    Public Function ToDataTable() As DataTable

        Dim ObjectsTable As New DataTable
        'DataSource, Type, DBName, TSName, Owner, Name
        With ObjectsTable.Columns
            .Add(New DataColumn With {.ColumnName = "DataSource", .DataType = GetType(String)})
            .Add(New DataColumn With {.ColumnName = "Type", .DataType = GetType(String)})
            .Add(New DataColumn With {.ColumnName = "DBName", .DataType = GetType(String)})
            .Add(New DataColumn With {.ColumnName = "TSName", .DataType = GetType(String)})
            .Add(New DataColumn With {.ColumnName = "Owner", .DataType = GetType(String)})
            .Add(New DataColumn With {.ColumnName = "Name", .DataType = GetType(String)})
        End With
        For Each ObjectItem In Me
            ObjectsTable.Rows.Add(Split(ObjectItem.ToString, Delimiter))
        Next
        Return ObjectsTable

    End Function
#End Region
End Class
<Serializable> Public NotInheritable Class SystemObject
#Region " CLASSES - ENUMS - STRUCTURES "
    Public Enum ObjectType
        None
        Table
        View
        Routine
        Trigger
    End Enum
#End Region
#Region " NEW "
    Public Sub New()
    End Sub
    Public Sub New(DataString As String)

        Dim DataElements As String() = Split(DataString, Delimiter)
        REM /// MY.SETTINGS.SystemObjects:  DSN      §   Type    §   DBNAME  §   TSNAME  §   OWNER   §   NAME
        '------------------------------ ----------- ----------- ----------- ----------- ----------- -----------
        REM /// MY.SETTINGS.SystemObjects:  CDNIW    §   Table   §   QBIMKTS §   MKTGTSS §   A085365 §   ACCESS
        DSN = DataElements.First
        Type = DirectCast([Enum].Parse(GetType(ObjectType), StrConv(DataElements(1), VbStrConv.ProperCase)), ObjectType)
        DBName = DataElements(2)
        TSName = DataElements(3)
        Owner = DataElements(4)
        Name = DataElements.Last

    End Sub
    Public Sub New(Row As DataRow)

        If Row IsNot Nothing Then
            Using Table As DataTable = Row.Table
                Dim Columns As DataColumnCollection = Table.Columns
                If Columns.Contains("DataSource") Then
                    DSN = Convert.ToString(Row("DataSource"), InvariantCulture)
                End If
                If Columns.Contains("Type") Then
                    Type = DirectCast([Enum].Parse(GetType(ObjectType), StrConv(Convert.ToString(Row("Type"), InvariantCulture), VbStrConv.ProperCase)), ObjectType)
                End If
                If Columns.Contains("DBName") Then
                    DBName = Convert.ToString(Row("DBName"), InvariantCulture)
                End If
                If Columns.Contains("TSName") Then
                    TSName = Convert.ToString(Row("TSName"), InvariantCulture)
                End If
                If Columns.Contains("Owner") Then
                    Owner = Convert.ToString(Row("Owner"), InvariantCulture)
                End If
                If Columns.Contains("Name") Then
                    Name = Convert.ToString(Row("Name"), InvariantCulture)
                End If
            End Using
        End If

    End Sub
#End Region
#Region " PROPERTIES - FUNCTIONS - METHODS "
    <NonSerialized> Friend _Parent As SystemObjectCollection
    Public ReadOnly Property Parent As SystemObjectCollection
        Get
            Return _Parent
        End Get
    End Property
    Public Property DSN As String
    Public Property Type As ObjectType
    Public Property DBName As String
    Public Property TSName As String
    Public Property Owner As String
    Public Property Name As String
    Public ReadOnly Property FullName As String
        Get
            Return Join({Owner, Name}, ".")
        End Get
    End Property
    Public ReadOnly Property Key As String
        Get
            Return Join({DSN, Owner, Name}, Delimiter)
        End Get
    End Property
    Public ReadOnly Property Connection As Connection
        Get
            Dim Connections = New ConnectionCollection
            Dim _Connections As New List(Of Connection)(Connections.Where(Function(x) x.DataSource = DSN))
            If _Connections.Any Then
                Return _Connections.First
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return Join({DSN, Type.ToString, DBName, TSName, Owner, Name}, Delimiter)
    End Function
#End Region
End Class
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
Public Class JobCollection
    Inherits List(Of Job)
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
            Handle.Dispose()
            ' Free any other managed objects here.
            For Each Job In Me
                Job.Dispose()
            Next
        End If
        disposed = True
    End Sub
#End Region
    Public Event Completed(sender As Object, e As ResponsesEventArgs)
    Public Event JobStarted(sender As Object, e As ResponsesEventArgs)
    Public Event JobEnded(sender As Object, e As ResponsesEventArgs)
    Private ReadOnly DataManagerDirectory As DirectoryInfo = Directory.CreateDirectory(MyDocuments & "\DataManager")
#Region " NEW "
    Public Sub New()

        For Each JobString In PathToList(Path)
            Add(New Job(JobString))
        Next

    End Sub
#End Region
#Region " PROPERTIES - FUNCTIONS - METHODS "
    Private ReadOnly Property Path As String
        Get
            Return DataManagerDirectory.FullName & "\Jobs.txt"
        End Get
    End Property
    Public Property Name As String
    Public Shadows Function Add(JobItem As Job) As Job

        If JobItem IsNot Nothing Then
            MyBase.Add(JobItem)
            JobItem._Parent = Me
            'RaiseEvent ItemAdded(JobItem, New ChangeEventArgs(JobItem))
        End If
        Return JobItem

    End Function
    Public Shadows Function Add(DDL As DDL) As Job

        If DDL IsNot Nothing Then
            Dim JobItem As New Job(DDL)
            MyBase.Add(JobItem)
            JobItem._Parent = Me
            'RaiseEvent ItemAdded(JobItem, New ChangeEventArgs(JobItem))
            Return JobItem
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Add(ETL As ETL) As Job

        If ETL IsNot Nothing Then
            Dim JobItem As New Job(ETL)
            MyBase.Add(JobItem)
            JobItem._Parent = Me
            'RaiseEvent ItemAdded(JobItem, New ChangeEventArgs(JobItem))
            Return JobItem
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Add(SQL As SQL) As Job

        If SQL IsNot Nothing Then
            Dim JobItem As New Job(SQL)
            MyBase.Add(JobItem)
            JobItem._Parent = Me
            'RaiseEvent ItemAdded(JobItem, New ChangeEventArgs(JobItem))
            Return JobItem
        Else
            Return Nothing
        End If

    End Function
    Public Shadows Function Item(Name As String) As Job

        If Count = 0 Then
            Return Nothing
        Else
            Dim Jobs = From j In Me Where j.Name.ToUpper(CultureInfo.InvariantCulture) = Name.ToUpper(CultureInfo.InvariantCulture)
            If Jobs.Any Then
                Return Jobs.First
            Else
                Return Nothing
            End If
        End If

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Sub SortCollection()

        Sort(Function(f1, f2)
                 Dim Level1 = String.Compare(f1.Request.ToString.ToUpperInvariant, f2.Request.ToString.ToUpperInvariant, StringComparison.Ordinal)
                 If Level1 <> 0 Then
                     Return Level1
                 Else
                     Dim Level2 = String.Compare(f1.Name.ToUpperInvariant, f2.Name.ToUpperInvariant, StringComparison.Ordinal)
                     Return Level2
                 End If
             End Function)

    End Sub
    Public Sub View()
        Using Message As New Prompt
            Dim DT As New DataTable
            With DT
                .Columns.Add(New DataColumn With {.ColumnName = "DSN", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "NAME", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "SCHEDULE", .DataType = GetType(Date)})
            End With
            Message.Datasource = DT
            Message.Show("Jobs Count=" & Count, "Current Jobs", Prompt.IconOption.Warning, Prompt.StyleOption.Grey)
        End Using
    End Sub
    Public Sub Save()
        Using SW As New StreamWriter(Path)
            SW.Write(ToString)
        End Using
    End Sub
    Public Function ToStringArray() As String()
        Return (From m In Me Select m.ToString & String.Empty).ToArray
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
    Public Sub Execute(Optional Sequential As Boolean = False)

        _Started = Now
        Responses.Clear()
        If Sequential Then
            For Each Job In Me
                AddHandler Job.Completed, AddressOf Job_Completed
            Next
            IterateJobs()
        Else
            For Each Job In Me
                AddHandler Job.Completed, AddressOf Job_Completed
                Job.Execute()
            Next
        End If

    End Sub
    Private Sub IterateJobs()

        Dim NotDone = From j In Me Where j.Responses.Count = 0
        If NotDone.Any Then
            With NotDone.First
                RaiseEvent JobStarted(NotDone.First, Nothing)
                .Execute()
            End With
        End If

    End Sub
    Private Sub Job_Completed(sender As Object, e As ResponsesEventArgs)

        Dim CurrentJob = DirectCast(sender, Job)
        With CurrentJob
            RemoveHandler .Completed, AddressOf Job_Completed
            RaiseEvent JobEnded(CurrentJob, e)
            Responses.AddRange(.Responses)
            IterateJobs()
            If AllCompleted Then
                _Ended = Now
                RaiseEvent Completed(Me, New ResponsesEventArgs(Responses))
            End If
        End With

    End Sub
    Public ReadOnly Property Elapsed As TimeSpan
        Get
            Return Ended - Started
        End Get
    End Property
    Public ReadOnly Property Started As Date
    Public ReadOnly Property Busy As Boolean
        Get
            Return Started <> New Date And Ended <> New Date
        End Get
    End Property
    Public ReadOnly Property Ended As Date
    Private ReadOnly Property AllCompleted As Boolean
        Get
            Dim Done = Where(Function(j) j.Ended > New Date)
            Return Done.Count = Count
        End Get
    End Property
    Public ReadOnly Property Succeeded As Boolean
        Get
            If AllCompleted Then
                Return Count = Where(Function(j) j.Succeeded).Count
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property Table As DataTable
        Get
            Dim Tables As New DataTable
            For Each Job In Where(Function(j) j.SQL IsNot Nothing)
                If Job.SQL IsNot Nothing AndAlso Job.SQL.Table IsNot Nothing Then
                    Try
                        Tables.Merge(Job.SQL.Table)
                    Catch ex As DataException
                    End Try
                End If
            Next
            For Each Job In Where(Function(j) j.ETL IsNot Nothing)
                Try
                    For Each Source In Job.ETL.Sources
                        Tables.Merge(Source.Table)
                    Next
                Catch ex As DataException
                End Try
            Next
            Return Tables
        End Get
    End Property
    Public ReadOnly Property Responses As New List(Of ResponseEventArgs)
#End Region
End Class
<Serializable> Public Class Job
    Implements IDisposable
#Region " DISPOSE "
    Dim disposed As Boolean = False
    <NonSerialized> ReadOnly Handle As SafeHandle = New Microsoft.Win32.SafeHandles.SafeFileHandle(IntPtr.Zero, True)
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If disposed Then Return
        If disposing Then
            Handle.Dispose()
            ' Free any other managed objects here.
            If DDL IsNot Nothing Then _DDL.Dispose()
            If ETL IsNot Nothing Then _ETL.Dispose()
            If SQL IsNot Nothing Then _SQL.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public Event Completed(sender As Object, e As ResponsesEventArgs)
#Region " NEW "
    Public Sub New(DDL As DDL)
        _DDL = DDL
        Request = Type.DDL
    End Sub
    Public Sub New(ETL As ETL)
        _ETL = ETL
        Request = Type.ETL
    End Sub
    Public Sub New(SQL As SQL)
        _SQL = SQL
        Request = Type.SQL
    End Sub
    Public Sub New(JobType As Type)
        Request = JobType
    End Sub
    Public Sub New(JobString As String)

        Dim JobElements As New List(Of String)(Split(JobString, Delimiter))
        Name = JobElements.First
        If JobElements.Count = 3 Then
            REM /// INITIAL JOB ADD WILL ONLY BE THE NAME, DDL And SERVER
            Instruction = JobElements(1)
            AddDate = Now
            LastRunDate = Nothing
        Else
            REM /// EXISTING WILL BE
            Dim Dates = From JE In JobElements Where IsDate(JE) Select Date.Parse(JE, InvariantCulture)
            AddDate = Dates.Min
            If Dates.Count = 2 Then LastRunDate = Dates.Max
            Dim Procedures As New List(Of String)(From JE In JobElements Where JE.ToUpperInvariant.Contains("SELECT"))
            If Procedures.Any Then Instruction = Procedures.First
        End If

    End Sub
#End Region
    Public Enum Type
        DDL
        ETL
        SQL
    End Enum
#Region " PROPERTIES - FUNCTIONS - METHODS "
    <NonSerialized> Friend _Parent As JobCollection
    Public ReadOnly Property Parent As JobCollection
        Get
            Return _Parent
        End Get
    End Property
    <NonSerialized> Private _DDL As DDL
    Public ReadOnly Property DDL As DDL
        Get
            Return _DDL
        End Get
    End Property
    <NonSerialized> Private ReadOnly _ETL As ETL
    Public ReadOnly Property ETL As ETL
        Get
            Return _ETL
        End Get
    End Property
    <NonSerialized> Private _SQL As SQL
    Public ReadOnly Property SQL As SQL
        Get
            Return _SQL
        End Get
    End Property
    Public ReadOnly Property Index As Integer
        Get
            If Parent Is Nothing Then
                Return -1
            Else
                Return Parent.IndexOf(Me)
            End If
        End Get
    End Property
    Public ReadOnly Property Request As Type
    Public Property SourceConnection As Connection
    Public Property Name As String
    Public Property Instruction As String
    Public ReadOnly Property Elapsed As TimeSpan
        Get
            Return Ended - Started
        End Get
    End Property
    Public ReadOnly Property Started As Date
    Public ReadOnly Property Ended As Date
    Public ReadOnly Property Succeeded As Boolean
    Public ReadOnly Property Responses As New List(Of ResponseEventArgs)
    Public Property Schedule As Frequency
    Public ReadOnly Property AddDate As Date
    Public ReadOnly Property LastRunDate As Date
    Public Sub Execute()

        _Started = Now
        Responses.Clear()
        If Request = Type.DDL Then
            If DDL Is Nothing Then _DDL = New DDL(SourceConnection.ToString, Instruction)
            AddHandler DDL.Completed, AddressOf Request_Completed
            DDL.Execute()

        ElseIf Request = Type.ETL Then
            AddHandler ETL.Completed, AddressOf Requests_Completed
            ETL.Execute()

        ElseIf Request = Type.SQL Then
            If SQL Is Nothing Then _SQL = New SQL(SourceConnection.ToString, Instruction)
            AddHandler SQL.Completed, AddressOf Request_Completed
            SQL.Execute()

        End If

    End Sub
    Private Sub Request_Completed(sender As Object, e As ResponseEventArgs)

        _Ended = Now
        If Request = Type.DDL Then
            With DDL
                _Succeeded = .Response.Succeeded
                Responses.Add(.Response)
                RaiseEvent Completed(Me, New ResponsesEventArgs(.Response))
            End With

        ElseIf Request = Type.SQL Then
            With SQL
                _Succeeded = .Response.Succeeded
                Responses.Add(.Response)
                RaiseEvent Completed(Me, New ResponsesEventArgs(.Response))
            End With

        End If

    End Sub
    Private Sub Requests_Completed(sender As Object, e As ResponsesEventArgs)

        _Ended = Now
        With ETL
            _Succeeded = .Succeeded
            Responses.AddRange(.Responses)
            RaiseEvent Completed(Me, New ResponsesEventArgs(.Responses))
        End With

    End Sub
#End Region
End Class
'▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
<ComVisible(False)> Public Class SQL
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
            Handle.Dispose()
            ' Free any other managed objects here.
            _Table.Dispose()
            If rf IsNot Nothing Then rf.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public Event Completed(sender As Object, e As ResponseEventArgs)
    Public ReadOnly Property ConnectionString As String
    Public ReadOnly Property Instruction As String
    Public ReadOnly Property Table As New DataTable
    Public ReadOnly Property Elapsed As TimeSpan
        Get
            Return Ended - Started
        End Get
    End Property
    Public ReadOnly Property Started As Date
    Public ReadOnly Property Busy As Boolean
    Public ReadOnly Property Ended As Date
    Public ReadOnly Property Response As ResponseEventArgs
    Public ReadOnly Property Connection As Connection
    Public ReadOnly Property Status As TriState
        Get
            If Response Is Nothing Then
                Return TriState.UseDefault
            Else
                If Response.Succeeded Then
                    Return TriState.True
                Else
                    Return TriState.False
                End If
            End If
        End Get
    End Property
    Public Property Name As String
    Public Property Tag As Object
    Private rf As ResponseFailure
    Public Sub New(ConnectionString As String, Instruction As String)

        Me.ConnectionString = If(ConnectionString, String.Empty)
        Me.Instruction = If(Instruction, String.Empty)
        Connection = New Connection(ConnectionString)

    End Sub
    Public Sub New(Connection As Connection, Instruction As String)

        ConnectionString = If(Connection Is Nothing, String.Empty, Connection.ToString)
        Me.Instruction = If(Instruction, String.Empty)
        Connection = New Connection(ConnectionString)

    End Sub
    Public Sub Execute(Optional RunInBackground As Boolean = True)

        _Started = Now
        _Busy = True
        If RunInBackground Then
            With New BackgroundWorker
                AddHandler .DoWork, AddressOf Execute
                AddHandler .RunWorkerCompleted, AddressOf Executed
                .RunWorkerAsync()
            End With
        Else
            Execute(Nothing, Nothing)
            Executed(Nothing, Nothing)
        End If

    End Sub
    Private Sub Execute(sender As Object, e As DoWorkEventArgs)

        If sender IsNot Nothing Then RemoveHandler DirectCast(sender, BackgroundWorker).DoWork, AddressOf Execute

        If ConnectionString.Any And Instruction.Any Then
#Region " BACKUP "
            If IsFile(ConnectionString) Then
                Dim Extension As String = Split(ConnectionString, ".").Last
#Region " /// COMPUTER FILE"
                Select Case True
                    Case Extension = "txt"
#Region " /// TEXTFILE "
                        _Table = TextFileToDataTable(ConnectionString)
                        _Ended = Now
                        If _Table Is Nothing Then
                            _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, New System.Text.StringBuilder("Text file conversion failed").ToString, Nothing)
                        Else
                            _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, Table, Ended - Started)
                        End If
#End Region
                    Case Regex.Match(Extension, "xl[a-z]{1,2}", RegexOptions.IgnoreCase).Success
#Region " /// EXCELFILE * NEED CODE TO READ AN EXCEL FILE WITHOUT A READER "
                        Try
                            Dim ExcelConnectionACE As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ConnectionString & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1;"""
                            Dim ExcelConnectionJet As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ConnectionString & ";Extended Properties=""Excel 8.0;HDR=Yes;"""

                            Dim Filter As String = Split(ConnectionString, ".").Last
                            Dim ExcelConnectionString As String = If(Filter = "xls", ExcelConnectionJet, ExcelConnectionACE)

                            Using ExcelConnection As New OleDbConnection(ExcelConnectionString)
                                Try
                                    Using Adapter As New OleDbDataAdapter(Instruction, ExcelConnection)
                                        Adapter.Fill(Table)
                                        Table.Locale = CultureInfo.InvariantCulture
                                        _Ended = Now
                                        _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, Table, Ended - Started)

                                    End Using
                                Catch ex As OleDbException
                                    _Ended = Now
                                    _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, ex.Message, New Errors(ex.Message))

                                End Try
                            End Using

                        Catch ex As OleDbException
                            _Ended = Now
                            _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, ex.Message, New Errors(ex.Message))

                        End Try
#End Region
                    Case Else
                        _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, New System.Text.StringBuilder("Unknown source").ToString, Nothing)

                End Select
#End Region
            Else
#Region " DATABASE "
                Using Connection As New OdbcConnection(ConnectionString)
                    Try
                        Connection.Open()
                        Try
                            'Using someCommand As New SqlCommand()
                            '    someCommand.Parameters.Add("@username", SqlDbType.NChar).Value = Name
                            'End Using
                            Using Adapter As New OdbcDataAdapter(Instruction, Connection)
                                Adapter.Fill(Table)
                                Connection.Close()
                                _Ended = Now
                                Table.Locale = CultureInfo.InvariantCulture
                                Table.Namespace = "<DB2>"
                                _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, Table, Ended - Started)

                            End Using
                        Catch odbcException As OdbcException
                            Connection.Close()
                            _Ended = Now
                            _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, odbcException.Message, New Errors(odbcException.Message))

                        End Try

                    Catch odbcOpenException As OdbcException
                        Connection.Close()
                        _Ended = Now
                        _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, odbcOpenException.Message, New Errors(odbcOpenException.Message))

                    End Try
                End Using
#End Region
            End If
#End Region
        Else
            Dim MissingMessage As New System.Text.StringBuilder
            If ConnectionString.Length = 0 Then
                MissingMessage.Append("Missing connection")
                If Instruction.Length = 0 Then MissingMessage.Append("and instruction")
            Else
                MissingMessage.Append("Missing instruction")
            End If
            _Response = New ResponseEventArgs(InstructionType.SQL, ConnectionString, Instruction, MissingMessage.ToString, New Errors(MissingMessage.ToString))
        End If

    End Sub
    Private Sub Executed(sender As Object, e As RunWorkerCompletedEventArgs)

        If sender IsNot Nothing Then RemoveHandler DirectCast(sender, BackgroundWorker).RunWorkerCompleted, AddressOf Executed
        _Busy = False
        RaiseEvent Completed(Me, Response)

    End Sub
    Private Sub Me_Completed(sender As Object, e As ResponseEventArgs) Handles Me.Completed
        _Busy = False
        If Not e.Succeeded Then rf = New ResponseFailure(e)
    End Sub
    Public Overrides Function ToString() As String
        Return If(Name Is Nothing, String.Empty, Name & BlackOut) & Join({If(Connection Is Nothing, "DSN=?", Connection.DataSource), If(Response Is Nothing, "Not executed", "Succeeded=" & Response.Succeeded)}, BlackOut)
    End Function
End Class
Public Class DDL
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
            Handle.Dispose()
            ' Free any other managed objects here.
            If rf IsNot Nothing Then rf.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public Event Completed(sender As Object, e As ResponseEventArgs)
    Public ReadOnly Property ConnectionString As String
    Public ReadOnly Property Instruction As String
    Public ReadOnly Property RequiresInput As Boolean
    Public ReadOnly Property GetRowCount As Boolean
    Public ReadOnly Property Elapsed As TimeSpan
        Get
            Return Ended - Started
        End Get
    End Property
    Public ReadOnly Property Started As Date
    Public ReadOnly Property Busy As Boolean
        Get
            Return Started <> New Date And Ended = New Date
        End Get
    End Property
    Public ReadOnly Property Ended As Date
    Public ReadOnly Property Response As ResponseEventArgs
    Public ReadOnly Property Status As TriState
        Get
            If Response Is Nothing Then
                Return TriState.UseDefault
            Else
                If Response.Succeeded Then
                    Return TriState.True
                Else
                    Return TriState.False
                End If
            End If
        End Get
    End Property
    Public Property Name As String
    Public Property Tag As Object
    Public ReadOnly Property Procedures As List(Of Procedure)
        Get
            If Regex.Match(Instruction, "(CREATE|ALTER|DROP)(\s{1,}OR REPLACE){0,1}\s{1,}(FUNCTION|PROCEDURE|TRIGGER)[\s]{1,}", RegexOptions.IgnoreCase).Success Then
                'Functions, Procs, and Triggers contain DDL within separated by ;
                Return {New Procedure(Instruction)}.ToList
            Else
                Return Split(Instruction, ";").Select(Function(p) New Procedure(p)).ToList
            End If
        End Get
    End Property
    Public ReadOnly Property ProceduresOK As New List(Of Procedure)
    Private rf As ResponseFailure
    Public Sub New(ConnectionString As String, Instruction As String, Optional PromptForInput As Boolean = False, Optional GetRowCount As Boolean = False)

        If ConnectionString IsNot Nothing Then
            Me.ConnectionString = ConnectionString
            Me.Instruction = Instruction
            RequiresInput = PromptForInput
            Me.GetRowCount = GetRowCount
            If RequiresInput Then
                GetInput()
            Else
                ProceduresOK.AddRange(Procedures)
            End If
        End If

    End Sub
    Public Sub New(Connection As Connection, Instruction As String, Optional PromptForInput As Boolean = False, Optional GetRowCount As Boolean = False)

        If Connection IsNot Nothing Then
            ConnectionString = Connection.ToString
            Me.Instruction = Instruction
            RequiresInput = PromptForInput
            Me.GetRowCount = GetRowCount
            If RequiresInput Then
                GetInput()
            Else
                ProceduresOK.AddRange(Procedures)
            End If
        End If

    End Sub
    Private Sub GetInput()

        Dim PromptProcedures As New List(Of Procedure)(Procedures)
        If Not PromptProcedures.Any Then Stop
        ProceduresOK.Clear()
        For Each Procedure In PromptProcedures
            Dim TableRowCount As Integer = -1
            Dim alterRow As DataRow = Nothing
            If Procedure.FetchStatement IsNot Nothing Then
                With New SQL(ConnectionString, Procedure.FetchStatement)
                    .Execute()
                    Do While .Response Is Nothing
                    Loop
                    If .Response.Succeeded Then
                        If .Response.Columns = 1 Then
                            TableRowCount = Convert.ToInt32(.Response.Value, InvariantCulture)
                        Else
                            If .Table.Rows.Count > 0 Then alterRow = .Table.Rows(0)
                        End If
                    End If
                    Procedure.Fetches.Add(New KeyValuePair(Of String, Integer)(Procedure.ObjectName, TableRowCount))
                End With
            End If
            With Procedure
                Dim PromptMessage As String = If(.ObjectAction = Procedure.Action.Drop,
                        Join({"You are about to Drop", .ObjectType.ToString, .ObjectName}),
                        Join({"You are about to", .ObjectAction.ToString, TableRowCount, "Rows in Table", .ObjectName}))
                If alterRow IsNot Nothing Then
                    Dim requestString As String() = Split(.ObjectName, BlackOut)
                    Dim tableName As String = requestString.First
                    Dim columnName As String = requestString(1)
                    Dim newDataType As String = requestString(2)
                    Dim currentDataType As String = alterRow("COLUMN_DATA").ToString
                    PromptMessage = Join({"You are about to Alter Column", columnName, "in table", tableName, "from", currentDataType, "to", newDataType})
                End If
                Using message As New Prompt
                    .Execute = message.Show("Proceed?", PromptMessage, Prompt.IconOption.YesNo) = DialogResult.Yes
                End Using
                If .Execute Then ProceduresOK.Add(Procedure)
            End With
        Next
        If Not ProceduresOK.Any Then
            _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, Instruction, Nothing, Ended - Started)
        End If

    End Sub
    Public Sub Execute(Optional RunInBackground As Boolean = False)

        If ProceduresOK.Any Then
            If RunInBackground Then
                With New BackgroundWorker
                    AddHandler .DoWork, AddressOf Execute
                    AddHandler .RunWorkerCompleted, AddressOf Executed
                    .RunWorkerAsync()
                End With
            Else
                Execute(Nothing, Nothing)
                Executed(Nothing, Nothing)
            End If
        End If

    End Sub
    Private Sub Execute(sender As Object, e As DoWorkEventArgs)

        If IsFile(ConnectionString) Then
            _Started = Now
            Try
                Dim ExcelConnectionACE As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ConnectionString & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1;"""
                Dim ExcelConnectionJet As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ConnectionString & ";Extended Properties=""Excel 8.0;HDR=Yes;"""

                Dim Filter As String = Split(ConnectionString, ".").Last
                Dim ExcelConnectionString As String = If(Filter = "xls", ExcelConnectionJet, ExcelConnectionACE)
                Using ExcelConnection As New OleDbConnection(ExcelConnectionString)
                    Try
                        ExcelConnection.Open()
                        Using Command As New OleDbCommand(Instruction, ExcelConnection)
                            'Command.Parameters.AddWithValue("@username", SqlDbType.NChar).Value = String.Empty
                            'Command.Parameters.AddWithValue("@instruction", Instruction)
                            'UPDATE[`R US all billing$`] SET `Item Amt` = 0'
                            Command.ExecuteNonQuery()
                        End Using
                        _Ended = Now
                        _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, Instruction, Nothing, Ended - Started)

                    Catch ex As OleDbException
                        _Ended = Now
                        _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, Instruction, ex.Message, New Errors(ex.Message))

                    End Try
                End Using

            Catch ex As OleDbException
                _Ended = Now
                _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, Instruction, ex.Message, New Errors(ex.Message))

            End Try
        Else
            _Started = Now
            Using New CursorBusy
                Using _Connection As New OdbcConnection(ConnectionString)
                    Dim DDL_Instruction As String = Join((From ok In ProceduresOK Select ok.Instruction).ToArray, ";")
                    Try
                        _Connection.Open()
                        Using Command As New OdbcCommand(DDL_Instruction, _Connection)
                            'Command.Parameters.AddWithValue("@instruction", DDL_Instruction)
                            Try
                                Command.ExecuteNonQuery()
                                _Ended = Now
                                _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, DDL_Instruction, Nothing, Ended - Started)

                            Catch ProcedureError As OdbcException
                                _Ended = Now
                                _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, DDL_Instruction, ProcedureError.Message, New Errors(ProcedureError.Message))

                            End Try
                        End Using

                    Catch ODBC_RunError As OdbcException
                        _Ended = Now
                        _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, DDL_Instruction, ODBC_RunError.Message, New Errors(ODBC_RunError.Message))

                    End Try
                End Using
            End Using
        End If

    End Sub
    Private Sub Executed(sender As Object, e As RunWorkerCompletedEventArgs)

        If sender IsNot Nothing Then RemoveHandler DirectCast(sender, BackgroundWorker).RunWorkerCompleted, AddressOf Executed
        RaiseEvent Completed(Me, Response)

    End Sub
    Private Sub Me_Completed(sender As Object, e As ResponseEventArgs) Handles Me.Completed
        If Not e.Succeeded Then rf = New ResponseFailure(e)
    End Sub
End Class
Public NotInheritable Class Procedure
    Public Sub New(Instruction As String)

        Me.Instruction = If(Instruction, String.Empty)
        Dim Match_Drop As Match = Regex.Match(Instruction, "DROP[\s]{1,}(TABLE|VIEW|FUNCTION|TRIGGER)[\s]{1,}" & ObjectPattern, RegexOptions.IgnoreCase)
        If Match_Drop.Success Then
#Region " DROP OBJECT REQUEST "
            ObjectAction = Action.Drop
            Dim dropObject As String = Regex.Match(Match_Drop.Value, "TABLE|VIEW|FUNCTION|TRIGGER", RegexOptions.IgnoreCase).Value
            ObjectType = ParseEnum(Of Type)(dropObject)
            ObjectName = Regex.Match(Instruction, "(?<=TABLE|VIEW|FUNCTION|TRIGGER)[\s]{1,}([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})", RegexOptions.IgnoreCase).Value
            ObjectName = Trim(Split(ObjectName, ".").Last)
#End Region
        Else
            Dim Match_Insert As Match = Regex.Match(Instruction, "INSERT[\s]{1,}INTO[\s]{1,}" + ObjectPattern + "([\s]{0,}\([A-Z0-9!%{}^~_@#$]{1,}(,[\s]{0,}[A-Z0-9!%{}^~_@#$]{1,}){0,}\)){0,}", RegexOptions.IgnoreCase)
            If Match_Insert.Success Then
#Region " INSERT ROWS REQUEST "
                ObjectAction = Action.Insert
                ObjectType = Type.Table
                ObjectName = Regex.Match(Match_Insert.Value, "INTO ([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})", RegexOptions.IgnoreCase).Value
                ObjectName = Split(ObjectName.Remove(0, 5), ".").Last
#End Region
            Else
                Dim Match_Update As Match = Regex.Match(Instruction, "UPDATE[\s]{1,}" + ObjectPattern + "([\s]{1,}([A-Z0-9!%{}^~_@#$]{1,})){0,1}[\s]{1,}SET[\s]{1,}", RegexOptions.IgnoreCase)
                If Match_Update.Success Then
#Region " UPDATE FIELDS REQUEST "
                    ObjectAction = Action.Update
                    ObjectType = Type.Table
                    Dim Elements As String() = Split(Match_Update.Value, " ")
                    ObjectName = Elements(1)
                    _FetchStatement = "SELECT COUNT(*) C From " & ObjectName
                    '==================================
                    'UPDATE PROFILES
                    'SET (PASSWORD, XYZ)=SELECT ('W0W0W0W0', ...
                    'WHERE USERID ='Q085365'

                    'Becomes...

                    'SELECT COUNT(*) C
                    'From PROFILES
                    'Where USERID ='Q085365'
                    Dim Wheres = RegexMatches(Instruction, "where\s{1,}", RegexOptions.IgnoreCase)
                    If Wheres.Any Then
                        Dim FirstWhere As Integer = Wheres.First.Index
                        _FetchStatement &= Instruction.Substring(FirstWhere, Instruction.Length - FirstWhere)
                    End If
#End Region
                Else
                    Dim match_Alter As New List(Of String)(Regex.Split(Instruction, "ALTER\s+TABLE\s+|\s+ALTER\s+COLUMN\s+|\s+SET\s+DATA\s+TYPE\s+", RegexOptions.IgnoreCase).Skip(1))
                    If match_Alter.Any Then
                        'ALTER TABLE C085365.ACTIONS_EXTRA ALTER COLUMN OA SET DATA TYPE VARCHAR(2003)
                        If match_Alter.Count = 3 Then
                            Dim tableName As String = match_Alter.First
                            Dim columnName As String = match_Alter(1)
                            Dim newDataType As String = match_Alter(2)
                            ObjectAction = Action.Alter
                            ObjectType = Type.Column
                            ObjectName = Join({tableName, columnName, newDataType}, BlackOut)
                            FetchStatement = Replace(Replace(My.Resources.SQL_ColumnTypes, "///OWNER_TABLE///", tableName), "--AND C.NAME='//COLUMN_NAME//'", "AND C.NAME=" & ValueToField(columnName))
                        End If

                    Else
                        Dim Match_Delete As Match = Regex.Match(Instruction, "DELETE[\s]{1,}FROM[\s]{1,}" + ObjectPattern, RegexOptions.IgnoreCase)
                        If Match_Delete.Success Then
#Region " DELETE ROWS REQUEST "
                            ObjectAction = Action.Delete
                            ObjectType = Type.Table
                            ObjectName = Regex.Replace(Match_Delete.Value, "DELETE[\s]{1,}FROM[\s]{1,}", String.Empty, RegexOptions.IgnoreCase)
                            ObjectName = Split(ObjectName, ".").Last
                            FetchStatement = Instruction.Remove(Match_Delete.Index, "DELETE".Length)
                            FetchStatement = FetchStatement.Insert(Match_Delete.Index, "SELECT COUNT(*) C ")
#End Region
                        Else
                            Dim Match_GrantRevoke As Match = Regex.Match(Instruction, "(GRANT|REVOKE)\s{1,}(SELECT|INSERT|ALTER|UPDATE|DELETE)\s{1,}" + ObjectPattern, RegexOptions.IgnoreCase)
                            If Match_GrantRevoke.Success Then
#Region " GRANT Or REVOKE REQUEST "
                                ObjectAction = If(Match_GrantRevoke.Value.ToUpperInvariant.StartsWith("GRANT", StringComparison.InvariantCulture), Action.Grant, Action.Revoke)
                                ObjectType = Type.Table
                                ObjectName = Regex.Replace(Match_GrantRevoke.Value, "(GRANT|REVOKE)\s{1,}(SELECT|INSERT|ALTER|UPDATE|DELETE)\s{1,}", String.Empty, RegexOptions.IgnoreCase)
                                ObjectName = Split(ObjectName, ".").Last
#End Region
                            End If

                        End If
                    End If
                End If
            End If
        End If
        If FetchStatement IsNot Nothing Then
            'With New SQL("", FetchStatement)

            'End With
        End If

    End Sub
    Public ReadOnly Property Instruction As String
    Public ReadOnly Property ObjectName As String
    Public ReadOnly Property ObjectAction As Action
    Public ReadOnly Property ObjectType As Type
    Public Property Execute As Boolean = True
    Public Enum Action
        Delete
        Insert
        Update
        Alter
        Create
        Drop
        Grant
        Revoke
        CommentOn
    End Enum
    Public Enum Type
        NickName    'Alias
        Utility     'Function
        Mask
        Permission
        Procedure
        Sequence
        Table
        Column
        Trigger
        View
        Index
        Schema
    End Enum
    Public ReadOnly Property RowCount As Integer
    Public ReadOnly Property FetchStatement As String
    Public ReadOnly Property Fetches As New List(Of KeyValuePair(Of String, Integer))
End Class
Public Class ETL
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
            Handle.Dispose()
            ' Free any other managed objects here.
            Sources.Dispose()
            Destinations.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public Event Completed(sender As Object, e As ResponsesEventArgs)
    Public Sub New()
    End Sub
    Public Property Name As String
    Public Property Description As String
    Public ReadOnly Property Elapsed As TimeSpan
        Get
            Return Ended - Started
        End Get
    End Property
    Public ReadOnly Property Started As Date
    Public ReadOnly Property Busy As Boolean
        Get
            Return Started <> New Date And Ended = New Date
        End Get
    End Property
    Public ReadOnly Property Ended As Date
    Private WithEvents Sources_ As New SourceCollection(Me)
    Public ReadOnly Property Sources As SourceCollection
        Get
            Return Sources_
        End Get
    End Property
    Private WithEvents Destinations_ As New DestinationCollection(Me)
    Public ReadOnly Property Destinations As DestinationCollection
        Get
            Return Destinations_
        End Get
    End Property
    Public ReadOnly Property Responses As New List(Of ResponseEventArgs)
    Public ReadOnly Property Succeeded As Boolean = False
    Public Class SourceCollection
        Inherits List(Of Source)
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
                Handle.Dispose()
                ' Free any other managed objects here.
                Data.Dispose()
                For Each source In Me
                    source.Dispose()
                Next
            End If
            disposed = True
        End Sub
#End Region
        Public ReadOnly Property Parent As ETL
        Public Event Completed(sender As Object, e As ResponsesEventArgs)
        Public ReadOnly Property Responses As New List(Of ResponseEventArgs)
        Public ReadOnly Property Data As New DataSet
        Public Sub New(ExtractTransformLoad As ETL)
            Parent = ExtractTransformLoad
        End Sub
        Public Shadows Function Add(Item As Source) As Source

            If Item IsNot Nothing Then
                MyBase.Add(Item)
                Item.Parent = Me
                AddHandler Item.Retrieved, AddressOf Source_Completed
            End If
            Return Item

        End Function
        Public Shadows Function Add(SQL As SQL) As Source

            If SQL IsNot Nothing Then
                Dim SourceItem As Source = New Source(SQL)
                MyBase.Add(SourceItem)
                SourceItem.Parent = Me
                AddHandler SourceItem.Retrieved, AddressOf Source_Completed
                Return SourceItem
            Else
                Return Nothing
            End If

        End Function
        Public Shadows Function Add(Table As DataTable) As Source

            If Table IsNot Nothing Then
                Dim SourceItem As Source = New Source(Table)
                MyBase.Add(SourceItem)
                SourceItem.Parent = Me
                AddHandler SourceItem.Retrieved, AddressOf Source_Completed
                Return SourceItem
            Else
                Return Nothing
            End If

        End Function
        Private Sub Source_Completed(sender As Object, e As ResponseEventArgs)

            With DirectCast(sender, Source)
                RemoveHandler .Retrieved, AddressOf Source_Completed
            End With
            Responses.Add(e)
            If Where(Function(s) s.SQL Is Nothing OrElse s.SQL.Response IsNot Nothing).Count = Count Then
                RaiseEvent Completed(Me, New ResponsesEventArgs(Responses))
            End If

        End Sub
    End Class
    Public Class Source
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
                Handle.Dispose()
                ' Free any other managed objects here.
                If _Table IsNot Nothing Then _Table.Dispose()
                If _SQL IsNot Nothing Then _SQL.Dispose()
            End If
            disposed = True
        End Sub
#End Region
        Public Event Retrieved(sender As Object, e As ResponseEventArgs)
        Public Sub New(ConnectionString As String, Instruction As String)
            SQL = New SQL(ConnectionString, Instruction)
        End Sub
        Public Sub New(Connection As Connection, Instruction As String)
            SQL = New SQL(Connection, Instruction)
        End Sub
        Public Sub New(Table As DataTable)
            Me.Table = Table
        End Sub
        Public Sub New(SQL As SQL)
            Me.SQL = SQL
        End Sub
        Friend Parent As SourceCollection
        Public Property Name As String
        Public ReadOnly Property SQL As SQL
        Public ReadOnly Property Started As Date
        Public ReadOnly Property Ended As Date
        Public ReadOnly Property Table As DataTable
        Public Sub Retrieve()

            _Started = Now
            If Table Is Nothing Then
                AddHandler SQL.Completed, AddressOf SQL_Completed
                SQL.Execute()

            Else
                'Section for when the Datatable was already provided and does not need retrieval
                Parent.Data.Tables.Add(Table.Copy)
                _Ended = Now
                If SQL Is Nothing Then
                    RaiseEvent Retrieved(Me, New ResponseEventArgs(InstructionType.SQL, String.Empty, String.Empty, Table, Ended - Started))
                Else
                    RaiseEvent Retrieved(Me, New ResponseEventArgs(InstructionType.SQL, SQL.ConnectionString, SQL.Instruction, Table, Ended - Started))
                End If

            End If

        End Sub
        Private Sub SQL_Completed(sender As Object, e As ResponseEventArgs)

            RemoveHandler SQL.Completed, AddressOf SQL_Completed
            If e.Succeeded Then
                _Table = e.Table
                Parent.Data.Tables.Add(Table.Copy)
            End If
            _Ended = Now
            RaiseEvent Retrieved(Me, e)

        End Sub
    End Class
    Public Class DestinationCollection
        Inherits List(Of Destination)
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
                Handle.Dispose()
                ' Free any other managed objects here.
                For Each destination In Me
                    destination.Dispose()
                Next
            End If
            disposed = True
        End Sub
#End Region
        Public Event Completed(sender As Object, e As ResponsesEventArgs)
        Friend ReadOnly Parent As ETL
        Public ReadOnly Property Responses As New List(Of ResponseEventArgs)
        Friend ResponseCount As Integer
        Public Sub New(ExtractTransformLoad As ETL)
            Parent = ExtractTransformLoad
        End Sub
        Public Shadows Function Add(Item As Destination) As Destination

            If Item IsNot Nothing Then
                MyBase.Add(Item)
                Item.Parent = Me
                AddHandler Item.Completed, AddressOf Destination_Completed
            End If
            Return Item

        End Function
        Public Shadows Function Add(Location As String) As Destination

            Dim Item = New Destination(Location)
            MyBase.Add(Item)
            Item.Parent = Me
            AddHandler Item.Completed, AddressOf Destination_Completed
            Return Item

        End Function
        Private Sub Destination_Completed(sender As Object, e As ResponsesEventArgs)

            With DirectCast(sender, Destination)
                RemoveHandler .Completed, AddressOf Destination_Completed
            End With
            Responses.AddRange(e.Responses)
            ResponseCount += 1
            If Count = ResponseCount Then
                RaiseEvent Completed(Me, New ResponsesEventArgs(Responses))
            End If

        End Sub
    End Class
    Public Class Destination
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
                Handle.Dispose()
                ' Free any other managed objects here.
                If Table IsNot Nothing Then Table.Dispose()
                If Message IsNot Nothing Then Message.Dispose()
                If DDL IsNot Nothing Then _DDL.Dispose()
            End If
            disposed = True
        End Sub
#End Region
        Public Event Completed(sender As Object, e As ResponsesEventArgs)
        Public Event BlockInserted(sender As Object, e As ResponseEventArgs)
        Public Sub New(Location As String)
            ConnectionString = Location
        End Sub
        Public Sub New(Connection As Connection, TableName As String)

            Dim Connections As New ConnectionCollection
            Me.Connection = Connection
            If Connection IsNot Nothing And TableName IsNot Nothing Then
                ConnectionString = Connection.ToString
                Me.TableName = TableName.ToUpperInvariant
            End If
            CreateTable = False

        End Sub
        Public Sub New(ConnectionString As String, TableName As String)

            Dim Connections As New ConnectionCollection
            Me.ConnectionString = ConnectionString
            If ConnectionString IsNot Nothing And TableName IsNot Nothing Then
                Connection = Connections.Item(ConnectionString)
                Me.TableName = TableName.ToUpperInvariant
            End If
            CreateTable = False

        End Sub
        Public Sub New(Connection As Connection, TableSpace As String, TableName As String)

            Dim Connections As New ConnectionCollection
            Me.Connection = Connection
            If Connection IsNot Nothing And TableSpace IsNot Nothing And TableName IsNot Nothing Then
                ConnectionString = Connection.ToString
                Me.TableSpace = TableSpace
                Me.TableName = TableName.ToUpperInvariant
                CreateTable = True
            End If

        End Sub
        Public Sub New(ConnectionString As String, TableSpace As String, TableName As String)

            Dim Connections As New ConnectionCollection
            Me.ConnectionString = ConnectionString
            If ConnectionString IsNot Nothing And TableSpace IsNot Nothing And TableName IsNot Nothing Then
                Connection = Connections.Item(ConnectionString)
                Me.TableSpace = TableSpace
                Me.TableName = TableName.ToUpperInvariant
                CreateTable = True
            End If

        End Sub
        Friend Parent As DestinationCollection
        Public Property Name As String
        Public ReadOnly Property DDL As DDL
        Public ReadOnly Property Blocks As New List(Of ResponseEventArgs)
        Public ReadOnly Property Started As Date
        Public ReadOnly Property Ended As Date
        Public ReadOnly Property ConnectionString As String
        Public ReadOnly Property Connection As Connection
        Public ReadOnly Property TableName As String
        Public ReadOnly Property TableSpace As String
        Public ReadOnly Property CreateTable As Boolean
        Public ReadOnly Property Columns As Dictionary(Of String, ColumnProperties)
        Public Property ClearTable As Boolean = True
        Public ReadOnly Property Table As DataTable
            Get
                Dim ConsolidatedData As New DataTable
                Dim Sources = Parent.Parent.Sources.Data
                For Each Table In Sources.Tables
                    ConsolidatedData.Merge(Table)
                Next
                Return ConsolidatedData
            End Get
        End Property
        Public ReadOnly Property TableDDL() As String
            Get
                Dim DDL As New List(Of String) From {
                    Join({"CREATE TABLE", TableName, "("})
                }
                For Each Column In _Columns
                    Dim Comma As String = If(Column.Value.Index = 0, String.Empty, ", ")
                    DDL.Add(Comma + DB2ColumnNamingConvention(Column.Key.ToUpperInvariant) & StrDup(6, vbTab) & Column.Value.DataFormat)
                Next Column
                DDL.Add(")")
                If If(TableSpace, String.Empty).Length > 0 Then DDL.Add(" IN " & TableSpace)
                Return Join(DDL.ToArray, vbNewLine)
            End Get
        End Property
        Private ReadOnly Message As New Prompt
        Friend Sub Fill()

            _Started = Now
            If IsFile(ConnectionString) Then
                If GetFileNameExtension(ConnectionString).Value = ExtensionNames.Text Then
                    DataTableToTextFile(Table, ConnectionString)
                ElseIf GetFileNameExtension(ConnectionString).Value = ExtensionNames.Excel Then
                    DataTableToExcel(Table, ConnectionString, False, False, False, True, True)
                Else
                End If
                Dim Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, String.Empty, Table, Now - Started)
                RaiseEvent Completed(Me, New ResponsesEventArgs(Response))

            ElseIf Table Is Nothing Then
                Dim Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, String.Empty, "No Table".ToString(InvariantCulture), Nothing)
                RaiseEvent Completed(Me, New ResponsesEventArgs(Response))

            ElseIf Table.Columns.Count = 0 Then
                Dim Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, String.Empty, "No Columns".ToString(InvariantCulture), Nothing)
                RaiseEvent Completed(Me, New ResponsesEventArgs(Response))

            ElseIf Table.Rows.Count = 0 Then
                Dim Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, String.Empty, Table, Now - Started)
                RaiseEvent Completed(Me, New ResponsesEventArgs(Response))

            Else
                If CreateTable Then
                    ClearTable = False
                    Create_Table()

                ElseIf ClearTable Then
                    Clear_Table()
                Else
                    Table_Structure()
                End If
            End If

        End Sub
#Region " CREATE TABLE "
        Private Sub Create_Table()

            Using newTable As DataTable = Table.Copy
                newTable.TableName = "newTable"
                _Columns = DataTableToListOfColumnProperties(newTable).ToDictionary(Function(x) x.Name, Function(y) y)
            End Using
            With New DDL(ConnectionString, TableDDL)
                AddHandler .Completed, AddressOf Table_CreateResponded
                .Execute()
            End With

        End Sub
        Private Sub Table_CreateResponded(sender As Object, e As ResponseEventArgs)

            _DDL = DirectCast(sender, DDL)
            RemoveHandler DDL.Completed, AddressOf Table_CreateResponded
            If e.Succeeded Then
                DataTable_DB2()
            Else
                'Exit now, can't move forward
                RaiseEvent Completed(Me, New ResponsesEventArgs(e))
            End If

        End Sub
#End Region
#Region " CLEAR TABLE "
        Private Sub Clear_Table()

            With New DDL(ConnectionString, "DELETE FROM " & TableName)
                AddHandler .Completed, AddressOf Table_ClearResponded
                .Execute()
            End With

        End Sub
        Private Sub Table_ClearResponded(sender As Object, e As ResponseEventArgs)

            _DDL = DirectCast(sender, DDL)
            RemoveHandler DDL.Completed, AddressOf Table_ClearResponded
            If e.Succeeded Then
                Table_Structure()
            Else
                'Exit now, can't move forward
                RaiseEvent Completed(Me, New ResponsesEventArgs(e))
            End If

        End Sub
#End Region
#Region " GET TABLE STRUCTURE FROM DB2 "
        Private Sub Table_Structure()

            Dim Instruction As String = ColumnSQL(TableName)
            With New SQL(ConnectionString, Instruction)
                AddHandler .Completed, AddressOf Table_StructureResponded
                .Execute()
            End With

        End Sub
        Private Sub Table_StructureResponded(sender As Object, e As ResponseEventArgs)

            With DirectCast(sender, SQL)
                RemoveHandler .Completed, AddressOf Table_StructureResponded
                If e.Succeeded Then
                    _Columns = DataTableToListOfColumnsProperties(.Table).ToDictionary(Function(x) x.Name, Function(y) y)
                    If Columns.Any Then
                        DataTable_DB2()
                    Else
                        'Exit now, can't move forward
                        Dim Response = New ResponseEventArgs(InstructionType.SQL, .ConnectionString, .Instruction, "No match in Database".ToString(InvariantCulture), Nothing)
                        RaiseEvent Completed(Me, New ResponsesEventArgs(Response))
                    End If
                Else
                    'Exit now, can't move forward
                    RaiseEvent Completed(Me, New ResponsesEventArgs(e))
                End If
            End With

        End Sub
#End Region
        Private Sub DataTable_DB2()

#Region " COLUMN SIZING - STRING...CHAR IF MIN=MAX, ELSE VARCHAR"
            'Select Case MIN(LENGTH(Trim(ENTERPRISE_NBR))) X
            ', MAX(LENGTH(TRIM(ENTERPRISE_NBR))) Y
            'From METRICS.AR_CA_S1F S
            'LIMIT 10
#End Region

            Dim SourceColumns As New List(Of DataColumn)
            For Each Column As DataColumn In Table.Columns
                Column.ColumnName = Column.ColumnName.ToUpper(CultureInfo.InvariantCulture)
                SourceColumns.Add(Column)
            Next

            Dim DestinationColumnIndices = Columns.Values.ToDictionary(Function(x) x.Index, Function(y) y)
            Dim IsZeroBased As Boolean = DestinationColumnIndices.Values.Select(Function(i) i.Index).Min = 0

            Dim ColumnTable As New DataTable
            With ColumnTable
                .Columns.Add(New DataColumn With {.ColumnName = "Index", .DataType = GetType(Integer)})
                .Columns.Add(New DataColumn With {.ColumnName = "Source Column", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "Source Type", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "Destination Column", .DataType = GetType(String)})
                .Columns.Add(New DataColumn With {.ColumnName = "Destination Type", .DataType = GetType(String)})
            End With

            Dim ColumnsParity As New List(Of ColumnParity)
            Dim ColumnsDifferentName As New Dictionary(Of String, String)
            Dim ColumnsOutOfSequence As New List(Of String)

            For ColumnIndex = 1 To {SourceColumns.Count, Table.Columns.Count}.Max
                Dim SourceIndex As Integer = ColumnIndex - 1
                Dim DestinationIndex As Integer = ColumnIndex - If(IsZeroBased, 1, 0)
                Dim CP As New ColumnParity With {.Index = ColumnIndex,
                    .SourceName = String.Empty,
                    .SourceType = GetType(String),
                    .DestinationName = String.Empty,
                    .DestinationType = String.Empty}
#Region " NAMES + TYPES "
                If SourceIndex >= 0 And SourceIndex < SourceColumns.Count Then
                    CP.SourceName = SourceColumns(SourceIndex).ColumnName
                    CP.SourceType = SourceColumns(SourceIndex).DataType
                Else
                    CP.SourceName = "X" & ColumnIndex
                    CP.SourceType = GetType(Object)
                End If
                If DestinationColumnIndices.ContainsKey(DestinationIndex) Then
                    CP.DestinationName = DestinationColumnIndices(DestinationIndex).Name
                    CP.DestinationType = DestinationColumnIndices(DestinationIndex).DataFormat
                Else
                    CP.DestinationName = "X" & ColumnIndex
                    CP.DestinationName = GetType(Object).ToString & ColumnIndex
                End If
#End Region
                ColumnsParity.Add(CP)
                ColumnTable.Rows.Add(CP.ToArray)
                If CP.SourceName <> CP.DestinationName Then ColumnsDifferentName.Add(CP.SourceName, CP.DestinationName)
            Next
            Dim SourceNames = ColumnsParity.Select(Function(c) c.SourceName).ToList
            Dim DestinationNames = ColumnsParity.Select(Function(c) c.DestinationName).ToList
            For Each Column As ColumnParity In ColumnsParity
                Dim SourceIndex As Integer = SourceNames.IndexOf(Column.SourceName)
                Dim DestinationIndex As Integer = DestinationNames.IndexOf(Column.DestinationName)
                If SourceIndex <> DestinationIndex Then
                    ColumnsOutOfSequence.Add(Column.SourceName & "@" & SourceIndex & "|" & Column.DestinationName & "@" & DestinationIndex)
                End If
            Next

            Dim CanProceed As Boolean = SourceColumns.Count = Columns.Count
            If CanProceed Then
                CanProceed = Not ColumnsOutOfSequence.Any
                If ColumnsDifferentName.Any Then CanProceed = Message.Show("Datasource Columns names are not present in the destination Table.", "Select Yes to continue or No to cancel.", Prompt.IconOption.YesNo) = DialogResult.Yes

            Else
                REM /// CAN NOT PROCEED - SOURCE COLUMNS COUNT MUST EQUAL DESTINATION COLUMNS COUNT...FOR NOW
                Dim MissingInDestinationTable As New List(Of String)(From CDN In ColumnsDifferentName.Keys Where CDN.Length > 0 And Not Columns.ContainsKey(CDN) Select CDN)
                Dim MissingInSourceTable As New List(Of String)(From CDN In ColumnsDifferentName.Values Where CDN.Length > 0 And Not Table.Columns.Contains(CDN) Select CDN)
                Dim WiderTable As String = String.Empty
                Dim NarrowerTable As String = String.Empty
                If SourceColumns.Count > Columns.Count Then
                    WiderTable = "Source DataTable"
                    NarrowerTable = "Destination DB2 Table (" & TableName & ")"
                Else
                    WiderTable = "Destination DB2 Table (" & TableName & ")"
                    NarrowerTable = "Source DataTable"
                End If
                Message.Datasource = ColumnTable
                Message.Show("Insert cancelled",
                         Join({"The number of columns in the", WiderTable, "exceeds that of the", NarrowerTable}),
                         Prompt.IconOption.Critical,
                         Prompt.StyleOption.Earth)
                Dim Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, String.Empty, Join({"The number of columns in the", WiderTable, "exceeds that of the", NarrowerTable}), Nothing)
                RaiseEvent Completed(Me, New ResponsesEventArgs(Response))
            End If
            If CanProceed Then
                Dim Rows As New Dictionary(Of Integer, List(Of Object))
                For Each Row As DataRow In Table.Rows
                    Dim Values As New List(Of Object)
                    'SourceName + SourceType + DestinationName
                    For Each Column As ColumnParity In ColumnsParity
                        Try     'If Table.Columns.Contains(Column.SourceName) Then **** SLOW!!! ( 250ms / Row )
                            Dim Value As Object = Row(Column.SourceName)
                            Dim EmptyValue As Boolean = IsDBNull(Value) Or Value.ToString.Length = 0
                            REM /// SOME DATA TYPES REQUIRE ' AROUND THE VALUE
                            If Columns.ContainsKey(Column.DestinationName) Then
                                Dim CP As ColumnProperties = Columns(Column.DestinationName)
                                With CP
                                    If EmptyValue AndAlso .Nulls = True Then
                                        Values.Add("CAST(NULL AS " & .DataFormat & ")")
                                    Else
                                        Select Case .DataType
                                            Case "CHAR", "VARCHAR", "LONG VARCHAR"
                                                If EmptyValue Then
                                                    Values.Add("''")

                                                Else
                                                    REM /// ToString is to account for Boolean which becomes "True" or "False"
                                                    Dim StringValue As String = Replace(Value.ToString, "'", "`")
                                                    If .DataType.Contains("VAR") Then StringValue = Trim(StringValue)
                                                    Values.Add("CAST('" & StringValue & "' As " & .DataFormat & ")")

                                                End If

                                            Case "BIGINT"
                                                Values.Add(Value)

                                            Case "DECIMAL", "SMALLINT", "INTEGER", "BIGINT", "DECFLOAT"
                                                REM /// NO FORMATTING NEEDED FOR NUMBERS
                                                If EmptyValue Then
                                                    Values.Add(0)
                                                Else
                                                    If Column.SourceType = GetType(Boolean) Then
                                                        Values.Add(If(Value.ToString.ToUpperInvariant = "FALSE", 0, 1))
                                                    Else
                                                        Values.Add(Value)
                                                        If .DataType = "BIGINT" Then Stop
                                                    End If
                                                End If

                                            Case "DATE"
                                                If EmptyValue Then
                                                    Values.Add("'1900-01-01'")
                                                Else
                                                    Values.Add(DateToDB2Date(Date.Parse(Value.ToString, InvariantCulture)))
                                                End If

                                            Case "TIMESTAMP"
                                                Dim DateValue As New Date
                                                Dim Success = Date.TryParse(Value.ToString, DateValue)
                                                If Success Then
                                                    Values.Add(DateToDB2Timestamp(DateValue))
                                                Else
                                                    Values.Add("'1900-01-01-00.00.00.000000'")
                                                End If

                                            Case "ROWID"
                                            Case "BLOB", "CLOB", "DBCLOB"
                                            Case "GRAPHIC", "LONG VARGRAPHIC", "VARGRAPHIC"
                                            Case "REAL"
                                            Case "FLOAT"
                                            Case "BINARY", "VARBINARY"
                                            Case Else
                                                Clipboard.SetText(.DataType)
                                                Stop

                                        End Select
                                    End If
                                End With
                            Else
                                Dim Message As String = Join({Column.SourceName, "not present in", TableName})
                                Dim Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, String.Empty, Message, Nothing)
                                RaiseEvent Completed(Me, New ResponsesEventArgs(Response))
                                Exit Sub
                            End If
                        Catch ex As KeyNotFoundException
                            Dim Message As String = Join({Column.SourceName, "not present in", TableName})
                            Dim Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, String.Empty, Message, Nothing)
                            RaiseEvent Completed(Me, New ResponsesEventArgs(Response))
                            Exit Sub
                        End Try
                    Next Column
                    Rows.Add(Rows.Count, Values)
                Next Row

                Dim Blocks = From R In Rows.Keys Select New With {.BlockNbr = QuotientRound(R, 254), .Select = "SELECT " + Join(Rows(R).ToArray, ",") + " FROM SYSIBM.SYSDUMMY1"}
                Dim Inserts = From B In Blocks Group B By BlockNbr = B.BlockNbr Into BlockGroup = Group Select New With {.Index = BlockNbr, .SQL = (From BG In BlockGroup Select BG.Select).ToArray}
                Dim BlockIndex As Integer

                For Each Block In Inserts
                    BlockIndex += 1
                    Dim Insert As String = "INSERT INTO " + TableName.ToUpperInvariant + vbNewLine + Join(Block.SQL, vbNewLine + "UNION ALL" + vbNewLine)
                    _DDL = New DDL(Connection, Insert)
                    With DDL
                        AddHandler .Completed, AddressOf Block_Completed
                        .Name = Join({BlockIndex, "of", Inserts.Count})
                        .Execute()
                    End With
                Next
            End If

        End Sub
        Private Sub Block_Completed(sender As Object, e As ResponseEventArgs)

            With DirectCast(sender, DDL)
                RemoveHandler .Completed, AddressOf Block_Completed
                Blocks.Add(e)
                RaiseEvent BlockInserted(sender, e)
                Dim Index As Integer = Integer.Parse(Split(.Name, " ").First, InvariantCulture)
                Dim Count As Integer = Integer.Parse(Split(.Name, " ").Last, InvariantCulture)
                If Index = Count Then
                    _Ended = Now
                    RaiseEvent Completed(Me, New ResponsesEventArgs(Blocks))
                End If
            End With

        End Sub
    End Class
    Public Sub Execute()
        Responses.Clear()
        _Started = Now
        RetrieveSourceData()
    End Sub
#Region " MULTIPLE SOURCES ==> ONE DESTINATION "
    Private Sub RetrieveSourceData()

        For Each Source In Sources
            Source.Retrieve()
        Next

    End Sub
    Private Sub ExportSourceData(sender As Object, e As ResponsesEventArgs) Handles Sources_.Completed

        For Each Destination In Destinations
            Destination.Fill()
        Next

    End Sub
    Private Sub SourceDataExported() Handles Destinations_.Completed

        Responses.AddRange(Sources.Responses)
        Responses.AddRange(Destinations.Responses)
        Dim Failures = Responses.Where(Function(r) Not r.Succeeded)
        _Succeeded = Not Failures.Any
        _Ended = Now
        RaiseEvent Completed(Me, New ResponsesEventArgs(Responses))

    End Sub
#End Region
End Class
Public Structure ColumnParity
    Implements IEquatable(Of ColumnParity)
    Public Property Index As Integer
    Public Property SourceName As String
    Public Property SourceType As Type
    Public Property DestinationName As String
    Public Property DestinationType As String
    Public Shadows Function ToArray() As String()
        Return {Index.ToString(InvariantCulture), SourceName, Convert.ToString(SourceType, InvariantCulture), DestinationName, DestinationType}
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return SourceName.GetHashCode Xor DestinationName.GetHashCode Xor DestinationType.GetHashCode Xor Index.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As ColumnParity) As Boolean Implements IEquatable(Of ColumnParity).Equals
        Return Index = other.Index AndAlso SourceName = other.SourceName
    End Function
    Public Shared Operator =(ByVal value1 As ColumnParity, ByVal value2 As ColumnParity) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(ByVal value1 As ColumnParity, ByVal value2 As ColumnParity) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is ColumnParity Then
            Return CType(obj, ColumnParity) = Me
        Else
            Return False
        End If
    End Function
End Structure
Public Structure ColumnProperties
    Implements IEquatable(Of ColumnProperties)
    Public Sub New(systemType As Type, Optional values As List(Of Object) = Nothing)

        Nulls = True
        If systemType = GetType(Date) Then
            'Might also be TIMESTAMP ... DataTable.DataType can't hold Systsem.DateAndTime, so if systemType comes from a DataTable then determine if it should be DATE vs TIMESTAMP
            Dim nonNullValues As New List(Of Object)(values?.Where(Function(v) Not (IsDBNull(v) Or v Is Nothing)))
            If nonNullValues.Any Then
                Dim valuesType As Type = GetDataType(nonNullValues)
                If valuesType = GetType(DateAndTime) Then
                    DataType = "TIMESTAMP"
                    Length = 10
                    Scale = 6
                Else
                    DataType = "DATE"
                    Length = 4
                End If
            Else
                DataType = "DATE"
                Length = 4
            End If
        Else
            If systemType = GetType(DateAndTime) Then
                DataType = "TIMESTAMP"
                Length = 10
                Scale = 6
            Else
                If {GetType(Byte), GetType(Short)}.Contains(systemType) Then
                    DataType = "SMALLINT"
                    Length = 2
                Else
                    If systemType = GetType(Integer) Then
                        DataType = "INTEGER"
                        Length = 4
                    Else
                        If systemType = GetType(Long) Then
                            DataType = "BIGINT"
                            Length = 8
                        Else
                            If {GetType(Decimal), GetType(Double)}.Contains(systemType) Then
                                '1234567.11 IS DECIMAL(9, 2)
                                Dim nonNullValues As New List(Of Object)(values?.Where(Function(v) Not (IsDBNull(v) Or v Is Nothing)))
                                If nonNullValues.Any Then
                                    Dim maxLength As Integer
                                    Dim maxScale As Integer
                                    For Each value In nonNullValues
                                        Dim number As Double
                                        If Double.TryParse(value.ToString, number) Then
                                            Dim kvp = DoubleSplit(number)
                                            Dim wholeLength As Integer = kvp.Key.ToString(InvariantCulture).Length
                                            Dim decimalLength As Integer = kvp.Value.ToString(InvariantCulture).Length - 2
                                            If maxLength < wholeLength Then maxLength = wholeLength
                                            If maxScale < decimalLength Then maxScale = decimalLength
                                        End If
                                    Next
                                    DataType = "DECIMAL"
                                    Length = CShort(maxLength)
                                    Scale = CShort(maxScale)

                                Else
                                    'No values in column? Set DECIMAL(10, 2) as default
                                    DataType = "DECIMAL"
                                    Length = 10
                                    Scale = 2

                                End If
                            Else
                                If systemType = GetType(Boolean) Then
                                    'IBM® DB2® 9.x does Not implement a Boolean SQL type.
                                    'Solution: The DB2 database interface converts BOOLEAN type to CHAR(1) columns And stores '1' or '0' values in the column.
                                    DataType = "CHAR"
                                    Length = 1
                                    Scale = 0
                                Else
                                    If systemType = GetType(String) Then
#Region " CHAR + VARCHAR "
                                        Dim lengths As New List(Of Integer)(values?.Where(Function(v) Not (IsDBNull(v) Or v Is Nothing)).Select(Function(v) v.ToString.Length))
                                        If lengths.Any Then
                                            Dim minLength As Integer = {lengths.Min, 2003}.Min
                                            Dim maxLength As Integer = {lengths.Max, 2003}.Min
                                            DataType = If(minLength = maxLength, "CHAR", "VARCHAR") 'Same length for each value=CHAR, Variant lengths for values=VARCHAR
                                            Length = CShort(maxLength)
                                        Else
                                            'No values in column? Set VARCHAR(50) as default
                                            DataType = "VARCHAR"
                                            Length = 50
                                        End If
#End Region
                                    Else
                                        'DUMMY VALUE!
                                        Stop
                                        DataType = "VARCHAR"
                                        Length = 50
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub
    Public Property SystemInfo As SystemObject
    Public Property Name As String
    Public Property Index As Integer
    Public Property DataType As String
    Public ReadOnly Property DataFormat As String
        Get
            If DataType Is Nothing Then
                Return If(DataType, String.Empty)
            Else
                If {"CHAR", "VARCHAR"}.Contains(DataType) Then
                    Return DataType & "(" & Length & ")"
                Else
                    If DataType = "DECIMAL" Then
                        '1234567.11 IS DECIMAL(9, 2) ... 7 Whole + 2 Decimal
                        Return "DECIMAL(" & Length + Scale & ", " & Scale & ")"
                    Else
                        Return DataType
                    End If
                End If
            End If
        End Get
    End Property
    Public Property Length As Short
    Public Property Scale As Short
    Public Property Nulls As Boolean
    Public Overrides Function ToString() As String
        Return Join({DataFormat, Nulls.ToString(InvariantCulture)}, BlackOut)
    End Function
    Public Overrides Function GetHashCode() As Integer
        Return SystemInfo.GetHashCode Xor Name.GetHashCode Xor Index.GetHashCode Xor DataType.GetHashCode Xor DataFormat.GetHashCode Xor Length.GetHashCode Xor Scale.GetHashCode Xor Nulls.GetHashCode
    End Function
    Public Overloads Function Equals(ByVal other As ColumnProperties) As Boolean Implements IEquatable(Of ColumnProperties).Equals
        Return Index = other.Index AndAlso Name = other.Name
    End Function
    Public Shared Operator =(ByVal value1 As ColumnProperties, ByVal value2 As ColumnProperties) As Boolean
        Return value1.Equals(value2)
    End Operator
    Public Shared Operator <>(ByVal value1 As ColumnProperties, ByVal value2 As ColumnProperties) As Boolean
        Return Not value1 = value2
    End Operator
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        If TypeOf obj Is ColumnProperties Then
            Return CType(obj, ColumnProperties) = Me
        Else
            Return False
        End If
    End Function
End Structure
Public Module iData
    Public Event Alerts(sender As Object, e As AlertEventArgs)
    Public Enum InstructionType
        DDL
        SQL
    End Enum
    Public Enum Frequency
        Never
        Always
        Daily
        Weekly
        Monthly
    End Enum
    Public Function RunSQL(ConnectionString As String, Instruction As String) As ResponseEventArgs

        With New SQL(ConnectionString, Instruction)
            .Execute()
            Do While .Response Is Nothing
            Loop
            If .Response.Succeeded Then
                Do While .Table Is Nothing
                Loop
                Return .Response
            Else
                Return .Response
            End If
        End With

    End Function
    Friend Function ColumnSQL(TableName As String) As String
        Return Replace(My.Resources.SQL_ColumnTypes, "'///OWNER_TABLE///'", ValueToField(TableName))
    End Function
    Friend Function ColumnSQL(TableNames As String()) As String

        If TableNames.Any Then
            Return Replace(My.Resources.SQL_ColumnTypes, "'///OWNER_TABLE///'", Join(TableNames.Select(Function(x) ValueToField(x)).ToArray, ","))
        Else
            Return My.Resources.SQL_ColumnTypes
        End If

    End Function
#Region " TO AND FROM DATATABLES "
    Public Function RetrieveData(ByVal Source As String, ByVal SQL As String) As DataTable

        Dim SQL_Table As DataTable
        With New SQL(Source, SQL)
            .Execute()
            Do While .Response Is Nothing
            Loop
            SQL_Table = .Table
        End With
        Return SQL_Table

    End Function
    Public Function DictionaryToProcedures(Connection As Connection, TableName As String, Rows As Dictionary(Of Integer, List(Of Object))) As List(Of DDL)

        Dim Procedures As New List(Of DDL)
        If Connection Is Nothing Or TableName Is Nothing Or Rows Is Nothing Then
        Else
            Dim Blocks = From R In Rows.Keys Select New With {.BlockNbr = QuotientRound(R, 254), .Select = "SELECT " + Join(Rows(R).ToArray, ",") + " FROM SYSIBM.SYSDUMMY1"}
            Dim Inserts = From B In Blocks Group B By BlockNbr = B.BlockNbr Into BlockGroup = Group Select New With {.Index = BlockNbr, .SQL = (From BG In BlockGroup Select BG.Select).ToArray}
            Dim BlockIndex As Integer

            For Each Block In Inserts
                BlockIndex += 1
                Dim Insert As String = "INSERT INTO " + TableName.ToUpperInvariant + vbNewLine + Join(Block.SQL, vbNewLine + "UNION ALL" + vbNewLine)
                Procedures.Add(New DDL(Connection, Insert) With {.Name = Join({BlockIndex, "of", Inserts.Count})})
            Next
        End If
        Return Procedures

    End Function
#Region " DataTable <===> .txt "
    Public Sub DataTableToTextFile(DataTable As DataTable, FilePath As String)

        If DataTable IsNot Nothing Then
            Dim Headers As New List(Of String)(From C In DataTable.Columns Select DirectCast(C, DataColumn).ColumnName)
            Dim Rows As New List(Of DataRow)(From R In DataTable.Rows Select DirectCast(R, DataRow))

            If Rows.Any Then
                Dim TextData As New List(Of String)(From R In Rows Select Join((From C In R.ItemArray Select If(IsDBNull(C), String.Empty, C.ToString)).ToArray, Delimiter))

                Using SR As New StreamWriter(FilePath)
                    SR.WriteLine(Join(Headers.ToArray, Delimiter))
                    For Each Row In TextData.Take(TextData.Count - 1)
                        SR.WriteLine(Row)
                    Next
                End Using
                My.Computer.FileSystem.WriteAllText(FilePath, TextData.Last, True)
            End If
        End If

    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function TextFileToDataTable(FilePath As String) As DataTable

        If FileExists(FilePath) Then
            Dim txt_Data As String = String.Empty
            Using SR As New StreamReader(FilePath)
                txt_Data = SR.ReadToEnd
            End Using
            Return StringToDataTable(txt_Data)
        Else
            Return Nothing
        End If

    End Function
    Public Function TextFileToDataTable(FilePath As String, Delimiter As String) As DataTable

        If FileExists(FilePath) Then
            Dim txt_Data As String = String.Empty
            Using SR As New StreamReader(FilePath)
                txt_Data = SR.ReadToEnd
            End Using
            Return StringToDataTable(txt_Data, Delimiter)
        Else
            Return Nothing
        End If

    End Function
#End Region
#Region " DataTable <===> String, List(Of String) "
    Public Function DataTableToString(StringTable As DataTable) As String
        Return Join(DataTableToList(StringTable).ToArray, vbNewLine)
    End Function
    Public Function DataTableToString(StringTable As DataTable, Distinct As Boolean) As String
        Return Join(DataTableToList(StringTable, Distinct).ToArray, vbNewLine)
    End Function
    Public Function DataTableToList(DataTable As DataTable, Distinct As Boolean) As List(Of String)
        Return If(Distinct, DataTableToList(DataTable).Distinct.ToList, DataTableToList(DataTable))
    End Function
    Public Function DataTableToList(DataTable As DataTable) As List(Of String)

        Dim StringTable As New List(Of String)
        If DataTable IsNot Nothing Then
            Dim Columns As New List(Of String)(From C In DataTable.Columns Select DirectCast(C, DataColumn).ColumnName)
            StringTable.Add(Join(Columns.ToArray, Delimiter))

            Dim Rows As New List(Of DataRow)(From R In DataTable.Rows Select DirectCast(R, DataRow))
            StringTable.AddRange(From R In Rows Select Join((From C In R.ItemArray Select If(IsDBNull(C), String.Empty, C.ToString)).ToArray, Delimiter))
        End If
        Return StringTable

    End Function
    Public Function DataTableToList(DataTable As DataTable, Column As String) As List(Of String)

        Dim StringTable As New List(Of String)
        If DataTable IsNot Nothing Then
            Try
                Dim Values As New List(Of Object)(From R In DataTable.Rows Select DirectCast(R, DataRow)(Column))
                StringTable.AddRange(From v In Values Select If(IsDBNull(v), String.Empty, v.ToString))
            Catch ex As KeyNotFoundException
            End Try
        End If
        Return StringTable

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function StringToDataTable(Body As String, Delimiter As String) As DataTable

        Dim TextTable As New DataTable

        Dim ClipboardRows As New List(Of String)(Split(Body, vbNewLine))
        ClipboardRows = (From CR In ClipboardRows Where CR.Length > 0).ToList
        REM ////////////////////////////////////// A-Z, a-z, 0-9 ARE NOT VALID DELIMITERS
        Dim HeaderRow As String = ClipboardRows(0)
        Dim BodyRows As String() = ClipboardRows.Skip(1).ToArray

        Dim Columns As New List(Of String)(Split(HeaderRow, Delimiter))
        Dim Rows As New List(Of String())(From CR In BodyRows Select Split(CR, Delimiter))

        Dim columnIndex As Integer = 0
        For Each column In Columns
            Dim columnValues As New List(Of String)(From R In Rows Select R(columnIndex))
            Dim columnType As Type = GetDataType(columnValues.Take(1000).ToList, column = "INVDATE")
            columnType = If(columnType Is GetType(DateAndTime), GetType(Date), columnType) 'DateAndTime is a kind of Flag for the DataViewer but is actually Date
            TextTable.Columns.Add(New DataColumn With {
                                  .ColumnName = column,
                                  .DataType = columnType})
            'If column = "INVDATE" Then Stop
            columnIndex += 1
        Next
        For Each Row In Rows
            TextTable.Rows.Add(Row)
        Next

        Return TextTable

    End Function
    Public Function StringToDataTable(Body As String) As DataTable

        Body = If(Body, String.Empty)
        Dim Delimiter As String = String.Empty
        Dim ClipboardRows As String() = (From l In Split(Body, vbNewLine) Where l.Any).ToArray

        If ClipboardRows.Any Then
            REM ////////////////////////////////////// A-Z, a-z, 0-9 ARE NOT VALID DELIMITERS
            Dim HeaderRow As String = ClipboardRows(0)
            Dim BodyRows As String() = ClipboardRows.Skip(1).ToArray
            Dim ValidDelimiters As String = Regex.Replace(HeaderRow, "[A-Z0-9]", String.Empty, RegexOptions.IgnoreCase)
            Dim PotentialDelimiters = (From C In ValidDelimiters.ToCharArray Order By C Group C By ASCC = Asc(C) Into CharacterGroup = Group Select New With {.ColumnName = "C_" & ASCC, ._Count = CharacterGroup.Count})

            If PotentialDelimiters.Count = 1 Then
                Delimiter = Chr(Convert.ToInt32(Split(PotentialDelimiters.First.ColumnName, "_").Last, InvariantCulture))

            Else
                Using CharacterTable As New DataTable
                    For Each Character In PotentialDelimiters
                        CharacterTable.Columns.Add(New DataColumn With {.ColumnName = Character.ColumnName, .DataType = GetType(Integer), .DefaultValue = 0})
                    Next
                    For Each BodyRow In BodyRows
                        Dim RowCharacters = (From C In BodyRow.ToCharArray Order By C Group C By ASCC = Asc(C) Into CharacterGroup = Group Select New With {.ColumnName = "C_" & ASCC, ._Count = CharacterGroup.Count})
                        Dim CurrentRow As DataRow = CharacterTable.Rows.Add()
                        For Each Character In RowCharacters
                            Dim ColumnName As String = Character.ColumnName
                            If CharacterTable.Columns.Contains(ColumnName) Then
                                CurrentRow(ColumnName) = Character._Count
                            End If
                        Next
                    Next
                    Dim Delimiters = (From PD In PotentialDelimiters Where (From CT In CharacterTable.AsEnumerable Where DirectCast(CT(PD.ColumnName), Integer) = PD._Count).Count = CharacterTable.Rows.Count)
                    Delimiters = Delimiters.OrderByDescending(Function(x) x._Count)
                    If Not Delimiters.Any Then
                        Return Nothing
                    End If
                    Delimiter = Chr(Convert.ToInt32(Split(Delimiters.First.ColumnName, "_").Last, InvariantCulture))
                End Using
            End If
            Return StringToDataTable(Body, Delimiter)
        Else
            Return Nothing

        End If

    End Function
#End Region
#Region " DataTable <===> HTML "
    Public Function HTMLToDataSet(ByVal HTML As String) As DataSet

        Dim ds As New DataSet
        Dim dt As DataTable
        Dim dr As DataRow
        'Dim dc As DataColumn
        Dim TableExpression As String = "<table[^>]*>(.*?)</table>"
        Dim HeaderExpression As String = "<th[^>]*>(.*?)</th>"
        Dim RowExpression As String = "<tr[^>]*>(.*?)</tr>"
        Dim ColumnExpression As String = "<td[^>]*>(.*?)</td>"
        Dim HeadersExist As Boolean
        Dim iCurrentColumn As Integer
        Dim iCurrentRow As Integer

        ' Get a match for all the tables in the HTML  
        Dim Tables As MatchCollection = Regex.Matches(HTML, TableExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

        ' Loop through each table element  
        For Each Table As Match In Tables

            ' Reset the current row counter and the header flag  
            iCurrentRow = 0
            HeadersExist = False

            ' Add a new table to the DataSet  
            dt = New DataTable

            ' Create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)  
            If Table.Value.Contains("<th") Then
                ' Set the HeadersExist flag  
                HeadersExist = True

                ' Get a match for all the rows in the table  
                Dim Headers As MatchCollection = Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                ' Loop through each header element  
                For Each Header As Match In Headers
                    dt.Columns.Add(Header.Groups(1).ToString)
                Next
            Else
                For iColumns As Integer = 1 To Regex.Matches(Regex.Matches(Regex.Matches(Table.Value, TableExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Item(0).ToString, RowExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Item(0).ToString, ColumnExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase).Count
                    dt.Columns.Add("Column " & iColumns)
                Next
            End If

            ' Get a match for all the rows in the table  
            Dim Rows As MatchCollection = Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

            ' Loop through each row element  
            For Each Row As Match In Rows

                ' Only loop through the row if it isn't a header row  
                If Not (iCurrentRow = 0 And HeadersExist = True) Then

                    ' Create a new row and reset the current column counter  
                    dr = dt.NewRow
                    iCurrentColumn = 0

                    ' Get a match for all the columns in the row  
                    Dim Columns As MatchCollection = Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline Or RegexOptions.Singleline Or RegexOptions.IgnoreCase)

                    ' Loop through each column element  
                    For Each Column As Match In Columns

                        ' Add the value to the DataRow  
                        dr(iCurrentColumn) = Column.Groups(1).ToString

                        ' Increase the current column  
                        iCurrentColumn += 1
                    Next

                    ' Add the DataRow to the DataTable  
                    dt.Rows.Add(dr)

                End If

                ' Increase the current row counter  
                iCurrentRow += 1
            Next

            ' Add the DataTable to the DataSet  
            ds.Tables.Add(dt)

        Next

        Return (ds)

    End Function
    Public Function DataTableToHtml(DataTable As DataTable, Optional HeaderBackColor As Color = Nothing, Optional HeaderForeColor As Color = Nothing, Optional Script As String = Nothing) As String

        Dim HTML As String = String.Empty
        If DataTable IsNot Nothing Then
            Dim Columns = From C In DataTable.Columns Select DirectCast(C, DataColumn)
            Using TableFont As New Drawing.Font("Calibri", 9)
                Dim ColumnsValuesAsStrings = From C In Columns Select New With {.Name = C.ColumnName, .Values = From R In DataTable.AsEnumerable Select Trim(If(C.DataType Is GetType(Date) And Not IsDBNull(R(C.ColumnName)), If(DirectCast(R(C.ColumnName), Date).TimeOfDay.Ticks = 0, DirectCast(R(C.ColumnName), Date).ToShortDateString, DirectCast(R(C.ColumnName), Date).ToString("M/d/yyyy h:mm tt", InvariantCulture)), R(C.ColumnName).ToString))}
                Dim ColumnWidths = From CV In ColumnsValuesAsStrings Select New With {.Name = CV.Name.ToString(InvariantCulture), .Width = 18 + {TextRenderer.MeasureText(CV.Name, TableFont).Width, (From V In CV.Values Select TextRenderer.MeasureText(V, TableFont).Width).Max}.Max}
                Dim Top As New List(Of String)
                If HeaderBackColor.Name = "0" Then
                    HeaderBackColor = Color.DarkGray
                ElseIf HeaderForeColor.Name = "0" Then
                    HeaderForeColor = Color.White
                End If
                Dim HBS As String = GetHexColor(HeaderBackColor)
                Dim HFS As String = GetHexColor(HeaderForeColor)
#Region " CSS Table Properties "
                Top.Clear()
                Top.Add("<!DOCTYPE html>")
                Top.Add("<html>")
                Top.Add("<head>")
                Top.Add("<style>")
                Top.Add("body {font-family: " & TableFont.FontFamily.Name & "}")
                Top.Add("table {border-collapse:collapse; border: 1px solid #778db3; width: 100%;}")
                Top.Add("th {background-color:" & HBS & "; color:" & HFS & "; text-align:center; font-weight:bold; font-size:0." & TableFont.Size & "em; border: 1px solid #778db3; white-space: nowrap;}")
                Top.Add("td {text-align:left; font-size:0." & (TableFont.Size - 1) & "em; border: 1px #696969; white-space: nowrap;}")
                'Top.Add("tr:nth-child(even) {background-color: #F5F5DC;}")     DOESN'T WORK!
                Top.Add("</style>")
                Top.Add("</head>")
                Top.Add("<body>")
                If Not IsNothing(Script) Then Top.Add("<p>" & Script & "</p>")
                Top.Add("<table>")
                Top.Add("<tr>" & Join((From C In ColumnWidths Select "<th width=" & C.Width & ";>" & C.Name & "</th>").ToArray, "") & "</tr>")
#End Region
                '#F5F5F5, #F5F5DC
                Dim Middle As New List(Of String)
                Dim Rows As New List(Of DataRow)(From R In DataTable.Rows Select DirectCast(R, DataRow))
                For Each Row In Rows
                    Dim ItemArray As New List(Of String)(From CV In ColumnsValuesAsStrings Select CV.Values(Rows.IndexOf(Row)))
                    Middle.Add("<tr style=background-color:" & IIf(Middle.Count Mod 2 = 0, "#F5F5F5", "#FFFFFF").ToString & ";>" + Join((From IA In ItemArray Select "<td>" + IA + "</td>").ToArray, "") + "</tr>")
                Next

                Dim Bottom As New List(Of String) From {
                "</table>",
                "</body>",
                "</html>"
            }
                Dim All As New List(Of String)
                All.AddRange(Top)
                All.AddRange(Middle)
                All.AddRange(Bottom)
                HTML = Join(All.ToArray, vbNewLine)
            End Using
        End If
        Return HTML

    End Function
#End Region
#Region " DataTable <===> EXCEL "
    Private ExcelPath_ As String, SheetName_ As String, Table_ As DataTable
    Private ReadOnly Watch As New Stopwatch
    Public Sub DataTableToExcel(Table As DataTable,
                             ExcelPath As String,
                             Optional FormatSheet As Boolean = False,
                             Optional ShowFile As Boolean = False,
                             Optional DisplayMessages As Boolean = False,
                             Optional IncludeHeaders As Boolean = True,
                             Optional NotifyCreatedFormattedFile As Boolean = False)
        If Table IsNot Nothing Then
            Using ds As New DataSet
                ds.Tables.Add(Table)
                DataSetToExcel(ds, ExcelPath, FormatSheet, ShowFile, DisplayMessages, IncludeHeaders, NotifyCreatedFormattedFile)
            End Using
        End If

    End Sub
    Public Sub DataSetToExcel(TableCollection As DataSet,
                             ExcelPath As String,
                             Optional FormatSheet As Boolean = False,
                             Optional ShowFile As Boolean = False,
                             Optional DisplayMessages As Boolean = False,
                             Optional IncludeHeaders As Boolean = True,
                             Optional NotifyCreatedFormattedFile As Boolean = False)

        If TableCollection IsNot Nothing Then
            Dim App As New Excel.Application
            Dim Book As Excel.Workbook = App.Workbooks.Add
            ExcelPath_ = ExcelPath
            With App
                .Visible = ShowFile
                .DisplayAlerts = DisplayMessages
            End With

            For Each Table As DataTable In TableCollection.Tables
                Dim Sheet As Excel.Worksheet = DirectCast(Book.Sheets.Add, Excel.Worksheet)
                Dim col, row As Integer

                ' Copy the DataTable to an object array - multi-dimensional array ( defined column and row count )
                Dim rawData(Table.Rows.Count, Table.Columns.Count - 1) As Object

                If IncludeHeaders Then
                    ' Copy the column names to the first row of the object array
                    For col = 0 To Table.Columns.Count - 1
                        Dim headerName As String = Table.Columns(col).ColumnName.ToUpperInvariant
                        rawData(0, col) = headerName
                    Next
                End If

                ' Copy the values to the object array
                Dim RowOffset As Integer = If(IncludeHeaders, 1, 0)
                For col = 0 To Table.Columns.Count - 1
                    For row = 0 To Table.Rows.Count - 1
                        rawData(row + RowOffset, col) = Table.Rows(row).ItemArray(col)
                    Next
                Next

                With Sheet
                    .Name = Table.TableName
                    SheetName_ = .Name
                    Dim TableRange As String = String.Format(InvariantCulture, "A1:{0}{1}", ExcelColName(Table.Columns.Count), Table.Rows.Count + 1)
                    .Range(TableRange, Type.Missing).Value2 = rawData
                End With
                ReleaseObject(Sheet)
                Table_ = Table
            Next
            DirectCast(Book.Sheets("Sheet1"), Excel.Worksheet).Delete()

            Try
                Book.Close(True, ExcelPath)
            Catch ex As ExternalException
                Using MESSAGE As New Prompt
                    MESSAGE.Show("Error!", ex.Message, Prompt.IconOption.Critical)
                End Using
            End Try

            ReleaseObject(Book)

            'Release the Application object
            App.Quit()
            ReleaseObject(App)

            'Collect the unreferenced objects
            GC.Collect()
            GC.WaitForPendingFinalizers()

            Dim Windows = Process.GetProcesses
            For Each ExcelProcess In From w In Windows Where w.ProcessName.ToUpperInvariant.Contains("EXCEL") And w.MainWindowTitle.Length = 0 Select w
                Try
                    ExcelProcess.Kill()
                Catch ex As Win32Exception
                End Try
            Next
            If FormatSheet Then
                With New BackgroundWorker
                    AddHandler .DoWork, AddressOf ExcelWorker_Start
                    AddHandler .RunWorkerCompleted, AddressOf ExcelWorker_End
                    .WorkerReportsProgress = NotifyCreatedFormattedFile
                    .RunWorkerAsync()
                End With
            End If
        End If

    End Sub
    Private Sub ExcelWorker_Start(sender As Object, e As DoWorkEventArgs)

        With DirectCast(sender, BackgroundWorker)
            RemoveHandler .DoWork, AddressOf ExcelWorker_Start
        End With
        Watch.Start()
        RaiseEvent Alerts(sender, New AlertEventArgs(Join({"Formatting Excel Workbook", ExcelPath_, "at", Now.ToLongTimeString})))
        FormatSheet(ExcelPath_, SheetName_, Table_)

    End Sub
    Private Sub ExcelWorker_End(sender As Object, e As RunWorkerCompletedEventArgs)

        Watch.Stop()
        With DirectCast(sender, BackgroundWorker)
            RemoveHandler .RunWorkerCompleted, AddressOf ExcelWorker_End
            If .WorkerReportsProgress Then
                RaiseEvent Alerts(sender, New AlertEventArgs(Join({"Formatted Excel Workbook", ExcelPath_, "in", Math.Round(Watch.Elapsed.TotalSeconds, 1), "seconds"})))
            End If
        End With
        If Table_ IsNot Nothing Then Table_.Dispose()

    End Sub
    Public Sub FormatSheet(ExcelPath As String, SheetName As String, Table As DataTable)

        If Table IsNot Nothing Then
            Dim App As New Excel.Application
            Dim Book As Excel.Workbook = App.Workbooks.Open(ExcelPath)
            Dim Sheet As Excel.Worksheet = DirectCast(Book.Sheets(SheetName), Excel.Worksheet)
            Dim TableRange As String = String.Format(InvariantCulture, "A1:{0}{1}", ExcelColName(Table.Columns.Count), Table.Rows.Count + 1)
            With App
                .DisplayAlerts = False
                .ActiveWindow.SplitRow = 1
                .ActiveWindow.FreezePanes = True
            End With
            With Sheet
                .Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                With .Range(TableRange, Type.Missing)
                    .AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, True)
                    With .FormatConditions
                        .Delete()
                        .Add(Excel.XlFormatConditionType.xlExpression, Formula1:="=MOD(ROW(A2),2)=1")
                        With DirectCast(.Item(1), Excel.FormatCondition)
                            .SetFirstPriority()
                            With .Interior
                                .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
                                .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
                                .TintAndShade = -0.0499893185216834
                            End With
                            .StopIfTrue = False
                        End With
                    End With
                    .WrapText = False
                    With .Font
                        .Bold = True
                        .Name = "Trebuchet MS"      'Sakkal Majalla
                        .Size = 8
                    End With
                    For Each Column As DataColumn In Table.Columns
                        Dim ColumnRange As Excel.Range = .Range(.Cells(2, Column.Ordinal + 1), .Cells(2, Column.Ordinal + 1)).EntireColumn
                        Dim ColumnType As Type = GetDataType(Column)
                        Dim CR As String = ColumnRange.Address
                        Select Case ColumnType
                            Case GetType(Byte), GetType(Short), GetType(Integer), GetType(Long)
                                ColumnRange.NumberFormat = "0"
                                ColumnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                            Case GetType(Decimal), GetType(Double)
                                ColumnRange.NumberFormat = "#,##0.00"
                                ColumnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

                            Case GetType(Date)
                                ColumnRange.NumberFormat = "m/d/yyyy"
                                ColumnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                            Case GetType(String)
                                Dim Objects As New List(Of Object)(From r In Table.AsEnumerable Select r(Column))
                                Dim Strings As New List(Of String)(From o In Objects Where Not IsDBNull(o) Select Trim(Convert.ToString(o, InvariantCulture)))
                                If Strings.Any Then
                                    If Strings.Min(Function(s) s.Length) = Strings.Max(Function(s) s.Length) Then
                                        'Strings with uniform width look better centered 
                                        ColumnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                    Else
                                        'Strings with mixed width look better aligned left
                                        ColumnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                                    End If
                                End If

                        End Select
                    Next
                    Dim TopRowRange As String = String.Format(InvariantCulture, "A1:{0}{1}", ExcelColName(Table.Columns.Count), 1)
                    With .Range(TopRowRange, Type.Missing)
                        .Interior.Color = Color.Gainsboro
                        With .Font
                            .Bold = True
                            .Name = "Trebuchet MS"
                            .Size = 9
                        End With
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    End With
                    .EntireColumn.AutoFit()
                End With
            End With
#Region " CLEANUP "
            Try
                Book.Close(True, ExcelPath)

            Catch ex As ExternalException
                Using MESSAGE As New Prompt
                    MESSAGE.Show("Error!", ex.Message, Prompt.IconOption.Critical)
                End Using
            End Try

            ReleaseObject(Book)

            'Release the Application object
            App.Quit()
            ReleaseObject(App)

            'Collect the unreferenced objects
            GC.Collect()
            GC.WaitForPendingFinalizers()

            Dim Windows = Process.GetProcesses
            For Each ExcelProcess In From w In Windows Where w.ProcessName.ToUpperInvariant.Contains("EXCEL") And w.MainWindowTitle.Length = 0 Select w
                Try
                    ExcelProcess.Kill()
                Catch ex As Win32Exception
                End Try
            Next
#End Region
        End If

    End Sub
    Private Function ExcelColName(ByVal Col As Integer) As String

        If Col < 0 And Col > 256 Then
            Return Nothing
        Else
            Dim i As Integer
            Dim r As Integer
            Dim S As String
            If Col <= 26 Then
                S = Chr(Col + 64)
            Else
                r = Col Mod 26
                i = CInt(Math.Floor(Col / 26))
                If r = 0 Then
                    r = 26
                    i -= 1
                End If
                S = Chr(i + 64) & Chr(r + 64)
            End If
            Return S
        End If

    End Function
    Public Sub ReleaseObject(ByVal Item As Object)
        Try
            Marshal.ReleaseComObject(Item)
            Item = Nothing
        Catch ex As NullReferenceException
            Item = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
#End Region
#Region " DATABASE COLUMN TYPES "
    Public Function RetrieveColumnTypes(ConnectionString As String, OwnerAndTableName As String) As DataTable
        Return RetrieveColumnTypes(ConnectionString, {OwnerAndTableName}.ToList)
    End Function
    Public Function RetrieveColumnTypes(ConnectionString As String, OwnersAndTableNames As List(Of String)) As DataTable

        If OwnersAndTableNames Is Nothing Then
            Return Nothing
        Else
            Dim TableNames As New List(Of String)(OwnersAndTableNames.Select(Function(x) ValueToField(x)))
            Dim SQL As String = ColumnSQL(TableNames.ToArray)
            Dim ColumnsTable As DataTable = RetrieveData(ConnectionString, SQL)
            ColumnsTable.Namespace = "<Retrieved>"
            Return ColumnsTable
        End If

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function DataTableToListOfColumnProperties(columnTable As DataTable) As List(Of ColumnProperties)

        Dim Columns As New List(Of ColumnProperties)
        If columnTable IsNot Nothing Then
            If columnTable.Namespace = "<Retrieved>" Then
#Region " FULL DATABASE DETAIL "
                REM /// CERTAIN THAT THE BELOW COLUMNNAMES ARE IN THE TABLE SINCE...
                REM /// THIS REQUEST COMES FROM CONNECTION TO A DATABASE USING THE MY.SETTINGS.ColumnTypes SQL
                REM /// EACH ROW IN THE TABLE IS FOR COLUMN PROPERTIES...EACH COLUMN IN THE ROW IS A PROPERTY
                Columns = DataTableToListOfColumnsProperties(columnTable)
#End Region
            Else
                REM /// THIS IS CASTING FROM USER QUERY
#Region " SELECT CAST(... ==> .NET "
                '------------ FILLING A DATATABLE FROM A DATABASE SOURCE
                Dim SQL As New List(Of String) From {"SELECT CAST(0 AS SMALLINT) SMALLINT_",
", CAST(0 AS INT) INT_",
", CAST(0 AS BIGINT) BIGINT_",
", CAST(0 AS DECIMAL(15, 2)) DECIMAL_",
", CAST(0 AS FLOAT) DOUBLE_",
", CAST(CURRENT_TIMESTAMP AS TIMESTAMP) TIMESTAMP_",
", CAST(CURRENT_DATE AS DATE) DATE_",
", CAST('X' AS CHAR(10)) CHAR_",
", CAST('X' AS VARCHAR(20)) VARCHAR_",
"FROM DSNA1.SYSIBM.SYSDUMMY1"}
                '------------ ...RESULTS IN THEN BELOW DATACOLUMN.DATATYPES
                '------------ ...GOOD TRANSLATION BUT System.Decimal, DateTime, and String LOSE THEIR ORIGINAL CASTING
                'SMALLINT_	→	System.Int16
                'INT_		→	System.Int32
                'BIGINT_	→	System.Int64
                'DECIMAL_	→	System.Decimal
                'DOUBLE_	→	System.Decimal
                'TIMESTAMP_	→	System.DateTime
                'DATE_		→	System.DateTime
                'CHAR_		→	String
                'VARCHAR_	→	String
#End Region

                For Each Column As DataColumn In columnTable.Columns
                    Dim columnType As Type = If(columnTable.Namespace = "<DB2>", Column.DataType, GetDataType(DataColumnToList(Column)))
                    Dim values As New List(Of Object)(DataColumnToList(Column))
                    Columns.Add(New ColumnProperties(Column.DataType, values) With {
                            .Name = DatabaseColumnName(Column.ColumnName),
                            .Index = Column.Ordinal})
                Next
            End If
        End If
        Return Columns

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function RetrieveColumnPropertiesTypes(ObjectItem As SystemObject) As List(Of ColumnProperties)

        If ObjectItem Is Nothing Then
            Return Nothing
        Else
            Return RetrieveColumnPropertiesTypes(ObjectItem.Connection.ToString, {ObjectItem.FullName}.ToList)
        End If

    End Function
    Public Function RetrieveColumnPropertiesTypes(ConnectionString As String, OwnerAndTableName As String) As List(Of ColumnProperties)
        Return RetrieveColumnPropertiesTypes(ConnectionString, {OwnerAndTableName}.ToList)
    End Function
    Public Function RetrieveColumnPropertiesTypes(ConnectionString As String, OwnerTableNames As List(Of String)) As List(Of ColumnProperties)

        Dim ColumnTable As DataTable = RetrieveColumnTypes(ConnectionString, OwnerTableNames)
        If IsNothing(ColumnTable) Then
            Return Nothing
        Else
            Return DataTableToListOfColumnsProperties(ColumnTable)
        End If

    End Function
    Public Function DataTableToListOfColumnsProperties(Table As DataTable) As List(Of ColumnProperties)

        REM /// DATATABLE IS RESULT OF COLUMN_TYPES QUERY
        Dim Columns As New List(Of ColumnProperties)
        If Table IsNot Nothing Then
            For Each DataRow As DataRow In Table.Rows
                Dim DB2_Column As New ColumnProperties
                With DB2_Column
                    .SystemInfo = New SystemObject With {
                                .DBName = DataRow.Item("DBNAME").ToString,
                                .Name = DataRow.Item("TABLE_NAME").ToString,
                                .Owner = DataRow.Item("CREATOR").ToString,
                                .TSName = DataRow.Item("TSNAME").ToString,
                                .Type = DirectCast([Enum].Parse(GetType(SystemObject.ObjectType), StrConv(DataRow.Item("OBJECT_TYPE").ToString, VbStrConv.ProperCase)), SystemObject.ObjectType),
                                .DSN = DataRow.Item("DSN").ToString
                    }
                    .Name = DataRow.Item("COLUMN_NAME").ToString
                    .Index = CInt(DataRow.Item("COL#"))
                    .DataType = DataRow.Item("COLTYPE").ToString
                    .Length = CShort(DataRow.Item("LENGTH"))
                    .Scale = CShort(DataRow.Item("SCALE"))
                    .Nulls = DataRow.Item("NULLS").ToString.Contains("Y")
                End With
                Columns.Add(DB2_Column)
            Next
        End If
        Return Columns

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function ColumnPropertiesToTableViewProcedure(Columns As List(Of ColumnProperties)) As String

        'DROP TABLE UNTIED
        '; CREATE TABLE UNTIED (
        'CUST#       INTEGER
        ', INV#		 CHAR(7)
        ', TIME#	 DECIMAL(15, 0)
        ') IN W75DFLTD.W75CCITS
        If Columns Is Nothing Then
            Return Nothing
        Else
            If Columns.Any Then
                Dim Elements = Columns.First.SystemInfo
                '.SystemInfo = {
                '0                    DataRow.Item("DBNAME").ToString,
                '1                    DataRow.Item("TABLE_NAME").ToString,
                '2                    DataRow.Item("CREATOR").ToString,
                '3                    DataRow.Item("TSNAME").ToString,
                '4                    StrConv(DataRow.Item("OBJECT_TYPE").ToString, VbStrConv.ProperCase),
                '5                    DataRow.Item("DSN").ToString
                '        }
                Dim SystemObject As SystemObject = Columns.First.SystemInfo
                Dim CreateObject As New List(Of String) From {
                    Join({"DROP", SystemObject.Type.ToString.ToUpperInvariant, SystemObject.FullName}),
                    Join({"; CREATE", SystemObject.Type.ToString.ToUpperInvariant, SystemObject.FullName, "("})
                }

                For Each ColumnProperties In Columns
                    Dim Line As String = Join({ColumnProperties.Name, ColumnProperties.DataFormat}, StrDup(4, vbTab))
                    If ColumnProperties.Index = 1 Then
                        CreateObject.Add(Line)
                    Else
                        CreateObject.Add(", " & Line)
                    End If
                Next
                CreateObject.Add(") IN " & SystemObject.TSName)
                Return Join(CreateObject.ToArray, vbNewLine)
            Else
                Return String.Empty
            End If
        End If

    End Function
#End Region
    Public Function FileExists(FilePath As String) As Boolean

        Try
            Using SR As New StreamReader(FilePath)
            End Using
            Return True

        Catch ex As FileNotFoundException
            Return False
        End Try

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function ValueToField(Value As Object) As String

        If Value Is Nothing Then
            Return Nothing
        Else
            Return ValueToField(Value, Value.GetType)
        End If

    End Function
    Public Function ValuesToFields(Values As Object()) As String

        If Values Is Nothing Then
            Return Nothing
        Else
            Dim Items As New List(Of String)
            For Each Value In Values
                Items.Add(ValueToField(Value, GetDataType(Value)))
            Next
            Return "(" & Join(Items.ToArray, ",") & ")"
        End If

    End Function
    Public Function ValueToField(Value As Object, ValueType As Type) As String

        If Value Is Nothing Then
            Return Nothing
        Else
            Select Case ValueType
                Case GetType(String)
                    Return Join({"'", Value.ToString, "'"}, String.Empty)

                Case GetType(Date)
                    Dim DateValue As Date = DirectCast(Value, Date)
                    If DateValue.TimeOfDay = New TimeSpan(0) Then
                        Return DateToDB2Date(DateValue)
                    Else
                        Return DateToDB2Timestamp(DateValue)
                    End If

                Case Else
                    Return Value.ToString

            End Select
        End If

    End Function
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Function DatabaseColumnName(DataColumnName As String) As String

        If DataColumnName Is Nothing Then
            Return Nothing
        Else
            DataColumnName = DataColumnName.ToUpperInvariant
            REM A DB2 COLUMN NAME MUST START WITH EITHER A [A-Z] Or @ Or # Or $
            DataColumnName = Regex.Replace(DataColumnName, "^[^A-Z@#$]", "#", RegexOptions.IgnoreCase)
            REM A DB2 COLUMN NAME MUST NOT HAVE ILLEGAL CHARACTERS IN POSITION 2...n
            DataColumnName = Regex.Replace(DataColumnName, "(?<=^[^■])[^A-Z0-9_@#$\n\r]{1,}", "#", RegexOptions.IgnoreCase)
            Return Trim(DataColumnName)
        End If

    End Function

    Public Function ConnectionStringToCredentials(Source As String) As KeyValuePair(Of String, String)

        'https://www.ibm.com/support/knowledgecenter/en/SSEPGG_9.7.0/com.ibm.swg.im.dbclient.adonet.ref.doc/doc/DB2ConnectionClassConnectionStringProperty.html
        'User ID | UID
        'Password | PWD
        Dim UID As String = String.Empty
        Dim PWD As String = String.Empty

        Dim Credentials As String() = Split(Source, ";")
        Dim UID_Field = Credentials.Where(Function(u) Regex.Match(u, "User ID|UID", RegexOptions.IgnoreCase).Success)
        If UID_Field.Any Then
            UID = UID_Field.First
        End If

        Dim PWD_Field = Credentials.Where(Function(u) Regex.Match(u, "Password|PWD", RegexOptions.IgnoreCase).Success)
        If PWD_Field.Any Then
            PWD = PWD_Field.First
        End If
        Return New KeyValuePair(Of String, String)(UID, PWD)

    End Function
    Friend Function CursorDirection(Point1 As Point, Point2 As Point) As Cursor

        If Point1.X = Point2.X And Point1.Y = Point2.Y Then
            Return Cursors.Default

        ElseIf Point1.X = Point2.X And Point1.Y < Point2.Y Then
            Return Cursors.PanNorth

        ElseIf Point1.X = Point2.X And Point1.Y > Point2.Y Then
            Return Cursors.PanSouth

        ElseIf Point1.X < Point2.X And Point1.Y = Point2.Y Then
            Return Cursors.PanWest

        ElseIf Point1.X > Point2.X And Point1.Y = Point2.Y Then
            Return Cursors.PanEast

        ElseIf Point1.X < Point2.X And Point1.Y < Point2.Y Then
            Return Cursors.PanNW

        ElseIf Point1.X < Point2.X And Point1.Y > Point2.Y Then
            Return Cursors.PanSW

        ElseIf Point1.X > Point2.X And Point1.Y < Point2.Y Then
            Return Cursors.PanNE

        ElseIf Point1.X > Point2.X And Point1.Y > Point2.Y Then
            Return Cursors.PanSE

        Else
            Return Cursors.Default

        End If

    End Function
#End Region
    Friend Function SaferQuery(connection As String, Instruction As String) As DataTable

        Using someConnection As New SqlConnection(connection)
            Using someCommand As New SqlCommand()
                someCommand.Connection = someConnection
                someCommand.Parameters.Add(
                "@Instruction",
                SqlDbType.NChar).Value = Instruction
                someCommand.CommandText = "@Instruction"

                someConnection.Open()

                Using da As New SqlDataAdapter(someCommand)
                    Using dt As New DataTable
                        da.Fill(dt)
                        Return dt
                    End Using
                End Using

                someConnection.Close()

            End Using
        End Using

    End Function
End Module