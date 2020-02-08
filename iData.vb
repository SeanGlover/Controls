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
            Handle.Dispose()
            ' Free any other managed objects here.
            ElementsWorker.Dispose()
        End If
        disposed = True
    End Sub
#End Region
    Public Event ConnectionChanged(sender As Object, e As ConnectionChangedEventArgs)
    Public Event TypeChanged(sender As Object, e As ScriptTypeChangedEventArgs)
    Public Event Completed(sender As Object, e As EventArgs)

    Private WithEvents ElementsWorker As New BackgroundWorker With {.WorkerReportsProgress = False}
    Private WithEvents ChangedTimer As New Timer With {.Interval = 400}         'When Connection or Instruction ( Text ) changes

    Private Const BlackOut As String = "■"
    Private Const NonCharacter As String = "©"
    Private Const SelectPattern As String = "SELECT[^■]{1,}?(?=FROM)"
    Private Const CommentPattern As String = "--[^\r\n]{1,}(?=\r|\n|$)"
    Private Const ObjectPattern As String = "([A-Z0-9!%{}^~_@#$]{1,}([.][A-Z0-9!%{}^~_@#$]{1,}){0,2})"     'DataSource.Owner.Name
    'Private Const OrderByPattern As String = "ORDER\s+BY\s+" & ObjectPattern & "(,\s+" & ObjectPattern & "){0,}"           * U N U S E D - BUT USEFUL
    Private Const FromJoinCommaPattern As String = "(?<=FROM |JOIN )[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}|(?<=,)[\s]{0,}[A-Z0-9!%{}^~_@#$♥]{1,}([.][A-Z0-9!%{}^~_@#$♥]{1,}){0,2}([\s]{1,}[A-Z0-9!%{}^~_@#$]{1,}){0,1}"
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public Sub New()
        IsBusy = True
        Initializing = True
        Connections = New ConnectionCollection()
        Objects = New SystemObjectCollection
    End Sub
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
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
    Private _Connection As Connection
    Public Property Connection As Connection
        Get
            Return _Connection
        End Get
        Set(value As Connection)
            ConnectionChange = value <> _Connection
            If ConnectionChange Then
                _Connection = value
                _IsBusy = False
                ChangedTimer_Tick()
            End If
        End Set
    End Property
    Public ReadOnly Property DerivedConnection As Connection
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
                            If IsNetezza() Then
                                Lines.Add("LIMIT " & Limit)
                            Else
                                Lines.Add("FETCH FIRST " & Limit & " ROWS ONLY")
                            End If

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
    Public ReadOnly Property InstructionType As ExecutionType = ExecutionType.Null
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
    Public ReadOnly Property TablesNeedObject As List(Of String)
        Get
            If TablesElement.Any Then
                'TablesElement is the Body.Text- what the user writes vs TablesObject which is saved SystemObjects (Table)
                'If an item in the Body is not in the saved SystemObjects, then add so an SystemObject can be retrieved from the Database
                Dim BodyTables As New List(Of String)(TablesElement.Select(Function(te) te.FullName))
                Dim HaveObjects As New List(Of String)
                For Each Item In TablesObject
                    If BodyTables.Contains(Item.FullName) Then
                        HaveObjects.Add(Item.FullName)
                    End If
                Next
                Return BodyTables.Except(HaveObjects).ToList
            Else
                Return New List(Of String)
            End If
        End Get
    End Property
    Private ReadOnly Property TablesObject As New List(Of SystemObject)
    Public ReadOnly Property Withs As New List(Of InstructionElement)
    Public ReadOnly Property DataSources As New Dictionary(Of String, List(Of SystemObject))
    Public ReadOnly Property DataSource As SystemObject
    Public ReadOnly Property ElementObjects As New Dictionary(Of InstructionElement, List(Of SystemObject))
    '▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬
    Public ReadOnly Property IsNetezza As Boolean
        Get
            If IsNothing(Connection) Then
                Return False
            Else
                Return Connection.IsNetezza
            End If
        End Get
    End Property
    Public ReadOnly Property IsDB2 As Boolean
        Get
            If IsNothing(Connection) Then
                Return False
            Else
                Return Connection.IsDB2
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
            '#1----------------------------COMMENTED TEXT...Commented text MUST be ignored
            _UncommentedText = StripComments(Text)

            '#2----------------------------Determineif DDL Or SQL
            Dim LastType As ExecutionType = InstructionType
            _InstructionType = GetInstructionType()
            If LastType <> InstructionType Then RaiseEvent TypeChanged(Me, New ScriptTypeChangedEventArgs(LastType, InstructionType))

            '----------------------------CLASSIFY TEXT
            AssignLabels()

            '----------------------------CROSS-REFERENCE TEXT AS AN OBJECT THAT RESIDES IN A DATABASE
            AddSystemObjects()
            '----------------------------DETEREMINE DATASOURCE
            GetDataSources()
            '----------------------------SET THE CONNECTION
            Dim LastConnection = Connection
            _DerivedConnection = GetConnection()
            If LastConnection = DerivedConnection Then
            Else
                _Connection = DerivedConnection
                RaiseEvent ConnectionChanged(Me, New ConnectionChangedEventArgs(LastConnection, _DerivedConnection))
            End If
        End If
        '----------------------------SET THE DATABASE TEXT
        _SystemText = If(GetSystemText(), Text)


    End Sub
    Private Sub EndObjectsWork(sender As Object, e As RunWorkerCompletedEventArgs) Handles ElementsWorker.RunWorkerCompleted
        _Initializing = False
        _IsBusy = False
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
            Dim Match_CreateAlterDrop As Match = Regex.Match(UncommentedText, "(CREATE|ALTER|DROP)(\s{1,}OR REPLACE){0,1}\s{1,}(TABLE|VIEW|Function|TRIGGER)[\s]{1,}" + ObjectPattern, RegexOptions.IgnoreCase)
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
                    If Element.Source = InstructionElement.LabelName.Limit And Connection.IsDB2 Then
                        Dim Limit As InstructionElement = Element
                        With Limit
                            Dim RowCount As Integer = Integer.Parse(Regex.Match(.Block.Value, "[0-9]{1,}", RegexOptions.None).Value, Globalization.CultureInfo.InvariantCulture)
                            Dim LimitText As String = DatabaseText.Substring(.Block.Start, .Block.Length)
                            If LimitText.ToUpper(Globalization.CultureInfo.InvariantCulture).StartsWith("LIMIT", StringComparison.InvariantCulture) Then
                                DatabaseText = DatabaseText.Remove(.Block.Start, .Block.Length)
                                DatabaseText = DatabaseText.Insert(.Block.Start, Join({"FETCH FIRST", RowCount.ToString(Globalization.CultureInfo.InvariantCulture), "ROWS ONLY"}))
                                If Not Regex.Match(DatabaseText, "FETCH\s+FIRST\s+[0-9]{1,9}\s+ROWS\s+ONLY", RegexOptions.IgnoreCase).Success Then
                                    Stop
                                End If
                            End If
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
        Return Item(New Connection(ConnectionString))
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
    Public Shadows ReadOnly Property ToString As String
        Get
            If IsFile Then
                Return Properties.Keys.First
            Else
                Return Join(Properties.Select(Function(x) Join({x.Key, x.Value}, "=")).ToArray, ";")
            End If
        End Get
    End Property
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
    Public Shadows ReadOnly Property ToString As String
        Get
            Return Join({DSN, Type.ToString, DBName, TSName, Owner, Name}, Delimiter)
        End Get
    End Property
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

    End Sub
    Public Sub New(Connection As Connection, Instruction As String)

        ConnectionString = If(Connection Is Nothing, String.Empty, Connection.ToString)
        Me.Instruction = If(Instruction, String.Empty)

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
#Region " /// EXCELFILE "
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
    Private rf As ResponseFailure
    Public Sub New(ConnectionString As String, Instruction As String, Optional PromptForInput As Boolean = False, Optional GetRowCount As Boolean = False)
        Me.ConnectionString = ConnectionString
        Me.Instruction = Instruction
        RequiresInput = PromptForInput
        Me.GetRowCount = GetRowCount
    End Sub
    Public Sub New(Connection As Connection, Instruction As String, Optional PromptForInput As Boolean = False, Optional GetRowCount As Boolean = False)

        If Connection IsNot Nothing Then
            ConnectionString = Connection.ToString
            Me.Instruction = Instruction
            RequiresInput = PromptForInput
            Me.GetRowCount = GetRowCount
        End If

    End Sub
    Public Sub Execute()

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
            Dim PromptProcedures As New List(Of Procedure)(Procedures)
            If RequiresInput Then
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
                    End With
                Next
            End If
            Dim OkdProcedures As New List(Of String)(From p In PromptProcedures Where p.Execute Select p.Instruction)
            If OkdProcedures.Any Then
                _Started = Now
                Using New CursorBusy
                    Using _Connection As New OdbcConnection(ConnectionString)
                        Dim DDL_Instruction As String = Join(OkdProcedures.ToArray, ";")
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
            Else
                Using message As New Prompt
                    message.Show("Request(s) cancelled", "Action(s) cancelled", Prompt.IconOption.TimedMessage)
                End Using
                _Response = New ResponseEventArgs(InstructionType.DDL, ConnectionString, Instruction, Nothing, Ended - Started)
            End If
        End If
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
                        Dim tableName As String = match_Alter.First
                        Dim columnName As String = match_Alter(1)
                        Dim newDataType As String = match_Alter(2)
                        ObjectAction = Action.Alter
                        ObjectType = Type.Column
                        ObjectName = Join({tableName, columnName, newDataType}, BlackOut)
                        FetchStatement = Replace(Replace(My.Resources.SQL_ColumnTypes, "///OWNER_TABLE///", tableName), "--AND C.NAME='//COLUMN_NAME//'", "AND C.NAME=" & ValueToField(columnName))

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
                If GetFileNameExtension(ConnectionString).Value = Extensions.Text Then
                    DataTableToTextFile(Table, ConnectionString)
                ElseIf GetFileNameExtension(ConnectionString).Value = Extensions.Excel Then
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

            _Columns = DataTableToListOfColumnProperties(Table).ToDictionary(Function(x) x.Name, Function(y) y)
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

                                            Case "DECIMAL", "SMALLINT", "INTEGER", "BIGINT", "DECFLOAT"
                                                REM /// NO FORMATTING NEEDED FOR NUMBERS
                                                If EmptyValue Then
                                                    Values.Add(0)
                                                Else
                                                    If Column.SourceType = GetType(Boolean) Then
                                                        Values.Add(If(Value.ToString.ToUpperInvariant = "FALSE", 0, 1))
                                                    Else
                                                        Values.Add(Value)
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
    Public Property SystemInfo As SystemObject
    Public Property Name As String
    Public Property Index As Integer
    Public Property DataType As String
    Public Property DataFormat As String
    Public Property Length As Short
    Public Property Scale As Short
    Public Property Nulls As Boolean
    Public Overrides Function ToString() As String
        Return Join({SystemInfo.ToString, Join({Name, Index, DataFormat, Nulls.ToString(InvariantCulture)}, Delimiter)}, BlackOut)
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
        Dim ColumnTypes = (From C In Columns Select New With {.Name = C, .DataType = GetDataType((From R In Rows Select R(Columns.IndexOf(C))).Take(1000).ToList)})

        For Each Column In ColumnTypes
            TextTable.Columns.Add(New DataColumn With {.ColumnName = Column.Name, .DataType = Column.DataType})
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

                ' Copy the DataTable to an object array
                Dim rawData(Table.Rows.Count, Table.Columns.Count - 1) As Object

                If IncludeHeaders Then
                    ' Copy the column names to the first row of the object array
                    For col = 0 To Table.Columns.Count - 1
                        rawData(0, col) = Table.Columns(col).ColumnName.ToUpperInvariant
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
        RaiseEvent Alerts(Nothing, New AlertEventArgs(Join({"Formatting Excel Workbook", ExcelPath_, "at", Now.ToLongTimeString})))
        FormatSheet(ExcelPath_, SheetName_, Table_)

    End Sub
    Private Sub ExcelWorker_End(sender As Object, e As RunWorkerCompletedEventArgs)

        With DirectCast(sender, BackgroundWorker)
            RemoveHandler .RunWorkerCompleted, AddressOf ExcelWorker_End
            If .WorkerReportsProgress Then
                Using Message As New Prompt()
                    Message.Show("File saved", ExcelPath_, Prompt.IconOption.TimedMessage)
                End Using
            End If
        End With
        Watch.Stop()
        RaiseEvent Alerts(Nothing, New AlertEventArgs(Join({"Formated Excel Workbook", ExcelPath_, "in", Math.Round(Watch.Elapsed.TotalSeconds, 1), "seconds"})))
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
    Public Function DataTableToListOfColumnProperties(ColumnTable As DataTable) As List(Of ColumnProperties)

        Dim Columns As New List(Of ColumnProperties)
        If ColumnTable IsNot Nothing Then
            Dim QueryToColumnTypes As Boolean = ColumnTable.Namespace = "<Retrieved>"
            Dim Query_Generic As Boolean = ColumnTable.Namespace = "<DB2>"

            If QueryToColumnTypes Then
#Region " FULL DATABASE DETAIL "
                REM /// CERTAIN THAT THE BELOW COLUMNNAMES ARE IN THE TABLE SINCE...
                REM /// THIS REQUEST COMES FROM CONNECTION TO A DATABASE USING THE MY.SETTINGS.ColumnTypes SQL
                REM /// EACH ROW IN THE TABLE IS FOR COLUMN PROPERTIES...EACH COLUMN IN THE ROW IS A PROPERTY
                Columns = DataTableToListOfColumnsProperties(ColumnTable)
#End Region
            ElseIf Query_Generic Then
#Region " PARTIAL DATABASE DETAIL "
                REM /// THIS REQUEST IS TO DETERMINE THE DB2 COLUMNTYPES FROM ANY SQL
                REM /// ...WON'T GET TABLENAME, TABLESPACE, ETC BUT CAN CLASS THE DATATYPES
                REM /// DATATABLE COLUMN.DATATYPES ARE FILLED BY AN ADAPTER AND SHOULD USE THE DATATABLE'S COLUMN.DATATYPE INFO
                For Each _DataColumn As DataColumn In ColumnTable.Columns
                    Columns.Add(KnownSourceToColumnProperties(_DataColumn))
                Next
#End Region
            Else
#Region " NO DATABASE DETAIL "
                REM /// THIS REQUEST IS TO DETERMINE THE DB2 COLUMNTYPES FROM A NON-DB2 SOURCE
                REM /// USAGE IS TAKING .txt Or .xlsx FILES TO CREATE A TABLE IN DB2 SPACE
                For Each _DataColumn As DataColumn In ColumnTable.Columns
                    Columns.Add(UnknownSourceToColumnProperties(_DataColumn))
                Next
#End Region
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
                    .Index = Convert.ToInt32(DataRow.Item("COL#"), InvariantCulture)
                    .DataType = DataRow.Item("COLTYPE").ToString
                    .DataFormat = DataRow.Item("COLUMN_DATA").ToString
                    .Length = Convert.ToInt16(DataRow.Item("LENGTH"), InvariantCulture)
                    .Scale = Convert.ToInt16(DataRow.Item("SCALE"), InvariantCulture)
                    .Nulls = (DataRow.Item("NULLS").ToString.Contains("Y"))
                End With
                Columns.Add(DB2_Column)
            Next
        End If
        Return Columns

    End Function
    Public Function DataColumnToColumnProperties(Column As DataColumn) As ColumnProperties

        If Column Is Nothing Then
            Return Nothing
        Else
            If Column.Table.Namespace = "<DB2>" Then
                Return KnownSourceToColumnProperties(Column)
            Else
                Return UnknownSourceToColumnProperties(Column)
            End If
        End If

    End Function
    Public Function KnownSourceToColumnProperties(TableColumn As DataColumn) As ColumnProperties

        If TableColumn Is Nothing Then
            Return Nothing
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
            Dim _DataTable As DataTable = TableColumn.Table
            Dim Database_Column As New ColumnProperties
            With Database_Column
                .SystemInfo = Nothing
                .Name = DatabaseColumnName(TableColumn.ColumnName)
                .Index = TableColumn.Ordinal
                .Nulls = True
                .DataFormat = String.Empty
                Dim Values As New List(Of Object)(From R As DataRow In _DataTable.AsEnumerable Where Not (IsDBNull(R.Item(TableColumn.ColumnName)) Or R.Item(TableColumn.ColumnName).ToString.Length = 0) Select R.Item(TableColumn.ColumnName))
                Select Case TableColumn.DataType
#Region " SMALLINT + INTEGER + BIGINT "
                    Case GetType(Short)
                        .DataType = "SMALLINT"
                        .Length = 2
                        .DataFormat = .DataType

                    Case GetType(Integer)
                        .DataType = "INTEGER"
                        .Length = 4
                        .DataFormat = .DataType

                    Case GetType(Long)
                        .DataType = "BIGINT"
                        .Length = 8
                        .DataFormat = .DataType
#End Region
                    Case GetType(Decimal), GetType(Double)
#Region " DECIMAL(n, n) "
                        REM /// CAST(10.32 AS FLOAT) CAME THROUGH AS Net.System.Decimal DESPITE <FLOAT>=<Net.System.Double>
                        REM /// ie) Case GetType(Double) WILL NOT HAPPEN MEANING MAY HAVE BEEN CAST AS FLOAT BUT END UP AS DECIMAL
                        Dim Amounts As New List(Of String)(From A In Values Select CStr(A))
                        If Amounts.Any Then
                            .Length = CShort((From A In Amounts Select A.Length).Max)
                        Else
                            .Length = 0
                        End If
                        .Scale = 0
                        If Amounts.Any Then
                            .Scale = CShort((From A In Amounts Where UBound(Split(A, ".")) > 0 Select Split(A, ".")(1).Length).Max)
                        End If
                        .DataType = "DECIMAL"
                        .DataFormat = "DECIMAL(" & .Length & ", " & .Scale & ")"
#End Region
                    Case GetType(Date)
#Region " DATE + TIMESTAMP "
                        Dim Dates As New List(Of Date)(From D In Values Select Date.Parse(D.ToString, InvariantCulture))
                        REM /// TEST IF ALWAYS MIDNIGHT
                        REM /// DATABASE DATE HAS NO TIME VALUE, TIMESTAMP HAS DATE+TIME
                        If Values.Count = (From D In Dates Where D.TimeOfDay.Ticks = 0 Select D).Count Then
                            .DataType = "DATE"
                            .Length = 4
                        Else
                            .DataType = "TIMESTAMP"
                            .Length = 10
                            .Scale = 6
                        End If
                        .DataFormat = .DataType
#End Region
                    Case GetType(String)
#Region " CHAR + VARCHAR "
                        Dim ColumnEmpty As Boolean = Not Values.Any
                        If ColumnEmpty Then
                            .Length = Convert.ToInt16(.Name.Length)
                        Else
                            .Length = Convert.ToInt16({(From O In Values Select O.ToString.Length).Max, 2003}.Min)
                        End If
                        .Scale = 0
                        Dim Strings = From v In Values Select Convert.ToString(v, InvariantCulture)
                        If ColumnEmpty Then
                            REM /// PREFER VARCHAR
                            .DataType = "VARCHAR"
                            .DataFormat = "VARCHAR(" & .Length & ")"
                        Else
                            If (From s In Strings Where s.Length <> Database_Column.Length).Any Then
                                REM /// LENGTHS VARY...MUST BE VARCHAR
                                .DataType = "VARCHAR"
                                .DataFormat = "VARCHAR(" & .Length & ")"
                            Else
                                REM /// LENGTH OF EACH FIELD IS THE SAME...MUST BE CHAR
                                .DataType = "CHAR"
                                .DataFormat = "CHAR(" & .Length & ")"
                            End If
                        End If
#End Region
                End Select
            End With
            Return Database_Column
        End If

    End Function
    Public Function UnknownSourceToColumnProperties(Column As DataColumn) As ColumnProperties

        If Column Is Nothing Then
            Return Nothing
        Else
            Dim _DataTable As DataTable = Column.Table
            Dim Database_Column As New ColumnProperties
            With Database_Column
                .Index = Column.Ordinal
                .Name = DatabaseColumnName(Column.ColumnName)
                .Nulls = True
                Dim Values As New List(Of Object)(From R As DataRow In _DataTable.AsEnumerable Where Not (IsDBNull(R.Item(Column.ColumnName)) Or R.Item(Column.ColumnName).ToString.Length = 0) Select R.Item(Column.ColumnName))
                Dim _DataType As Type
                If Values.Any Then
                    _DataType = GetDataType((From V In Values.Take(1000) Select CStr(V)).ToList)
                Else
                    _DataType = Column.DataType
                End If

                Select Case _DataType
#Region " DATE + TIMESTAMP "
                    Case GetType(Date), GetType(Date)
                        Dim Dates As New List(Of Date)(From D In Values Select Date.Parse(D.ToString, InvariantCulture))
                        REM /// TEST FOR ALWAYS MIDNIGHT
                        If Values.Count = (From D In Dates Where D.TimeOfDay.Ticks = 0 Select D).Count Then
                            .DataType = "DATE"
                            .Length = 4
                        Else
                            .DataType = "TIMESTAMP"
                            .Length = 10
                            .Scale = 6
                        End If
                        .DataFormat = .DataType
#End Region
                    Case GetType(Decimal), GetType(Double), GetType(Short), GetType(Integer), GetType(Long), GetType(Byte)
                        REM /// TEST FOR INTEGER, Or IF A DECIMAL, THEN HOW MANY PLACES
                        Dim Amounts As New List(Of String)(From A In Values Select CStr(A))
                        If Values.Count = (From A In Amounts Where UBound(Split(A, ".")) = 0 Select A).Count Then
#Region " SMALLINT + INTEGER + BIGINT "
                            REM /// IS INTEGER, NOW CHECK FOR SIZE
                            Dim INT_16 As Short
                            Dim INT_32 As Integer
                            If (From Small In Amounts Where Short.TryParse(Small, INT_16)).Count = Values.Count Then
                                .DataType = "SMALLINT"
                                .Length = 2
                            Else
                                If (From Medium In Amounts Where Integer.TryParse(Medium, INT_32)).Count = Values.Count Then
                                    .DataType = "INTEGER"
                                    .Length = 4
                                Else
                                    .DataType = "BIGINT"
                                    .Length = 8
                                End If
                            End If
                            .DataFormat = .DataType
#End Region
                        Else
#Region " DECIMAL(n, n) "
                            REM /// HAS DECIMAL PLACES
                            REM 1234567.11 IS DECIMAL(9, 2)
                            .Length = CShort((From A In Amounts Select A.Length).Max)
                            .Scale = 0
                            If Amounts.Any Then
                                .Scale = CShort((From A In Amounts Where UBound(Split(A, ".")) > 0 Select Split(A, ".")(1).Length).Max)
                            End If
                            .DataType = "DECIMAL"
                            .DataFormat = "DECIMAL(" & .Length & ", " & .Scale & ")"
#End Region
                        End If

                    Case GetType(String)
#Region " CHAR + VARCHAR "
                        Dim ColumnEmpty As Boolean = Not Values.Any
                        If ColumnEmpty Then
                            .Length = Convert.ToInt16(.Name.Length)
                        Else
                            .Length = Convert.ToInt16({(From O In Values Select O.ToString.Length).Max, 2003}.Min)
                        End If
                        .Scale = 0
                        Dim Strings = From v In Values Select Convert.ToString(v, InvariantCulture)
                        If ColumnEmpty Then
                            REM /// PREFER VARCHAR
                            .DataType = "VARCHAR"
                            .DataFormat = "VARCHAR(" & .Length & ")"
                        Else
                            If (From s In Strings Where s.Length <> Database_Column.Length).Any Then
                                REM /// LENGTHS VARY...MUST BE VARCHAR
                                .DataType = "VARCHAR"
                                .DataFormat = "VARCHAR(" & .Length & ")"
                            Else
                                REM /// LENGTH OF EACH FIELD IS THE SAME...MUST BE CHAR
                                .DataType = "CHAR"
                                .DataFormat = "CHAR(" & .Length & ")"
                            End If
                        End If
#End Region
                    Case GetType(Boolean)
#Region " CHAR(1) "
                        'IBM® DB2® 9.x does Not implement a Boolean SQL type.
                        'Solution: The DB2 database interface converts BOOLEAN type to CHAR(1) columns And stores '1' or '0' values in the column.
                        .Length = 1
                        .Scale = 0
                        .DataType = "CHAR"
                        .DataFormat = "CHAR(" & .Length & ")"
#End Region
                    Case Else

                End Select
            End With
            Return Database_Column
        End If

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