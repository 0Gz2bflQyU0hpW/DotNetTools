Option Explicit On 
Option Strict On
Imports System.IO
Public Class Services
#Region "Dispose"
    Inherits ReleaseObj
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not mConection Is Nothing Then
                If Not InTransaction Then
                    With mConection
                        .Close()
                        .Dispose()
                    End With
                    mConection = Nothing
                    MyBase.Dispose(Disposing)
                End If
            End If
        End If
    End Sub
#End Region
#Region "Vars"
    Private Stringsql As String
    Private VsystemName As String = ""
    Private Vmodule As String = ""
    Private Vmethod As String = ""
    Private Vuser As String = ""
    Private mConection As System.Data.SqlClient.SqlConnection
    Private mTransaction As System.Data.SqlClient.SqlTransaction
    Public InTransaction As Boolean
    Private VTimeOut As String = ""
#End Region
#Region "Properties"
    Public Property StringConection() As String
        Get
            If Stringsql.Trim.Length = 0 Then
                Throw New System.Exception("No se puede establecer la StringCommand de conexión")
            End If
            Return Stringsql
        End Get
        Set(ByVal Value As String)
            Stringsql = Value
        End Set
    End Property
    Public Property SystemName() As String
        Get
            SystemName = VsystemName
        End Get
        Set(ByVal Value As String)
            If VsystemName.Trim.Length = 0 Then
                VsystemName = Value.Trim
            Else
                If LastValue(VsystemName) <> Value Then VsystemName = VsystemName.Trim & "." & Value.Trim
            End If
        End Set
    End Property
    Public Property ModuleName() As String
        Get
            ModuleName = Vmodule
        End Get
        Set(ByVal Value As String)
            If Vmodule.Trim.Length = 0 Then
                Vmodule = Value.Trim
            Else
                If LastValue(Vmodule) <> Value Then Vmodule = Vmodule.Trim & "." & Value.Trim
            End If
        End Set
    End Property
    Public Property Method() As String
        Get
            Method = Vmethod
        End Get
        Set(ByVal Value As String)
            If Vmethod.Trim.Length = 0 Then
                Vmethod = Value.Trim
            Else
                If LastValue(Vmethod) <> Value Then Vmethod = Vmethod.Trim & "." & Value.Trim
            End If
        End Set
    End Property
    Public Property User() As String
        Get
            User = Vuser
        End Get
        Set(ByVal Value As String)
            If Vuser.Trim.Length = 0 Then
                Vuser = Value.Trim
            Else
                If Vuser.Trim <> Value.Trim Then Vuser = Vuser.Trim & "." & Value.Trim
            End If
        End Set
    End Property
    Public Property TimeOut() As String
        Get
            TimeOut = VTimeOut
        End Get
        Set(ByVal Value As String)
            VTimeOut = Value.Trim
        End Set
    End Property
#End Region
#Region "Class Constructor"
    Sub New(ByVal StringConection As String)
        Stringsql = StringConection
    End Sub
#End Region
#Region "Users Functions"
#Region "DeleteWithFilter"
    Public Function DeleteWithFilter(ByVal Table As String, ByVal Filter As String) As Integer
        Dim sql As String
        sql = Table & "_DX_" & Filter
        Return Execute(sql)
    End Function
    Public Function DeleteWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Args As Integer) As Integer
        Dim sql As String
        sql = Table & "_DX_" & Filter
        Return Execute(sql, Args)
    End Function
    Public Function DeleteWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Args As DataTable) As Integer
        Dim sql As String
        sql = Table & "_DX_" & Filter
        Return Execute(sql, Args)
    End Function
    Public Function DeleteWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Args() As String) As Integer
        Dim sql As String
        sql = Table & "_DX_" & Filter
        Return Execute(sql, Args)
    End Function
#End Region
#Region "Delete"
    Public Function Delete(ByVal Table As String, ByVal Id As Integer) As Integer
        Delete = Execute(Table & "_E", Id)
    End Function
#End Region
#Region "GetAll"
    Public Function GetAll(ByVal Table As String) As DataTable
        Dim Sql As String
        Sql = Table & "_TT"
        GetAll = GetRecords(Sql)
    End Function
#End Region
#Region "GetWithFilter"
    Public Function GetWithFilter(ByVal Table As String, ByVal Filter As String, ByVal StringCommand As String) As DataTable
        Dim Sql As String
        Sql = Table & "_TX_" & Filter
        GetWithFilter = GetRecords(Sql, StringCommand)
    End Function
    Public Function GetWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Args As Integer) As DataTable
        Dim Sql As String
        Sql = Table & "_TX_" & Filter
        GetWithFilter = GetRecords(Sql, Args)
    End Function
    Public Function GetWithFilter(ByVal Table As String, ByVal Filter As String) As DataTable
        Dim Sql As String
        Sql = Table & "_TX_" & Filter
        GetWithFilter = GetRecords(Sql)
    End Function
    Public Function GetWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Args() As String) As DataTable
        Dim Sql As String
        Sql = Table & "_TX_" & Filter
        GetWithFilter = GetRecords(Sql, Args)
    End Function
    Public Function GetWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Args As DataTable) As DataTable
        Dim Sql As String
        Sql = Table & "_TX_" & Filter
        GetWithFilter = GetRecords(Sql, Args)
    End Function
#End Region
#Region "GetValue"
    Public Function GetValue(ByVal Table As String, ByVal Filter As String) As String
        Dim Sql As String
        Sql = Table & "_TV_" & Filter
        GetValue = ExecuteValue(Sql).ToString
    End Function
    Public Function GetValue(ByVal Table As String, ByVal Filter As String, ByVal Args As Integer) As String
        Dim Sql As String
        Sql = Table & "_TV_" & Filter
        GetValue = ExecuteValue(Sql, Args).ToString
    End Function
    Public Function GetValue(ByVal Table As String, ByVal Filter As String, ByVal StringCommand As String) As String
        Dim Sql As String
        Sql = Table & "_TV_" & Filter
        GetValue = ExecuteValue(Sql, StringCommand).ToString
    End Function
    Public Function GetValue(ByVal Table As String, ByVal Filter As String, ByVal Args() As String) As String
        Dim Sql As String
        Sql = Table & "_TV_" & Filter
        GetValue = ExecuteValue(Sql, Args).ToString
    End Function
    Public Function GetValue(ByVal Table As String, ByVal Filter As String, ByVal Args As DataTable) As String
        Dim Sql As String
        Sql = Table & "_TV_" & Filter
        GetValue = ExecuteValue(Sql, Args).ToString
    End Function
#End Region
#Region "GetOne"
    Public Function GetOne(ByVal Table As String, ByVal Id As Integer) As DataTable
        Dim Sql As String
        Sql = Table & "_T"
        GetOne = GetRecords(Sql, Id)
    End Function
#End Region
#Region "Add"
    Public Function Add(ByVal Table As String, ByVal Args As DataRow) As Integer
        Dim sql As String
        sql = Table & "_A"
        Add = Execute(sql, Args)
    End Function
    Public Function Add(ByVal Table As String, ByVal Args As DataTable) As Integer
        Dim sql As String
        sql = Table & "_A"
        Add = Execute(sql, Args)
    End Function
    Public Function Add(ByVal Table As String, ByVal Args() As String) As Integer
        Dim sql As String
        sql = Table & "_A"
        Add = Execute(sql, Args)
    End Function
#End Region
#Region "Update"
    Public Function Update(ByVal Table As String, ByVal Args() As String) As Integer
        Dim sql As String
        sql = Table & "_M"
        Update = Execute(sql, Args)
    End Function
    Public Function Update(ByVal Table As String, ByVal Args As DataRow) As Integer
        Dim sql As String
        sql = Table & "_M"
        Update = Execute(sql, Args)
    End Function
    Public Function Update(ByVal Table As String, ByVal Args As DataTable) As Integer
        Dim sql As String
        sql = Table & "_M"
        Update = Execute(sql, Args)
    End Function
    Public Function Update(ByVal Table As String, ByVal Filter As String) As Integer
        Dim sql As String
        sql = Table & "_M_" & Filter.Trim
        Update = Execute(sql)
    End Function
#End Region
#Region "UpdateWithFilter"
    Public Function UpdateWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Id As Integer) As Integer
        Dim sql As String
        sql = Table & "_MX_" & Filter
        UpdateWithFilter = Execute(sql, Id)
    End Function
    Public Function UpdateWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Args() As String) As Integer
        Dim Sql As String
        Sql = Table & "_MX_" & Filter
        UpdateWithFilter = Execute(Sql, Args)
    End Function
    Public Function UpdateWithFilter(ByVal Table As String, ByVal Filter As String, ByVal Datos As DataTable) As Integer
        Dim Sql As String
        Sql = Table & "_MX_" & Filter
        UpdateWithFilter = Execute(Sql, Datos)
    End Function
    Public Function UpdateWithFilter(ByVal Table As String, ByVal Filter As String) As Integer
        Dim sql As String
        sql = Table & "_MX_" & Filter
        UpdateWithFilter = Execute(sql)
    End Function
#End Region
    Public Function CheckConection() As String
        Try
            Dim AuxiConec As New System.Data.SqlClient.SqlConnection(Stringsql)
            AuxiConec.Open()
            AuxiConec.Close()
            CheckConection = "OK"
        Catch ex As Exception
            CheckConection = ex.Message
        End Try
    End Function
#End Region
#Region "Private Database Functions"
#Region "Conection"
    Private Sub Conection()
        If mConection Is Nothing Then mConection = New System.Data.SqlClient.SqlConnection(Stringsql)
        With mConection
            If mConection.State <> ConnectionState.Open Then .Open()
        End With
    End Sub
#End Region
#Region "ExecuteValue"
    Private Function ExecuteValue(ByVal Sql As String, ByVal Args() As String) As String
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Try
            SqlComm = LoadParameters(SqlComm, Args)
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                .Parameters(.Parameters.Count - 1).Value = ""
                .ExecuteNonQuery()
                ExecuteValue = .Parameters(.Parameters.Count - 1).Value.ToString.Trim
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            If Not SqlComm Is Nothing Then SqlComm.Dispose()
            SqlComm = Nothing
            ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
            Throw ex
        End Try
    End Function
    Private Function ExecuteValue(ByVal Sql As String, ByVal Args As DataTable) As String
        Dim lAffected As Integer
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                For lAffected = 0 To Args.Rows.Count - 1
                    .Parameters(lAffected + 1).Value = Args.Rows(0)(lAffected)
                Next
                .Parameters(.Parameters.Count - 1).Value = ""
                .ExecuteNonQuery()
                ExecuteValue = .Parameters(.Parameters.Count - 1).Value.ToString.Trim
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            If Not SqlComm Is Nothing Then SqlComm.Dispose()
            SqlComm = Nothing
            ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
            Throw ex
        End Try
    End Function
    Private Function ExecuteValue(ByVal Sql As String, ByVal Args As Integer) As String
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                If .Parameters.Count <> 3 Then
                    ExecuteValue = "Stored con demasiados Parametros"
                Else
                    .Parameters(1).Value = Args
                    Dim N As Integer
                    For N = 1 To .Parameters.Count - 1
                        If .Parameters(N).Direction = ParameterDirection.InputOutput Then
                            .Parameters(N).Value = ""
                            Exit For
                        End If
                    Next
                    .ExecuteNonQuery()
                    ExecuteValue = .Parameters(N).Value.ToString.Trim
                    .Dispose()
                End If
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            If Not SqlComm Is Nothing Then SqlComm.Dispose()
            SqlComm = Nothing
            ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
            Throw ex
        End Try
    End Function
    Private Function ExecuteValue(ByVal Sql As String, ByVal StringCommand As String) As String
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                If .Parameters.Count <> 3 Then
                    ExecuteValue = "Stored con demasiados Parametros"
                Else
                    .Parameters(1).Value = StringCommand
                    Dim N As Integer
                    For N = 1 To SqlComm.Parameters.Count - 1
                        If .Parameters(N).Direction = ParameterDirection.InputOutput Then
                            .Parameters(N).Value = ""
                            Exit For
                        End If
                    Next
                    .ExecuteNonQuery()
                    ExecuteValue = .Parameters(N).Value.ToString.Trim
                    .Dispose()
                End If
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            If Not SqlComm Is Nothing Then SqlComm.Dispose()
            SqlComm = Nothing
            ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
            Throw ex
        End Try
    End Function
    Private Function ExecuteValue(ByVal Sql As String) As String
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                .ExecuteNonQuery()
                ExecuteValue = .Parameters(1).Value.ToString.Trim
                .Dispose()
            End With
        Catch ex As Exception
            If Not SqlComm Is Nothing Then SqlComm.Dispose()
            ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
            Throw ex
        End Try
    End Function
#End Region
#Region "Execute"
    Public Function Execute(ByVal Sql As String, ByVal Args() As String) As Integer
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                SqlComm = LoadParameters(SqlComm, Args)
                .ExecuteNonQuery()
                Execute = ToInt(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                If Not SqlComm Is Nothing Then .Dispose()
                SqlComm = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function Execute(ByVal Sql As String) As Integer
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                .ExecuteNonQuery()
                If .Parameters.Count > 0 Then
                    Execute = ToInt(.Parameters(0).Value.ToString)
                End If
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                If Not SqlComm Is Nothing Then .Dispose()
                SqlComm = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function Execute(ByVal Sql As String, ByVal Args As DataRow) As Integer
        Dim ColData As Integer
        Dim IdNoVa As Integer = 0
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        If Sql.Substring(Sql.Trim.Length - 2, 2) = "_A" Then IdNoVa = 1
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                For ColData = 0 + IdNoVa To Args.Table.Columns.Count - 1
                    .Parameters(ColData - IdNoVa + 1).Value = Args(ColData)
                Next
                .ExecuteNonQuery()
                Execute = Convert.ToInt32(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                If Not SqlComm Is Nothing Then .Dispose()
                SqlComm = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function Execute(ByVal Sql As String, ByVal Args As DataTable) As Integer
        Dim Posicion As Integer, ColData As Integer
        Dim IdNoVa As Integer = 0
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        If Sql.Substring(Sql.Trim.Length - 2, 2) = "_A" Then
            IdNoVa = 1
        End If
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                For Posicion = 0 To Args.Rows.Count - 1
                    For ColData = 0 + IdNoVa To Args.Columns.Count - 1
                        .Parameters(ColData - IdNoVa + 1).Value = Args.Rows(Posicion)(ColData)
                    Next
                    .ExecuteNonQuery()
                Next
                Execute = ToInt(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
                Args.Dispose()
            Catch ex As Exception
                If Not SqlComm Is Nothing Then .Dispose()
                Args.Dispose()
                SqlComm = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function Execute(ByVal Sql As String, ByVal Args As Integer) As Integer
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                .Parameters(1).Value = Args
                .ExecuteNonQuery()
                Execute = ToInt(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                If Not SqlComm Is Nothing Then .Dispose()
                SqlComm = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function Execute(ByVal Sql As String, ByVal Args As String) As Integer
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                .Parameters(1).Value = Args
                .ExecuteNonQuery()
                Execute = ToInt(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                If Not SqlComm Is Nothing Then .Dispose()
                SqlComm = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function ExecuteTxtCommand(ByVal Sql As String) As Integer
        Conection()
        Dim mCommand As New System.Data.SqlClient.SqlCommand(Sql, mConection)
        With mCommand
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
            .Connection = mConection
            .Transaction = mTransaction
        End With
        With mCommand
            Try
                .ExecuteNonQuery()
                .Dispose()
                mCommand = Nothing
            Catch ex As Exception
                If Not mCommand Is Nothing Then .Dispose()
                mCommand = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
#End Region
#Region "CreateCommand"
    Private Function CreateCommand(ByVal Procedimiento As String) As System.Data.SqlClient.SqlCommand
        'Dim ConectionAux As New System.Data.SqlClient.SqlConnection(StringConection)
        'ConectionAux.Open()
        Dim NewConection As Boolean = True
        Dim ConectionAux As System.Data.SqlClient.SqlConnection
        If Not mConection Is Nothing Then
            If mConection.State = ConnectionState.Open And Not InTransaction Then
                ConectionAux = mConection
                NewConection = False
            Else
                ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
                ConectionAux.Open()
            End If
        Else
            ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
            ConectionAux.Open()
        End If
        Dim mCommand As New System.Data.SqlClient.SqlCommand(Procedimiento, ConectionAux)
        Dim mConstructor As New System.Data.SqlClient.SqlCommandBuilder
        mCommand.CommandType = CommandType.StoredProcedure
        If ParameterCache.IsParametersCached(mCommand) Then
            Dim param As System.Data.SqlClient.SqlParameter() = DirectCast(ParameterCache.GetCachedParameters(mCommand), System.Data.SqlClient.SqlParameter())
            mCommand.Parameters.AddRange(param)
        Else
            SqlClient.SqlCommandBuilder.DeriveParameters(mCommand)
            ParameterCache.CacheParameters(mCommand)
        End If
        With mCommand
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
            .Connection = mConection
            .Transaction = mTransaction
        End With
        If NewConection Then
            With ConectionAux
                .Close()
                .Dispose()
            End With
            ConectionAux = Nothing
        End If
        mConstructor.Dispose()
        Return mCommand
    End Function
    Private Function CreateCommandFromTxt(ByVal command As String) As System.Data.SqlClient.SqlCommand
        'Dim ConectionAux As New System.Data.SqlClient.SqlConnection(StringConection)
        'ConectionAux.Open()
        Dim NewConection As Boolean = True
        Dim ConectionAux As System.Data.SqlClient.SqlConnection
        If Not mConection Is Nothing Then
            If mConection.State = ConnectionState.Open And Not InTransaction Then
                ConectionAux = mConection
                NewConection = False
            Else
                ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
                ConectionAux.Open()
            End If
        Else
            ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
            ConectionAux.Open()
        End If
        Dim mCommand As New System.Data.SqlClient.SqlCommand(command)
        Dim mConstructor As New System.Data.SqlClient.SqlCommandBuilder
        mCommand.CommandType = CommandType.Text
        With mCommand
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
            .Connection = mConection
            .Transaction = mTransaction
        End With
        If NewConection Then
            With ConectionAux
                .Close()
                .Dispose()
            End With
            ConectionAux = Nothing
        End If
        mConstructor.Dispose()
        Return mCommand
    End Function
#End Region
#Region "LoadParameters"
    Private Function LoadParameters(ByVal Command As System.Data.SqlClient.SqlCommand, ByVal Args() As String) As System.Data.SqlClient.SqlCommand
        Dim Posicion As Integer
        Dim TipoParametro As SqlDbType
        With Command
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
            For Posicion = 0 To Args.Length - 1
                TipoParametro = .Parameters(Posicion + 1).SqlDbType
                If Args(Posicion).Trim.ToUpper = "DBNULL" Then
                    .Parameters(Posicion + 1).Value = System.DBNull.Value
                Else
                    Select Case TipoParametro
                        Case SqlDbType.BigInt
                            .Parameters(Posicion + 1).Value = CType(Args(Posicion), Int64)
                        Case SqlDbType.Binary
                        Case SqlDbType.Bit
                            .Parameters(Posicion + 1).Value = CLogico(Args(Posicion)).ToString.Trim
                        Case SqlDbType.Char
                            .Parameters(Posicion + 1).Value = Args(Posicion).Trim
                        Case SqlDbType.DateTime
                            .Parameters(Posicion + 1).Value = Args(Posicion)
                        Case SqlDbType.Decimal
                            .Parameters(Posicion + 1).Value = CDecimal(Args(Posicion))
                        Case SqlDbType.Float
                            .Parameters(Posicion + 1).Value = CDoble(Args(Posicion))
                        Case SqlDbType.Image
                        Case SqlDbType.Int
                            .Parameters(Posicion + 1).Value = ToInt(Args(Posicion))
                        Case SqlDbType.Money
                            .Parameters(Posicion + 1).Value = CDoble(Args(Posicion))
                        Case SqlDbType.NChar
                            .Parameters(Posicion + 1).Value = Args(Posicion).Trim
                        Case SqlDbType.NText
                            .Parameters(Posicion + 1).Value = Args(Posicion).Trim
                        Case SqlDbType.NVarChar
                            .Parameters(Posicion + 1).Value = Args(Posicion).Trim
                        Case SqlDbType.Real
                        Case SqlDbType.SmallDateTime
                            .Parameters(Posicion + 1).Value = Args(Posicion)
                        Case SqlDbType.SmallInt
                            .Parameters(Posicion + 1).Value = ToInt(Args(Posicion))
                        Case SqlDbType.SmallMoney
                            .Parameters(Posicion + 1).Value = CDoble(Args(Posicion))
                        Case SqlDbType.Text
                            .Parameters(Posicion + 1).Value = Args(Posicion).Trim
                        Case SqlDbType.Xml
                            .Parameters(Posicion + 1).Value = Args(Posicion).Trim
                        Case SqlDbType.Timestamp
                        Case SqlDbType.TinyInt
                        Case SqlDbType.UniqueIdentifier
                        Case SqlDbType.VarBinary
                        Case SqlDbType.VarChar
                            .Parameters(Posicion + 1).Value = Args(Posicion).Trim
                        Case SqlDbType.Variant
                    End Select
                End If
            Next
        End With
        LoadParameters = Command
    End Function
#End Region
#Region "GetRecords"
    Public Function GetRecords(ByVal Sql As String, ByVal Args() As String) As DataTable
        Dim Datatable As New DataTable
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
            Try
                SqlComm = LoadParameters(SqlComm, Args)
                DataAdapter.Fill(Datatable)
                GetRecords = Datatable
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
            Catch ex As Exception
                DataAdapter.Dispose()
                DataAdapter = Nothing
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function GetRecords(ByVal Sql As String, ByVal Args As DataTable) As DataTable
        Dim DataTable As New DataTable
        Dim lAffected As Integer
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
            Try
                For lAffected = 0 To Args.Rows.Count() - 1
                    .Parameters(lAffected + 1).Value = Args.Rows(0)(lAffected)
                Next
                DataAdapter.Fill(DataTable)
                GetRecords = DataTable
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
                Args.Dispose()
            Catch ex As Exception
                DataAdapter.Dispose()
                DataAdapter = Nothing
                .Dispose()
                SqlComm = Nothing
                Args.Dispose()
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function GetRecords(ByVal Sql As String) As DataTable
        Dim DataTable As New DataTable
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        Try
            DataAdapter.Fill(DataTable)
            GetRecords = DataTable
            SqlComm.Dispose()
            DataAdapter.Dispose()
            DataTable.Dispose()
            SqlComm = Nothing
            DataAdapter = Nothing
        Catch ex As Exception
            SqlComm.Dispose()
            DataAdapter.Dispose()
            DataTable.Dispose()
            SqlComm = Nothing
            DataAdapter = Nothing
            ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
            Throw ex
        End Try
    End Function
    Public Function GetRecords(ByVal Sql As String, ByVal Args As Integer) As DataTable
        Dim Datatable As New DataTable
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                .Parameters(1).Value = Args
                DataAdapter.Fill(Datatable)
                GetRecords = Datatable
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
            Catch ex As Exception
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function GetRecords(ByVal Sql As String, ByVal StringCommand As String) As DataTable
        Dim Datatable As New DataTable
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommand(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            Try
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = ToInt(VTimeOut)
                If StringCommand.Trim.ToUpper = "DBNULL" Then
                    .Parameters(1).Value = System.DBNull.Value
                Else
                    .Parameters(1).Value = StringCommand.Trim
                End If
                DataAdapter.Fill(Datatable)
                GetRecords = Datatable
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
            Catch ex As Exception
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
                ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
                Throw ex
            End Try
        End With
    End Function
    Public Function GetRecordsFromTxtCommand(ByVal Command As String) As DataTable
        Dim DataTable As New DataTable
        Conection()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = CreateCommandFromTxt(Command)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        Try
            DataAdapter.Fill(DataTable)
            GetRecordsFromTxtCommand = DataTable
            SqlComm.Dispose()
            DataAdapter.Dispose()
            DataTable.Dispose()
            SqlComm = Nothing
            DataAdapter = Nothing
        Catch ex As Exception
            SqlComm.Dispose()
            DataAdapter.Dispose()
            DataTable.Dispose()
            SqlComm = Nothing
            DataAdapter = Nothing
            ErrsAnalyzer(ex.Message, ex.Source, ex.StackTrace.Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX").Replace("Rslutzki", "XXXX").Replace("ColmanSaks", "XXXX"))
            Throw ex
        End Try
    End Function
#End Region
#Region "ErrsAnalyzer"
    Private Sub ErrsAnalyzer(ByVal TxtError As String, ByVal Source As String, ByVal StackTrace As String)
        Dim Aplication As String = ""
        Dim ProductName As String = "SqlDataServer"
        Dim File As String = ""
        Dim Drive As String = ""
        If System.Configuration.ConfigurationManager.AppSettings("DriveAppErr") Is Nothing Then
            If Drive.Trim.Length = 0 Then Drive = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location).Substring(0, 2)
            If Drive.Trim.Length = 0 Then Drive = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location).Substring(0, 2)
        Else
            Drive = System.Configuration.ConfigurationManager.AppSettings("DriveAppErr").ToString
        End If
        If File.Trim.Length = 0 Then
            Dim Path As String = ""
            If Drive.Trim.Length = 0 Then
                Path = "c:\AppErrs\"
            Else
                Path = Drive & "\AppErrs\"
            End If
            Path = Path & DateTime.Now.Year.ToString & "\" & (DateTime.Now.Month + 100).ToString.Substring(1, 2) & "\" & (DateTime.Now.Day + 100).ToString.Substring(1, 2)
            If Not IO.Directory.Exists(Path) Then IO.Directory.CreateDirectory(Path)
            If Aplication.Trim.Length = 0 Then
                File = Path & "\Errs.xml"
            Else
                File = Path & "\" & Aplication.Trim & ".xml"
            End If
        End If
        Dim MidataSet As New DataSet
        MidataSet.DataSetName = "MisErrores"
        If IO.File.Exists(File) Then
            MidataSet.ReadXml(File)
        Else
            Dim MisRegistros As New DataTable
            With MisRegistros
                .Columns.Add("ProductName", System.Type.GetType("System.String"))
                .Columns.Add("Aplication", System.Type.GetType("System.String"))
                .Columns.Add("Class", System.Type.GetType("System.String"))
                .Columns.Add("Function", System.Type.GetType("System.String"))
                .Columns.Add("Date", System.Type.GetType("System.String"))
                .Columns.Add("Time", System.Type.GetType("System.String"))
                .Columns.Add("Source", System.Type.GetType("System.String"))
                .Columns.Add("ErrDescription", System.Type.GetType("System.String"))
                .Columns.Add("User", System.Type.GetType("System.String"))
                .Columns.Add("StackTrace", System.Type.GetType("System.String"))
            End With
            MidataSet.Tables.Add(MisRegistros)
            MisRegistros.Dispose()
        End If
        Dim Args(9) As String
        Args(0) = ProductName
        Args(1) = ""
        Args(2) = ""
        Args(3) = ""
        Args(4) = System.DateTime.Now.Day.ToString.Trim & "/" & System.DateTime.Now.Month.ToString.Trim & "/" & System.DateTime.Now.Year.ToString.Trim
        Args(5) = System.DateTime.Now.Hour.ToString.Trim & ":" & System.DateTime.Now.Minute.ToString.Trim & ":" & System.DateTime.Now.Second.ToString.Trim
        Args(6) = Source
        Args(7) = TxtError
        Args(8) = User
        Args(9) = StackTrace
        With MidataSet
            .Tables(0).Rows.Add(Args)
            .WriteXml(File)
            .Dispose()
        End With
        Args = Nothing
    End Sub
    Public Function ErrCode(ByVal TheErr As String, ByVal Code As Boolean) As String
        Dim ConstErr As String = "DELETE statement conflicted with COLUMN REFERENCE constraint"
        If ConstErr.Trim.Length <= TheErr.Trim.Length Then
            If TheErr.Substring(0, ConstErr.Length).Trim = ConstErr Then
                If Code Then
                    Return "-3"
                Else
                    Return "DELETE statement conflicted with COLUMN REFERENCE constraint"
                End If
            End If
        End If
        ConstErr = "Violation of UNIQUE KEY"
        If ConstErr.Trim.Length <= TheErr.Trim.Length Then
            If TheErr.Substring(0, ConstErr.Length).Trim = ConstErr Then
                If Code Then
                    Return "-1"
                Else
                    Return "Violation of UNIQUE KEY"
                End If
            End If
        End If
        ConstErr = "INSERT statement conflicted with COLUMN FOREIGN KEY constraint"
        If ConstErr.Trim.Length <= TheErr.Trim.Length Then
            If TheErr.Substring(0, ConstErr.Length).Trim = ConstErr Then
                If Code Then
                    Return "-6"
                Else
                    Return "INSERT statement conflicted with COLUMN FOREIGN KEY constraint"
                End If
            End If
        End If
        ConstErr = "The stored procedure"
        If ConstErr.Trim.Length <= TheErr.Trim.Length Then
            If TheErr.Substring(0, ConstErr.Length).Trim = ConstErr Then
                If TheErr.Substring(TheErr.Trim.Length - 14, 14) = "doesn't exist." Then
                    If Code Then
                        Return "-5"
                    Else
                        Return "A stored procedure was not found"
                    End If
                End If
            End If
        End If
        Select Case TheErr
            Case Is = "SQL Server does not exist or access denied."
                If Code Then
                    Return "-2"
                Else
                    Return TheErr
                End If
            Case Is = "The ConnectionString property has not been initialized."
                If Code Then
                    Return "-4"
                Else
                    Return TheErr
                End If
            Case Else
                If Code Then
                    Return "-999999"
                Else
                    Return "Unknow Err"
                End If
        End Select
    End Function
#End Region
#End Region
#Region "Funciones de Formato Privadas"
#Region "EsNumero"
    Private Function EsNumero(ByVal Numero As String) As Boolean
        Try
            Convert.ToDouble(Numero)
        Catch
            Return False
        End Try
        Return True
    End Function
#End Region
#Region "Cdecimal"
    Private Function CDecimal(ByVal Numero As String) As Double
        If Numero.Trim.Length = 0 Then Return 0
        Try
            Dim entero As New System.Text.StringBuilder
            Dim decimales As New System.Text.StringBuilder
            Dim Divisor As New System.Text.StringBuilder
            Divisor.Append("1")
            Dim caracter As String
            Dim bandera As Boolean = False
            Dim n As Integer
            For n = 0 To Numero.Trim.Length - 1
                caracter = Izquierda(Derecha(Numero, Numero.Length - n), 1)
                If Not EsNumero(caracter) And caracter <> "-" And caracter <> "+" Then
                    bandera = True
                Else
                    If Not bandera Then
                        entero.Append(caracter)
                    Else
                        decimales.Append(caracter)
                        Divisor.Append("0")
                    End If
                End If
            Next
            CDecimal = Decimal.Parse(entero.ToString)
            If decimales.ToString.Trim.Length <> 0 Then
                If CDecimal >= 0 Then
                    CDecimal = CDecimal + (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                Else
                    CDecimal = CDecimal - (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                End If
            End If
            If entero.ToString = "-0" Then CDecimal = CDecimal * -1
        Catch
            Return 0
        End Try
    End Function
#End Region
    Private Function ToInt(ByVal Numero As String) As Integer
        If Numero.ToLower = "true" Or Numero.ToLower = "verdadero" Then
            ToInt = 1
            Exit Function
        End If
        If Numero.ToLower = "false" Or Numero.ToLower = "falso" Then
            ToInt = 0
            Exit Function
        End If
        Try
            ToInt = Integer.Parse(Numero)
        Catch
            Return 0
        End Try
    End Function
    Private Function CLogico(ByVal Numero As String) As Boolean
        If Numero.Trim = "0" Then
            CLogico = False
            Exit Function
        End If
        If Numero.Trim = "1" Then
            CLogico = True
            Exit Function
        End If
        Try
            CLogico = Boolean.Parse(Numero)
        Catch
            Return False
        End Try
    End Function
#Region "Cdoble"
    Private Function CDoble(ByVal Numero As String) As Double
        If Numero.Trim.Length = 0 Then Return 0
        Try
            Dim entero As New System.Text.StringBuilder
            Dim decimales As New System.Text.StringBuilder
            Dim Divisor As New System.Text.StringBuilder
            Divisor.Append("1")
            Dim caracter As String
            Dim bandera As Boolean = False
            Dim n As Integer
            For n = 0 To Numero.Trim.Length - 1
                caracter = Izquierda(Derecha(Numero, Numero.Length - n), 1)
                If Not EsNumero(caracter) And caracter <> "-" And caracter <> "+" Then
                    bandera = True
                Else
                    If Not bandera Then
                        entero.Append(caracter)
                    Else
                        decimales.Append(caracter)
                        Divisor.Append("0")
                    End If
                End If
            Next
            CDoble = Double.Parse(entero.ToString)
            If decimales.ToString.Trim.Length <> 0 Then
                If CDoble >= 0 Then
                    CDoble = CDoble + (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                Else
                    CDoble = CDoble - (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                End If
            End If
            If entero.ToString = "-0" Then CDoble = CDoble * -1
        Catch
            Return 0
        End Try
    End Function
#End Region
#Region "Izquierda"
    Private Function Izquierda(ByVal StringCommand As String, ByVal Posiciones As Integer) As String
        If Posiciones > StringCommand.Trim.Length Then Return StringCommand
        Return StringCommand.Trim.Substring(0, Posiciones)
    End Function
#End Region
#Region "Derecha"
    Private Function Derecha(ByVal StringCommand As String, ByVal Posiciones As Integer) As String
        If Posiciones > StringCommand.Trim.Length Then Return StringCommand
        Return StringCommand.Trim.Substring(StringCommand.Trim.Length - Posiciones, Posiciones)
    End Function
#End Region
    Private Function CFecha(ByVal Fecha As String) As DateTime
        Try
            CFecha = DateTime.Parse(Fecha)
        Catch
            Return System.DateTime.Now
        End Try
    End Function
    Public Shared Function LastValue(ByVal Texto As String) As String
        LastValue = ""
        Texto = Texto.Trim
        If Texto.Length = 0 Then Return LastValue
        Dim Largo As Integer = Texto.Length
        Dim N As Integer
        For N = Largo - 1 To 0 Step -1
            If Texto.Substring(N, 1) = "." Then
                LastValue = Texto.Substring(N + 1, Largo - 1 - N)
                Exit For
            End If
        Next
        If LastValue = "" Then LastValue = Texto
    End Function
#End Region
#Region "Transacciones"
    Public Sub BeginTransaccion()
        Conection()
        mTransaction = mConection.BeginTransaction
        InTransaction = True
    End Sub
    Public Sub EndTransaccion()
        Try
            With mTransaction
                .Commit()
                .Dispose()
            End With
            InTransaction = False
            mTransaction = Nothing
        Catch ex As System.Exception
            mTransaction.Connection.Close()
            InTransaction = False
            mTransaction = Nothing
            Throw ex
        End Try
    End Sub
    Public Sub CancelTransaccion()
        Try
            If InTransaction Then
                With mTransaction
                    .Rollback()
                    .Dispose()
                End With
            End If
            mTransaction = Nothing
            InTransaction = False
        Catch Ex As System.Exception
            If Not mTransaction Is Nothing Then
                mTransaction.Connection.Close()
            End If
            mTransaction = Nothing
            InTransaction = False
            Throw Ex
        End Try
    End Sub
#End Region
End Class



