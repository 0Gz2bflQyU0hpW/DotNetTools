Option Explicit On 
Option Strict On
Imports System.IO
Public Class Servicios
#Region "Implementacion del Dispose"
    Inherits LiberarObj
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not mConexion Is Nothing Then
                If Not EnTransaccion Then
                    With mConexion
                        .Close()
                        .Dispose()
                    End With
                    mConexion = Nothing
                    MyBase.Dispose(Disposing)
                End If
            End If
        End If
    End Sub
#End Region
#Region "Variables"
    Private Stringsql As String
    Private VSistema As String = ""
    Private VModulo As String = ""
    Private VMetodo As String = ""
    Private Vusuario As String = ""
    Private mConexion As System.Data.SqlClient.SqlConnection
    Private mTransaccion As System.Data.SqlClient.SqlTransaction
    Public EnTransaccion As Boolean
    Private VTimeOut As String = ""
#End Region
#Region "Propiedades"
    Public Property CadenaConexion() As String
        Get
            If Stringsql.Trim.Length = 0 Then
                Throw New System.Exception("No se puede establecer la cadena de conexión")
            End If
            Return Stringsql
        End Get
        Set(ByVal Value As String)
            Stringsql = Value
        End Set
    End Property
    Public Property Sistema() As String
        Get
            Sistema = VSistema
        End Get
        Set(ByVal Value As String)
            If VSistema.Trim.Length = 0 Then
                VSistema = Value.Trim
            Else
                If UltimoValor(VSistema) <> Value Then VSistema = VSistema.Trim & "." & Value.Trim
            End If
        End Set
    End Property
    Public Property Modulo() As String
        Get
            Modulo = VModulo
        End Get
        Set(ByVal Value As String)
            If VModulo.Trim.Length = 0 Then
                VModulo = Value.Trim
            Else
                If UltimoValor(VModulo) <> Value Then VModulo = VModulo.Trim & "." & Value.Trim
            End If
        End Set
    End Property
    Public Property Metodo() As String
        Get
            Metodo = VMetodo
        End Get
        Set(ByVal Value As String)
            If VMetodo.Trim.Length = 0 Then
                VMetodo = Value.Trim
            Else
                If UltimoValor(VMetodo) <> Value Then VMetodo = VMetodo.Trim & "." & Value.Trim
            End If
        End Set
    End Property
    Public Property Usuario() As String
        Get
            Usuario = Vusuario
        End Get
        Set(ByVal Value As String)
            If Vusuario.Trim.Length = 0 Then
                Vusuario = Value.Trim
            Else
                If Vusuario.Trim <> Value.Trim Then Vusuario = Vusuario.Trim & "." & Value.Trim
            End If
        End Set
    End Property
#End Region
#Region "Constructores de la clase"
    Sub New(ByVal CadenaConexion As String)
        Stringsql = CadenaConexion
    End Sub
    Sub New(ByVal CadenaConexion As String, ByVal TimeOut As Integer)
        VTimeOut = TimeOut.ToString.Trim
        If CadenaConexion.Trim.Substring(CadenaConexion.Trim.Length - 1, 1) <> ";" Then
            CadenaConexion = CadenaConexion & ";Connect Timeout=" & TimeOut.ToString
        Else
            CadenaConexion = CadenaConexion & "Connect Timeout=" & TimeOut.ToString
        End If
        Stringsql = CadenaConexion
    End Sub
#End Region
#Region "Funciones de Interfase con el usuario"
#Region "BorrarConFiltro"
    Public Function BorrarConFiltro(ByVal Tabla As String, ByVal Filtro As String) As Integer
        Dim sql As String
        sql = Tabla & "_DX_" & Filtro
        Return Ejecutar(sql)
    End Function
    Public Function BorrarConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args As Integer) As Integer
        Dim sql As String
        sql = Tabla & "_DX_" & Filtro
        Return Ejecutar(sql, Args)
    End Function
    Public Function BorrarConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args As DataTable) As Integer
        Dim sql As String
        sql = Tabla & "_DX_" & Filtro
        Return Ejecutar(sql, Args)
    End Function
    Public Function BorrarConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args() As String) As Integer
        Dim sql As String
        sql = Tabla & "_DX_" & Filtro
        Return Ejecutar(sql, Args)
    End Function
#End Region
#Region "Borrar"
    Public Function Borrar(ByVal Tabla As String, ByVal Id As Integer) As Integer
        Borrar = Ejecutar(Tabla & "_E", Id)
    End Function
#End Region
#Region "TraerTodos"
    Public Function TraerTodos(ByVal Tabla As String) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TT"
        TraerTodos = TraerRegistros(Sql)
    End Function
#End Region
#Region "TraerConFiltro"
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Cadena As String) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql, Cadena)
    End Function
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args As Integer) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql, Args)
    End Function
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql)
    End Function
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args() As String) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql, Args)
    End Function
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args As DataTable) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql, Args)
    End Function
#End Region
#Region "TraerValor"
    Public Function TraerValor(ByVal Tabla As String, ByVal Filtro As String) As String
        Dim Sql As String
        Sql = Tabla & "_TV_" & Filtro
        TraerValor = EjecutarValor(Sql).ToString
    End Function
    Public Function TraerValor(ByVal Tabla As String, ByVal Filtro As String, ByVal Args As Integer) As String
        Dim Sql As String
        Sql = Tabla & "_TV_" & Filtro
        TraerValor = EjecutarValor(Sql, Args).ToString
    End Function
    Public Function TraerValor(ByVal Tabla As String, ByVal Filtro As String, ByVal Cadena As String) As String
        Dim Sql As String
        Sql = Tabla & "_TV_" & Filtro
        TraerValor = EjecutarValor(Sql, Cadena).ToString
    End Function
    Public Function TraerValor(ByVal Tabla As String, ByVal Filtro As String, ByVal Args() As String) As String
        Dim Sql As String
        Sql = Tabla & "_TV_" & Filtro
        TraerValor = EjecutarValor(Sql, Args).ToString
    End Function
    Public Function TraerValor(ByVal Tabla As String, ByVal Filtro As String, ByVal Args As DataTable) As String
        Dim Sql As String
        Sql = Tabla & "_TV_" & Filtro
        TraerValor = EjecutarValor(Sql, Args).ToString
    End Function
#End Region
#Region "TraerUno"
    Public Function TraerUno(ByVal Tabla As String, ByVal Id As Integer) As DataTable
        Dim Sql As String
        Sql = Tabla & "_T"
        TraerUno = TraerRegistros(Sql, Id)
    End Function
#End Region
#Region "Agregar"
    Public Function Agregar(ByVal Tabla As String, ByVal Args As DataTable) As Integer
        Dim sql As String
        sql = Tabla & "_A"
        Agregar = Ejecutar(sql, Args)
    End Function
    Public Function Agregar(ByVal Tabla As String, ByVal Args() As String) As Integer
        Dim sql As String
        sql = Tabla & "_A"
        Agregar = Ejecutar(sql, Args)
    End Function
#End Region
#Region "Modificar"
    Public Function Modificar(ByVal Tabla As String, ByVal Args() As String) As Integer
        Dim sql As String
        sql = Tabla & "_M"
        Modificar = Ejecutar(sql, Args)
    End Function
    Public Function Modificar(ByVal Tabla As String, ByVal Args As DataTable) As Integer
        Dim sql As String
        sql = Tabla & "_M"
        Modificar = Ejecutar(sql, Args)
    End Function
    Public Function Modificar(ByVal Tabla As String, ByVal Filtro As String) As Integer
        Dim sql As String
        sql = Tabla & "_M_" & Filtro.Trim
        Modificar = Ejecutar(sql)
    End Function
#End Region
#Region "Modificar con Filtro"
    Public Function ModificarConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Id As Integer) As Integer
        Dim sql As String
        sql = Tabla & "_MX_" & Filtro
        ModificarConFiltro = Ejecutar(sql, Id)
    End Function
    Public Function ModificarConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args() As String) As Integer
        Dim Sql As String
        Sql = Tabla & "_MX_" & Filtro
        ModificarConFiltro = Ejecutar(Sql, Args)
    End Function
    Public Function ModificarConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Datos As DataTable) As Integer
        Dim Sql As String
        Sql = Tabla & "_MX_" & Filtro
        ModificarConFiltro = Ejecutar(Sql, Datos)
    End Function
    Public Function ModificarConFiltro(ByVal Tabla As String, ByVal Filtro As String) As Integer
        Dim sql As String
        sql = Tabla & "_MX_" & Filtro
        ModificarConFiltro = Ejecutar(sql)
    End Function
#End Region
#End Region
#Region "Funciones Privadas de Base da datos"
#Region "Conexion"
    Private Sub Conexion()
        If mConexion Is Nothing Then
            mConexion = New System.Data.SqlClient.SqlConnection(Stringsql)
        End If
        With mConexion
            If .State <> ConnectionState.Open Then .Open()
        End With
    End Sub
#End Region
#Region "Ejecutar devolviendo valor"
    Private Function EjecutarValor(ByVal Sql As String, ByVal Args() As String) As String
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Try
            SqlComm = CargarParametros(SqlComm, Args)
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
                .Parameters(.Parameters.Count - 1).Value = ""
                .ExecuteNonQuery()
                EjecutarValor = .Parameters(.Parameters.Count - 1).Value.ToString
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            SqlComm.Dispose()
            SqlComm = Nothing
            EjecutarValor = AnalizoError(ex.Message, ex.Source)
            Throw ex
        End Try
    End Function
    Private Function EjecutarValor(ByVal Sql As String, ByVal Args As DataTable) As String
        Dim lAffected As Integer
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
                For lAffected = 0 To Args.Rows.Count - 1
                    .Parameters(lAffected + 1).Value = Args.Rows(0)(lAffected)
                Next
                .Parameters(.Parameters.Count - 1).Value = ""
                .ExecuteNonQuery()
                EjecutarValor = .Parameters(.Parameters.Count - 1).Value.ToString
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            SqlComm.Dispose()
            SqlComm = Nothing
            EjecutarValor = AnalizoError(ex.Message, ex.Source)
            Throw ex
        End Try
    End Function
    Private Function EjecutarValor(ByVal Sql As String, ByVal Args As Integer) As String
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
                If .Parameters.Count <> 3 Then
                    EjecutarValor = "Stored con demasiados Parametros"
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
                    EjecutarValor = .Parameters(N).Value.ToString
                    .Dispose()
                End If
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            SqlComm.Dispose()
            SqlComm = Nothing
            EjecutarValor = AnalizoError(ex.Message, ex.Source)
            Throw ex
        End Try
    End Function
    Private Function EjecutarValor(ByVal Sql As String, ByVal Cadena As String) As String
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
                If .Parameters.Count <> 3 Then
                    EjecutarValor = "Stored con demasiados Parametros"
                Else
                    .Parameters(1).Value = Cadena
                    Dim N As Integer
                    For N = 1 To SqlComm.Parameters.Count - 1
                        If .Parameters(N).Direction = ParameterDirection.InputOutput Then
                            .Parameters(N).Value = ""
                            Exit For
                        End If
                    Next
                    .ExecuteNonQuery()
                    EjecutarValor = .Parameters(N).Value.ToString
                    .Dispose()
                End If
                .Dispose()
            End With
            SqlComm = Nothing
        Catch ex As Exception
            SqlComm.Dispose()
            SqlComm = Nothing
            EjecutarValor = AnalizoError(ex.Message, ex.Source)
            Throw ex
        End Try
    End Function
    Private Function EjecutarValor(ByVal Sql As String) As String
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Try
            With SqlComm
                If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
                .ExecuteNonQuery()
                EjecutarValor = .Parameters(1).Value.ToString
                .Dispose()
            End With
        Catch ex As Exception
            SqlComm.Dispose()
            EjecutarValor = AnalizoError(ex.Message, ex.Source)
            Throw ex
        End Try
    End Function
#End Region
#Region "Ejecutar"
    Public Function Ejecutar(ByVal Sql As String, ByVal Args() As String) As Integer
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                SqlComm = CargarParametros(SqlComm, Args)
                .ExecuteNonQuery()
                Ejecutar = CEntero(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                .Dispose()
                SqlComm = Nothing
                Ejecutar = CEntero(AnalizoError(ex.Message, ex.Source))
                Throw ex
            End Try
        End With
    End Function
    Public Function Ejecutar(ByVal Sql As String) As Integer
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                .ExecuteNonQuery()
                If .Parameters.Count > 0 Then
                    Ejecutar = CEntero(.Parameters(0).Value.ToString)
                End If
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                .Dispose()
                SqlComm = Nothing
                Ejecutar = CEntero(AnalizoError(ex.Message, ex.Source))
                Throw ex
            End Try
        End With
    End Function
    Public Function Ejecutar(ByVal Sql As String, ByVal Args As DataTable) As Integer
        Dim Posicion As Integer, ColData As Integer
        Dim IdNoVa As Integer = 0
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        If Sql.Substring(Sql.Trim.Length - 2, 2) = "_A" Then
            IdNoVa = 1
        End If
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                For Posicion = 0 To Args.Rows.Count - 1
                    For ColData = 0 + IdNoVa To Args.Columns.Count - 1
                        .Parameters(ColData - IdNoVa + 1).Value = Args.Rows(Posicion)(ColData)
                    Next
                    .ExecuteNonQuery()
                Next
                Ejecutar = CEntero(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
                Args.Dispose()
            Catch ex As Exception
                .Dispose()
                Args.Dispose()
                SqlComm = Nothing
                Ejecutar = CEntero(AnalizoError(ex.Message, ex.Source))
                Throw ex
            End Try
        End With
    End Function
    Public Function Ejecutar(ByVal Sql As String, ByVal Args As Integer) As Integer
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                .Parameters(1).Value = Args
                .ExecuteNonQuery()
                Ejecutar = CEntero(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                .Dispose()
                SqlComm = Nothing
                Ejecutar = CEntero(AnalizoError(ex.Message, ex.Source))
                Throw ex
            End Try
        End With
    End Function
    Public Function Ejecutar(ByVal Sql As String, ByVal Args As String) As Integer
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                .Parameters(1).Value = Args
                .ExecuteNonQuery()
                Ejecutar = CEntero(.Parameters(0).Value.ToString)
                .Dispose()
                SqlComm = Nothing
            Catch ex As Exception
                .Dispose()
                SqlComm = Nothing
                Ejecutar = CEntero(AnalizoError(ex.Message, ex.Source))
                Throw ex
            End Try
        End With
    End Function
#End Region
#Region "ArmarComando"
    Private Function ArmarComando(ByVal Procedimiento As String) As System.Data.SqlClient.SqlCommand
        'Dim ConexionAux As New System.Data.SqlClient.SqlConnection(CadenaConexion)
        'ConexionAux.Open()
        Dim NewConection As Boolean = True
        Dim ConectionAux As System.Data.SqlClient.SqlConnection
        If Not mConexion Is Nothing Then
            If mConexion.State = ConnectionState.Open And Not EnTransaccion Then
                ConectionAux = mConexion
                NewConection = False
            Else
                ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
                ConectionAux.Open()
            End If
        Else
            ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
            ConectionAux.Open()
        End If
        Dim mComando As New System.Data.SqlClient.SqlCommand(Procedimiento, ConectionAux)
        Dim mConstructor As New System.Data.SqlClient.SqlCommandBuilder
        mComando.CommandType = CommandType.StoredProcedure
        If ParameterCache.IsParametersCached(mComando) Then
            Dim param As System.Data.SqlClient.SqlParameter() = DirectCast(ParameterCache.GetCachedParameters(mComando), System.Data.SqlClient.SqlParameter())
            mComando.Parameters.AddRange(param)
        Else
            SqlClient.SqlCommandBuilder.DeriveParameters(mComando)
            ParameterCache.CacheParameters(mComando)
        End If
        With mComando
            .Connection = mConexion
            .Transaction = mTransaccion
        End With
        If NewConection Then
            With ConectionAux
                .Close()
                .Dispose()
            End With
            ConectionAux = Nothing
        End If
        mConstructor.Dispose()
        Return mComando
    End Function
#End Region
#Region "CargarParametros"
    Private Function CargarParametros(ByVal Comando As System.Data.SqlClient.SqlCommand, ByVal Args() As String) As System.Data.SqlClient.SqlCommand
        Dim Posicion As Integer
        Dim TipoParametro As SqlDbType
        With Comando
            For Posicion = 0 To Args.Length - 1
                TipoParametro = .Parameters(Posicion + 1).SqlDbType
                If Args(Posicion).Trim.ToUpper = "DBNULL" Then
                    .Parameters(Posicion + 1).Value = System.DBNull.Value
                Else
                    Select Case TipoParametro
                        Case SqlDbType.BigInt
                            .Parameters(Posicion + 1).Value = CEntero(Args(Posicion))
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
                            .Parameters(Posicion + 1).Value = CEntero(Args(Posicion))
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
                            .Parameters(Posicion + 1).Value = CEntero(Args(Posicion))
                        Case SqlDbType.SmallMoney
                            .Parameters(Posicion + 1).Value = CDoble(Args(Posicion))
                        Case SqlDbType.Text
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
        CargarParametros = Comando
    End Function
#End Region
#Region "TraerRegistros"
    Private Function TraerRegistros(ByVal Sql As String, ByVal Args() As String) As DataTable
        Dim Datatable As New DataTable
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                SqlComm = CargarParametros(SqlComm, Args)
                DataAdapter.Fill(Datatable)
                TraerRegistros = Datatable
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
                AnalizoError(ex.Message, ex.Source)
                Throw ex
            End Try
        End With
    End Function
    Private Function TraerRegistros(ByVal Sql As String, ByVal Args As DataTable) As DataTable
        Dim DataTable As New DataTable
        Dim lAffected As Integer
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                For lAffected = 0 To Args.Rows.Count() - 1
                    .Parameters(lAffected + 1).Value = Args.Rows(0)(lAffected)
                Next
                DataAdapter.Fill(DataTable)
                TraerRegistros = DataTable
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
                AnalizoError(ex.Message, ex.Source)
                Throw ex
            End Try
        End With
    End Function
    Private Function TraerRegistros(ByVal Sql As String) As DataTable
        Dim DataTable As New DataTable
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        Try
            DataAdapter.Fill(DataTable)
            TraerRegistros = DataTable
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
            AnalizoError(ex.Message, ex.Source)
            Throw ex
        End Try
    End Function
    Private Function TraerRegistros(ByVal Sql As String, ByVal Args As Integer) As DataTable
        Dim Datatable As New DataTable
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                .Parameters(1).Value = Args
                DataAdapter.Fill(Datatable)
                TraerRegistros = Datatable
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
            Catch ex As Exception
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
                AnalizoError(ex.Message, ex.Source)
                Throw ex
            End Try
        End With
    End Function
    Private Function TraerRegistros(ByVal Sql As String, ByVal Cadena As String) As DataTable
        Dim Datatable As New DataTable
        Conexion()
        Dim SqlComm As System.Data.SqlClient.SqlCommand = ArmarComando(Sql)
        Dim DataAdapter As New System.Data.SqlClient.SqlDataAdapter(SqlComm)
        With SqlComm
            If Me.VTimeOut.Trim <> "" Then .CommandTimeout = CEntero(VTimeOut)
            Try
                If Cadena.Trim.ToUpper = "DBNULL" Then
                    .Parameters(1).Value = System.DBNull.Value
                Else
                    .Parameters(1).Value = Cadena.Trim
                End If
                DataAdapter.Fill(Datatable)
                TraerRegistros = Datatable
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
            Catch ex As Exception
                DataAdapter.Dispose()
                .Dispose()
                SqlComm = Nothing
                DataAdapter = Nothing
                AnalizoError(ex.Message, ex.Source)
                Throw ex
            End Try
        End With
    End Function
#End Region
#Region "Errores"
    Private Function AnalizoError(ByVal TxtError As String, ByVal Origen As String) As String
        Dim Ruta As String = "C:\NetErr"
        Dim Archivo As String = Ruta & "\Errores.xml"
        If Not Directory.Exists(Ruta) Then
            Directory.CreateDirectory(Ruta)
        End If
        Dim MidataSet As New DataSet
        MidataSet.DataSetName = "MisErrores"
        If File.Exists(Archivo) Then
            MidataSet.ReadXml(Archivo)
        Else
            Dim MisRegistros As New DataTable
            With MisRegistros
                .Columns.Add("Sistema", System.Type.GetType("System.String"))
                .Columns.Add("Fecha", System.Type.GetType("System.String"))
                .Columns.Add("Hora", System.Type.GetType("System.String"))
                .Columns.Add("Modulo", System.Type.GetType("System.String"))
                .Columns.Add("Metodo", System.Type.GetType("System.String"))
                .Columns.Add("Origen", System.Type.GetType("System.String"))
                .Columns.Add("CodigoError", System.Type.GetType("System.Int32"))
                .Columns.Add("ErrorPuro", System.Type.GetType("System.String"))
                .Columns.Add("ErrorInterpretado", System.Type.GetType("System.String"))
                .Columns.Add("Usuario", System.Type.GetType("System.String"))
            End With
            MidataSet.Tables.Add(MisRegistros)
            MisRegistros.Dispose()
        End If
        Dim Args(9) As String
        Args(0) = VSistema
        Args(1) = System.DateTime.Now.Day.ToString.Trim & "/" & System.DateTime.Now.Month.ToString.Trim & "/" & System.DateTime.Now.Year.ToString.Trim
        Args(2) = System.DateTime.Now.Hour.ToString.Trim & ":" & System.DateTime.Now.Minute.ToString.Trim & ":" & System.DateTime.Now.Second.ToString.Trim
        Args(3) = VModulo
        Args(4) = VMetodo
        Args(5) = Origen
        Args(6) = CodigoDeError(TxtError.Trim, True)
        Args(7) = TxtError.Trim
        Args(8) = CodigoDeError(TxtError.Trim, False)
        Args(9) = Vusuario
        With MidataSet
            .Tables(0).Rows.Add(Args)
            .WriteXml(Archivo)
            .Dispose()
        End With
        AnalizoError = Args(6)
        Args = Nothing
    End Function
    Public Function CodigoDeError(ByVal ElError As String, ByVal Codigo As Boolean) As String
        Dim ConstanteDeError As String = "DELETE statement conflicted with COLUMN REFERENCE constraint"
        If ElError.IndexOf(ConstanteDeError) <> -1 Then
            If Codigo Then
                Return "-3"
            Else
                Return "No se puede borrar el registro porque tiene datos asociados"
            End If
        End If
        ConstanteDeError = "The DELETE statement conflicted with the REFERENCE constraint"
        If ElError.IndexOf(ConstanteDeError) <> -1 Then
            If Codigo Then
                Return "-3"
            Else
                Return "No se puede borrar el registro porque tiene datos asociados"
            End If
        End If
        ConstanteDeError = "Violation of UNIQUE KEY"
        If ElError.IndexOf(ConstanteDeError) <> -1 Then
            If Codigo Then
                Return "-1"
            Else
                Return "Esta intentando cargar un registro existente en la base"
            End If
        End If
        ConstanteDeError = "Cannot insert duplicate key row in object"
        If ElError.IndexOf(ConstanteDeError) <> -1 Then
            If Codigo Then
                Return "-1"
            Else
                Return "Esta intentando cargar un registro existente en la base"
            End If
        End If
        ConstanteDeError = "Infracción de la restricción UNIQUE KEY"
        If ElError.IndexOf(ConstanteDeError) <> -1 Then
            If Codigo Then
                Return "-1"
            Else
                Return "Esta intentando cargar un registro existente en la base"
            End If
        End If
        ConstanteDeError = "INSERT statement conflicted with COLUMN FOREIGN KEY constraint"
        If ElError.IndexOf(ConstanteDeError) <> -1 Then
            If Codigo Then
                Return "-6"
            Else
                Return "El registro no se pudo ingresar porque algun dato relacional es incorrecto"
            End If
        End If
        ConstanteDeError = "The stored procedure"
        If ConstanteDeError.Trim.Length <= ElError.Trim.Length Then
            If ElError.Substring(0, ConstanteDeError.Length).Trim = ConstanteDeError Then
                If ElError.Substring(ElError.Trim.Length - 14, 14) = "doesn't exist." Then
                    If Codigo Then
                        Return "-5"
                    Else
                        Return "Falta un procedimiento almacenado"
                    End If
                End If
            End If
        End If
        Select Case ElError
            Case Is = "SQL Server does not exist or access denied."
                If Codigo Then
                    Return "-2"
                Else
                    Return "No existe el Sql al cual se esta haciendo referencia"
                End If
            Case Is = "The ConnectionString property has not been initialized."
                If Codigo Then
                    Return "-4"
                Else
                    Return "La conexion al sql no fue inicializada"
                End If
            Case Else
                If Codigo Then
                    Return "-999999"
                Else
                    Return "Error desconocido"
                End If
        End Select
    End Function
#End Region
#Region "GetRecords"
    Public Function GetRecordsFromTxtCommand(ByVal Command As String) As DataTable
        Conexion()
        Dim DataTable As New DataTable
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
            AnalizoError(ex.Message, ex.Source)
            Throw ex
        End Try
    End Function
    Private Function CreateCommandFromTxt(ByVal command As String) As System.Data.SqlClient.SqlCommand
        'Dim ConectionAux As New System.Data.SqlClient.SqlConnection(Stringsql)
        'ConectionAux.Open()
        Dim NewConection As Boolean = True
        Dim ConectionAux As System.Data.SqlClient.SqlConnection
        If Not mConexion Is Nothing Then
            If mConexion.State = ConnectionState.Open And Not EnTransaccion Then
                ConectionAux = mConexion
                NewConection = False
            Else
                ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
                ConectionAux.Open()
            End If
        Else
            ConectionAux = New System.Data.SqlClient.SqlConnection(Stringsql)
            ConectionAux.Open()
        End If
        Dim mCommand As New System.Data.SqlClient.SqlCommand(command, ConectionAux)
        Dim mConstructor As New System.Data.SqlClient.SqlCommandBuilder
        mCommand.CommandType = CommandType.Text
        With mCommand
            .Connection = mConexion
            .Transaction = mTransaccion
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
    Public Function CheckConection() As String
        Try
            Dim AuxiConec As New System.Data.SqlClient.SqlConnection(Stringsql)
            AuxiConec.Open()
            AuxiConec.Close()
            AuxiConec.Dispose()
            AuxiConec = Nothing
            CheckConection = "OK"
        Catch ex As Exception
            CheckConection = ex.Message
        End Try
    End Function
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
    Private Function CEntero(ByVal Numero As String) As Integer
        If Numero.ToLower = "true" Or Numero.ToLower = "verdadero" Then
            CEntero = 1
            Exit Function
        End If
        If Numero.ToLower = "false" Or Numero.ToLower = "falso" Then
            CEntero = 0
            Exit Function
        End If
        Try
            CEntero = Integer.Parse(Numero)
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
    Private Function Izquierda(ByVal Cadena As String, ByVal Posiciones As Integer) As String
        If Posiciones > Cadena.Trim.Length Then Return Cadena
        Return Cadena.Trim.Substring(0, Posiciones)
    End Function
#End Region
#Region "Derecha"
    Private Function Derecha(ByVal Cadena As String, ByVal Posiciones As Integer) As String
        If Posiciones > Cadena.Trim.Length Then Return Cadena
        Return Cadena.Trim.Substring(Cadena.Trim.Length - Posiciones, Posiciones)
    End Function
#End Region
    Private Function CFecha(ByVal Fecha As String) As DateTime
        Try
            CFecha = DateTime.Parse(Fecha)
        Catch
            Return System.DateTime.Now
        End Try
    End Function
    Public Shared Function UltimoValor(ByVal Texto As String) As String
        UltimoValor = ""
        Texto = Texto.Trim
        If Texto.Length = 0 Then Return UltimoValor
        Dim Largo As Integer = Texto.Length
        Dim N As Integer
        For N = Largo - 1 To 0 Step -1
            If Texto.Substring(N, 1) = "." Then
                UltimoValor = Texto.Substring(N + 1, Largo - 1 - N)
                Exit For
            End If
        Next
        If UltimoValor = "" Then UltimoValor = Texto
    End Function
#End Region
#Region "Transacciones"
    Public Sub IniciarTransaccion()
        Conexion()
        mTransaccion = mConexion.BeginTransaction
        EnTransaccion = True
    End Sub
    Public Sub TerminarTransaccion()
        Try
            With mTransaccion
                .Commit()
                .Dispose()
            End With
            EnTransaccion = False
            mTransaccion = Nothing
        Catch ex As System.Exception
            mTransaccion.Connection.Close()
            EnTransaccion = False
            mTransaccion = Nothing
            Throw ex
        End Try
    End Sub
    Public Sub AbortarTransaccion()
        Try
            With mTransaccion
                .Rollback()
                .Dispose()
            End With
            mTransaccion = Nothing
            EnTransaccion = False
        Catch Ex As System.Exception
            mTransaccion.Connection.Close()
            mTransaccion = Nothing
            EnTransaccion = False
            Throw Ex
        End Try
    End Sub
#End Region
End Class



