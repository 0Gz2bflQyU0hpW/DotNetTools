Option Explicit On 
Option Strict On
Imports System.Data.OleDb
Imports ServidorUniversal.Functions
Public Class Access
#Region "Variables"
    Private VArchivoConRuta As String
#End Region
#Region "Propiedades"
    Public ReadOnly Property ArchivoConectado() As String
        Get
            Return VArchivoConRuta
        End Get
    End Property
#End Region
#Region "Constructores de la clase"
    Sub New(ByVal ArchivoConRuta As String)
        VArchivoConRuta = ArchivoConRuta
    End Sub
#End Region
#Region "Metodos"
#Region "ConexionOLEDB"
    Private Function ConexionOLEDB() As String
        Dim Cadena As New System.Text.StringBuilder
        With Cadena
            .Append("Provider=Microsoft.Jet.OLEDB.4.0; ")
            .Append("Data Source=")
            .Append(VArchivoConRuta & ";")
            .Append("User ID=Admin;")
            .Append("Password=")
        End With
        Return Cadena.ToString
    End Function
#End Region
#Region "TraerRegistrosDesdeComandoSql"
    Public Function TraerRegistrosDesdeComandoSql(ByVal Comando As String) As DataTable
        TraerRegistrosDesdeComandoSql = New DataTable
        Try
            Dim StringConexion As String = ConexionOLEDB()
            Dim Conexion As New OleDbConnection(StringConexion)
            Conexion.Open()
            Dim Lector As New OleDbDataAdapter(Comando, Conexion)
            Lector.Fill(TraerRegistrosDesdeComandoSql)
            Lector.Dispose()
            Conexion.Close()
            Conexion.Dispose()
        Catch ex As Exception
            Throw New SystemException(ex.Message)
        End Try
    End Function
#End Region
#Region "EjecutarDesdeComandoSql"
    Public Sub EjecutarDesdeComandoSql(ByVal comandoSql As String)
        Dim StringConexion As String
        StringConexion = ConexionOLEDB()
        Dim conexion As OleDbConnection
        conexion = New OleDbConnection(StringConexion)
        conexion.Open()
        Dim Comando As New OleDbCommand(comandoSql, conexion)
        Try
            Comando.ExecuteNonQuery()
            Comando.Dispose()
        Catch ex As Exception
            Comando.Dispose()
            conexion.Dispose()
            Throw New SystemException(ex.Message)
        End Try
        conexion.Dispose()
    End Sub
#End Region
#Region "TraerTodos"
    Public Function TraerTodos(ByVal Tabla As String) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TT"
        TraerTodos = TraerRegistros(Sql)
    End Function
#End Region
#Region "Agregar"
    Public Function Agregar(ByVal Tabla As String, ByVal Valores(,) As String) As Integer
        Dim Sql As String
        Sql = Tabla & "_A"
        Agregar = Ejecutar(Sql, Valores)
    End Function
#End Region
#Region "Modificar"
    Public Function Modificar(ByVal Tabla As String, ByVal Valores(,) As String) As Integer
        Dim Sql As String
        Sql = Tabla & "_M"
        Modificar = Ejecutar(Sql, Valores)
    End Function
#End Region
#Region "TraerConFiltro"
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Args(,) As String) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql, Args)
    End Function
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql)
    End Function
    Public Function TraerConFiltro(ByVal Tabla As String, ByVal Filtro As String, ByVal Id As Integer) As DataTable
        Dim Sql As String
        Sql = Tabla & "_TX_" & Filtro
        TraerConFiltro = TraerRegistros(Sql, Id)
    End Function
#End Region
#Region "BorrarTodo"
    Public Function BorrarTodo(ByVal Tabla As String) As Integer
        Dim Sql As String
        Sql = Tabla & "_DXA"
        BorrarTodo = Ejecutar(Sql)
    End Function
#End Region
#Region "Ejecutar"
    Private Function Ejecutar(ByVal Sql As String, ByVal Args(,) As String) As Integer
        Try
            Dim StringConexion As String = ConexionOLEDB()
            Dim Conexion As New OleDbConnection(StringConexion)
            Conexion.Open()
            Dim Comando As New OleDbCommand(Sql, Conexion)
            With Comando
                Dim n As Integer
                .CommandType = CommandType.StoredProcedure
                For n = 0 To Args.GetLength(0) - 1
                    .Parameters.Add(DameParametro(Args(n, 0), Args(n, 1)))
                Next
            End With
            Comando.ExecuteNonQuery()
            Comando.Dispose()
            Conexion.Close()
            Conexion.Dispose()
            Comando.Dispose()
            Args = Nothing
        Catch ex As Exception
            Throw New SystemException(ex.Message)
        End Try
    End Function
    Private Function Ejecutar(ByVal Sql As String) As Integer
        Try
            Dim StringConexion As String = ConexionOLEDB()
            Dim Conexion As New OleDbConnection(StringConexion)
            Conexion.Open()
            Dim Comando As New OleDbCommand(Sql, Conexion)
            With Comando
                .CommandType = CommandType.StoredProcedure
            End With
            Comando.ExecuteNonQuery()
            Comando.Dispose()
            Conexion.Close()
            Conexion.Dispose()
            Comando.Dispose()
        Catch ex As Exception
            Throw New SystemException(ex.Message)
        End Try
    End Function
#End Region
#Region "TraerRegistros"
    Private Function TraerRegistros(ByVal Sql As String) As DataTable
        TraerRegistros = New DataTable
        Try
            Dim StringConexion As String = ConexionOLEDB()
            Dim Conexion As New OleDbConnection(StringConexion)
            Conexion.Open()
            Dim Comando As New OleDbCommand(Sql, Conexion)
            Comando.CommandType = CommandType.StoredProcedure
            Dim Lector As New OleDbDataAdapter(Comando)
            Lector.Fill(TraerRegistros)
            Comando.Dispose()
            Lector.Dispose()
            Conexion.Close()
            Conexion.Dispose()
            Comando.Dispose()
        Catch ex As Exception
            Throw New SystemException(ex.Message)
        End Try
    End Function
    Private Function TraerRegistros(ByVal Sql As String, ByVal Args(,) As String) As DataTable
        TraerRegistros = New DataTable
        Try
            Dim StringConexion As String = ConexionOLEDB()
            Dim Conexion As New OleDbConnection(StringConexion)
            Conexion.Open()
            Dim Comando As New OleDbCommand(Sql, Conexion)
            With Comando
                Dim n As Integer
                .CommandType = CommandType.StoredProcedure
                For n = 0 To Args.GetLength(0) - 1
                    .Parameters.Add(DameParametro(Args(n, 0), Args(n, 1)))
                Next
            End With
            Dim Lector As New OleDbDataAdapter(Comando)
            Lector.Fill(TraerRegistros)
            Comando.Dispose()
            Lector.Dispose()
            Conexion.Close()
            Conexion.Dispose()
            Comando.Dispose()
            Args = Nothing
        Catch ex As Exception
            Throw New SystemException(ex.Message)
        End Try
    End Function
    Private Function TraerRegistros(ByVal Sql As String, ByVal Id As Integer) As DataTable
        TraerRegistros = New DataTable
        Try
            Dim StringConexion As String = ConexionOLEDB()
            Dim Conexion As New OleDbConnection(StringConexion)
            Conexion.Open()
            Dim Comando As New OleDbCommand(Sql, Conexion)
            With Comando
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add(DameParametro(Id.ToString, "int"))
            End With
            Dim Lector As New OleDbDataAdapter(Comando)
            Lector.Fill(TraerRegistros)
            Comando.Dispose()
            Lector.Dispose()
            Conexion.Close()
            Conexion.Dispose()
            Comando.Dispose()
        Catch ex As Exception
            Throw New SystemException(ex.Message)
        End Try
    End Function
#End Region
#Region "DameParametro"
    Private Function DameParametro(ByVal Valor As String, ByVal Tipo As String) As OleDb.OleDbParameter
        DameParametro = New OleDb.OleDbParameter
        DameParametro.Value = Valor
        Select Case Tipo.Trim.ToLower
            Case "int"
                DameParametro.OleDbType = OleDb.OleDbType.Integer
            Case "chr"
                DameParametro.OleDbType = OleDb.OleDbType.Char
        End Select
    End Function
#End Region
#End Region
End Class
