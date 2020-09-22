Option Explicit On 
Option Strict On
Imports System.Data.OleDb
Public Class Dbf
#Region "Variables"
    Private Vruta As String
    Public Enum TipoDbf As Integer
        DbaseIII = 1
        FoxPro = 2
    End Enum
#End Region
#Region "Propiedades"
    Public Property Ruta() As String
        Get
            Return Vruta
        End Get
        Set(ByVal Value As String)
            Vruta = Value
        End Set
    End Property
#End Region
#Region "Constructores de la clase"
    Sub New(ByVal Ruta As String)
        Vruta = Ruta
    End Sub
#End Region
#Region "Metodos"
    Public Function TraerRegistros(ByVal DbfType As TipoDbf, ByVal Archivo As String, ByVal Filtro As String, ByVal TopNumber As Integer, ByVal Orden As String, ByVal Campos As String) As DataTable
        Dim Ors As New DataTable
        Dim StringConexion As String
        Dim Cadena As New System.Text.StringBuilder
        If Campos.Trim.Length = 0 Then Campos = "*"
        With Cadena
            If TopNumber = 0 Then
                .Append("select ")
                .Append(Campos)
                .Append(" From ")
            Else
                .Append("select ")
                .Append(Campos & ",0 as CampoEliminacion")
                .Append(" From ")
            End If
            .Append(Archivo)
            .Append(".dbf")
            If Filtro.Trim.Length <> 0 Then
                .Append(" where ")
                .Append(Filtro)
            End If
            If Orden.Trim.Length <> 0 Then
                .Append(" order by ")
                .Append(Orden)
            End If
        End With
        Try
            Select Case DbfType
                Case TipoDbf.DbaseIII
                    StringConexion = ConexionOLEDB()
                    Dim conexion As OleDbConnection
                    conexion = New OleDbConnection(StringConexion)
                    conexion.Open()
                    Dim Lector As New OleDbDataAdapter(Cadena.ToString, conexion)
                    Lector.Fill(Ors)
                    Lector.Dispose()
                    conexion.Close()
                    conexion.Dispose()
                Case TipoDbf.FoxPro
                    StringConexion = ConexionODBC()
                    Dim conexion As Odbc.OdbcConnection
                    conexion = New Odbc.OdbcConnection(StringConexion)
                    conexion.Open()
                    Dim Lector As New Odbc.OdbcDataAdapter(Cadena.ToString, conexion)
                    Lector.Fill(Ors)
                    Lector.Dispose()
                    conexion.Close()
                    conexion.Dispose()
            End Select
        Catch ex As Exception
            Ors = New DataTable
        End Try
        If TopNumber <> 0 Then
            Dim n As Integer
            If TopNumber > Ors.Rows.Count Then TopNumber = Ors.Rows.Count
            For n = 0 To TopNumber
                Ors.Rows(n)("CampoEliminacion") = "1"
            Next
            Dim Otool As New Registros
            Ors = Otool.FiltrarRegistros(Ors, "CampoEliminacion=1")
        End If
        Return Ors
    End Function
    Public Function TraerRegistros(ByVal DbfType As TipoDbf, ByVal Comando As String) As DataTable
        Dim Ors As New DataTable
        Dim StringConexion As String
        Select Case DbfType
            Case TipoDbf.DbaseIII
                StringConexion = ConexionOLEDB()
                Dim conexion As OleDbConnection
                conexion = New OleDbConnection(StringConexion)
                conexion.Open()
                Dim Lector As New OleDbDataAdapter(Comando.ToString, conexion)
                Lector.Fill(Ors)
                Lector.Dispose()
                conexion.Close()
                conexion.Dispose()
            Case TipoDbf.FoxPro
                Try
                    StringConexion = ConexionODBC()
                    Dim conexion As New Odbc.OdbcConnection(StringConexion)
                    conexion.Open()
                    Dim Lector As New Odbc.OdbcDataAdapter(Comando.ToString, conexion)
                    Lector.Fill(Ors)
                    Lector.Dispose()
                    conexion.Close()
                    conexion.Dispose()
                Catch ex As Exception

                End Try
        End Select
        Return Ors
    End Function
    Public Function BorrarRegistros(ByVal DbfType As TipoDbf, ByVal Archivo As String, ByVal Filtro As String) As String
        Dim StringConexion As String
        Dim Cadena As New System.Text.StringBuilder
        With Cadena
            .Append("delete from ")
            .Append(Archivo)
            .Append(".dbf")
            If Filtro.Trim.Length <> 0 Then
                .Append(" where ")
                .Append(Filtro)
            End If
        End With
        Select Case DbfType
            Case TipoDbf.DbaseIII
                StringConexion = ConexionOLEDB()
                Dim conexion As OleDbConnection
                conexion = New OleDbConnection(StringConexion)
                conexion.Open()
                Dim Comando As New OleDbCommand(Cadena.ToString, conexion)
                Try
                    Comando.ExecuteNonQuery()
                    Comando.Dispose()
                Catch ex As Exception
                    Comando.Dispose()
                End Try
                conexion.Dispose()
        End Select
        BorrarRegistros = "OK"
    End Function
    Public Function Ejecutarordenes(ByVal DbfType As TipoDbf, ByVal ComandoSql As String) As Integer
        Dim StringConexion As String
        Select Case DbfType
            Case TipoDbf.DbaseIII
                StringConexion = ConexionOLEDB()
                Dim conexion As OleDbConnection
                conexion = New OleDbConnection(StringConexion)
                conexion.Open()
                Dim Comando As OleDbCommand
                Try
                    Comando = New OleDbCommand(ComandoSql, conexion)
                    Comando.ExecuteNonQuery()
                    Comando.Dispose()
                Catch ex As Exception
                End Try
                conexion.Dispose()
            Case TipoDbf.FoxPro
                StringConexion = ConexionODBC()
                Dim conexion As Odbc.OdbcConnection
                conexion = New Odbc.OdbcConnection(StringConexion)
                conexion.Open()
                Dim Comando As Odbc.OdbcCommand
                Try
                    Comando = New Odbc.OdbcCommand(ComandoSql, conexion)
                    Comando.ExecuteNonQuery()
                    Comando.Dispose()
                Catch ex As Exception
                End Try
                conexion.Dispose()
        End Select
    End Function
#End Region
#Region "Conexion"
    Private Function ConexionODBC() As String
        Dim Cadena As New System.Text.StringBuilder
        With Cadena
            .Append("Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDb=")
            .Append(Vruta)
            .Append(";Exclusive=No")
        End With
        Return Cadena.ToString
    End Function
    Private Function ConexionOLEDB() As String
        Dim Cadena As New System.Text.StringBuilder
        With Cadena
            .Append("Provider=Microsoft.Jet.OLEDB.4.0; ")
            .Append(" Data Source=")
            .Append(Vruta)
            .Append(";Extended Properties=dBASE IV; ")

        End With
        Return Cadena.ToString
    End Function
#End Region
End Class
