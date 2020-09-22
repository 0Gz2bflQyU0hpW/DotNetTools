Option Explicit On 
Option Strict On
Imports System.Data.OleDb
Public Class Xls
#Region "Conexion"
    Private Function ConexionOLEDB() As String
        Dim Cadena As New System.Text.StringBuilder
        With Cadena
            .Append("Provider=Microsoft.Jet.OLEDB.4.0; ")
            .Append("Extended Properties=Excel 8.0; ")
            .Append(" Data Source=")
            .Append(Varchivo)
        End With
        Return Cadena.ToString
    End Function
#End Region
#Region "Variables"
    Private Varchivo As String
#End Region
#Region "Constructores de la clase"
    Sub New(ByVal Archivo As String)
        If System.IO.File.Exists(Archivo) Then
            Varchivo = Archivo
        Else
            Throw New System.Exception("File " & Archivo & " not found")
        End If
    End Sub
#End Region
#Region "Metodos"
    Public Function TraerRegistros(ByVal Hoja As String, ByVal Filtro As String, ByVal Seccion As String) As DataTable
        Dim StringConexion As String
        StringConexion = ConexionOLEDB()
        Dim conexion As OleDbConnection = New OleDbConnection(StringConexion)
        Try
            conexion.Open()
            Dim Ors As New DataTable
            Dim Cadena As New System.Text.StringBuilder
            If Hoja.Trim.Length <> 0 Then
                With Cadena
                    .Append("select * from [")
                    .Append(Hoja)
                    .Append("$]")
                    If Filtro.Trim.Length <> 0 Then
                        .Append(" where ")
                        .Append(Filtro)
                    End If
                End With
            Else
                With Cadena
                    .Append("select * from ")
                    .Append(Seccion)
                    If Filtro.Trim.Length <> 0 Then
                        .Append(" where ")
                        .Append(Filtro)
                    End If
                End With
            End If
            Dim Lector As New OleDbDataAdapter(Cadena.ToString, conexion)
            Lector.Fill(Ors)
            Lector.Dispose()
            Lector = Nothing
            TraerRegistros = Ors
            Ors.Dispose()
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        Finally
            conexion.Close()
            conexion.Dispose()
            conexion = Nothing
        End Try
    End Function
    Public Function TraerRegistros(ByVal Hoja As String) As DataTable
        Dim Ors As New DataTable
        Dim StringConexion As String
        Dim Cadena As New System.Text.StringBuilder
        With Cadena
            .Append("select * from [")
            .Append(Hoja)
            .Append("$]")
        End With
        StringConexion = ConexionOLEDB()
        Dim conexion As OleDbConnection = New OleDbConnection(StringConexion)
        conexion.Open()
        Dim Lector As New OleDbDataAdapter(Cadena.ToString, conexion)
        Lector.Fill(Ors)
        Lector.Dispose()
        conexion.Close()
        conexion.Dispose()
        Return Ors
    End Function
    Public Function ExisteHoja(ByVal Hoja As String) As Boolean
        Dim StringConexion As String
        Dim dtXlsSchema As DataTable
        Dim i As Integer
        StringConexion = ConexionOLEDB()
        Dim conexion As OleDbConnection = New OleDbConnection(StringConexion)
        conexion.Open()
        dtXlsSchema = conexion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        ExisteHoja = False
        For i = 0 To dtXlsSchema.Rows.Count - 1
            If dtXlsSchema.Rows(i).Item("Table_Name").ToString.ToLower = Hoja.ToLower & "$" Then
                ExisteHoja = True
            End If
        Next
        conexion.Close()
        conexion.Dispose()
        dtXlsSchema.Dispose()
        conexion = Nothing
        dtXlsSchema = Nothing
    End Function
    Public Function PrimeraHoja() As String
        Dim StringConexion As String
        Dim dtXlsSchema As DataTable
        StringConexion = ConexionOLEDB()
        Dim conexion As OleDbConnection = New OleDbConnection(StringConexion)
        conexion.Open()
        dtXlsSchema = conexion.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
        PrimeraHoja = dtXlsSchema.Rows(0).Item("Table_Name").ToString.ToLower
        If PrimeraHoja.EndsWith("$") Then
            PrimeraHoja = PrimeraHoja.Substring(0, PrimeraHoja.Length - 1)
        End If
        conexion.Close()
        conexion.Dispose()
        dtXlsSchema.Dispose()
        conexion = Nothing
        dtXlsSchema = Nothing
    End Function
    Public Function EjecutarOrdenes(ByVal Hoja As String, ByVal Comando As String) As String
        EjecutarOrdenes = "OK"
        Dim StringConexion As String
        StringConexion = ConexionOLEDB()
        Dim conexion As OleDbConnection = New OleDbConnection(StringConexion)
        conexion.Open()
        Dim OComando As OleDbCommand
        Try
            OComando = New OleDbCommand(Comando, conexion)
            OComando.ExecuteNonQuery()
            OComando.Dispose()
        Catch ex As Exception
            EjecutarOrdenes = ex.Message
        Finally
            conexion.Dispose()
        End Try
    End Function
#End Region
End Class
