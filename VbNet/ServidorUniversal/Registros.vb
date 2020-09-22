Imports ServidorUniversal.Functions
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Data
Public Class Registros
    Public Function AgrupacionDeRegistros(ByVal Datos As DataTable, ByVal CampoAgrupacion As String, ByVal CampoAgrupacion2 As String, ByVal CampoAgrupacion3 As String, ByVal CampoCorte As String, ByVal OrdenDeUsuario As String) As DataTable
        AgrupacionDeRegistros = Datos.Clone
        If OrdenDeUsuario.Trim.Length = 0 Then
            Datos = OrdenarRegistros(Datos, CampoCorte)
        Else
            Datos = OrdenarRegistros(Datos, OrdenDeUsuario)
        End If
        Dim Posicion As Integer
        Dim Posicion2 As Integer
        Dim Posicion3 As Integer
        Dim n As Integer
        Dim BreakFieldIsChar As Boolean = False
        With Datos
            For n = 0 To .Columns.Count - 1
                If .Columns(n).ColumnName.Trim.ToLower = CampoCorte.Trim.ToLower Then
                    If .Columns(n).DataType Is System.Type.GetType("System.String") Then BreakFieldIsChar = True
                End If
                If .Columns(n).ColumnName.Trim.ToLower = CampoAgrupacion.Trim.ToLower Then Posicion = n
                If .Columns(n).ColumnName.Trim.ToLower = CampoAgrupacion2.Trim.ToLower Then Posicion2 = n
                If .Columns(n).ColumnName.Trim.ToLower = CampoAgrupacion3.Trim.ToLower Then Posicion3 = n
            Next
            Dim ArgsClone() As Object
            For n = 0 To .Rows.Count - 1
                If BreakFieldIsChar Then
                    If Toint(AgrupacionDeRegistros.Compute("count(" & CampoCorte & ")", CampoCorte & "='" & Datos.Rows(n)(CampoCorte).ToString.Replace("'", "''") & "'").ToString) = 0 Then
                        ArgsClone = .Rows(n).ItemArray()
                        ArgsClone(Posicion) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion & ")", CampoCorte & "='" & Datos.Rows(n)(CampoCorte).ToString.Replace("'", "''") & "'").ToString)
                        If CampoAgrupacion2.Trim.Length <> 0 Then
                            ArgsClone(Posicion2) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion2 & ")", CampoCorte & "='" & Datos.Rows(n)(CampoCorte).ToString.Replace("'", "''") & "'").ToString)
                        End If
                        If CampoAgrupacion3.Trim.Length <> 0 Then
                            ArgsClone(Posicion3) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion3 & ")", CampoCorte & "='" & Datos.Rows(n)(CampoCorte).ToString.Replace("'", "''") & "'").ToString)
                        End If
                        AgrupacionDeRegistros.Rows.Add(ArgsClone)
                    End If
                Else
                    If Toint(AgrupacionDeRegistros.Compute("count(" & CampoCorte & ")", CampoCorte & "=" & Datos.Rows(n)(CampoCorte).ToString).ToString) = 0 Then
                        ArgsClone = .Rows(n).ItemArray()
                        ArgsClone(Posicion) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion & ")", CampoCorte & "=" & Datos.Rows(n)(CampoCorte).ToString).ToString)
                        If CampoAgrupacion2.Trim.Length <> 0 Then
                            ArgsClone(Posicion2) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion2 & ")", CampoCorte & "=" & Datos.Rows(n)(CampoCorte).ToString).ToString)
                        End If
                        If CampoAgrupacion3.Trim.Length <> 0 Then
                            ArgsClone(Posicion3) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion3 & ")", CampoCorte & "=" & Datos.Rows(n)(CampoCorte).ToString).ToString)
                        End If
                        AgrupacionDeRegistros.Rows.Add(ArgsClone)
                    End If
                End If
            Next
            ArgsClone = Nothing
        End With
    End Function
    Public Function AgrupacionDeRegistros(ByVal Datos As DataTable, ByVal CampoAgrupacion As String, ByVal CampoCorte As String, ByVal OrdenDeUsuario As String) As DataTable
        AgrupacionDeRegistros = Datos.Clone
        If OrdenDeUsuario.Trim.Length = 0 Then
            Datos = OrdenarRegistros(Datos, CampoCorte)
        Else
            Datos = OrdenarRegistros(Datos, OrdenDeUsuario)
        End If
        Dim Posicion As Integer
        Dim n As Integer
        Dim BreakFieldIsChar As Boolean = False
        With Datos
            For n = 0 To .Columns.Count - 1
                If .Columns(n).ColumnName.Trim.ToLower = CampoCorte.Trim.ToLower Then
                    If .Columns(n).DataType Is System.Type.GetType("System.String") Then BreakFieldIsChar = True
                End If
                If .Columns(n).ColumnName.Trim.ToLower = CampoAgrupacion.Trim.ToLower Then Posicion = n
            Next
            Dim ArgsClone() As Object
            For n = 0 To .Rows.Count - 1
                If BreakFieldIsChar Then
                    If Toint(AgrupacionDeRegistros.Compute("count(" & CampoCorte & ")", CampoCorte & "='" & Datos.Rows(n)(CampoCorte).ToString.Replace("'", "''") & "'").ToString) = 0 Then
                        ArgsClone = .Rows(n).ItemArray()
                        ArgsClone(Posicion) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion & ")", CampoCorte & "='" & Datos.Rows(n)(CampoCorte).ToString.Replace("'", "''") & "'").ToString)
                        AgrupacionDeRegistros.Rows.Add(ArgsClone)
                    End If
                Else
                    If Toint(AgrupacionDeRegistros.Compute("count(" & CampoCorte & ")", CampoCorte & "=" & Datos.Rows(n)(CampoCorte).ToString).ToString) = 0 Then
                        ArgsClone = .Rows(n).ItemArray()
                        ArgsClone(Posicion) = ToDouble(Datos.Compute("sum(" & CampoAgrupacion & ")", CampoCorte & "=" & Datos.Rows(n)(CampoCorte).ToString).ToString)
                        AgrupacionDeRegistros.Rows.Add(ArgsClone)
                    End If
                End If

            Next
            ArgsClone = Nothing
        End With
    End Function
    Public Function OrdenarRegistros(ByVal Datos As DataTable, ByVal Orden As String) As DataTable
        Dim n As Integer
        Dim x As Integer
        Dim Vista As New DataView(Datos)
        Dim WorkTable As DataTable = Vista.Table.Clone
        Dim Registros As Integer = Vista.Count
        Dim Columnas As Integer = Vista.Table.Columns.Count
        Dim Args(Columnas - 1) As Object
        Vista.Sort = Orden
        For n = 0 To Registros - 1
            For x = 0 To Columnas - 1
                Args(x) = Vista(n)(x)
            Next
            WorkTable.Rows.Add(Args)
        Next
        Vista.Dispose()
        Return WorkTable
    End Function
    Public Function FiltrarRegistros(ByVal Datos As DataTable, ByVal Filtro As String) As DataTable
        Dim n As Integer
        Dim x As Integer
        Dim Vista As New DataView(Datos)
        Dim WorkTable As DataTable = Vista.Table.Clone
        Vista.RowFilter = Filtro
        Dim Registros As Integer = Vista.Count
        Dim Columnas As Integer = Vista.Table.Columns.Count
        Dim Args(Columnas - 1) As Object
        For n = 0 To Registros - 1
            For x = 0 To Columnas - 1
                Args(x) = Vista(n)(x)
            Next
            WorkTable.Rows.Add(Args)
        Next
        Vista.Dispose()
        Return WorkTable
    End Function
    Public Function ExportarExcelXml(ByVal Titulo As String, ByVal Tabla As String, ByVal Request As System.Web.HttpRequest) As String
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim MiArchivo As String
        Dim MiFile As StreamWriter
        Dim Cadena As New System.Text.StringBuilder
        RutaExportacion = Request.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("xls")
        MiArchivo = RutaExportacion & MiArchivoLimpio
        With Cadena
            .Append("<html xmlns:x=")
            .Append(Convert.ToChar(34))
            .Append(Convert.ToChar(34))
            .Append("urn:schemas-microsoft-com:office:excel")
            .Append(Convert.ToChar(34))
            .Append(Convert.ToChar(34))
            .Append(">")
            .Append("<head>")
            .Append("<x:ExcelWorkbook>")
            .Append("<x:ExcelWorksheets>")
            .Append("<x:ExcelWorksheets>")
            .Append("<x:Name>")
            .Append(Titulo)
            .Append("</x:Name>")
            .Append("<x:WorksheetOptions>")
            .Append("<x:ValidPrinterInfo/>")
            .Append("</x:Print>")
            .Append("</x:WorksheetOptions>")
            .Append("</x:ExcelWorksheet>")
            .Append("</x:ExcelWorksheets>")
            .Append("</x:ExcelWorkbook>")
            .Append("</xml>")
            .Append("</head>")
            .Append("<body>")
            .Append(Tabla)
            .Append("</body>")
            .Append("</html>")
        End With
        MiFile = New StreamWriter(MiArchivo)
        MiFile.Write(Cadena.ToString)
        MiFile.Close()
        Return MiArchivoLimpio
    End Function
    Public Function ExportarXml(ByVal Datos As DataTable, ByVal SacaCaracteresEspeciales As Boolean, ByVal Request As System.Web.HttpRequest, ByVal Ndataset As String, ByVal NTabla As String) As String
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim MiArchivo As String
        RutaExportacion = Request.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("xml")
        MiArchivo = RutaExportacion & MiArchivoLimpio
        Dim MiDataSet As New DataSet
        MiDataSet.DataSetName = Ndataset
        Dim Odtaux As DataTable = New DataTable
        Odtaux = Datos.Copy
        Odtaux.TableName = NTabla
        MiDataSet.Tables.Add(Odtaux)
        Try
            MiDataSet.WriteXml(MiArchivo)
        Catch ex As Exception
            MiArchivoLimpio = ""
        End Try
        MiDataSet.Dispose()
        Datos.Dispose()
        Odtaux.Dispose()
        If SacaCaracteresEspeciales Then
            'FALTA DESARROLLAR
        End If
        Return MiArchivoLimpio
    End Function
    Public Function ExportarXml(ByVal Datos As DataSet, ByVal SacaCaracteresEspeciales As Boolean, ByVal Request As System.Web.HttpRequest, ByVal Ndataset As String, ByVal NTabla As String) As String
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim MiArchivo As String
        RutaExportacion = Request.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("xml")
        MiArchivo = RutaExportacion & MiArchivoLimpio
        Try
            Datos.WriteXml(MiArchivo)
        Catch ex As Exception
            MiArchivoLimpio = ""
        End Try
        Datos.Dispose()
        If SacaCaracteresEspeciales Then
            'FALTA DESARROLLAR
        End If
        Return MiArchivoLimpio
    End Function
    Public Function ExportarTemporalXml(ByVal Datos As DataTable) As String
        ExportarTemporalXml = ""
        Dim MiDataSet As New DataSet
        MiDataSet.Tables.Add(Datos)
        Try
            MiDataSet.WriteXml("C:\Inetpub\wwwroot\Antares\Exportaciones\datos.xml")
        Catch ex As Exception
        End Try
        MiDataSet.Dispose()
        Datos.Dispose()
    End Function
    Public Function Merge(ByVal MyData As DataTable, ByVal MyData2 As DataTable, ByVal Sort As String) As DataTable
        Dim n As Integer
        With MyData2
            Dim ArgsObj() As Object
            For n = 0 To .Rows.Count - 1
                ArgsObj = .Rows(n).ItemArray
                MyData.Rows.Add(ArgsObj)
            Next
            If Sort.Trim.Length <> 0 Then MyData = Me.OrdenarRegistros(MyData, Sort)
            ArgsObj = Nothing
        End With
        Return MyData
    End Function
    Public Function Distinct(ByVal Ors As DataTable, ByVal Field As String, ByVal NumericData As Boolean) As Data.DataTable
        Distinct = Ors.Clone
        Dim n As Integer
        For n = 0 To Ors.Rows.Count - 1
            If NumericData Then
                If Me.FiltrarRegistros(Distinct, Field & "=" & Ors.Rows(n)(Field).ToString).Rows.Count = 0 Then
                    Distinct.Rows.Add(Ors.Rows(n).ItemArray)
                End If
            Else
                If Me.FiltrarRegistros(Distinct, Field & "='" & Ors.Rows(n)(Field).ToString & "'").Rows.Count = 0 Then
                    Distinct.Rows.Add(Ors.Rows(n).ItemArray)
                End If
            End If
        Next
    End Function
    Public Sub DataTableToXls(ByVal dt As DataTable, ByVal FileName As String)
        Const FieldSeparator As String = vbTab
        Const RowSeparator As String = vbLf
        Dim output As New Text.StringBuilder
        For Each dc As DataColumn In dt.Columns
            output.Append(dc.ColumnName)
            output.Append(FieldSeparator)
        Next
        output.Append(RowSeparator)
        For Each item As DataRow In dt.Rows
            For Each value As Object In item.ItemArray
                output.Append(value.ToString())
                output.Append(FieldSeparator)
            Next
            output.Append(RowSeparator)
        Next
        Dim sw As New StreamWriter(FileName)
        sw.Write(output.ToString())
        sw.Close()
    End Sub
End Class

