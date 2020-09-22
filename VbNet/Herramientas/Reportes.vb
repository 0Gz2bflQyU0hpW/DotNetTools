Option Explicit On 
Option Strict On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports LibreriaNet.Funciones
Imports Microsoft.VisualBasic
Imports System.IO
Public Class Reportes
    Inherits ReportDocument
    Private NombreReporte_p As String
    Private Request_p As System.Web.HttpRequest
    Private MiArchivo As String
    Public CC As String = ""
    Public Sub New(ByVal Nombre As String, ByVal Request As System.Web.HttpRequest)
        Dim Ilen As Integer
        Request_p = Request
        NombreReporte_p = Request_p.ServerVariables("PATH_TRANSLATED")
        While (Derecha(NombreReporte_p, 1) <> "\" And NombreReporte_p.Length <> 0)
            Ilen = NombreReporte_p.Length - 1
            NombreReporte_p = NombreReporte_p.Substring(0, Ilen)
        End While
        NombreReporte_p = NombreReporte_p & Nombre
    End Sub
    Public Sub AgregarParametros(ByVal ParamName As String, ByVal Value As String)
        If Value.Trim.Length >= 254 Then Value = Value.Substring(0, 254)
        'ESTO ES PORQUE SI PASAS UN PARAMETRO DE TEXTO MAYOR A 254 CARACTERES TIRA ERROR
        Dim MiParametro As New ParameterDiscreteValue
        Dim Valores As New ParameterValues
        MiParametro.Value = Value
        With Valores
            .Clear()
            .Add(MiParametro)
        End With
        Me.DataDefinition.ParameterFields(ParamName).ApplyCurrentValues(Valores)
    End Sub
    Public Overloads Sub Load(ByVal Source As DataTable)
        Me.Load(NombreReporte_p)
        Me.SetDataSource(Source)
    End Sub
    Public Overloads Sub Load()
        Me.Load(NombreReporte_p)
    End Sub
    Public Sub Mostrar(ByVal Response As System.Web.HttpResponse)
        LlenarParametrosNoPasados()
        Dim rutaexportacion As String
        Dim MiArchivoLimpio As String
        Dim OpcionesExportacion As ExportOptions
        Dim OpcionesArchivoExportacion As New DiskFileDestinationOptions
        rutaexportacion = Request_p.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(rutaexportacion) = False Then
            System.IO.Directory.CreateDirectory(Request_p.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("pdf")
        MiArchivo = rutaexportacion & MiArchivoLimpio
        OpcionesArchivoExportacion.DiskFileName = MiArchivo
        OpcionesExportacion = Me.ExportOptions
        With OpcionesExportacion
            .DestinationOptions = OpcionesArchivoExportacion
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
        End With
        Me.Export()
        If CC.Trim <> "" Then
            Try
                IO.File.Copy(MiArchivo, CC)
            Catch ex As Exception
            End Try
        End If
        With Response
            .Clear()
            .Buffer = True
            .ClearContent()
            .ClearHeaders()
            .ContentType = "application/pdf"
            .WriteFile(MiArchivo)
            .Flush()
            .Close()
        End With
        System.IO.File.Delete(MiArchivo)
    End Sub
    Public Sub CargarSubReporte(ByVal Subreporte As String, ByVal Datos As DataTable)
        Dim oSubRpt As New ReportDocument
        Dim crSections As Sections
        Dim crSection As Section
        Dim crReportObjects As ReportObjects
        Dim crReportObject As ReportObject
        Dim crSubreportObject As SubreportObject
        crSections = Me.ReportDefinition.Sections
        For Each crSection In crSections
            crReportObjects = crSection.ReportObjects
            For Each crReportObject In crReportObjects
                If crReportObject.Kind = ReportObjectKind.SubreportObject Then
                    crSubreportObject = CType(crReportObject, SubreportObject)
                    oSubRpt = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName)
                    If oSubRpt.Name.Trim = Subreporte Then
                        oSubRpt.SetDataSource(Datos)
                    End If
                End If
            Next
        Next
    End Sub
    Public Sub LlenarParametrosNoPasados()
        Dim N As Integer
        For N = 0 To Me.DataDefinition.ParameterFields.Count - 1
            If Not Me.DataDefinition.ParameterFields(N).HasCurrentValue Then
                Select Case Me.DataDefinition.ParameterFields(N).ParameterValueKind.ToString.ToLower
                    Case "stringparameter"
                        AgregarParametros(Me.DataDefinition.ParameterFields(N).Name, "**" & Me.DataDefinition.ParameterFields(N).Name & "**")
                    Case "numberparameter"
                        AgregarParametros(Me.DataDefinition.ParameterFields(N).Name, "-99999")
                    Case "dateparameter"
                    Case Else
                        Throw New System.Exception("Parameter Mising......!")
                End Select
            End If
        Next
    End Sub

    Public Function CrearPDF(ByVal Source As DataTable, ByVal Response As System.Web.HttpResponse) As String
        LlenarParametrosNoPasados()
        Dim rutaexportacion As String
        Dim MiArchivoLimpio As String
        Dim OpcionesExportacion As ExportOptions
        Dim OpcionesArchivoExportacion As New DiskFileDestinationOptions
        Me.SetDataSource(Source)
        rutaexportacion = Request_p.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(rutaexportacion) = False Then
            System.IO.Directory.CreateDirectory(Request_p.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("pdf")
        MiArchivo = rutaexportacion & MiArchivoLimpio
        OpcionesArchivoExportacion.DiskFileName = MiArchivo
        OpcionesExportacion = Me.ExportOptions
        With OpcionesExportacion
            .DestinationOptions = OpcionesArchivoExportacion
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
        End With
        Me.Export()
        Return MiArchivoLimpio
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
    Public Function CrearXLS(ByVal Response As System.Web.HttpResponse) As String
        LlenarParametrosNoPasados()
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim OpcionesExportacion As ExportOptions
        Dim OpcionesArchivoExportacion As New DiskFileDestinationOptions
        RutaExportacion = Request_p.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request_p.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("xls")
        MiArchivo = RutaExportacion & MiArchivoLimpio
        OpcionesArchivoExportacion.DiskFileName = MiArchivo
        OpcionesExportacion = Me.ExportOptions
        With OpcionesExportacion
            .DestinationOptions = OpcionesArchivoExportacion
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.Excel
        End With
        Me.Export()
        Return MiArchivoLimpio
    End Function
    Public Function CrearPDF() As String
        LlenarParametrosNoPasados()
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim OpcionesExportacion As ExportOptions
        Dim OpcionesArchivoExportacion As New DiskFileDestinationOptions
        RutaExportacion = Request_p.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request_p.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("pdf")
        MiArchivo = RutaExportacion & MiArchivoLimpio
        OpcionesArchivoExportacion.DiskFileName = MiArchivo
        OpcionesExportacion = Me.ExportOptions
        With OpcionesExportacion
            .DestinationOptions = OpcionesArchivoExportacion
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
        End With
        Me.Export()
        Return MiArchivoLimpio
    End Function
    Public Function CrearPDF(ByVal Response As System.Web.HttpResponse) As String
        LlenarParametrosNoPasados()
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim OpcionesExportacion As ExportOptions
        Dim OpcionesArchivoExportacion As New DiskFileDestinationOptions
        RutaExportacion = Request_p.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request_p.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreArchivoTemporal("pdf")
        MiArchivo = RutaExportacion & MiArchivoLimpio
        OpcionesArchivoExportacion.DiskFileName = MiArchivo
        OpcionesExportacion = Me.ExportOptions
        With OpcionesExportacion
            .DestinationOptions = OpcionesArchivoExportacion
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
        End With
        Me.Export()
        Return MiArchivoLimpio
    End Function
    Public Function CrearPDF(ByVal Response As System.Web.HttpResponse, ByVal File As String) As String
        LlenarParametrosNoPasados()
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim OpcionesExportacion As ExportOptions
        Dim OpcionesArchivoExportacion As New DiskFileDestinationOptions
        RutaExportacion = Request_p.PhysicalApplicationPath & "Exportaciones\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request_p.PhysicalApplicationPath & "Exportaciones\")
        End If
        MiArchivoLimpio = NombreDeArchivo(File)
        MiArchivo = RutaExportacion & MiArchivoLimpio
        OpcionesArchivoExportacion.DiskFileName = MiArchivo
        OpcionesExportacion = Me.ExportOptions
        With OpcionesExportacion
            .DestinationOptions = OpcionesArchivoExportacion
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
        End With
        Me.Export()
        Return MiArchivoLimpio
    End Function

End Class