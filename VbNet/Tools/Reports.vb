Option Explicit On 
Option Strict On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports FunctionLibrary.Functions
Public Class Reports
    Inherits ReportDocument
    Private NameReport_p As String
    Private Request_p As System.Web.HttpRequest
    Private MiArchivo As String
    Public Sub New(ByVal Name As String, ByVal Request As System.Web.HttpRequest)
        Dim Ilen As Integer
        Request_p = Request
        NameReport_p = Request_p.ServerVariables("PATH_TRANSLATED")
        While (Right(NameReport_p, 1) <> "\" And NameReport_p.Length <> 0)
            Ilen = NameReport_p.Length - 1
            NameReport_p = NameReport_p.Substring(0, Ilen)
        End While
        NameReport_p = NameReport_p & Name
    End Sub
    Public Sub AddParameters(ByVal ParamName As String, ByVal Value As String)
        If Value.Trim.Length > 254 Then Value = Value.Substring(0, 254)
        'ESTO ES PORQUE SI PASAS UN PARAMETRO DE TEXTO MAYOR A254 CARACTERES TIRA ERROR
        Dim MyParameter As New ParameterDiscreteValue
        MyParameter.Value = Value
        Dim Values As New ParameterValues
        MyParameter.Value = Value
        With Values
            .Clear()
            .Add(MyParameter)
        End With
        Me.DataDefinition.ParameterFields(ParamName).ApplyCurrentValues(Values)
    End Sub
    Public Overloads Sub Load(ByVal Source As DataTable)
        Me.Load(NameReport_p)
        Me.SetDataSource(Source)
    End Sub
    Public Sub Show(ByVal Response As System.Web.HttpResponse)
        FillParametersNoSent()
        Dim ExportationPath As String
        Dim MyClearFile As String
        Dim ExportationOptions As ExportOptions
        Dim ExportattionOptionFile As New DiskFileDestinationOptions
        ExportationPath = Request_p.PhysicalApplicationPath & "Exportations\"
        If System.IO.Directory.Exists(ExportationPath) = False Then System.IO.Directory.CreateDirectory(Request_p.PhysicalApplicationPath & "Exportations\")
        MyClearFile = TemporaryFileName("pdf")
        MiArchivo = ExportationPath & MyClearFile
        ExportattionOptionFile.DiskFileName = MiArchivo
        ExportationOptions = Me.ExportOptions
        With ExportationOptions
            .DestinationOptions = ExportattionOptionFile
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
        End With
        Me.Export()
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
    Public Sub LoadSubreport(ByVal SubReport As String, ByVal Mydata As DataTable)
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
                    If oSubRpt.Name.Trim = SubReport Then oSubRpt.SetDataSource(Mydata)
                End If
            Next
        Next
    End Sub
    Private Sub FillParametersNoSent()
        Dim N As Integer
        For N = 0 To Me.DataDefinition.ParameterFields.Count - 1
            If Not Me.DataDefinition.ParameterFields(N).HasCurrentValue Then
                Select Case Me.DataDefinition.ParameterFields(N).ParameterValueKind.ToString.ToLower
                    Case "stringparameter"
                        AddParameters(Me.DataDefinition.ParameterFields(N).Name, "**" & Me.DataDefinition.ParameterFields(N).Name & "**")
                    Case "numberparameter"
                        AddParameters(Me.DataDefinition.ParameterFields(N).Name, "-99999")
                End Select
            End If
        Next
    End Sub
End Class