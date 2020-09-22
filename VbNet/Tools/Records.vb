Option Infer Off
Option Explicit On
Option Strict On
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.IO
Public Class Records
#Region "ToDouble"
    Public Function ToDouble(ByVal Numero As String) As Double
        Numero = Numero.Replace("$", "").Trim
        If Numero.Trim.Length = 0 Then Return 0
        Try
            Dim entero As New System.Text.StringBuilder
            Dim Mydecimals As New System.Text.StringBuilder
            Dim Divisor As New System.Text.StringBuilder
            Divisor.Append("1")
            Dim MyCharacter As String
            Dim flag As Boolean = False
            Dim n As Integer
            For n = 0 To Numero.Trim.Length - 1
                MyCharacter = Left(Right(Numero, Numero.Length - n), 1)
                If Not IsNumeric(MyCharacter) And MyCharacter <> "-" And MyCharacter <> "+" Then
                    flag = True
                Else
                    If Not flag Then
                        entero.Append(MyCharacter)
                    Else
                        Mydecimals.Append(MyCharacter)
                        Divisor.Append("0")
                    End If
                End If
            Next
            ToDouble = Double.Parse(entero.ToString)
            If Mydecimals.ToString.Trim.Length <> 0 Then
                If ToDouble >= 0 Then
                    ToDouble = ToDouble + (Long.Parse(Mydecimals.ToString) / Long.Parse(Divisor.ToString))
                Else
                    ToDouble = ToDouble - (Long.Parse(Mydecimals.ToString) / Long.Parse(Divisor.ToString))
                End If
            End If
            If entero.ToString = "-0" Then ToDouble = ToDouble * -1
        Catch
            Return 0
        End Try
    End Function
#End Region
#Region "ToInt"
    Public Function ToInt(ByVal Number As String) As Integer
        If Number Is Nothing Then Return 0
        If Number = "" Then Return 0
        If Number.Trim.Length = 0 Then Return 0
        If Number.ToLower = "true" Or Number.ToLower = "verdadero" Then Return 1
        If Number.ToLower = "false" Or Number.ToLower = "falso" Then Return 0
        Try
            ToInt = Integer.Parse(Number)
        Catch
            Return 0
        End Try
    End Function
#End Region
    Public Function GroupRecords(ByVal Mydata As DataTable, ByVal GroupField As String, ByVal BreakField As String) As DataTable
        GroupRecords = Mydata.Clone
        Mydata = Sort(Mydata, GroupField)
        Dim Posicion As Integer
        Dim n As Integer
        Dim BreakFieldIsChar As Boolean = False
        With Mydata
            For n = 0 To .Columns.Count - 1
                If .Columns(n).ColumnName.Trim.ToLower = BreakField.Trim.ToLower Then
                    If .Columns(n).DataType Is System.Type.GetType("System.String") Then
                        BreakFieldIsChar = True
                    End If
                End If
                If .Columns(n).ColumnName.Trim.ToLower = GroupField.Trim.ToLower Then
                    Posicion = n
                End If
            Next
            Dim ArgsClone() As Object
            For n = 0 To .Rows.Count - 1
                Mydata.Rows(n)(BreakField) = Mydata.Rows(n)(BreakField).ToString
                If BreakFieldIsChar Then
                    If ToInt(GroupRecords.Compute("count(" & BreakField & ")", BreakField & "='" & Mydata.Rows(n)(BreakField).ToString.Replace("'", "''") & "'").ToString) = 0 Then
                        ArgsClone = .Rows(n).ItemArray()
                        ArgsClone(Posicion) = ToDouble(Mydata.Compute("sum(" & GroupField & ")", BreakField & "='" & Mydata.Rows(n)(BreakField).ToString.Replace("'", "''") & "'").ToString)
                        GroupRecords.Rows.Add(ArgsClone)
                    End If
                Else
                    If ToInt(GroupRecords.Compute("count(" & BreakField & ")", BreakField & "=" & Mydata.Rows(n)(BreakField).ToString.Replace("'", "''")).ToString) = 0 Then
                        ArgsClone = .Rows(n).ItemArray()
                        ArgsClone(Posicion) = ToDouble(Mydata.Compute("sum(" & GroupField & ")", BreakField & "=" & Mydata.Rows(n)(BreakField).ToString.Replace("'", "''")).ToString)
                        GroupRecords.Rows.Add(ArgsClone)
                    End If
                End If
            Next
            ArgsClone = Nothing
        End With
        With GroupRecords
            For n = 0 To .Rows.Count - 1
                .Rows(n)(BreakField) = .Rows(n)(BreakField).ToString
            Next
        End With
    End Function
    Public Function Sort(ByVal MyData As DataTable, ByVal Order As String) As DataTable
        Dim n As Integer
        Dim x As Integer
        Dim View As New DataView(MyData)
        Dim WorkTable As DataTable = View.Table.Clone
        Dim records As Integer = View.Count
        Dim Columns As Integer = View.Table.Columns.Count
        Dim Args(Columns - 1) As Object
        View.Sort = Order
        For n = 0 To records - 1
            For x = 0 To Columns - 1
                Args(x) = View(n)(x)
            Next
            WorkTable.Rows.Add(Args)
        Next
        View.Dispose()
        Return WorkTable
    End Function
    Public Function Filter(ByVal Mydata As DataTable, ByVal vFilter As String) As DataTable
        Dim n As Integer
        Dim x As Integer
        Dim View As New DataView(Mydata)
        Dim WorkTable As DataTable = View.Table.Clone
        View.RowFilter = vFilter
        Dim records As Integer = View.Count
        Dim Columns As Integer = View.Table.Columns.Count
        Dim Args(Columns - 1) As Object
        For n = 0 To records - 1
            For x = 0 To Columns - 1
                Args(x) = View(n)(x)
            Next
            WorkTable.Rows.Add(Args)
        Next
        View.Dispose()
        Return WorkTable
    End Function
#Region "TemporaryFileName"
    Public Function TemporaryFileName(ByVal Extension As String) As String
        Dim r As Random = New Random
        Dim NumeroRandom As String = r.Next(1, 999999).ToString
        TemporaryFileName = Date.Today.Year.ToString & Date.Today.Month.ToString & Date.Today.Day.ToString & Date.Today.Minute.ToString & Date.Today.Second.ToString & NumeroRandom.ToString & "." & Extension
    End Function
#End Region
    Public Function ExportToExcelXml(ByVal Title As String, ByVal Table As String, ByVal Request As System.Web.HttpRequest) As String
        Dim RutaExportacion As String
        Dim MiArchivoLimpio As String
        Dim MiArchivo As String
        Dim MiFile As StreamWriter
        Dim Cadena As New System.Text.StringBuilder
        RutaExportacion = Request.PhysicalApplicationPath & "Exportations\"
        If System.IO.Directory.Exists(RutaExportacion) = False Then
            System.IO.Directory.CreateDirectory(Request.PhysicalApplicationPath & "Exportations\")
        End If
        MiArchivoLimpio = TemporaryFileName("xls")
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
            .Append(Title)
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
            .Append(Table)
            .Append("</body>")
            .Append("</html>")
        End With
        MiFile = New StreamWriter(MiArchivo)
        MiFile.Write(Cadena.ToString)
        MiFile.Close()
        Return MiArchivoLimpio
    End Function
    Public Function Merge(ByVal MyData As DataTable, ByVal MyData2 As DataTable, ByVal Sort As String) As DataTable
        Dim n As Integer
        With MyData2
            Dim ArgsObj() As Object
            For n = 0 To .Rows.Count - 1
                ArgsObj = .Rows(n).ItemArray
                MyData.Rows.Add(ArgsObj)
            Next
            If Sort.Trim.Length <> 0 Then MyData = Me.Sort(MyData, Sort)
            ArgsObj = Nothing
        End With
        Return MyData
    End Function
    Public Sub DataTableToXls(ByVal dt As DataTable, ByVal FileName As String)
        Const FieldSeparator As String = vbTab
        Const RowSeparator As String = vbLf
        Dim output As New Text.StringBuilder()
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
    Public Sub DataTableToXML(ByVal dt As DataTable, ByVal FileName As String)
        Dim MyDataSet As New Data.DataSet
        MyDataSet.Tables.Add(dt.Copy)
        MyDataSet.WriteXml(FileName)
        MyDataSet.Dispose()
    End Sub
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
End Class

