
Imports System.Data
Public Class Distinct
    Private _LastValues() As Object = {} ' stores last values... used to determine distinct values
    Dim _SourceTable As DataTable
    Dim _RecordCount As Long
    Dim _Fields() As String
#Region " ...ctor "
    Public Sub New()
        _SourceTable = Nothing
        _RecordCount = 0
        _Fields = Nothing
        _LastValues = Nothing
    End Sub

    Public Sub New(ByVal dt As DataTable)
        _SourceTable = dt
        _RecordCount = 0
        _Fields = Nothing
        _LastValues = Nothing
    End Sub

    Public Sub New(ByVal dt As DataTable, ByVal Fields() As String)
        Dim iIndex As Integer = 0
        _SourceTable = dt
        _RecordCount = 0
        _Fields = Fields
        _LastValues = Nothing
        TrimFields(_Fields)
        ReDim _LastValues(_Fields.Length - 1)
    End Sub

    Public Sub New(ByVal dt As DataTable, ByVal Fields As String)
        Dim iIndex As Integer = 0
        _SourceTable = dt
        _RecordCount = 0
        _Fields = Fields.Split(CChar(","))
        TrimFields(_Fields)
        ReDim _LastValues(_Fields.Length - 1)
    End Sub

    Public Sub Dispose()
        _SourceTable = Nothing
        _RecordCount = 0
        _Fields = Nothing
        _LastValues = Nothing
    End Sub
#End Region

#Region " Properties "
    Public Property Table() As DataTable
        Get
            Return _SourceTable
        End Get
        Set(ByVal value As DataTable)
            _SourceTable = value
        End Set
    End Property

    Public ReadOnly Property RecordCount() As Long
        Get
            Return _RecordCount
        End Get
    End Property

    Public Property Fields() As String()
        Get
            Return _Fields
        End Get
        Set(ByVal value As String())
            _Fields = value
        End Set
    End Property
#End Region

#Region " Private Methods "
    Private Function RowsToTable(ByVal drs As DataRow()) As DataTable 'converts dr() to datatable
        Dim dt As New DataTable
        Dim HasColumns As Boolean = False
        For Each dr As DataRow In drs
            If _Fields Is Nothing Then _Fields = GetFields(dr)
            If Not HasColumns Then
                For Each field As String In _Fields
                    dt.Columns.Add(field.Trim, dr.Table.Columns(field.Trim).DataType)
                Next
                HasColumns = True
            End If
            dt.ImportRow(dr)
        Next
        Return dt
    End Function

    Private Sub TrimFields(ByRef fields As String())
        For a As Integer = 0 To fields.Length - 1
            fields(a) = fields(a).Trim
        Next
    End Sub

    Private Function GetFields(ByVal dr As DataRow) As String() ' creates string array of column names
        Dim strFields() As String = Nothing
        Dim a As Integer = 0
        For Each col As DataColumn In dr.Table.Columns
            ReDim Preserve strFields(a)
            strFields(a) = col.ColumnName
            a += 1
        Next
        Return strFields
    End Function

    Private Function Exists(ByVal dr As DataRow) As Boolean ' compares each column to field array
        Dim oValue As Object
        Dim iIndex As Integer
        Dim bExists As Boolean = True
        If _LastValues Is Nothing OrElse _LastValues.Length = 0 Then
            ReDim _LastValues(_Fields.Length - 1)
            bExists = False
        End If
        For Each field As String In _Fields
            oValue = dr.Item(field)
            If bExists Then
                If TypeOf oValue Is String Then
                    If _LastValues(iIndex) Is DBNull.Value OrElse Not oValue.Equals(_LastValues(iIndex)) Then bExists = False
                Else
                    If Not oValue.Equals(_LastValues(iIndex)) Then bExists = False
                End If
            End If
            _LastValues(iIndex) = oValue
            iIndex = iIndex + 1
        Next
        Return bExists
    End Function
#End Region

#Region " Public Methods "
    'Internal object overloads
    Public Function SelectDistinct() As DataTable   ' uses internal source table object. No filtering or sorting.
        Return SelectDistinct(_SourceTable, _Fields, "", "")
    End Function

    Public Function SelectDistinct(ByVal Sort As String) As DataTable ' uses internal source table object. No filtering.
        Return SelectDistinct(_SourceTable, _Fields, "", Sort)
    End Function

    Public Function SelectDistinct(ByVal Fields() As String, ByVal Sort As String) As DataTable ' uses internal source table object. No filtering.
        Return SelectDistinct(_SourceTable, Fields, Sort)
    End Function

    Public Function SelectDistinct(ByVal Fields() As String, ByVal Filter As String, ByVal Sort As String) As DataTable ' uses internal source table object.
        Return SelectDistinct(_SourceTable, Fields, Filter, Sort)
    End Function

    Public Function SelectDistinct(ByVal Filter As String, ByVal Sort As String) As DataTable   ' uses internal source table object.
        Return SelectDistinct(_SourceTable, _Fields, Filter, Sort)
    End Function

    'Source Table Overloads
    Public Function SelectDistinct(ByVal SourceTable As DataTable) As DataTable ' no sorting or filtering/conditions
        Return SelectDistinct(SourceTable, _Fields, "", "")
    End Function

    Public Function SelectDistinct(ByVal SourceTable As DataTable, ByVal Fields() As String) As DataTable   ' no sorting or filtering/conditions
        Return SelectDistinct(SourceTable, Fields, "", "")
    End Function

    Public Function SelectDistinct(ByVal SourceTable As DataTable, ByVal Fields() As String, ByVal Sort As String) As DataTable ' no filtering/conditions
        Return SelectDistinct(SourceTable, Fields, "", Sort)
    End Function

    Public Function SelectDistinct(ByVal SourceTable As DataTable, ByVal Fields() As String, ByVal Filter As String, ByVal Sort As String) As DataTable
        Dim dt As New DataTable
        Dim t As Long = DateTime.Now.Ticks ' timing
        _Fields = Fields
        TrimFields(_Fields)
        If SourceTable.Columns.Count = 0 And SourceTable.Rows.Count = 0 Then Return Nothing
        Try
            SourceTable = RowsToTable(SourceTable.Select(Filter)) ' reduces the fields down to the passed fields
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
            Return Nothing
        End Try
        For Each field As String In Fields
            If SourceTable.Columns(field) Is Nothing Then Return SourceTable
            dt.Columns.Add(field, SourceTable.Columns(field).DataType)
        Next
        For Each dr As DataRow In SourceTable.Rows
            If Not Exists(dr) Then dt.ImportRow(dr)
        Next
        Try
            dt = RowsToTable(dt.Select("", Sort)) ' sorting
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
        _RecordCount = dt.Rows.Count
        _SourceTable = dt
        Return _SourceTable
    End Function

    'String Overloads
    Public Function SelectDistinct(ByVal SourceTable As DataTable, ByVal Fields As String) As DataTable   ' no sorting or filtering/conditions
        _Fields = Fields.Split(CChar(","))
        Return SelectDistinct(SourceTable, _Fields, "", "")
    End Function

    Public Function SelectDistinct(ByVal SourceTable As DataTable, ByVal Fields As String, ByVal Sort As String) As DataTable      ' no sorting or filtering/conditions
        _Fields = Fields.Split(CChar(","))
        Return SelectDistinct(SourceTable, _Fields, "", Sort)
    End Function

    Public Function SelectDistinct(ByVal SourceTable As DataTable, ByVal Fields As String, ByVal Filter As String, ByVal Sort As String) As DataTable    ' no sorting or filtering/conditions
        _Fields = Fields.Split(CChar(","))
        Return SelectDistinct(SourceTable, _Fields, Filter, Sort)
    End Function
#End Region
End Class