Option Infer Off
Option Explicit On
Option Strict On
Imports System.Web.UI
Imports System.Drawing
Imports System.Web.UI.WebControls
Public Class Tables
    Public Function CreateCell(ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal Texto As String, ByVal Negrita As Boolean, ByVal ColorDeFondo As String, ByVal AlineacionH As HorizontalAlign, ByVal AlineacionV As VerticalAlign, ByVal ColorDeBorde As String, ByVal EstiloDeBorde As BorderStyle, ByVal BorderWidth As Integer, ByVal FontSize As Integer, ByVal Fuente As String, ByVal CellHeight As Integer, ByVal ColorDeFuente As String, ByVal Css As String) As TableCell
        Dim TheCell As New TableCell
        With TheCell
            .Text = Texto
            If Css.Trim.Length = 0 Then
                .Font.Bold = Negrita
                .BackColor = Color.FromName(ColorDeFondo)
                .BorderColor = Color.FromName(ColorDeBorde)
                .BorderStyle = EstiloDeBorde
                .ForeColor = Color.FromName(ColorDeFuente)
                If FontSize <> 0 Then .Font.Size = WebControls.FontUnit.Point(FontSize)
                If Fuente <> "" Then .Font.Name = Fuente
                If BorderWidth <> 0 Then .BorderWidth = WebControls.Unit.Pixel(BorderWidth)
                If CellHeight <> 0 Then .Height = WebControls.Unit.Pixel(CellHeight)
            Else
                .CssClass = Css
            End If
            .HorizontalAlign = AlineacionH
            .VerticalAlign = AlineacionV
            If Rowspan <> 0 Then .RowSpan = Rowspan
            If ColumnSpan <> 0 Then .ColumnSpan = ColumnSpan
        End With
        Return TheCell
    End Function
    Public Function CreateRow(ByVal Valores() As String, ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal Negrita As Boolean, ByVal ColorDeFondo As String, ByVal AlineacionH As HorizontalAlign, ByVal AlineacionV As VerticalAlign, ByVal ColorDeBorde As String, ByVal EstiloDeBorde As BorderStyle, ByVal BorderWidth As Integer, ByVal FontSize As Integer, ByVal Fuente As String, ByVal CellHeight As Integer, ByVal ColorDeFuente As String, ByVal Css As String) As TableRow
        Dim LaFila As New TableRow
        Dim TheCell As New TableCell
        Dim N As Integer
        For N = 0 To Valores.Length - 1
            LaFila.Cells.Add(CreateCell(Rowspan, ColumnSpan, Valores(N), Negrita, ColorDeFondo, AlineacionH, AlineacionV, ColorDeBorde, EstiloDeBorde, BorderWidth, FontSize, Fuente, CellHeight, ColorDeFuente, Css))
        Next
        Return LaFila
    End Function
End Class

