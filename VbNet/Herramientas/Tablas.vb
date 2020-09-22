Imports System.Web.UI
Imports System.Drawing
Imports System.Web.UI.WebControls
Public Class Tablas
    Public Function ArmarCelda(ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal Texto As String, ByVal Negrita As Boolean, ByVal ColorDeFondo As String, ByVal AlineacionH As HorizontalAlign, ByVal AlineacionV As VerticalAlign, ByVal ColorDeBorde As String, ByVal EstiloDeBorde As BorderStyle, ByVal AnchoDeBorde As Integer, ByVal TamañoDeFuente As Integer, ByVal Fuente As String, ByVal AlturaDecelda As Integer, ByVal ColorDeFuente As String, ByVal Css As String) As TableCell
        Dim Estilo As String = ""
        Dim Style As String = ""
        If Texto.IndexOf("|") <> -1 Then
            Estilo = Texto.Substring(Texto.IndexOf("|")).Trim.ToUpper
            Texto = Texto.Substring(0, Texto.IndexOf("|"))
            Dim Attributes() As String = Estilo.Split(CChar("|"))
            For n As Integer = 0 To Attributes.Length - 1
                If Attributes(n).StartsWith("WIDTH") Then
                    Style &= "width:" & Attributes(n).Replace("WIDTH:", "") & ";"
                End If
                If Attributes(n).StartsWith("HA") Then
                    Select Case Attributes(n).Replace("HA:", "")
                        Case "L"
                            Style &= "text-align:left;"
                        Case "R"
                            Style &= "text-align:right;"
                        Case "C"
                            Style &= "text-align:center;"
                    End Select
                End If
            Next
        End If
        Dim LaCelda As New TableCell
        With LaCelda
            .Text = Texto
            If Css.Trim.Length = 0 And Css.Trim <> "." Then
                .Font.Bold = Negrita
                .BackColor = Color.FromName(ColorDeFondo)
                .BorderColor = Color.FromName(ColorDeBorde)
                .BorderStyle = EstiloDeBorde
                .ForeColor = Color.FromName(ColorDeFuente)
                If AnchoDeBorde <> 0 Then .BorderWidth = WebControls.Unit.Pixel(AnchoDeBorde)
                If TamañoDeFuente <> 0 Then .Font.Size = WebControls.FontUnit.Point(TamañoDeFuente)
                If AlturaDecelda <> 0 Then .Height = WebControls.Unit.Pixel(AlturaDecelda)
                If Fuente <> "" Then .Font.Name = Fuente
                .HorizontalAlign = AlineacionH
                .VerticalAlign = AlineacionV
            Else
                .CssClass = Css
                Select Case AlineacionH
                    Case HorizontalAlign.Center
                        Style &= "text-align:center;"
                    Case HorizontalAlign.Justify
                        Style &= "text-align:center;"
                    Case HorizontalAlign.Left
                        Style &= "text-align:left;"
                    Case HorizontalAlign.Right
                        Style &= "text-align:right;"
                End Select
                Select Case AlineacionV
                    Case VerticalAlign.Middle
                        Style &= "vertical-align:middle;"
                    Case VerticalAlign.Bottom
                        Style &= "vertical-align:bottom;"
                    Case VerticalAlign.Top
                        Style &= "vertical-align:top;"
                End Select
            End If
            If Not ColorDeFuente Is Nothing AndAlso ColorDeFuente.Trim.Length > 0 Then
                Style &= "color:" & ColorDeFuente.Trim & ";"
            End If
            If Style.Trim.Length > 0 Then
                .Attributes.Add("style", Style)
            End If
            If Rowspan <> 1 Then .RowSpan = Rowspan
            If ColumnSpan <> 1 Then .ColumnSpan = ColumnSpan
        End With
        Return LaCelda
    End Function
    Public Function ArmarCelda(ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal Texto As String, ByVal Negrita As Boolean, ByVal ColorDeFondo As String, ByVal AlineacionH As HorizontalAlign, ByVal AlineacionV As VerticalAlign, ByVal ColorDeBorde As String, ByVal EstiloDeBorde As BorderStyle, ByVal AnchoDeBorde As Integer, ByVal TamañoDeFuente As Integer, ByVal Fuente As String, ByVal AlturaDecelda As Integer, ByVal ColorDeFuente As String) As TableCell
        Dim LaCelda As New TableCell
        With LaCelda
            .Text = Texto
            .Font.Bold = Negrita
            .BackColor = Color.FromName(ColorDeFondo)
            .HorizontalAlign = AlineacionH
            .VerticalAlign = AlineacionV
            .BorderColor = Color.FromName(ColorDeBorde)
            .BorderStyle = EstiloDeBorde
            .ForeColor = Color.FromName(ColorDeFuente)
            If AnchoDeBorde <> 0 Then .BorderWidth = WebControls.Unit.Pixel(AnchoDeBorde)
            If TamañoDeFuente <> 0 Then .Font.Size = WebControls.FontUnit.Point(TamañoDeFuente)
            If AlturaDecelda <> 0 Then .Height = WebControls.Unit.Pixel(AlturaDecelda)
            If Fuente <> "" Then .Font.Name = Fuente
            If Rowspan <> 0 Then .RowSpan = Rowspan
            If ColumnSpan <> 0 Then .ColumnSpan = ColumnSpan
        End With
        Return LaCelda
    End Function
    Public Function ArmarCelda(ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal Texto As String, ByVal CssClass As String) As TableCell
        Dim Estilo As String = ""
        Dim Style As String = ""
        If Texto.IndexOf("|") <> -1 Then
            Estilo = Texto.Substring(Texto.IndexOf("|")).Trim.ToUpper
            Texto = Texto.Substring(0, Texto.IndexOf("|"))
            Dim Attributes() As String = Estilo.Split(CChar("|"))
            For n As Integer = 0 To Attributes.Length - 1
                If Attributes(n).StartsWith("WIDTH") Then
                    Style &= "width:" & Attributes(n).Replace("WIDTH:", "") & ";"
                End If
                If Attributes(n).StartsWith("HA") Then
                    Select Case Attributes(n).Replace("HA:", "")
                        Case "L"
                            Style &= "text-align:left;"
                        Case "R"
                            Style &= "text-align:right;"
                        Case "C"
                            Style &= "text-align:center;"
                    End Select
                End If
            Next
        End If
        Dim LaCelda As New TableCell
        With LaCelda
            .Text = Texto
            If Rowspan <> 0 Then .RowSpan = Rowspan
            If ColumnSpan <> 0 Then .ColumnSpan = ColumnSpan
            If CssClass.Trim.Length > 0 Then .CssClass = CssClass
            If Style.Trim.Length > 0 Then
                .Attributes.Add("style", Style)
            End If
        End With
        Return LaCelda
    End Function
    Public Function ArmarFila(ByVal Valores() As String, ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal Negrita As Boolean, ByVal ColorDeFondo As String, ByVal AlineacionH As HorizontalAlign, ByVal AlineacionV As VerticalAlign, ByVal ColorDeBorde As String, ByVal EstiloDeBorde As BorderStyle, ByVal AnchoDeBorde As Integer, ByVal TamañoDeFuente As Integer, ByVal Fuente As String, ByVal AlturaDecelda As Integer, ByVal ColorDeFuente As String, ByVal CssClass As String) As TableRow
        Dim LaFila As New TableRow
        Dim LaCelda As New TableCell
        Dim N As Integer
        For N = 0 To Valores.Length - 1
            LaFila.Cells.Add(ArmarCelda(Rowspan, ColumnSpan, Valores(N), Negrita, ColorDeFondo, AlineacionH, AlineacionV, ColorDeBorde, EstiloDeBorde, AnchoDeBorde, TamañoDeFuente, Fuente, AlturaDecelda, ColorDeFuente, ""))
        Next
        LaFila.CssClass = CssClass
        Return LaFila
    End Function
    Public Function ArmarFila(ByVal Valores() As String, ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal Negrita As Boolean, ByVal ColorDeFondo As String, ByVal AlineacionH As HorizontalAlign, ByVal AlineacionV As VerticalAlign, ByVal ColorDeBorde As String, ByVal EstiloDeBorde As BorderStyle, ByVal AnchoDeBorde As Integer, ByVal TamañoDeFuente As Integer, ByVal Fuente As String, ByVal AlturaDecelda As Integer, ByVal ColorDeFuente As String) As TableRow
        Dim LaFila As New TableRow
        Dim LaCelda As New TableCell
        Dim N As Integer
        For N = 0 To Valores.Length - 1
            LaFila.Cells.Add(ArmarCelda(Rowspan, ColumnSpan, Valores(N), Negrita, ColorDeFondo, AlineacionH, AlineacionV, ColorDeBorde, EstiloDeBorde, AnchoDeBorde, TamañoDeFuente, Fuente, AlturaDecelda, ColorDeFuente))
        Next
        Return LaFila
    End Function
    Public Function ArmarFila(ByVal Valores() As String, ByVal Rowspan As Integer, ByVal ColumnSpan As Integer, ByVal CssClass As String) As TableRow
        Dim LaFila As New TableRow
        Dim LaCelda As New TableCell
        Dim N As Integer
        For N = 0 To Valores.Length - 1
            LaFila.Cells.Add(ArmarCelda(Rowspan, ColumnSpan, Valores(N), ""))
        Next
        If CssClass.Trim.Length > 0 Then
            LaFila.CssClass = CssClass
        End If
        Return LaFila
    End Function

    Public Function CabeceraEspacial(ByVal Titulo As String, ByVal Ancho As Integer) As String
        Dim cadena As New System.Text.StringBuilder
        With cadena
            .Append("<table width='")
            .Append(Ancho.ToString)
            .Append("%' border='0' cellpadding='0' cellspacing='0' background='../ImgTablas/cat_back.gif'>")
            .Append("<TR>")
            .Append("<td width='140' height='27' align='left' valign='top'><img src='../ImgTablas/cat_top_ls.gif' width='140' height='27' alt='' border='0'></td> ")
            .Append("<td width='100%' background='../ImgTablas/cat_back.gif' valign='middle' align='center'>")
            .Append(Titulo)
            .Append("<td width='140' height='27' align='right' valign='top'><img src='../ImgTablas/cat_top_rs.gif' width='140' height='27' alt='' border='0'></td>")
            .Append("</TR>")
            .Append("</table>")
            Return .ToString
        End With
    End Function
    Public Function TituloColumnasEspacial(ByVal Elementos() As String, ByVal Ancho As Integer) As String
        Dim cadena As New System.Text.StringBuilder
        Dim n As Integer
        With cadena
            .Append("<table class='tborder' cellpadding='6' cellspacing='1' border='0' width='")
            .Append(Ancho.ToString)
            .Append("%' align='center'>")
            .Append("<tr align='center'> ")
            For n = 0 To Elementos.Length - 1
                .Append("<td class='thead'>")
                .Append(Elementos(n))
                .Append("</td>")
            Next
            .Append("</tr>")
            Return .ToString
        End With
    End Function
    Public Function PieEspacial(ByVal Ancho As Integer) As String
        Dim cadena As New System.Text.StringBuilder
        With cadena
            .Append("<table width='")
            .Append(Ancho.ToString)
            .Append("%' border='0' cellpadding='0' cellspacing='0' background='../ImgTablas/cat_back.gif'>")
            .Append("<TR>")
            .Append("<td width='70' align='left' valign='top'><img src='../ImgTablas/ls_main_table_bottom.gif' width='70' height='14' alt='' border='0'></td>")
            .Append("<td width='100%' Background='../ImgTablas/extended_main_table_bottom.gif'><img src='../ImgTablas/clear.gif' width='100%' height='14' alt='' border='0'></td>")
            .Append("<td width='70' align='right' valign='top'><img src='../ImgTablas/rs_main_table_bottom.gif' width='70' height='14' alt='' border='0'></td>")
            .Append("</TR>")
            .Append("</table>")
            Return .ToString
        End With
    End Function
End Class

