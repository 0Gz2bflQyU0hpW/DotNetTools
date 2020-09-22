Option Explicit On
Option Strict On
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Net
Imports SHDocVw
Imports System.Web
Imports System.Drawing
Imports System.Xml
Public Class Funciones
    'CLASS SUMMARY
    'Librería de funciones.
    'END CLASS SUMMARY
    Public Shared Function DateFullFormat(ByVal Fecha As String) As String
        Dim Vfecha As Date = TxtToFecha(Fecha)
        DateFullFormat = Vfecha.Day.ToString
        Select Case Vfecha.Month
            Case 1
                DateFullFormat = DateFullFormat & " de " & "Enero" & " de " & Vfecha.Year.ToString
            Case 2
                DateFullFormat = DateFullFormat & " de " & "Febrero" & " de " & Vfecha.Year.ToString
            Case 3
                DateFullFormat = DateFullFormat & " de " & "Marzo" & " de " & Vfecha.Year.ToString
            Case 4
                DateFullFormat = DateFullFormat & " de" & "Abril" & " de " & Vfecha.Year.ToString
            Case 5
                DateFullFormat = DateFullFormat & " de " & "Mayo" & " de " & Vfecha.Year.ToString
            Case 6
                DateFullFormat = DateFullFormat & " de " & "Junio" & " de " & Vfecha.Year.ToString
            Case 7
                DateFullFormat = DateFullFormat & " de " & "Julio" & " de " & Vfecha.Year.ToString
            Case 8
                DateFullFormat = DateFullFormat & " de " & "Agosto" & " de " & Vfecha.Year.ToString
            Case 9
                DateFullFormat = DateFullFormat & " de " & "Septiembre" & " de " & Vfecha.Year.ToString
            Case 10
                DateFullFormat = DateFullFormat & " de " & "Octubre" & " de " & Vfecha.Year.ToString
            Case 11
                DateFullFormat = DateFullFormat & " de " & "Noviembre" & " de " & Vfecha.Year.ToString
            Case 12
                DateFullFormat = DateFullFormat & " de " & "Diciembre" & " de " & Vfecha.Year.ToString
        End Select
    End Function
#Region "MilSecondsToFormatHour"
    Public Shared Function MilSecondsToFormatHour(ByVal Milseconds As Integer, ByVal whole_seconds As Boolean) As String
        If Milseconds < 1000 Then
            Return "00:00:00:" & Milseconds.ToString.Trim
        End If
        Dim MySpan As New TimeSpan(0, 0, 0, 0, Milseconds)
        Dim Txt As New StringBuilder
        If MySpan.Hours > 0 Then
            Txt.Append(MySpan.Hours.ToString())
            MySpan = MySpan.Subtract(New TimeSpan(0, MySpan.Hours, 0, 0))
        Else
            Txt.Append("00")
        End If
        If MySpan.Minutes > 0 Then
            Txt.Append(":" & (MySpan.Minutes + 100).ToString().Substring(1, 2))
            MySpan = MySpan.Subtract(New TimeSpan(0, 0, MySpan.Minutes, 0))
        Else
            Txt.Append(":00")
        End If
        If whole_seconds Then
            ' Display only whole seconds.
            If MySpan.Seconds > 0 Then
                Txt.Append(":" & (MySpan.Seconds + 100).ToString().Substring(1, 2))
            Else
                Txt.Append(":00")
            End If
            If MySpan.Milliseconds > 0 Then
                Txt.Append(":" & MySpan.Milliseconds.ToString)
            End If
        Else
            ' Display fractional seconds.
            Txt.Append(":" & MySpan.TotalSeconds.ToString())
        End If
        Return Txt.ToString
    End Function
#End Region
#Region "IsDigitNumeric"
    Public Shared Function IsDigitNumeric(ByVal Text As String) As Boolean
        Select Case Text.Trim
            Case "0"
                IsDigitNumeric = True
            Case "1"
                IsDigitNumeric = True
            Case "2"
                IsDigitNumeric = True
            Case "3"
                IsDigitNumeric = True
            Case "4"
                IsDigitNumeric = True
            Case "5"
                IsDigitNumeric = True
            Case "6"
                IsDigitNumeric = True
            Case "7"
                IsDigitNumeric = True
            Case "8"
                IsDigitNumeric = True
            Case "9"
                IsDigitNumeric = True
            Case Else
                IsDigitNumeric = False
        End Select
    End Function
#End Region
#Region "GoogleMapReferences"
    Public Shared Function GoogleMapReferences(ByVal Address As String) As XmlDocument
        GoogleMapReferences = New XmlDocument
        Try
            Dim url As String = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & Address & "&sensor=true&key=AIzaSyBqBTl02kWThbJICaDvKI2KcF3rnwnRuVs"
            Dim request As HttpWebRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim reader As New StreamReader(response.GetResponseStream)
            GoogleMapReferences.LoadXml(reader.ReadToEnd())
            request = Nothing
            response = Nothing
            reader.Dispose()
        Catch ex As Exception
            SaveError("Antares", "LibreriaNet", "Funciones", "GoogleMapReferences", ex.Source, ex.Message, "", "", ex.StackTrace, "")
        End Try
    End Function
#End Region
#Region "HtmlDecode"
    Public Shared Function HtmlDecode(ByVal Text As String) As String
        Text = Text.Replace("&amp;#39;", "'")
        Text = Text.Replace("<br>", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("&lt;br&gt;", "")
        Text = Text.Replace("%20", " ")
        Text = Text.Replace("%21", "!")
        Text = Text.Replace("%22", Convert.ToChar(34))
        Text = Text.Replace("%23", "#")
        Text = Text.Replace("%24", "$")
        Text = Text.Replace("%25", "%")
        Text = Text.Replace("%26", "&")
        Text = Text.Replace("%27", "'")
        Text = Text.Replace("%28", "(")
        Text = Text.Replace("%29", ")")
        Text = Text.Replace("%2A", "*")
        Text = Text.Replace("%2B", "+")
        Text = Text.Replace("%2C", ",")
        Text = Text.Replace("%2D", "-")
        Text = Text.Replace("%2E", ".")
        Text = Text.Replace("%2F", "/")
        Text = Text.Replace("%30", "0")
        Text = Text.Replace("%31", "1")
        Text = Text.Replace("%32", "2")
        Text = Text.Replace("%33", "3")
        Text = Text.Replace("%34", "4")
        Text = Text.Replace("%2F", "/")
        Text = Text.Replace("%ED", "í")
        Text = Text.Replace("%E1", "á")
        Text = Text.Replace("%E9", "é")
        Text = Text.Replace("%C9", "É")
        Text = Text.Replace("%C1", "Á")
        Text = Text.Replace("%F3", "ó")
        Text = Text.Replace("%CD", "Í")
        Text = Text.Replace("%D1", "Ú")
        Text = Text.Replace("%D3", "Ó")
        Text = Text.Replace("%DA", "Ú")
        Text = Text.Replace("%FA", "ú")
        Text = Text.Replace("%F1", "Ñ")
        Text = Text.Replace("%3A", ":")
        Text = Text.Replace("%35", "5")
        Text = Text.Replace("%36", "6")
        Text = Text.Replace("%37", "7")
        Text = Text.Replace("%38", "8")
        Text = Text.Replace("%39", "9")
        Text = Text.Replace("%3A", ":")
        Text = Text.Replace("%3B", ";")
        Text = Text.Replace("%3C", "<")
        Text = Text.Replace("%3D", "=")
        Text = Text.Replace("%3E", ">")
        Text = Text.Replace("%3F", "?")
        Text = Text.Replace("%40", "@")
        Text = Text.Replace("%41", "A")
        Text = Text.Replace("%42", "B")
        Text = Text.Replace("%43", "C")
        Text = Text.Replace("%44", "D")
        Text = Text.Replace("%45", "E")
        Text = Text.Replace("%46", "F")
        Text = Text.Replace("%47", "G")
        Text = Text.Replace("%48", "H")
        Text = Text.Replace("%49", "I")
        Text = Text.Replace("%4A", "J")
        Text = Text.Replace("%4B", "K")
        Text = Text.Replace("%4C", "L")
        Text = Text.Replace("%4D", "M")
        Text = Text.Replace("%4E", "N")
        Text = Text.Replace("%4F", "O")
        Text = Text.Replace("%50", "P")
        Text = Text.Replace("%51", "Q")
        Text = Text.Replace("%52", "R")
        Text = Text.Replace("%53", "S")
        Text = Text.Replace("%54", "T")
        Text = Text.Replace("%55", "V")
        Text = Text.Replace("%56", "V")
        Text = Text.Replace("%57", "W")
        Text = Text.Replace("%58", "X")
        Text = Text.Replace("%59", "Y")
        Text = Text.Replace("%5A", "Z")
        Text = Text.Replace("%5B", "[")
        Text = Text.Replace("%5C", "\")
        Text = Text.Replace("%5D", "]")
        Text = Text.Replace("%5E", "^")
        Text = Text.Replace("%5F", "_")
        Text = Text.Replace("%60", "`")
        Text = Text.Replace("%61", "a")
        Text = Text.Replace("%62", "b")
        Text = Text.Replace("%63", "c")
        Text = Text.Replace("%64", "d")
        Text = Text.Replace("%65", "e")
        Text = Text.Replace("%66", "f")
        Text = Text.Replace("%67", "g")
        Text = Text.Replace("%68", "h")
        Text = Text.Replace("%69", "i")
        Text = Text.Replace("%6A", "j")
        Text = Text.Replace("%6B", "k")
        Text = Text.Replace("%6C", "l")
        Text = Text.Replace("%6D", "m")
        Text = Text.Replace("%6E", "n")
        Text = Text.Replace("%6F", "o")
        Text = Text.Replace("%70", "p")
        Text = Text.Replace("%71", "q")
        Text = Text.Replace("%72", "r")
        Text = Text.Replace("%73", "s")
        Text = Text.Replace("%74", "t")
        Text = Text.Replace("%75", "u")
        Text = Text.Replace("%76", "v")
        Text = Text.Replace("%77", "w")
        Text = Text.Replace("%78", "x")
        Text = Text.Replace("%79", "y")
        Text = Text.Replace("%7A", "z")
        Text = Text.Replace("%7B", "{")
        Text = Text.Replace("%7C", "|")
        Text = Text.Replace("%7D", "}")
        Text = Text.Replace("%7E", "~")
        Text = Text.Replace("%7F", " ")
        Text = Text.Replace("%80", "€")
        Text = Text.Replace("%81", " ")
        Text = Text.Replace("%82", "‚")
        Text = Text.Replace("%83", "ƒ")
        Text = Text.Replace("%84", "„")
        Text = Text.Replace("%85", "…")
        Text = Text.Replace("%86", "†")
        Text = Text.Replace("%87", "‡")
        Text = Text.Replace("%88", "ˆ")
        Text = Text.Replace("%89", "‰")
        Text = Text.Replace("%8A", "Š")
        Text = Text.Replace("%8B", "‹")
        Text = Text.Replace("%8C", "Œ")
        Text = Text.Replace("%8D", " ")
        Text = Text.Replace("%8E", "Ž")
        Text = Text.Replace("%8F", " ")
        Text = Text.Replace("%90", " ")
        Text = Text.Replace("%91", "‘")
        Text = Text.Replace("%92", "’")
        'Text = Text.Replace("%93","“")
        'Text = Text.Replace("%94","”")
        Text = Text.Replace("%95", "•")
        Text = Text.Replace("%96", "–")
        Text = Text.Replace("%97", "—")
        Text = Text.Replace("%98", "˜")
        Text = Text.Replace("%99", "™")
        Text = Text.Replace("%9A", "š")
        Text = Text.Replace("%9B", "›")
        Text = Text.Replace("%9C", "œ")
        Text = Text.Replace("%9D", " ")
        Text = Text.Replace("%9E", "ž")
        Text = Text.Replace("%9F", "Ÿ")
        Text = Text.Replace("%A0", " ")
        Text = Text.Replace("%A1", "¡")
        Text = Text.Replace("%A2", "¢")
        Text = Text.Replace("%A3", "£")
        Text = Text.Replace("%A4", " ")
        Text = Text.Replace("%A5", "¥")
        Text = Text.Replace("%A6", "|")
        Text = Text.Replace("%A7", "§")
        Text = Text.Replace("%A8", "¨")
        Text = Text.Replace("%A9", "©")
        Text = Text.Replace("%AA", "ª")
        Text = Text.Replace("%AB", "«")
        Text = Text.Replace("%AC", "¬")
        Text = Text.Replace("%AD", "¯")
        Text = Text.Replace("%AE", "®")
        Text = Text.Replace("%AF", "¯")
        Text = Text.Replace("%B0", "°")
        Text = Text.Replace("%B1", "±")
        Text = Text.Replace("%B2", "²")
        Text = Text.Replace("%B3", "³")
        Text = Text.Replace("%B4", "´")
        Text = Text.Replace("%B5", "µ")
        Text = Text.Replace("%B6", "¶")
        Text = Text.Replace("%B7", "·")
        Text = Text.Replace("%B8", "¸")
        Text = Text.Replace("%B9", "¹")
        Text = Text.Replace("%BA", "º")
        Text = Text.Replace("%BB", "»")
        Text = Text.Replace("%BC", "¼")
        Text = Text.Replace("%BD", "½")
        Text = Text.Replace("%BE", "¾")
        Text = Text.Replace("%BF", "¿")
        Text = Text.Replace("%C0", "À")
        Text = Text.Replace("%C1", "Á")
        Text = Text.Replace("%C2", "Â")
        Text = Text.Replace("%C3", "Ã")
        Text = Text.Replace("%C4", "Ä")
        Text = Text.Replace("%C5", "Å")
        Text = Text.Replace("%C6", "Æ")
        Text = Text.Replace("%C7", "Ç")
        Text = Text.Replace("%C8", "È")
        Text = Text.Replace("%C9", "É")
        Text = Text.Replace("%CA", "Ê")
        Text = Text.Replace("%CB", "Ë")
        Text = Text.Replace("%CC", "Ì")
        Text = Text.Replace("%CD", "Í")
        Text = Text.Replace("%CE", "Î")
        Text = Text.Replace("%CF", "Ï")
        Text = Text.Replace("%D0", "Ð")
        Text = Text.Replace("%D1", "Ñ")
        Text = Text.Replace("%D2", "Ò")
        Text = Text.Replace("%D3", "Ó")
        Text = Text.Replace("%D4", "Ô")
        Text = Text.Replace("%D5", "Õ")
        Text = Text.Replace("%D6", "Ö")
        Text = Text.Replace("%D7", " ")
        Text = Text.Replace("%D8", "Ø")
        Text = Text.Replace("%D9", "Ù")
        Text = Text.Replace("%DA", "Ú")
        Text = Text.Replace("%DB", "Û")
        Text = Text.Replace("%DC", "Ü")
        Text = Text.Replace("%DD", "Ý")
        Text = Text.Replace("%DE", "Þ")
        Text = Text.Replace("%DF", "ß")
        Text = Text.Replace("%E0", "à")
        Text = Text.Replace("%E1", "á")
        Text = Text.Replace("%E2", "â")
        Text = Text.Replace("%E3", "ã")
        Text = Text.Replace("%E4", "ä")
        Text = Text.Replace("%E5", "å")
        Text = Text.Replace("%E6", "æ")
        Text = Text.Replace("%E7", "ç")
        Text = Text.Replace("%E8", "è")
        Text = Text.Replace("%E9", "é")
        Text = Text.Replace("%EA", "ê")
        Text = Text.Replace("%EB", "ë")
        Text = Text.Replace("%EC", "ì")
        Text = Text.Replace("%ED", "í")
        Text = Text.Replace("%EE", "î")
        Text = Text.Replace("%EF", "ï")
        Text = Text.Replace("%F0", "ð")
        Text = Text.Replace("%F1", "ñ")
        Text = Text.Replace("%F2", "ò")
        Text = Text.Replace("%F3", "ó")
        Text = Text.Replace("%F4", "ô")
        Text = Text.Replace("%F5", "õ")
        Text = Text.Replace("%F6", "ö")
        Text = Text.Replace("%F7", "÷")
        Text = Text.Replace("%F8", "ø")
        Text = Text.Replace("%F9", "ù")
        Text = Text.Replace("%FA", "ú")
        Text = Text.Replace("%FB", "û")
        Text = Text.Replace("%FC", "ü")
        Text = Text.Replace("%FD", "ý")
        Text = Text.Replace("%FE", "þ")
        Text = Text.Replace("%FF", "ÿ")

        Return Text
    End Function
#End Region
#Region "DameEmail"
    Public Shared Function DameEmail(ByVal Lineatexto As String) As String
        'METHOD SUMMARY
        'Extrae un email de una cadena de texto.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Lineatexto=Cadena de caracteres
        'END PARAMETERS SUMMARY
        DameEmail = ""
        If Lineatexto = Nothing Then Exit Function
        If Lineatexto.Trim.Length = 0 Then Exit Function
        Dim n As Integer = 0
        Dim PrimeraParte As String = ""
        Dim SegundaParte As String = ""
        Dim caracter As String
        Dim Barroba As Boolean = False
        For n = 0 To Lineatexto.Length - 1
            caracter = Lineatexto.Substring(n, 1)
            If caracter = "@" Then
                Barroba = True
            Else
                Select Case caracter.Trim
                    Case ""
                        If SegundaParte.Length <> 0 Then
                            Exit For
                        Else
                            PrimeraParte = ""
                        End If
                    Case "/"
                    Case Else
                        If Not Barroba Then
                            PrimeraParte = PrimeraParte & caracter
                        Else
                            SegundaParte = SegundaParte & caracter
                        End If
                End Select
            End If
        Next
        If SegundaParte.Trim.Length <> 0 And PrimeraParte.Trim.Length <> 0 Then
            DameEmail = PrimeraParte.Trim & "@" & SegundaParte.Trim
            If Not MailValido(DameEmail) Then DameEmail = ""
        End If
    End Function
#End Region
#Region "Encriptador"
    Private Shared ReadOnly Property LlavePrimaria() As Byte()
        'PROPERTY SUMMARY
        'Clave primaria para encryptar datos
        'END PROPERTY SUMMARY
        Get
            Dim Clave As String = "%/k@%q=?%%$%&9z{+]@2mxFh"
            Return Encoding.Default.GetBytes(Clave)
        End Get
    End Property
    Private Shared ReadOnly Property LlaveSecundaria() As Byte()
        'PROPERTY SUMMARY
        'Clave secundaria para encryptar datos
        'END PROPERTY SUMMARY
        Get
            Dim Clave As String = "1p!9z$%&"
            Return Encoding.Default.GetBytes(Clave)
        End Get
    End Property
    Public Shared Function Encryptador(ByVal Text As String) As String
        'METHOD SUMMARY
        'Encrypta una cadena de caracteres usando las librerias del framewaok de .NET
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Text=Cadena de texto a encryptar
        'END PARAMETERS SUMMARY
        Dim Des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        Return Transform(Text, Des.CreateEncryptor(LlavePrimaria, LlaveSecundaria)).Replace(Convert.ToChar(34), "JAR1").Replace(" ", "JAR2").Replace(Convert.ToChar(1), "JAR3").Trim
    End Function
    Public Shared Function DesEncryptador(ByVal encryptedText As String) As String
        'METHOD SUMMARY
        'Desencrypta una cadena encriptada con la funcion Encryptador.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'encryptedText=Texto a desencryptar
        'END PARAMETERS SUMMARY
        Dim Des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        Return Transform(encryptedText.ToString.Replace("JAR1", Convert.ToChar(34)).Replace("JAR3", Convert.ToChar(1)).Trim.Replace("JAR2", " "), Des.CreateDecryptor(LlavePrimaria, LlaveSecundaria))
    End Function
    Private Shared Function Transform(ByVal Text As String, ByVal CryptoTransform As ICryptoTransform) As String
        'METHOD SUMMARY
        'Función usada por Encryptador y DesEncryptador.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Text=Texto a encriptar
        'CryptoTransform=Clase de NET para encryptar
        'END PARAMETERS SUMMARY
        Dim Stream As MemoryStream = New MemoryStream
        Dim CryptoStream As CryptoStream = New CryptoStream(Stream, CryptoTransform, CryptoStreamMode.Write)
        Dim Input() As Byte = Encoding.Default.GetBytes(Text)
        CryptoStream.Write(Input, 0, Input.Length)
        Try
            CryptoStream.FlushFinalBlock()
            Return Encoding.Default.GetString(Stream.ToArray())
        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region
#Region "FechaATexto"
    Public Shared Function FechaATexto(ByVal Fecha As Date) As String
        'METHOD SUMMARY
        'Convierte una fecha a texto en formato DD/MM/YYYY.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a convertir
        'END PARAMETERS SUMMARY
        Dim Dia As String, Mes As String, Ano As String
        Dia = Fecha.Day.ToString
        Mes = Fecha.Month.ToString
        Ano = Fecha.Year.ToString
        If Dia.Length = 1 Then
            Dia = "0" & Dia
        End If
        If Mes.Length = 1 Then
            Mes = "0" & Mes
        End If
        FechaATexto = Dia & "/" & Mes & "/" & Ano
    End Function
#End Region
#Region "TxtFechaHoy"
    Public Shared Function TxtFechaHoy() As String
        'METHOD SUMMARY
        'Devuelve la fecha actual como un texo en formato DD/MM/YYYY.
        'END METHOD SUMMARY
        Dim Dia As String, Mes As String, Ano As String
        Dia = System.DateTime.Today().Day.ToString
        Mes = System.DateTime.Today().Month.ToString
        Ano = System.DateTime.Today().Year.ToString
        If Dia.Length = 1 Then
            Dia = "0" & Dia
        End If
        If Mes.Length = 1 Then
            Mes = "0" & Mes
        End If
        TxtFechaHoy = Dia & "/" & Mes & "/" & Ano
    End Function
#End Region
#Region "TxtFechaPrimeroMes"
    Public Shared Function TxtFechaPrimeroMes(ByVal Mes As Integer, ByVal Ano As Integer, ByVal Americana As Boolean) As String
        'METHOD SUMMARY
        'Devuelve una fecha como texto la cual corresponde al primero de mes
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Mes=Mes de la fecha
        'Ano=Año de la fecha
        'Americana=Devuelve en fomato americana
        'END PARAMETERS SUMMARY
        If Mes > 12 Or Mes < 1 Then Return ""
        Dim Dia As String, Vmes As String, Vano As String
        Dia = "01"
        Vmes = Mes.ToString.Trim
        Vano = Ano.ToString.Trim
        If Dia.Length = 1 Then
            Dia = "0" & Dia
        End If
        If Vmes.Length = 1 Then
            Vmes = "0" & Vmes
        End If
        If Americana Then
            TxtFechaPrimeroMes = Vmes & "/" & Dia & "/" & Vano
        Else
            TxtFechaPrimeroMes = Dia & "/" & Vmes & "/" & Vano
        End If
    End Function
#End Region
#Region "ValidoTxtFecha"
    Public Shared Function ValidoTxtFecha(ByVal Fecha As String) As Boolean
        'METHOD SUMMARY
        'Valida si una fecha esta en el formato DD/MM/YYYY.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Texto conteniendo una fecha en formato DD/MM/YYYY
        'END PARAMETERS SUMMARY
        If Fecha.Trim.Length <> 10 Then
            Return False
        End If
        Dim Dia As Integer = CEntero(Fecha.Substring(0, 2))
        Dim Mes As Integer = CEntero(Fecha.Substring(3, 2))
        Dim Ano As Integer = CEntero(Fecha.Substring(6, 4))
        Try
            Dim FechaDate As New System.DateTime(Ano, Mes, Dia)
            If FechaDate.Day <> Dia Or FechaDate.Month <> Mes Or FechaDate.Year <> Ano Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function
#End Region
#Region "EsNumero"
    Public Shared Function EsNumero(ByVal Numero As String) As Boolean
        'METHOD SUMMARY
        'Valida si una cadena de caracteres puede convertirse a numero.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Numero=Valor a analizar
        'END PARAMETERS SUMMARY
        Try
            Convert.ToDouble(Numero)
        Catch
            Return False
        End Try
        Return True
    End Function
#End Region
#Region "Izquierda"
    Public Shared Function Izquierda(ByVal Cadena As String, ByVal Posiciones As Integer) As String
        'METHOD SUMMARY
        'Extrae la parte izquierda de una cadena de caracteres.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Posiciones=Posiciones a devolver
        'Cadena=Texto en donde se extraera la cadena
        'END PARAMETERS SUMMARY
        If Posiciones > Cadena.Trim.Length Then Return Cadena
        Return Cadena.Trim.Substring(0, Posiciones)
    End Function
#End Region
#Region "Derecha"
    Public Shared Function Derecha(ByVal Cadena As String, ByVal Posiciones As Integer) As String
        'METHOD SUMMARY
        'Extrae la parte derecha de una cadena de caracteres.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Posiciones=Posiciones a extraer.
        'Cadena=Cadena de texto
        'END PARAMETERS SUMMARY
        If Posiciones > Cadena.Trim.Length Then Return Cadena
        Return Cadena.Trim.Substring(Cadena.Trim.Length - Posiciones, Posiciones)
    End Function
#End Region
#Region "IsSubmit"
    Public Shared Function IsSubmit(ByVal Cadena As String) As String
        'METHOD SUMMARY
        'Sirve para anlizar si el valor de un QueryString existe. Si es nothing lo devuelve como un texto vacio.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Cadena=Texto del QueryString
        'END PARAMETERS SUMMARY
        If Cadena = Nothing Then
            Return ""
        Else
            Return Cadena.Trim
        End If
    End Function
#End Region
#Region "CaracterNumerico"
    Private Shared Function CaracterNumerico(ByVal Valor As String) As Boolean
        If Valor.Trim = "1" Or Valor.Trim = "2" Or Valor.Trim = "3" Or Valor.Trim = "4" Or Valor.Trim = "5" Or Valor.Trim = "6" Or Valor.Trim = "7" Or Valor.Trim = "8" Or Valor.Trim = "9" Or Valor.Trim = "0" Or Valor.Trim = "-" Then
            CaracterNumerico = True
        Else
            CaracterNumerico = False
        End If
    End Function
#End Region
#Region "Centero"
    Public Shared Function CEntero(ByVal Numero As String) As Integer
        'METHOD SUMMARY
        'Convierte a Entero una cadena de texto.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Numero=Texto a convertir
        'END PARAMETERS SUMMARY
        If Numero = Nothing Then Return 0
        If Numero.Trim.Length = 0 Then Return 0
        Numero = Numero.Trim
        If Numero.ToLower = "-" Then Return 0
        If Numero.ToLower = "true" Or Numero.ToLower = "verdadero" Then Return 1
        If Numero.ToLower = "false" Or Numero.ToLower = "falso" Then Return 0
        For n As Integer = 0 To Numero.Length - 1
            If Not CaracterNumerico(Numero.Substring(n, 1)) Then
                Return 0
            End If
        Next
        'Try
        CEntero = Integer.Parse(Numero)
        'Catch
        'Return 0
        'End Try
    End Function
#End Region
#Region "TiempoEntreFechasHoras"
    Public Shared Function TiempoEntreFechasHoras(ByVal Diai As Integer, ByVal Mesi As Integer, ByVal Anoi As Integer, ByVal Horai As Integer, ByVal Minutosi As Integer, ByVal Segundosi As Integer, ByVal Diaf As Integer, ByVal Mesf As Integer, ByVal Anof As Integer, ByVal Horaf As Integer, ByVal Minutosf As Integer, ByVal Segundosf As Integer, ByVal CalculoDeHoras As Boolean) As String
        'METHOD SUMMARY
        'Devuelve la cantidad de dias entre 2 fechas.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Diai=Dia Inicial
        'Mesi=Mes Inicial
        'Horai=Hora Inicial
        'Minutosi=Minutos Iniciales
        'Segundosi=Segundos Iniciales
        'Anoi=Año Inicial
        'Diaf=Dia final
        'Mesf=Mes final
        'Anof=Año Final
        'Horaf=Hora Final
        'Minutosf=Minutos Final
        'Segundosf=Segundos Final
        'CalculoDeHoras=Bandera que especifica si el resultado devuelto es la unida de medidas de Horas.
        'END PARAMETERS SUMMARY
        If Diai > 31 Or Diai < 1 Then
            TiempoEntreFechasHoras = "Error en diai"
            Exit Function
        End If
        If Diaf > 31 Or Diaf < 1 Then
            TiempoEntreFechasHoras = "Error en diaf"
            Exit Function
        End If
        If Mesi > 12 Or Mesi < 1 Then
            TiempoEntreFechasHoras = "Error en mesi"
            Exit Function
        End If
        If Mesf > 12 Or Mesf < 1 Then
            TiempoEntreFechasHoras = "Error en mesf"
            Exit Function
        End If
        If Horai >= 24 Or Horai < 0 Then
            TiempoEntreFechasHoras = "Error en horai"
            Exit Function
        End If
        If Horaf >= 24 Or Horaf < 0 Then
            TiempoEntreFechasHoras = "Error en horaf"
            Exit Function
        End If
        If Minutosi >= 60 Or Minutosi < 0 Then
            TiempoEntreFechasHoras = "Error en minutosi"
            Exit Function
        End If
        If Minutosf >= 60 Or Minutosf < 0 Then
            TiempoEntreFechasHoras = "Error en minutosf"
            Exit Function
        End If
        If Segundosi >= 60 Or Segundosi < 0 Then
            TiempoEntreFechasHoras = "Error en segundosi"
            Exit Function
        End If
        If Segundosf >= 60 Or Segundosf < 0 Then
            TiempoEntreFechasHoras = "Error en segundosf"
            Exit Function
        End If
        Dim Mydate1 As New DateTime(Anoi, Mesi, Mesi, Horai, Minutosi, Segundosi)
        Dim Mydate2 As New DateTime(Anof, Mesf, Mesf, Horaf, Minutosf, Segundosf)
        Dim TiempoFechas As New System.TimeSpan(Mydate2.Ticks - Mydate1.Ticks)
        If TiempoFechas.Ticks < 0 Then
            If Not CalculoDeHoras Then
                TiempoEntreFechasHoras = "00:00:00:00"
            Else
                TiempoEntreFechasHoras = "00:00:00"
            End If
        Else
            If Not CalculoDeHoras Then
                TiempoEntreFechasHoras = (TiempoFechas.Days + 100).ToString.Substring(1, 2).Trim & ":" & (TiempoFechas.Hours + 100).ToString.Substring(1, 2).Trim & ":" & (TiempoFechas.TotalMinutes + 100).ToString.Substring(1, 2).Trim & ":" & (TiempoFechas.Seconds + 100).ToString.Substring(1, 2).Trim
            Else
                TiempoEntreFechasHoras = (TiempoFechas.Hours + 100).ToString.Substring(1, 2).Trim & ":" & (TiempoFechas.Minutes + 100).ToString.Substring(1, 2).Trim & ":" & (TiempoFechas.Seconds + 100).ToString.Substring(1, 2).Trim
            End If
        End If
    End Function
#End Region
#Region "FormatoDecimal"
    Public Shared Function FormatoDecimal(ByVal Valor As String, ByVal CantDec As Integer) As String
        'METHOD SUMMARY
        'Formatea un numero convertido a texto a una x cantidad de decimales. Funciona independientemente del formato de la maquina. Simpre devuelve como separador decimal el .
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'CantDec=Cantidad de decimales
        'Valor=Importe a formatear
        'END PARAMETERS SUMMARY
        Dim Bandera As Boolean
        Dim N As Integer, Largo As Integer
        Dim Cadena As String, Cadena2 As String = "", Pesos As String
        Dim Veces As Integer, Veces2 As Integer
        Dim Resto As String = "", X As Integer, Contador As Integer
        Dim Negativo As Boolean
        For X = 1 To CantDec
            Resto = Resto & "0"
        Next
        Largo = Valor.Trim.Length
        If Valor.Substring(0, 1) = "-" Then
            Negativo = True
            Valor = Derecha(Valor, Largo - 1)
        Else
            Negativo = False
        End If
        Valor = Math.Round(CDoble(Valor), CEntero(CantDec.ToString)).ToString
        Largo = Valor.Trim.Length
        If EsNumero(Valor) = False Then
            If CantDec = 0 Then
                FormatoDecimal = "0"
            Else
                FormatoDecimal = "0." & Resto
            End If
            Exit Function
        End If
        If Largo = 1 Then
            If EsNumero(Valor.Trim) = False Then
                If CantDec = 0 Then
                    FormatoDecimal = "0"
                Else
                    FormatoDecimal = "0." & Resto
                End If
                Exit Function
            Else
                FormatoDecimal = Valor.Trim & "." & Resto
                If Negativo = True Then FormatoDecimal = "-" & FormatoDecimal
                Exit Function
            End If
        End If
        If Largo = 0 Then
            If CantDec = 0 Then
                FormatoDecimal = "0"
            Else
                FormatoDecimal = "0." & Resto
            End If
            Exit Function
        End If
        For N = 0 To Valor.Trim.Length - 1
            Cadena = Izquierda(Derecha(Valor.Trim, Largo - N), 1)
            If EsNumero(Cadena) = False Then
                If Cadena = "-" Then
                    Veces2 = Veces2 + 1
                    Cadena2 = Cadena2 & Cadena
                Else
                    Cadena2 = Cadena2 & "."
                    Veces = Veces + 1
                End If
            Else
                Cadena2 = Cadena2 & Cadena
            End If
        Next
        If Veces > 1 Then
            If CantDec = 0 Then
                FormatoDecimal = "0"
            Else
                FormatoDecimal = "0." & Resto
            End If
            Exit Function
        End If
        If Veces2 > 1 Then
            If CantDec = 0 Then
                FormatoDecimal = "0"
            Else
                FormatoDecimal = "0." & Resto
            End If
            Exit Function
        End If
        Bandera = False
        Pesos = Cadena2
        If Pesos.Trim.Length = 1 Then
            Pesos = Pesos.Trim & "." & Resto
        Else
            Largo = Pesos.Trim.Length
            Contador = 0
            For N = 0 To Pesos.Trim.Length - 1
                Cadena = Izquierda(Derecha(Pesos.Trim, Largo - N), 1)
                If Contador > 0 Then Contador = Contador + 1
                If Cadena = "." Then
                    Contador = Contador + 1
                End If
            Next
        End If
        If Contador = 0 Then
            Pesos = Pesos & "." & Resto
        Else
            Cadena = ""
            For N = 1 To CantDec - (Contador - 1)
                Cadena = Cadena & "0"
            Next
            Pesos = Pesos & Cadena
        End If
        Largo = Pesos.Trim.Length
        For N = 0 To Pesos.Trim.Length - 1
            Cadena = Izquierda(Derecha(Pesos.Trim, Largo - N), 1)
            If Cadena = "." Or Cadena = "," Then
                Bandera = True
            End If
        Next
        If Bandera = False Then
            Pesos = Pesos.Trim & "." & Resto
        End If
        Bandera = False
        If Izquierda(Derecha(Pesos.Trim, CantDec + 1), 1) <> "." Then
            If CantDec = 0 Then
                FormatoDecimal = "0"
            Else
                FormatoDecimal = "0." & Resto
            End If
            Exit Function
        End If
        If Negativo = True Then Pesos = "-" & Pesos
        If CantDec = 0 And Pesos.Trim.Substring(Pesos.Trim.Length - 1, 1) = "." Then
            Pesos = Pesos.Trim.Substring(0, Pesos.Trim.Length - 1)
        End If
        FormatoDecimal = Pesos
        Exit Function
    End Function
#End Region
#Region "HoraSistema"
    Public Shared Function HoraSistema() As String
        'METHOD SUMMARY
        'Devuelve la hora del sistema como un texto en formato HH:MM:SS
        '
        '
        'END METHOD SUMMARY
        Dim Hora As String
        If System.DateTime.Now.Hour.ToString.Trim.Length = 1 Then
            Hora = "0" & System.DateTime.Now.Hour.ToString & ":"
        Else
            Hora = System.DateTime.Now.Hour.ToString & ":"
        End If
        If System.DateTime.Now.Minute.ToString.Trim.Length = 1 Then
            Hora = Hora & "0" & System.DateTime.Now.Minute.ToString & ":"
        Else
            Hora = Hora & System.DateTime.Now.Minute.ToString & ":"
        End If
        If System.DateTime.Now.Second.ToString.Trim.Length = 1 Then
            Hora = Hora & "0" & System.DateTime.Now.Second.ToString
        Else
            Hora = Hora & System.DateTime.Now.Second.ToString
        End If
        HoraSistema = Hora
    End Function
#End Region
#Region "FechaAmericana"
    Public Shared Function FechaAmericana(ByVal Fecha As String) As String
        'METHOD SUMMARY
        'Convierte un texto conteniendo una fecha en formato DD/MM/YYYY a MM/DD/YYYY
        '
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a formatear
        'END PARAMETERS SUMMARY
        If Fecha.Trim.Length <> 10 Then
            FechaAmericana = Fecha
            Exit Function
        End If
        Dim Dia As String, Mes As String, Ano As String
        Dia = Fecha.Trim.Substring(0, 2)
        Ano = Derecha(Fecha.Trim, 4)
        Mes = Derecha(Izquierda(Fecha.Trim, 5), 2)
        FechaAmericana = Mes & "/" & Dia & "/" & Ano
    End Function
#End Region
#Region "FechaBritanica"
    Public Shared Function FechaBritanica(ByVal Fecha As String) As String
        'METHOD SUMMARY
        'Convierte un texto conteniendo una fecha en formato MM/DD/YYYY a DD/MM/YYYY
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a formatear
        'END PARAMETERS SUMMARY
        If Fecha.Trim.Length <> 10 Then
            FechaBritanica = Fecha
            Exit Function
        End If
        Dim Dia As String, Mes As String, Ano As String
        Mes = Fecha.Trim.Substring(0, 2)
        Ano = Derecha(Fecha.Trim, 4)
        Dia = Derecha(Izquierda(Fecha.Trim, 5), 2)
        FechaBritanica = Dia & "/" & Mes & "/" & Ano
    End Function
#End Region
#Region "Extraer"
    Public Shared Function Extraer(ByVal Cadena As String, ByVal Separador As String, ByVal Posicion As Integer) As String
        'METHOD SUMMARY
        'Busca y extrae una posicion de una cadena de texto.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Cadena=Cadena  a extraer el valor
        'Separador=Separador de elementos
        'Posicion=Posicion a extraer
        'END PARAMETERS SUMMARY
        Dim Largo As Integer
        Dim Caracter As String
        Dim Separadores As Integer
        Dim Pinicio As Integer
        Dim Tresul As String = ""
        Dim auxcant As Integer
        Dim N As Integer
        Largo = Cadena.Trim.Length
        Pinicio = 1
        auxcant = 0
        For N = 0 To (Largo - 1)
            Caracter = Izquierda(Derecha(Cadena.Trim, (Largo - N)), 1)
            auxcant = auxcant + 1
            If Caracter.Trim = Separador.Trim Then
                Separadores = Separadores + 1
                If Separadores = Posicion Then
                    Tresul = Cadena.Substring(Pinicio - 1, auxcant - 1)
                    N = Largo - 1
                Else
                    Pinicio = N + 2
                End If
                auxcant = 0
            End If
        Next
        Extraer = Tresul
    End Function
#End Region
#Region "ImporteLetrasEspañol"
    Public Shared Function ImporteLetrasEspanol(ByVal Monto As Double, ByVal Moneda As String) As String
        'METHOD SUMMARY
        'Pasandole un numero la funcion devuelve el importe en letras.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Monto=Importe
        'Moneda=Simbolo monetario
        'END PARAMETERS SUMMARY
        Dim Txt As String = ""
        Dim Im As String
        Dim Centavos As String
        Dim U As Integer
        Dim D As Integer
        Dim C As Integer
        Dim M As Integer
        Dim A As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim Z As Integer
        'Formatea la variable monto.
        Im = Microsoft.VisualBasic.Format(Monto, "0.00")
        'Calcula el largo entero.
        A = Im.Length - 3
        'Genera los centavos.
        Centavos = "CON " & Derecha(Im, 2) & " CENTAVOS."
        'Cifras de 1 a 9.
        If A = 1 Then
            U = CEntero(Microsoft.VisualBasic.Mid$(Im, 1, 1))
        End If
        'Cifras de 1 a 99.
        If A = 2 Then
            U = CEntero(Microsoft.VisualBasic.Mid$(Im, 2, 1))
            D = CEntero(Microsoft.VisualBasic.Mid$(Im, 1, 1))
        End If
        'Cifras de 1 a 999.
        If A = 3 Then
            U = CEntero(Microsoft.VisualBasic.Mid$(Im, 3, 1))
            D = CEntero(Microsoft.VisualBasic.Mid$(Im, 2, 1))
            C = CEntero(Microsoft.VisualBasic.Mid$(Im, 1, 1))
        End If
        'Cifras de 1 a 9999.
        If A = 4 Then
            U = CEntero(Microsoft.VisualBasic.Mid$(Im, 4, 1))
            D = CEntero(Microsoft.VisualBasic.Mid$(Im, 3, 1))
            C = CEntero(Microsoft.VisualBasic.Mid$(Im, 2, 1))
            M = CEntero(Microsoft.VisualBasic.Mid$(Im, 1, 1))
        End If
        'Cifras de 1 a 99999.
        If A = 5 Then
            U = CEntero(Microsoft.VisualBasic.Mid$(Im, 5, 1))
            D = CEntero(Microsoft.VisualBasic.Mid$(Im, 4, 1))
            C = CEntero(Microsoft.VisualBasic.Mid$(Im, 3, 1))
            M = CEntero(Microsoft.VisualBasic.Mid$(Im, 2, 1))
            X = CEntero(Microsoft.VisualBasic.Mid$(Im, 1, 1))
        End If
        'Cifras de 1 a 999999.
        If A = 6 Then
            U = CEntero(Microsoft.VisualBasic.Mid$(Im, 6, 1))
            D = CEntero(Microsoft.VisualBasic.Mid$(Im, 5, 1))
            C = CEntero(Microsoft.VisualBasic.Mid$(Im, 4, 1))
            M = CEntero(Microsoft.VisualBasic.Mid$(Im, 3, 1))
            X = CEntero(Microsoft.VisualBasic.Mid$(Im, 2, 1))
            Y = CEntero(Microsoft.VisualBasic.Mid$(Im, 1, 1))
        End If
        'Cifras de 1 a 9999999.
        If A = 7 Then
            U = CEntero(Microsoft.VisualBasic.Mid$(Im, 7, 1))
            D = CEntero(Microsoft.VisualBasic.Mid$(Im, 6, 1))
            C = CEntero(Microsoft.VisualBasic.Mid$(Im, 5, 1))
            M = CEntero(Microsoft.VisualBasic.Mid$(Im, 4, 1))
            X = CEntero(Microsoft.VisualBasic.Mid$(Im, 3, 1))
            Y = CEntero(Microsoft.VisualBasic.Mid$(Im, 2, 1))
            Z = CEntero(Microsoft.VisualBasic.Mid$(Im, 1, 1))
        End If
        'Genera los millones.
        If Z = 1 Then Txt = "UN MILLON "
        If Z = 2 Then Txt = "DOS MILLONES "
        If Z = 3 Then Txt = "TRES MILLONES "
        If Z = 4 Then Txt = "CUATRO MILLONES "
        If Z = 5 Then Txt = "CINCO MILLONES "
        If Z = 6 Then Txt = "SEIS MILLONES "
        If Z = 7 Then Txt = "SIETE MILLONES "
        If Z = 8 Then Txt = "OCHO MILLONES "
        If Z = 9 Then Txt = "NUEVE MILLONES "
        'Genera los cientos redondos.
        If Y = 1 And M = 0 And X = 0 Then Txt = Txt & "CIEN MIL "
        If Y = 2 And M = 0 And X = 0 Then Txt = Txt & "DOCIENTOS MIL "
        If Y = 3 And M = 0 And X = 0 Then Txt = Txt & "TRECIENTOS MIL "
        If Y = 4 And M = 0 And X = 0 Then Txt = Txt & "CUATROCIENTOS MIL "
        If Y = 5 And M = 0 And X = 0 Then Txt = Txt & "QUINIENTOS MIL "
        If Y = 6 And M = 0 And X = 0 Then Txt = Txt & "SEISCIENTOS MIL "
        If Y = 7 And M = 0 And X = 0 Then Txt = Txt & "SETECIENTOS MIL "
        If Y = 8 And M = 0 And X = 0 Then Txt = Txt & "OCHOCIENTOS MIL "
        If Y = 9 And M = 0 And X = 0 Then Txt = Txt & "NOVECIENTOS MIL "
        'Genera los cientos parciales.
        If Y = 1 And (M > 0 Or X > 0) Then Txt = Txt & "CIENTO "
        If Y = 2 And (M > 0 Or X > 0) Then Txt = Txt & "DOCIENTOS "
        If Y = 3 And (M > 0 Or X > 0) Then Txt = Txt & "TRECIENTOS "
        If Y = 4 And (M > 0 Or X > 0) Then Txt = Txt & "CUATROCIENTOS "
        If Y = 5 And (M > 0 Or X > 0) Then Txt = Txt & "QUINIENTOS "
        If Y = 6 And (M > 0 Or X > 0) Then Txt = Txt & "SEISCIENTOS "
        If Y = 7 And (M > 0 Or X > 0) Then Txt = Txt & "SETECIENTOS "
        If Y = 8 And (M > 0 Or X > 0) Then Txt = Txt & "OCHOCIENTOS "
        If Y = 9 And (M > 0 Or X > 0) Then Txt = Txt & "NOVECIENTOS "
        'Genera las decenas de mil redondas.
        If X = 1 And M = 0 Then Txt = Txt & "DIEZ MIL "
        If X = 2 And M = 0 Then Txt = Txt & "VEINTE MIL "
        If X = 3 And M = 0 Then Txt = Txt & "TREINTA MIL "
        If X = 4 And M = 0 Then Txt = Txt & "CUARENTA MIL "
        If X = 5 And M = 0 Then Txt = Txt & "CINCUENTA MIL "
        If X = 6 And M = 0 Then Txt = Txt & "SESENTA MIL "
        If X = 7 And M = 0 Then Txt = Txt & "SETENTA MIL "
        If X = 8 And M = 0 Then Txt = Txt & "OCHENTA MIL "
        If X = 9 And M = 0 Then Txt = Txt & "NOVENTA MIL "
        'Genera las decenas de mil parciales.
        If X = 1 And M > 5 Then Txt = Txt & "DIEZ Y "
        If X = 2 And M > 0 Then Txt = Txt & "VEINTI"
        If X = 3 And M > 0 Then Txt = Txt & "TREINTA Y "
        If X = 4 And M > 0 Then Txt = Txt & "CUARENTA Y "
        If X = 5 And M > 0 Then Txt = Txt & "CINCUENTA Y "
        If X = 6 And M > 0 Then Txt = Txt & "SESENTA Y "
        If X = 7 And M > 0 Then Txt = Txt & "SETENTA Y "
        If X = 8 And M > 0 Then Txt = Txt & "OCHENTA Y "
        If X = 9 And M > 0 Then Txt = Txt & "NOVENTA Y "
        If M = 1 And X = 1 Then Txt = Txt & "ONCE MIL "
        If M = 2 And X = 1 Then Txt = Txt & "DOCE MIL "
        If M = 3 And X = 1 Then Txt = Txt & "TRECE MIL "
        If M = 4 And X = 1 Then Txt = Txt & "CATORCE MIL "
        If M = 5 And X = 1 Then Txt = Txt & "QUINCE MIL "
        If M = 6 And X = 1 Then Txt = Txt & "SEIS MIL "
        If M = 7 And X = 1 Then Txt = Txt & "SIETE MIL "
        If M = 8 And X = 1 Then Txt = Txt & "OCHO MIL "
        If M = 9 And X = 1 Then Txt = Txt & "NUEVE MIL "
        If M = 1 And X <> 1 Then Txt = Txt & "UN MIL "
        If M = 2 And X <> 1 Then Txt = Txt & "DOS MIL "
        If M = 3 And X <> 1 Then Txt = Txt & "TRES MIL "
        If M = 4 And X <> 1 Then Txt = Txt & "CUATRO MIL "
        If M = 5 And X <> 1 Then Txt = Txt & "CINCO MIL "
        If M = 6 And X <> 1 Then Txt = Txt & "SEIS MIL "
        If M = 7 And X <> 1 Then Txt = Txt & "SIETE MIL "
        If M = 8 And X <> 1 Then Txt = Txt & "OCHO MIL "
        If M = 9 And X <> 1 Then Txt = Txt & "NUEVE MIL "
        If C = 1 And (D > 0 Or U > 0) Then Txt = Txt & "CIENTO "
        If C = 1 And (D = 0 And U = 0) Then Txt = Txt & "CIEN "
        If C = 2 Then Txt = Txt & "DOCIENTOS "
        If C = 3 Then Txt = Txt & "TRECIENTOS "
        If C = 4 Then Txt = Txt & "CUATROCIENTOS "
        If C = 5 Then Txt = Txt & "QUINIENTOS "
        If C = 6 Then Txt = Txt & "SEISCIENTOS "
        If C = 7 Then Txt = Txt & "SETECIENTOS "
        If C = 8 Then Txt = Txt & "OCHOCIENTOS "
        If C = 9 Then Txt = Txt & "NOVECIENTOS "
        If D = 1 And U = 0 Then Txt = Txt & "DIEZ "
        If D = 1 And U = 1 Then Txt = Txt & "ONCE "
        If D = 1 And U = 2 Then Txt = Txt & "DOCE "
        If D = 1 And U = 3 Then Txt = Txt & "TRECE "
        If D = 1 And U = 4 Then Txt = Txt & "CATORCE "
        If D = 1 And U = 5 Then Txt = Txt & "QUINCE "
        If D = 1 And U = 6 Then Txt = Txt & "DIEZ Y SEIS "
        If D = 1 And U = 7 Then Txt = Txt & "DIEZ Y SIETE "
        If D = 1 And U = 8 Then Txt = Txt & "DIEZ Y OCHO "
        If D = 1 And U = 9 Then Txt = Txt & "DIEZ Y NUEVE "
        If D = 2 And U = 0 Then Txt = Txt & "VEINTE "
        If D = 3 And U = 0 Then Txt = Txt & "TREINTA "
        If D = 4 And U = 0 Then Txt = Txt & "CUARENTA "
        If D = 5 And U = 0 Then Txt = Txt & "CINCUENTA "
        If D = 6 And U = 0 Then Txt = Txt & "SESENTA "
        If D = 7 And U = 0 Then Txt = Txt & "SETENTA "
        If D = 8 And U = 0 Then Txt = Txt & "OCHENTA "
        If D = 9 And U = 0 Then Txt = Txt & "NOVENTA "
        If D = 2 And U > 0 Then Txt = Txt & "VEINTI"
        If D = 3 And U > 0 Then Txt = Txt & "TREINTA Y "
        If D = 4 And U > 0 Then Txt = Txt & "CUARENTA Y "
        If D = 5 And U > 0 Then Txt = Txt & "CINCUENTA Y "
        If D = 6 And U > 0 Then Txt = Txt & "SESENTA Y "
        If D = 7 And U > 0 Then Txt = Txt & "SETENTA Y "
        If D = 8 And U > 0 Then Txt = Txt & "OCHENTA Y "
        If D = 9 And U > 0 Then Txt = Txt & "NOVENTA Y "
        If U = 1 And D <> 1 Then Txt = Txt & "UNO "
        If U = 2 And D <> 1 Then Txt = Txt & "DOS "
        If U = 3 And D <> 1 Then Txt = Txt & "TRES "
        If U = 4 And D <> 1 Then Txt = Txt & "CUATRO "
        If U = 5 And D <> 1 Then Txt = Txt & "CINCO "
        If U = 6 And D <> 1 Then Txt = Txt & "SEIS "
        If U = 7 And D <> 1 Then Txt = Txt & "SIETE "
        If U = 8 And D <> 1 Then Txt = Txt & "OCHO "
        If U = 9 And D <> 1 Then Txt = Txt & "NUEVE "
        If U = 0 And D = 0 And C = 0 And M = 0 And X = 0 And Y = 0 And Z = 0 Then Txt = "CERO "
        If Centavos <> "CON 00 CENTAVOS." Then
            ImporteLetrasEspanol = Moneda & " " & Txt & Centavos & "-"
        Else
            ImporteLetrasEspanol = Moneda & " " & Txt & ".-"
        End If
        ImporteLetrasEspanol = ImporteLetrasEspanol.Trim.ToLower
    End Function
#End Region
#Region "Clogico"
    Public Shared Function CLogico(ByVal Valor As String) As Boolean
        'METHOD SUMMARY
        'Convierte a boolean una cadena de texto.
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Valor=Texto a convertir
        'END PARAMETERS SUMMARY
        If Valor.Trim = "0" Or Valor.Trim.ToLower = "off" Or Valor.Trim.ToUpper = "N" Or Valor.Trim.ToUpper = "FALSE" Then Return False
        If Valor.Trim = "1" Or Valor.Trim.ToLower = "on" Or Valor.Trim.ToUpper = "Y" Or Valor.Trim.ToUpper = "TRUE" Or Valor.Trim.ToUpper = "S" Then Return True
        Try
            CLogico = Boolean.Parse(Valor)
        Catch
            Return False
        End Try
    End Function
#End Region
#Region "Cdoble"
    Public Shared Function CDoble(ByVal Numero As String) As Double
        'METHOD SUMMARY
        'Convierte a Doble una cadena de texto.
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Numero=Texto a convertir
        'END PARAMETERS SUMMARY
        If Numero Is Nothing Then Return 0
        Numero = Numero.Trim
        If Numero.Trim.Length = 0 Then Return 0
        Try
            Dim entero As New System.Text.StringBuilder
            Dim decimales As New System.Text.StringBuilder
            Dim Divisor As New System.Text.StringBuilder
            Divisor.Append("1")
            Dim caracter As String
            Dim bandera As Boolean = False
            Dim n As Integer
            For n = 0 To Numero.Trim.Length - 1
                caracter = Izquierda(Derecha(Numero, Numero.Length - n), 1)
                If Not EsNumero(caracter) And caracter <> "-" And caracter <> "+" Then
                    bandera = True
                Else
                    If Not bandera Then
                        entero.Append(caracter)
                    Else
                        decimales.Append(caracter)
                        Divisor.Append("0")
                    End If
                End If
            Next
            CDoble = Double.Parse(entero.ToString)
            If decimales.ToString.Trim.Length <> 0 Then
                If CDoble >= 0 Then
                    CDoble = CDoble + (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                Else
                    CDoble = CDoble - (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                End If
            End If
            If entero.ToString = "-0" Then CDoble = CDoble * -1
        Catch
            Return 0
        End Try
    End Function
#End Region
#Region "Cdecimal"
    Public Shared Function CDecimal(ByVal Numero As String) As Double
        'METHOD SUMMARY
        'Convierte a Decimal una cadena de texto.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Numero=Texto a convertir
        'END PARAMETERS SUMMARY
        If Numero.Trim.Length = 0 Then Return 0
        Try
            Dim entero As New System.Text.StringBuilder
            Dim decimales As New System.Text.StringBuilder
            Dim Divisor As New System.Text.StringBuilder
            Divisor.Append("1")
            Dim caracter As String
            Dim bandera As Boolean = False
            Dim n As Integer
            For n = 0 To Numero.Trim.Length - 1
                caracter = Izquierda(Derecha(Numero, Numero.Length - n), 1)
                If Not EsNumero(caracter) And caracter <> "-" And caracter <> "+" Then
                    bandera = True
                Else
                    If Not bandera Then
                        entero.Append(caracter)
                    Else
                        decimales.Append(caracter)
                        Divisor.Append("0")
                    End If
                End If
            Next
            CDecimal = Decimal.Parse(entero.ToString)
            If decimales.ToString.Trim.Length <> 0 Then
                If CDecimal >= 0 Then
                    CDecimal = CDecimal + (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                Else
                    CDecimal = CDecimal - (Long.Parse(decimales.ToString) / Long.Parse(Divisor.ToString))
                End If
            End If
            If entero.ToString = "-0" Then CDecimal = CDecimal * -1
        Catch
            Return 0
        End Try
    End Function
#End Region
#Region "Clargo"
    Public Shared Function Clargo(ByVal Numero As String) As Long
        'METHOD SUMMARY
        'Convierte a long una cadena de texto.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Numero=Texto a convertir
        'END PARAMETERS SUMMARY
        Dim SeparadorDecimalSistema As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        If SeparadorDecimalSistema <> "." Then
            Numero = Numero.Replace(".", SeparadorDecimalSistema)
        End If
        Try
            Clargo = Long.Parse(Numero)
        Catch
            Return 0
        End Try
    End Function
#End Region
#Region "TxtToFecha"
    Public Shared Function TxtToFecha(ByVal Fecha As String) As DateTime
        'METHOD SUMMARY
        'Convierte un texto a una fecha. El texto debe estar formateado en DD/MM/YYYY
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Texto a convertir a fecha
        'END PARAMETERS SUMMARY
        Try
            Dim Mifecha As New DateTime(CEntero(Fecha.Substring(6, 4)), CEntero(Fecha.Substring(3, 2)), CEntero(Fecha.Substring(0, 2)))
            Return Mifecha
        Catch
            Return System.DateTime.Now
        End Try
    End Function
#End Region
#Region "FinDeMes"
    Public Shared Function FinDeMes(ByVal Mes As Integer, ByVal Ano As Integer) As String
        FinDeMes = ""
        'METHOD SUMMARY
        'Devuelve la ultima fecha de un mes en formato de texto DD/MM/YYYY
        'END METHOD SUMMARY
        Dim TxtMes As String
        If Mes.ToString.Trim.Length = 1 Then
            TxtMes = "0" + Mes.ToString.Trim
        Else
            TxtMes = Mes.ToString.Trim
        End If
        Select Case Mes
            Case 1
                FinDeMes = "31/" + TxtMes + "/" + Ano.ToString.Trim
            Case 2
                If TxtToFecha("29/02/" + Ano.ToString) = System.DateTime.Now Then
                    FinDeMes = "28/" + TxtMes + "/" + Ano.ToString.Trim
                Else
                    FinDeMes = "29/" + TxtMes + "/" + Ano.ToString.Trim
                End If
            Case 3
                FinDeMes = "31/" + TxtMes + "/" + Ano.ToString.Trim
            Case 4
                FinDeMes = "30/" + TxtMes + "/" + Ano.ToString.Trim
            Case 5
                FinDeMes = "31/" + TxtMes + "/" + Ano.ToString.Trim
            Case 6
                FinDeMes = "30/" + TxtMes + "/" + Ano.ToString.Trim
            Case 7
                FinDeMes = "31/" + TxtMes + "/" + Ano.ToString.Trim
            Case 8
                FinDeMes = "31/" + TxtMes + "/" + Ano.ToString.Trim
            Case 9
                FinDeMes = "30/" + TxtMes + "/" + Ano.ToString.Trim
            Case 10
                FinDeMes = "31/" + TxtMes + "/" + Ano.ToString.Trim
            Case 11
                FinDeMes = "30/" + TxtMes + "/" + Ano.ToString.Trim
            Case 12
                FinDeMes = "31/" + TxtMes + "/" + Ano.ToString.Trim
        End Select
    End Function
#End Region
#Region "PrincipioDeMes"
    Public Shared Function PrincipioDeMes(ByVal Mes As Integer, ByVal Ano As Integer) As String
        'METHOD SUMMARY
        'Devuelve la primera fecha de un mes en formato de texto DD/MM/YYYY
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Mes=Mes
        'Ano=Año
        'END PARAMETERS SUMMARY
        Dim TxtMes As String
        If Mes.ToString.Trim.Length = 1 Then
            TxtMes = "0" + Mes.ToString.Trim
        Else
            TxtMes = Mes.ToString.Trim
        End If
        PrincipioDeMes = "01/" + TxtMes + "/" + Ano.ToString.Trim
    End Function
#End Region
#Region "TiempoEntreFechas"
#End Region
#Region "SumarDiasAFechas"
    Public Shared Function SumarDiasAFechas(ByVal Fecha As String, ByVal Dias As Integer) As String
        'METHOD SUMMARY
        'Suma o resta dias a una fecha y devuelve el resultado como un texto en formato DD/MM/YYYY
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a ser sumada los dias
        'Dias=Dias a sumar
        'END PARAMETERS SUMMARY
        If Not ValidoTxtFecha(Fecha) Then
            SumarDiasAFechas = "**/**/****"
            Exit Function
        End If
        Dim Dia As String = Fecha.Substring(0, 2)
        Dim Mes As String = Fecha.Substring(3, 2)
        Dim Ano As String = Fecha.Substring(6, 4)
        Dim FFechaAux As New System.DateTime(CEntero(Ano), CEntero(Mes), CEntero(Dia))
        Dim ffecha As DateTime = FFechaAux.AddDays(Dias)
        ffecha.AddDays(Dias)
        Dim Diar As String = (ffecha.Day + 100).ToString.Substring(1, 2)
        Dim mesr As String = (ffecha.Month + 100).ToString.Substring(1, 2)
        Dim anor As String = (ffecha.Year).ToString
        SumarDiasAFechas = Diar & "/" & mesr & "/" & anor
    End Function
#End Region
#Region "ConCaracteresNovalidos"
    Public Shared Function ConCaracteresNovalidos(ByVal Texto As String) As Boolean
        'METHOD SUMMARY
        'Valida si una cadena contiene caracteres no validos para pasar en queriesStrings.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Texto=Cadena de texto a analizar
        'END PARAMETERS SUMMARY
        Dim n As Integer
        For n = 0 To Texto.Trim.Length - 1
            Select Case Texto.Substring(n, 1)
                Case Is = ">"
                    Return True
                Case Is = "<"
                    Return True
                Case Is = """"
                    Return True
                Case Is = "?"
                    Return True
            End Select
        Next
    End Function
#End Region
#Region "MailValido"
    Public Shared Function MailValido(ByVal Mail As String) As Boolean
        'METHOD SUMMARY
        'Valida si un mail esta escripto en un formato valido.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Mail=Mail a validar
        'END PARAMETERS SUMMARY
        If Mail.Trim.Length <> 0 Then
            Dim Expresion As New Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")
            Dim Resultado As Match
            Resultado = Expresion.Match(Mail)
            Return Resultado.Success
        Else
            Return True
        End If
    End Function
#End Region
#Region "UrlValida"
    Public Shared Function UrlValida(ByVal Url As String) As Boolean
        'METHOD SUMMARY
        'Verifica si una direccion esta escrita en un formato valido.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Url=Url a validar
        'END PARAMETERS SUMMARY
        If Url.Trim.Length <> 0 Then
            Dim Expresion As New Regex("http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?")
            Dim Resultado As Match
            Resultado = Expresion.Match(Url)
            Return Resultado.Success
        Else
            Return True
        End If
    End Function
#End Region
#Region "ValidarRangoMeses"
    Public Shared Function ValidarRangoMeses(ByVal Mesd As Integer, ByVal Añod As Integer, ByVal Mesh As Integer, ByVal Añoh As Integer) As Boolean
        'METHOD SUMMARY
        'Valida si un mes es mayor o igual a otro.
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Mesd=Mes Inicial
        'Añod=Año Inicial
        'Mesh=Mes Final
        'Añoh=Año Final
        'END PARAMETERS SUMMARY
        Dim Fechad As String = PrincipioDeMes(Mesd, Añod)
        Dim Fechah As String = FinDeMes(Mesh, Añoh)
        If TiempoEntreFechas(Fechad, Fechah) < 0 Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region
#Region "ExtencionDeArchivo"
    Public Shared Function ExtencionDeArchivo(ByVal Archivo As String) As String
        'METHOD SUMMARY
        'Devuelve la extension del un archivo
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Archivo=Archivo a extrarer la extencion
        'END PARAMETERS SUMMARY
        ExtencionDeArchivo = ""
        Archivo = Archivo.Trim
        Dim largo As Integer = Archivo.Length
        Dim N As Integer
        For N = largo - 1 To 0 Step -1
            If Archivo.Substring(N, 1) = "." Then
                ExtencionDeArchivo = Archivo.Substring(N, largo - N)
                Exit For
            End If
        Next
    End Function
#End Region
#Region "NombreDeArchivo"
    Public Shared Function NombreDeArchivo(ByVal Ruta As String) As String
        'METHOD SUMMARY
        'Devuelve solo el nomnre de archivo de una cadena de caracteres que incluye el path.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Ruta=Nombre de archivo junto con su ruta
        'END PARAMETERS SUMMARY
        Dim N As Integer
        NombreDeArchivo = ""
        For n = Ruta.Trim.Length - 1 To 0 Step -1
            If Ruta.Trim.Substring(n, 1) = "\" Or Ruta.Trim.Substring(n, 1) = "/" Then
                Exit For
            Else
                NombreDeArchivo = Ruta.Trim.Trim.Substring(n, 1) & NombreDeArchivo
            End If
        Next
        'Dim N As Integer, Caracter As String = ""
        'For N = Ruta.Trim.Length To 1 Step -1
        '    Caracter = Derecha(Izquierda(Ruta, N), 1)
        '    If Caracter = "/" Or Caracter = "\" Then
        '        Caracter = Derecha(Ruta, Ruta.Trim.Length - N)
        '        Exit For
        '    End If
        'Next
        'NombreDeArchivo = Caracter
    End Function
#End Region
#Region "NumeroComprobanteValido"
    Public Shared Function NumeroComprobanteValido(ByVal Numero As String, ByVal Expresion As String) As Boolean
        'METHOD SUMMARY
        'Valida una exprecion si es valida
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Numero=Numero de comprobante
        'Expresion=Expresion regular
        'END PARAMETERS SUMMARY
        If Expresion.Trim.Length <> 0 Then
            Dim LExpresion As New Regex(Expresion)
            Dim Resultado As Match
            Resultado = LExpresion.Match(Numero)
            Return Resultado.Success
        Else
            Return True
        End If
    End Function
#End Region
#Region "NombreArchivoTemporal"
    Public Shared Function NombreArchivoTemporal(ByVal Extension As String) As String
        'METHOD SUMMARY
        'Genera un nombre random de un archivo temporal.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Extension=Extension de archivo a buscar.
        'END PARAMETERS SUMMARY
        Dim r As Random = New Random
        Dim NumeroRandom As String = r.Next(1, 999999).ToString
        NombreArchivoTemporal = Date.Today.Year.ToString & Date.Today.Month.ToString & Date.Today.Day.ToString & Date.Today.Minute.ToString & Date.Today.Second.ToString & NumeroRandom.ToString & "." & Extension
    End Function
#End Region
#Region "SinCaracteresExtendidos"
    Public Shared Function SinCaracteresExtendidos(ByVal Texto As String) As String
        'METHOD SUMMARY
        'Reemplaza algunos caracteres extendidos en una cadena de texto.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Texto=Cadena de texto
        'END PARAMETERS SUMMARY
        If Texto Is Nothing Then Return ""
        Texto = Texto.Replace("á", "a")
        Texto = Texto.Replace("é", "e")
        Texto = Texto.Replace("í", "i")
        Texto = Texto.Replace("ó", "o")
        Texto = Texto.Replace("ú", "u")
        Texto = Texto.Replace("ñ", "n")
        Return Texto
    End Function
#End Region
#Region "LetraCapital"
    Public Shared Function LetraCapital(ByVal Texto As String) As String
        'METHOD SUMMARY
        'Devuelve una cadena de texto en formato de letra capital.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Texto=Texto a convertir a letra capital
        'END PARAMETERS SUMMARY
        If Texto.Trim.Length = 0 Then
            Return Texto
        Else
            Return Texto.Trim.Substring(0, 1).ToUpper & Texto.Trim.Substring(1, Texto.Trim.Length - 1).ToLower
        End If
    End Function
#End Region
#Region "DameEnteros"
    Public Shared Function DameEnteros(ByVal Valor As String) As Long
        'METHOD SUMMARY
        'Devuelve la parte entera de un numero doble.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Valor=Numero doble
        'END PARAMETERS SUMMARY
        Dim N As Integer
        Dim Entero As String = ""
        Dim Cadena As String = Valor.Trim
        For N = 0 To Cadena.Length - 1
            If Cadena.Substring(N, 1) = "." Or Cadena.Substring(N, 1) = "," Then
                Exit For
            End If
            Entero = Entero & Cadena.Substring(N, 1)
        Next
        Return Clargo(Entero)
    End Function
#End Region
#Region "DameDecimales"
    Public Shared Function DameDecimales(ByVal Valor As String) As Long
        'METHOD SUMMARY
        'Devuelve la parte decimal de un numero doble.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Valor=Valor a analizar
        'END PARAMETERS SUMMARY
        Dim N As Integer
        Dim Decimales As String = ""
        Dim Bandera As Boolean = False
        Dim Cadena As String = Valor.Trim
        For N = 0 To Cadena.Length - 1
            If Bandera Then Decimales = Decimales & Cadena.Substring(N, 1)
            If Cadena.Substring(N, 1) = "." Or Cadena.Substring(N, 1) = "," Then
                Bandera = True
            End If
        Next
        Return Clargo(Decimales)
    End Function
#End Region
#Region "CompletoConCaracteres"
    Public Shared Function CompletoConCaracteres(ByVal Texto As String, ByVal Caracter As String, ByVal Izquierda As Boolean, ByVal LargoResultado As Integer) As String
        'METHOD SUMMARY
        'Rellena una cadena de caracteres.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Texto=Cadena de texto a completar
        'Caracter=caracter de relleno
        'Izquierda=del lado izquierdo
        'LargoResultado=Largo del resultado
        'END PARAMETERS SUMMARY
        Dim Largo As Integer = Texto.Trim.Length
        Dim n As Integer
        If Largo < LargoResultado Then
            Dim Cadena As New System.Text.StringBuilder
            For n = 1 To LargoResultado - Largo
                With Cadena
                    .Append(Caracter)
                End With
            Next
            If Izquierda Then
                Texto = Texto & Cadena.ToString
            Else
                Texto = Cadena.ToString & Texto
            End If
            Return Texto
        Else
            Return Texto
        End If
    End Function
#End Region
#Region "ImportesIguales"
    Public Shared Function ImportesIguales(ByVal Valor1 As Double, ByVal Valor2 As Double, ByVal DiferenciaAceptable As Double) As Boolean
        'METHOD SUMMARY
        'Redondea 2 numeros y devuelve si son iguales considerando un margen de diferencia.
        '
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Valor1=Valor1
        'Valor2=Valor2
        'DiferenciaAceptable=Margen de diferencia aceptada
        'END PARAMETERS SUMMARY
        Valor1 = Math.Round(Valor1, 2)
        Valor2 = Math.Round(Valor2, 2)
        If Valor1 <> Valor2 Then
            If Valor1 > Valor2 Then
                If (Valor1 - Valor2) < DiferenciaAceptable Then
                    Return False
                End If
            Else
                If (Valor2 - Valor1) < DiferenciaAceptable Then
                    Return False
                End If
            End If
        End If
        Return True
    End Function
#End Region
#Region "ValidoTxtHora"
    Public Shared Function ValidoTxtHora(ByVal Hora As String) As Boolean
        'METHOD SUMMARY
        'Valida si una hora esta en formato HH:MM:SS
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Hora=Texto conteniendo una hora en formato HH:MM:SS
        'END PARAMETERS SUMMARY
        If Hora.Trim.Length <> 8 Then
            Return False
        End If
        Dim horas As Integer = CEntero(Hora.Substring(0, 2))
        Dim minutos As Integer = CEntero(Hora.Substring(3, 2))
        Dim segundos As Integer = CEntero(Hora.Substring(6, 2))
        If minutos > 60 Or minutos < 0 Then Return False
        If segundos > 60 Or segundos < 0 Then Return False
        If horas > 23 Or horas < 0 Then Return False
        Try
            Dim HoraDate As New System.DateTime(1971, 6, 7, horas, minutos, segundos)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region
#Region "TiempoEnSegundos"
    Public Shared Function TiempoEnSegundos(ByVal Tiempo As String) As Integer
        'METHOD SUMMARY
        'Devuelve una hora en segundos
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Tiempo=String que representa una hora
        'END PARAMETERS SUMMARY
        If Not ValidoTxtHora(Tiempo) Then Return 0
        Dim hora As Integer
        Dim minuto As Integer
        Dim segundo As Integer
        hora = CEntero(Tiempo.Substring(0, 2))
        minuto = CEntero(Tiempo.Substring(3, 2))
        segundo = CEntero(Tiempo.Substring(6, 2))
        TiempoEnSegundos = segundo + (minuto * 60) + ((hora * 60) * 60)
    End Function
#End Region
#Region "SumarTiempos"
    Public Shared Function SumarTiempos(ByVal Tiempo1 As String, ByVal Tiempo2 As String) As String
        'METHOD SUMMARY
        'Suma 2 horas
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Tiempo1=Tiempo1
        'Tiempo2=Tiempo2
        'END PARAMETERS SUMMARY
        If Not ValidoTxtHora(Tiempo1) Then Return ""
        If Not ValidoTxtHora(Tiempo2) Then Return ""
        Dim segundos As Integer = 0
        Dim minutos As Integer = 0
        Dim horas As Integer = 0
        Dim seg1 As Integer = CEntero(Tiempo1.Substring(6, 2))
        Dim seg2 As Integer = CEntero(Tiempo2.Substring(6, 2))
        Dim min1 As Integer = CEntero(Tiempo1.Substring(3, 2))
        Dim min2 As Integer = CEntero(Tiempo2.Substring(3, 2))
        Dim hora1 As Integer = CEntero(Tiempo1.Substring(0, 2))
        Dim hora2 As Integer = CEntero(Tiempo2.Substring(0, 2))
        segundos = seg1 + seg2
        If segundos >= 60 Then
            While segundos >= 60
                segundos = segundos - 60
                minutos = minutos + 1
            End While
        End If
        minutos = minutos + min1 + min2
        If minutos >= 60 Then
            minutos = minutos - 60
            horas = horas + 1
        End If
        horas = horas + hora1 + hora2
        If horas.ToString.Trim.Length = 1 Then
            SumarTiempos = "0" & horas.ToString.Trim & ":"
        Else
            SumarTiempos = horas.ToString.Trim & ":"
        End If
        If minutos.ToString.Trim.Length = 1 Then
            SumarTiempos = SumarTiempos & "0" & minutos.ToString.Trim & ":"
        Else
            SumarTiempos = SumarTiempos & minutos.ToString.Trim & ":"
        End If
        If segundos.ToString.Trim.Length = 1 Then
            SumarTiempos = SumarTiempos & "0" & segundos.ToString.Trim
        Else
            SumarTiempos = SumarTiempos & segundos.ToString.Trim
        End If
    End Function
#End Region
#Region "RangoHorario"
    Public Shared Function RangoHorario(ByVal InicioRangoHora As String, ByVal FinRangoHora As String, ByVal Hora As String) As Boolean
        'METHOD SUMMARY
        'Valida si 2 horas estan dentro de un rango.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'InicioRangoHora=Hora de inicio
        'FinRangoHora=Hora Final
        'Hora=Hora a Evaluar
        'END PARAMETERS SUMMARY
        RangoHorario = False
        If Not ValidoTxtHora(InicioRangoHora) Then Return False
        If Not ValidoTxtHora(InicioRangoHora) Then Return False
        If Not ValidoTxtHora(InicioRangoHora) Then Return False
        Dim segundos As Integer = 0
        Dim minutos As Integer = 0
        Dim horas As Integer = 0
        Dim InicioRangoHora_s As Integer = CEntero(InicioRangoHora.Substring(6, 2))
        Dim InicioRangoHora_m As Integer = CEntero(InicioRangoHora.Substring(3, 2))
        Dim InicioRangoHora_h As Integer = CEntero(InicioRangoHora.Substring(0, 2))
        Dim FinRangoHora_s As Integer = CEntero(FinRangoHora.Substring(6, 2))
        Dim FinRangoHora_m As Integer = CEntero(FinRangoHora.Substring(3, 2))
        Dim FinRangoHora_h As Integer = CEntero(FinRangoHora.Substring(0, 2))
        Dim Hora_s As Integer = CEntero(Hora.Substring(6, 2))
        Dim Hora_m As Integer = CEntero(Hora.Substring(3, 2))
        Dim Hora_h As Integer = CEntero(Hora.Substring(0, 2))
        Do While True
            If InicioRangoHora_s = Hora_s And InicioRangoHora_m = Hora_m And InicioRangoHora_h = Hora_h Then
                RangoHorario = True
                Exit Do
            End If
            If InicioRangoHora_s = FinRangoHora_s And InicioRangoHora_m = FinRangoHora_m And InicioRangoHora_h = FinRangoHora_h Then
                Exit Do
            End If
            InicioRangoHora_s = InicioRangoHora_s + 1
            If InicioRangoHora_s >= 60 Then
                InicioRangoHora_s = 0
                InicioRangoHora_m = InicioRangoHora_m + 1
            End If
            If InicioRangoHora_m >= 60 Then
                InicioRangoHora_m = 0
                InicioRangoHora_h = InicioRangoHora_h + 1
            End If
            If InicioRangoHora_h >= 24 Then
                InicioRangoHora_h = 0
            End If
        Loop
    End Function
#End Region
#Region "ValidarCUIT"
    Public Shared Function ValidarCuit(ByVal Valor As String) As Boolean
        'METHOD SUMMARY
        'Valida si un numero cuit es valido.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Valor=Número de cuit a validar
        'END PARAMETERS SUMMARY
        Try
            Dim Expresion As New Regex("^[0-9]{2}-[0-9]{8}-[0-9]{1}$")
            Dim Cadena As String
            Dim n As Integer
            Dim Resultado As Match
            Resultado = Expresion.Match(Valor)
            If Not Resultado.Success Then
                Return False
            End If
            Cadena = Valor.Substring(0, 2) & Valor.Substring(3, 8) & Valor.Substring(12, 1)
            Dim MiVector(10) As String
            Dim Contador As Integer = 0
            For n = 10 To 0 Step -1
                MiVector(Contador) = Cadena.Substring(n, 1)
                Contador = Contador + 1
            Next
            Dim Suma As Double
            Dim Multi As Integer
            For n = 0 To 10
                If n + 1 < 8 Then
                    Multi = n + 1
                Else
                    Multi = (n + 1) - 6
                End If
                Suma = Suma + (CEntero(MiVector(n)) * Multi)
            Next
            Suma = Suma / 11
            If DameDecimales(Suma.ToString) <> 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region
#Region "AgregarMesesAFecha"
    Public Shared Function AgregarMesesAFecha(ByVal Fecha As String, ByVal Meses As Integer) As String
        'METHOD SUMMARY
        'Suma meses a una fecha.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a modificar
        'Meses=Meses
        'END PARAMETERS SUMMARY
        Try
            Dim Mifecha As New DateTime(CEntero(Fecha.Substring(6, 4)), CEntero(Fecha.Substring(3, 2)), CEntero(Fecha.Substring(0, 2)))
            Mifecha = Mifecha.AddMonths(Meses)
            Dim Dia As String, Mes As String, Ano As String
            Dia = Mifecha.Day.ToString
            Mes = Mifecha.Month.ToString
            Ano = Mifecha.Year.ToString
            If Dia.Length = 1 Then
                Dia = "0" & Dia
            End If
            If Mes.Length = 1 Then
                Mes = "0" & Mes
            End If
            AgregarMesesAFecha = Dia & "/" & Mes & "/" & Ano
        Catch
            Return TxtFechaHoy()
        End Try
    End Function
#End Region
#Region "Tabulaciones"
    Public Shared Function Tabulaciones(ByVal cantidad As Integer) As String
        'METHOD SUMMARY
        'Devuelve una cadena con x tabulaciones.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'cantidad=Cantidad de TAB a devolver
        'END PARAMETERS SUMMARY
        Dim n As Integer
        Tabulaciones = ""
        For n = 1 To cantidad
            Tabulaciones = Tabulaciones & Convert.ToChar(9)
        Next
    End Function
#End Region
#Region "AccesoDirectoWeb"
    Public Shared Function AccesoDirectoWeb(ByVal Imagen As String, ByVal TextoAlt As String, ByVal Link As String, ByVal Vuelta As String, ByVal Id As String, ByVal Valor As Integer) As String
        'METHOD SUMMARY
        'Arma un tag a href de Html.
        '
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Imagen=Imagen
        'TextoAlt=Alt del tag IMG
        'Link=Link Html
        'Vuelta=Devolucion
        'Id=Id
        'Valor=Valor
        'END PARAMETERS SUMMARY
        AccesoDirectoWeb = "<A href='" & Link
        If Vuelta.Trim.Length <> 0 Then
            AccesoDirectoWeb = AccesoDirectoWeb & "?Vuelta=" & Vuelta
            If Valor <> 0 Then
                AccesoDirectoWeb = AccesoDirectoWeb & "?Modi=" & Valor.ToString.Trim & "'"
            Else
                AccesoDirectoWeb = AccesoDirectoWeb & "'"
            End If
        Else
            AccesoDirectoWeb = AccesoDirectoWeb & "'"
        End If
        AccesoDirectoWeb = AccesoDirectoWeb & ">"
        AccesoDirectoWeb = AccesoDirectoWeb & "<IMG  alt='" & TextoAlt & "'" & " src='" & Imagen & "' border=0></A>"
    End Function
#End Region
#Region "DuplicarCaracter"
    Public Shared Function DuplicarCaracter(ByVal caracter As String, ByVal cantidad As Integer) As String
        'METHOD SUMMARY
        'Devuelve una cadena con x espacios.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'cantidad=Cantidad de espacios
        'END PARAMETERS SUMMARY
        Dim n As Integer
        DuplicarCaracter = ""
        For n = 0 To cantidad - 1
            DuplicarCaracter = DuplicarCaracter & caracter
        Next
    End Function
#End Region
#Region "Espacios"
    Public Shared Function Espacios(ByVal cantidad As Integer) As String
        'METHOD SUMMARY
        'Devuelve una cadena con x espacios.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'cantidad=Cantidad de espacios
        'END PARAMETERS SUMMARY
        Dim n As Integer
        Espacios = ""
        For n = 0 To cantidad - 1
            Espacios = Espacios & " "
        Next
    End Function
#End Region
#Region "FormatoNumeroArg"
    Public Shared Function FormatoNumeroArg(ByVal Valor As String) As String
        'METHOD SUMMARY
        'Formatea un numero con el separador decimal y de millares usados en Argentina.
        '
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Valor=Importe a formatear
        'END PARAMETERS SUMMARY
        If Valor = "" Then Return Valor
        Dim Condecimales As Boolean = False
        Dim n As Integer
        Dim cadena As String
        Dim Cola As String
        Dim Cuerpo As String
        Dim PosicionDecimal As Integer
        Dim Veces As Integer
        Dim cuerpoAux As String = ""
        Valor = Valor.Trim
        For n = 0 To Valor.Length - 1
            cadena = Valor.Substring(n, 1)
            If cadena.Equals(".") Or cadena.Equals(",") Then
                Condecimales = True
                PosicionDecimal = n
            End If
        Next
        If Condecimales Then
            Valor = Valor.Replace(".", ",")
            Cola = Valor.Substring(PosicionDecimal, Valor.Length - PosicionDecimal)
            Cuerpo = Valor.Substring(0, PosicionDecimal)
        Else
            Cola = ""
            Cuerpo = Valor
        End If
        Veces = 0
        For n = Cuerpo.Length - 1 To 0 Step -1
            Veces = Veces + 1
            cadena = Valor.Substring(n, 1)
            If Veces = 3 Then
                Veces = 0
                If n <> 0 Then
                    cuerpoAux = "." & cadena & cuerpoAux
                Else
                    cuerpoAux = cadena & cuerpoAux
                End If
            Else
                cuerpoAux = cadena & cuerpoAux
            End If
        Next
        Valor = cuerpoAux & Cola
        If Valor.Trim.Substring(Valor.Trim.Length - 1, 1) = "," Then
            Valor = Valor.Trim.Substring(0, Valor.Trim.Length - 1)
        End If
        Return Valor
    End Function
#End Region
#Region "CodificarQueryString"
    Public Shared Function CodificarQueryString(ByVal Text As String) As String
        'METHOD SUMMARY
        'Encrypta un Querystring
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Text=Texto a codificar
        'END PARAMETERS SUMMARY
        Dim oes As Encryption64 = New Encryption64
        Return oes.Encrypt(Text, "!#$aBJ?3")
    End Function

#End Region
#Region "DecodificarQueryString"
    Public Shared Function DecodificarQueryString(ByVal Text As String) As String
        'METHOD SUMMARY
        'DesEncrypta un string encrypytado con la funcion CodificarQueryString.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Text=texto a analizar
        'END PARAMETERS SUMMARY
        Text = Text.Replace(" ", "+")
        Dim oes As Encryption64 = New Encryption64
        Return oes.Decrypt(Text, "!#$aBJ?3")
    End Function
#End Region
#Region "FechaAtxt"
    Public Shared Function FechaAtxt(ByVal Fecha As Date) As String
        'METHOD SUMMARY
        'Transforma un a texto segun el formato DD/MM/YYYY.
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a convertir
        'END PARAMETERS SUMMARY
        Dim Dia As String, Mes As String, Ano As String
        Dia = Fecha.Day.ToString
        Mes = Fecha.Month.ToString
        Ano = Fecha.Year.ToString
        If Dia.Length = 1 Then
            Dia = "0" & Dia
        End If
        If Mes.Length = 1 Then
            Mes = "0" & Mes
        End If
        FechaAtxt = Dia & "/" & Mes & "/" & Ano
    End Function
#End Region
#Region "Clase Encryption64"
    Private Class Encryption64
        Private key() As Byte = {}
        Private IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Public Function Decrypt(ByVal stringToDecrypt As String, ByVal sEncryptionKey As String) As String
            'METHOD SUMMARY
            'Funcion usada por el DesEncryptador.
            'END METHOD SUMMARY
            'PARAMETERS SUMMARY
            'stringToDecrypt=Texto a desencryptar
            'sEncryptionKey=Clave de encriptacion
            'END PARAMETERS SUMMARY
            Dim inputByteArray(stringToDecrypt.Length) As Byte
            Try
                key = System.Text.Encoding.UTF8.GetBytes(Izquierda(sEncryptionKey, 8))
                Dim des As New DESCryptoServiceProvider
                inputByteArray = Convert.FromBase64String(stringToDecrypt)
                Dim ms As New MemoryStream
                Dim cs As New CryptoStream(ms, des.CreateDecryptor(key, IV), CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
                Return encoding.GetString(ms.ToArray())
            Catch e As Exception
                Return e.Message
            End Try
        End Function

        Public Function Encrypt(ByVal stringToEncrypt As String, ByVal SEncryptionKey As String) As String
            'METHOD SUMMARY
            'Funcion usada por el Encryptador.
            'END METHOD SUMMARY
            'PARAMETERS SUMMARY
            'stringToEncrypt=Valor a encryptar
            'SEncryptionKey=Clave de encriptacion
            'END PARAMETERS SUMMARY
            Try
                key = System.Text.Encoding.UTF8.GetBytes(Izquierda(SEncryptionKey, 8))
                Dim des As New DESCryptoServiceProvider
                Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(stringToEncrypt)
                Dim ms As New MemoryStream
                Dim cs As New CryptoStream(ms, des.CreateEncryptor(key, IV), CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                Return Convert.ToBase64String(ms.ToArray())
            Catch e As Exception
                Return e.Message
            End Try
        End Function
    End Class
#End Region
#Region "ContenidoLineaSinTags"
    Public Shared Function ContenidoLineaSinTags(ByVal texto As String) As String
        'METHOD SUMMARY
        'Extrae el valor de un Tag de Html.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'texto=Cadena de texto a analizar
        'END PARAMETERS SUMMARY
        Dim n As Integer
        Dim caracter As String
        Dim Capturando As Boolean = False
        ContenidoLineaSinTags = ""
        Dim Valor As String = ""
        For n = 0 To texto.Trim.Length - 1
            caracter = texto.Trim.Substring(n, 1)
            Select Case caracter
                Case "<"
                    Capturando = False
                Case ">"
                    Capturando = True
                Case Else
                    If Capturando Then
                        Valor = Valor & caracter
                    End If
            End Select
        Next
        Return Valor.Trim
    End Function
#End Region
#Region "EsNumero"
    Public Shared Function EsCaracterNumerico(ByVal caracter As String) As Boolean
        'METHOD SUMMARY
        'Valida si un caracter es numerico.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'caracter=Caracter para analizar
        'END PARAMETERS SUMMARY
        Select Case caracter
            Case "0"
                Return True
            Case "1"
                Return True
            Case "2"
                Return True
            Case "3"
                Return True
            Case "4"
                Return True
            Case "5"
                Return True
            Case "6"
                Return True
            Case "7"
                Return True
            Case "8"
                Return True
            Case "9"
                Return True
            Case Else
                Return False
        End Select
    End Function
#End Region
#Region "UltimoCaracter"
    Public Shared Function UltimoCaracter(ByVal Texto As String) As String
        'METHOD SUMMARY
        'Devuelve el ultimo carcater de una cadena.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Texto=Texto de entrada
        'END PARAMETERS SUMMARY
        If Texto.Trim.Length = 0 Then
            Return ""
        Else
            Return Texto.Trim.Substring(Texto.Trim.Length - 1, 1)
        End If
    End Function
#End Region
#Region "FechaEstaEnElRango"
    Public Shared Function FechaEstaEnElRango(ByVal Fecha As String, ByVal FechaI As String, ByVal FechaF As String) As Boolean
        'METHOD SUMMARY
        'Valida si una fecha esta dentro de un rango de fechas.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a evaluar
        'FechaI=Fecha Inicial
        'FechaF=Fecha Final
        'END PARAMETERS SUMMARY
        If Not ValidoTxtFecha(Fecha) Then
            FechaEstaEnElRango = False
            Exit Function
        End If
        If Not ValidoTxtFecha(FechaI) Then
            FechaEstaEnElRango = False
            Exit Function
        End If
        If Not ValidoTxtFecha(FechaF) Then
            FechaEstaEnElRango = False
            Exit Function
        End If
        FechaEstaEnElRango = False
        Dim Dia As String = Fecha.Substring(0, 2)
        Dim Ano As String = Fecha.Substring(6, 4)
        Dim Mes As String = Fecha.Substring(3, 2)
        Dim DiaI As String = FechaI.Substring(0, 2)
        Dim DiaF As String = FechaF.Substring(0, 2)
        Dim AnoI As String = FechaI.Substring(6, 4)
        Dim AnoF As String = FechaF.Substring(6, 4)
        Dim MesI As String = FechaI.Substring(3, 2)
        Dim MesF As String = FechaF.Substring(3, 2)
        Dim FFechaI As New System.DateTime(CEntero(Ano), CEntero(Mes), CEntero(Dia))
        Dim FFechaf As New System.DateTime(CEntero(AnoI), CEntero(MesI), CEntero(DiaI))
        Dim TiempoFechas As New System.TimeSpan(FFechaf.Ticks - FFechaI.Ticks)
        Select Case CDoble(TiempoFechas.TotalDays.ToString)
            Case Is = 0
                FechaEstaEnElRango = True
            Case Is > 0
                FechaEstaEnElRango = False
            Case Is < 0
                FechaEstaEnElRango = True
        End Select
        If FechaEstaEnElRango = True Then
            FFechaI = New System.DateTime(CEntero(AnoF), CEntero(MesF), CEntero(DiaF))
            FFechaf = New System.DateTime(CEntero(Ano), CEntero(Mes), CEntero(Dia))
            TiempoFechas = New System.TimeSpan(FFechaf.Ticks - FFechaI.Ticks)
            Select Case CDoble(TiempoFechas.TotalDays.ToString)
                Case Is = 0
                    FechaEstaEnElRango = True
                Case Is > 0
                    FechaEstaEnElRango = False
                Case Is < 0
                    FechaEstaEnElRango = True
            End Select
        End If
    End Function
#End Region
#Region "RemoteScriptingDeEncode"
    Public Shared Function RemoteScriptingDeEncode(ByVal Texto As String) As String
        'METHOD SUMMARY
        'Desencrypta una cadena encryptada con ScriptingEncode.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Texto=Cadena de texto
        'END PARAMETERS SUMMARY
        RemoteScriptingDeEncode = Texto
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[a]#", "á")
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[e]#", "é")
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[i]#", "í")
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[o]#", "ó")
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[u]#", "ú")
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[ñ]#", "ñ")
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[Ñ]#", "Ñ")
        RemoteScriptingDeEncode = RemoteScriptingDeEncode.Replace("#[ENTER]#", Convert.ToChar(13))
    End Function
#End Region
#Region "SaveError"
    Public Shared Sub SaveError(ByVal ProductName As String, ByVal Aplication As String, ByVal Vclass As String, ByVal Vfunction As String, ByVal Source As String, ByVal ErrDescription As String, ByVal User As String, ByVal File As String, ByVal StackTrace As String, ByVal Drive As String)
        'METHOD SUMMARY
        'Salva un error en un archivo XML
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'ProductName=Nombre del producto
        'Aplication=Aplicacion
        'Vclass=Clase
        'Vfunction=Funcion  o metodo.
        'Source=Origen del error
        'ErrDescription=Descripcion del error
        'User=Usuario
        'File=Archivo en donde se guardara el error
        'StackTrace=StackTrace del error
        'Drive=Drive a guardar el error.
        'END PARAMETERS SUMMARY
        Try
            If Drive.Trim.Length = 0 Then Drive = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location).Substring(0, 2)
            If File.Trim.Length = 0 Then
                Dim Path As String = ""
                If Drive.Trim.Length = 0 Then
                    Path = "c:\AppErrs\"
                Else
                    Path = Drive & "\AppErrs\"
                End If
                Path = Path & DateTime.Now.Year.ToString & "\" & (DateTime.Now.Month + 100).ToString.Substring(1, 2) & "\" & (DateTime.Now.Day + 100).ToString.Substring(1, 2)
                If Not IO.Directory.Exists(Path) Then IO.Directory.CreateDirectory(Path)
                If Aplication.Trim.Length = 0 Then
                    File = Path & "\Errs.xml"
                Else
                    File = Path & "\" & Aplication.Trim & ".xml"
                End If
            End If
            Dim MidataSet As New DataSet
            MidataSet.DataSetName = "MisErrores"
            If IO.File.Exists(File) Then
                MidataSet.ReadXml(File)
            Else
                Dim MisRegistros As New DataTable
                With MisRegistros
                    .Columns.Add("ProductName", System.Type.GetType("System.String"))
                    .Columns.Add("Aplication", System.Type.GetType("System.String"))
                    .Columns.Add("Class", System.Type.GetType("System.String"))
                    .Columns.Add("Function", System.Type.GetType("System.String"))
                    .Columns.Add("Date", System.Type.GetType("System.String"))
                    .Columns.Add("Time", System.Type.GetType("System.String"))
                    .Columns.Add("Source", System.Type.GetType("System.String"))
                    .Columns.Add("ErrDescription", System.Type.GetType("System.String"))
                    .Columns.Add("User", System.Type.GetType("System.String"))
                    .Columns.Add("StackTrace", System.Type.GetType("System.String"))
                End With
                MidataSet.Tables.Add(MisRegistros)
                MisRegistros.Dispose()
            End If
            Dim Args(9) As String
            Args(0) = ProductName
            Args(1) = Aplication
            Args(2) = Vclass
            Args(3) = Vfunction
            Args(4) = System.DateTime.Now.Day.ToString.Trim & "/" & System.DateTime.Now.Month.ToString.Trim & "/" & System.DateTime.Now.Year.ToString.Trim
            Args(5) = System.DateTime.Now.Hour.ToString.Trim & ":" & System.DateTime.Now.Minute.ToString.Trim & ":" & System.DateTime.Now.Second.ToString.Trim
            Args(6) = Source
            Args(7) = ErrDescription
            Args(8) = User
            Args(9) = StackTrace
            With MidataSet
                .Tables(0).Rows.Add(Args)
                .WriteXml(File)
                .Dispose()
            End With
            Args = Nothing
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "SinComas"
    Public Shared Function SinComas(ByVal Texto As String) As String
        'METHOD SUMMARY
        'Eliminas las dobles y simples comillas de una cadena de caracteres.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Texto=Cadena de texto
        'END PARAMETERS SUMMARY
        Texto = Texto.Replace(Convert.ToChar(34), "")
        Texto = Texto.Replace("'", "")
        Texto = Texto.Replace(Convert.ToChar(13) & Convert.ToChar(10), "")
        Return Texto
    End Function
#End Region
#Region "ErrorToJavascript"
    Public Shared Function ErrorToJavascript(ByVal Message As String) As String
        'METHOD SUMMARY
        'Formatea un texto para poder ser mostrado en una alert de javascript.
        'END METHOD SUMMARY
        Message = Message.Replace(Convert.ToChar(34), "")
        Message = Message.Replace("'", "")
        Message = Message.Replace(Convert.ToChar(13) & Convert.ToChar(10), "")
        Return Message
    End Function
#End Region
#Region "SqlErrAnalizer"
    Public Shared Function SqlErrAnalizer(ByVal Message As String) As String
        'METHOD SUMMARY
        'Analiza un error devuelto por SQL.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Message=Mensage para analizar
        'END PARAMETERS SUMMARY
        Message = Message.Trim
        If Message.Length > 60 Then
            If Message.Substring(0, 60).ToUpper = "DELETE STATEMENT CONFLICTED WITH COLUMN REFERENCE CONSTRAINT" Then
                Return "El registro no puede ser borrado, por tener datos vinculados. Intente deshabilitarlo"
            End If
        End If
        If Message.Length > 23 Then
            If Message.Substring(0, 23).ToUpper = "VIOLATION OF UNIQUE KEY" Then
                Return "El registro ya esta cargado"
            End If
        End If
        Return Message
    End Function
#End Region
#Region "CurrentSeconds"
    Public Shared Function CurrentSeconds() As String
        'METHOD SUMMARY
        'Devuelve los segundos actuales.
        'END METHOD SUMMARY
        Dim d As DateTime = DateTime.Now
        Dim ts As TimeSpan = d.TimeOfDay
        CurrentSeconds = CType(ts.TotalSeconds, Integer).ToString()
    End Function
#End Region
#Region "RepiteCaracter"
    Public Shared Function RepiteCaracter(ByVal Caracter As String, ByVal cantidad As Integer) As String
        'METHOD SUMMARY
        'Completa en una cadena de caracteres n caracteres.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Caracter=Caracter a completar la cadena de texto
        'cantidad=Cantidad de caracteres
        'END PARAMETERS SUMMARY
        Dim n As Integer
        RepiteCaracter = ""
        For n = 0 To cantidad - 1
            RepiteCaracter = RepiteCaracter & Caracter.Trim
        Next
    End Function
#End Region
#Region "AgregarDiasAFecha"
    Public Shared Function AgregarDiasAFecha(ByVal Fecha As String, ByVal Dias As Integer) As String
        'METHOD SUMMARY
        'Suma dias a una fecha.
        '
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Fecha=Fecha a modificar
        'Dias=Dias
        'END PARAMETERS SUMMARY
        Try
            Dim Mifecha As New DateTime(CEntero(Fecha.Substring(6, 4)), CEntero(Fecha.Substring(3, 2)), CEntero(Fecha.Substring(0, 2)))
            Mifecha = Mifecha.AddDays(Dias)
            Dim Dia As String, Mes As String, Ano As String
            Dia = Mifecha.Day.ToString
            Mes = Mifecha.Month.ToString
            Ano = Mifecha.Year.ToString
            If Dia.Length = 1 Then
                Dia = "0" & Dia
            End If
            If Mes.Length = 1 Then
                Mes = "0" & Mes
            End If
            AgregarDiasAFecha = Dia & "/" & Mes & "/" & Ano
        Catch
            Return TxtFechaHoy()
        End Try
    End Function
#End Region
#Region "CurrentTicks"
    Public Shared Function CurrentTicks() As Long
        'METHOD SUMMARY
        'Devuelve los Ticks actuales del sistema.
        'END METHOD SUMMARY
        Dim d As DateTime = DateTime.Now
        CurrentTicks = d.Ticks
    End Function
#End Region
#Region "KillExcel"
    Public Shared Sub KillExcel()
        'METHOD SUMMARY
        'Elimina todo porceso Excel.exe que este corriendo en la maquina
        'END METHOD SUMMARY
        Try
            Dim mp As System.Diagnostics.Process() = Process.GetProcessesByName("EXCEL")
            Dim ExcelList As New Process
            For Each ExcelList In mp
                ExcelList.Kill()
            Next ExcelList
            ExcelList.Dispose()
            mp = Nothing
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "MilesFormat"
    Public Shared Function MilesFormat(ByVal Number As String, ByVal AmericanFormat As Boolean) As String
        'METHOD SUMMARY
        'Formatea un numero con el separador decimal y de millares usados en USA.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'AmericanFormat=Devuelve el resultado en formato americanao xxx,xxx.xx
        'Number=Numero
        'END PARAMETERS SUMMARY
        MilesFormat = ""
        Dim WithDecimals As Boolean = False
        Dim n As Integer
        Dim Character As String = ""
        For n = 0 To Number.Trim.Length - 1
            Character = Number.Trim.Substring(n, 1)
            If AmericanFormat Then
                If Character = "." Then
                    WithDecimals = True
                    Exit For
                End If
            Else
                If Character = "," Then
                    WithDecimals = True
                    Exit For
                End If
            End If
        Next
        Dim CounterStart As Boolean
        If WithDecimals Then
            CounterStart = False
        Else
            CounterStart = True
        End If
        Dim Counter As Integer = 0
        For n = Number.Trim.Length - 1 To 0 Step -1
            If CounterStart = True Then Counter = Counter + 1
            If Counter = 4 Then
                If n <> Number.Trim.Length - 1 Then
                    If AmericanFormat Then
                        MilesFormat = "," & MilesFormat
                    Else
                        MilesFormat = "." & MilesFormat
                    End If
                End If
                Counter = 1
            End If
            Character = Number.Trim.Substring(n, 1)
            MilesFormat = Character & MilesFormat
            If AmericanFormat Then
                If Character = "." Then
                    CounterStart = True
                End If
            Else
                If Character = "." Then
                    CounterStart = True
                End If
            End If
        Next
        If MilesFormat.Trim.Length >= 2 Then If MilesFormat.Trim.Substring(0, 2) = "-," Then MilesFormat = MilesFormat.Replace("-,", "-")
    End Function
#End Region
#Region "ExternalIP"
    Public Shared Function ExternalIP() As String
        'METHOD SUMMARY
        'Devuelve la IP externa si se esta conectado a atraves de un Router.
        '
        'END METHOD SUMMARY
        ExternalIP = ""
        Try
            Dim Url As String = ""
            Dim Proxy As WebProxy
            Dim webrequest As WebRequest
            Url = "http://www.whatismyip.com/"
            Proxy = New WebProxy
            WebProxy.GetDefaultProxy()
            webrequest = Net.WebRequest.Create(Url)
            webrequest.Proxy = Proxy
            Dim objReader As New StreamReader(webrequest.GetResponse.GetResponseStream)
            Dim Html As String = objReader.ReadToEnd.ToUpper.Trim
            Dim OtherHtml As String
            OtherHtml = Html.Replace("WHATISMYIP.COM -", "Ñ").Trim
            Dim args() As String
            args = OtherHtml.Split(Convert.ToChar("Ñ"))
            OtherHtml = args(1).Replace("</TITLE>", "Ñ").Trim()
            args = OtherHtml.Split(Convert.ToChar("Ñ"))
            ExternalIP = args(0)
            args = Nothing
            webrequest = Nothing
            Proxy = Nothing
        Catch ex As Exception
        End Try
    End Function
#End Region
#Region "ClearTmpFiles"
    Public Shared Sub ClearTmpFiles(ByVal Path As String, ByVal Pattern As String, ByVal AgingInDays As Integer)
        'METHOD SUMMARY
        'Borra los archivos de un directorio segun su antiguedad.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Path=Ruta a Borrar los archivos
        'Pattern=Tipos de archivos a borrar, se debe usar los *.*, *.tmp, etc.
        'AgingInDays=Antiguedad de los archivos a borrar
        'END PARAMETERS SUMMARY
        Try
            Dim Files() As String
            Files = System.IO.Directory.GetFiles(Path, Pattern)
            Dim n As Integer
            Dim CurrentDate As New Date
            CurrentDate = Date.Now.AddDays(AgingInDays * -1)
            For n = 0 To Files.Length - 1
                If IO.File.GetCreationTime(Files(n)).Ticks <= CurrentDate.Ticks Then
                    IO.File.Delete(Files(n))
                End If
            Next
            Files = Nothing
        Catch ex As Exception

        End Try
    End Sub
#End Region
#Region "DivisionMod"
    Public Shared Function DivisionMod(ByVal Value As Double, ByVal Value2 As Double) As Integer
        'METHOD SUMMARY
        'Funcion equivalente al MOD de cualquier lenguaje de programacion.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'Value=Valor1
        'Value2=Valor2
        'END PARAMETERS SUMMARY
        Dim Result As String = (Value / Value2).ToString.Replace(",", ".")
        If Result.IndexOf(".") >= 0 Then
            DivisionMod = CEntero(Result.Substring(Result.IndexOf(".") + 1, Result.Length - Result.IndexOf(".") - 1))
        Else
            DivisionMod = 0
        End If
    End Function
#End Region
#Region "GetInternetIp"
    Public Shared Function GetInternetIp() As String
        'METHOD SUMMARY
        'Devuelve la IP externa si se esta conectado a atraves de un Router.
        'END METHOD SUMMARY
        GetInternetIp = ""
        Try
            Dim myWebClient As New System.Net.WebClient
            Dim myStream As System.IO.Stream = myWebClient.OpenRead("http://www.geekpedia.com/ip.php")
            Dim myStreamReader As New System.IO.StreamReader(myStream)
            Dim myIP As String = myStreamReader.ReadToEnd()
            myWebClient.Dispose()
            If myIP.Trim.Length = 0 Then
                myStream = myWebClient.OpenRead("http://vbnet.mvps.org/resources/tools/getpublicip.shtml")
                myStreamReader = New System.IO.StreamReader(myStream)
                GetInternetIp = myStreamReader.ReadToEnd()
            Else
                GetInternetIp = myIP
            End If
        Catch ex As Exception
        End Try
    End Function
#End Region
#Region "TiempoEntreFechas"
    Public Shared Function TiempoEntreFechas(ByVal FechaI As String, ByVal FechaF As String) As Double
        'METHOD SUMMARY
        'Devuelve los dias que hay entre 2 fechas.
        'END METHOD SUMMARY
        'PARAMETERS SUMMARY
        'FechaI=Fecha Inicial
        'FechaF=Fecha Final
        'END PARAMETERS SUMMARY
        If Not ValidoTxtFecha(FechaI) Then
            TiempoEntreFechas = -92417511
            Exit Function
        End If
        If Not ValidoTxtFecha(FechaF) Then
            TiempoEntreFechas = -92417511
            Exit Function
        End If
        TiempoEntreFechas = 0
        Dim DiaI As String = FechaI.Substring(0, 2)
        Dim DiaF As String = FechaF.Substring(0, 2)
        Dim AnoI As String = FechaI.Substring(6, 4)
        Dim AnoF As String = FechaF.Substring(6, 4)
        Dim MesI As String = FechaI.Substring(3, 2)
        Dim MesF As String = FechaF.Substring(3, 2)
        Dim FFechaI As New System.DateTime(CEntero(AnoI), CEntero(MesI), CEntero(DiaI))
        Dim FFechaf As New System.DateTime(CEntero(AnoF), CEntero(MesF), CEntero(DiaF))
        Dim TiempoFechas As New System.TimeSpan(FFechaf.Ticks - FFechaI.Ticks)
        TiempoEntreFechas = CDoble(TiempoFechas.TotalDays.ToString)
    End Function
#End Region
#Region "GetFilesFromFolder"
    Public Shared Function GetFilesFromFolder(ByVal Path As String, ByVal Result As DataTable, ByVal Pattern As String) As DataTable
        If Result Is Nothing Then
            Result = New DataTable
            Result.Columns.Add("File", System.Type.GetType("System.String"))
        End If
        Dim Files() As String
        Files = IO.Directory.GetFiles(Path, Pattern)
        Dim n As Integer
        For n = 0 To Files.Length - 1
            Dim args(0) As String
            args(0) = Files(n)
            If Result.Columns("File") Is Nothing Then
                Result.Columns.Add("File", System.Type.GetType("System.String"))
            End If
            Result.Rows.Add(args)
            args = Nothing
        Next
        Files = Nothing
        Dim Folders() As String
        Folders = IO.Directory.GetDirectories(Path, Pattern)
        For n = 0 To Folders.Length - 1
            GetFilesFromFolder(Folders(n), Result, Pattern)
        Next
        Folders = Nothing
        Return Result
    End Function
#End Region
#Region "GetMapFromAnimap"
    Public Shared Function GetMapFromAnimap(ByVal Calle As String, ByVal Altura As String, ByVal Localidad As String, ByVal Partido As String, ByVal Provincia As String, ByVal MapFile As String) As String
        GetMapFromAnimap = ""
        If Altura.Trim.Length = 0 Then
            Dim args() As String
            args = Calle.Split(CChar(" "))
            If CEntero(args(args.Length - 1)) > 0 Then
                Altura = args(args.Length - 1)
                Calle = Calle.Replace(args(args.Length - 1), "").Trim
            End If
            args = Nothing
        End If
        If Partido.Trim.ToUpper = "CAPITAL FEDERAL" Then Provincia = Partido
        Dim IE As SHDocVw.InternetExplorer
        Dim objMSHTML As New mshtml.HTMLDocument
        Try
            Dim ProcessAntes As System.Diagnostics.Process()
            Dim ProcessDespues As System.Diagnostics.Process()
            Dim DeletedProcess As Boolean = True
            Try
                ProcessAntes = Process.GetProcessesByName("iexplore")
            Catch ex As Exception
                DeletedProcess = False
            End Try
            IE = New SHDocVw.InternetExplorer
            Try
                ProcessDespues = Process.GetProcessesByName("iexplore")
            Catch ex As Exception
                DeletedProcess = False
            End Try
            IE.Navigate("www.animap.com.ar/")
            IE.Visible = False
            While IE.Busy
                'Application.DoEvents()
            End While
            objMSHTML = CType(IE.Document, mshtml.HTMLDocument)
            Do Until objMSHTML.readyState = "complete"
                'Application.DoEvents()
            Loop
            Dim ObjHtml As mshtml.HTMLInputElementClass
            Dim ObjHtml2 As mshtml.HTMLSelectElementClass
            ObjHtml = CType(objMSHTML.all.item("calle"), mshtml.HTMLInputElementClass)
            ObjHtml.value = Calle.Trim
            ObjHtml = CType(objMSHTML.all.item("altura"), mshtml.HTMLInputElementClass)
            ObjHtml.value = Altura.Trim
            ObjHtml = CType(objMSHTML.all.item("localidad"), mshtml.HTMLInputElementClass)
            If Localidad.Trim.Length = 0 Then
                ObjHtml.value = Partido.Trim
            Else
                ObjHtml.value = Localidad.Trim
            End If
            ObjHtml = CType(objMSHTML.all.item("partido"), mshtml.HTMLInputElementClass)
            ObjHtml.value = Partido.Trim
            ObjHtml2 = CType(objMSHTML.all.item("provincia"), mshtml.HTMLSelectElementClass)
            Select Case Provincia.Trim.ToUpper
                Case "BUENOS AIRES"
                    ObjHtml2.selectedIndex = 1
                Case "CAPITAL FEDERAL"
                    ObjHtml2.selectedIndex = 0
            End Select
            ObjHtml = CType(objMSHTML.all.item("Buscar"), mshtml.HTMLInputElementClass)
            ObjHtml.click()
            Do Until objMSHTML.readyState = "complete"
                '    'Application.DoEvents()
            Loop
            Dim CurrentLocationUrl As String = objMSHTML.url
            If CurrentLocationUrl.EndsWith("map_rubro=") Then
                Dim ObjHtml3 As mshtml.HTMLImgClass
                ObjHtml3 = CType(objMSHTML.all.item("mapagenerado"), mshtml.HTMLImgClass)
                GetMapFromAnimap = ObjHtml3.href
                ObjHtml3 = Nothing
                If MapFile.Trim.Length <> 0 Then
                    Try
                        Dim objStream As Stream
                        Dim Proxy As New WebProxy(GetMapFromAnimap, True)
                        Dim webrequest As WebRequest = Net.WebRequest.Create(GetMapFromAnimap)
                        webrequest = HttpWebRequest.Create(GetMapFromAnimap)
                        Dim response As WebResponse
                        response = webrequest.GetResponse
                        objStream = response.GetResponseStream
                        Dim myimage As Image
                        myimage = Image.FromStream(objStream)
                        myimage.Save(MapFile)
                        Proxy = Nothing
                        webrequest = Nothing
                        objStream.Close()
                        objStream = Nothing
                        myimage.Dispose()
                    Catch ex As Exception
                    End Try
                End If
                GetMapFromAnimap = CurrentLocationUrl
            Else
                GetMapFromAnimap = "busca_direccion_listado"
            End If
            ObjHtml = Nothing
            ObjHtml2 = Nothing
            objMSHTML.close()
            objMSHTML = Nothing
            IE = Nothing
            If DeletedProcess Then
                Dim OneProcess As New Process
                Dim OneProcess2 As New Process
                For Each OneProcess In ProcessDespues
                    For Each OneProcess2 In ProcessAntes
                        If OneProcess.Id = OneProcess2.Id Then Exit For
                        OneProcess.Kill()
                    Next
                Next OneProcess
                OneProcess.Dispose()
                OneProcess2.Dispose()
                Dim mp As System.Diagnostics.Process()
                mp = Process.GetProcessesByName("aAvgApi")
                Dim n As Integer
                For n = 0 To mp.Length - 1
                    mp(n).Kill()
                Next
                mp = Nothing
            End If
        Catch ex As Exception
            SaveError("", "", "", "", ex.Source, ex.Message, "", "", ex.StackTrace, "")
        End Try
    End Function
#End Region
#Region "KillIE"
    Public Shared Sub KillIE()
        Try
            Dim mp As System.Diagnostics.Process() = Process.GetProcessesByName("iexplore")
            Dim n As Integer
            If mp.Length >= 2 Then
                For n = 1 To mp.Length - 1
                    mp(n).Kill()
                Next
            End If
            mp = Process.GetProcessesByName("aAvgApi")
            For n = 0 To mp.Length - 1
                mp(n).Kill()
            Next
            mp = Nothing
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "MinuteBetweenDates"
    Public Shared Function MinutesBetweenDates(ByVal Date1 As String, ByVal Date2 As String, ByVal Time1 As String, ByVal Time2 As String, ByVal ShowNegativeResult As Boolean) As Integer
        MinutesBetweenDates = 0
        Dim d1 As New DateTime(CEntero(Date1.Substring(6, 4)), CEntero(Date1.Substring(3, 2)), CEntero(Date1.Substring(0, 2)), CEntero(Time1.Trim.Substring(0, 2)), CEntero(Time1.Trim.Substring(3, 2)), CEntero(Time1.Trim.Substring(6, 2)))
        Dim d2 As New DateTime(CEntero(Date2.Substring(6, 4)), CEntero(Date2.Substring(3, 2)), CEntero(Date2.Substring(0, 2)), CEntero(Time2.Trim.Substring(0, 2)), CEntero(Time2.Trim.Substring(3, 2)), CEntero(Time2.Trim.Substring(6, 2)))
        Dim ts1 As TimeSpan = d1.TimeOfDay
        Dim ts2 As TimeSpan = d2.TimeOfDay
        Dim ts3 As TimeSpan = d2 - d1
        MinutesBetweenDates = CType(ts3.TotalMinutes, Integer)
        If Not ShowNegativeResult Then If MinutesBetweenDates < 0 Then MinutesBetweenDates = 0
    End Function
#End Region
#Region "DateToTimeTxt"
    Public Shared Function DateToTimeTxt(ByVal TheDate As Date) As String
        Dim MyHour As String, MyMin As String, MySec As String
        MyHour = TheDate.Hour.ToString
        MyMin = TheDate.Minute.ToString
        MySec = TheDate.Second.ToString
        If MyHour.Length = 1 Then MyHour = "0" & MyHour
        If MyMin.Length = 1 Then MyMin = "0" & MyMin
        If MySec.Length = 1 Then MySec = "0" & MySec
        DateToTimeTxt = MyHour & ":" & MyMin & ":" & MySec
    End Function
#End Region
#Region "DateToTxt"
    Public Shared Function DateToTxt(ByVal TheDate As Date) As String
        Dim MyDay As String, MyMonth As String, MyYear As String
        MyDay = TheDate.Day.ToString
        MyMonth = TheDate.Month.ToString
        MyYear = TheDate.Year.ToString
        If MyDay.Length = 1 Then MyDay = "0" & MyDay
        If MyMonth.Length = 1 Then MyMonth = "0" & MyMonth
        DateToTxt = MyDay & "/" & MyMonth & "/" & MyYear
    End Function
#End Region
End Class


