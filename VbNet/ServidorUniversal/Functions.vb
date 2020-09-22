Public Class Functions
#Region "ToInt"
    Public Shared Function ToInt(ByVal Number As String) As Integer
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
#Region "ToLong"
    Public Shared Function Tolong(ByVal Number As String) As Long
        If Number = Nothing Then Return 0
        If Number.Trim.Length = 0 Then Return 0
        If Number.ToLower = "true" Or Number.ToLower = "verdadero" Then Return 1
        If Number.ToLower = "false" Or Number.ToLower = "falso" Then Return 0
        Try
            Tolong = CType(Number, Int64)
        Catch
            Return 0
        End Try
    End Function
#End Region
#Region "ToDouble"
    Public Shared Function ToDouble(ByVal Numero As String) As Double
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
#Region "Enter"
    Public Shared Function Enter() As String
        Return Convert.ToChar(13) & Convert.ToChar(10)
    End Function
#End Region
#Region "Left"
    Public Shared Function Left(ByVal Cadena As String, ByVal Posiciones As Integer) As String
        If Posiciones > Cadena.Trim.Length Then Return Cadena
        Return Cadena.Trim.Substring(0, Posiciones)
    End Function
#End Region
#Region "Right"
    Public Shared Function Right(ByVal Cadena As String, ByVal Posiciones As Integer) As String
        If Posiciones > Cadena.Trim.Length Then Return Cadena
        Return Cadena.Trim.Substring(Cadena.Trim.Length - Posiciones, Posiciones)
    End Function
#End Region
#Region "IsNumeric"
    Public Shared Function IsNumeric(ByVal Numero As String) As Boolean
        Try
            Convert.ToDouble(Numero)
        Catch
            Return False
        End Try
        Return True
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
End Class
