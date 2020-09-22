Option Explicit On
Option Strict On
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports System.Net
Imports System.Xml
Public Class Functions
#Region "BinToHex"
    Public Shared Function BinToHex(ByVal bits As String) As Long
        If (bits = "") Then
            BinToHex = 0
        Else
            BinToHex = 2 * BinToHex(Left(bits, Len(bits) - 1)) + CLng(Right(bits, 1))
        End If
    End Function
#End Region
#Region "CreatePassword"
    Public Shared Function CreatePassword(ByVal Length As Integer) As String
        CreatePassword = "J@92Ra741"
        Try
            Dim Chars As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
            CreatePassword = ""
            Dim i As Integer
            Dim MyRandom As New Random(Now.Hour + Now.Minute + Now.Second)
            For x As Integer = 0 To Length - 1
                i = CInt(Math.Floor(Convert.ToDecimal("0." & MyRandom.Next.ToString) * 62))
                CreatePassword += Chars.Substring(i, 1)
            Next
            Sleep(1000)
            CreatePassword = CreatePassword.Substring(0, 1).Trim.ToUpper & CreatePassword.Substring(1)
        Catch ex As Exception
            CreatePassword = "J@92Ra741"
        End Try
    End Function
#End Region
#Region "ToBinary"
    Public Shared Function ToBinary(ByVal Num As Decimal) As String
        ToBinary = ""
        Do
            ToBinary = IIf(Num / 2 <> Int(Num / 2), 1, 0).ToString & ToBinary
            Num = Int(Num / 2)
        Loop Until Num = 0
    End Function
#End Region
#Region "ClearTmpDirectories"
    Public Shared Sub ClearTmpDirectories(ByVal Path As String, ByVal Pattern As String, ByVal AgingInDays As Integer)
        Try
            Dim Files() As String
            Files = System.IO.Directory.GetDirectories(Path, Pattern)
            Dim n As Integer
            Dim CurrentDate As New Date
            CurrentDate = Date.Now.AddDays(AgingInDays * -1)
            For n = 0 To Files.Length - 1
                If IO.Directory.GetCreationTime(Files(n)) <= CurrentDate Then
                    IO.Directory.Delete(Files(n))
                End If
            Next
            Files = Nothing
        Catch ex As Exception

        End Try
    End Sub
#End Region
#Region "Encriptador"
    Private Shared ReadOnly Property Key1() As Byte()
        Get
            Dim Clave As String = "%/k@%q=?%%$%&9z{+]@2mxFh"
            Return Encoding.Default.GetBytes(Clave)
        End Get
    End Property
    Private Shared ReadOnly Property Key2() As Byte()
        Get
            Dim Clave As String = "1p!9z$%&"
            Return Encoding.Default.GetBytes(Clave)
        End Get
    End Property
    Public Shared Function Encrypt(ByVal Text As String) As String
        Dim Des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        Encrypt = Transform(Text, Des.CreateEncryptor(Key1, Key2))
        Encrypt = Encrypt.Replace(Convert.ToChar(34), "JAR1")
        Encrypt = Encrypt.Replace(Convert.ToChar(32), "JAR2")
        Encrypt = Encrypt.Replace(Convert.ToChar(13), "JAR3")
        Encrypt = Encrypt.Replace(Convert.ToChar(27), "JAR4")
        Encrypt = Encrypt.Replace(Microsoft.VisualBasic.Chr(10), "JAR5")
        Return Encrypt
    End Function
    Public Shared Function Decrypt(ByVal encryptedText As Object) As String
        Dim Des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        encryptedText = encryptedText.ToString.Replace("JAR1", Convert.ToChar(34)) 'Comillas dobles
        encryptedText = encryptedText.ToString.Replace("JAR2", Convert.ToChar(32)) ' Espacios
        encryptedText = encryptedText.ToString.Replace("JAR3", Convert.ToChar(13)) 'Enters
        encryptedText = encryptedText.ToString.Replace("JAR4", Convert.ToChar(27)) 'Esc
        encryptedText = encryptedText.ToString.Replace("JAR5", Microsoft.VisualBasic.Chr(10)) 'Esc
        Return Transform(encryptedText.ToString, Des.CreateDecryptor(Key1, Key2))
    End Function
    Private Shared Function Transform(ByVal Text As String, ByVal CryptoTransform As ICryptoTransform) As String
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
#Region "ToInt"
    Public Shared Function ToInt(ByVal Numero As String) As Integer
        Numero = Numero.Trim
        If Numero = Nothing Then Return 0
        If Numero.Trim.Length = 0 Then Return 0
        If Numero.ToLower = "true" Or Numero.ToLower = "verdadero" Then Return 1
        If Numero.ToLower = "false" Or Numero.ToLower = "falso" Then Return 0
        For n As Integer = 0 To Numero.Length - 1
            If Not CaracterNumerico(Numero.Substring(n, 1)) Then
                Return 0
            End If
        Next
        'Try
        ToInt = Integer.Parse(Numero)
        'Catch
        'Return 0
        'End Try
    End Function
#End Region
#Region "CaracterNumerico"
    Private Shared Function CaracterNumerico(ByVal Valor As String) As Boolean
        If Valor.Trim = "1" Or Valor.Trim = "2" Or Valor.Trim = "3" Or Valor.Trim = "4" Or Valor.Trim = "5" Or Valor.Trim = "6" Or Valor.Trim = "7" Or Valor.Trim = "8" Or Valor.Trim = "9" Or Valor.Trim = "0" Then
            CaracterNumerico = True
        Else
            CaracterNumerico = False
        End If
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
#Region "GiveMeFileName"
    Public Shared Function GiveMeFileName(ByVal ThePath As String) As String
        If ThePath.IndexOf("/") = -1 And ThePath.IndexOf("\") = -1 Then
            Return ThePath
        End If
        Dim N As Integer
        Dim TheCharacter As String = ""
        For N = ThePath.Trim.Length To 1 Step -1
            TheCharacter = Right(Left(ThePath, N), 1)
            If TheCharacter = "/" Or TheCharacter = "\" Then
                TheCharacter = Right(ThePath, ThePath.Trim.Length - N)
                Exit For
            End If
        Next
        GiveMeFileName = TheCharacter
    End Function
#End Region
#Region "TxtDateToday"
    Public Shared Function TxtDateToday() As String
        Dim MyDay As String, MyMonth As String, MyYear As String
        MyDay = System.DateTime.Today().Day.ToString
        MyMonth = System.DateTime.Today().Month.ToString
        MyYear = System.DateTime.Today().Year.ToString
        If MyDay.Length = 1 Then MyDay = "0" & MyDay
        If MyMonth.Length = 1 Then MyMonth = "0" & MyMonth
        TxtDateToday = MyDay & "/" & MyMonth & "/" & MyYear
    End Function
#End Region
#Region "SystemTime"
    Public Shared Function SystemTime() As String
        Dim Hour As String
        If System.DateTime.Now.Hour.ToString.Trim.Length = 1 Then
            Hour = "0" & System.DateTime.Now.Hour.ToString & ":"
        Else
            Hour = System.DateTime.Now.Hour.ToString & ":"
        End If
        If System.DateTime.Now.Minute.ToString.Trim.Length = 1 Then
            Hour = Hour & "0" & System.DateTime.Now.Minute.ToString & ":"
        Else
            Hour = Hour & System.DateTime.Now.Minute.ToString & ":"
        End If
        If System.DateTime.Now.Second.ToString.Trim.Length = 1 Then
            Hour = Hour & "0" & System.DateTime.Now.Second.ToString
        Else
            Hour = Hour & System.DateTime.Now.Second.ToString
        End If
        SystemTime = Hour
    End Function
#End Region
#Region "SystemTimeOtherCity"
    Public Shared Function SystemTimeOtherCity(ByVal Dif As Integer) As String
        Dim Hour As String
        Dim MyTime As New DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, System.DateTime.Now.Day, System.DateTime.Now.Hour + Dif, System.DateTime.Now.Minute, System.DateTime.Now.Second)
        If MyTime.Hour.ToString.Trim.Length = 1 Then
            Hour = "0" & MyTime.Hour.ToString & ":"
        Else
            Hour = MyTime.Hour.ToString & ":"
        End If
        If MyTime.Minute.ToString.Trim.Length = 1 Then
            Hour = Hour & "0" & MyTime.Minute.ToString & ":"
        Else
            Hour = Hour & MyTime.Minute.ToString & ":"
        End If
        If MyTime.Second.ToString.Trim.Length = 1 Then
            Hour = Hour & "0" & MyTime.Second.ToString
        Else
            Hour = Hour & MyTime.Second.ToString
        End If
        SystemTimeOtherCity = Hour
    End Function
#End Region
#Region "SaveErrr"
    Public Shared Sub SaveErr(ByVal ProductName As String, ByVal Aplication As String, ByVal Vclass As String, ByVal Vfunction As String, ByVal Source As String, ByVal ErrDescription As String, ByVal User As String, ByVal File As String, ByVal StackTrace As String, ByVal Drive As String)
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
    End Sub
#End Region
#Region "ToBoolean"
    Public Shared Function ToBoolean(ByVal Value As String) As Boolean
        If Value.Trim = "0" Or Value.Trim.ToLower = "off" Or Value.Trim.ToUpper = "N" Or Value.Trim.ToUpper = "NO" Or Value.Trim.ToUpper = "FALSE" Or Value.Trim.ToUpper = "" Or Value.Trim.ToUpper = ".F." Then Return False
        If Value.Trim = "1" Or Value.Trim.ToLower = "on" Or Value.Trim.ToUpper = "Y" Or Value.Trim.ToUpper = "SI" Or Value.Trim.ToUpper = "TRUE" Or Value.Trim.ToUpper = ".T." Then Return True
        Try
            ToBoolean = Boolean.Parse(Value)
        Catch ex As Exception
            ToBoolean = False
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
#Region "DecimalFormat"
    Public Shared Function DecimalFormat(ByVal TheValue As String, ByVal DecimalsDigits As Integer) As String
        If TheValue = "" Then
            Select Case DecimalsDigits
                Case 0
                    Return "0"
                Case 1
                    Return "0.0"
                Case 2
                    Return "0.00"
                Case 3
                    Return "0.000"
                Case 4
                    Return "0.0000"
                Case Else
                    Return "0.00"
            End Select
        End If
        Dim Flag As Boolean
        Dim N As Integer, Vlength As Integer
        Dim Cadena As String, MyString2 As String = "", Pesos As String
        Dim Veces As Integer, Veces2 As Integer
        Dim Rest As String = "", X As Integer, Counter As Integer
        Dim Negative As Boolean
        For X = 1 To DecimalsDigits
            Rest = Rest & "0"
        Next
        Vlength = TheValue.Trim.Length
        If TheValue.Substring(0, 1) = "-" Then
            Negative = True
            TheValue = Right(TheValue, Vlength - 1)
        Else
            Negative = False
        End If
        TheValue = Math.Round(ToDouble(TheValue), ToInt(DecimalsDigits.ToString)).ToString
        Vlength = TheValue.Trim.Length
        If IsNumeric(TheValue) = False Then
            DecimalFormat = "0." & Rest
            Exit Function
        End If
        If Vlength = 1 Then
            If IsNumeric(TheValue.Trim) = False Then
                DecimalFormat = "0." & Rest
                If DecimalsDigits = 0 Then DecimalFormat = DecimalFormat.Replace(".0", "").Replace(".00", "").Replace(".", "")
                Exit Function
            Else
                DecimalFormat = TheValue.Trim & "." & Rest
                If Negative = True Then DecimalFormat = "-" & DecimalFormat
                If DecimalsDigits = 0 Then DecimalFormat = DecimalFormat.Replace(".0", "").Replace(".00", "").Replace(".", "")
                Exit Function
            End If
        End If
        If Vlength = 0 Then
            DecimalFormat = "0." & Rest
            If DecimalsDigits = 0 Then DecimalFormat = DecimalFormat.Replace(".0", "").Replace(".00", "").Replace(".", "")
            Exit Function
        End If
        For N = 0 To TheValue.Trim.Length - 1
            Cadena = Left(Right(TheValue.Trim, Vlength - N), 1)
            If IsNumeric(Cadena) = False Then
                If Cadena = "-" Then
                    Veces2 = Veces2 + 1
                    MyString2 = MyString2 & Cadena
                Else
                    MyString2 = MyString2 & "."
                    Veces = Veces + 1
                End If
            Else
                MyString2 = MyString2 & Cadena
            End If
        Next
        If Veces > 1 Then
            DecimalFormat = "0." & Rest
            If DecimalsDigits = 0 Then DecimalFormat = DecimalFormat.Replace(".0", "").Replace(".00", "").Replace(".", "")
            Exit Function
        End If
        If Veces2 > 1 Then
            DecimalFormat = "0." & Rest
            If DecimalsDigits = 0 Then DecimalFormat = DecimalFormat.Replace(".0", "").Replace(".00", "").Replace(".", "")
            Exit Function
        End If
        Flag = False
        Pesos = MyString2
        If Pesos.Trim.Length = 1 Then
            Pesos = Pesos.Trim & "." & Rest
        Else
            Vlength = Pesos.Trim.Length
            Counter = 0
            For N = 0 To Pesos.Trim.Length - 1
                Cadena = Left(Right(Pesos.Trim, Vlength - N), 1)
                If Counter > 0 Then Counter = Counter + 1
                If Cadena = "." Then
                    Counter = Counter + 1
                End If
            Next
        End If
        If Counter = 0 Then
            Pesos = Pesos & "." & Rest
        Else
            Cadena = ""
            For N = 1 To DecimalsDigits - (Counter - 1)
                Cadena = Cadena & "0"
            Next
            Pesos = Pesos & Cadena
        End If
        Vlength = Pesos.Trim.Length
        For N = 0 To Pesos.Trim.Length - 1
            Cadena = Left(Right(Pesos.Trim, Vlength - N), 1)
            If Cadena = "." Or Cadena = "," Then
                Flag = True
            End If
        Next
        If Flag = False Then
            Pesos = Pesos.Trim & "." & Rest
        End If
        Flag = False
        If Left(Right(Pesos.Trim, DecimalsDigits + 1), 1) <> "." Then
            DecimalFormat = "0." & Rest
            If DecimalsDigits = 0 Then DecimalFormat = DecimalFormat.Replace(".0", "").Replace(".00", "").Replace(".", "")
            Exit Function
        End If
        If Negative = True Then Pesos = "-" & Pesos
        If DecimalsDigits = 0 And Pesos.Trim.Substring(Pesos.Trim.Length - 1, 1) = "." Then
            Pesos = Pesos.Trim.Substring(0, Pesos.Trim.Length - 1)
        End If
        DecimalFormat = Pesos
        If DecimalsDigits = 0 Then DecimalFormat = DecimalFormat.Replace(".0", "").Replace(".00", "").Replace(".", "")
        Exit Function
    End Function
#End Region
#Region "Extract"
    Public Shared Function Extract(ByVal Cadena As String, ByVal Separador As String, ByVal Posicion As Integer) As String
        Dim Largo As Integer
        Dim TheCharacter As String
        Dim Separadores As Integer
        Dim Pinicio As Integer
        Dim Tresul As String = ""
        Dim auxcant As Integer
        Dim N As Integer
        Largo = Cadena.Trim.Length
        Pinicio = 1
        auxcant = 0
        For N = 0 To (Largo - 1)
            TheCharacter = Left(Right(Cadena.Trim, (Largo - N)), 1)
            auxcant = auxcant + 1
            If TheCharacter.Trim = Separador.Trim Then
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
        Extract = Tresul
    End Function
#End Region
#Region "AmericanDate"
    Public Shared Function AmericanDate(ByVal MyDate As String) As String
        If MyDate.Trim.Length <> 10 Then
            AmericanDate = MyDate
            Exit Function
        End If
        Dim MyDay As String, MyMonth As String, MyYear As String
        MyDay = MyDate.Trim.Substring(0, 2)
        MyYear = Right(MyDate.Trim, 4)
        MyMonth = Right(Left(MyDate.Trim, 5), 2)
        AmericanDate = MyMonth & "/" & MyDay & "/" & MyYear
    End Function
#End Region
#Region "GiveMeFileExt"
    Public Shared Function GiveMeFileExt(ByVal MyFile As String) As String
        GiveMeFileExt = ""
        MyFile = MyFile.Trim
        Dim largo As Integer = MyFile.Length
        Dim N As Integer
        For N = largo - 1 To 0 Step -1
            If MyFile.Substring(N, 1) = "." Then
                GiveMeFileExt = MyFile.Substring(N, largo - N)
                Exit For
            End If
        Next
    End Function
#End Region
#Region "ValidateTxtDate"
    Public Shared Function ValidateTxtDate(ByVal MyDate As String) As Boolean
        If MyDate.Trim.Length <> 10 Then
            Return False
        End If
        Dim MyDay As Integer = ToInt(MyDate.Substring(0, 2))
        Dim MyMonth As Integer = ToInt(MyDate.Substring(3, 2))
        Dim MyYear As Integer = ToInt(MyDate.Substring(6, 4))
        Try
            Dim MyDateDate As New System.DateTime(MyYear, MyMonth, MyDay)
            If MyDateDate.Day <> MyDay Or MyDateDate.Month <> MyMonth Or MyDateDate.Year <> MyYear Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region
#Region "Sleep"
    Public Shared Sub Sleep(ByVal Time As Integer)
        System.Threading.Thread.Sleep(Time)
    End Sub
#End Region
#Region "TemporaryFileName"
    Public Shared Function TemporaryFileName(ByVal Extension As String) As String
        Dim r As Random = New Random
        Dim NumeroRandom As String = r.Next(1, 999999).ToString
        TemporaryFileName = Date.Today.Year.ToString & Date.Today.Month.ToString & Date.Today.Day.ToString & Date.Today.Minute.ToString & Date.Today.Second.ToString & NumeroRandom.ToString & "." & Extension
    End Function
#End Region
#Region "IsSubmit"
    Public Shared Function IsSubmit(ByVal Cadena As String) As String
        If Cadena = Nothing Then
            Return ""
        Else
            Return Cadena.Trim
        End If
    End Function
#End Region
#Region "BritishDate"
    Public Shared Function BritishDate(ByVal thedate As String) As String
        If thedate.Trim.Length <> 10 Then
            BritishDate = thedate
            Exit Function
        End If
        Dim day As String, Month As String, year As String
        Month = thedate.Trim.Substring(0, 2)
        year = Right(thedate.Trim, 4)
        day = Right(Left(thedate.Trim, 5), 2)
        BritishDate = day & "/" & Month & "/" & year
    End Function
#End Region
#Region "EndMonthDay"
    Public Shared Function EndMonthDay(ByVal Month As Integer, ByVal Year As Integer, ByVal Americana As Boolean) As String
        EndMonthDay = ""
        Dim TxtMonth As String
        If Month.ToString.Trim.Length = 1 Then
            TxtMonth = "0" + Month.ToString.Trim
        Else
            TxtMonth = Month.ToString.Trim
        End If
        Select Case Month
            Case 1
                EndMonthDay = "31/" + TxtMonth + "/" + Year.ToString.Trim
            Case 2
                If TxtToDate("29/02/" + Year.ToString) = System.DateTime.Now Then
                    EndMonthDay = "28/" + TxtMonth + "/" + Year.ToString.Trim
                Else
                    EndMonthDay = "29/" + TxtMonth + "/" + Year.ToString.Trim
                End If
            Case 3
                EndMonthDay = "31/" + TxtMonth + "/" + Year.ToString.Trim
            Case 4
                EndMonthDay = "30/" + TxtMonth + "/" + Year.ToString.Trim
            Case 5
                EndMonthDay = "31/" + TxtMonth + "/" + Year.ToString.Trim
            Case 6
                EndMonthDay = "30/" + TxtMonth + "/" + Year.ToString.Trim
            Case 7
                EndMonthDay = "31/" + TxtMonth + "/" + Year.ToString.Trim
            Case 8
                EndMonthDay = "31/" + TxtMonth + "/" + Year.ToString.Trim
            Case 9
                EndMonthDay = "30/" + TxtMonth + "/" + Year.ToString.Trim
            Case 10
                EndMonthDay = "31/" + TxtMonth + "/" + Year.ToString.Trim
            Case 11
                EndMonthDay = "30/" + TxtMonth + "/" + Year.ToString.Trim
            Case 12
                EndMonthDay = "31/" + TxtMonth + "/" + Year.ToString.Trim
        End Select
        If Americana Then
            EndMonthDay = AmericanDate(EndMonthDay)
        End If
    End Function
#End Region
#Region "TxtToDate"
    Public Shared Function TxtToDate(ByVal TheDate As String) As DateTime
        Try
            Dim MyDate As New DateTime(ToInt(TheDate.Substring(6, 4)), ToInt(TheDate.Substring(3, 2)), ToInt(TheDate.Substring(0, 2)))
            Return MyDate
        Catch
            Return System.DateTime.Now
        End Try
    End Function
#End Region
#Region "TxtDateFirstDay"
    Public Shared Function TxtDateFirstDay(ByVal Month As Integer, ByVal year As Integer, ByVal Americana As Boolean) As String
        If Month > 12 Or Month < 1 Then Return ""
        Dim day As String, Vmonth As String, Vyear As String
        day = "01"
        Vmonth = Month.ToString.Trim
        Vyear = year.ToString.Trim
        If day.Length = 1 Then
            day = "0" & day
        End If
        If Vmonth.Length = 1 Then
            Vmonth = "0" & Vmonth
        End If
        If Americana Then
            TxtDateFirstDay = Vmonth & "/" & day & "/" & Vyear
        Else
            TxtDateFirstDay = day & "/" & Vmonth & "/" & Vyear
        End If
    End Function
#End Region
#Region "Tabs"
    Public Shared Function Tabs(ByVal Quantity As Integer) As String
        Dim n As Integer
        Tabs = ""
        For n = 1 To Quantity
            Tabs = Tabs & Convert.ToChar(9)
        Next
    End Function
#End Region
#Region "FillCharacter"
    Public Shared Function FillCharacter(ByVal MyChar As String, ByVal Quantity As Integer) As String
        Dim n As Integer
        FillCharacter = ""
        For n = 1 To Quantity
            FillCharacter = FillCharacter & MyChar
        Next
    End Function
#End Region
#Region "Spaces"
    Public Shared Function Spaces(ByVal Quantity As Integer) As String
        Dim n As Integer
        Spaces = ""
        For n = 1 To Quantity
            Spaces = Spaces & " "
        Next
    End Function
#End Region
#Region "CurrentSeconds"
    Public Shared Function CurrentSeconds() As String
        Dim d As DateTime = DateTime.Now
        Dim ts As TimeSpan = d.TimeOfDay
        CurrentSeconds = CType(ts.TotalSeconds, Integer).ToString()
    End Function
#End Region
#Region "ErrorToJavascript"
    Public Shared Function ErrorToJavascript(ByVal MyMessage As String) As String
        MyMessage = MyMessage.Replace(Convert.ToChar(34), "")
        MyMessage = MyMessage.Replace("'", "")
        MyMessage = MyMessage.Replace(Convert.ToChar(13) & Convert.ToChar(10), "")
        MyMessage = MyMessage.Replace(Convert.ToChar(13), "")
        MyMessage = MyMessage.Replace(Convert.ToChar(10), "")
        Return MyMessage
    End Function
#End Region
#Region "SqlErrAnalizer"
    Public Shared Function SqlErrAnalizer(ByVal MyMessage As String) As String
        MyMessage = MyMessage.Trim
        If MyMessage.Length > 60 Then
            If MyMessage.Substring(0, 60).ToUpper = "DELETE STATEMENT CONFLICTED WITH COLUMN REFERENCE CONSTRAINT" Then
                Return "The record couldn't be deleted because it has related data. Change the enabled status"
            End If
        End If
        If MyMessage.Length > 23 Then
            If MyMessage.Substring(0, 23).ToUpper = "VIOLATION OF UNIQUE KEY" Then
                Return "You are duplicating a record. Operation was aborted."
            End If
        End If
        If MyMessage.Length > 39 Then
            If MyMessage.Substring(0, 39).ToUpper = "Infracción de la restricción UNIQUE KEY" Then
                Return "Esta intentando duplicar un registro"
            End If
        End If
        If MyMessage.StartsWith("No se puede insertar una fila de clave duplicada en el objeto dbo.Perfiles con índice único") Then
            Return "Esta intentando duplicar un registro o campo clave"
        End If
        Return MyMessage.Replace(Convert.ToChar(34), "'")
    End Function
#End Region
#Region "AddDaysToDates"
    ''' <summary>
    ''' Suma dias a una fecha
    ''' </summary>
    ''' <param name="MyDate">Fecha a sumar los dias</param>
    ''' <param name="Days">Dias a incrementar</param>
    Public Shared Function AddDaysToDates(ByVal MyDate As String, ByVal Days As Integer) As String
        If Not ValidateTxtDate(MyDate) Then
            AddDaysToDates = "**/**/****"
            Exit Function
        End If
        Dim MyDay As String = MyDate.Substring(0, 2)
        Dim MyMonth As String = MyDate.Substring(3, 2)
        Dim MyYear As String = MyDate.Substring(6, 4)
        Dim FMyDateAux As New System.DateTime(CInt(MyYear), CInt(MyMonth), CInt(MyDay))
        Dim fMyDate As DateTime = FMyDateAux.AddDays(Days)
        fMyDate.AddDays(Days)
        Dim MyDayr As String = (fMyDate.Day + 100).ToString.Substring(1, 2)
        Dim MyMonthr As String = (fMyDate.Month + 100).ToString.Substring(1, 2)
        Dim MyYearr As String = (fMyDate.Year).ToString
        AddDaysToDates = MyDayr & "/" & MyMonthr & "/" & MyYearr
    End Function
#End Region
#Region "TimeBetweenDates"
    Public Shared Function TimeBetweenDates(ByVal DateI As String, ByVal DateF As String) As Double
        If Not ValidateTxtDate(DateI) Then
            TimeBetweenDates = -92417511
            Exit Function
        End If
        If Not ValidateTxtDate(DateF) Then
            TimeBetweenDates = -92417511
            Exit Function
        End If
        TimeBetweenDates = 0
        Dim DayI As String = DateI.Substring(0, 2)
        Dim DayF As String = DateF.Substring(0, 2)
        Dim YearI As String = DateI.Substring(6, 4)
        Dim yearF As String = DateF.Substring(6, 4)
        Dim MonthI As String = DateI.Substring(3, 2)
        Dim MonthF As String = DateF.Substring(3, 2)
        Dim FDateI As New System.DateTime(ToInt(YearI), ToInt(MonthI), ToInt(DayI))
        Dim FDatef As New System.DateTime(ToInt(yearF), ToInt(MonthF), ToInt(DayF))
        Dim TimeDates As New System.TimeSpan(FDatef.Ticks - FDateI.Ticks)
        TimeBetweenDates = ToDouble(TimeDates.TotalDays.ToString)
    End Function
#End Region
#Region "MailValidate"
    Public Shared Function MailValidate(ByVal Mail As String) As Boolean
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
#Region "FileExt"
    Public Shared Function FileExt(ByVal File As String) As String
        FileExt = ""
        File = File.Trim
        Dim largo As Integer = File.Length
        Dim N As Integer
        For N = largo - 1 To 0 Step -1
            If File.Substring(N, 1) = "." Then
                FileExt = File.Substring(N, largo - N)
                Exit For
            End If
        Next
    End Function
#End Region
#Region "MinuteBetweenDates"
    Public Shared Function MinutesBetweenDates(ByVal Date1 As String, ByVal Date2 As String, ByVal Time1 As String, ByVal Time2 As String, ByVal ShowNegativeResult As Boolean) As Integer
        MinutesBetweenDates = 0
        Dim d1 As New DateTime(ToInt(Date1.Substring(6, 4)), ToInt(Date1.Substring(3, 2)), ToInt(Date1.Substring(0, 2)), ToInt(Time1.Trim.Substring(0, 2)), ToInt(Time1.Trim.Substring(3, 2)), ToInt(Time1.Trim.Substring(6, 2)))
        Dim d2 As New DateTime(ToInt(Date2.Substring(6, 4)), ToInt(Date2.Substring(3, 2)), ToInt(Date2.Substring(0, 2)), ToInt(Time2.Trim.Substring(0, 2)), ToInt(Time2.Trim.Substring(3, 2)), ToInt(Time2.Trim.Substring(6, 2)))
        Dim ts1 As TimeSpan = d1.TimeOfDay
        Dim ts2 As TimeSpan = d2.TimeOfDay
        Dim ts3 As TimeSpan = d2 - d1
        MinutesBetweenDates = CType(ts3.TotalMinutes, Integer)
        If Not ShowNegativeResult Then If MinutesBetweenDates < 0 Then MinutesBetweenDates = 0
    End Function
#End Region
#Region "DaysBetweenDates"
    Public Shared Function DaysBetweenDates(ByVal Date1 As String, ByVal Date2 As String, ByVal Time1 As String, ByVal Time2 As String, ByVal ShowNegativeResult As Boolean) As Integer
        If Time1.Trim.Length = 0 And Time2.Trim.Length = 0 Then
            Time1 = SystemTime()
            Time2 = SystemTime()
        End If
        DaysBetweenDates = 0
        Dim d1 As New DateTime(ToInt(Date1.Substring(6, 4)), ToInt(Date1.Substring(3, 2)), ToInt(Date1.Substring(0, 2)), ToInt(Time1.Trim.Substring(0, 2)), ToInt(Time1.Trim.Substring(3, 2)), ToInt(Time1.Trim.Substring(6, 2)))
        Dim d2 As New DateTime(ToInt(Date2.Substring(6, 4)), ToInt(Date2.Substring(3, 2)), ToInt(Date2.Substring(0, 2)), ToInt(Time2.Trim.Substring(0, 2)), ToInt(Time2.Trim.Substring(3, 2)), ToInt(Time2.Trim.Substring(6, 2)))
        Dim ts1 As TimeSpan = d1.TimeOfDay
        Dim ts2 As TimeSpan = d2.TimeOfDay
        Dim ts3 As TimeSpan = d2 - d1
        DaysBetweenDates = CType(ts3.TotalDays, Integer)
        If Not ShowNegativeResult Then If DaysBetweenDates < 0 Then DaysBetweenDates = 0
    End Function
#End Region
#Region "SecondsToFormatHour"
    Public Shared Function SecondsToFormatHour(ByVal seconds As Integer, ByVal whole_seconds As Boolean) As String
        Dim MySpan As New TimeSpan(0, 0, seconds)
        Dim txt As String = ""
        If MySpan.Days > 0 Then
            txt &= ", " & MySpan.Days.ToString() & " days"
            MySpan = MySpan.Subtract(New TimeSpan(MySpan.Days, 0, 0, 0))
        End If
        If MySpan.Hours > 0 Then
            txt &= ", " & MySpan.Hours.ToString() & " hours"
            MySpan = MySpan.Subtract(New TimeSpan(0, MySpan.Hours, 0, 0))
        End If
        If MySpan.Minutes > 0 Then
            txt &= ", " & MySpan.Minutes.ToString() & " " & "minutes"
            MySpan = MySpan.Subtract(New TimeSpan(0, 0, MySpan.Minutes, 0))
        End If
        If whole_seconds Then
            ' Display only whole seconds.
            If MySpan.Seconds > 0 Then
                txt &= ", " & MySpan.Seconds.ToString() & " " & "seconds"
            End If
        Else
            ' Display fractional seconds.
            txt &= ", " & MySpan.TotalSeconds.ToString() & " " & "seconds"
        End If
        ' Remove the leading ", ".
        If txt.Length > 0 Then txt = txt.Substring(2)
        ' Return the result.
        Return txt
    End Function
#End Region
#Region "MilSecondsToFormatHour"
    Public Shared Function MilSecondsToFormatHour(ByVal Milseconds As Integer, ByVal whole_seconds As Boolean) As String
        If Milseconds < 1000 Then
            Return "00:00:00:" & Milseconds.ToString.Trim
        End If
        Dim MySpan As New TimeSpan(0, 0, 0, 0, Milseconds)
        Dim txt As String = ""
        If MySpan.Hours > 0 Then
            txt &= MySpan.Hours.ToString()
            MySpan = MySpan.Subtract(New TimeSpan(0, MySpan.Hours, 0, 0))
        Else
            txt &= "00"
        End If
        If MySpan.Minutes > 0 Then
            txt &= ":" & (MySpan.Minutes + 100).ToString().Substring(1, 2)
            MySpan = MySpan.Subtract(New TimeSpan(0, 0, MySpan.Minutes, 0))
        Else
            txt &= ":00"
        End If
        If whole_seconds Then
            ' Display only whole seconds.
            If MySpan.Seconds > 0 Then
                txt &= ":" & (MySpan.Seconds + 100).ToString().Substring(1, 2)
            Else
                txt &= ":00"
            End If
        Else
            ' Display fractional seconds.
            txt &= ":" & MySpan.TotalSeconds.ToString()
        End If
        Return txt
    End Function
#End Region
#Region "MinutesToFormatHour"
    Public Shared Function MinutesToFormatHour(ByVal minutes As Integer) As String
        MinutesToFormatHour = ""
        Dim Hour As Integer = 0
        Do While minutes >= 60
            minutes = minutes - 60
            Hour = Hour + 1
        Loop
        MinutesToFormatHour = (Hour + 100).ToString.Substring(1, 2) & ":" & (minutes + 100).ToString.Substring(1, 2) & ":00"
    End Function
#End Region
#Region "RoundTime"
    Public Shared Function RoundTime(ByVal Time As String, ByVal Method As Integer) As String
        RoundTime = ""
        Select Case Method
            Case 1
                Dim Minutes As Integer
                Dim Hour As Integer
                Minutes = ToInt(Time.Substring(3, 2))
                Hour = ToInt(Time.Substring(0, 2))
                Select Case Minutes
                    Case 0, 1, 2, 3, 4, 5, 6, 7
                        Minutes = 0
                    Case 8, 9, 10, 11, 12, 13, 14, 15
                        Minutes = 15
                    Case 16, 17, 18, 19, 20, 21, 22, 23
                        Minutes = 15
                    Case 24, 25, 26, 27, 28, 29, 30
                        Minutes = 30
                    Case 31, 32, 33, 34, 35, 36, 37
                        Minutes = 30
                    Case 38, 39, 40, 41, 42, 43, 44, 45
                        Minutes = 45
                    Case 46, 47, 48, 49, 50, 51, 52
                        Minutes = 45
                    Case 53, 54, 55, 56, 57, 58, 59, 60
                        Minutes = 0
                        Hour = Hour + 1
                        If Hour = 24 Then Hour = 0
                End Select
                RoundTime = (Hour + 100).ToString.Substring(1, 2) & ":" & (Minutes + 100).ToString.Substring(1, 2) & ":00"
        End Select
    End Function
#End Region
#Region "ClearTmpFiles"
    Public Shared Sub ClearTmpFiles(ByVal Path As String, ByVal Pattern As String, ByVal AgingInDays As Integer)
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
#Region "ValidateTxtTime"
    Public Shared Function ValidateTxtTime(ByVal MyTime As String) As Boolean
        If MyTime.Trim.Length <> 8 Then Return False
        Dim MyHour As Integer = ToInt(MyTime.Substring(0, 2))
        Dim MyMin As Integer = ToInt(MyTime.Substring(3, 2))
        Dim MySec As Integer = ToInt(MyTime.Substring(6, 2))
        If MyHour > 23 Or MyHour < 0 Then Return False
        If MyMin > 59 Or MyMin < 0 Then Return False
        If MySec > 59 Or MySec < 0 Then Return False
        Return True
    End Function
#End Region
#Region "ValidMail"
    Public Shared Function ValidMail(ByVal Mail As String) As Boolean
        If Mail.Trim.Length <> 0 Then
            Dim Expresion As New Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")
            Dim Result As Match
            Result = Expresion.Match(Mail)
            Return Result.Success
        Else
            Return True
        End If
    End Function
#End Region
#Region "ValidUrl"
    Public Shared Function ValidUrl(ByVal Url As String) As Boolean
        If Url.Trim.Length <> 0 Then
            Dim Expresion As New Regex("http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?")
            Dim Result As Match
            Result = Expresion.Match(Url)
            Return Result.Success
        Else
            Return True
        End If
    End Function
#End Region
#Region "AmericanTime"
    Public Shared Function AmericanTime(ByVal Time As String, ByVal NotSeconds As Boolean) As String
        If Time.Trim.StartsWith("01") Then Time = Time & " am"
        If Time.Trim.StartsWith("02") Then Time = Time & " am"
        If Time.Trim.StartsWith("03") Then Time = Time & " am"
        If Time.Trim.StartsWith("04") Then Time = Time & " am"
        If Time.Trim.StartsWith("05") Then Time = Time & " am"
        If Time.Trim.StartsWith("06") Then Time = Time & " am"
        If Time.Trim.StartsWith("07") Then Time = Time & " am"
        If Time.Trim.StartsWith("08") Then Time = Time & " am"
        If Time.Trim.StartsWith("09") Then Time = Time & " am"
        If Time.Trim.StartsWith("10") Then Time = Time & " am"
        If Time.Trim.StartsWith("11") Then Time = Time & " am"
        If Time.Trim.StartsWith("00") Then
            Time = "12" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " am"
        End If
        If Time.Trim.StartsWith("12") Then
            Time = "12" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("13") Then
            Time = "01" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("14") Then
            Time = "02" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("15") Then
            Time = "03" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("16") Then
            Time = "04" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("17") Then
            Time = "05" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("18") Then
            Time = "06" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("19") Then
            Time = "07" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("20") Then
            Time = "08" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("21") Then
            Time = "09" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("22") Then
            Time = "10" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If Time.Trim.StartsWith("23") Then
            Time = "11" & Time.Trim.Substring(2, Time.Trim.Length - 2) & " pm"
        End If
        If NotSeconds Then
            Time = Time.Substring(0, 5) & " " & Time.Substring(9, 2)
        End If
        Return Time
    End Function
#End Region
#Region "ProtectCard"
    Public Shared Function ProtectCard(ByVal CardNumber As String) As String
        Try
            Dim n2 As Integer
            ProtectCard = ""
            For n2 = 0 To CardNumber.Trim.Length - 5
                ProtectCard = ProtectCard & "x"
            Next
            ProtectCard = ProtectCard & CardNumber.Trim.Substring(CardNumber.Trim.Length - 4, 4)
        Catch ex As Exception
            ProtectCard = "xxxxxxxxxxxxxxxx"
        End Try
    End Function
#End Region
#Region "CurrentTicks"
    Public Shared Function CurrentTicks() As Long
        Dim d As DateTime = DateTime.Now
        CurrentTicks = d.Ticks
    End Function
#End Region
#Region "AddMinutesToTime"
    Public Shared Function AddMinutesToTime(ByVal MyTime As String, ByVal Minutes As Integer, ByVal RoundSeconds As Boolean) As String
        If Not ValidateTxtTime(MyTime) Then
            AddMinutesToTime = "**:**:**"
            Exit Function
        End If
        Dim Myhour As String = MyTime.Substring(0, 2)
        Dim Myminutes As String = MyTime.Substring(3, 2)
        Dim Myseconds As String = MyTime.Substring(6, 2)
        Dim FMyDateAux As New System.DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, System.DateTime.Now.Day, ToInt(Myhour), ToInt(Myminutes), ToInt(Myseconds))
        Dim FMyDate As DateTime = FMyDateAux.AddMinutes(Minutes)
        Dim MyHourr As String = (FMyDate.Hour + 100).ToString.Substring(1, 2)
        Dim MyMinutesr As String = (FMyDate.Minute + 100).ToString.Substring(1, 2)
        Dim MySecondsr As String = (FMyDate.Second + 100).ToString.Substring(1, 2)
        If RoundSeconds Then MySecondsr = "00"
        AddMinutesToTime = MyHourr & ":" & MyMinutesr & ":" & MySecondsr
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
#Region "TimeToTxt"
    Public Shared Function TimeToTxt(ByVal TheDate As Date) As String
        Dim Myhour As String, MyMinutes As String, MySeconds As String
        Myhour = TheDate.Hour.ToString
        MyMinutes = TheDate.Minute.ToString
        MySeconds = TheDate.Second.ToString
        If Myhour.Length = 1 Then Myhour = "0" & Myhour
        If MyMinutes.Length = 1 Then MyMinutes = "0" & MyMinutes
        If Myhour.Length = 1 Then Myhour = "0" & Myhour
        If MySeconds.Length = 1 Then MySeconds = "0" & MySeconds
        TimeToTxt = Myhour & ":" & MyMinutes & ":" & MySeconds
    End Function
#End Region
#Region "GetDecimalPart"
    Private Shared Function GetDecimalPart(ByVal Amount As String, ByVal Decimals As Integer) As Long
        Dim Numero As String = Amount.ToString.Trim
        If Numero.Trim.Length = 0 Then Return 0
        Dim entero As New System.Text.StringBuilder
        Dim Mydecimals As New System.Text.StringBuilder
        Dim Divisor As New System.Text.StringBuilder
        Divisor.Append("1")
        Dim MyCharacter As String
        Dim flag As Boolean = False
        Dim n As Integer
        Dim DecimalsCount As Integer = 0
        For n = 0 To Numero.Trim.Length - 1
            MyCharacter = Left(Right(Numero, Numero.Length - n), 1)
            If Not IsNumeric(MyCharacter) And MyCharacter <> "-" And MyCharacter <> "+" Then
                flag = True
            Else
                If Not flag Then
                    entero.Append(MyCharacter)
                Else
                    DecimalsCount = DecimalsCount + 1
                    If DecimalsCount <= Decimals Then
                        Mydecimals.Append(MyCharacter)
                        Divisor.Append("0")
                    End If
                End If
            End If
        Next
        Select Case Mydecimals.ToString.Trim.Length
            Case 0
                GetDecimalPart = 0
            Case 1
                GetDecimalPart = Convert.ToInt64(Mydecimals.ToString & "0")
            Case Else
                GetDecimalPart = Convert.ToInt64((ToInt(Mydecimals.ToString) + 100).ToString.Substring(1, 2))
        End Select
    End Function
#End Region
#Region "GetIntegerPart"
    Private Shared Function GetIntegerPart(ByVal Amount As Double) As Long
        Dim Numero As String = Amount.ToString.Trim
        If Numero.Trim.Length = 0 Then Return 0
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
        GetIntegerPart = Convert.ToInt64(entero.ToString)
    End Function
#End Region
#Region "AmountInWords"
    Public Shared Function AmountInWords(ByVal Amount As Double, ByVal TxtDollars As Boolean, ByVal TxtExactly As Boolean) As String
        AmountInWords = ""
        If TxtExactly Then
            AmountInWords = "EXACTLY "
        End If
        If TxtDollars Then
            AmountInWords = AmountInWords & NumAText(GetIntegerPart(Amount)) & " DOLLARS AND "
        Else
            AmountInWords = AmountInWords & NumAText(GetIntegerPart(Amount)) & " AND "
        End If
        AmountInWords = AmountInWords & ((GetDecimalPart(Amount.ToString, 2) + 100).ToString.Substring(1, 2)).ToString & "/100"
    End Function
    Public Shared Function NumAText(ByVal Val As Long) As String
        NumAText = ""
        Dim AA As String
        Dim Val1 As Long
        Dim Val2 As Long
        AA = "EXACTLY DOLLARS "
        Val1 = 0
        Val2 = 0
        Select Case Val
            Case 0
                Return "ZERO"
            Case Is <= 1000
                AA = NumaTex1(Val, 0)
                Return AA
            Case Is < 1000000
                Val1 = GetIntegerPart(Val / 1000)
                Val2 = Val - Val1 * 1000
                If Val2 = 0 Then
                    AA = NumaTex1(Val1, 1) + " THOUSAND "
                Else
                    AA = NumaTex1(Val1, 1) + " THOUSAND " + NumaTex1(Val2, 0)
                End If
                Return AA
            Case Is < 2000000
                Val2 = Val - 1000000
                If Val2 = 0 Then
                    AA = "ONE MILLION "
                Else
                    AA = "ONE MILLION " + NumAText(Val2)
                End If
                Return (AA)
            Case Is < 1000000000000
                Val1 = GetIntegerPart(Val / 1000000)
                Val2 = Val - Val1 * 1000000
                If Val2 = 0 Then
                    AA = NumAText(Val1) + " MILLIONS "
                Else
                    AA = NumAText(Val1) + " MILLIONS " + NumAText(Val2)
                End If
                Return (AA)
            Case Else
                Return ("OVERFLOW")
        End Select
    End Function
    Public Shared Function NumaTex1(ByVal Cifra1 As Long, ByVal Miles As Double) As String
        NumaTex1 = ""
        Dim Cifra3 As Long
        Dim Cifra2 As Long
        Dim Txx As String
        Dim Txy As String
        Select Case Cifra1
            Case Is <= 100
                Select Case Cifra1
                    Case Is < 21
                        Select Case Cifra1
                            Case 1
                                If Miles = 0 Then
                                    Return "ONE"
                                Else
                                    Return "ONE"
                                End If
                            Case 2
                                Return "TWO"
                            Case 3
                                Return "THREE"
                            Case 4
                                Return "FOUR"
                            Case 5
                                Return "FIVE"
                            Case 6
                                Return "SIX"
                            Case 7
                                Return "SEVEN"
                            Case 8
                                Return "EIGHT"
                            Case 9
                                Return "NINE"
                            Case 10
                                Return "TEN"
                            Case 11
                                Return "ELEVEN"
                            Case 12
                                Return "TWELVE"
                            Case 13
                                Return "THIRTEEN"
                            Case 14
                                Return "FOURTEEN"
                            Case 15
                                Return "FIFTEEN"
                            Case 16
                                Return "SIXTEEN"
                            Case 17
                                Return "SEVENTEEN"
                            Case 18
                                Return "EIGHTEEN"
                            Case 19
                                Return "NINETEEN"
                            Case 20
                                Return "TWENTY"
                        End Select
                    Case Is < 30
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        Txx = NumaTex1(Cifra2, Miles)
                        Return "TWENTY " + Txx
                    Case Is < 40
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        If Cifra2 = 0 Then
                            Return "THIRTY"
                        Else
                            Txx = NumaTex1(Cifra2, Miles)
                            Return "THIRTY " + Txx
                        End If
                    Case Is < 50
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        If Cifra2 = 0 Then
                            Return "FORTY"
                        Else
                            Txx = NumaTex1(Cifra2, Miles)
                            Return "FORTY  " + Txx
                        End If
                    Case Is < 60
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        If Cifra2 = 0 Then
                            Return "FIFTY"
                        Else
                            Txx = NumaTex1(Cifra2, Miles)
                            Return "FIFTY " + Txx
                        End If
                    Case Is < 70
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        If Cifra2 = 0 Then
                            Return "SIXTY"
                        Else
                            Txx = NumaTex1(Cifra2, Miles)
                            Return "SIXTY " + Txx
                        End If
                    Case Is < 80
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        If Cifra2 = 0 Then
                            Return "SEVENTY"
                        Else
                            Txx = NumaTex1(Cifra2, Miles)
                            Return "SEVENTY " + Txx
                        End If
                    Case Is < 90
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        If Cifra2 = 0 Then
                            Return "EIGHTY"
                        Else
                            Txx = NumaTex1(Cifra2, Miles)
                            Return "EIGHTY " + Txx
                        End If
                    Case Is < 100
                        Cifra2 = Cifra1 - GetIntegerPart(Cifra1 / 10) * 10
                        If Cifra2 = 0 Then
                            Return "NINETY"
                        Else
                            Txx = NumaTex1(Cifra2, Miles)
                            Return "NINETY " + Txx
                        End If
                    Case Else
                        Return "ONE HUNDRED"
                End Select
            Case Is < 200
                Cifra3 = Cifra1 - GetIntegerPart(Cifra1 / 100) * 100
                Txx = NumaTex1(Cifra3, Miles)
                Return "ONE HUNDRED " + Txx
            Case Is < 500
                Cifra3 = GetIntegerPart(Cifra1 / 100)
                Txy = NumaTex1(Cifra3, Miles)
                Cifra3 = Cifra1 - Cifra3 * 100
                If Cifra3 = 0 Then
                    Txx = ""
                Else
                    Txx = NumaTex1(Cifra3, Miles)
                End If
                Return Txy + " HUNDRED " + Txx
            Case Is < 600
                Cifra3 = Cifra1 - GetIntegerPart(Cifra1 / 100) * 100
                If Cifra3 = 0 Then
                    Txx = ""
                Else
                    Txx = NumaTex1(Cifra3, Miles)
                End If
                Return "FIVE HUNDRED " + Txx
            Case Is < 700
                Cifra3 = Cifra1 - GetIntegerPart(Cifra1 / 100) * 100
                If Cifra3 = 0 Then
                    Txx = ""
                Else
                    Txx = NumaTex1(Cifra3, Miles)
                End If
                Return ("SIX HUNDRED " + Txx)
            Case Is < 800
                Cifra3 = Cifra1 - GetIntegerPart(Cifra1 / 100) * 100
                If Cifra3 = 0 Then
                    Txx = ""
                Else
                    Txx = NumaTex1(Cifra3, Miles)
                End If
                Return "SEVEN HUNDRED " + Txx
            Case Is < 900
                Cifra3 = Cifra1 - GetIntegerPart(Cifra1 / 100) * 100
                If Cifra3 = 0 Then
                    Txx = ""
                Else
                    Txx = NumaTex1(Cifra3, Miles)
                End If
                Return "EIGHT HUNDRED " + Txx
            Case Is < 1000
                Cifra3 = Cifra1 - GetIntegerPart(Cifra1 / 100) * 100
                If Cifra3 = 0 Then
                    Txx = ""
                Else
                    Txx = NumaTex1(Cifra3, Miles)
                End If
                Return "NINE HUNDRED " + Txx
            Case Else
                Return "ONE THOUSAND"
        End Select
    End Function
#End Region
#Region "TwipsToPixels"
    Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hdc As Int32, ByVal nIndex As Int32) As Int32
    Const WU_LOGPIXELSX As Int32 = 88
    Const WU_LOGPIXELSY As Int32 = 90
    Const TwipsPerInch As Int32 = 1440
    Public Shared Function TwipsToPixels(ByVal Twips As Int32, ByVal IsHorizontal As Boolean) As Int32
        Dim PixelsPerInch As Int32
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero)
        Dim DC As IntPtr = g.GetHdc
        Dim intDC As Int32 = DC.ToInt32

        If IsHorizontal Then
            PixelsPerInch = GetDeviceCaps(intDC, WU_LOGPIXELSX)
        Else
            PixelsPerInch = GetDeviceCaps(intDC, WU_LOGPIXELSY)
        End If
        g.ReleaseHdc(DC)

        Return CType((Twips / TwipsPerInch) * PixelsPerInch, Int32)
    End Function
#End Region
#Region "PixelsToTwips"
    Public Shared Function PixelsToTwips(ByVal Pixels As Int32, _
    ByVal IsHorizontal As Boolean) As Int32
        Dim PixelsPerInch As Int32
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero)
        Dim DC As IntPtr = g.GetHdc
        Dim intDC As Int32 = DC.ToInt32
        If IsHorizontal Then
            PixelsPerInch = GetDeviceCaps(intDC, WU_LOGPIXELSX)
        Else
            PixelsPerInch = GetDeviceCaps(intDC, WU_LOGPIXELSY)
        End If
        Dim numInches As Double = Pixels / PixelsPerInch
        Return CType(numInches * TwipsPerInch, Int32)
    End Function
#End Region
#Region "MilesFormat"
    Public Shared Function MilesFormat(ByVal Number As String, ByVal AmericanFormat As Boolean) As String
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
        If MilesFormat.Trim = "0." Then MilesFormat = "0"
    End Function
#End Region
#Region "WithInvalidCharacters"
    Public Shared Function WithInvalidCharacters(ByVal MyText As String) As Boolean
        Dim n As Integer
        For n = 0 To MyText.Trim.Length - 1
            Select Case MyText.Substring(n, 1)
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
#Region "SumTimes"
    Public Shared Function SumTimes(ByVal Time1 As String, ByVal Time2 As String) As String
        SumTimes = "00:00:00"
        Dim args() As String
        args = Time1.Split(CType(":", Char))
        If args.Length <> 3 Then Exit Function
        Dim Myhour As Integer = ToInt(args(0))
        Dim Myminutes As Integer = ToInt(args(1))
        Dim Myseconds As Integer = ToInt(args(2))
        args = Time2.Split(CType(":", Char))
        If args.Length <> 3 Then Exit Function
        Myhour = Myhour + ToInt(args(0))
        Myminutes = Myminutes + ToInt(args(1))
        Myseconds = Myseconds + ToInt(args(2))
        While 1 = 1
            If Myseconds >= 60 Then
                Myseconds = Myseconds - 60
                Myminutes = Myminutes + 1
            Else
                Exit While
            End If
        End While
        While 1 = 1
            If Myminutes >= 60 Then
                Myminutes = Myminutes - 60
                Myhour = Myhour + 1
            Else
                Exit While
            End If
        End While
        SumTimes = Myhour.ToString & ":" & (Myminutes + 100).ToString.Substring(1, 2) & ":" & (Myseconds + 100).ToString.Substring(1, 2)
        args = Nothing
    End Function
#End Region
#Region "SumFieldsTimes"
    Public Shared Function SumFieldsTimes(ByVal Data As Data.DataTable, ByVal Field As String) As String
        SumFieldsTimes = "00:00:00"
        Dim n As Integer
        With Data
            For n = 0 To .Rows.Count - 1
                If Data.Columns(Field).DataType Is System.Type.GetType("System.String") Then
                    SumFieldsTimes = SumTimes(SumFieldsTimes, Data.Rows(n)(Field).ToString)
                Else
                    SumFieldsTimes = SumTimes(SumFieldsTimes, MinutesToFormatHour(ToInt(Data.Rows(n)(Field).ToString)))
                End If
            Next
        End With
    End Function
#End Region
#Region "KillExcel"
    Public Shared Sub KillExcel()
        Try
            Dim mp As System.Diagnostics.Process() = Process.GetProcessesByName("EXCEL")
            Dim ExcelList As Process
            For Each ExcelList In mp
                ExcelList.Kill()
            Next ExcelList
            mp = Nothing
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "KillWord"
    Public Shared Sub KillWord()
        Try
            Dim mp As System.Diagnostics.Process() = Process.GetProcessesByName("WinWord")
            Dim WinWordList As Process
            For Each WinWordList In mp
                WinWordList.Kill()
            Next WinWordList
            mp = Nothing
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "KillInfoPath"
    Public Shared Sub KillInfoPath()
        Try
            Dim mp As System.Diagnostics.Process() = Process.GetProcessesByName("INFOPATH")
            Dim InfoPathList As Process
            For Each InfoPathList In mp
                InfoPathList.Kill()
            Next InfoPathList
            mp = Nothing
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "DatesList"
    Public Shared Function DatesList(ByVal StartDay As Integer, ByVal StartMonth As Integer, ByVal StartYear As Integer, ByVal EndDay As Integer, ByVal EndMonth As Integer, ByVal EndYear As Integer, ByVal AmericanFormat As Boolean) As DataTable
        DatesList = New DataTable
        DatesList.Columns.Add("DateTxt", System.Type.GetType("System.String"))
        Dim StartDate As New DateTime(StartYear, StartMonth, StartDay)
        Dim EndDate As New DateTime(EndYear, EndMonth, EndDay)
        If EndDate >= StartDate Then
            Dim Args(0) As String
            Do While 1 = 1
                If StartDate > EndDate Then
                    Exit Do
                Else
                    If AmericanFormat Then
                        Args(0) = AmericanDate(DateToTxt(StartDate))
                    Else
                        Args(0) = DateToTxt(StartDate)
                    End If
                    DatesList.Rows.Add(Args)
                    StartDate = StartDate.AddDays(1)
                End If
            Loop
            Args = Nothing
        End If
    End Function
#End Region
#Region "Alltrim"
    Public Shared Function Alltrim(ByVal Text As String) As String
        Alltrim = ""
        Text = Text.Trim
        Dim n As Integer = 0
        Dim Character As String = ""
        For n = 0 To Text.Length - 1
            Character = Text.Substring(n, 1)
            If Character <> " " Then Alltrim = Alltrim & Character
        Next
    End Function
#End Region
#Region "ElapsedTime"
    Public Shared Function ElapsedTime(ByVal Time1 As String, ByVal Time2 As String) As String
        Try
            ElapsedTime = ""
            Time1 = Time1.Trim
            Time2 = Time2.Trim
            Dim Sec1 As Integer = ToInt(Time1.Substring(6, 2))
            Dim Min1 As Integer = ToInt(Time1.Substring(3, 2))
            Dim Hour1 As Integer = ToInt(Time1.Substring(0, 2))
            Dim Sec2 As Integer = ToInt(Time2.Substring(6, 2))
            Dim Min2 As Integer = ToInt(Time2.Substring(3, 2))
            Dim Hour2 As Integer = ToInt(Time2.Substring(0, 2))
            Dim SecResult As Integer = 0
            Dim MinResult As Integer = 0
            Dim HourResult As Integer = 0
            While 1 = 1
                If Min1 = Min2 And Sec1 = Sec2 And Hour1 = Hour2 Then Exit While
                Sec1 = Sec1 + 1
                SecResult = SecResult + 1
                If Sec1 = 60 Then
                    Sec1 = 0
                    Min1 = Min1 + 1
                    If Min1 = 60 Then
                        Min1 = 0
                        Hour1 = Hour1 + 1
                    End If
                End If
            End While
            While 1 = 1
                If SecResult < 60 Then Exit While
                SecResult = SecResult - 60
                MinResult = MinResult + 1
                If MinResult = 60 Then
                    MinResult = 0
                    HourResult = HourResult + 1
                End If
            End While
            Return (HourResult + 100).ToString.Substring(1, 2) & ":" & (MinResult + 100).ToString.Substring(1, 2) & ":" & (SecResult + 100).ToString.Substring(1, 2)
        Catch ex As Exception
            ElapsedTime = "??:??:??"
        End Try
    End Function
#End Region
#Region "HourInRange"
    Public Shared Function HourInRange(ByVal Hour As String, ByVal HourFrom As String, ByVal HourTo As String) As Boolean
        HourInRange = False
        Try
            Hour = Hour.Trim
            HourFrom = HourFrom.Trim
            HourTo = HourTo.Trim
            Dim VSecFrom As Integer = ToInt(HourFrom.Substring(6, 2))
            Dim VMinFrom As Integer = ToInt(HourFrom.Substring(3, 2))
            Dim VHourFrom As Integer = ToInt(HourFrom.Substring(0, 2))
            Dim VSecTo As Integer = ToInt(HourTo.Substring(6, 2))
            Dim VMinTo As Integer = ToInt(HourTo.Substring(3, 2))
            Dim VHourTo As Integer = ToInt(HourTo.Substring(0, 2))
            Dim VSec As Integer = ToInt(Hour.Substring(6, 2))
            Dim VMin As Integer = ToInt(Hour.Substring(3, 2))
            Dim VHour As Integer = ToInt(Hour.Substring(0, 2))

            Dim Time1 As New DateTime(1971, 6, 7, VHour, VMin, VSec)
            Dim Time2 As New DateTime(1971, 6, 7, VHourFrom, VMinFrom, VSecFrom)
            Dim Time3 As New DateTime(1971, 6, 7, VHourTo, VMinTo, VSecTo)
            If Time2 > Time1 Then Time2 = New DateTime(1971, 6, 6, VHourTo, VMinTo, VSecTo)
            If Time3 < Time1 Then Time3 = New DateTime(1971, 6, 8, VHourTo, VMinTo, VSecTo)
            If Time1 >= Time2 And Time1 <= Time3 Then HourInRange = True
        Catch ex As Exception
        End Try
    End Function
#End Region
#Region "StrToPOSTNET"
    Public Shared Function StrToPostNet(ByVal Zipcode5 As Integer, ByVal Zipcode4 As Integer, ByVal DeliveryPoint As Integer) As String
        'POSTNET (Postal Numeric Encoding Technique)
        StrToPostNet = "*" & (Zipcode5 + 100000).ToString.Substring(1, 5) & (Zipcode4 + 10000).ToString.Substring(1, 4) & (DeliveryPoint + 1000).ToString.Substring(1, 3) & "*"
    End Function
#End Region
#Region "DivisionMod"
    Public Shared Function DivisionMod(ByVal Value As Double, ByVal Value2 As Double) As Integer
        Dim Result As String = (Value / Value2).ToString.Replace(",", ".")
        If Result.IndexOf(".") >= 0 Then
            DivisionMod = ToInt(Result.Substring(Result.IndexOf(".") + 1, Result.Length - Result.IndexOf(".") - 1))
        Else
            DivisionMod = 0
        End If
    End Function
#End Region
#Region "GiveMeDecimals"
    Public Shared Function GiveMeDecimals(ByVal Valor As String) As Long
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
        Return Tolong(Decimales)
    End Function
#End Region
#Region "GeneratePassword"
    Public Shared Function GeneratePassword() As String
        Dim r As Random = New Random
        GeneratePassword = r.Next(1, 9999).ToString
        Dim n As Integer
        For n = 1 To 4
            Select Case r.Next(1, 10).ToString
                Case "1"
                    GeneratePassword = GeneratePassword & "Z"
                Case "2"
                    GeneratePassword = GeneratePassword & "X"
                Case "3"
                    GeneratePassword = GeneratePassword & "C"
                Case "4"
                    GeneratePassword = GeneratePassword & "V"
                Case "5"
                    GeneratePassword = GeneratePassword & "Y"
                Case "6"
                    GeneratePassword = GeneratePassword & "P"
                Case "7"
                    GeneratePassword = GeneratePassword & "W"
                Case "8"
                    GeneratePassword = GeneratePassword & "A"
                Case "9"
                    GeneratePassword = GeneratePassword & "U"
                Case "10"
                    GeneratePassword = GeneratePassword & "P"
            End Select
        Next
    End Function
#End Region
#Region "GetJulianDate"
    Public Shared Function GetJulianDate(ByVal pdtmDate As DateTime) As Int64
        Dim Temp As New TimeSpan(pdtmDate.Ticks)
        GetJulianDate = CLng(Temp.TotalDays + 1721426)
    End Function
#End Region
#Region "Asc"
    Public Shared Function Asc(ByVal Txt As String) As Integer
        Return Microsoft.VisualBasic.Asc(Txt)
    End Function
#End Region
#Region "Enter"
    Public Shared Function Enter() As String
        Return Convert.ToChar(13) & Convert.ToChar(10)
    End Function
#End Region
#Region "RemoteDebug"
    Public Shared Sub RemoteDebug(ByVal Value As String)
        Dim MyFile As New IO.StreamWriter("c:\MyDebug.txt", True)
        MyFile.WriteLine(Value)
        MyFile.Close()
        MyFile = Nothing
    End Sub
#End Region
#Region "TimeBetweenDatesAndTime"
    Public Shared Function TimeBetweenDatesAndTime(ByVal StartDate As String, ByVal EndDate As String, ByVal StartTime As String, ByVal EndTime As String, ByVal Language As Integer) As String
        Try
            If Not ValidateTxtDate(StartDate) Then
                TimeBetweenDatesAndTime = ""
                Exit Function
            End If
            If Not ValidateTxtTime(StartTime) Then
                TimeBetweenDatesAndTime = ""
                Exit Function
            End If
            If Not ValidateTxtDate(EndDate) Then
                TimeBetweenDatesAndTime = ""
                Exit Function
            End If
            TimeBetweenDatesAndTime = ""
            Dim DayS As Integer = CInt(StartDate.Substring(0, 2))
            Dim DayE As Integer = CInt(EndDate.Substring(0, 2))
            Dim YearS As Integer = CInt(StartDate.Substring(6, 4))
            Dim yearE As Integer = CInt(EndDate.Substring(6, 4))
            Dim MonthS As Integer = CInt(StartDate.Substring(3, 2))
            Dim MonthE As Integer = CInt(EndDate.Substring(3, 2))
            Dim HourS As Integer = CInt(StartTime.Substring(0, 2))
            Dim HourE As Integer = CInt(EndTime.Substring(0, 2))
            Dim MinuteS As Integer = CInt(StartTime.Substring(3, 2))
            Dim MinuteE As Integer = CInt(EndTime.Substring(3, 2))
            Dim SecondS As Integer = CInt(StartTime.Substring(6, 2))
            Dim SecondE As Integer = CInt(EndTime.Substring(6, 2))
            Dim FDateS As New System.DateTime(YearS, MonthS, DayS, HourS, MinuteS, SecondS)
            Dim FDateE As New System.DateTime(yearE, MonthE, DayE, HourE, MinuteE, SecondE)
            Dim TimeDates As New System.TimeSpan(FDateE.Ticks - FDateS.Ticks)
            Dim TSecond As Double = TimeDates.TotalSeconds
            Dim Tminutes As Double = 0
            Dim Tdays As Double = 0
            Dim Thour As Double = 0
            Dim DaysTxt As String = " days "
            Dim HourTxt As String = " hours "
            Dim SecsTxt As String = " secs"
            If Language = 2 Then
                HourTxt = " horas "
                SecsTxt = " segs"
                DaysTxt = " días "
            End If
            If TSecond < 60 Then
                TimeBetweenDatesAndTime = TSecond.ToString & SecsTxt
            Else
                Do While TSecond >= 60
                    Tminutes = Tminutes + 1
                    TSecond = TSecond - 60
                Loop
                If Tminutes < 60 Then
                    TimeBetweenDatesAndTime = Tminutes.ToString & " min " & TSecond.ToString & SecsTxt
                Else
                    Do While Tminutes >= 60
                        Thour = Thour + 1
                        Tminutes = Tminutes - 60
                    Loop
                    Do While Thour >= 24
                        Tdays = Tdays + 1
                        Thour = Thour - 24
                    Loop
                    If Thour > 0 Then
                        TimeBetweenDatesAndTime = Tdays.ToString & DaysTxt & Thour.ToString & HourTxt & Tminutes.ToString & " min " & TSecond.ToString & SecsTxt
                    Else
                        TimeBetweenDatesAndTime = Thour.ToString & HourTxt & Tminutes.ToString & " min " & TSecond.ToString & SecsTxt
                    End If
                End If
            End If
        Catch ex As Exception
            TimeBetweenDatesAndTime = ""
        End Try
    End Function
#End Region
#Region "SendToKeyBoard"
    Public Shared Sub SendToKeyBoard(ByVal Value As String)
        My.Computer.Keyboard.SendKeys(Value, True)
    End Sub
#End Region
#Region "IsAdministrator"
    Public Shared Function isAdministrator() As Boolean
        ' Check if the user is authenticated before continuing.
        If My.User.IsAuthenticated Then
            ' If the user is in the administrators group.
            If My.User.IsInRole("Administrators") OrElse My.User.IsInRole("Administrador") Then
                Return True
            End If
        End If
        ' Return false because the user isn't an administrator,
        ' or authenticated.
        Return False
    End Function
#End Region
#Region "KillAcrobat"
    Public Shared Sub KillAcrobat()
        Try
            Dim mp As System.Diagnostics.Process() = Process.GetProcessesByName("Acrobat")
            Dim AcrobatList As Process
            For Each AcrobatList In mp
                AcrobatList.Kill()
            Next AcrobatList
            mp = Nothing
        Catch ex As Exception
        End Try
    End Sub
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
#Region "ConvertSpecialChars"
    Public Shared Function ConvertSpecialChars(ByVal Text As String) As String
        ConvertSpecialChars = Text
        ConvertSpecialChars = ConvertSpecialChars.Replace("á", "a")
        ConvertSpecialChars = ConvertSpecialChars.Replace("é", "e")
        ConvertSpecialChars = ConvertSpecialChars.Replace("í", "i")
        ConvertSpecialChars = ConvertSpecialChars.Replace("ó", "o")
        ConvertSpecialChars = ConvertSpecialChars.Replace("ú", "u")
        ConvertSpecialChars = ConvertSpecialChars.Replace("Á", "A")
        ConvertSpecialChars = ConvertSpecialChars.Replace("É", "E")
        ConvertSpecialChars = ConvertSpecialChars.Replace("Í", "I")
        ConvertSpecialChars = ConvertSpecialChars.Replace("Ó", "O")
        ConvertSpecialChars = ConvertSpecialChars.Replace("Ú", "U")
        ConvertSpecialChars = ConvertSpecialChars.Replace("ñ", "n")
        ConvertSpecialChars = ConvertSpecialChars.Replace("Ñ", "n")
    End Function
#End Region
#Region "GoogleMapReferences"
    Public Shared Function GoogleMapReferences(ByVal Address As String) As XmlDocument
        GoogleMapReferences = New XmlDocument
        Try
            Dim url As String = "https://maps.google.com.ar/maps/api/geocode/xml?address=" & Address & "&sensor=true&key=AIzaSyBqBTl02kWThbJICaDvKI2KcF3rnwnRuVs"
            Dim request As HttpWebRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim reader As New StreamReader(response.GetResponseStream)
            GoogleMapReferences.LoadXml(reader.ReadToEnd())
            request = Nothing
            response = Nothing
            reader.Dispose()
        Catch ex As Exception
        End Try
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
#Region "IPwhois"
    Public Shared Function IPwhois(ByVal Ip As String) As Data.DataTable
        IPwhois = New DataTable
        With IPwhois
            .Columns.Add("inetnum", Type.GetType("System.String"))
            .Columns.Add("owner", Type.GetType("System.String"))
            .Columns.Add("ownerid", Type.GetType("System.String"))
            .Columns.Add("responsible", Type.GetType("System.String"))
            .Columns.Add("address", Type.GetType("System.String"))
            .Columns.Add("country", Type.GetType("System.String"))
            .Columns.Add("phone", Type.GetType("System.String"))
        End With
        Try
            Dim url As String = "http://www.ipmango.com/whois.php?ip=" & Ip
            Dim request As HttpWebRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim reader As New StreamReader(response.GetResponseStream)
            Dim Result As String = reader.ReadToEnd()
            If Result.IndexOf("inetnum") <> -1 Then
                Result = Result.Substring(Result.IndexOf("inetnum"))
                Dim Result2 As String = Result.Replace("<br />", "|")
                Dim Array() As String
                Array = Result2.Split(CChar("|"))
                IPwhois.Rows.Add(IPwhois.NewRow)
                If Array.Length >= 15 Then
                    For n As Integer = 0 To 14
                        If Array(n).IndexOf("inetnum:") <> -1 Then IPwhois.Rows(0)("inetnum") = Array(n).Replace("inetnum:", "").Trim
                        If Array(n).IndexOf("owner:") <> -1 Then IPwhois.Rows(0)("owner") = Array(n).Replace("owner:", "").Trim
                        If Array(n).IndexOf("ownerid:") <> -1 Then IPwhois.Rows(0)("ownerid") = Array(n).Replace("ownerid:", "").Trim
                        If Array(n).IndexOf("responsible:") <> -1 Then IPwhois.Rows(0)("responsible") = Array(n).Replace("responsible:", "").Trim
                        If Array(n).IndexOf("address:") <> -1 Then IPwhois.Rows(0)("address") = IPwhois.Rows(0)("address").ToString.Trim & " " & Array(n).Replace("address:", "").Trim
                        If Array(n).IndexOf("country:") <> -1 Then IPwhois.Rows(0)("country") = Array(n).Replace("country:", "").Trim
                        If Array(n).IndexOf("phone:") <> -1 Then IPwhois.Rows(0)("phone") = Array(n).Replace("phone:", "").Trim.Replace("]", "").Replace("[", "")
                    Next
                End If
            End If
            request = Nothing
            response = Nothing
            reader.Dispose()
        Catch ex As Exception
        End Try
    End Function
#End Region
#Region "DireccionLimpia"
    Public Shared Function DireccionLimpia(ByVal Direccion As String) As String
        Direccion = Direccion.Replace("|", "%")
        Direccion = FunctionLibrary.Functions.HtmlDecode(Direccion)
        Dim PrimerNumeroEncontrado As Boolean = False
        Dim NuevaDireccion As String = ""
        Dim ExtraRemovido As Boolean = False
        Dim FinNumeracion As Boolean
        For n As Integer = 0 To Direccion.Length - 1
            If ToInt(Direccion.Substring(n, 1)) > 0 And ExtraRemovido = False Then
                PrimerNumeroEncontrado = True
            End If
            If PrimerNumeroEncontrado And (ToInt(Direccion.Substring(n, 1)) = 0 And Direccion.Substring(n, 1) <> "0") And ExtraRemovido = False Then
                If Direccion.Substring(n, 1) = "," Then
                    NuevaDireccion = NuevaDireccion & Direccion.Substring(n, 1)
                    ExtraRemovido = True
                End If
                FinNumeracion = True
            Else
                If ToInt(Direccion.Substring(n, 1)) > 0 And FinNumeracion Then
                Else
                    NuevaDireccion = NuevaDireccion & Direccion.Substring(n, 1)
                End If
            End If
        Next
        DireccionLimpia = NuevaDireccion
    End Function
#End Region
#Region "PutJARAChars"
    Public Shared Function PutJARAChars(ByVal Text As String) As String
        If Text.Trim <> "" Then
            Text = Text.Replace("á", "JAR1")
            Text = Text.Replace("Á", "JAR2")
            Text = Text.Replace("é", "JAR3")
            Text = Text.Replace("É", "JAR4")
            Text = Text.Replace("í", "JAR5")
            Text = Text.Replace("Í", "JAR6")
            Text = Text.Replace("ó", "JAR7")
            Text = Text.Replace("Ó", "JAR8")
            Text = Text.Replace("ú", "JAR9")
            Text = Text.Replace("Ú", "JARA10")
            Text = Text.Replace("ñ", "JARA11")
            Text = Text.Replace("Ñ", "JARA12")
        End If
        Return Text
    End Function
#End Region
#Region "ValidateTxtDate"
    Public Shared Function ExcelDateInCsv(ByVal MyDate As String) As String
        If MyDate.Trim.Length <> 10 Then
            Return MyDate
        End If
        Dim MyDay As String = MyDate.Substring(0, 2)
        Dim MyMonth As String = MyDate.Substring(3, 2)
        Dim MyYear As String = MyDate.Substring(6, 4)
        Return MyYear & "-" & MyMonth & "-" & MyDay
    End Function
#End Region
#Region "DateXlsFormat"
    Public Shared Function DateXlsFormat(ByVal TheDate As Date) As String
        DateXlsFormat = TheDate.Year & "-" & (TheDate.Month + 100).ToString.Trim.Substring(1, 2) & "-" & (TheDate.Day + 100).ToString.Trim.Substring(1, 2)
    End Function
#End Region
End Class