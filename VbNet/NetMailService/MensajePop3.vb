Option Explicit On 
Option Strict On
Imports System.IO
Public Class MensajePop3
#Region "Variables"
    Public Numero As Long
    Public Bytes As Long
    Public Traido As Boolean
    Public Mensaje As String
    Private vDesde As String
    Private vPara As String
    Private vBody As String
    Private vFormatoMIME As Boolean
    Private vSeparadorMIME As String
    Private vFechaEnTicks As Long
#End Region
#Region "Propiedades"
    Public ReadOnly Property FormatoMIME() As Boolean
        Get
            FormatoMIME = vFormatoMIME
        End Get
    End Property
    Public ReadOnly Property Desde() As String
        Get
            Desde = vDesde
        End Get
    End Property
    Public ReadOnly Property Para() As String
        Get
            Para = vPara
        End Get
    End Property
    Public ReadOnly Property Body() As String
        Get
            Body = vBody
        End Get
    End Property
    Public ReadOnly Property SeparadorMIME() As String
        Get
            SeparadorMIME = vSeparadorMIME
        End Get
    End Property
    Public ReadOnly Property FechaEnTicks() As Long
        Get
            FechaEnTicks = vFechaEnTicks
        End Get
    End Property
#End Region
#Region "CargarMensajeDesdeTexto"
    Public Sub CargarMensajeDesdeTexto(ByVal Texto As String, ByVal Archivo As String)
        Dim Strpath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location) & "\"
        Dim MyFile As New System.io.StreamWriter(Strpath & Archivo)
        MyFile.Write(Texto)
        MyFile.Close()
        CargarMensajeDesdeArchivo(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location), Archivo)
    End Sub
#End Region
#Region "CargarMensajeDesdeArchivo"
    Public Sub CargarMensajeDesdeArchivo(ByVal Path As String, ByVal File As String)
        Try
            Dim Ubicacion As New CDO.DropDirectory
            Dim ListaMensajes As CDO.IMessages
            Dim Mensaje As New CDO.Message
            Dim ElMensaje As CDO.IMessage

            Dim n As Integer = 0
            Dim NombreArchivo As String
            Dim Posicion As Integer = -1
            ListaMensajes = Ubicacion.GetMessages(Path)
            For Each Mensaje In ListaMensajes
                n = n + 1
                NombreArchivo = ListaMensajes.FileName(Mensaje)
                If NombreArchivo = Path & "\" & File Then
                    Posicion = n
                End If
            Next
            If Posicion = -1 Then
                Throw New System.Exception("El mail " & Path & "\" & File & " no fue encontrado")
            Else
                ElMensaje = ListaMensajes.Item(Posicion)
                'If ElMensaje.Attachments.Count <> 0 Then ElMensaje.Attachments.Item(2).SaveToFile("c:\" & ElMensaje.Attachments.Item(2).FileName)
                Traido = True
                vDesde = ElMensaje.From
                vPara = ElMensaje.To
                vBody = ElMensaje.HTMLBody
                vFormatoMIME = ElMensaje.MimeFormatted
                vFechaEnTicks = ElMensaje.ReceivedTime.Ticks
                'ListaMensajes.Delete(Posicion)
            End If
            ListaMensajes = Nothing
            Ubicacion = Nothing
            Mensaje = Nothing
            ElMensaje = Nothing
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub
#End Region
#Region "DamePrimeraLineaDelTexto"
    Private Function DamePrimeraLineaDelTexto(ByVal Texto As String) As String
        Dim n As Integer
        Dim Encontre13 As Boolean = False
        Dim Encontre10 As Boolean = False
        Dim ValorAscii As Integer
        Dim Caracter As Char
        DamePrimeraLineaDelTexto = ""
        For n = 0 To Texto.Length - 1
            Caracter = CType(Texto.Substring(n, 1), Char)
            ValorAscii = Convert.ToInt32(CType(Caracter, Char))
            If ValorAscii = 13 Then
                Encontre13 = True
            Else
                If ValorAscii = 10 Then
                    Encontre10 = True
                Else
                    DamePrimeraLineaDelTexto = DamePrimeraLineaDelTexto & Caracter
                End If
            End If
            If Encontre10 And Encontre13 Then Exit For
        Next
        DamePrimeraLineaDelTexto = DamePrimeraLineaDelTexto.Trim
    End Function
#End Region
End Class
