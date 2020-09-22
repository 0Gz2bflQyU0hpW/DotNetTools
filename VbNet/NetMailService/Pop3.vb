Option Explicit On 
Option Strict On
Imports System.Net.Sockets
Imports Microsoft.VisualBasic
Public Class Pop3
    Inherits System.Net.Sockets.TcpClient
#Region "Connect"
    Public Overloads Sub Connect(ByVal server As String, ByVal username As String, ByVal password As String)
        Try
            Dim message, response As String
            Dim o As New System.Net.Sockets.TcpClient
            Connect(server, 110)
            response = Respond()
            If response.Substring(0, 3) <> "+OK" Then
                Throw New System.Exception(response)
            End If
            message = "USER " + username + vbCrLf
            write(message)
            response = Respond()
            If response.Substring(0, 3) <> "+OK" Then
                Throw New System.Exception(response)
            End If
            message = "PASS " + password + vbCrLf
            write(message)
            response = Respond()
            If response.Substring(0, 3) <> "+OK" Then
                Throw New System.Exception(response)
            End If
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub
#End Region
#Region "Borrar"
    Public Sub Borrar(ByVal rhs As MensajePop3)
        Dim message As String
        Dim response As String
        message = "DELE " & rhs.Numero.ToString & vbCrLf
        write(message)
        response = Respond()
    End Sub
#End Region
#Region "Desconectar"
    Public Sub Desconectar()
        Dim message, response As String
        message = "QUIT" + ControlChars.CrLf
        write(message)
        response = Respond()
        If response.Substring(0, 3) <> "+OK" Then
            Throw New System.Exception(response)
        End If
    End Sub
#End Region
#Region "TraerLista"
    Public Function TraerLista() As ArrayList
        TraerLista = New ArrayList
        Dim veri As String()
        Dim msg As MensajePop3
        Dim message, response As String
        Dim retval As New ArrayList
        message = "LIST" + vbCrLf
        write(message)
        response = Respond()
        If response.Substring(0, 3) <> "+OK" Then
            Throw New System.Exception(response)
        End If
        While (1 = 1)
            response = Respond()
            If response = "." + vbCrLf Then
                Return retval
            Else
                msg = New MensajePop3
                veri = response.Split(CType(" ", Char))
                msg.Numero = Int32.Parse(veri(0))
                msg.Bytes = Int32.Parse(veri(1))
                msg.Traido = False
                retval.Add(msg)
            End If
        End While
    End Function
#End Region
#Region "TraerMensaje"
    Public Function Retrieve(ByVal rhs As MensajePop3) As MensajePop3
        Dim message, response As String
        Dim msg As New MensajePop3
        msg.Bytes = rhs.Bytes
        msg.Numero = rhs.Numero
        message = "RETR " + rhs.Numero.ToString + ControlChars.CrLf
        write(message)
        response = Respond()
        If response.Substring(0, 3) <> "+OK" Then
            Throw New System.Exception(response)
        End If
        msg.Traido = True
        While (1 = 1)
            response = Respond()
            If response = "." + vbCrLf Then
                Exit While
            Else
                msg.Mensaje += response
            End If
        End While
        Return msg
    End Function
#End Region
#Region "Private Functions"
    Private Function Respond() As String
        Dim enc As New System.Text.ASCIIEncoding
        Dim serverbuff(1024), buff(0) As Byte
        Dim count, bytes As Integer
        Dim stream As NetworkStream
        stream = GetStream()
        While 1 = 1
            bytes = stream.Read(buff, 0, 1)
            If bytes = 1 Then
                serverbuff(count) = buff(0)
                count = count + 1
                If buff(0) = Asc(vbLf) Then
                    Exit While
                End If
            Else
                Exit While
            End If
        End While
        Return enc.GetString(serverbuff, 0, count)
    End Function
    Private Sub write(ByVal message As String)
        Dim en As New System.Text.ASCIIEncoding
        Dim writebuffer(1024) As Byte
        writebuffer = en.GetBytes(message)
        Dim stream As NetworkStream = GetStream()
        stream.Write(writebuffer, 0, writebuffer.Length)
    End Sub
#End Region
End Class