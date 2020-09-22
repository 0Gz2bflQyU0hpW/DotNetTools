Option Explicit On 
Option Strict On
Option Infer Off
Imports System.IO
Imports System.Net.Mail
Public Class Smtp
#Region "Vars"
    Private vFrom As String
    Private vToEmail As String
    Private vBody As String = ""
    Private vBodyHtml As String = ""
    Private vSubject As String
    Private vMimeText As String = ""
    Private vAttachFile As String = ""
    Private vBodyHtmlACrear As String = ""
    Private vSmtpServer As String = ""
    Private vSmtpPort As String = ""
    Private vSmtpUser As String = ""
    Private vSmtpPassword As String = ""
    Private vHideCopy As String = ""
    Private vSsl As Boolean
#End Region
#Region "Propiedades"
    Public Property HideCopy() As String
        Get
            HideCopy = vHideCopy
        End Get
        Set(ByVal Value As String)
            vHideCopy = Value
        End Set
    End Property
    Public Property From() As String
        Get
            From = vFrom
        End Get
        Set(ByVal Value As String)
            vFrom = Value
        End Set
    End Property
    Public Property ToEmail() As String
        Get
            ToEmail = vToEmail
        End Get
        Set(ByVal Value As String)
            vToEmail = Value
        End Set
    End Property
    Public Property Body() As String
        Get
            Body = vBody
        End Get
        Set(ByVal Value As String)
            vBody = Value
        End Set
    End Property
    Public Property BodyHtml() As String
        Get
            BodyHtml = vBodyHtml
        End Get
        Set(ByVal Value As String)
            vBodyHtml = Value
        End Set
    End Property
    Public Property Subject() As String
        Get
            Subject = vSubject
        End Get
        Set(ByVal Value As String)
            vSubject = Value
        End Set
    End Property
    Public Property MimeText() As String
        Get
            MimeText = vMimeText
        End Get
        Set(ByVal Value As String)
            vMimeText = Value
        End Set
    End Property
    Public Property AttachFile() As String
        Get
            AttachFile = vAttachFile
        End Get
        Set(ByVal Value As String)
            vAttachFile = Value
        End Set
    End Property
    Public Property BodyHtmlACrear() As String
        Get
            BodyHtmlACrear = vBodyHtmlACrear
        End Get
        Set(ByVal Value As String)
            vBodyHtmlACrear = Value
        End Set
    End Property
    Public Property SmtpServer() As String
        Get
            SmtpServer = vSmtpServer
        End Get
        Set(ByVal Value As String)
            vSmtpServer = Value
        End Set
    End Property
    Public Property SmtpPort() As String
        Get
            SmtpPort = vSmtpPort
        End Get
        Set(ByVal Value As String)
            vSmtpPort = Value
        End Set
    End Property
    Public Property SmtpUser() As String
        Get
            SmtpUser = vSmtpUser
        End Get
        Set(ByVal Value As String)
            vSmtpUser = Value
        End Set
    End Property
    Public Property SmtpPassword() As String
        Get
            SmtpPassword = vSmtpPassword
        End Get
        Set(ByVal Value As String)
            vSmtpPassword = Value
        End Set
    End Property
    Public Property Ssl() As Boolean
        Get
            Ssl = vSsl
        End Get
        Set(ByVal Value As Boolean)
            vSsl = Value
        End Set
    End Property
#End Region
#Region "Metodos"
    Public Function Send() As String
        Try
            Send = ""
            If vSmtpPassword.Trim.Length <> 0 And vSmtpUser.Trim.Length <> 0 And vSmtpPort.Trim.Length <> 0 And vSmtpServer.Trim.Length <> 0 Then
                Dim oEmail As New System.Net.Mail.MailMessage
                Dim smtp As Net.Mail.SmtpClient
                smtp = New Net.Mail.SmtpClient(SmtpServer)
                smtp.Credentials = New Net.NetworkCredential(SmtpUser, SmtpPassword)
                smtp.Port = Convert.ToInt32(vSmtpPort)
                With oEmail
                    If vBodyHtml.Trim.Length <> 0 Then
                        .IsBodyHtml = True
                        .Body = vBodyHtml
                    Else
                        .IsBodyHtml = False
                        .Body = vBody
                    End If
                    If AttachFile.Trim.Length <> 0 Then
                        If AttachFile.Trim.IndexOf(";") > 0 Then
                            Dim Argsf() As String
                            Argsf = AttachFile.Trim.Split(CChar(";"))
                            Dim x As Integer
                            For x = 0 To Argsf.Length - 1
                                If Argsf(x).Trim.Length <> 0 Then
                                    .Attachments.Add(New System.Net.Mail.Attachment(Argsf(x).Replace(";", "")))
                                End If
                            Next
                            Argsf = Nothing
                        Else
                            .Attachments.Add(New System.Net.Mail.Attachment(vAttachFile))
                        End If
                    End If
                    .From = New Net.Mail.MailAddress(vFrom)
                    Dim Args() As String
                    Args = vToEmail.Split(CChar(";"))
                    Dim n As Integer
                    For n = 0 To Args.Length - 1
                        If Args(n).Trim.Length > 0 Then
                            If Args(n).IndexOf("@") > 0 Then
                                .To.Add(Args(n))
                            End If
                        End If
                    Next
                    If Me.vHideCopy.Trim.Length <> 0 Then
                        Args = vHideCopy.Split(CChar(";"))
                        For n = 0 To Args.Length - 1
                            If Args(n).IndexOf("@") > 0 Then
                                .CC.Add(Args(n))
                            End If
                        Next
                    End If
                    Args = Nothing
                    .Subject = vSubject
                End With
                smtp.EnableSsl = vSsl
                smtp.Send(oEmail)
                Send = "OK"
            Else
                Dim mensaje As New CDO.Message
                With mensaje
                    .From = vFrom
                    .To = vToEmail
                    If Me.vHideCopy.Trim.Length <> 0 Then .BCC = vHideCopy
                    .Subject = vSubject
                    If vBody.Trim.Length <> 0 Then .TextBody = vBody
                    If vBodyHtml.Trim.Length <> 0 Then .HTMLBody = vBodyHtml
                    If vAttachFile.Trim.Length <> 0 Then .AddAttachment(vAttachFile)
                    If vBodyHtmlACrear.Trim.Length <> 0 Then .CreateMHTMLBody(vBodyHtmlACrear)
                    .Send()
                End With
                Send = "OK"
            End If
        Catch ex As Exception
            Send = ex.Message
        End Try
    End Function
#End Region
End Class
