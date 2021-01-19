#Region " Email "

Public Class clsEmail

    Private WithEvents bgwEmail As New System.ComponentModel.BackgroundWorker
    Private Smtp As New SmtpClient
    Private sUserEmail As String
    Private SentException As Exception
    Public Event EmailSent(ErrMsg As Exception)

    Sub New()

    End Sub
    Sub New(UserEmail As String, UserPassword As String, Optional port As Integer = 587)
        sUserEmail = UserEmail
        CreateSmtp(UserPassword, port)
    End Sub

    Public Sub ChangeUser(UserEmail As String, UserPassword As String, Optional port As Integer = 587)
        sUserEmail = UserEmail
        CreateSmtp(UserPassword, port)
    End Sub

    Private Sub CreateSmtp(password As String, port As Integer)
        Smtp.Host = "smtp." & Split(sUserEmail, "@")(1)
        Smtp.Port = port '587 or 465
        Smtp.EnableSsl = True
        Smtp.DeliveryMethod = SmtpDeliveryMethod.Network
        Smtp.Timeout = 7000
        Smtp.Credentials = New Net.NetworkCredential(sUserEmail, password)
    End Sub

    Public Function CreateMail(email As String, header As String, message As String) As MailMessage
        Dim mail As New MailMessage
        mail.From = New MailAddress(sUserEmail)
        mail.To.Add(email)
        mail.IsBodyHtml = False
        mail.Body = header
        mail.Subject = message
        Return mail
    End Function

    Public Sub Send(Mail As MailMessage)
        bgwEmail.RunWorkerAsync(Mail)
    End Sub
    Private Sub bgwEmail_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles bgwEmail.DoWork
        Dim Mail = CType(e.Argument, MailMessage)
        Try
            Smtp.Send(Mail)
            SentException = Nothing
        Catch Ex As Exception 'operation timeout error no.
            SentException = Ex
        End Try
    End Sub

    Private Sub bgwEmail_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgwEmail.RunWorkerCompleted
        RaiseEvent EmailSent(SentException)
    End Sub

End Class

#End Region