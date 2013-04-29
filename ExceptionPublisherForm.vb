Imports System.net
Imports System.Net.Mail
Imports System.Threading
Friend Class ExceptionPublisherForm
    Private _sendOK As Boolean = False
    Private _finished As AutoResetEvent
    Public Sub New(ByVal ex As Exception)
        If ex Is Nothing Then Throw New ArgumentNullException("ex")

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ExMsg.Text = ex.Message
        txtDetails.Text = ex.ToString
        SplitContainer1.Panel2Collapsed = True
        Me.Height = Me.Height - 180
    End Sub
    Private Sub SendAsMail()
        If Not My.Computer.Network.IsAvailable Then Exit Sub
        SendMail()
    End Sub
    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click
        SendAsMail()
        Me.Close()
    End Sub

    Private Sub btnDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetails.Click
        If SplitContainer1.Panel2Collapsed Then
            SplitContainer1.Panel2Collapsed = False
            btnDetails.Text = "Hide Details"
            Me.Height = Me.Height + 180
        Else
            btnDetails.Text = "Show Details"
            SplitContainer1.Panel2Collapsed = True
            Me.Height = Me.Height - 180
        End If
    End Sub

    Private Sub ctxmSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctxmSelectAll.Click
        txtDetails.SelectAll()
    End Sub

    Private Sub ctxmCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctxmCopy.Click
        System.Windows.Forms.Clipboard.SetDataObject(txtDetails.SelectedText)
    End Sub

    Private Function ExternalIP() As String
        Try
            Dim request As HttpWebRequest = WebRequest.Create("http://checkip.dyndns.org/")
            Dim response As HttpWebResponse = request.GetResponse()
            Dim reader As IO.StreamReader = New IO.StreamReader(response.GetResponseStream())
            Dim sb As New Text.StringBuilder
            Dim str As String = reader.ReadLine()
            Do While str IsNot Nothing
                sb.Append(str)
                str = reader.ReadLine()
            Loop
            str = sb.ToString
            str = str.Substring(InStr(str, ":")).Trim
            str = str.Substring(0, InStr(str, "<") - 1)
            Return str
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Sub SendMail()
        'Builed The MSG
        Dim xip As String = ExternalIP()
        Dim ipText As New Text.StringBuilder
        Dim name As String = My.Computer.Name
        For Each addr As IPAddress In System.Net.Dns.GetHostEntry(name).AddressList
            Select Case addr.AddressFamily
                Case Sockets.AddressFamily.InterNetwork
                    If ipText.Length > 0 Then ipText.Append(", ")
                    ipText.Append("IPV4: ")
                    ipText.Append(addr.ToString())
                Case Sockets.AddressFamily.InterNetworkV6
                    If ipText.Length > 0 Then ipText.Append(",")
                    ipText.Append("IPV6: ")
                    ipText.Append(addr.ToString())
                Case Else
            End Select
        Next
        Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
        msg.To.Add("ghull@amcomsoft.com")
        msg.From = New MailAddress("donotreply@amcomsoft.com", "Amcom Automated Exception Message", System.Text.Encoding.UTF8)
        msg.Subject = "Automated Exception Message from " & App.Name
        msg.SubjectEncoding = System.Text.Encoding.UTF8
        Dim sb As New Text.StringBuilder
        sb.Append(vbCrLf)
        sb.Append("********************************************************************************")
        sb.Append(vbCrLf)
        sb.Append("* Amcom Software, Inc.   Automated Exception Report Message    ")
        sb.Append(vbCrLf)
        sb.Append("*")
        sb.Append(vbCrLf)
        sb.Append("* " & Now.ToLongDateString & " " & Now.ToShortTimeString & " " & TimeZone.CurrentTimeZone.StandardName)
        sb.Append(vbCrLf)
        sb.Append("*")
        sb.Append(vbCrLf)
        sb.Append("* System: " & name)
        sb.Append(vbCrLf)
        sb.Append("* Site External IP: " & xip)
        sb.Append(vbCrLf)
        ipText.Insert(0, "* Internal IP Address(es): ")
        sb.Append(ipText.ToString)
        sb.Append(vbCrLf)
        sb.Append("*")
        sb.Append(vbCrLf)
        sb.Append("* General Failure Details: ")
        sb.Append(vbCrLf)
        sb.Append("* ")
        sb.Append(ExMsg.Text)
        sb.Append(vbCrLf)
        sb.Append("********************************************************************************")
        sb.Append(vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("Details:------------------------------------------------------------------------")
        sb.Append(vbCrLf)
        sb.Append(txtDetails.Text)
        msg.Body = sb.ToString
        msg.BodyEncoding = System.Text.Encoding.UTF8
        msg.IsBodyHtml = False
        msg.Priority = MailPriority.High
        msg.From = New MailAddress("sdcautomatedsupport@gmail.com")
        msg.ReplyTo = New MailAddress("sdcautomatedsupport@gmail.com")

        'Add the Creddentials
        Dim client As SmtpClient = New SmtpClient()
        client.Credentials = New System.Net.NetworkCredential("sdcautomatedsupport@gmail.com", "@thinksdc!")
        client.Port = 587 'or use 587            
        client.Host = "smtp.gmail.com"
        client.EnableSsl = True
        Dim userState As Object = msg
        _sendOK = False
        Dim t As New Threading.Thread(AddressOf SendMessage)
        t.IsBackground = True
        Dim o() As Object = {msg, client}
        t.Start(CObj(o))
        Me.Cursor = Windows.Forms.Cursors.WaitCursor
        _finished = New AutoResetEvent(False)
        _finished.WaitOne(Timeout.Infinite)
        Me.Cursor = Windows.Forms.Cursors.Default
        If Not _sendOK Then
            MsgBox("We were unable to send an automated report message to support. ", MsgBoxStyle.Exclamation, "Automated Exception Message")
        Else
            MsgBox("Message Sent Successfully.", MsgBoxStyle.Information, "Automated Exception Message")

        End If
    End Sub
    Private Sub SendMessage(ByVal o As Object)
        Dim t As Object() = o
        Dim s As SmtpClient = t(1)
        Dim text As System.Net.Mail.MailMessage = t(0)
        Try
            s.Send(text)
            _sendOK = True
        Catch ex As Exception
            _sendOK = False
        End Try
        _finished.Set()
    End Sub
    Private Sub ExMsg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ExMsg.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub ExMsg_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExMsg.TextChanged

    End Sub
    Private Sub cmdToClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToClipboard.Click
        Dim t As New System.Threading.Thread(AddressOf CopyThread)
        _sendOK = False
        _finished = New AutoResetEvent(False)
        t.SetApartmentState(Threading.ApartmentState.STA)
        Me.Cursor = Windows.Forms.Cursors.WaitCursor
        t.Start(CObj(Me.txtDetails.Text))
        _finished.WaitOne(Timeout.Infinite)
        If _sendOK Then
            System.Windows.Forms.MessageBox.Show(Me, "Details are on clipboard", "Application Exception", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Information)
        Else
            System.Windows.Forms.MessageBox.Show(Me, "Failed to put details on the clipboard", "Application Exception", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub CopyThread(ByVal o As Object)
        Try
            System.Windows.Forms.Clipboard.Clear()
            System.Windows.Forms.Clipboard.SetText(CStr(o))
            _sendOK = True
            Me.Cursor = Windows.Forms.Cursors.Default
        Catch ex As Exception

        End Try
        _finished.Set()
    End Sub
End Class