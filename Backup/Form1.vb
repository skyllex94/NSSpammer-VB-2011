Imports System.Net.Mail
Imports System.Net
Public Class Form1
#Region "InitSpammer"
    Sub New()
        Control.CheckForIllegalCrossThreadCalls = False
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Dim threadCli As Threading.Thread = New Threading.Thread(AddressOf SendMail)
    ' Dim threadClick As Threading.Thread = New Threading.Thread(AddressOf SendMail)
    Private Sub StartChain(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GlassButton1.Click
        If TextBox1.Text = "" Then MsgBox("Input Username") : Exit Sub
        If TextBox2.Text = "" Then MsgBox("Input Password") : Exit Sub
        If TextBox3.Text = "" Then MsgBox("Forgot the Message") : Exit Sub
        If TextBox5.Text = "" Then MsgBox("Input Email of the VictiM") : Exit Sub
        If txtsubj.Text = "" Then MsgBox("Forgot the Subject") : Exit Sub
        If txtsmtp.Text = "" Then Label4.Text = "Input SMTP" : Exit Sub
        Try : Dim testmessage As New MailMessage
            testmessage.Subject = "NSSpammer"
            testmessage.To.Add(Me.TextBox5.Text)
            testmessage.From = New MailAddress(TextBox1.Text)
            testmessage.Body = "Your mail works. You can Spam now"
            Dim client As New SmtpClient(txtsmtp.Text, portselect.Text)
            client.EnableSsl = True
            client.Credentials = New NetworkCredential(TextBox1.Text, TextBox2.Text)
            client.Send(testmessage)
        Catch exception1 As Exception
            Dim exception As Exception = exception1
            Interaction.MsgBox("Invalid Mail or Options, check them agian!", MsgBoxStyle.OkOnly, Nothing) : Exit Sub
        End Try
        Try : threadCli.Resume() : Catch asd As Exception : End Try
        Try : threadCli.Start() : Catch ex As Exception : End Try
        GlassButton2.Enabled = True : GlassButton4.Enabled = False : GlassButton1.Enabled = False
        GroupBox2.Enabled = False : GroupBox4.Enabled = False : TextBox3.Enabled = False
        txtsubj.Enabled = False : CheckBox10.Enabled = False : engroupbox.Enabled = False
        enGroupBox.Enabled = False : GroupBox5.Enabled = False : GroupBox3.Enabled = False
    End Sub
#End Region

#Region "SendMails"
    Dim lb4 As Integer = 0
    Private Sub SendMail()
        Dim mail As New MailMessage
        Dim a As New SmtpClient(txtsmtp.Text, portselect.Text)
        Try : If Not TextBox4.Text = Nothing Then mail.Attachments.Add(New Attachment(TextBox4.Text))
            If Not TextBox6.Text = Nothing Then mail.Attachments.Add(New Attachment(TextBox6.Text))
            If Not TextBox7.Text = Nothing Then mail.Attachments.Add(New Attachment(TextBox7.Text))
            mail.Subject = txtsubj.Text
            mail.To.Add(TextBox5.Text)
            mail.From = New MailAddress(TextBox1.Text)
            mail.Body = (TextBox3.Text)
            a.EnableSsl = True
            a.Credentials = New NetworkCredential(TextBox1.Text, TextBox2.Text)
            TrackBar1.Enabled = False : Label3.Enabled = False
            Label12.Enabled = False : Label13.Enabled = False
            Label4.Text = "Starting NSSpammer"
            If TrackBar1.Value = 3 Then
                Do Until lb4.ToString = ComboBox2.Text
                    a.Send(mail) : lb4 += 1
                    Label4.Text = lb4 & " Messages Sent"
                    Threading.Thread.Sleep(3000)
                    ' If me.Button2_Click then ....
                Loop : End If
            If TrackBar1.Value = 2 Then
                Do Until lb4.ToString = ComboBox2.Text
                    a.Send(mail) : lb4 += 1
                    Label4.Text = lb4 & " Messages Sent"
                    Threading.Thread.Sleep(1000)
                Loop : End If
            If TrackBar1.Value = 1 Then
                Do Until lb4.ToString = ComboBox2.Text
                    a.Send(mail) : lb4 += 1
                    Label4.Text = lb4 & " Messages Sent"
                Loop : End If
        Catch 'exz As Exception
            'MsgBox("Error: " & exz.Message, MsgBoxStyle.Critical, "Error: ")
            GlassButton1.Enabled = False : GlassButton2.Enabled = False
            GroupBox2.Enabled = False : GroupBox4.Enabled = False : TextBox3.Enabled = False
            enGroupBox.Enabled = False : GroupBox5.Enabled = False : GroupBox3.Enabled = False
            CheckBox10.Enabled = False : GlassButton4.Enabled = False : Exit Sub
        End Try
        If lb4.ToString = ComboBox2.Text Then
            GlassButton2.Enabled = False
            MsgBox("Finished Spamming at " & lb4 & " Messages") : End If
    End Sub
    '   Private Sub Sendsecondmail()
    '       Do Until lb4.ToString = ComboBox2.Text
    '          a2.Send(mail2)
    'lb4 += 1
    'Label4.Text = lb4 & " Messages Sent"
    '     Loop
    '    Catch exz As Exception
    '       MsgBox("Error: " & exz.Message, MsgBoxStyle.Critical, "Error: ")
    'Button1.Enabled = True : Button5.Enabled = False
    'GroupBox2.Enabled = True : GroupBox4.Enabled = True : TextBox3.Enabled = True
    'enGroupBox.Enabled = True : GroupBox5.Enabled = True : GroupBox3.Enabled = True
    'CheckBox10.Enabled = True : Button2.Enabled = True : Exit Sub
    '  End Try
    'End Sub
#End Region

#Region "TestMail"
    Private Sub GlassButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GlassButton4.Click
        If TextBox1.Text = "" Then MsgBox("Input Username") : Exit Sub
        If TextBox2.Text = "" Then MsgBox("Input Password") : Exit Sub
        If txtsmtp.Text = "" Then Label4.Text = "Input SMTP" : Exit Sub
        Try : Dim testmessage As New MailMessage
            testmessage.Subject = "NSSpammer"
            testmessage.To.Add(Me.TextBox1.Text)
            testmessage.From = New MailAddress(TextBox1.Text)
            testmessage.Body = "Your mail works. You can Spam now"
            Dim client As New SmtpClient(txtsmtp.Text, portselect.Text)
            client.EnableSsl = True
            client.Credentials = New NetworkCredential(TextBox1.Text, TextBox2.Text)
            client.Send(testmessage)
            Interaction.MsgBox("Email is valid, u can Spam now!", MsgBoxStyle.OkOnly, Nothing)
        Catch exception1 As Exception
            Interaction.MsgBox("Invalid Mail or Options, check them agian!", MsgBoxStyle.OkOnly, Nothing)
        End Try
    End Sub
#End Region

#Region "Choose attachments/exit/options"
    Private Sub engroupbox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles engroupbox.CheckedChanged
        If engroupbox.Checked Then : GroupBox3.Enabled = True
        ElseIf engroupbox.Checked = False Then : GroupBox3.Enabled = False : End If
    End Sub
    Private Sub CheckBox10_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked Then : GroupBox5.Enabled = True
        ElseIf CheckBox10.Checked = False Then : GroupBox5.Enabled = False : End If
    End Sub
    Private Sub GlassButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GlassButton5.Click
        Dim file(4) As String
        file = Nothing
        Try : OpenFileDialog1.ShowDialog()
            file = OpenFileDialog1.FileNames
            TextBox4.Text = file(0)
        Catch asd As IndexOutOfRangeException : End Try
        Try : TextBox6.Text = file(1)
        Catch ex As IndexOutOfRangeException
        End Try
        Try : TextBox7.Text = file(2)
        Catch ex As IndexOutOfRangeException
        End Try
    End Sub
#End Region

#Region "Notify Icon"
    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub
    Private Sub GlassButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GlassButton3.Click
        WindowState = FormWindowState.Minimized
        Me.Hide()
        NotifyIcon1.Visible = True
    End Sub
#End Region

#Region "Stop Spammer"
    Dim asd As Boolean = True
    Private Sub GlassButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GlassButton2.Click
        threadCli.Suspend()
        GlassButton1.Enabled = True : GlassButton4.Enabled = True : GlassButton2.Enabled = False
        GroupBox2.Enabled = True : GroupBox4.Enabled = True : TextBox3.Enabled = True
        GroupBox5.Enabled = False : GroupBox3.Enabled = False : TextBox5.ReadOnly = True
        TextBox1.ReadOnly = True : TextBox2.ReadOnly = True : TextBox3.ReadOnly = True
        txtsubj.Enabled = True : txtsubj.ReadOnly = True : Label4.Height = 199
        Label4.Text = "Stoped at:" & lb4.ToString & " Messages"
    End Sub
#End Region

#Region "PassChar"
    Private Sub CheckBox12_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked Then : TextBox2.PasswordChar = ""
        ElseIf CheckBox12.Checked = False Then : TextBox2.PasswordChar = "*" : End If
    End Sub
#End Region

#Region "Form1_FormClosing"
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try : Try : threadCli.Resume() : Catch : End Try
            Try : threadCli.Abort() : Catch : End Try : Catch : End Try
    End Sub
#End Region

#Region "Trackbar"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = Me.Text & (My.Computer.Name)
        TrackBar1.SetRange(1, 3)
        TrackBar1.Value = 2
        For FadeIn = 0.0 To 1.0 Step 0.1
            Me.Opacity = FadeIn : Me.Refresh()
            Threading.Thread.Sleep(50) : Next
    End Sub

    Private Sub TrackBar1_Scroll_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        If TrackBar1.Value = 3 Then Label4.Text = "Slow speed of sending"
        If TrackBar1.Value = 2 Then Label4.Text = "Medium speed of sending"
        If TrackBar1.Value = 1 Then Label4.Text = "Fast speed of sending "
    End Sub
#End Region

#Region "Mail_Select"
    Private Sub gmailr_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gmailr.CheckedChanged
        txtsmtp.ReadOnly = True : portselect.ReadOnly = True
        txtsmtp.Text = "smtp.gmail.com" : portselect.Text = 587
    End Sub
    Private Sub aolr_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aolr.CheckedChanged
        txtsmtp.ReadOnly = True : portselect.ReadOnly = True
        txtsmtp.Text = "smtp.uk.aol.com" : portselect.Text = 587
    End Sub
    Private Sub hotmailr_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hotmailr.CheckedChanged
        txtsmtp.ReadOnly = True : portselect.ReadOnly = True
        txtsmtp.Text = "smtp.live.com" : portselect.Text = 25
    End Sub
    Private Sub yahoor_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yahoor.CheckedChanged
        txtsmtp.ReadOnly = True : portselect.ReadOnly = True
        txtsmtp.Text = "plus.smtp.mail.yahoo.com" : portselect.Text = 465
    End Sub
    Private Sub otherr_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles otherr.CheckedChanged
        engroupbox.Checked = True : GroupBox3.Enabled = True
        txtsmtp.Text = "Write SMTP... " : portselect.Text = "Port here..."
        txtsmtp.ReadOnly = False : portselect.ReadOnly = False
    End Sub
#End Region
End Class
