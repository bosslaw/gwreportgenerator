Public Class frmUploader
    Dim sourcefile As String
    Dim ftpurl As String = "ftp://ftp.manresa-sj.com"
    Dim ftpusername As String = "renzlo@manresa-sj.com"
    Dim ftppassword As String = "nica@5685647"

    Public Sub New(_srcfile As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        sourcefile = _srcfile
    End Sub
    Private Sub frmUploader_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Application.DoEvents()
        UploadFileToFTP(sourcefile)
    End Sub

    Private Sub UploadFileToFTP(source As String)
        Application.DoEvents()

        Try
            Dim filename As String = IO.Path.GetFileName(source)
            Dim ftpfullpath As String = ftpurl & "/files/" & IO.Path.GetFileName(source)
            Dim ftp As Net.FtpWebRequest = DirectCast(Net.FtpWebRequest.Create(New Uri(ftpfullpath)), Net.FtpWebRequest)
            ftp.Credentials = New Net.NetworkCredential(ftpusername, ftppassword)

            ftp.KeepAlive = True
            ftp.UseBinary = True
            ftp.Method = Net.WebRequestMethods.Ftp.UploadFile

            Dim fs As IO.FileStream = IO.File.OpenRead(source)
            Dim buffer As Byte() = New Byte(fs.Length - 1) {}
            fs.Read(buffer, 0, buffer.Length)
            fs.Close()

            Dim ftpstream As IO.Stream = ftp.GetRequestStream()
            ftpstream.Write(buffer, 0, buffer.Length)
            ftpstream.Close()

            MessageBox.Show("Successfully uploaded!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub frmUploader_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frmCaptainsListPreview.Activate() : GC.SuppressFinalize(Me) : Me.Dispose()
    End Sub

    Private Sub frmUploader_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GC.Collect()
    End Sub
End Class