Imports System.ComponentModel
Imports OfficeOpenXml

Public Class frmCaptainsListPreview
    Dim captainsdata As DataTable
    Dim mergedpdffullpath As String
    Dim ftpurl As String = "ftp://ftp.manresa-sj.com"
    Dim ftpusername As String = "renzlo@manresa-sj.com"
    Dim ftppassword As String = "nica@5685647"
    Dim GenCaptTable As DataTable

    Private Sub frmCaptainsListPreview_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.butPrint.Enabled = False
        Me.butUploadToManresa.Enabled = False
        Me.butSave.Enabled = False
        Me.AxAcroPDF1.Visible = False
        If IO.Directory.Exists(IO.Path.GetTempPath & "CaptainsList") Then
            Try
                IO.Directory.Delete(IO.Path.GetTempPath & "CaptainsList", True)
            Catch ex As Exception
            End Try
        End If
        captainsdata = New DataTable
        captainsdata = DB_EXECUTE_SELECT("select CONVERT(VARCHAR(10),b.StartDateTime,110) as 'Start Date',b.Name as 'Event Name',c.lName as 'Last Name',c.fName as 'First Name',c.mName as 'Middle Name',c.suffix as 'Suffix' from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and b.FKCategoryId = 7 and b.StartDateTime > GETDATE() and a.Status = 2 order by b.StartDateTime,c.lName")
        Dim ReviewForm As New frmReview(captainsdata)
        ReviewForm.ShowDialog()
        If NewCaptainsData Is Nothing Then
            GenCaptTable = captainsdata
        Else
            GenCaptTable = NewCaptainsData
        End If
        Me.bgWorker.RunWorkerAsync()
    End Sub

    Private Sub GenerateCaptainsReport(_outfile As String, ddtable As DataTable, Optional ByVal IsHeader As Boolean = False)
        Dim retrycounter As Integer = 0
GenPDF:
        Try
            Dim templateFile As IO.FileInfo = Nothing
            If IsHeader Then
                templateFile = New IO.FileInfo("templates\CaptHeaderTemplate.xlsx")
            Else
                templateFile = New IO.FileInfo("templates\CaptChildTemplate.xlsx")
            End If
            Dim newFile As New IO.FileInfo(_outfile)
            If newFile.Exists Then
                newFile.Delete()
                newFile = New IO.FileInfo(_outfile)
            End If
            Using package As New ExcelPackage(newFile, templateFile)
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(1)
                If IsHeader Then
                    worksheet.Cells("B6").Value = ddtable.Rows(0)(0).ToString
                    worksheet.Cells("D6").Value = ddtable.Rows.Count
                    worksheet.Cells("A8").Value = ddtable.Rows(0)(1).ToString
                    worksheet.Cells("A9").Value = ddtable.Rows(0)(0).ToString
                    worksheet.Cells("B9").Value = String.Format("Total: {0}", ddtable.Rows.Count)
                    ddtable.Columns.RemoveAt(0)
                    ddtable.Columns.RemoveAt(0)
                    worksheet.Cells("A10").LoadFromDataTable(ddtable, False)
                Else
                    worksheet.Cells("B3").Value = ddtable.Rows(0)(0).ToString
                    worksheet.Cells("D3").Value = ddtable.Rows.Count
                    worksheet.Cells("A5").Value = ddtable.Rows(0)(1).ToString
                    worksheet.Cells("A6").Value = ddtable.Rows(0)(0).ToString
                    worksheet.Cells("B6").Value = String.Format("Total: {0}", ddtable.Rows.Count)
                    ddtable.Columns.RemoveAt(0)
                    ddtable.Columns.RemoveAt(0)
                    worksheet.Cells("A7").LoadFromDataTable(ddtable, False)
                End If
                'worksheet.Cells.AutoFitColumns()
                package.Save()
            End Using
            PDFConvert(_outfile)
        Catch ex As Exception
            retrycounter += 1
            If retrycounter <= 3 Then
                GoTo GenPDF
            End If
        End Try
    End Sub

    Private Sub bgWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bgWorker.DoWork
        Dim outfolder As New IO.DirectoryInfo(IO.Path.GetTempPath & "CaptainsList")
        If outfolder.Exists Then outfolder.Delete(True)
        If outfolder.Exists = False Then outfolder.Create()
        Dim futuredatelist As New List(Of String)
        For Each frow As DataRow In GenCaptTable.Rows
            If futuredatelist.Contains(frow(0).ToString) = False Then futuredatelist.Add(frow(0).ToString)
        Next
        Dim futcounter As Integer = 0
        For Each futstring As String In futuredatelist
            futcounter += 1
            Dim ftable As DataTable = GenCaptTable.Select(String.Format("[Start Date] = '{0}'", futstring)).CopyToDataTable
            GenerateCaptainsReport(String.Concat(outfolder.FullName, "\", futcounter.ToString.PadLeft(8, "0"), ".xlsx"), ftable, IIf(futcounter = 1, True, False))
        Next
        'merge now
        Dim pdflist As New List(Of String)
        Dim payls As IO.FileInfo() = outfolder.GetFiles("*.pdf")
        For Each payl As IO.FileInfo In payls
            pdflist.Add(payl.FullName)
        Next
        MergePdfFiles(pdflist, outfolder.FullName & "\CaptainsList.pdf")
        mergedpdffullpath = outfolder.FullName & "\CaptainsList.pdf"
    End Sub

    Private Sub bgWorker_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgWorker.ProgressChanged

    End Sub

    Private Sub bgWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgWorker.RunWorkerCompleted
        Me.AxAcroPDF1.LoadFile(mergedpdffullpath)
        Me.butPrint.Enabled = True
        Me.AxAcroPDF1.Visible = True
        Me.butSave.Enabled = True
        Me.butUploadToManresa.Enabled = True
    End Sub
    Private Sub butPrint_Click(sender As Object, e As EventArgs) Handles butPrint.Click
        Me.AxAcroPDF1.printWithDialog()
    End Sub

    Private Sub butUploadToManresa_Click(sender As Object, e As EventArgs) Handles butUploadToManresa.Click
        If MessageBox.Show("Are you sure you want to upload this captainslist on https://manresa-sj.org?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then Return
        Me.butUploadToManresa.Text = "Uploading, please wait..."
        Me.butUploadToManresa.Enabled = False
        Application.DoEvents()
        UploadFileToFTP(mergedpdffullpath)
    End Sub

    Private Sub butSave_Click(sender As Object, e As EventArgs) Handles butSave.Click
        Dim saveas As New SaveFileDialog
        saveas.Filter = "PDF File|*.pdf"
        saveas.FileName = "CaptainsList"
        saveas.OverwritePrompt = True
        If saveas.ShowDialog = DialogResult.OK Then
            If IO.File.Exists(mergedpdffullpath) Then
                My.Computer.FileSystem.CopyFile(mergedpdffullpath, saveas.FileName, True)
                MessageBox.Show(String.Format("Successfully saved in {0}.", saveas.FileName), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("No generated file to save!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub UploadFileToFTP(source As String)
        If IO.File.Exists(String.Concat(IO.Path.GetTempPath, "CaptainsList.pdf")) Then IO.File.Delete(String.Concat(IO.Path.GetTempPath, "CaptainsList.pdf"))
        'SecurePDF(source, String.Concat(IO.Path.GetTempPath, "CaptainsList.pdf"), "captain")
        'source = String.Concat(IO.Path.GetTempPath, "CaptainsList.pdf")
        Try
            Dim filename As String = IO.Path.GetFileName(source)
            Dim ftpfullpath As String = ftpurl & "/manresa-sj.org/wp-content/uploads/" & IO.Path.GetFileName(source)
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

            Me.butUploadToManresa.Enabled = True
            Me.butUploadToManresa.Text = "Upload to Manresa Website"
            MessageBox.Show("Successfully uploaded!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class