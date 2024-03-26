Imports iTextSharp.text.pdf
Imports Microsoft.Office.Interop.Excel

Module GlobalAny
    Public FieldList As List(Of String)
    Public RoomAssignmentFieldList As List(Of String)
    Public FatherFieldList As List(Of String)
    Public KitchenFieldList As List(Of String)
    Public DonationFieldList As List(Of String)
    Public TransferFieldList As List(Of String)
    Public FieldsData As System.Data.DataTable
    Dim WithEvents writer As PdfWriter
    Public NewCaptainsData As System.Data.DataTable
    Public captainprev As frmCaptainsListPreview

    Public Sub LoadDefaultFieldList()
        RoomAssignmentFieldList = New List(Of String)
        RoomAssignmentFieldList.Add("c.display as 'Participant Name'")
        RoomAssignmentFieldList.Add("c.custom152 as 'Retreats#'")
        RoomAssignmentFieldList.Add("c.nickname as 'Called'")
        RoomAssignmentFieldList.Add("CONVERT(VARCHAR(10), c.custom50, 101) as 'BirthDate'")
        RoomAssignmentFieldList.Add("(case when c.custom19 = 'No' then '' else c.custom19 end) as 'Capt'")
        RoomAssignmentFieldList.Add("c.custom18 as 'Floor'")
        RoomAssignmentFieldList.Add("c.custom2 as 'Room Pref/Notes'")
        'RoomAssignmentFieldList.Add("c.custom3 as 'Notes'")
        RoomAssignmentFieldList.Add("'' as 'RoomAssigned'")
        FatherFieldList = New List(Of String)
        FatherFieldList.Add("c.display as 'Participant Name'")
        FatherFieldList.Add("c.nickname as 'Called'")
        FatherFieldList.Add("(case when c.phoneName = 'Home Phone' OR c.phoneName = 'Business Phone' then c.phone else '' End) as 'Home Phone'")
        FatherFieldList.Add("(case when c.phoneName = 'Mobile Phone' then c.phone else '' end) as 'Cell Phone'")
        FatherFieldList.Add("c.custom152 as 'Retreats#'")
        'FatherFieldList.Add("'' as 'Spouse'")
        'FatherFieldList.Add("c.profession as 'Occupation'")
        KitchenFieldList = New List(Of String)
        KitchenFieldList.Add("c.lName as 'LastName'")
        KitchenFieldList.Add("c.fName as 'FirstName'")
        KitchenFieldList.Add("c.mName as 'MiddleName'")
        KitchenFieldList.Add("(case when c.phoneName = 'Home Phone' OR c.phoneName = 'Mobile Phone' then c.phone else '' End) as 'Home Phone'")
        KitchenFieldList.Add("(case when c.phoneName = 'Business Phone' then c.phone else '' end) as 'Work Phone'")
        KitchenFieldList.Add("c.custom4 as 'Dietary'")
        DonationFieldList = New List(Of String)
        DonationFieldList.Add("c.display as 'Participant Name'")
        DonationFieldList.Add("c.nickname as 'Called'")
        DonationFieldList.Add("(select max('$'+ CONVERT(varchar(20),k.AmountPaid, 1)) from events_event_reservations k where k.FKDonorId = c.id and k.FKEventId = a.FKEventId) as 'Deposit'")
        'DonationFieldList.Add("'$'+ CONVERT(varchar(20),k.AmountPaid, 1) as 'Deposit'")
        'DonationFieldList.Add("'$0.00' as 'Deposit'")
        DonationFieldList.Add("'' as 'Donation'")
        DonationFieldList.Add("'' as 'Book'")
        DonationFieldList.Add("'' as 'Candy'")
        DonationFieldList.Add("'' as 'Photo'")
        DonationFieldList.Add("'' as 'Total'")
        TransferFieldList = New List(Of String)
        TransferFieldList.Add("c.display as 'Participant Name'")
        TransferFieldList.Add("c.custom152 as 'Retreats#'")
        TransferFieldList.Add("c.custom152 + 1000000 as 'RetreatsSum'")
    End Sub

    Public Sub PDFConvert(_xlsfile As String)
        Dim fileName As String = _xlsfile
        Dim xlsApp = New Microsoft.Office.Interop.Excel.Application
        xlsApp.ScreenUpdating = False
        Dim xlsBook As Microsoft.Office.Interop.Excel.Workbook
        Dim paramExportFormat As XlFixedFormatType = XlFixedFormatType.xlTypePDF
        Dim paramExportQuality As XlFixedFormatQuality = XlFixedFormatQuality.xlQualityStandard
        Dim paramOpenAfterPublish As Boolean = False
        Dim paramIncludeDocProps As Boolean = True
        Dim paramIgnorePrintAreas As Boolean = True
        Dim paramFromPage As Object = Type.Missing
        Dim paramToPage As Object = Type.Missing
        xlsBook = xlsApp.Workbooks.Open(fileName, UpdateLinks:=False, ReadOnly:=False)
        xlsBook.ExportAsFixedFormat(paramExportFormat, IO.Path.ChangeExtension(fileName, ".pdf"), paramExportQuality, paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage, paramToPage, paramOpenAfterPublish)
        xlsBook.Close(SaveChanges:=False)
        xlsApp.Quit()
    End Sub

    Public Function MergePdfFiles(ByVal pdfFiles As List(Of String), ByVal outputPath As String) As Boolean
        Dim result As Boolean = False
        Dim pdfCount As Integer = 0     'total input pdf file count
        Dim f As Integer = 0    'pointer to current input pdf file
        Dim fileName As String
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pageCount As Integer = 0
        Dim pdfDoc As iTextSharp.text.Document = Nothing    'the output pdf document
        writer = Nothing
        Dim cb As PdfContentByte = Nothing

        Dim page As PdfImportedPage = Nothing
        Dim rotation As Integer = 0

        Try
            pdfCount = pdfFiles.Count
            If pdfCount > 1 Then
                'Open the 1st item in the array PDFFiles
                fileName = pdfFiles(f)
                reader = New iTextSharp.text.pdf.PdfReader(fileName)
                'Get page count
                pageCount = reader.NumberOfPages

                pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1), 18, 18, 18, 18)

                writer = PdfWriter.GetInstance(pdfDoc, New IO.FileStream(outputPath, IO.FileMode.OpenOrCreate))
                With pdfDoc
                    .Open()
                End With
                'Instantiate a PdfContentByte object
                cb = writer.DirectContent
                'Now loop thru the input pdfs
                While f < pdfCount
                    'Declare a page counter variable
                    Dim i As Integer = 0
                    'Loop thru the current input pdf's pages starting at page 1
                    While i < pageCount
                        i += 1
                        'Get the input page size
                        pdfDoc.SetPageSize(reader.GetPageSizeWithRotation(i))
                        'Create a new page on the output document
                        pdfDoc.NewPage()
                        'If it is the 1st page, we add bookmarks to the page
                        'Now we get the imported page
                        page = writer.GetImportedPage(reader, i)
                        'Read the imported page's rotation
                        rotation = reader.GetPageRotation(i)
                        'Then add the imported page to the PdfContentByte object as a template based on the page's rotation
                        If rotation = 90 Then
                            cb.AddTemplate(page, 0, -1.0F, 1.0F, 0, 0, reader.GetPageSizeWithRotation(i).Height)
                        ElseIf rotation = 270 Then
                            cb.AddTemplate(page, 0, 1.0F, -1.0F, 0, reader.GetPageSizeWithRotation(i).Width + 60, -30)
                        Else
                            cb.AddTemplate(page, 1.0F, 0, 0, 1.0F, 0, 0)
                        End If
                    End While
                    'Increment f and read the next input pdf file
                    f += 1
                    If f < pdfCount Then
                        fileName = pdfFiles(f)
                        reader = New iTextSharp.text.pdf.PdfReader(fileName)
                        pageCount = reader.NumberOfPages
                    End If
                End While
                'When all done, we close the document so that the pdfwriter object can write it to the output file
                pdfDoc.Close()
                result = True
            End If
        Catch ex As Exception
            Return False
        End Try
        Return result
    End Function

    Public Sub SecurePDF(_inPath As String, _outPath As String, _pw As String)
        Dim InputFileStream As System.IO.FileStream = New System.IO.FileStream(_inPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
        Dim OutputFileStream As System.IO.FileStream = New System.IO.FileStream(_outPath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
        OutputFileStream.SetLength(0)
        Dim PdfReader As PdfReader = New PdfReader(InputFileStream)
        PdfEncryptor.Encrypt(PdfReader, OutputFileStream, True, _pw, _pw, PdfWriter.ALLOW_SCREENREADERS)
        InputFileStream.Close()
        InputFileStream.Dispose()
    End Sub
End Module
