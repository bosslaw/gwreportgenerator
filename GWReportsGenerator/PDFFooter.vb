Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class PDFFooter
    Inherits PdfPageEventHelper
    ' write on top of document
    Public Overrides Sub OnOpenDocument(writer As PdfWriter, document As Document)
        MyBase.OnOpenDocument(writer, document)
        'Dim tabFot As New PdfPTable(New Single() {1.0F})
        'tabFot.SpacingAfter = 10.0F
        'Dim cell As PdfPCell
        'tabFot.TotalWidth = 300.0F
        'cell = New PdfPCell(New Phrase("Header"))
        'tabFot.AddCell(cell)
        'tabFot.WriteSelectedRows(0, -1, 150, document.Top, writer.DirectContent)
    End Sub

    ' write on start of each page
    Public Overrides Sub OnStartPage(writer As PdfWriter, document As Document)
        MyBase.OnStartPage(writer, document)
    End Sub

    ' write on end of each page
    Public Overrides Sub OnEndPage(writer As PdfWriter, document As Document)
        MyBase.OnEndPage(writer, document)
        Dim tabFot As New PdfPTable(New Single() {1.0F})
        Dim cell As PdfPCell
        tabFot.TotalWidth = 300.0F
        cell = New PdfPCell(New Phrase("Footer"))
        tabFot.AddCell(cell)
        tabFot.WriteSelectedRows(0, -1, 150, document.Bottom, writer.DirectContent)
    End Sub

    'write on close of document
    Public Overrides Sub OnCloseDocument(writer As PdfWriter, document As Document)
        MyBase.OnCloseDocument(writer, document)
    End Sub
End Class