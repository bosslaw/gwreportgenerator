Imports System.Deployment.Application
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class frmMain
    Dim blTitle As String = "GiftWorks Reports Generator v"
    Dim eventnamesdata As DataTable
    Dim radgvdata As DataTable
    Dim fathersdata As DataTable
    Dim kitchendata As DataTable
    Dim donationdata As DataTable
    Dim transferdata As DataTable


    'Public Sub New()

    '    ' This call is required by the designer.
    '    InitializeComponent()

    '    ' Add any initialization after the InitializeComponent() call.
    '    Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
    '    Me.SetStyle(ControlStyles.UserPaint, True)
    '    Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
    '    Me.SetStyle(ControlStyles.ResizeRedraw, True)
    'End Sub

    Private Sub txtEventName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEventName.TextChanged
        If Me.txtEventName.Text.Trim.Length = 0 Then
            Me.lbEventList.DataSource = eventnamesdata
            Me.lbEventList.DisplayMember = "Name"
            Me.lbEventList.ValueMember = "Id"
            Me.lbEventList.ClearSelected()
            LoadFilteredValues()
        Else
            LoadFilteredValues()
        End If
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            blTitle = blTitle & My.Application.Info.Version.ToString
        Catch ex As Exception
        End Try
        Me.Text = blTitle
        LoadDefaultFieldList()
        LoadEventNames()
        LoadCategories()
        LoadStatuses()
        'newly added
        Me.cmbCategories.Text = "Conference Retreat"
        Me.cmbCategories.Enabled = False
        'Me.cmBStatuses.Text = "Open"
        'Me.cmBStatuses.Enabled = False
        'LoadFieldsData()
        LoadFilteredValues()
        'Me.dgvRoomAssignments.ContextMenuStrip = cmsMenu
        'Me.dgvFathersList.ContextMenuStrip = cmsMenu
        'Me.dgvKitchenReport.ContextMenuStrip = cmsMenu
        'Me.dgvDonationList.ContextMenuStrip = cmsMenu
    End Sub

    Private Sub LoadEventNames()
        Me.lbEventList.SuspendLayout()
        eventnamesdata = New DataTable
        eventnamesdata = DB_EXECUTE_SELECT("select * from events_event_info where Id > 0 and (FKStatusID = 1 or FKStatusID = 2) order by Name")
        Me.lbEventList.DataSource = eventnamesdata
        Me.lbEventList.DisplayMember = "Name"
        Me.lbEventList.ValueMember = "Id"
        Me.lbEventList.ClearSelected()
        Me.lbEventList.ResumeLayout()
    End Sub

    Private Sub LoadCategories()
        Dim allcatdata As New DataTable
        allcatdata.Columns.Add("id", GetType(Integer))
        allcatdata.Columns.Add("name")
        allcatdata.Rows.Add("-1", "All Categories")
        Dim categoriesdata As DataTable = DB_EXECUTE_SELECT("select id,name from events_list_eventcategories where id > 0 order by id")
        allcatdata.Merge(categoriesdata)
        Me.cmbCategories.DataSource = allcatdata
        Me.cmbCategories.DisplayMember = "name"
        Me.cmbCategories.ValueMember = "id"
    End Sub

    Private Sub LoadStatuses()
        Dim allstatdata As New DataTable
        allstatdata.Columns.Add("id", GetType(Integer))
        allstatdata.Columns.Add("name")
        allstatdata.Rows.Add("-1", "All Statuses")
        allstatdata.Rows.Add("1", "Open")
        allstatdata.Rows.Add("2", "Closed")
        'Dim statusesdata As DataTable = DB_EXECUTE_SELECT("select id,name from events_list_eventstatuses where id > 0 order by id")
        'allstatdata.Merge(statusesdata)
        Me.cmBStatuses.DataSource = allstatdata
        Me.cmBStatuses.DisplayMember = "name"
        Me.cmBStatuses.ValueMember = "id"
    End Sub

    Private Sub LoadCategStatus()
        Me.cmbCategories.Items.Add("Conference Retreat")
        Me.cmbCategories.SelectedIndex = 0
        Me.cmbCategories.Enabled = False
        Me.cmBStatuses.Items.Add("Open")
        Me.cmBStatuses.SelectedIndex = 0
        Me.cmBStatuses.Enabled = False
    End Sub

    Private Sub LoadFieldsData()
        FieldsData = New DataTable
        FieldsData.Columns.Add("Real")
        FieldsData.Columns.Add("Fake")
        Dim basa As New IO.StreamReader("GWReportsGenerator.tbl", System.Text.Encoding.UTF8)
        While Not basa.EndOfStream
            Dim layn As String = basa.ReadLine.Trim
            If layn.Length > 0 Then
                Dim slayn As String() = layn.Split(vbTab)
                FieldsData.Rows.Add(slayn)
            End If
        End While
        basa.Close() : basa.Dispose()
    End Sub

    Private Sub LoadFilteredValues()
        Try
            Dim filteredeventdata As New DataTable
            Select Case Me.cmbCategories.SelectedValue
                Case -1
                    Select Case Me.cmBStatuses.SelectedValue
                        Case -1
                            filteredeventdata = eventnamesdata.Select(String.Format("Name like '%{0}%'", Me.txtEventName.Text)).CopyToDataTable
                        Case Else
                            filteredeventdata = eventnamesdata.Select(String.Format("Name like '%{0}%' and fkstatusid = {1}", Me.txtEventName.Text, Me.cmBStatuses.SelectedValue)).CopyToDataTable
                    End Select
                Case Else
                    Select Case Me.cmBStatuses.SelectedValue
                        Case -1
                            filteredeventdata = eventnamesdata.Select(String.Format("Name like '%{0}%' and fkcategoryid = {1}", Me.txtEventName.Text, Me.cmbCategories.SelectedValue)).CopyToDataTable
                        Case Else
                            filteredeventdata = eventnamesdata.Select(String.Format("Name like '%{0}%' and fkcategoryid = {1} and fkstatusid = {2}", Me.txtEventName.Text, Me.cmbCategories.SelectedValue, Me.cmBStatuses.SelectedValue)).CopyToDataTable
                    End Select
            End Select
            filteredeventdata = filteredeventdata.Select("", "StartDateTime DESC").CopyToDataTable
            Me.lbEventList.DataSource = filteredeventdata
            Me.lbEventList.DisplayMember = "Name"
            Me.lbEventList.ValueMember = "Id"
            Me.lbEventList.ClearSelected()
        Catch ex As Exception
            Me.lbEventList.DataSource = Nothing
        End Try
    End Sub

    Private Sub cmbCategories_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCategories.SelectionChangeCommitted
        LoadFilteredValues()
    End Sub

    Private Sub cmBStatuses_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmBStatuses.SelectionChangeCommitted
        LoadFilteredValues()
    End Sub

    Private Sub dgvRoomAssignments_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRoomAssignments.DataSourceChanged
        Try
            Me.labRATitle.Text = String.Format("Room Assignments For Participants Coming on: {0}", CDate(Me.lbEventList.SelectedItem("StartDateTime").ToString).ToString("MMMM dd, yyyy"))
            Me.labRASubTitle.Text = String.Format("{0} Participant(s)", Me.dgvRoomAssignments.Rows.Count)
        Catch ex As Exception
            Me.labRATitle.Text = "..."
            Me.labRASubTitle.Text = "..."
        End Try
    End Sub

    Private Sub lbEventList_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbEventList.SelectedValueChanged
        'If lbEventList.SelectedValue Is Nothing Then Return
        'If lbEventList.SelectedValue.ToString = "System.Data.DataRowView" Then Return
        Me.Cursor = Cursors.WaitCursor
        radgvdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", RoomAssignmentFieldList), Me.lbEventList.SelectedValue))
        Me.dgvRoomAssignments.DataSource = radgvdata
        fathersdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", FatherFieldList), Me.lbEventList.SelectedValue))
        Me.dgvFathersList.DataSource = fathersdata
        kitchendata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", KitchenFieldList), Me.lbEventList.SelectedValue))
        Me.dgvKitchenReport.DataSource = kitchendata
        donationdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", DonationFieldList), Me.lbEventList.SelectedValue))
        'donationdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a left join events_event_info b on a.FKEventId = b.Id left join donor_donors c on a.FKDonorId = c.id inner join events_event_reservations k on k.FKDonorId = c.id and k.FKEventId = a.FKEventId where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", DonationFieldList), Me.lbEventList.SelectedValue))
        Me.dgvDonationList.DataSource = donationdata
        transferdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", TransferFieldList), Me.lbEventList.SelectedValue))
        Try
            Dim dataView As New DataView(transferdata)
            dataView.Sort = "RetreatsSum DESC,Participant Name"
            Dim newtransdata As DataTable = dataView.ToTable
            newtransdata.Columns.RemoveAt(2)
            Me.dgvTransferList.DataSource = newtransdata
        Catch ex As Exception
        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub dgvFathersList_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvFathersList.DataSourceChanged
        Try
            Me.labFatherDate.Text = CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToLongDateString
            Me.labFatherRetreat.Text = Me.lbEventList.SelectedItem(1).ToString
            Me.labFatherNumParticipant.Text = Me.dgvFathersList.Rows.Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub dgvKitchenReport_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvKitchenReport.DataSourceChanged
        Try
            Me.labKitDate.Text = CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToLongDateString
            Me.labKitRetreat.Text = Me.lbEventList.SelectedItem(1).ToString
            Me.labKitNumParticipants.Text = Me.dgvKitchenReport.Rows.Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub dgvDonationList_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDonationList.DataSourceChanged
        Try
            Dim fundname As String = DB_EXECUTE_SELECT(String.Format("select value from donor_funds where id={0}", Me.lbEventList.SelectedItem("fkfundid"))).Rows(0)(0)
            Me.labDonationFundLabel.Text = fundname
            Me.labDonationTitle.Text = String.Format("Donation List For Participants Coming on: {0}", CDate(Me.lbEventList.SelectedItem("StartDateTime").ToString).ToString("MMMM dd, yyyy"))
            Me.labDonationParticipants.Text = String.Format("{0} Participant(s)", Me.dgvDonationList.Rows.Count)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CustomizeFieldsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim CustFieldForm As New frmCustomize(Me.TabControl1.SelectedIndex)
        CustFieldForm.ShowDialog()
        Select Case Me.TabControl1.SelectedIndex
            Case 0
                radgvdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", RoomAssignmentFieldList), Me.lbEventList.SelectedValue))
                Me.dgvRoomAssignments.DataSource = radgvdata
            Case 1
                fathersdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", FatherFieldList), Me.lbEventList.SelectedValue))
                Me.dgvFathersList.DataSource = fathersdata
            Case 2
                kitchendata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", KitchenFieldList), Me.lbEventList.SelectedValue))
                Me.dgvKitchenReport.DataSource = kitchendata
            Case 3
                donationdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", DonationFieldList), Me.lbEventList.SelectedValue))
                Me.dgvDonationList.DataSource = donationdata
            Case 4
                transferdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", TransferFieldList), Me.lbEventList.SelectedValue))
                transferdata = DB_EXECUTE_SELECT(String.Format("select {0} from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and a.FKEventId = {1} and a.Status = 2", String.Join(",", TransferFieldList), Me.lbEventList.SelectedValue))
                Try
                    Dim dataView As New DataView(transferdata)
                    dataView.Sort = "RetreatsSum DESC,Participant Name"
                    Dim newtransdata As DataTable = dataView.ToTable
                    newtransdata.Columns.RemoveAt(2)
                    Me.dgvTransferList.DataSource = newtransdata
                Catch ex As Exception
                End Try
        End Select
    End Sub

    'Private Sub cmsMenu_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
    '    Dim dgv As DataGridView = DirectCast(cmsMenu.SourceControl, DataGridView)
    '    If dgv.Rows.Count = 0 Then
    '        CustomizeFieldsToolStripMenuItem.Enabled = False
    '    Else
    '        CustomizeFieldsToolStripMenuItem.Enabled = True
    '    End If
    'End Sub

    'Protected Overrides ReadOnly Property CreateParams() As CreateParams
    '    Get
    '        Dim cp As CreateParams = MyBase.CreateParams
    '        cp.ExStyle = cp.ExStyle Or &H2000000
    '        Return cp
    '    End Get
    'End Property

    Private Sub CurrentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrentToolStripMenuItem.Click
        Dim saveas As New SaveFileDialog
        saveas.Filter = "Excel File|*.xlsx"
        Select Case Me.TabControl1.SelectedIndex
            Case 0
                saveas.FileName = String.Format("Room Assignments of {0}", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"))
            Case 1
                saveas.FileName = String.Format("Father's List of {0}", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"))
            Case 2
                saveas.FileName = String.Format("Kitchen Report of {0}", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"))
            Case 3
                saveas.FileName = String.Format("Donation List of {0}", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"))
            Case 4
                saveas.FileName = String.Format("Transfer List of {0}", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"))
        End Select
        If saveas.ShowDialog = Windows.Forms.DialogResult.OK Then
            Select Case Me.TabControl1.SelectedIndex
                Case 0
                    GenerateRoomAssigment(radgvdata, saveas.FileName)
                Case 1
                    GenerateFatherList(fathersdata, saveas.FileName)
                Case 2
                    GenerateKitchenReport(kitchendata, saveas.FileName)
                Case 3
                    GenerateDonationList(donationdata, saveas.FileName)
                Case 4
                    GenerateTranserListReport(DirectCast(dgvTransferList.DataSource, DataTable), saveas.FileName)
            End Select
            MessageBox.Show(String.Format("Successfully extracted in {0}.", saveas.FileName), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub GenerateRoomAssigment(ByVal ddtable As DataTable, ByVal outputfile As String)
        Dim ofile As New IO.FileInfo(outputfile)
        If ofile.Directory.Exists = False Then ofile.Directory.Create()
        If ofile.Exists Then ofile.Delete()
        Using pck As New ExcelPackage(ofile)
            Dim ws As ExcelWorksheet = Nothing
            ws = pck.Workbook.Worksheets.Add("Room Assignment")
            ws.Cells(1, 1).Value = String.Format("Room Assignments For Participants Coming on: {0}", CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToString("MMMM dd, yyyy"))
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Merge = True
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.Font.Bold = True
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.Font.Size = 16
            ws.Cells(2, 1).Value = String.Format("{0} Participant(s)", ddtable.Rows.Count)
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Merge = True
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Style.Font.Size = 14
            'column headers
            ws.Cells("A4").LoadFromDataTable(ddtable, True)
            For j As Integer = 4 To ddtable.Rows.Count + 4
                For i As Integer = 1 To ddtable.Columns.Count
                    ws.Cells(j, i).Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    If i <> 1 Then ws.Cells(j, i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    ws.Cells(j, i).Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    If j = 4 Then ws.Cells(j, i).Style.Font.Bold = True
                    If i = 5 Then
                        If ws.Cells(j, i).Value = "Yes" Then
                            ws.Cells(j, i).Value = "✔"
                            ws.Cells(j, i).Style.Font.Size = 20
                            ws.Cells(j, i).AutoFitColumns()
                        End If
                    End If
                    If i <> 5 Then ws.Cells(j, i).Style.Font.Size = 12
                Next
                ws.Row(j).Height = 22
            Next
            ws.Cells.AutoFitColumns()
            ws.PrinterSettings.Orientation = eOrientation.Landscape
            ws.PrinterSettings.LeftMargin = 1 / 2.54
            ws.PrinterSettings.RightMargin = 1 / 2.54
            ws.PrinterSettings.Scale = 71
            pck.Save()
        End Using
    End Sub

    Private Sub GenerateFatherList(ByVal ddtable As DataTable, ByVal outputfile As String)
        Dim ofile As New IO.FileInfo(outputfile)
        If ofile.Directory.Exists = False Then ofile.Directory.Create()
        If ofile.Exists Then ofile.Delete()
        Using pck As New ExcelPackage(ofile)
            Dim ws As ExcelWorksheet = Nothing
            ws = pck.Workbook.Worksheets.Add("Father's List")
            ws.Cells(1, 1).Value = "Father's List"
            ws.Cells(1, 1).Style.Font.SetFromFont(New Font("Times New Roman", 20))
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Merge = True
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.Font.Bold = True
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
            ws.Cells(2, 1).Value = "Manresa Jesuit Retreat Center"
            ws.Cells(2, 1).Style.Font.SetFromFont(New Font("Times New Roman", 10))
            ws.Cells(2, 1).Style.Font.Bold = True
            ws.Cells(2, 1).Style.Font.Italic = True
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Merge = True
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
            ws.Cells(4, 1).Value = "Date:"
            ws.Cells(4, 2).Value = CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToLongDateString
            ws.Cells(4, 2, 4, ddtable.Columns.Count).Merge = True
            ws.Cells(5, 1).Value = "Retreat(s):"
            ws.Cells(5, 2).Value = Me.lbEventList.SelectedItem(1).ToString
            ws.Cells(5, 2, 5, ddtable.Columns.Count).Merge = True
            ws.Cells(6, 1).Value = "No. of Participants:"
            ws.Cells(6, 2).Value = Me.dgvFathersList.Rows.Count
            ws.Cells(6, 2, 6, ddtable.Columns.Count).Merge = True
            ws.Cells(6, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

            ws.Cells(4, 1, 6, 2).Style.Font.SetFromFont(New Font("Times New Roman", 12))
            ws.Cells(4, 1, 6, 2).Style.Font.Bold = True
            ''column headers 
            ws.Cells("A8").LoadFromDataTable(ddtable, True)
            For j As Integer = 8 To ddtable.Rows.Count + 8
                For i As Integer = 1 To ddtable.Columns.Count
                    ws.Cells(j, i).Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    If i <> 1 Then ws.Cells(j, i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    ws.Cells(j, i).Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    If j = 8 Then ws.Cells(j, i).Style.Font.Bold = True
                    ws.Cells(j, i).Style.Font.Size = 12
                Next
                'ws.Row(j).Height = 22
            Next
            ws.Cells.AutoFitColumns()
            pck.Save()
        End Using
    End Sub

    Private Sub GenerateKitchenReport(ByVal ddtable As DataTable, ByVal outputfile As String)
        Dim ofile As New IO.FileInfo(outputfile)
        If ofile.Exists Then ofile.Delete()
        If ofile.Directory.Exists = False Then ofile.Directory.Create()
        Using pck As New ExcelPackage(ofile)
            Dim ws As ExcelWorksheet = Nothing
            ws = pck.Workbook.Worksheets.Add("Kitchen Report")
            ws.Cells(1, 1).Value = "Kitchen Report"
            ws.Cells(1, 1).Style.Font.SetFromFont(New Font("Times New Roman", 20))
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Merge = True
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.Font.Bold = True
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
            ws.Cells(2, 1).Value = "Manresa Jesuit Retreat Center"
            ws.Cells(2, 1).Style.Font.SetFromFont(New Font("Times New Roman", 10))
            ws.Cells(2, 1).Style.Font.Bold = True
            ws.Cells(2, 1).Style.Font.Italic = True
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Merge = True
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
            ws.Cells(4, 1).Value = "Date:"
            ws.Cells(4, 2).Value = CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToLongDateString
            ws.Cells(4, 2, 4, ddtable.Columns.Count).Merge = True
            ws.Cells(5, 1).Value = "Retreat(s):"
            ws.Cells(5, 2).Value = Me.lbEventList.SelectedItem(1).ToString
            ws.Cells(5, 2, 5, ddtable.Columns.Count).Merge = True
            ws.Cells(6, 1).Value = "No. of Participants:"
            ws.Cells(6, 2).Value = Me.dgvFathersList.Rows.Count
            ws.Cells(6, 2, 6, ddtable.Columns.Count).Merge = True
            ws.Cells(6, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

            ws.Cells(4, 1, 6, 2).Style.Font.SetFromFont(New Font("Times New Roman", 12))
            ws.Cells(4, 1, 6, 2).Style.Font.Bold = True
            ''column headers 
            ws.Cells("A8").LoadFromDataTable(ddtable, True)
            For j As Integer = 8 To ddtable.Rows.Count + 8
                For i As Integer = 1 To ddtable.Columns.Count
                    ws.Cells(j, i).Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    If i <> 1 Then ws.Cells(j, i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    ws.Cells(j, i).Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    If j = 8 Then ws.Cells(j, i).Style.Font.Bold = True
                    ws.Cells(j, i).Style.Font.Size = 12
                Next
                'ws.Row(j).Height = 22
            Next
            ws.Cells.AutoFitColumns()
            pck.Save()
        End Using
    End Sub

    Private Sub GenerateDonationList(ByVal ddtable As DataTable, ByVal outputfile As String)
        Dim ofile As New IO.FileInfo(outputfile)
        Dim fundname As String = DB_EXECUTE_SELECT(String.Format("select value from donor_funds where id={0}", Me.lbEventList.SelectedItem("fkfundid"))).Rows(0)(0)
        If ofile.Exists Then ofile.Delete()
        If ofile.Directory.Exists = False Then ofile.Directory.Create()
        Using pck As New ExcelPackage(ofile)
            Dim ws As ExcelWorksheet = Nothing
            ws = pck.Workbook.Worksheets.Add("Donation List")
            ws.Cells(1, 1).Value = fundname
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Merge = True
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
            ws.Cells(1, 1, 1, ddtable.Columns.Count).Style.Font.Size = 14
            ws.Cells(2, 1).Value = String.Format("Donation List For Participants Coming on: {0}", CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToString("MMMM dd, yyyy"))
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Merge = True
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Style.Font.Bold = True
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells(2, 1, 2, ddtable.Columns.Count).Style.Font.Size = 16
            ws.Cells(3, 1).Value = String.Format("{0} Participant(s)", ddtable.Rows.Count)
            ws.Cells(3, 1, 3, ddtable.Columns.Count).Merge = True
            ws.Cells(3, 1, 3, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells(3, 1, 3, ddtable.Columns.Count).Style.Font.Size = 14
            'column headers
            ws.Cells("A5").LoadFromDataTable(ddtable, True)
            For j As Integer = 5 To ddtable.Rows.Count + 5
                For i As Integer = 1 To ddtable.Columns.Count
                    ws.Cells(j, i).Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    If i <> 1 Then
                        ws.Cells(j, i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    End If
                    ws.Cells(j, i).Style.VerticalAlignment = ExcelVerticalAlignment.Center

                    If j = 5 Then ws.Cells(j, i).Style.Font.Bold = True
                    ws.Cells(j, i).Style.Font.Size = 12
                Next
                ws.Row(j).Height = 22
            Next
            ws.Cells.AutoFitColumns()
            pck.Save()
        End Using
    End Sub

    Private Sub GenerateTranserListReport(ByVal ddtable As DataTable, ByVal outputfile As String)
        Dim ofile As New IO.FileInfo(outputfile)
        If ofile.Exists Then ofile.Delete()
        Using pck As New ExcelPackage(ofile)
            Dim ws As ExcelWorksheet = Nothing
            ws = pck.Workbook.Worksheets.Add("Transfer List")
            ws.Cells(1, 1).Value = "Participant list in order of Number of Retreats for Program"
            ws.Cells(1, 1).Style.Font.SetFromFont(New Font("Times New Roman", 14))
            ws.Column(1).Width = 40
            ws.Column(2).Width = 40
            ws.Cells(1, 1, 2, ddtable.Columns.Count).Merge = True
            ws.Cells(1, 1, 2, ddtable.Columns.Count).Style.Font.Bold = True
            ws.Cells(1, 1, 2, ddtable.Columns.Count).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
            ws.Cells(1, 1, 2, ddtable.Columns.Count).Style.VerticalAlignment = ExcelHorizontalAlignment.General
            ws.Cells(2, 1).Value = "Date:"
            ws.Cells(2, 2).Value = CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToLongDateString
            ws.Cells(3, 1).Value = "Retreat(s):"
            ws.Cells(3, 2).Value = Me.lbEventList.SelectedItem(1).ToString
            ws.Cells(4, 1).Value = "No. of Participants:"
            ws.Cells(4, 2).Value = Me.dgvFathersList.Rows.Count
            ws.Cells(4, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
            'ws.Cells(4, 1, 6, 2).Style.Font.SetFromFont(New Font("Times New Roman", 12))
            ws.Cells(5, 1, 5, 2).Style.Font.Bold = True
            ''column headers 
            ws.Cells("A5").LoadFromDataTable(ddtable, True)
            For j As Integer = 5 To ddtable.Rows.Count + 5
                For i As Integer = 1 To ddtable.Columns.Count
                    ws.Cells(j, i).Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    If i = 2 Then ws.Cells(j, i).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    ws.Cells(j, i).Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    ws.Cells(j, i).Style.Font.Size = 12
                Next
                'ws.Row(j).Height = 22
            Next
            'ws.Cells.AutoFitColumns()
            pck.Save()
        End Using
    End Sub

    Private Sub AllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllToolStripMenuItem.Click
        Dim saveas As New SaveFileDialog
        saveas.Filter = "Excel File|*.xlsx"
        saveas.FileName = "All Reports"
        If saveas.ShowDialog = Windows.Forms.DialogResult.OK Then
            GenerateRoomAssigment(radgvdata, String.Format("{0}\Room Assignments of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            GenerateFatherList(fathersdata, String.Format("{0}\Father's List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            GenerateKitchenReport(kitchendata, String.Format("{0}\Kitchen Report of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            GenerateDonationList(donationdata, String.Format("{0}\Donation List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            Try
                Dim dataView As New DataView(transferdata)
                dataView.Sort = "RetreatsSum DESC"
                Dim newtransdata As DataTable = dataView.ToTable
                newtransdata.Columns.RemoveAt(2)
                GenerateTranserListReport(newtransdata, String.Format("{0}\Transfer List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            Catch ex As Exception
            End Try
            MessageBox.Show(String.Format("Successfully extracted in {0}.", IO.Path.GetDirectoryName(saveas.FileName)), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub dgvTransferList_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTransferList.DataSourceChanged
        Try
            Me.labTransDate.Text = CDate(Me.lbEventList.SelectedItem("StartDateTime")).ToLongDateString
            Me.labTransRetreat.Text = Me.lbEventList.SelectedItem(1).ToString
            Me.labTransParticipants.Text = Me.dgvTransferList.Rows.Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub butExportToFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butExportToFile.Click
        If MessageBox.Show("Do you want me to save the files in S:\RetreatReportsFromGiftworks?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.Cursor = Cursors.WaitCursor
            If IO.Directory.Exists("S:\RetreatReportsFromGiftworks") Then
                Dim evname As String = Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-").Trim
                If evname.Length > 20 Then
                    evname = evname.Substring(0, 23) & "-etc"
                End If
                GenerateRoomAssigment(radgvdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Room Assignments of {1}.xlsx", evname, evname))
                GenerateFatherList(fathersdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Father's List of {1}.xlsx", evname, evname))
                GenerateKitchenReport(kitchendata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Kitchen Report of {1}.xlsx", evname, evname))
                GenerateDonationList(donationdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Donation List of {1}.xlsx", evname, evname))
                'GenerateRoomAssigment(radgvdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Room Assignments of {1}.xlsx", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
                'GenerateFatherList(fathersdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Father's List of {1}.xlsx", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
                'GenerateKitchenReport(kitchendata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Kitchen Report of {1}.xlsx", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
                'GenerateDonationList(donationdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Donation List of {1}.xlsx", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
                Try
                    Dim dataView As New DataView(transferdata)
                    dataView.Sort = "RetreatsSum DESC"
                    Dim newtransdata As DataTable = dataView.ToTable
                    newtransdata.Columns.RemoveAt(2)
                    GenerateTranserListReport(newtransdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Transfer List of {1}.xlsx", evname, evname))
                    'GenerateTranserListReport(newtransdata, String.Format("S:\RetreatReportsFromGiftworks\{0}\Transfer List of {1}.xlsx", Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-"), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
                Catch ex As Exception
                End Try
                Me.Cursor = Cursors.Default
                MessageBox.Show("Successfully extracted in S:\RetreatReportsFromGiftworks.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                CustomSaving()
            End If
        Else
                CustomSaving()
        End If
    End Sub

    Private Sub CustomSaving()
        Dim saveas As New SaveFileDialog
        saveas.Filter = "Excel File|*.xlsx"
        saveas.FileName = "All Reports"
        If saveas.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim evname As String = Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-").Trim
            If evname.Length > 20 Then
                evname = evname.Substring(0, 23) & "-etc"
            End If
            GenerateRoomAssigment(radgvdata, String.Format("{0}\Room Assignments of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), evname))
            GenerateFatherList(fathersdata, String.Format("{0}\Father's List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), evname))
            GenerateKitchenReport(kitchendata, String.Format("{0}\Kitchen Report of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), evname))
            GenerateDonationList(donationdata, String.Format("{0}\Donation List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), evname))
            'GenerateRoomAssigment(radgvdata, String.Format("{0}\Room Assignments of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            'GenerateFatherList(fathersdata, String.Format("{0}\Father's List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            'GenerateKitchenReport(kitchendata, String.Format("{0}\Kitchen Report of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            'GenerateDonationList(donationdata, String.Format("{0}\Donation List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            Try
                Dim dataView As New DataView(transferdata)
                dataView.Sort = "RetreatsSum DESC"
                Dim newtransdata As DataTable = dataView.ToTable
                newtransdata.Columns.RemoveAt(2)
                GenerateTranserListReport(newtransdata, String.Format("{0}\Transfer List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), evname))
                'GenerateTranserListReport(newtransdata, String.Format("{0}\Transfer List of {1}.xlsx", IO.Path.GetDirectoryName(saveas.FileName), Me.lbEventList.SelectedItem(1).ToString.Replace("/", "-")))
            Catch ex As Exception
            End Try
            MessageBox.Show(String.Format("Successfully extracted in {0}.", IO.Path.GetDirectoryName(saveas.FileName)), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub butPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butPrint.Click
        Select Case TabControl1.SelectedIndex
            Case 0
                DataGridViewPrinter.StartPrint(Me.dgvRoomAssignments, True, True, "Retreat: " & Me.labRATitle.Text & " with " & Me.labRASubTitle.Text)
            Case 1
                DataGridViewPrinter.StartPrint(Me.dgvFathersList, False, True, "Retreat: " & Me.labFatherRetreat.Text & vbTab & "No. of Participants: " & Me.labFatherNumParticipant.Text)
            Case 2
                DataGridViewPrinter.StartPrint(Me.dgvKitchenReport, False, True, "Retreat: " & Me.labKitRetreat.Text & vbTab & "No. of Participants: " & Me.labKitNumParticipants.Text)
            Case 3
                DataGridViewPrinter.StartPrint(Me.dgvDonationList, False, True, "Retreat: " & Me.labDonationTitle.Text & vbTab & "No. of Participants: " & Me.labDonationParticipants.Text)
            Case 4
                DataGridViewPrinter.StartPrint(Me.dgvTransferList, False, True, "Retreat: " & Me.labTransRetreat.Text & vbTab & "No. of Participants: " & Me.labTransParticipants.Text)
                'Case 5
                '    DataGridViewPrinter.StartPrint(Me.dgvCaptains, True, True, "Guests Registered For Future Retreats")
        End Select

    End Sub



    Private Sub butGenerateCaptainsList_Click(sender As Object, e As EventArgs) Handles butGenerateCaptainsList.Click
        captainprev = New frmCaptainsListPreview()
        captainprev.ShowDialog()
        ''captains report
        'captainsdata = DB_EXECUTE_SELECT("select CONVERT(VARCHAR(10),b.StartDateTime,110) as 'Start Date',b.Name as 'Event Name',c.lName as 'Last Name',c.fName as 'First Name',c.mName as 'Middle Name',c.suffix as 'Suffix' from events_event_participants a inner join events_event_info b on a.FKEventId = b.Id inner join donor_donors c on a.FKDonorId = c.id where a.id > 0 and b.FKCategoryId = 7 and b.StartDateTime > GETDATE() and a.Status = 2;")
        'Dim ofolder As New FolderBrowserDialog
        'ofolder.RootFolder = Environment.SpecialFolder.MyComputer
        'If ofolder.ShowDialog = DialogResult.OK Then
        '    Dim outfolder As New IO.DirectoryInfo(ofolder.SelectedPath & "CaptainsList")
        '    If outfolder.Exists Then outfolder.Delete(True)
        '    If outfolder.Exists = False Then outfolder.Create()
        '    Dim futuredatelist As New List(Of String)
        '    For Each frow As DataRow In captainsdata.Rows
        '        If futuredatelist.Contains(frow(0).ToString) = False Then futuredatelist.Add(frow(0).ToString)
        '    Next
        '    Dim futcounter As Integer = 0
        '    For Each futstring As String In futuredatelist
        '        futcounter += 1
        '        Dim ftable As DataTable = captainsdata.Select(String.Format("[Start Date] = '{0}'", futstring)).CopyToDataTable
        '        GenerateCaptainsReport(String.Concat(outfolder, "\", futstring, ".xlsx"), ftable, IIf(futcounter = 1, True, False))
        '    Next
        '    MessageBox.Show(String.Format("Successfully extracted in {0}.", outfolder), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'End If

    End Sub

    Public Sub SetPrintablePageHeight(ws As ExcelWorksheet, newRowCount As Integer)
        If newRowCount <= 0 Then
            Throw New NotSupportedException("Row count must be a positive integer")
        End If

        Dim currentEndRow As Integer = ws.Dimension.[End].Row
        Dim currentEndCol As String = GetColumnLetterFromAddress(ws.PrinterSettings.PrintArea.[End].Address)
        Dim newEndAddress As String = String.Concat(currentEndCol, newRowCount)
        Dim startAddress As String = ws.PrinterSettings.PrintArea.Start.Address
        ws.PrinterSettings.PrintArea = ws.Cells(Convert.ToString(startAddress & Convert.ToString(":")) & newEndAddress)

        ' remove any other page breaks
        For row As Integer = 1 To currentEndRow
            If ws.Row(row).PageBreak Then
                ws.Row(row).PageBreak = False
            End If
        Next

        ws.Row(newRowCount).PageBreak = True
    End Sub

    Private Function GetColumnLetterFromAddress(address As String) As String
        Dim ret As String = String.Empty
        For Each c As Char In address
            If [Char].IsLetter(c) Then
                ret += c
            End If
        Next
        Return ret
    End Function


End Class
