Public Class frmReview
    Private indata As DataTable
    Private tempdata As DataTable

    Public Sub New(ByVal _indata As DataTable)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        indata = _indata
    End Sub

    Private Sub frmReview_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        captainprev.Activate() : GC.SuppressFinalize(Me) : Me.Dispose()
    End Sub

    Private Sub frmReview_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        GC.Collect()
    End Sub

    Private Sub frmReview_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        NewCaptainsData = Nothing
        Dim distinctDT As DataTable = indata.DefaultView.ToTable(True, "Event Name")
        tempdata = New DataTable
        tempdata.Columns.Add("Include", GetType(Boolean))
        tempdata.Columns.Add("Event Name")
        For Each drow As DataRow In distinctDT.Rows
            tempdata.Rows.Add(True, drow(0).ToString)
        Next
        Me.dgvRecords.DataSource = tempdata
    End Sub

    Private Sub dgvRecords_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRecords.DataSourceChanged
        Me.dgvRecords.Columns(0).FillWeight = 10
        Me.dgvRecords.Columns(0).ReadOnly = False
        Me.dgvRecords.Columns(1).ReadOnly = True
        If Me.dgvRecords.RowCount = 0 Then
            Me.ckCheckAll.Enabled = False
        Else
            Me.ckCheckAll.Enabled = True
        End If
    End Sub

    Private Sub ckCheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ckCheckAll.CheckedChanged
        Try
            If Me.ckCheckAll.Checked Then
                For Each dtrow As DataRow In tempdata.Rows
                    dtrow(0) = True
                Next
                Me.ckCheckAll.Text = "Uncheck All"
            Else
                For Each dtrow As DataRow In tempdata.Rows
                    dtrow(0) = False
                Next
                Me.ckCheckAll.Text = "Check All"
            End If

            Me.dgvRecords.DataSource = tempdata
        Catch ex As Exception
        End Try
    End Sub

    Private Sub butConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butConfirm.Click
        NewCaptainsData = indata.Copy
        For Each temprow As DataRow In tempdata.Rows
            If temprow(0) = False Then
                Dim qrow As DataRow() = NewCaptainsData.Select(String.Format("NOT [Event Name] = '{0}'", temprow(1).ToString))
                NewCaptainsData = qrow.CopyToDataTable
            End If
        Next
        NewCaptainsData.Columns.Add("Order")
        For Each drow As DataRow In NewCaptainsData.Rows
            Dim petsa As String() = drow(0).ToString.Split("-")
            drow(NewCaptainsData.Columns.Count - 1) = String.Concat(petsa(2), petsa(0), petsa(1))
        Next
        NewCaptainsData.DefaultView.Sort = "Order ASC"
        NewCaptainsData = NewCaptainsData.DefaultView.ToTable
        NewCaptainsData.Columns.RemoveAt(NewCaptainsData.Columns.Count - 1)
        Me.Close()
    End Sub
End Class