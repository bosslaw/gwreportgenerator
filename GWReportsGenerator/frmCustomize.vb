Public Class frmCustomize
    Private CurSelIndex As Integer

    Public Sub New(ByVal _selindx As Integer)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        CurSelIndex = _selindx
    End Sub

    'Private Sub frmCustomize_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    frmMain.Activate() : GC.SuppressFinalize(Me) : Me.Dispose()
    'End Sub

    'Private Sub frmCustomize_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    GC.Collect()
    'End Sub

    'Private Sub frmCustomize_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    LoadAllFields()
    'End Sub

    'Private Sub LoadAllFields()
    '    Me.clbFields.Items.Clear()
    '    For Each FLDRow As DataRow In FieldsData.Rows
    '        Me.clbFields.Items.Add(FLDRow(1).ToString)
    '    Next
    '    CheckFields()
    'End Sub

    'Private Sub CheckFields()
    '    For i As Integer = 0 To clbFields.Items.Count - 1

    '        Dim lrow As DataTable = FieldsData.Select(String.Format("Fake = '{0}'", clbFields.Items(i).ToString))
    '        Select Case CurSelIndex
    '            Case 0
    '                If RoomAssignmentFieldList.Contains(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'")) Then
    '                    clbFields.SetItemCheckState(i, 1)
    '                End If
    '            Case 1
    '                If FatherFieldList.Contains(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'")) Then
    '                    clbFields.SetItemCheckState(i, 1)
    '                End If
    '            Case 2
    '                If KitchenFieldList.Contains(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'")) Then
    '                    clbFields.SetItemCheckState(i, 1)
    '                End If
    '            Case 3
    '                If DonationFieldList.Contains(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'")) Then
    '                    clbFields.SetItemCheckState(i, 1)
    '                End If
    '        End Select
    '    Next
    'End Sub

    'Private Sub butOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butOK.Click
    '    For i As Integer = 0 To clbFields.Items.Count - 1
    '        ' Get the selected item's check state.
    '        Dim lrow As DataRow() = FieldsData.Select(String.Format("Fake = '{0}'", clbFields.Items(i).ToString))
    '        Dim chkstate As CheckState
    '        chkstate = clbFields.GetItemCheckState(i)
    '        If chkstate = 1 Then
    '            Select Case CurSelIndex
    '                Case 0
    '                    RoomAssignmentFieldList.Add(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '                Case 1
    '                    FatherFieldList.Add(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '                Case 2
    '                    KitchenFieldList.Add(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '                Case 3
    '                    DonationFieldList.Add(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '            End Select
    '        Else
    '            Select Case CurSelIndex
    '                Case 0
    '                    RoomAssignmentFieldList.Remove(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '                Case 1
    '                    FatherFieldList.Remove(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '                Case 2
    '                    KitchenFieldList.Remove(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '                Case 3
    '                    DonationFieldList.Remove(String.Concat(lrow(0)("Real").ToString, " as '", lrow(0)("Fake").ToString, "'"))
    '            End Select
    '        End If
    '    Next
    '    Me.Close()
    'End Sub
End Class