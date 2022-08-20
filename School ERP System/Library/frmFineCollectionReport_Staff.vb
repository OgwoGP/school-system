Imports System.Data.SqlClient

Public Class frmFineCollectionReport_Staff

    Sub Reset()
        DateTimePicker1.Text = Now
        DateTimePicker2.Text = Today
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT StaffName,Designation,Book.AccessionNo,BookTitle,Fine FROM (((Staff INNER JOIN BookIssue_Staff ON Staff.St_ID = BookIssue_Staff.StaffID) INNER JOIN Return_Staff ON BookIssue_Staff.ID = Return_Staff.IssueID) INNER JOIN Book ON BookIssue_Staff.AccessionNo = Book.AccessionNo) where ReturnDate between @d1 and @d2 and Fine > 0 order by returndate", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker2.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker1.Value.Date
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            DataGridView1.DataSource = dtable
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("FineCollection_Staff.xml")
            Dim rpt As New rptFine_Staff
            rpt.SetDataSource(ds)
            rpt.SetParameterValue("p1", DateTimePicker2.Value.Date)
            rpt.SetParameterValue("p2", DateTimePicker1.Value.Date)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub DateTimePicker1_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker1.Validating
        If (DateTimePicker2.Value.Date) > (DateTimePicker1.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DateTimePicker1.Focus()
        End If
    End Sub
End Class
