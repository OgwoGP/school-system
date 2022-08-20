Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmAdvanceEntryRecord

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Public Sub GetData()
        Try
            Total.Visible = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(AdvanceEntry.ID) as [ID], Convert(DateTime,workingdate,131) as [Entry Date],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(Amount) as [Advance] from Advanceentry,Staff where Staff.St_ID=AdvanceEntry.StaffID and Amount > 0 order by workingdate", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "AdvanceEntry")
            myDA.Fill(myDataSet, "Staff")
            dgw.DataSource = myDataSet.Tables("AdvanceEntry").DefaultView
            dgw.DataSource = myDataSet.Tables("Staff").DefaultView
            Dim sum As Double = 0
            For Each r As DataGridViewRow In Me.dgw.Rows
                sum = sum + r.Cells(5).Value
            Next
            sum = Math.Round(sum, 2)
            TotalAdvance.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub
    Sub Reset()
        txtStaffName.Text = ""
        DateFrom.Value = Today
        DateTo.Value = Now
        Total.Visible = False
        GetData()
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnExportExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub dgw_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Advance Entry" Then
                    Me.Hide()
                    frmAdvanceEntry.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    frmAdvanceEntry.txtID.Text = dr.Cells(0).Value.ToString()
                    frmAdvanceEntry.dtpEntryDate.Text = dr.Cells(1).Value.ToString()
                    frmAdvanceEntry.txtStID.Text = dr.Cells(2).Value.ToString()
                    frmAdvanceEntry.txtStaffID.Text = dr.Cells(3).Value.ToString()
                    frmAdvanceEntry.txtStaffName.Text = dr.Cells(4).Value.ToString()
                    frmAdvanceEntry.txtAmount.Text = dr.Cells(5).Value.ToString()
                    frmAdvanceEntry.btnSave.Enabled = False
                    frmAdvanceEntry.btnUpdate.Enabled = True
                    frmAdvanceEntry.btnDelete.Enabled = True
                    frmAdvanceEntry.dtpEntryDate.Enabled = False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub dgw_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub txtStaffName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            Total.Visible = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(AdvanceEntry.ID) as [ID], Convert(DateTime,workingdate,131) as [Entry Date],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(Amount) as [Advance] from Advanceentry,Staff where Staff.St_ID=AdvanceEntry.StaffID and Amount > 0 and StaffName like '%" & txtStaffName.Text & "%' order by workingdate", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "AdvanceEntry")
            myDA.Fill(myDataSet, "Staff")
            dgw.DataSource = myDataSet.Tables("AdvanceEntry").DefaultView
            dgw.DataSource = myDataSet.Tables("Staff").DefaultView
            Dim sum As Double = 0
            For Each r As DataGridViewRow In Me.dgw.Rows
                sum = sum + r.Cells(5).Value
            Next
            sum = Math.Round(sum, 2)
            TotalAdvance.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            Total.Visible = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(AdvanceEntry.ID) as [ID], Convert(DateTime,workingdate,131) as [Entry Date],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(Amount) as [Advance] from Advanceentry,Staff where Staff.St_ID=AdvanceEntry.StaffID and Amount > 0 and WorkingDate Between @d1 and @d2 order by workingdate", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "DateIN").Value = DateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "DateIN").Value = DateTo.Value
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "AdvanceEntry")
            myDA.Fill(myDataSet, "Staff")
            dgw.DataSource = myDataSet.Tables("AdvanceEntry").DefaultView
            dgw.DataSource = myDataSet.Tables("Staff").DefaultView
            Dim sum As Double = 0
            For Each r As DataGridViewRow In Me.dgw.Rows
                sum = sum + r.Cells(5).Value
            Next
            sum = Math.Round(sum, 2)
            TotalAdvance.Text = sum
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
