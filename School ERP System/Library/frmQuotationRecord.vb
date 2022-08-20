Imports System.Data.SqlClient

Public Class frmQuotationRecord
    Sub fillOrderNo()
        Try
            Dim CN As New SqlConnection(cs)
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(OrderNo) FROM Quotation", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbOrderNo.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbOrderNo.Items.Add(drow(0).ToString())
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(OrderID) as [Order ID], RTRIM(OrderNo) as [Order No.], (OrderDate) as [Order Date],RTRIM(DepartmentID) as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM([From]) as [Sender's Name], RTRIM(From_Address) as [Sender's Address], RTRIM(From_ContactNo) as [Sender's CN], RTRIM([To]) as [Receiver's Name], RTRIM(To_Address) as [Receiver's Address], RTRIM(To_ContactNo) as [Receiver's CN],  RTRIM(Remarks) as [Remarks] from Quotation,Department where Quotation.DepartmentID=Department.ID order by OrderDate", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Order")
            dgw.DataSource = myDataSet.Tables("Order").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
        fillOrderNo()
    End Sub
    Sub Reset()
        cmbOrderNo.Text = ""
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Now
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
                Me.Hide()
                frmQuotation.Show()
                ' or simply use column name instead of index
                'dr.Cells["id"].Value.ToString();
                frmQuotation.txtOrderID.Text = dr.Cells(0).Value.ToString()
                frmQuotation.txtOrderNo.Text = dr.Cells(1).Value.ToString()
                frmQuotation.dtpOrderDate.Text = dr.Cells(2).Value.ToString()
                frmQuotation.txtDepartmentID.Text = dr.Cells(3).Value.ToString()
                frmQuotation.cmbDepartment.Text = dr.Cells(4).Value.ToString()
                frmQuotation.txtSenderName.Text = dr.Cells(5).Value.ToString()
                frmQuotation.txtSenderAddress.Text = dr.Cells(6).Value.ToString()
                frmQuotation.txtSenderContactNo.Text = dr.Cells(7).Value.ToString()
                frmQuotation.txtReceiverName.Text = dr.Cells(8).Value.ToString()
                frmQuotation.txtReceiverAddress.Text = dr.Cells(9).Value.ToString()
                frmQuotation.txtReceiverCN.Text = dr.Cells(10).Value.ToString()
                frmQuotation.txtRemarks.Text = dr.Cells(11).Value.ToString()
                frmQuotation.btnSave.Enabled = False
                frmQuotation.btnDelete.Enabled = True
                frmQuotation.btnUpdate.Enabled = True
                frmQuotation.btnPrint.Enabled = True
                frmQuotation.btnRemove.Enabled = False
                con = New SqlConnection(cs)
                con.Open()
                Dim Sql As String = "Select Book.AccessionNo,BookTitle,Author,Publisher,Qty from Quotation,Quotation_Join,Department,Book where Quotation.OrderId=Quotation_Join.QJ_OrderID and Quotation.DepartmentID=Department.ID and Book.AccessionNo=Quotation_Join.AccessionNo and OrderID=" & dr.Cells(0).Value & ""
                cmd = New SqlCommand(Sql, con)
                rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                frmQuotation.DataGridView1.Rows.Clear()
                While (rdr.Read() = True)
                    frmQuotation.DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4))
                End While
                con.Close()
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(OrderID) as [Order ID], RTRIM(OrderNo) as [Order No.], (OrderDate) as [Order Date],RTRIM(DepartmentID) as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM([From]) as [Sender's Name], RTRIM(From_Address) as [Sender's Address], RTRIM(From_ContactNo) as [Sender's CN], RTRIM([To]) as [Receiver's Name], RTRIM(To_Address) as [Receiver's Address], RTRIM(To_ContactNo) as [Receiver's CN],  RTRIM(Remarks) as [Remarks] from Quotation,Department where Quotation.DepartmentID=Department.ID and OrderDate between @d1 and @d2 order by OrderDate", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Order")
            dgw.DataSource = myDataSet.Tables("Order").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbBillNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOrderNo.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(OrderID) as [Order ID], RTRIM(OrderNo) as [Order No.], (OrderDate) as [Order Date],RTRIM(DepartmentID) as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM([From]) as [Sender's Name], RTRIM(From_Address) as [Sender's Address], RTRIM(From_ContactNo) as [Sender's CN], RTRIM([To]) as [Receiver's Name], RTRIM(To_Address) as [Receiver's Address], RTRIM(To_ContactNo) as [Receiver's CN],  RTRIM(Remarks) as [Remarks] from Quotation,Department where Quotation.DepartmentID=Department.ID and OrderNo='" & cmbOrderNo.Text & "' order by OrderDate", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Order")
            dgw.DataSource = myDataSet.Tables("Order").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtpDateTo_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles dtpDateTo.Validating
        If (dtpDateFrom.Value.Date) > (dtpDateTo.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtpDateTo.Focus()
        End If
    End Sub
End Class
