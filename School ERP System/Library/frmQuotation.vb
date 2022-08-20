Imports System.Data.SqlClient
Public Class frmQuotation
  
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Sub Reset()
        txtAccessionNo.Text = ""
        cmbDepartment.SelectedIndex = -1
        txtAuthor.Text = ""
        txtBookTitle.Text = ""
        txtReceiverCN.Text = ""
        txtOrderID.Text = ""
        txtOrderNo.Text = ""
        txtPublisher.Text = ""
        txtQty.Text = ""
        txtReceiverAddress.Text = ""
        txtReceiverName.Text = ""
        txtSenderAddress.Text = ""
        txtSenderContactNo.Text = ""
        txtSenderName.Text = ""
        dtpOrderDate.Text = Today
        DataGridView1.Rows.Clear()
        btnPrint.Enabled = False
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        btnAdd.Enabled = True
        btnRemove.Enabled = False
        Clear()
        auto()
    End Sub

    Sub Clear()
        txtAuthor.Text = ""
        txtBookTitle.Text = ""
        txtAccessionNo.Text = ""
        txtQty.Text = ""
        txtPublisher.Text = ""
    End Sub

    Private Sub auto()
        Dim Num As Integer = 0
        con = New SqlConnection(cs)
        con.Open()
        Dim Sql As String = ("SELECT MAX(OrderID) from Quotation")
        cmd = New SqlCommand(Sql)
        cmd.Connection = con
        If (IsDBNull(cmd.ExecuteScalar)) Then
            Num = 1
            txtOrderID.Text = Num.ToString
            txtOrderNo.Text = "OD-" + Num.ToString
        Else
            Num = cmd.ExecuteScalar + 1
            txtOrderID.Text = Num.ToString
            txtOrderNo.Text = "OD-" + Num.ToString
        End If
        cmd.Dispose()
        con.Close()
        con.Dispose()
    End Sub
    Sub fillCombo()
        Try
            Dim CN As New SqlConnection(cs)
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct Departmentname FROM Department", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            Dim dtable As DataTable = ds.Tables(0)
            cmbDepartment.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbDepartment.Items.Add(drow(0).ToString())
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Print()
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptQuotation 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Quotation.OrderID, Quotation.OrderNo, Quotation.[From], Quotation.From_Address, Quotation.From_ContactNo, Quotation.[To], Quotation.To_Address, Quotation.To_ContactNo, Quotation.OrderDate,Quotation.DepartmentID, Quotation.Remarks, Quotation_Join.QJ_ID, Quotation_Join.AccessionNo,Quotation_Join.Qty, Quotation_Join.QJ_OrderID, Book.AccessionNo AS Expr1, Book.BookTitle, Book.EntryDate, Book.Author,Book.JointAuthors, Book.SubCategoryID, Book.Barcode, Book.ISBN, Book.Volume, Book.Edition, Book.Publisher, Book.PlaceOfPublisher, Book.PublishingYear, Book.[Section], Book.[Language],Book.BookPosition, Book.Price, Book.SupplierID, Book.BillNo, Book.BillDate, Book.NoOfPages, Book.Condition, Book.Status, Book.Remarks AS Expr2, Department.ID, Department.DepartmentName FROM (((Quotation INNER JOIN Quotation_Join ON Quotation.OrderID = Quotation_Join.QJ_OrderID) INNER JOIN Book ON Quotation_Join.AccessionNo = Book.AccessionNo) INNER JOIN Department ON Quotation.DepartmentID = Department.ID) where OrderNo='" & txtOrderNo.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Quotation")
            myDA.Fill(myDS, "Quotation_Join")
            myDA.Fill(myDS, "Department")
            myDA.Fill(myDS, "Book")
            rpt.SetDataSource(myDS)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "delete from Quotation where OrderID=" & txtOrderID.Text & ""
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            If RowsAffected > 0 Then
                Dim st As String = "deleted the quotation having order no.'" & txtOrderNo.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        Try
            If txtAccessionNo.Text = "" Then
                MessageBox.Show("Please retrieve accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAccessionNo.Focus()
                Exit Sub
            End If
            If txtQty.Text = "" Then
                MessageBox.Show("Please enter quantity", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtQty.Focus()
                Exit Sub
            End If
            If DataGridView1.Rows.Count = 0 Then
                DataGridView1.Rows.Add(txtAccessionNo.Text, txtBookTitle.Text, txtAuthor.Text, txtPublisher.Text, txtQty.Text)
                Clear()
                Exit Sub
            End If
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                If r.Cells(0).Value = txtAccessionNo.Text Then
                    r.Cells(0).Value = txtAccessionNo.Text
                    r.Cells(1).Value = txtBookTitle.Text
                    r.Cells(2).Value = txtAuthor.Text
                    r.Cells(3).Value = txtPublisher.Text
                    r.Cells(4).Value = Val(r.Cells(4).Value) + Val(txtQty.Text)
                    Clear()
                    Exit Sub
                End If
            Next
            DataGridView1.Rows.Add(txtAccessionNo.Text, txtBookTitle.Text, txtAuthor.Text, txtPublisher.Text, txtQty.Text)
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As System.EventArgs) Handles btnRemove.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
            Next
            btnRemove.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_MouseClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        btnRemove.Enabled = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub cmbDepartment_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbDepartment.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID FROM Department where DepartmentName=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbDepartment.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtDepartmentID.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Try
            If Len(Trim(cmbDepartment.Text)) = 0 Then
                MessageBox.Show("Please select department", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbDepartment.Focus()
                Exit Sub
            End If
            If Len(Trim(txtSenderName.Text)) = 0 Then
                MessageBox.Show("Please enter sender's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSenderName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtReceiverName.Text)) = 0 Then
                MessageBox.Show("Please enter receiver's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtReceiverName.Focus()
                Exit Sub
            End If

            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("sorry no data added to grid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Quotation(OrderID, OrderNo, OrderDate,DepartmentID, [From], From_Address, From_ContactNo, [To], To_Address, To_ContactNo,  Remarks) Values (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtOrderID.Text)
            cmd.Parameters.AddWithValue("@d2", txtOrderNo.Text)
            cmd.Parameters.AddWithValue("@d3", dtpOrderDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtDepartmentID.Text)
            cmd.Parameters.AddWithValue("@d5", txtSenderName.Text)
            cmd.Parameters.AddWithValue("@d6", txtSenderAddress.Text)
            cmd.Parameters.AddWithValue("@d7", txtSenderContactNo.Text)
            cmd.Parameters.AddWithValue("@d8", txtReceiverName.Text)
            cmd.Parameters.AddWithValue("@d9", txtReceiverAddress.Text)
            cmd.Parameters.AddWithValue("@d10", txtReceiverCN.Text)
            cmd.Parameters.AddWithValue("@d11", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into Quotation_Join(AccessionNo,Qty,QJ_OrderID) VALUES (@d1,@d2," & txtOrderID.Text & ")"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d2", row.Cells(4).Value)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            Dim st As String = "added the new quotation having order no.'" & txtOrderNo.Text & "'"
            LogFunc(lblUser.Text, st)
            btnSave.Enabled = False
            MessageBox.Show("Successfully saved", "Quotation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Print()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmQuotation_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillCombo()
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        Try
            If Len(Trim(cmbDepartment.Text)) = 0 Then
                MessageBox.Show("Please select department", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbDepartment.Focus()
                Exit Sub
            End If
            If Len(Trim(txtSenderName.Text)) = 0 Then
                MessageBox.Show("Please enter sender's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSenderName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtReceiverName.Text)) = 0 Then
                MessageBox.Show("Please enter receiver's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtReceiverName.Focus()
                Exit Sub
            End If

            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("sorry no data added to grid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update quotation set OrderNo=@d2, OrderDate=@d3,DepartmentID=@d4, [From]=@d5, From_Address=@d6, From_ContactNo=@d7, [To]=@d8, To_Address=@d9, To_ContactNo=@d10,  Remarks=@d11 where OrderID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtOrderNo.Text)
            cmd.Parameters.AddWithValue("@d3", dtpOrderDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtDepartmentID.Text)
            cmd.Parameters.AddWithValue("@d5", txtSenderName.Text)
            cmd.Parameters.AddWithValue("@d6", txtSenderAddress.Text)
            cmd.Parameters.AddWithValue("@d7", txtSenderContactNo.Text)
            cmd.Parameters.AddWithValue("@d8", txtReceiverName.Text)
            cmd.Parameters.AddWithValue("@d9", txtReceiverAddress.Text)
            cmd.Parameters.AddWithValue("@d10", txtReceiverCN.Text)
            cmd.Parameters.AddWithValue("@d11", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d1", txtOrderID.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "delete from Quotation_Join where QJ_OrderID=" & txtOrderID.Text & ""
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into Quotation_Join(AccessionNo,Qty,QJ_OrderID) VALUES (@d1,@d2," & txtOrderID.Text & ")"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d1", row.Cells(4).Value)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            Dim st As String = "updated the quotation having order no.'" & txtOrderNo.Text & "'"
            LogFunc(lblUser.Text, st)
            btnUpdate.Enabled = False
            MessageBox.Show("Successfully updated", "Quotation", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click_1(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Print()
    End Sub

    Private Sub txtQty_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtQty.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmBookEntryRecord.lblSet.Text = "Quotation"
        frmBookEntryRecord.ShowDialog()
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmQuotationRecord.lblSet.Text = "Quotation"
        frmQuotationRecord.Reset()
        frmQuotationRecord.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub
End Class
