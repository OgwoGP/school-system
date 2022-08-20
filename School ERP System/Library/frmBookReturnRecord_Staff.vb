Imports System.Data.SqlClient
Imports System.IO

Public Class frmBookReturnRecord_Staff
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(Return_Staff.ID) as [ID],IssueID as [Issue ID],RTRIM(Book.AccessionNo) as [Accession No],RTRIM(BookTitle) as [Book Title],RTRIM(Author) as [Author],RTRIM(JointAuthors) as [Joint Authors],ST_ID as [SID],RTRIM(Staff.StaffID) as [StaffID],RTRIM(Staffname) as [Staff Name],IssueDate as [Date],DueDate as [Due Date],ReturnDate as [Return date],Fine as [Fine], RTRIM(Return_Staff.Remarks) as [Remarks] from Book,BookIssue_Staff,Staff,Return_Staff where Book.AccessionNo=BookIssue_Staff.AccessionNo and BookIssue_Staff.StaffID=Staff.St_ID and return_Staff.IssueID=BookIssue_Staff.ID order by ReturnDate desc", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Return_Staff")
            dgw.DataSource = ds.Tables("Return_Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub Reset()
        txtStaffName.Text = ""
        dtpDateTo.Value = Today
        dtpdateFrom.Value = Today
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Book Return" Then
                    Me.Hide()
                    frmBookReturn.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    'JM_Name, SubscriptionNo, SubscriptionDate, Subscription, SubscriptionDateFrom, SubscriptionDateTo, BillNo, BillDate, Amount, PaidOn, IssueNo,IssueDate, Months, Jm_Year, Volume, V_num, DateOfReceipt, SupplierName, Department, Remarks
                    frmBookReturn.txtID1.Text = dr.Cells(0).Value.ToString()
                    frmBookReturn.txtIssueID_Staff.Text = dr.Cells(1).Value.ToString()
                    frmBookReturn.txtAccessionNo1.Text = dr.Cells(2).Value.ToString()
                    frmBookReturn.txtBookTitle1.Text = dr.Cells(3).Value.ToString()
                    frmBookReturn.txtAuthor1.Text = dr.Cells(4).Value.ToString()
                    frmBookReturn.txtJointAuthors1.Text = dr.Cells(5).Value.ToString()
                    frmBookReturn.txtS_ID.Text = dr.Cells(6).Value.ToString()
                    frmBookReturn.txtStaffID.Text = dr.Cells(7).Value.ToString()
                    frmBookReturn.txtStaffName.Text = dr.Cells(8).Value.ToString()
                    frmBookReturn.dtpIssueDate1.Text = dr.Cells(9).Value.ToString()
                    frmBookReturn.dtpDueDate1.Text = dr.Cells(10).Value.ToString()
                    frmBookReturn.dtpReturnDate1.Text = dr.Cells(11).Value.ToString()
                    frmBookReturn.txtFine1.Text = dr.Cells(12).Value.ToString()
                    frmBookReturn.txtRemarks1.Text = dr.Cells(13).Value.ToString()
                    frmBookReturn.btnUpdate1.Enabled = True
                    frmBookReturn.btnDelete1.Enabled = True
                    frmBookReturn.btnSave1.Enabled = False
                    frmBookReturn.Button4.Enabled = False
                    frmBookReturn.FillData()
                    Me.lblSet.Text = ""
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgw_RowPostPaint(sender As Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub


    Private Sub txtName_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub txtStudentName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(Return_Staff.ID) as [ID],IssueID as [Issue ID],RTRIM(Book.AccessionNo) as [Accession No],RTRIM(BookTitle) as [Book Title],RTRIM(Author) as [Author],RTRIM(JointAuthors) as [Joint Authors],ST_ID as [SID],RTRIM(Staff.StaffID) as [StaffID],RTRIM(Staffname) as [Staff Name],IssueDate as [Date],DueDate as [Due Date],ReturnDate as [Return date],Fine as [Fine], RTRIM(Return_Staff.Remarks) as [Remarks] from Book,BookIssue_Staff,Staff,Return_Staff where Book.AccessionNo=BookIssue_Staff.AccessionNo and BookIssue_Staff.StaffID=Staff.St_ID and return_Staff.IssueID=BookIssue_Staff.ID and StaffName like '%" & txtStaffName.Text & "%' order by ReturnDate desc", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "return_Staff")
            dgw.DataSource = ds.Tables("return_Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(Return_Staff.ID) as [ID],IssueID as [Issue ID],RTRIM(Book.AccessionNo) as [Accession No],RTRIM(BookTitle) as [Book Title],RTRIM(Author) as [Author],RTRIM(JointAuthors) as [Joint Authors],ST_ID as [SID],RTRIM(Staff.StaffID) as [StaffID],RTRIM(Staffname) as [Staff Name],IssueDate as [Date],DueDate as [Due Date],ReturnDate as [Return date],Fine as [Fine], RTRIM(Return_Staff.Remarks) as [Remarks] from Book,BookIssue_Staff,Staff,Return_Staff where Book.AccessionNo=BookIssue_Staff.AccessionNo and BookIssue_Staff.StaffID=Staff.St_ID and return_Staff.IssueID=BookIssue_Staff.ID and ReturnDate between @d1 and @d2 order by ReturnDate desc", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpdateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "return_Staff")
            dgw.DataSource = ds.Tables("return_Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(Return_Staff.ID) as [ID],IssueID as [Issue ID],RTRIM(Book.AccessionNo) as [Accession No],RTRIM(BookTitle) as [Book Title],RTRIM(Author) as [Author],RTRIM(JointAuthors) as [Joint Authors],ST_ID as [SID],RTRIM(Staff.StaffID) as [StaffID],RTRIM(Staffname) as [Staff Name],IssueDate as [Date],DueDate as [Due Date],ReturnDate as [Return date],Fine as [Fine], RTRIM(Return_Staff.Remarks) as [Remarks] from Book,BookIssue_Staff,Staff,Return_Staff where Book.AccessionNo=BookIssue_Staff.AccessionNo and BookIssue_Staff.StaffID=Staff.St_ID and return_Staff.IssueID=BookIssue_Staff.ID and ReturnDate between @d1 and @d2 and Fine > 0 order by ReturnDate", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker2.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker1.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "return_Staff")
            dgw.DataSource = ds.Tables("return_Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtpDateTo_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles dtpDateTo.Validating
        If (dtpdateFrom.Value.Date) > (dtpDateTo.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtpDateTo.Focus()
        End If
    End Sub

    Private Sub DateTimePicker1_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker1.Validating
        If (DateTimePicker2.Value.Date) > (DateTimePicker1.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DateTimePicker1.Focus()
        End If
    End Sub
End Class
