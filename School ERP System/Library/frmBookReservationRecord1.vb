Imports System.Data.SqlClient
Imports System.IO

Public Class frmBookReservationRecord1
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(BookReservation.ID) as [Reservation ID],RTRIM(Book.AccessionNo) as [Accession No],RTRIM(BookTitle) as [Book Title],RTRIM(Author) as [Author],RTRIM(JointAuthors) as [Joint Authors],ST_ID as [SID],RTRIM(Staff.StaffID) as [StaffID],RTRIM(Staffname) as [Staff Name],R_Date as [Date], RTRIM(BookReservation.Status) as [Status], RTRIM(BookReservation.Remarks) as [Remarks] from Book,BookReservation,Staff where Book.AccessionNo=BookReservation.AccessionNo and BookReservation.StaffID=Staff.St_ID and BookReservation.Status='Reserved' order by R_Date", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub


    Sub Reset()
        txtStaffName.Text = ""
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Book Reservation Entry" Then
                    Me.Hide()
                    frmBookReservation.Show()
                    frmBookReservation.txtID.Text = dr.Cells(0).Value.ToString()
                    frmBookReservation.txtAccessionNo.Text = dr.Cells(1).Value.ToString()
                    frmBookReservation.txtBookTitle.Text = dr.Cells(2).Value.ToString()
                    frmBookReservation.txtAuthor.Text = dr.Cells(3).Value.ToString()
                    frmBookReservation.txtJointAuthor.Text = dr.Cells(4).Value.ToString()
                    frmBookReservation.txtS_ID.Text = dr.Cells(5).Value.ToString()
                    frmBookReservation.txtStaffID.Text = dr.Cells(6).Value.ToString()
                    frmBookReservation.txtStaffName.Text = dr.Cells(7).Value.ToString()
                    frmBookReservation.dtpReservationDate.Text = dr.Cells(8).Value.ToString()
                    frmBookReservation.txtStatus.Text = dr.Cells(9).Value.ToString()
                    frmBookReservation.txtRemarks.Text = dr.Cells(10).Value.ToString()
                    frmBookReservation.btnSave.Enabled = False
                    frmBookReservation.btnUpdate.Enabled = True
                    frmBookReservation.btnDelete.Enabled = True
                    frmBookReservation.btnCancelReservation.Enabled = True
                    frmBookReservation.Button1.Enabled = False
                    frmBookReservation.Button2.Enabled = False
                    Me.lblSet.Text = ""
                End If
                If lblSet.Text = "Book Issue" Then
                    Me.Hide()
                    frmBookIssue.Show()
                    frmBookIssue.txtReservationID.Text = dr.Cells(0).Value.ToString()
                    frmBookIssue.txtAccessionNo1.Text = dr.Cells(1).Value.ToString()
                    frmBookIssue.txtBookTitle1.Text = dr.Cells(2).Value.ToString()
                    frmBookIssue.txtAuthor1.Text = dr.Cells(3).Value.ToString()
                    frmBookIssue.txtJointAuthors1.Text = dr.Cells(4).Value.ToString()
                    frmBookIssue.txtS_ID.Text = dr.Cells(5).Value.ToString()
                    frmBookIssue.txtStaffID.Text = dr.Cells(6).Value.ToString()
                    frmBookIssue.txtStaffName.Text = dr.Cells(7).Value.ToString()
                    frmBookIssue.btnSave1.Enabled = False
                    frmBookIssue.btnIssue_Reservation.Enabled = True
                    frmBookIssue.Button3.Enabled = False
                    frmBookIssue.FillData()
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

    Private Sub txtStaffName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(BookReservation.ID) as [Reservation ID],RTRIM(Book.AccessionNo) as [Accession No],RTRIM(BookTitle) as [Book Title],RTRIM(Author) as [Author],RTRIM(JointAuthors) as [Joint Authors],ST_ID as [SID],RTRIM(Staff.StaffID) as [StaffID],RTRIM(Staffname) as [Staff Name],R_Date as [Date], RTRIM(BookReservation.Status) as [Status], RTRIM(BookReservation.Remarks) as [Remarks] from Book,BookReservation,Staff where Book.AccessionNo=BookReservation.AccessionNo and BookReservation.StaffID=Staff.St_ID and BookReservation.Status='Reserved' and Staffname like '%" & txtStaffName.Text & "%' order by R_Date", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
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

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(BookReservation.ID) as [Reservation ID],RTRIM(Book.AccessionNo) as [Accession No],RTRIM(BookTitle) as [Book Title],RTRIM(Author) as [Author],RTRIM(JointAuthors) as [Joint Authors],ST_ID as [SID],RTRIM(Staff.StaffID) as [StaffID],RTRIM(Staffname) as [Staff Name],R_Date as [Date], RTRIM(BookReservation.Status) as [Status], RTRIM(BookReservation.Remarks) as [Remarks] from Book,BookReservation,Staff where Book.AccessionNo=BookReservation.AccessionNo and BookReservation.StaffID=Staff.St_ID and BookReservation.Status='Reserved' and R_Date between @d1 and @d2 order by R_Date", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub
End Class
