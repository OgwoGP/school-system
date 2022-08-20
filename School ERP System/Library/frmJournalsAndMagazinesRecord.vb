Imports System.Data.SqlClient
Imports System.IO

Public Class frmJournalsAndMagazinesRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(JournalandMagazines.ID) as [ID], RTRIM(JM_Name) as [Title], RTRIM(SubscriptionNo) as [Subscription No], SubscriptionDate as [Subscription Date], RTRIM(Subscription) as [Subscription], SubscriptionDateFrom as [Subscription Date From], SubscriptionDateTo as [Subscription Date To], RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], Amount, PaidOn as [Paid On], RTRIM(IssueNo) as [Issue No],IssueDate as [Issue Date], RTRIM(Months) as [Month], Jm_Year as [Year], RTRIM(Volume) as [Volume], RTRIM(V_num) as [Number], DateOfReceipt as [Date Of Receipt],Supplier.ID as [SID],RTRIM(Supplier.SupplierID) as [Supplier ID], RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name], Department.ID as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM(JournalAndMagazines.Remarks) as [Remarks] from JournalAndMagazines,Supplier,Department where Supplier.ID=JournalandMagazines.SupplierID and Department.ID=JournalandMagazines.DepartmentID order by JM_name", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "JandM")
            dgw.DataSource = ds.Tables("JandM").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillDepartment()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(DepartmentName) FROM JournalAndMagazines,Department where JournalAndMagazines.DepartmentID=Department.ID", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbDepartment.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbDepartment.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub Reset()
        txtName.Text = ""
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DateTimePicker3.Text = Today
        DateTimePicker4.Text = Today
        DateTimePicker5.Text = Today
        DateTimePicker6.Text = Today
        cmbDepartment.SelectedIndex = -1
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        GetData()
        fillDepartment()
    End Sub

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "JandM Entry" Then
                    Me.Hide()
                    frmJournalsandMagazines.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    'JM_Name, SubscriptionNo, SubscriptionDate, Subscription, SubscriptionDateFrom, SubscriptionDateTo, BillNo, BillDate, Amount, PaidOn, IssueNo,IssueDate, Months, Jm_Year, Volume, V_num, DateOfReceipt, SupplierName, Department, Remarks
                    frmJournalsandMagazines.txtID.Text = dr.Cells(0).Value.ToString()
                    frmJournalsandMagazines.txtName.Text = dr.Cells(1).Value.ToString()
                    frmJournalsandMagazines.txtSubNo.Text = dr.Cells(2).Value.ToString()
                    frmJournalsandMagazines.dtpSubDate.Text = dr.Cells(3).Value.ToString()
                    frmJournalsandMagazines.txtSub.Text = dr.Cells(4).Value.ToString()
                    frmJournalsandMagazines.dtpSubDateFrom.Text = dr.Cells(5).Value.ToString()
                    frmJournalsandMagazines.dtpSubDateTo.Text = dr.Cells(6).Value.ToString()
                    frmJournalsandMagazines.txtBillNo.Text = dr.Cells(7).Value.ToString()
                    frmJournalsandMagazines.dtpBillDate.Text = dr.Cells(8).Value.ToString()
                    frmJournalsandMagazines.txtAmount.Text = dr.Cells(9).Value.ToString()
                    frmJournalsandMagazines.dtpPaidOn.Text = dr.Cells(10).Value.ToString()
                    frmJournalsandMagazines.txtIssueNo.Text = dr.Cells(11).Value.ToString()
                    frmJournalsandMagazines.dtpDate.Text = dr.Cells(12).Value.ToString()
                    frmJournalsandMagazines.cmbMonth.Text = dr.Cells(13).Value.ToString()
                    frmJournalsandMagazines.txtYear.Text = dr.Cells(14).Value.ToString()
                    frmJournalsandMagazines.txtVolume.Text = dr.Cells(15).Value.ToString()
                    frmJournalsandMagazines.txtNumber.Text = dr.Cells(16).Value.ToString()
                    frmJournalsandMagazines.dtpDateOfReceipt.Text = dr.Cells(17).Value.ToString()
                    frmJournalsandMagazines.txtS_ID.Text = dr.Cells(18).Value.ToString()
                    frmJournalsandMagazines.txtSupplierID.Text = dr.Cells(19).Value.ToString()
                    frmJournalsandMagazines.txtLastName.Text = dr.Cells(20).Value.ToString()
                    frmJournalsandMagazines.txtFirstName.Text = dr.Cells(21).Value.ToString()
                    frmJournalsandMagazines.txtDepartmentID.Text = dr.Cells(22).Value.ToString()
                    frmJournalsandMagazines.cmbDepartment.Text = dr.Cells(23).Value.ToString()
                    frmJournalsandMagazines.txtRemarks.Text = dr.Cells(24).Value.ToString()
                    frmJournalsandMagazines.btnUpdate.Enabled = True
                    frmJournalsandMagazines.btnDelete.Enabled = True
                    frmJournalsandMagazines.btnSave.Enabled = False
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

    Private Sub txtName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(JournalandMagazines.ID) as [ID], RTRIM(JM_Name) as [Title], RTRIM(SubscriptionNo) as [Subscription No], SubscriptionDate as [Subscription Date], RTRIM(Subscription) as [Subscription], SubscriptionDateFrom as [Subscription Date From], SubscriptionDateTo as [Subscription Date To], RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], Amount, PaidOn as [Paid On], RTRIM(IssueNo) as [Issue No],IssueDate as [Issue Date], RTRIM(Months) as [Month], Jm_Year as [Year], RTRIM(Volume) as [Volume], RTRIM(V_num) as [Number], DateOfReceipt as [Date Of Receipt],Supplier.ID as [SID],RTRIM(Supplier.SupplierID) as [Supplier ID], RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name], Department.ID as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM(JournalAndMagazines.Remarks) as [Remarks] from JournalAndMagazines,Supplier,Department where Supplier.ID=JournalandMagazines.SupplierID and Department.ID=JournalandMagazines.DepartmentID and JM_Name like '%" & txtName.Text & "%' order by JM_name", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "JandM")
            dgw.DataSource = ds.Tables("JandM").DefaultView
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

    Private Sub DateTimePicker1_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker1.Validating
        If (DateTimePicker2.Value.Date) > (DateTimePicker1.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DateTimePicker1.Focus()
        End If
    End Sub

    Private Sub DateTimePicker3_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker3.Validating
        If (DateTimePicker4.Value.Date) > (DateTimePicker3.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DateTimePicker3.Focus()
        End If
    End Sub

    Private Sub DateTimePicker5_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker5.Validating
        If (DateTimePicker6.Value.Date) > (DateTimePicker5.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            DateTimePicker5.Focus()
        End If
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(JournalandMagazines.ID) as [ID], RTRIM(JM_Name) as [Title], RTRIM(SubscriptionNo) as [Subscription No], SubscriptionDate as [Subscription Date], RTRIM(Subscription) as [Subscription], SubscriptionDateFrom as [Subscription Date From], SubscriptionDateTo as [Subscription Date To], RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], Amount, PaidOn as [Paid On], RTRIM(IssueNo) as [Issue No],IssueDate as [Issue Date], RTRIM(Months) as [Month], Jm_Year as [Year], RTRIM(Volume) as [Volume], RTRIM(V_num) as [Number], DateOfReceipt as [Date Of Receipt],Supplier.ID as [SID],RTRIM(Supplier.SupplierID) as [Supplier ID], RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name], Department.ID as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM(JournalAndMagazines.Remarks) as [Remarks] from JournalAndMagazines,Supplier,Department where Supplier.ID=JournalandMagazines.SupplierID and Department.ID=JournalandMagazines.DepartmentID and IssueDate between @d1 and @d2  order by IssueDate desc", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "JandM")
            dgw.DataSource = ds.Tables("JandM").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(JournalandMagazines.ID) as [ID], RTRIM(JM_Name) as [Title], RTRIM(SubscriptionNo) as [Subscription No], SubscriptionDate as [Subscription Date], RTRIM(Subscription) as [Subscription], SubscriptionDateFrom as [Subscription Date From], SubscriptionDateTo as [Subscription Date To], RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], Amount, PaidOn as [Paid On], RTRIM(IssueNo) as [Issue No],IssueDate as [Issue Date], RTRIM(Months) as [Month], Jm_Year as [Year], RTRIM(Volume) as [Volume], RTRIM(V_num) as [Number], DateOfReceipt as [Date Of Receipt],Supplier.ID as [SID],RTRIM(Supplier.SupplierID) as [Supplier ID], RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name], Department.ID as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM(JournalAndMagazines.Remarks) as [Remarks] from JournalAndMagazines,Supplier,Department where Supplier.ID=JournalandMagazines.SupplierID and Department.ID=JournalandMagazines.DepartmentID and DateOfReceipt between @d1 and @d2  order by DateOfReceipt desc", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker2.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker1.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "JandM")
            dgw.DataSource = ds.Tables("JandM").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(JournalandMagazines.ID) as [ID], RTRIM(JM_Name) as [Title], RTRIM(SubscriptionNo) as [Subscription No], SubscriptionDate as [Subscription Date], RTRIM(Subscription) as [Subscription], SubscriptionDateFrom as [Subscription Date From], SubscriptionDateTo as [Subscription Date To], RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], Amount, PaidOn as [Paid On], RTRIM(IssueNo) as [Issue No],IssueDate as [Issue Date], RTRIM(Months) as [Month], Jm_Year as [Year], RTRIM(Volume) as [Volume], RTRIM(V_num) as [Number], DateOfReceipt as [Date Of Receipt],Supplier.ID as [SID],RTRIM(Supplier.SupplierID) as [Supplier ID], RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name], Department.ID as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM(JournalAndMagazines.Remarks) as [Remarks] from JournalAndMagazines,Supplier,Department where Supplier.ID=JournalandMagazines.SupplierID and Department.ID=JournalandMagazines.DepartmentID and BillDate between @d1 and @d2 order by BillDate desc", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker4.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker3.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "JandM")
            dgw.DataSource = ds.Tables("JandM").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(JournalandMagazines.ID) as [ID], RTRIM(JM_Name) as [Title], RTRIM(SubscriptionNo) as [Subscription No], SubscriptionDate as [Subscription Date], RTRIM(Subscription) as [Subscription], SubscriptionDateFrom as [Subscription Date From], SubscriptionDateTo as [Subscription Date To], RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], Amount, PaidOn as [Paid On], RTRIM(IssueNo) as [Issue No],IssueDate as [Issue Date], RTRIM(Months) as [Month], Jm_Year as [Year], RTRIM(Volume) as [Volume], RTRIM(V_num) as [Number], DateOfReceipt as [Date Of Receipt],Supplier.ID as [SID],RTRIM(Supplier.SupplierID) as [Supplier ID], RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name], Department.ID as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM(JournalAndMagazines.Remarks) as [Remarks] from JournalAndMagazines,Supplier,Department where Supplier.ID=JournalandMagazines.SupplierID and Department.ID=JournalandMagazines.DepartmentID and SubscriptionDateTo between @d1 and @d2 order by SubscriptionDateTo", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker6.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker5.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "JandM")
            dgw.DataSource = ds.Tables("JandM").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbDepartment_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbDepartment.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(JournalandMagazines.ID) as [ID], RTRIM(JM_Name) as [Title], RTRIM(SubscriptionNo) as [Subscription No], SubscriptionDate as [Subscription Date], RTRIM(Subscription) as [Subscription], SubscriptionDateFrom as [Subscription Date From], SubscriptionDateTo as [Subscription Date To], RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], Amount, PaidOn as [Paid On], RTRIM(IssueNo) as [Issue No],IssueDate as [Issue Date], RTRIM(Months) as [Month], Jm_Year as [Year], RTRIM(Volume) as [Volume], RTRIM(V_num) as [Number], DateOfReceipt as [Date Of Receipt],Supplier.ID as [SID],RTRIM(Supplier.SupplierID) as [Supplier ID], RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name], Department.ID as [Department ID],RTRIM(DepartmentName) as [Department], RTRIM(JournalAndMagazines.Remarks) as [Remarks] from JournalAndMagazines,Supplier,Department where Supplier.ID=JournalandMagazines.SupplierID and Department.ID=JournalandMagazines.DepartmentID and DepartmentName=@d1 order by JM_name", con)
            cmd.Parameters.AddWithValue("@d1", cmbDepartment.Text)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "JandM")
            dgw.DataSource = ds.Tables("JandM").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
