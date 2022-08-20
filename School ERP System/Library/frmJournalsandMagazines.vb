Imports System.Data.SqlClient
Imports System.IO

Public Class frmJournalsandMagazines

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmSupplierRecord.lblSet.Text = "JandM Entry"
        frmSupplierRecord.Getdata()
        frmSupplierRecord.ShowDialog()
    End Sub
    Sub fillDepartment()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Departmentname) FROM Department", con)
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
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from JournalAndMagazines where ID = " & txtID.Text & ""
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the Journal or magazine '" & txtName.Text & "' has ID= " & txtID.Text & "")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Autocomplete()
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub
    Sub Autocomplete()
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cmd As New SqlCommand("SELECT distinct JM_name FROM JournalAndMagazines", con)
            Dim ds As New DataSet()
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(ds, "JournalAndMagazines")
            Dim col As New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("JM_name").ToString())
            Next
            txtName.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtName.AutoCompleteCustomSource = col
            txtName.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM JournalandMagazines")
            cmd = New SqlCommand(sql)
            cmd.Connection = con
            If (IsDBNull(cmd.ExecuteScalar)) Then
                Num = 1
                txtID.Text = Num.ToString
            Else
                Num = cmd.ExecuteScalar + 1
                txtID.Text = Num.ToString
            End If
            cmd.Dispose()
            con.Close()
            con.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmStaff1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Autocomplete()
        fillDepartment()
    End Sub
    Sub Reset()
        dtpBillDate.Text = Today
        dtpPaidOn.Text = Today
        dtpSubDate.Text = Today
        dtpSubDateFrom.Value = Today
        dtpSubDateTo.Value = Today
        txtIssueNo.Text = ""
        txtVolume.Text = ""
        txtRemarks.Text = ""
        txtYear.Text = ""
        cmbMonth.SelectedIndex = -1
        txtName.Text = ""
        txtNumber.Text = ""
        txtSub.Text = ""
        txtSubNo.Text = ""
        txtAmount.Text = ""
        txtBillNo.Text = ""
        cmbDepartment.SelectedIndex = -1
        dtpDate.Text = Today
        txtSupplierID.Text = ""
        txtFirstName.Text = ""
        txtLastName.Text = ""
        dtpDateOfReceipt.Text = Today
        txtName.Focus()
        auto()
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Try
            If Len(Trim(txtName.Text)) = 0 Then
                MessageBox.Show("Please Enter Title", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtBillNo.Text)) = 0 Then
                MessageBox.Show("Please Enter Bill No.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtBillNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtAmount.Text)) = 0 Then
                MessageBox.Show("Please Enter Amount", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbMonth.Text)) = 0 Then
                MessageBox.Show("Please select month", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbMonth.Focus()
                Exit Sub
            End If
            If Len(Trim(txtYear.Text)) = 0 Then
                MessageBox.Show("Please Enter Year", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtYear.Focus()
                Exit Sub
            End If
            If Len(Trim(txtSupplierID.Text)) = 0 Then
                MessageBox.Show("Please retrieve supplier info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSupplierID.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbDepartment.Text)) = 0 Then
                MessageBox.Show("Please select department", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbDepartment.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into JournalAndMagazines(ID, JM_Name, SubscriptionNo, SubscriptionDate, Subscription, SubscriptionDateFrom, SubscriptionDateTo, BillNo, BillDate, Amount, PaidOn, IssueNo,IssueDate, Months, Jm_Year, Volume, V_num, DateOfReceipt, SupplierID, DepartmentID, Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17,@d18,@d19,@d20,@d21)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtName.Text)
            cmd.Parameters.AddWithValue("@d3", txtSubNo.Text)
            cmd.Parameters.AddWithValue("@d4", dtpSubDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", txtSub.Text)
            cmd.Parameters.AddWithValue("@d6", dtpSubDateFrom.Value.Date)
            cmd.Parameters.AddWithValue("@d7", dtpSubDateTo.Value.Date)
            cmd.Parameters.AddWithValue("@d8", txtBillNo.Text)
            cmd.Parameters.AddWithValue("@d9", dtpBillDate.Value.Date)
            cmd.Parameters.AddWithValue("@d10", txtAmount.Text)
            cmd.Parameters.AddWithValue("@d11", dtpPaidOn.Value.Date)
            cmd.Parameters.AddWithValue("@d12", txtIssueNo.Text)
            cmd.Parameters.AddWithValue("@d13", dtpDate.Value.Date)
            cmd.Parameters.AddWithValue("@d14", cmbMonth.Text)
            cmd.Parameters.AddWithValue("@d15", txtYear.Text)
            cmd.Parameters.AddWithValue("@d16", txtVolume.Text)
            cmd.Parameters.AddWithValue("@d17", txtNumber.Text)
            cmd.Parameters.AddWithValue("@d18", dtpDateOfReceipt.Value.Date)
            cmd.Parameters.AddWithValue("@d19", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d20", txtDepartmentID.Text)
            cmd.Parameters.AddWithValue("@d21", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            LogFunc(lblUser.Text, "added the new Journal or magazine '" & txtName.Text & "' has ID= " & txtID.Text & "")
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Autocomplete()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtYear_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtYear.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtAmount_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtAmount.Text
            Dim selectionStart = Me.txtAmount.SelectionStart
            Dim selectionLength = Me.txtAmount.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        Try
            If Len(Trim(txtName.Text)) = 0 Then
                MessageBox.Show("Please Enter Title", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtBillNo.Text)) = 0 Then
                MessageBox.Show("Please Enter Bill No.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtBillNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtAmount.Text)) = 0 Then
                MessageBox.Show("Please Enter Amount", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbMonth.Text)) = 0 Then
                MessageBox.Show("Please select month", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbMonth.Focus()
                Exit Sub
            End If
            If Len(Trim(txtYear.Text)) = 0 Then
                MessageBox.Show("Please Enter Year", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtYear.Focus()
                Exit Sub
            End If
            If Len(Trim(txtSupplierID.Text)) = 0 Then
                MessageBox.Show("Please retrieve supplier info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSupplierID.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbDepartment.Text)) = 0 Then
                MessageBox.Show("Please select department", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbDepartment.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update JournalAndMagazines set JM_Name=@d2, SubscriptionNo=@d3, SubscriptionDate=@d4, Subscription=@d5, SubscriptionDateFrom=@d6, SubscriptionDateTo=@d7, BillNo=@d8, BillDate=@d9, Amount=@d10, PaidOn=@d11, IssueNo=@d12,IssueDate=@d13, Months=@d14, Jm_Year=@d15, Volume=@d16, V_num=@d17, DateOfReceipt=@d18, SupplierID=@d19, DepartmentID=@d20, Remarks=@d21 where ID=" & txtID.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtName.Text)
            cmd.Parameters.AddWithValue("@d3", txtSubNo.Text)
            cmd.Parameters.AddWithValue("@d4", dtpSubDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", txtSub.Text)
            cmd.Parameters.AddWithValue("@d6", dtpSubDateFrom.Value.Date)
            cmd.Parameters.AddWithValue("@d7", dtpSubDateTo.Value.Date)
            cmd.Parameters.AddWithValue("@d8", txtBillNo.Text)
            cmd.Parameters.AddWithValue("@d9", dtpBillDate.Value.Date)
            cmd.Parameters.AddWithValue("@d10", txtAmount.Text)
            cmd.Parameters.AddWithValue("@d11", dtpPaidOn.Value.Date)
            cmd.Parameters.AddWithValue("@d12", txtIssueNo.Text)
            cmd.Parameters.AddWithValue("@d13", dtpDate.Value.Date)
            cmd.Parameters.AddWithValue("@d14", cmbMonth.Text)
            cmd.Parameters.AddWithValue("@d15", txtYear.Text)
            cmd.Parameters.AddWithValue("@d16", txtVolume.Text)
            cmd.Parameters.AddWithValue("@d17", txtNumber.Text)
            cmd.Parameters.AddWithValue("@d18", dtpDateOfReceipt.Value.Date)
            cmd.Parameters.AddWithValue("@d19", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d20", txtDepartmentID.Text)
            cmd.Parameters.AddWithValue("@d21", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            LogFunc(lblUser.Text, "update the Journal or magazine '" & txtName.Text & "' has ID= " & txtID.Text & "")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            Autocomplete()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmJournalsAndMagazinesRecord.Reset()
        frmJournalsAndMagazinesRecord.lblSet.Text = "JandM Entry"
        frmJournalsAndMagazinesRecord.ShowDialog()
    End Sub
End Class
