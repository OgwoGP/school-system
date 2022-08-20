Imports System.Data.SqlClient
Public Class frmNewspaper
    Dim Status As String
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM Newspaper")
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
    Sub fillCombo()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (NewsPaper) FROM Newspaper_Master", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbNewsPaper.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbNewsPaper.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        cmbNewsPaper.SelectedIndex = -1
        txtRemarks.Text = ""
        dtpDate.Text = Today
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        auto()
        cmbNewsPaper.Focus()
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(cmbNewsPaper.Text)) = 0 Then
            MessageBox.Show("Please select Paper Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbNewsPaper.Focus()
            Exit Sub
        End If
        If (RadioButton1.Checked = False And RadioButton2.Checked = False) Then
            MessageBox.Show("Please select status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If (RadioButton1.Checked = True) Then
            Status = RadioButton1.Text
        End If
        If (RadioButton2.Checked = True) Then
            Status = RadioButton2.Text
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select Papername,N_date from NewsPaper where Papername=@d1 and N_Date=@d2"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", cmbNewsPaper.Text)
            cmd.Parameters.AddWithValue("@d2", dtpDate.Value.Date)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Record Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into Newspaper(ID,PaperName, N_Date, Status, Remarks) VALUES (@d1,@d2,@d3,@d4,@d5)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", cmbNewsPaper.Text)
            cmd.Parameters.AddWithValue("@d3", dtpDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", Status)
            cmd.Parameters.AddWithValue("@d5", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "added the new entry for newspaper '" & cmbNewsPaper.Text & "' having ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully Saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(cmbNewsPaper.Text)) = 0 Then
            MessageBox.Show("Please select Paper Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbNewsPaper.Focus()
            Exit Sub
        End If
        If (RadioButton1.Checked = False And RadioButton2.Checked = False) Then
            MessageBox.Show("Please select status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If (RadioButton1.Checked = True) Then
            Status = RadioButton1.Text
        End If
        If (RadioButton2.Checked = True) Then
            Status = RadioButton2.Text
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Newspaper set PaperName=@d2, N_Date=@d3, Status=@d4, Remarks=@d5 where ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", cmbNewsPaper.Text)
            cmd.Parameters.AddWithValue("@d3", dtpDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", Status)
            cmd.Parameters.AddWithValue("@d5", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "updated the entry for newspaper '" & cmbNewsPaper.Text & "' having ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub DeleteRecord()

        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Newspaper where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the entry for newspaper '" & cmbNewsPaper.Text & "' having ID='" & txtID.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            Else
                MessageBox.Show("No Record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmNewspaper_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillCombo()
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmNewspaperRecord.lblSet.Text = "Newspaper Entry"
        frmNewspaperRecord.Reset()
        frmNewspaperRecord.ShowDialog()
    End Sub
End Class
