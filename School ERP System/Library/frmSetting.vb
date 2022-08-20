Imports System.Data.SqlClient
Public Class frmSetting

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Sub Reset()
        txtMaxDaysAllowedStaff.Text = ""
        txtFinePerDayStaff.Text = ""
        txtFinePerDayStudent.Text = ""
        txtMaxDaysAllowedStudent.Text = ""
        txtMaxBooksAllowedStaff.Text = ""
        txtMaxBooksAllowedStudent.Text = ""
        cmbBookType.SelectedIndex = -1
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        txtMaxDaysAllowedStaff.Focus()
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If cmbBookType.Text = "" Then
            MessageBox.Show("Please select book type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbBookType.Focus()
            Return
        End If
        If txtMaxDaysAllowedStaff.Text = "" Then
            MessageBox.Show("Please enter max. days allowed (Staff)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxDaysAllowedStaff.Focus()
            Return
        End If
        If txtMaxDaysAllowedStudent.Text = "" Then
            MessageBox.Show("Please enter max. days allowed (Student)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxDaysAllowedStudent.Focus()
            Return
        End If
        If txtFinePerDayStaff.Text = "" Then
            MessageBox.Show("Please enter fine per day (Staff)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFinePerDayStaff.Focus()
            Return
        End If
        If txtFinePerDayStudent.Text = "" Then
            MessageBox.Show("Please enter fine per day (Student)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFinePerDayStudent.Focus()
            Return
        End If
        If txtMaxBooksAllowedStaff.Text = "" Then
            MessageBox.Show("Please enter max. books allowed (Staff)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxBooksAllowedStaff.Focus()
            Return
        End If
        If txtMaxBooksAllowedStudent.Text = "" Then
            MessageBox.Show("Please enter max. books allowed (Student)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxBooksAllowedStudent.Focus()
            Return
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select BookType from Setting where BookType=@d1"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", cmbBookType.Text)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If rdr.Read() Then
                MessageBox.Show("Setting for selected book type Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into Setting(BookType, MaxDays_Staff, MaxDays_Student, FinePerDay_Student, FinePerDay_Staff,MaxBooks_Staff,MaxBooks_Student) VALUES ('" & cmbBookType.Text & "'," & txtMaxDaysAllowedStaff.Text & "," & txtMaxDaysAllowedStudent.Text & "," & txtFinePerDayStudent.Text & "," & txtFinePerDayStaff.Text & "," & txtMaxBooksAllowedStaff.Text & "," & txtMaxBooksAllowedStudent.Text & ")"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "added the setting for Book Type '" & cmbBookType.Text & "' for book issue and return"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully Saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Getdata()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If cmbBookType.Text = "" Then
            MessageBox.Show("Please select book type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbBookType.Focus()
            Return
        End If
        If txtMaxDaysAllowedStaff.Text = "" Then
            MessageBox.Show("Please enter max. days allowed (Staff)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxDaysAllowedStaff.Focus()
            Return
        End If
        If txtMaxDaysAllowedStudent.Text = "" Then
            MessageBox.Show("Please enter max. days allowed (Student)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxDaysAllowedStudent.Focus()
            Return
        End If
        If txtFinePerDayStaff.Text = "" Then
            MessageBox.Show("Please enter fine per day (Staff)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFinePerDayStaff.Focus()
            Return
        End If
        If txtFinePerDayStudent.Text = "" Then
            MessageBox.Show("Please enter fine per day (Student)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFinePerDayStudent.Focus()
            Return
        End If
        If txtMaxBooksAllowedStaff.Text = "" Then
            MessageBox.Show("Please enter max. books allowed (Staff)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxBooksAllowedStaff.Focus()
            Return
        End If
        If txtMaxBooksAllowedStudent.Text = "" Then
            MessageBox.Show("Please enter max. books allowed (Student)", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaxBooksAllowedStudent.Focus()
            Return
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Setting set MaxDays_Staff=" & txtMaxDaysAllowedStaff.Text & ", MaxDays_Student=" & txtMaxDaysAllowedStudent.Text & ", FinePerDay_Student=" & txtFinePerDayStudent.Text & ", FinePerDay_Staff=" & txtFinePerDayStaff.Text & ",MaxBooks_Staff=" & txtMaxBooksAllowedStaff.Text & ",MaxBooks_Student=" & txtMaxBooksAllowedStudent.Text & " where BookType='" & cmbBookType.Text & "'"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "updated the setting for Book Type '" & cmbBookType.Text & "' for book issue and return"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            Getdata()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub DeleteRecord()

        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Setting where BookType=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", cmbBookType.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the setting for Book Type '" & cmbBookType.Text & "' for book issue and return"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Getdata()
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

    Private Sub dgw_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            Dim dr As DataGridViewRow = dgw.SelectedRows(0)
            cmbBookType.Text = dr.Cells(0).Value.ToString()
            txtMaxDaysAllowedStaff.Text = dr.Cells(1).Value.ToString()
            txtMaxDaysAllowedStudent.Text = dr.Cells(2).Value.ToString()
            txtFinePerDayStaff.Text = dr.Cells(3).Value.ToString()
            txtFinePerDayStudent.Text = dr.Cells(4).Value.ToString()
            txtMaxBooksAllowedStaff.Text = dr.Cells(5).Value.ToString()
            txtMaxBooksAllowedStudent.Text = dr.Cells(6).Value.ToString()
            btnUpdate.Enabled = True
            btnDelete.Enabled = True
            btnSave.Enabled = False
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
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(BookType), MaxDays_Staff, MaxDays_Student, FinePerDay_Staff, FinePerDay_Student,MaxBooks_Staff,MaxBooks_Student from setting order by BookType", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmDesignation_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Getdata()
    End Sub

    Private Sub txtFinePerDayStaff_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtFinePerDayStaff.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtFinePerDayStaff.Text
            Dim selectionStart = Me.txtFinePerDayStaff.SelectionStart
            Dim selectionLength = Me.txtFinePerDayStaff.SelectionLength

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

    Private Sub txtFinePerDayStudent_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtFinePerDayStudent.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtFinePerDayStudent.Text
            Dim selectionStart = Me.txtFinePerDayStudent.SelectionStart
            Dim selectionLength = Me.txtFinePerDayStudent.SelectionLength

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

    Private Sub txtMaxDaysAllowedStaff_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxDaysAllowedStaff.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMaxDaysAllowedStudent_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxDaysAllowedStudent.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMaxBooksAllowedStaff_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxBooksAllowedStaff.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMaxBooksAllowedStudent_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxBooksAllowedStudent.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
End Class
