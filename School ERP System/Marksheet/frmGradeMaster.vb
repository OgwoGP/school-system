Imports System.Data.SqlClient
Public Class frmGradeMaster


    Sub Reset()
        txtMarksStart.Text = ""
        txtGrade.Text = ""
        txtGradePoint.Text = ""
        txtMarksEnd.Text = ""
        txtGrade.Focus()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        Getdata()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtGrade.Text)) = 0 Then
            MessageBox.Show("Please enter grade", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtGrade.Focus()
            Exit Sub
        End If
        If Len(Trim(txtGradePoint.Text)) = 0 Then
            MessageBox.Show("Please enter grade point", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtGradePoint.Focus()
            Exit Sub
        End If
        If Len(Trim(txtMarksStart.Text)) = 0 And Len(Trim(txtMarksEnd.Text)) = 0 Then
            MessageBox.Show("Please enter marks", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMarksStart.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select Grade from GradeMaster where Grade=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Grade Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into GradeMaster(Grade,GradePoint,MarksStartFrom,MarksEndTo) VALUES (@d1,@d2,@d3,@d4)"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            cmd.Parameters.AddWithValue("@d2", Val(txtGradePoint.Text))
            cmd.Parameters.AddWithValue("@d3", Val(txtMarksStart.Text))
            cmd.Parameters.AddWithValue("@d4", Val(txtMarksEnd.Text))
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "added the new Grade '" & txtGrade.Text & "'")
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Getdata()
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
    Private Sub DeleteRecord()

        Try

            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = "select GradeMaster.Grade from GradeMaster,MarksEntry_Join where GradeMaster.Grade=MarksEntry_Join.OGTheory and GradeMaster.Grade=@d1"
            cmd = New SqlCommand(sql)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Marks Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim sql1 As String = "select GradeMaster.Grade from GradeMaster,MarksEntry_Join where GradeMaster.Grade=MarksEntry_Join.OGPractical and GradeMaster.Grade=@d1"
            cmd = New SqlCommand(sql1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Marks Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim sql2 As String = "select GradeMaster.Grade from GradeMaster,MarksEntry_Join where GradeMaster.Grade=MarksEntry_Join.FinalGrade and GradeMaster.Grade=@d1"
            cmd = New SqlCommand(sql2)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Marks Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from GradeMaster where Grade=@d1"
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the Grade '" & txtGrade.Text & "'")
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

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtGrade.Text)) = 0 Then
            MessageBox.Show("Please enter grade", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtGrade.Focus()
            Exit Sub
        End If
        If Len(Trim(txtGradePoint.Text)) = 0 Then
            MessageBox.Show("Please enter grade point", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtGradePoint.Focus()
            Exit Sub
        End If
        If Len(Trim(txtMarksStart.Text)) = 0 And Len(Trim(txtMarksEnd.Text)) = 0 Then
            MessageBox.Show("Please enter marks", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMarksStart.Focus()
            Exit Sub
        End If
        Try
            If txtGrade.Text <> txtGrade1.Text Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ct As String = "select Grade from GradeMaster where Grade=@d1"
                cmd = New SqlCommand(ct)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    MessageBox.Show("Grade Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    Return
                End If
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "update MarksEntry_Join set OGTheory=@d1 where OGTheory=@d2"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            cmd.Parameters.AddWithValue("@d2", txtGrade1.Text)
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb2 As String = "update MarksEntry_Join set OGPractical=@d1 where OGPractical=@d2"
            cmd = New SqlCommand(cb2)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            cmd.Parameters.AddWithValue("@d2", txtGrade1.Text)
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb3 As String = "update MarksEntry_Join set FinalGrade=@d1 where FinalGrade=@d2"
            cmd = New SqlCommand(cb3)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            cmd.Parameters.AddWithValue("@d2", txtGrade1.Text)
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update GradeMaster set Grade=@d1,GradePoint=@d2,MarksStartFrom=@d3,MarksEndTo=@d4 where Grade=@d5"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtGrade.Text)
            cmd.Parameters.AddWithValue("@d2", Val(txtGradePoint.Text))
            cmd.Parameters.AddWithValue("@d3", Val(txtMarksStart.Text))
            cmd.Parameters.AddWithValue("@d4", Val(txtMarksEnd.Text))
            cmd.Parameters.AddWithValue("@d5", txtGrade1.Text)
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "Updated the grade '" & txtGrade.Text & "' details")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            Getdata()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(Grade),(GradePoint),(MarksStartFrom),MarksEndTo from GradeMaster order by GradePoint desc", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
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

    Private Sub frmCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Getdata()
    End Sub

    Private Sub dgw_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                txtGrade.Text = dr.Cells(0).Value.ToString()
                txtGrade1.Text = dr.Cells(0).Value.ToString()
                txtGradePoint.Text = dr.Cells(1).Value.ToString()
                txtMarksStart.Text = dr.Cells(2).Value.ToString()
                txtMarksEnd.Text = dr.Cells(3).Value.ToString()
                btnUpdate.Enabled = True
                btnDelete.Enabled = True
                btnSave.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbCourse_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs)
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub
    Private Sub txtGradePoint_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtGradePoint.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtGradePoint.Text
            Dim selectionStart = Me.txtGradePoint.SelectionStart
            Dim selectionLength = Me.txtGradePoint.SelectionLength

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

    Private Sub txtMarksStart_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMarksStart.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMarksEnd_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMarksEnd.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
End Class
