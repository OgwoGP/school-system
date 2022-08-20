Imports System.Data.SqlClient
Public Class frmSubjectEntry

    Sub fillCourse()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(ClassName) FROM Class", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbClass.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbClass.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub Reset()
        cmbClass.SelectedIndex = -1
        txtSubjectCode.Text = ""
        txtSearchByClass.Text = ""
        txtSubjectName.Text = ""
        txtSearchBySubject.Text = ""
        txtSubjectCode.Focus()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        Getdata()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtSubjectCode.Text)) = 0 Then
            MessageBox.Show("Please enter subject code", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSubjectCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtSubjectName.Text)) = 0 Then
            MessageBox.Show("Please enter subject name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSubjectName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbClass.Text)) = 0 Then
            MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbClass.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select SubjectCode from Subject where SubjectCode=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Subject Code Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Subject(SubjectCode,SubjectName,Class) VALUES (@d1,@d2,@d3)"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
            cmd.Parameters.AddWithValue("@d2", txtSubjectName.Text)
            cmd.Parameters.AddWithValue("@d3", cmbClass.Text)
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "added the new Subject '" & txtSubjectName.Text & "' having subject code='" & txtSubjectCode.Text & "'")
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
            Dim sql As String = "select Subject.SubjectCode from Subject,AttendanceMaster where Subject.SubjectCode=AttendanceMaster.SubjectCode and Subject.SubjectCode=@d1"
            cmd = New SqlCommand(sql)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Attendance Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim sql2 As String = "select Subject.SubjectCode from Subject,MarksEntry_Join where Subject.SubjectCode=MarksEntry_Join.SubCode and Subject.SubjectCode=@d1"
            cmd = New SqlCommand(sql2)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
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
            Dim cl As String = "select Subject.SubjectCode from Subject,Exam_Subject where Subject.SubjectCode=Exam_Subject.SubjectCode and Subject.SubjectCode=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Exam Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Subject where SubjectCode=@d1"
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the Subject '" & txtSubjectName.Text & "' having subject code='" & txtSubjectCode.Text & "'")
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
        If Len(Trim(txtSubjectCode.Text)) = 0 Then
            MessageBox.Show("Please enter subject code", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSubjectCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtSubjectName.Text)) = 0 Then
            MessageBox.Show("Please enter subject name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSubjectName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbClass.Text)) = 0 Then
            MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbClass.Focus()
            Exit Sub
        End If
        Try
            If txtSubjectCode.Text <> txtSubjectCode1.Text Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ct As String = "select SubjectCode from Subject where SubjectCode=@d1"
                cmd = New SqlCommand(ct)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    MessageBox.Show("Subject Code Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    Return
                End If
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "update MarksEntry_Join set SubCode=@d1 where SubCode=@d2"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
            cmd.Parameters.AddWithValue("@d2", txtSubjectCode1.Text)
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Subject set SubjectCode=@d1, SubjectName=@d2,Class=@d3 where SubjectCode=@d4"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubjectCode.Text)
            cmd.Parameters.AddWithValue("@d2", txtSubjectName.Text)
            cmd.Parameters.AddWithValue("@d3", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d4", txtSubjectCode1.Text)
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "updated the Subject '" & txtSubjectName.Text & "' having subject code='" & txtSubjectCode.Text & "'")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Getdata()
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(SubjectCode),RTRIM(SubjectName),RTRIM(Class) from Subject order by Subjectname", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2))
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
        fillCourse()
    End Sub

    Private Sub dgw_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                txtSubjectCode.Text = dr.Cells(0).Value.ToString()
                txtSubjectCode1.Text = dr.Cells(0).Value.ToString()
                txtSubjectName.Text = dr.Cells(1).Value.ToString()
                cmbClass.Text = dr.Cells(2).Value.ToString()
                btnUpdate.Enabled = True
                btnDelete.Enabled = True
                btnSave.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub cmbCourse_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbClass.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub

    Private Sub txtSearchByClass_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSearchByClass.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(SubjectCode),RTRIM(SubjectName),RTRIM(Class) from Subject where Class Like '%" & txtSearchByClass.Text & "%' order by Subjectname", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearchBySubject_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSearchBySubject.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(SubjectCode),RTRIM(SubjectName),RTRIM(Class) from Subject where SubjectName Like '%" & txtSearchBySubject.Text & "%' order by Subjectname", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
