Imports System.Data.SqlClient
Public Class frmExamsEntry
    Sub Reset()
        txtClass.Text = ""
        txtSubjectCode.Text = ""
        txtSearchByClass.Text = ""
        txtSubjectName.Text = ""
        txtSearchBySubject.Text = ""
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        DataGridView1.Rows.Clear()
        cmbSchoolName.SelectedIndex = -1
        cmbExamType.SelectedIndex = -1
        Clear()
        Getdata()
        Getdata1()
        cmbExamType.DropDownStyle = ComboBoxStyle.DropDownList
        auto()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(cmbExamType.Text)) = 0 Then
            MessageBox.Show("Please select exam type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbExamType.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbSchoolName.Text)) = 0 Then
            MessageBox.Show("Please select college name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbSchoolName.Focus()
            Exit Sub
        End If
        If (Me.DataGridView1.Rows.Count = 0) Then
            MessageBox.Show("Sorry..No subject info added to list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Exam(ID,ExamType,SchoolID) VALUES (@d1,@d2," & txtSchoolID.Text & ")"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", cmbExamType.Text)
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cmdText1 As String = "insert into Exam_Subject(ExamID,SubjectCode,ExamDate) VALUES (@d1,@d2,@d3)"
            cmd = New SqlCommand(cmdText1)
            cmd.Connection = con
            cmd.Prepare()
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", txtID.Text)
                    cmd.Parameters.AddWithValue("@d2", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d3", Convert.ToDateTime(row.Cells(3).Value))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            LogFunc(lblUser.Text, "added the new exam record having exam id='" & txtID.Text & "'")
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Getdata1()
            btnSave.Enabled = False
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
            Dim cq As String = "delete from Exam where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the exam record having exam id='" & txtID.Text & "'")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Getdata1()
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
        If Len(Trim(cmbExamType.Text)) = 0 Then
            MessageBox.Show("Please select exam type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbExamType.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbSchoolName.Text)) = 0 Then
            MessageBox.Show("Please select college name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbSchoolName.Focus()
            Exit Sub
        End If
        If (Me.DataGridView1.Rows.Count = 0) Then
            MessageBox.Show("Sorry..No subject info added to list", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update exam set ExamType=@d2,SchoolID=" & txtSchoolID.Text & " where ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", cmbExamType.Text)
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Exam_Subject where ExamID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cmdText1 As String = "insert into Exam_Subject(ExamID,SubjectCode,ExamDate) VALUES (@d1,@d2,@d3)"
            cmd = New SqlCommand(cmdText1)
            cmd.Connection = con
            cmd.Prepare()
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", txtID.Text)
                    cmd.Parameters.AddWithValue("@d2", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d3", Convert.ToDateTime(row.Cells(3).Value))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            LogFunc(lblUser.Text, "updated the exam record having exam id='" & txtID.Text & "'")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Getdata1()
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(SubjectCode),RTRIM(SubjectName),RTRIM(Class) from Subject order by SubjectName", con)
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
    Public Sub Getdata1()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(ID),RTRIM(ExamType),SchoolID,RTRIM(SchoolName) from Exam,SchoolInfo where Exam.SchoolID=SchoolInfo.S_ID order by ID", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView2.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView2.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM Exam")
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
    Private Sub frmCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Getdata()
        Getdata1()
        fillSchool()
        fillExamType()
    End Sub

    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Clear()
    End Sub

    Private Sub btnAddToList_Click(sender As System.Object, e As System.EventArgs) Handles btnAddToList.Click
        Try
            If Me.txtSubjectCode.Text = "" Then
                MessageBox.Show("Please retrieve subject code", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.txtSubjectCode.Focus()
                Exit Sub
            End If
            For Each row As DataGridViewRow In DataGridView1.Rows
                If txtSubjectCode.Text = row.Cells(0).Value Then
                    MessageBox.Show("Subject code already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next
            Me.DataGridView1.Rows.Add(txtSubjectCode.Text, txtSubjectName.Text, txtClass.Text, dtpExamDate.Value.Date)
            Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnListRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnListRemove.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
            Next
            Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnListUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnListUpdate.Click
        Try
            If Me.txtSubjectCode.Text = "" Then
                MessageBox.Show("Please retrieve subject code", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.txtSubjectCode.Focus()
                Exit Sub
            End If
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
            Next
            Me.DataGridView1.Rows.Add(txtSubjectCode.Text, txtSubjectName.Text, txtClass.Text, dtpExamDate.Value.Date)
            Me.Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub Clear()
        Me.txtClass.Text = ""
        Me.txtSubjectCode.Text = ""
        txtSubjectName.Text = ""
        dtpExamDate.Text = Today
        Me.btnAddToList.Enabled = True
        Me.btnListRemove.Enabled = False
        Me.btnListUpdate.Enabled = False
    End Sub

    Private Sub DataGridView1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        Try
            If DataGridView1.Rows.Count > 0 Then
                Me.btnListRemove.Enabled = True
                Me.btnListUpdate.Enabled = True
                Me.btnAddToList.Enabled = False
                Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
                txtSubjectCode.Text = dr.Cells(0).Value.ToString()
                txtSubjectName.Text = dr.Cells(1).Value.ToString()
                txtClass.Text = dr.Cells(2).Value.ToString()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowPostPaint(sender As Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub DataGridView2_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseClick
        Try
            If DataGridView2.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
                txtID.Text = dr.Cells(0).Value.ToString()
                cmbExamType.Text = dr.Cells(1).Value.ToString()
                txtSchoolID.Text = dr.Cells(2).Value.ToString()
                cmbSchoolName.Text = dr.Cells(3).Value.ToString()
                con = New SqlConnection(cs)
                con.Open()
                cmd = New SqlCommand("select RTRIM(Subject.SubjectCode),RTRIM(SubjectName),RTRIM(Class),ExamDate from Class,Subject,Exam,Exam_Subject where Subject.Class=Class.Classname and Exam.ID=Exam_Subject.ExamID and Subject.SubjectCode=Exam_Subject.SubjectCode and Exam.ID=" & dr.Cells(0).Value & " ", con)
                cmd.Parameters.AddWithValue("@d1", dr.Cells(0).Value)
                rdr = cmd.ExecuteReader()
                DataGridView1.Rows.Clear()
                While rdr.Read()
                    DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
                End While
                con.Close()
                Me.btnDelete.Enabled = True
                Me.btnUpdate.Enabled = True
                Me.btnSave.Enabled = False
                cmbExamType.DropDownStyle = ComboBoxStyle.DropDown
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_RowPostPaint(sender As Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub cmbCollegeType_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSchoolName.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT S_ID FROM SchoolInfo where SchoolName=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbSchoolName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtSchoolID.Text = rdr.GetValue(0)
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
    Sub fillSchool()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (SchoolName) FROM SchoolInfo", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSchoolName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSchoolName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillExamType()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (type) FROM ExamType", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbExamType.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbExamType.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbExamType_Format(sender As Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbExamType.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                txtSubjectCode.Text = dr.Cells(0).Value.ToString()
                txtSubjectName.Text = dr.Cells(1).Value.ToString()
                txtClass.Text = dr.Cells(2).Value.ToString()
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
