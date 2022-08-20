Imports System.Data.SqlClient

Public Class frmStudentsAttendance_Record
    Dim Status As String
  
    Sub Reset()
        dgw.Rows.Clear()
        cmbSubjectName.Text = ""
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        DateTimePicker1.Value = Today
        DateTimePicker2.Value = Today
        txtSubjectCode.Text = ""

        fillSubjectname()
        GetData()
    End Sub

 
    Private Sub btnClose_Click_1(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnNew_Click_1(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub cmbSubjectName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSubjectName.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT SubjectCode FROM Subject WHERE SubjectName=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbSubjectName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtSubjectCode.Text = rdr.GetValue(0)
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

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub
    Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT   AttendanceMaster.ID, RTRIM(AttendanceMaster.AttendanceType), AttendanceMaster.Date, RTRIM(Subject.SubjectCode), RTRIM(Subject.SubjectName), RTRIM(SchoolInfo.SchoolName),RTRIM(Class.ClassName),RTRIM(Session_AM),Staff.St_ID, RTRIM(Staff.StaffID), RTRIM(Staff.StaffName) FROM            AttendanceMaster,Subject,Staff,Class,SchoolInfo where AttendanceMaster.SubjectCode=Subject.SubjectCode and Subject.Class=Class.ClassName and Staff.ST_ID=AttendanceMaster.StaffID and SchoolInfo.S_ID=Staff.SchoolID order by AttendanceMaster.Date", con)
            rdr = cmd.ExecuteReader()
            dgw.Rows.Clear()
            While rdr.Read()
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnSearch_Click(sender As System.Object, e As System.EventArgs) Handles btnSearch.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT   AttendanceMaster.ID, RTRIM(AttendanceMaster.AttendanceType), AttendanceMaster.Date, RTRIM(Subject.SubjectCode), RTRIM(Subject.SubjectName), RTRIM(SchoolInfo.SchoolName),RTRIM(Class.ClassName),RTRIM(Session_AM),Staff.St_ID, RTRIM(Staff.StaffID), RTRIM(Staff.StaffName) FROM            AttendanceMaster,Subject,Staff,Class,SchoolInfo where AttendanceMaster.SubjectCode=Subject.SubjectCode and Subject.Class=Class.ClassName and Staff.ST_ID=AttendanceMaster.StaffID and SchoolInfo.S_ID=Staff.SchoolID and AttendanceMaster.Date between @d1 and @d2 order by AttendanceMaster.Date", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            rdr = cmd.ExecuteReader()
            dgw.Rows.Clear()
            While rdr.Read()
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillSubjectname()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(SubjectName) FROM Subject order by 1", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSubjectName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSubjectName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            If Len(Trim(cmbSubjectName.Text)) = 0 Then
                MessageBox.Show("Please select subject name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSubjectName.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT AttendanceMaster.ID, RTRIM(AttendanceMaster.AttendanceType), AttendanceMaster.Date, RTRIM(Subject.SubjectCode), RTRIM(Subject.SubjectName), RTRIM(SchoolInfo.SchoolName),RTRIM(Class.ClassName),RTRIM(Session_AM),Staff.St_ID, RTRIM(Staff.StaffID), RTRIM(Staff.StaffName) FROM            AttendanceMaster,Subject,Staff,Class,SchoolInfo where AttendanceMaster.SubjectCode=Subject.SubjectCode and Subject.Class=Class.ClassName and Staff.ST_ID=AttendanceMaster.StaffID and SchoolInfo.S_ID=Staff.SchoolID and AttendanceMaster.Date between @d1 and @d2 and AttendanceMaster.SubjectCode=@d3 order by AttendanceMaster.Date", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker2.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker1.Value.Date
            cmd.Parameters.AddWithValue("@d3", txtSubjectCode.Text)
            rdr = cmd.ExecuteReader()
            dgw.Rows.Clear()
            While rdr.Read()
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmStudentsAttendance_Record_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSubjectname()
        GetData()
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                Me.Hide()
                frmStudentAttendance.Show()
                frmStudentAttendance.txtID.Text = dr.Cells(0).Value.ToString()
                frmStudentAttendance.cmbAttendanceType.Text = dr.Cells(1).Value.ToString()
                frmStudentAttendance.dtpDate.Text = dr.Cells(2).Value.ToString()
                frmStudentAttendance.cmbSchool.Text = dr.Cells(5).Value.ToString()
                frmStudentAttendance.cmbSession.Text = dr.Cells(7).Value.ToString()
                frmStudentAttendance.cmbClass.Text = dr.Cells(6).Value.ToString()
                frmStudentAttendance.txtSubjectCode.Text = dr.Cells(3).Value.ToString()
                frmStudentAttendance.cmbSubjectName.Text = dr.Cells(4).Value.ToString()
                frmStudentAttendance.txtSt_ID.Text = dr.Cells(8).Value.ToString()
                frmStudentAttendance.txtStaffID.Text = dr.Cells(9).Value.ToString()
                frmStudentAttendance.cmbStaffName.Text = dr.Cells(10).Value.ToString()
                frmStudentAttendance.btnUpdate.Enabled = True
                frmStudentAttendance.btnDelete.Enabled = True
                frmStudentAttendance.btnSave.Enabled = False
               
                con = New SqlConnection(cs)
                con.Open()
                cmd = New SqlCommand("Select distinct RTRIM(Student.AdmissionNo),RTRIM(Student.Studentname),RTRIM(Attendance.Status) from Student,AttendanceMaster,Attendance where Student.AdmissionNo=Attendance.AdmissionNo and Attendance.AttendanceID=AttendanceMaster.ID and AttendanceMaster.ID=" & Val(dr.Cells(0).Value) & " order by 2", con)
                rdr = cmd.ExecuteReader()
                frmStudentAttendance.listView1.Items.Clear()
                While rdr.Read()
                    Dim item = New ListViewItem()
                    item.Text = rdr(0).ToString()
                    item.SubItems.Add(rdr(1).ToString())
                    item.SubItems.Add(rdr(2).ToString())
                    frmStudentAttendance.listView1.Items.Add(item)
                    For i As Integer = frmStudentAttendance.listView1.Items.Count - 1 To 0 Step -1
                        If frmStudentAttendance.listView1.Items(i).SubItems(2).Text = "P" Then
                            frmStudentAttendance.listView1.Items(i).Checked = True
                        Else
                            frmStudentAttendance.listView1.Items(i).Checked = False
                        End If
                    Next
                End While
                con.Close()
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
End Class
