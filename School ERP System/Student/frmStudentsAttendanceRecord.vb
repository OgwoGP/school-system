Imports System.Data.SqlClient

Public Class frmStudentsAttendanceRecord
    Dim Status As String
    Sub fillschool()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (SchoolName) FROM SchoolInfo", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSchool.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSchool.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Sub Reset()
        cmbSchool.SelectedIndex = -1
        cmbclass.SelectedIndex = -1
        cmbSession.SelectedIndex = -1
        dgw.Rows.Clear()
        cmbSession.Enabled = False
        cmbclass.Enabled = False
        cmbSubjectName.SelectedIndex = -1
        cmbSubjectName.Enabled = False
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        DateTimePicker1.Value = Today
        DateTimePicker2.Value = Today
        txtSubjectCode.Text = ""
        lblTotalClasses.Visible = False
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnSearch.Click
        Try
            If Len(Trim(cmbSchool.Text)) = 0 Then
                MessageBox.Show("Please select school name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSchool.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbSession.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbClass.Text)) = 0 Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select distinct RTRIM(Student.AdmissionNo),RTRIM(Student.StudentName),Count(Attendance.Status),(Count(Attendance.Status) * 100)/(Select Count(distinct Attendance.AttendanceID) from Student,class,Section,SchoolInfo,Attendance,AttendanceMaster where Student.SectionID=Section.ID and class.classname=Section.class and SchoolInfo.S_ID=Student.SchoolID and AttendanceMaster.ID=Attendance.AttendanceID and Attendance.AdmissionNo=Student.AdmissionNo and Session_AM=@d1 and classname=@d2 and SchoolName=@d3 and AttendanceMaster.Date between @date1 and @date2 ) from Student,class,Section,SchoolInfo,Attendance,AttendanceMaster where Student.SectionID=Section.ID and class.classname=Section.class and SchoolInfo.S_ID=Student.SchoolID and AttendanceMaster.ID=Attendance.AttendanceID and Attendance.AdmissionNo=Student.AdmissionNo and Session_AM=@d1 and classname=@d2 and Schoolname=@d3 and Attendance.Status='P' and AttendanceMaster.Date between @date1 and @date2 group by Student.AdmissionNo,Student.StudentName order by 2", con)
            cmd.Parameters.Add("@date1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@date2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbclass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSchool.Text)
            rdr = cmd.ExecuteReader()
            dgw.Rows.Clear()
            While rdr.Read()
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
            lblTotalClasses.Visible = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select Count(distinct Attendance.AttendanceID) from Student,class,Section,SchoolInfo,Attendance,AttendanceMaster where Student.SectionID=Section.ID and class.classname=Section.class and SchoolInfo.S_ID=Student.SchoolID and AttendanceMaster.ID=Attendance.AttendanceID and Attendance.AdmissionNo=Student.AdmissionNo and Session_AM=@d1 and classname=@d2 and SchoolName=@d3 and AttendanceMaster.Date between @date1 and @date2", con)
            cmd.Parameters.Add("@date1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@date2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbclass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSchool.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                lblTotalClasses.Text = rdr.GetValue(0)
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmDiscount_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillschool()
    End Sub

    Private Sub cmbSession_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles cmbSession.SelectedIndexChanged
        Try
            cmbClass.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(className) FROM Student,class,Section,SchoolInfo,AttendanceMaster,Attendance where Student.SectionID=Section.ID and Section.class=class.classname and Student.SchoolID=SchoolInfo.S_ID and Student.AdmissionNo=Attendance.AdmissionNo and AttendanceMaster.ID=Attendance.AttendanceID and Session_AM=@d1 and SchoolName=@d2"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbSchool.Text)
            rdr = cmd.ExecuteReader()
            cmbClass.Items.Clear()
            While rdr.Read
                cmbClass.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbclass_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbClass.SelectedIndexChanged
        Try
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select distinct RTRIM(SubjectName) from Student,Class,Section,SchoolInfo,Subject,Attendance,AttendanceMaster where Student.SectionID=Section.ID and Class.Classname=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Subject.Class=Class.ClassName and Student.AdmissionNo=Attendance.AdmissionNo and AttendanceMaster.ID=Attendance.AttendanceID and Session_AM=@d1 and classname=@d2 and Schoolname=@d3 order by 1", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSchool.Text)
            rdr = cmd.ExecuteReader()
            cmbSubjectName.Enabled = True
            cmbSubjectName.Items.Clear()
            While rdr.Read()
                cmbSubjectName.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub


    Private Sub cmbSchool_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSchool.SelectedIndexChanged
        Try
            cmbSession.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Session_AM) FROM Student,SchoolInfo,AttendanceMaster,Attendance where Student.SchoolID=SchoolInfo.S_ID and Student.AdmissionNo=Attendance.AdmissionNo and AttendanceMaster.ID=Attendance.AttendanceID and SchoolName=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSchool.Text)
            rdr = cmd.ExecuteReader()
            cmbSession.Items.Clear()
            While rdr.Read()
                cmbSession.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub btnClose_Click_1(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnNew_Click_1(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
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


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            If Len(Trim(cmbSchool.Text)) = 0 Then
                MessageBox.Show("Please select school name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSchool.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbSession.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbclass.Text)) = 0 Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbclass.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbSubjectName.Text)) = 0 Then
                MessageBox.Show("Please select subject name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSubjectName.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select distinct RTRIM(Student.AdmissionNo),RTRIM(Student.StudentName),Count(Attendance.Status),(Count(Attendance.Status) * 100)/(Select Count(distinct Attendance.AttendanceID) from Student,class,Section,SchoolInfo,Attendance,AttendanceMaster,Subject where Student.SectionID=Section.ID and class.classname=Section.Class and SchoolInfo.S_ID=Student.SchoolID and AttendanceMaster.ID=Attendance.AttendanceID and Attendance.AdmissionNo=Student.AdmissionNo and Subject.SubjectCode=AttendanceMaster.SubjectCode and Session_AM=@d1 and classname=@d2 and SchoolName=@d3 and SubjectName=@d4 and AttendanceMaster.Date between @date1 and @date2 ) from Subject, Student,class,Section,SchoolInfo,Attendance,AttendanceMaster where Student.SectionID=Section.ID and class.classname=Section.class and SchoolInfo.S_ID=Student.SchoolID and AttendanceMaster.ID=Attendance.AttendanceID and Attendance.AdmissionNo=Student.AdmissionNo and Subject.SubjectCode=AttendanceMaster.SubjectCode and Session_AM=@d1 and classname=@d2 and Schoolname=@d3 and SubjectName=@d4 and Attendance.Status='P' and AttendanceMaster.Date between @date1 and @date2 group by Student.AdmissionNo,Student.StudentName order by 2", con)
            cmd.Parameters.Add("@date1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker2.Value.Date
            cmd.Parameters.Add("@date2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker1.Value.Date
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSchool.Text)
            cmd.Parameters.AddWithValue("@d4", cmbSubjectName.Text)
            rdr = cmd.ExecuteReader()
            dgw.Rows.Clear()
            While rdr.Read()
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
            lblTotalClasses.Visible = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select Count(distinct Attendance.AttendanceID) from Student,class,Section,SchoolInfo,Attendance,AttendanceMaster,Subject where Student.SectionID=Section.ID and class.classname=Section.class and SchoolInfo.S_ID=Student.SchoolID and AttendanceMaster.ID=Attendance.AttendanceID and Attendance.AdmissionNo=Student.AdmissionNo and Subject.SubjectCode=AttendanceMaster.SubjectCode and Session_AM=@d1 and classname=@d2 and SchoolName=@d3 and SubjectName=@d4 and AttendanceMaster.Date between @date1 and @date2", con)
            cmd.Parameters.Add("@date1", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker2.Value.Date
            cmd.Parameters.Add("@date2", SqlDbType.DateTime, 30, "Date").Value = DateTimePicker1.Value.Date
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbclass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSchool.Text)
            cmd.Parameters.AddWithValue("@d4", cmbSubjectName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                lblTotalClasses.Text = rdr.GetValue(0)
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        ExportExcel(dgw)
    End Sub
End Class
