Imports System.Data.SqlClient
Public Class frmStudentAttendance
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

    Sub fillStaffName()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT RTRIM(StaffName) FROM Staff", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbStaffName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbStaffName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillAttendanceType()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT RTRIM(Type) FROM AttendanceType", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbAttendanceType.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbAttendanceType.Items.Add(drow(0).ToString())
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
        cmbClass.SelectedIndex = -1
        cmbSession.SelectedIndex = -1
        txtSubjectCode.Text = ""
        listView1.Items.Clear()
        cmbSession.Enabled = False
        cmbClass.Enabled = False
        btnUpdate.Enabled = False
        btnSave.Enabled = True
        btnDelete.Enabled = False
        listView1.Items.Clear()
        cmbSubjectName.SelectedIndex = -1
        cmbSubjectName.DropDownStyle = ComboBoxStyle.DropDownList
        cmbClass.DropDownStyle = ComboBoxStyle.DropDownList
        cmbSubjectName.Enabled = False
        dtpDate.Value = Today
        cmbAttendanceType.SelectedIndex = -1
        cmbStaffName.Text = ""
        txtStaffID.Text = ""
        txtSt_ID.Text = ""
        auto()
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
            listView1.Items.Clear()
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select distinct RTRIM(Student.AdmissionNo),RTRIM(Student.StudentName) from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.Classname=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Session=@d1 and Classname=@d2 and SchoolName=@d3 order by 2", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSchool.Text)
            rdr = cmd.ExecuteReader()
            listView1.Items.Clear()
            While rdr.Read()
                Dim item = New ListViewItem()
                item.Text = rdr(0).ToString().Trim()
                item.SubItems.Add(rdr(1).ToString().Trim())
                listView1.Items.Add(item)
            End While
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select distinct RTRIM(SubjectName) from Student,Class,Section,SchoolInfo,Subject where Student.SectionID=Section.ID and Class.Classname=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Subject.Class=Class.ClassName and Session=@d1 and classname=@d2 and Schoolname=@d3 order by 1", con)
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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM AttendanceMaster")
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
    Private Sub frmDiscount_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillschool()
        fillAttendanceType()
        fillStaffName()
    End Sub

    Private Sub cmbSession_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles cmbSession.SelectedIndexChanged
        Try
            cmbClass.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(className) FROM Student,class,Section,SchoolInfo where Student.SectionID=Section.ID and Section.class=class.classname and Student.SchoolID=SchoolInfo.S_ID and Session=@d1 and SchoolName=@d2"
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
    Private Sub cmbschoolName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSchool.SelectedIndexChanged
        Try
            cmbSession.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Session) FROM Student,SchoolInfo where Student.SchoolID=SchoolInfo.S_ID and SchoolName=@d1"
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

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
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
            If Len(Trim(cmbSubjectName.Text)) = 0 Then
                MessageBox.Show("Please select subject name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSubjectName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbStaffName.Text)) = 0 Then
                MessageBox.Show("Please select staff name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbStaffName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtStaffID.Text)) = 0 Then
                MessageBox.Show("Please retrieve staff id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtStaffID.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbAttendanceType.Text)) = 0 Then
                MessageBox.Show("Please select attendance type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbAttendanceType.Focus()
                Exit Sub
            End If

            If listView1.Items.Count = 0 Then
                MessageBox.Show("Sorry nothing to save.." & vbCrLf & "Please retrieve data in listview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            con = New SqlConnection(cs)
            Dim cd As String = "Insert Into AttendanceMaster(ID,AttendanceType,StaffID,Date,SubjectCode,Session_AM) Values(@d1,@d2,@d3,@d4,@d5,@d6)"
            cmd = New SqlCommand(cd)
            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            cmd.Parameters.AddWithValue("@d2", cmbAttendanceType.Text)
            cmd.Parameters.AddWithValue("@d3", Val(txtSt_ID.Text))
            cmd.Parameters.AddWithValue("@d4", dtpDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", txtSubjectCode.Text)
            cmd.Parameters.AddWithValue("@d6", cmbSession.Text)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            For i As Integer = listView1.Items.Count - 1 To 0 Step -1

                If listView1.Items(i).Checked = True Then
                    Status = "P"
                Else
                    Status = "A"
                End If
                con = New SqlConnection(cs)
                Dim cd1 As String = "Insert Into Attendance(AdmissionNo,Status,AttendanceID) Values(@d1,@d2,@d3)"
                cmd = New SqlCommand(cd1)
                cmd.Parameters.AddWithValue("@d1", listView1.Items(i).SubItems(0).Text)
                cmd.Parameters.AddWithValue("@d2", Status)
                cmd.Parameters.AddWithValue("@d3", Val(txtID.Text))
                cmd.Connection = con
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            LogFunc(lblUser.Text, "added the new attendance record having attendance id '" & txtID.Text & "'")
            auto()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
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

    Private Sub cmbStaffName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbStaffName.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT St_ID,StaffID FROM Staff WHERE StaffName=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbStaffName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtSt_ID.Text = rdr.GetValue(0)
                txtStaffID.Text = rdr.GetValue(1)
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

    Private Sub cmbStaffName_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbStaffName.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
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
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from AttendanceMaster where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the attendance record having attendance id '" & txtID.Text & "'")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
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
            If Len(Trim(cmbSubjectName.Text)) = 0 Then
                MessageBox.Show("Please select subject name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSubjectName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbStaffName.Text)) = 0 Then
                MessageBox.Show("Please select staff name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbStaffName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtStaffID.Text)) = 0 Then
                MessageBox.Show("Please retrieve staff id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtStaffID.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbAttendanceType.Text)) = 0 Then
                MessageBox.Show("Please select attendance type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbAttendanceType.Focus()
                Exit Sub
            End If

            If listView1.Items.Count = 0 Then
                MessageBox.Show("Sorry nothing to save.." & vbCrLf & "Please retrieve data in listview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            con = New SqlConnection(cs)
            Dim cd As String = "Update AttendanceMaster set AttendanceType=@d2,StaffID=@d3,Date=@d4,SubjectCode=@d5,Session_AM=@d6 where ID=@d1"
            cmd = New SqlCommand(cd)
            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            cmd.Parameters.AddWithValue("@d2", cmbAttendanceType.Text)
            cmd.Parameters.AddWithValue("@d3", Val(txtSt_ID.Text))
            cmd.Parameters.AddWithValue("@d4", dtpDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", txtSubjectCode.Text)
            cmd.Parameters.AddWithValue("@d6", cmbSession.Text)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Attendance where AttendanceID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", Val(txtID.Text))
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            For i As Integer = listView1.Items.Count - 1 To 0 Step -1

                If listView1.Items(i).Checked = True Then
                    Status = "P"
                Else
                    Status = "A"
                End If
                con = New SqlConnection(cs)
                Dim cd1 As String = "Insert Into Attendance(AdmissionNo,Status,AttendanceID) Values(@d1,@d2,@d3)"
                cmd = New SqlCommand(cd1)
                cmd.Parameters.AddWithValue("@d1", listView1.Items(i).SubItems(0).Text)
                cmd.Parameters.AddWithValue("@d2", Status)
                cmd.Parameters.AddWithValue("@d3", txtID.Text)
                cmd.Connection = con
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            LogFunc(lblUser.Text, "updated the attendance record having attendance id '" & txtID.Text & "'")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub cmbClass_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbClass.SelectedIndexChanged
        Try
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select distinct RTRIM(SubjectName) from Student,Class,Section,SchoolInfo,Subject where Student.SectionID=Section.ID and Class.Classname=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Subject.Class=Class.ClassName and Session=@d1 and classname=@d2 and Schoolname=@d3 order by 1", con)
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

    Private Sub btnGetData_Click_1(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmStudentsAttendance_Record.Reset()
        frmStudentsAttendance_Record.ShowDialog()
    End Sub
End Class
