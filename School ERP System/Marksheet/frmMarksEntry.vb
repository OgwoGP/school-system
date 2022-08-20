Imports System.Data.SqlClient
Imports System.IO

Public Class frmMarksEntry
    Dim stX As String
    Public Sub auto()
        Try
            Try
                Dim Num As Integer = 0
                con = New SqlConnection(cs)
                con.Open()
                Dim sql As String = ("SELECT MAX(M_ID) FROM MarksEntry")
                cmd = New SqlCommand(sql)
                cmd.Connection = con
                If (IsDBNull(cmd.ExecuteScalar)) Then
                    Num = 1
                    txtMarksID.Text = Num.ToString
                Else
                    Num = cmd.ExecuteScalar + 1
                    txtMarksID.Text = Num.ToString
                End If
                cmd.Dispose()
                con.Close()
                con.Dispose()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Sub Reset()
        txtAdmissionNo.Text = ""
        txtClass.Text = ""
        txtCreditPoint.Text = ""
        txtEnrollmentNo.Text = ""
        txtFinalGrade.Text = ""
        txtGradePoint.Text = ""
        txtMaxMarksT.Text = ""
        txtOGPractical.Text = ""
        txtOGTheory.Text = ""
        txtOMPractical.Text = ""
        txtOMTheory.Text = ""
        txtSchoolName.Text = ""
        txtStudentName.Text = ""
        txtSubjectCode.Text = ""
        txtSession.Text = ""
        cmbSubjectName.SelectedIndex = -1
        dgw.Rows.Clear()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        btnSelection.Enabled = True
        btnAddOS.Enabled = True
        btnRemoveFromGridOS.Enabled = False
        btnPrint.Enabled = False
        btnPrint1.Enabled = False
        rbPass.Checked = True
        rbFail.Checked = False
        clear()
        GetData()
        cmbSession.SelectedIndex = -1
        cmbClass.Enabled = False
        cmbClass.SelectedIndex = -1
    End Sub
    Sub FillSubject()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RTRIM(SubjectName) from Subject where Class=@d1"
            cmd.Parameters.AddWithValue("@d1", txtClass.Text)
            cmbSubjectName.Items.Clear()
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                cmbSubjectName.Items.Add(rdr(0))
            End While
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles btnSelection.Click
        frmActiveStudentRecord.lblSet.Text = "Marks Entry"
        frmActiveStudentRecord.Reset()
        frmActiveStudentRecord.ShowDialog()
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
        Reset()
    End Sub
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],Photo from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            DataGridView1.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtAdmissionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve student info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        If dgw.Rows.Count = 0 Then
            MessageBox.Show("Please add subject and grade info in a datagrid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select Session,AdmissionNo from MarksEntry where Session=@d1 and AdmissionNo=@d2"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSession.Text)
            cmd.Parameters.AddWithValue("@d2", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Record already exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            If rbPass.Checked = True Then
                stX = rbPass.Text
            Else
                stX = rbFail.Text
            End If
            auto()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into MarksEntry(M_ID, AdmissionNo, Session,EntryDate,Student_Class,Result) VALUES (@d1,@d2,@d3,@d4,@d5,@d6)"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", Val(txtMarksID.Text))
            cmd.Parameters.AddWithValue("@d2", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d3", txtSession.Text)
            cmd.Parameters.AddWithValue("@d4", System.DateTime.Now)
            cmd.Parameters.AddWithValue("@d5", txtClass.Text)
            cmd.Parameters.AddWithValue("@d6", stX)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into MarksEntry_Join(MarksID,SubCode, MaxMarks,MMPractical, CreditHour, OMTheory,OMPractical ,OGTheory, OGPractical, FinalGrade, GradePoint) VALUES (" & txtMarksID.Text & ",@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10)"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d2", Val(row.Cells(2).Value))
                    cmd.Parameters.AddWithValue("@d3", Val(row.Cells(3).Value))
                    cmd.Parameters.AddWithValue("@d4", Val(row.Cells(4).Value))
                    cmd.Parameters.AddWithValue("@d5", Val(row.Cells(5).Value))
                    cmd.Parameters.AddWithValue("@d6", Val(row.Cells(6).Value))
                    cmd.Parameters.AddWithValue("@d7", row.Cells(7).Value)
                    cmd.Parameters.AddWithValue("@d8", row.Cells(8).Value)
                    cmd.Parameters.AddWithValue("@d9", row.Cells(9).Value)
                    cmd.Parameters.AddWithValue("@d10", Val(row.Cells(10).Value))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            Dim st As String = "Added marks for student having admission no.'" & txtAdmissionNo.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            btnAddOS.Enabled = False
            btnRemoveFromGridOS.Enabled = False
            btnPrint.Enabled = True
            btnPrint1.Enabled = True
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete the record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
                delete_records()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub delete_records()
        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from MarksEntry where M_ID= " & txtMarksID.Text & ""
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the marks for student having admission no.'" & txtAdmissionNo.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
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
        If Len(Trim(txtAdmissionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve student info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        If dgw.Rows.Count = 0 Then
            MessageBox.Show("Please add subject and grade info in a datagrid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            If rbPass.Checked = True Then
                stX = rbPass.Text
            Else
                stX = rbFail.Text
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update MarksEntry set Result=@d2 where M_ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", Val(txtMarksID.Text))
            cmd.Parameters.AddWithValue("@d2", stX)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from MarksEntry_Join where MarksID= " & txtMarksID.Text & ""
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into MarksEntry_Join(MarksID,SubCode, MaxMarks,MMPractical, CreditHour, OMTheory,OMPractical ,OGTheory, OGPractical, FinalGrade, GradePoint) VALUES (" & txtMarksID.Text & ",@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10)"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d2", Val(row.Cells(2).Value))
                    cmd.Parameters.AddWithValue("@d3", Val(row.Cells(3).Value))
                    cmd.Parameters.AddWithValue("@d4", Val(row.Cells(4).Value))
                    cmd.Parameters.AddWithValue("@d5", Val(row.Cells(5).Value))
                    cmd.Parameters.AddWithValue("@d6", Val(row.Cells(6).Value))
                    cmd.Parameters.AddWithValue("@d7", row.Cells(7).Value)
                    cmd.Parameters.AddWithValue("@d8", row.Cells(8).Value)
                    cmd.Parameters.AddWithValue("@d9", row.Cells(9).Value)
                    cmd.Parameters.AddWithValue("@d10", Val(row.Cells(10).Value))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            Dim st As String = "updated the marks of student having admission no.'" & txtAdmissionNo.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmMarksEntryRecord.lblSet.Text = "Marks Entry"
        frmMarksEntryRecord.lblUserType.Text = lblUserType.Text
        frmMarksEntryRecord.Reset()
        frmMarksEntryRecord.ShowDialog()
    End Sub
    Private Sub txtMaxMarks_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxMarksT.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtCreditPoint_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCreditPoint.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtOMTheory_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtOMTheory.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtOMPractical_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtOMPractical.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub cmbSubjectName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSubjectName.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RTRIM(SubjectCode) from Subject where SubjectName=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbSubjectName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtSubjectCode.Text = rdr.GetValue(0)
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub frmMarksEntry_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        GetData()
        fillSession()
    End Sub

    Private Sub txtOMTheory_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtOMTheory.TextChanged
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 90 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 100 Then
            txtOGTheory.Text = "A+"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 80 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 89 Then
            txtOGTheory.Text = "A"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 70 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 79 Then
            txtOGTheory.Text = "B+"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 60 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 69 Then
            txtOGTheory.Text = "B"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 50 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 59 Then
            txtOGTheory.Text = "C+"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 40 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 49 Then
            txtOGTheory.Text = "C"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 30 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 39 Then
            txtOGTheory.Text = "D+"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 20 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 29 Then
            txtOGTheory.Text = "D"
        End If
        If ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 0 And ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 19 Then
            txtOGTheory.Text = "E"
        End If
        If txtOMTheory.Text = "" Then
            txtOGTheory.Text = ""
        End If
        FinalGrade()
        GetGradePoint()
    End Sub

    Private Sub OMPractical_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtOMPractical.TextChanged
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 90 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 100 Then
            txtOGPractical.Text = "A+"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 80 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 89 Then
            txtOGPractical.Text = "A"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 70 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 79 Then
            txtOGPractical.Text = "B+"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 60 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 69 Then
            txtOGPractical.Text = "B"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 50 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 59 Then
            txtOGPractical.Text = "C+"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 40 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 49 Then
            txtOGPractical.Text = "C"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 30 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 39 Then
            txtOGPractical.Text = "D+"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 20 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 29 Then
            txtOGPractical.Text = "D"
        End If
        If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) >= 0 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) <= 19 Then
            txtOGPractical.Text = "E"
        End If
        If txtOMPractical.Text = "" Then
            txtOGPractical.Text = ""
        End If
        FinalGrade()
        GetGradePoint()
    End Sub
    Sub FinalGrade()
        If txtOMPractical.Text <> "" And txtOMTheory.Text <> "" Then
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 90 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 100 Then
                txtFinalGrade.Text = "A+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 80 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 89 Then
                txtFinalGrade.Text = "A"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 70 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 79 Then
                txtFinalGrade.Text = "B+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 60 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 69 Then
                txtFinalGrade.Text = "B"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 50 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 59 Then
                txtFinalGrade.Text = "C+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 40 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 49 Then
                txtFinalGrade.Text = "C"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 30 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 39 Then
                txtFinalGrade.Text = "D+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 20 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 29 Then
                txtFinalGrade.Text = "D"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 0 And ((((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 19 Then
                txtFinalGrade.Text = "E"
            End If
        End If
        If txtOMPractical.Text = "" And txtOMTheory.Text <> "" Then
            If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 90 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 100 Then
                txtFinalGrade.Text = "A+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 80 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 89 Then
                txtFinalGrade.Text = "A"
            End If
            If ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 70 And ((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 79 Then
                txtFinalGrade.Text = "B+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 60 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 69 Then
                txtFinalGrade.Text = "B"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 50 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 59 Then
                txtFinalGrade.Text = "C+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 40 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 49 Then
                txtFinalGrade.Text = "C"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 30 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 39 Then
                txtFinalGrade.Text = "D+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 20 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 29 Then
                txtFinalGrade.Text = "D"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 0 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 19 Then
                txtFinalGrade.Text = "E"
            End If
        End If
        If txtOMPractical.Text <> "" And txtOMTheory.Text = "" Then
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 90 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 100 Then
                txtFinalGrade.Text = "A+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 80 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 89 Then
                txtFinalGrade.Text = "A"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 70 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 79 Then
                txtFinalGrade.Text = "B+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 60 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 69 Then
                txtFinalGrade.Text = "B"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 50 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 59 Then
                txtFinalGrade.Text = "C+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 40 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 49 Then
                txtFinalGrade.Text = "C"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 30 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 39 Then
                txtFinalGrade.Text = "D+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 20 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 29 Then
                txtFinalGrade.Text = "D"
            End If
            If (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 0 And (((Val(txtOMPractical.Text) * 100) / Val(txtMaxMarksP.Text)) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 19 Then
                txtFinalGrade.Text = "E"
            End If
        End If
        If txtOMPractical.Text <> "" And txtOMTheory.Text <> "" And txtMaxMarksP.Text = "" Then
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 90 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 100 Then
                txtFinalGrade.Text = "A+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 80 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 89 Then
                txtFinalGrade.Text = "A"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 70 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 79 Then
                txtFinalGrade.Text = "B+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 60 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 69 Then
                txtFinalGrade.Text = "B"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 50 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 59 Then
                txtFinalGrade.Text = "C+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 40 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 49 Then
                txtFinalGrade.Text = "C"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 30 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 39 Then
                txtFinalGrade.Text = "D+"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 20 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 29 Then
                txtFinalGrade.Text = "D"
            End If
            If ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) >= 0 And ((((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) / 2) <= 19 Then
                txtFinalGrade.Text = "E"
            End If
        End If
        If txtOMPractical.Text = "" And txtOMTheory.Text <> "" And txtMaxMarksP.Text = "" Then
            If ((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 90 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 100 Then
                txtFinalGrade.Text = "A+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 80 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 89 Then
                txtFinalGrade.Text = "A"
            End If
            If ((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) >= 70 And ((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text)) <= 79 Then
                txtFinalGrade.Text = "B+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 60 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 69 Then
                txtFinalGrade.Text = "B"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 50 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 59 Then
                txtFinalGrade.Text = "C+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 40 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 49 Then
                txtFinalGrade.Text = "C"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 30 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 39 Then
                txtFinalGrade.Text = "D+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 20 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 29 Then
                txtFinalGrade.Text = "D"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 0 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 19 Then
                txtFinalGrade.Text = "E"
            End If
        End If
        If txtOMPractical.Text <> "" And txtOMTheory.Text = "" And txtMaxMarksP.Text = "" Then
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 90 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 100 Then
                txtFinalGrade.Text = "A+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 80 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 89 Then
                txtFinalGrade.Text = "A"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 70 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 79 Then
                txtFinalGrade.Text = "B+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 60 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 69 Then
                txtFinalGrade.Text = "B"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 50 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 59 Then
                txtFinalGrade.Text = "C+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 40 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 49 Then
                txtFinalGrade.Text = "C"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 30 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 39 Then
                txtFinalGrade.Text = "D+"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 20 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 29 Then
                txtFinalGrade.Text = "D"
            End If
            If (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) >= 0 And (((Val(txtOMPractical.Text) * 100) / 1) + ((Val(txtOMTheory.Text) * 100) / Val(txtMaxMarksT.Text))) <= 19 Then
                txtFinalGrade.Text = "E"
            End If
        End If
        If txtOMPractical.Text = "" And txtOMTheory.Text = "" Then
            txtFinalGrade.Text = ""
            txtGradePoint.Text = ""
        End If
    End Sub

    Private Sub btnRemoveFromGridOS_Click(sender As System.Object, e As System.EventArgs) Handles btnRemoveFromGridOS.Click
        Try
            If dgw.Rows.Count > 0 Then
                For Each row As DataGridViewRow In dgw.SelectedRows
                    dgw.Rows.Remove(row)
                Next
                btnRemoveFromGridOS.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            btnRemoveFromGridOS.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddOS_Click(sender As System.Object, e As System.EventArgs) Handles btnAddOS.Click
        Try

            If cmbSubjectName.Text = "" Then
                MessageBox.Show("Please select subject name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSubjectName.Focus()
                Exit Sub
            End If
            If Val(txtMaxMarksT.Text) + Val(txtMaxMarksP.Text) > 100 Then
                MessageBox.Show("Max. Marks(T) + Max Marks(P) must be less than or equal to 100", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If txtCreditPoint.Text = "" Then
                MessageBox.Show("Please enter credit point", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCreditPoint.Focus()
                Exit Sub
            End If
            If txtOMPractical.Text = "" And txtOMTheory.Text = "" Then
                MessageBox.Show("Please enter obtained marks", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If Val(txtOMPractical.Text) + Val(txtOMTheory.Text) > 100 Then
                MessageBox.Show("Obtained marks(Theory + Practical) must be less than or equal to 100", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            For Each row As DataGridViewRow In dgw.Rows
                If row.Cells(0).Value = txtSubjectCode.Text Then
                    MessageBox.Show("Subject code already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    clear()
                    Exit Sub
                End If
            Next
            dgw.Rows.Add(txtSubjectCode.Text, cmbSubjectName.Text, txtMaxMarksT.Text, txtMaxMarksP.Text, txtCreditPoint.Text, txtOMTheory.Text, txtOMPractical.Text, txtOGTheory.Text, txtOGPractical.Text, txtFinalGrade.Text, txtGradePoint.Text)
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub clear()
        txtCreditPoint.Text = ""
        txtFinalGrade.Text = ""
        txtGradePoint.Text = ""
        txtMaxMarksT.Text = ""
        txtMaxMarksP.Text = ""
        txtOGPractical.Text = ""
        txtOGTheory.Text = ""
        txtOMPractical.Text = ""
        txtOMTheory.Text = ""
        txtSubjectCode.Text = ""
        cmbSubjectName.SelectedIndex = -1
    End Sub

    Private Sub txtGradePoint_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtGradePoint.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Sub GetGradePoint()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT GradePoint from GradeMaster where Grade=@d1"
            cmd.Parameters.AddWithValue("@d1", txtFinalGrade.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtGradePoint.Text = rdr.GetValue(0)
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub txtFinalGrade_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtFinalGrade.TextChanged
        GetGradePoint()
        If txtOMPractical.Text = "" And txtOMTheory.Text = "" Then
            txtFinalGrade.Text = ""
            txtGradePoint.Text = ""
        End If
    End Sub
    Sub Print()
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptMarksheet 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT MarksEntry.Result, SchoolInfo.S_Id, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName, Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.ContactNo AS Expr1, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Session, Student.Caste, Student.Religion, Student.SectionID, Student.Photo, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.PassPercentage, Student.Status, MarksEntry.M_Id, MarksEntry.AdmissionNo AS Expr2, MarksEntry.Session AS Expr3, MarksEntry.EntryDate, MarksEntry.Student_Class, MarksEntry_Join.MJ_Id, MarksEntry_Join.SubCode, MarksEntry_Join.MaxMarks, MarksEntry_Join.CreditHour, MarksEntry_Join.OGTheory, MarksEntry_Join.OMTheory, MarksEntry_Join.OGPractical, MarksEntry_Join.OMPractical, MarksEntry_Join.FinalGrade, MarksEntry_Join.GradePoint, MarksEntry_Join.MarksID, Subject.SubjectCode, Subject.SubjectName, Subject.Class FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN MarksEntry ON Student.AdmissionNo = MarksEntry.AdmissionNo INNER JOIN MarksEntry_Join ON MarksEntry.M_Id = MarksEntry_Join.MarksID INNER JOIN Subject ON MarksEntry_Join.SubCode = Subject.SubjectCode where M_ID='" & txtMarksID.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "SchoolInfo")
            myDA.Fill(myDS, "Student")
            myDA.Fill(myDS, "MarksEntry")
            myDA.Fill(myDS, "MarksEntry_Join")
            myDA.Fill(myDS, "Subject")
            rpt.SetDataSource(myDS)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Print()
    End Sub
    Sub Print1()
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptMarksheetPercentage 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT MMPractical,MarksEntry.Result, SchoolInfo.S_Id, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName, Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.ContactNo AS Expr1, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Session, Student.Caste, Student.Religion, Student.SectionID, Student.Photo, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.PassPercentage, Student.Status, MarksEntry.M_Id, MarksEntry.AdmissionNo AS Expr2, MarksEntry.Session AS Expr3, MarksEntry.EntryDate, MarksEntry.Student_Class, MarksEntry_Join.MJ_Id, MarksEntry_Join.SubCode, MarksEntry_Join.MaxMarks, MarksEntry_Join.CreditHour, MarksEntry_Join.OGTheory, MarksEntry_Join.OMTheory, MarksEntry_Join.OGPractical, MarksEntry_Join.OMPractical, MarksEntry_Join.FinalGrade, MarksEntry_Join.GradePoint, MarksEntry_Join.MarksID, Subject.SubjectCode, Subject.SubjectName, Subject.Class FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN MarksEntry ON Student.AdmissionNo = MarksEntry.AdmissionNo INNER JOIN MarksEntry_Join ON MarksEntry.M_Id = MarksEntry_Join.MarksID INNER JOIN Subject ON MarksEntry_Join.SubCode = Subject.SubjectCode where M_ID='" & txtMarksID.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "SchoolInfo")
            myDA.Fill(myDS, "Student")
            myDA.Fill(myDS, "MarksEntry")
            myDA.Fill(myDS, "MarksEntry_Join")
            myDA.Fill(myDS, "Subject")
            rpt.SetDataSource(myDS)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnPrint1_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint1.Click
        Print1()
    End Sub

    Private Sub txtMaxMarksP_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaxMarksP.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtOMTheory_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles txtOMTheory.Validating
        If Val(txtOMTheory.Text) > Val(txtMaxMarksT.Text) Then
            MessageBox.Show("Obtained marks must be less than or equal to max. marks", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtOMTheory.Text = ""
            txtOMTheory.Focus()
        End If
    End Sub

    Private Sub txtOMPractical_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles txtOMPractical.Validating
        If Val(txtOMPractical.Text) > Val(txtMaxMarksP.Text) Then
            MessageBox.Show("Obtained marks must be less than or equal to max. marks", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtOMPractical.Text = ""
            txtOMPractical.Focus()
        End If
    End Sub
    Sub fillSession()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM Student", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSession.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSession.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
   

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],Photo from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' and Session=@d1 and ClassName=@d2 Order by AdmissionNo", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            DataGridView1.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbSession_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSession.SelectedIndexChanged
        Try
            cmbClass.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(ClassName) FROM Student,Section,Class where Student.SectionID=Section.ID and Section.Class=Class.ClassName and Session=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
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

    Private Sub DataGridView1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        Try
            If DataGridView1.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
                txtAdmissionNo.Text = dr.Cells(0).Value.ToString()
                txtEnrollmentNo.Text = dr.Cells(1).Value.ToString()
                txtStudentName.Text = dr.Cells(5).Value.ToString()
                txtClass.Text = dr.Cells(19).Value.ToString()
                txtSchoolName.Text = dr.Cells(22).Value.ToString()
                txtSession.Text = dr.Cells(8).Value.ToString()
                FillSubject()
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
End Class
