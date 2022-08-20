Imports System.Data.SqlClient
Imports System.IO

Public Class frmStudent
    Dim s1 As String
    Dim st2 As String
    Dim Photoname As String = ""
    Dim IsImageChanged As Boolean = False
    Sub Reset()
        txtAdmissionNo.Text = ""
        txtEnrollmentNo.Text = ""
        txtClass.Text = ""
        txtContactNo.Text = ""
        txtDocID.Text = ""
        txtGRNo.Text = ""
        txtUID.Text = ""
        txtLastSchoolAttended.Text = ""
        cmbResult.SelectedIndex = -1
        txtPercentage.Text = ""
        txtEmailID.Text = ""
        txtEnrollmentNo.Text = ""
        txtFatherContactNo.Text = ""
        txtFathername.Text = ""
        txtMotherName.Text = ""
        txtNationality.Text = "Pakistani"
        txtPermanentAddress.Text = ""
        txtSection.Text = ""
        cmbSession.SelectedIndex = -1
        txtStudentName.Text = ""
        txtTemrorayAddress.Text = ""
        cmbSchoolName.SelectedIndex = -1
        cmbCategory.SelectedIndex = -1
        cmbDocumentsSubmitted.SelectedIndex = -1
        cmbHouse.SelectedIndex = -1
        cmbReligion.SelectedIndex = -1
        cmbStatus.SelectedIndex = 0
        dtpAdmissionDate.Value = Today
        dtpDOB.Value = Today
        txtSSSM_ID.Text = ""
        txtAccountNo.Text = ""
        txtAccountName.Text = ""
        txtBank.Text = ""
        txtBranch.Text = ""
        txtIFSCcode.Text = ""
        ListView1.Items.Clear()
        rbFemale.Checked = False
        rbMale.Checked = False
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        btnSave.Enabled = True
        txtPercentage.ReadOnly = True
        Picture.Image = My.Resources.photo
        btnRemoveSubject.Enabled = False
        cmbSubject.Items.Clear()
        cmbSubject.SelectedIndex = -1
        cmbSubCategory.Enabled = False
        cmbSubCategory.Text = ""
        dgw.Rows.Clear()
        txtEnrollmentNo.Text = DateTime.Now.Day.ToString() + "" + DateTime.Now.Month.ToString() + "" + DateTime.Now.Year.ToString() + "" + DateTime.Now.Hour.ToString() + "" + DateTime.Now.Second.ToString() + "" + DateTime.Now.Millisecond.ToString()
        auto()
    End Sub


    Private Sub Browse_Click_1(sender As System.Object, e As System.EventArgs) Handles Browse.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Picture.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub BRemove_Click(sender As System.Object, e As System.EventArgs) Handles BRemove.Click
        Picture.Image = My.Resources.photo
    End Sub

    Private Sub BStartCapture_Click(sender As System.Object, e As System.EventArgs) Handles BStartCapture.Click
        Dim k As New frmCamera
        k.ShowDialog()
        If TempFileNames2.Length > 0 Then

            Picture.Image = Image.FromFile(TempFileNames2)
            Photoname = TempFileNames2
            IsImageChanged = True
        End If
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub
    Private Function GenerateID() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 (A_ID) FROM Student ORDER BY A_ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("A_ID")
            End If
            rdr.Close()
            ' Increase the ID by 1
            value += 1
            ' Because incrementing a string with an integer removes 0's
            ' we need to replace them. If necessary.
            If value <= 9 Then 'Value is between 0 and 10
                value = "000" & value
            ElseIf value <= 99 Then 'Value is between 9 and 100
                value = "00" & value
            ElseIf value <= 999 Then 'Value is between 999 and 1000
                value = "0" & value
            End If
        Catch ex As Exception
            ' If an error occurs, check the connection state and close it if necessary.
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            value = "0000"
        End Try
        Return value
    End Function
    Public Sub auto()
        Try
            txtA_ID.Text = GenerateID()
            txtAdmissionNo.Text = "A-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Try
            If Len(Trim(cmbSchoolName.Text)) = 0 Then
                MessageBox.Show("Please select school name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSchoolName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtAdmissionNo.Text)) = 0 Then
                MessageBox.Show("Please enter admission no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAdmissionNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtStudentName.Text)) = 0 Then
                MessageBox.Show("Please enter student name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtStudentName.Focus()
                Exit Sub
            End If
            If rbMale.Checked = False And rbFemale.Checked = False Then
                MessageBox.Show("Please check gender", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If Len(Trim(cmbSession.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbReligion.Text)) = 0 Then
                MessageBox.Show("Please select religion", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbReligion.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbCategory.Text)) = 0 Then
                MessageBox.Show("Please select category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbCategory.Focus()
                Exit Sub
            End If
            If Len(Trim(txtFathername.Text)) = 0 Then
                MessageBox.Show("Please enter father's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFathername.Focus()
                Exit Sub
            End If
            If Len(Trim(txtFatherContactNo.Text)) = 0 Then
                MessageBox.Show("Please enter father's contact no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFatherContactNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtMotherName.Text)) = 0 Then
                MessageBox.Show("Please enter mother's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtMotherName.Focus()
                Exit Sub
            End If
            If ListView1.Items.Count = 0 Then
                MessageBox.Show("Please add documents submitted to listview", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbDocumentsSubmitted.Focus()
                Exit Sub
            End If
            If Len(Trim(txtTemrorayAddress.Text)) = 0 Then
                MessageBox.Show("Please enter temporary address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtTemrorayAddress.Focus()
                Exit Sub
            End If
            If Len(Trim(txtPermanentAddress.Text)) = 0 Then
                MessageBox.Show("Please enter permanent address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPermanentAddress.Focus()
                Exit Sub
            End If
            If Len(Trim(txtContactNo.Text)) = 0 Then
                MessageBox.Show("Please enter contact no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtContactNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtClass.Text)) = 0 Then
                MessageBox.Show("Please retrieve class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtClass.Focus()
                Exit Sub
            End If
            If Len(Trim(txtSection.Text)) = 0 Then
                MessageBox.Show("Please retrieve section", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSection.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbStatus.Text)) = 0 Then
                MessageBox.Show("Please select status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbStatus.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select AdmissionNo from Student where AdmissionNo=@d1"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If rdr.Read() Then
                MessageBox.Show("Admission No. Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                txtAdmissionNo.Text = ""
                txtAdmissionNo.Focus()
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            If rbMale.Checked = True Then
                s1 = rbMale.Text
            End If
            If rbFemale.Checked = True Then
                s1 = rbFemale.Text
            End If
            auto()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Student(AdmissionNo, EnrollmentNo, GRNo, UID, StudentName, FatherName, MotherName, FatherCN, PermanentAddress, TemporaryAddress, ContactNo, EmailID, DOB, Gender, AdmissionDate, Session, Caste, Religion,SectionID, Nationality, SchoolID, LastSchoolAttended, Result, PassPercentage, Status,Photo,House,SSSM_ID,AccountNo,Accountname,Bank,Branch,IFSCCode,A_ID,SubCategory) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17,@d18,@d19,@d20,@d21,@d22,@d23,@d24,@d25,@d26,@d27,@d28,@d29,@d30,@d31,@d32,@d33," & Val(txtA_ID.Text) & ",@d34)"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d2", txtEnrollmentNo.Text)
            cmd.Parameters.AddWithValue("@d3", txtGRNo.Text)
            cmd.Parameters.AddWithValue("@d4", txtUID.Text)
            cmd.Parameters.AddWithValue("@d5", txtStudentName.Text)
            cmd.Parameters.AddWithValue("@d6", txtFathername.Text)
            cmd.Parameters.AddWithValue("@d7", txtMotherName.Text)
            cmd.Parameters.AddWithValue("@d8", txtFatherContactNo.Text)
            cmd.Parameters.AddWithValue("@d9", txtPermanentAddress.Text)
            cmd.Parameters.AddWithValue("@d10", txtTemrorayAddress.Text)
            cmd.Parameters.AddWithValue("@d11", txtContactNo.Text)
            cmd.Parameters.AddWithValue("@d12", txtEmailID.Text)
            cmd.Parameters.AddWithValue("@d13", dtpDOB.Value.Date)
            cmd.Parameters.AddWithValue("@d14", s1)
            cmd.Parameters.AddWithValue("@d15", dtpAdmissionDate.Value.date)
            cmd.Parameters.AddWithValue("@d16", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d17", cmbCategory.Text)
            cmd.Parameters.AddWithValue("@d18", cmbReligion.Text)
            cmd.Parameters.AddWithValue("@d19", Val(txtSectionID.Text))
            cmd.Parameters.AddWithValue("@d20", txtNationality.Text)
            cmd.Parameters.AddWithValue("@d21", Val(txtSchoolID.Text))
            cmd.Parameters.AddWithValue("@d22", txtLastSchoolAttended.Text)
            cmd.Parameters.AddWithValue("@d23", cmbResult.Text)
            cmd.Parameters.AddWithValue("@d24", txtPercentage.Text)
            cmd.Parameters.AddWithValue("@d25", cmbStatus.Text)
            cmd.Parameters.AddWithValue("@d27", cmbHouse.Text)
            cmd.Parameters.AddWithValue("@d28", txtSSSM_ID.Text)
            cmd.Parameters.AddWithValue("@d29", txtAccountNo.Text)
            cmd.Parameters.AddWithValue("@d30", txtAccountName.Text)
            cmd.Parameters.AddWithValue("@d31", txtBank.Text)
            cmd.Parameters.AddWithValue("@d32", txtBranch.Text)
            cmd.Parameters.AddWithValue("@d33", txtIFSCcode.Text)
            cmd.Parameters.AddWithValue("@d34", cmbSubCategory.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(Picture.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New SqlParameter("@d26", SqlDbType.Image)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.ExecuteNonQuery()
            For i = 0 To ListView1.Items.Count - 1
                con = New SqlConnection(cs)
                Dim cd As String = "insert Into StudentDocSubmitted(AdmissionNo,DocID) VALUES (@d1,@d2)"
                cmd = New SqlCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
                cmd.Parameters.AddWithValue("@d2", CInt(ListView1.Items(i).SubItems(1).Text))
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    con = New SqlConnection(cs)
                    Dim cd As String = "insert Into Student_Subjects(SubjectCode,AdmissionNo) VALUES (@d1,@d2)"
                    cmd = New SqlCommand(cd)
                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value.ToString())
                    cmd.Parameters.AddWithValue("@d2", txtAdmissionNo.Text)
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                End If
            Next
            con = New SqlConnection(cs)
            con.Open()
            Dim cb3 As String = "insert into Discount(AdmissionNo,FeeType,Discount) VALUES (@d1,'Class',0.00)"
            cmd = New SqlCommand(cb3)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb4 As String = "insert into Discount(AdmissionNo,FeeType,Discount) VALUES (@d1,'Hostel',0.00)"
            cmd = New SqlCommand(cb4)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb5 As String = "insert into Discount(AdmissionNo,FeeType,Discount) VALUES (@d1,'Bus',0.00)"
            cmd = New SqlCommand(cb5)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            If CheckForInternetConnection() = True Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ctn As String = "select RTRIM(APIURL) from SMSSetting where IsDefault='Yes' and IsEnabled='Yes'"
                cmd = New SqlCommand(ctn)
                cmd.Connection = con
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    st2 = rdr.GetValue(0)
                    Dim st3 As String = "Hello Dear, " & txtStudentName.Text & " your admission is successful having admission no '" & txtAdmissionNo.Text & "'"
                    SMSFunc(txtContactNo.Text, st3, st2)
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                End If
            End If
            LogFunc(lblUser.Text, "added new student '" & txtStudentName.Text & "' has admission no. '" & txtAdmissionNo.Text & "'")
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False

            If con.State = ConnectionState.Open Then
                con.Close()
            End If
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
    Public Sub DeleteRecord()
        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim clB As String = "select Student.AdmissionNo from Student,Attendance where Student.AdmissionNo=Attendance.AdmissionNo and Student.AdmissionNo=@d1"
            cmd = New SqlCommand(clB)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
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
            Dim sql As String = "select Student.AdmissionNo from Student,MarksEntry where Student.AdmissionNo=MarksEntry.AdmissionNo and Student.AdmissionNo=@d1"
            cmd = New SqlCommand(sql)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
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
            Dim clZ As String = "select Student.AdmissionNo from Student,BookIssue_Student where Student.AdmissionNo=BookIssue_Student.AdmissionNo and Student.AdmissionNo=@d1"
            cmd = New SqlCommand(clZ)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Book Issue", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If

            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "select Student.AdmissionNo from Student,Hosteler where Student.AdmissionNo=Hosteler.AdmissionNo and Student.AdmissionNo=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Hosteler Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cl2 As String = "select Student.AdmissionNo from Student,BusCardHolder_Student where Student.AdmissionNo=BusCardHolder_Student.AdmissionNo and Student.AdmissionNo=@d1"
            cmd = New SqlCommand(cl2)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Bus Card Holder[Student] Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cl1 As String = "select Student.AdmissionNo from Student,CourseFeePayment where Student.AdmissionNo=CourseFeePayment.AdmissionNo and Student.AdmissionNo=@d1"
            cmd = New SqlCommand(cl1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Course Fee Payment Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Student where AdmissionNo=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the student '" & txtStudentName.Text & "' has admission no. '" & txtAdmissionNo.Text & "'")
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
            If Len(Trim(cmbSchoolName.Text)) = 0 Then
                MessageBox.Show("Please select school name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSchoolName.Focus()
                Exit Sub
            End If
            If Len(Trim(txtAdmissionNo.Text)) = 0 Then
                MessageBox.Show("Please enter admission no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAdmissionNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtStudentName.Text)) = 0 Then
                MessageBox.Show("Please enter student name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtStudentName.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbSession.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If rbMale.Checked = False And rbFemale.Checked = False Then
                MessageBox.Show("Please check gender", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If Len(Trim(cmbReligion.Text)) = 0 Then
                MessageBox.Show("Please select religion", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbReligion.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbCategory.Text)) = 0 Then
                MessageBox.Show("Please select category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbCategory.Focus()
                Exit Sub
            End If
            If Len(Trim(txtFathername.Text)) = 0 Then
                MessageBox.Show("Please enter father's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFathername.Focus()
                Exit Sub
            End If
            If Len(Trim(txtFatherContactNo.Text)) = 0 Then
                MessageBox.Show("Please enter father's contact no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFatherContactNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtMotherName.Text)) = 0 Then
                MessageBox.Show("Please enter mother's name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtMotherName.Focus()
                Exit Sub
            End If
            If ListView1.Items.Count = 0 Then
                MessageBox.Show("Please add documents submitted to listview", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbDocumentsSubmitted.Focus()
                Exit Sub
            End If
            If Len(Trim(txtTemrorayAddress.Text)) = 0 Then
                MessageBox.Show("Please enter temporary address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtTemrorayAddress.Focus()
                Exit Sub
            End If
            If Len(Trim(txtPermanentAddress.Text)) = 0 Then
                MessageBox.Show("Please enter permanent address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPermanentAddress.Focus()
                Exit Sub
            End If
            If Len(Trim(txtContactNo.Text)) = 0 Then
                MessageBox.Show("Please enter contact no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtContactNo.Focus()
                Exit Sub
            End If
            If Len(Trim(txtClass.Text)) = 0 Then
                MessageBox.Show("Please retrieve class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtClass.Focus()
                Exit Sub
            End If
            If Len(Trim(txtSection.Text)) = 0 Then
                MessageBox.Show("Please retrieve section", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSection.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbStatus.Text)) = 0 Then
                MessageBox.Show("Please select status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbStatus.Focus()
                Exit Sub
            End If

            If rbMale.Checked = True Then
                s1 = rbMale.Text
            End If
            If rbFemale.Checked = True Then
                s1 = rbFemale.Text
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "update StudentDocSubmitted set AdmissionNo=@d1 where AdmissionNo=@d2"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d2", txtAdmissionNo1.Text)
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Student set  EnrollmentNo=@d2, GRNo=@d3, UID=@d4, StudentName=@d5, FatherName=@d6, MotherName=@d7, FatherCN=@d8, PermanentAddress=@d9, TemporaryAddress=@d10, ContactNo=@d11, EmailID=@d12, DOB=@d13, Gender=@d14, AdmissionDate=@d15, Session=@d16, Caste=@d17, Religion=@d18,SectionID=@d19, Nationality=@d20, SchoolID=@d21, LastSchoolAttended=@d22, Result=@d23, PassPercentage=@d24, Status=@d25,Photo=@d26,AdmissionNo=@d27,House=@d28,SSSM_ID=@d29,AccountNo=@d30,Accountname=@d31,Bank=@d32,Branch=@d33,IFSCCode=@d34,SubCategory=@d35 where AdmissionNo=@d1"
            cmd = New SqlCommand(cb)
            cmd.Connection = con

            cmd.Parameters.AddWithValue("@d2", txtEnrollmentNo.Text)
            cmd.Parameters.AddWithValue("@d3", txtGRNo.Text)
            cmd.Parameters.AddWithValue("@d4", txtUID.Text)
            cmd.Parameters.AddWithValue("@d5", txtStudentName.Text)
            cmd.Parameters.AddWithValue("@d6", txtFathername.Text)
            cmd.Parameters.AddWithValue("@d7", txtMotherName.Text)
            cmd.Parameters.AddWithValue("@d8", txtFatherContactNo.Text)
            cmd.Parameters.AddWithValue("@d9", txtPermanentAddress.Text)
            cmd.Parameters.AddWithValue("@d10", txtTemrorayAddress.Text)
            cmd.Parameters.AddWithValue("@d11", txtContactNo.Text)
            cmd.Parameters.AddWithValue("@d12", txtEmailID.Text)
            cmd.Parameters.AddWithValue("@d13", dtpDOB.Value.Date)
            cmd.Parameters.AddWithValue("@d14", s1)
            cmd.Parameters.AddWithValue("@d15", dtpAdmissionDate.Value.date)
            cmd.Parameters.AddWithValue("@d16", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d17", cmbCategory.Text)
            cmd.Parameters.AddWithValue("@d18", cmbReligion.Text)
            cmd.Parameters.AddWithValue("@d19", Val(txtSectionID.Text))
            cmd.Parameters.AddWithValue("@d20", txtNationality.Text)
            cmd.Parameters.AddWithValue("@d21", Val(txtSchoolID.Text))
            cmd.Parameters.AddWithValue("@d22", txtLastSchoolAttended.Text)
            cmd.Parameters.AddWithValue("@d23", cmbResult.Text)
            cmd.Parameters.AddWithValue("@d24", txtPercentage.Text)
            cmd.Parameters.AddWithValue("@d25", cmbStatus.Text)
            cmd.Parameters.AddWithValue("@d28", cmbHouse.Text)
            cmd.Parameters.AddWithValue("@d29", txtSSSM_ID.Text)
            cmd.Parameters.AddWithValue("@d30", txtAccountNo.Text)
            cmd.Parameters.AddWithValue("@d31", txtAccountName.Text)
            cmd.Parameters.AddWithValue("@d32", txtBank.Text)
            cmd.Parameters.AddWithValue("@d33", txtBranch.Text)
            cmd.Parameters.AddWithValue("@d34", txtIFSCcode.Text)
            cmd.Parameters.AddWithValue("@d35", cmbSubCategory.Text)
            Dim ms As New MemoryStream()
            Dim bmpImage As New Bitmap(Picture.Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Dim data As Byte() = ms.GetBuffer()
            Dim p As New SqlParameter("@d26", SqlDbType.Image)
            p.Value = data
            cmd.Parameters.Add(p)
            cmd.Parameters.AddWithValue("@d27", txtAdmissionNo1.Text)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from StudentDocSubmitted where AdmissionNo=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            For i = 0 To ListView1.Items.Count - 1
                con = New SqlConnection(cs)
                Dim cd As String = "insert Into StudentDocSubmitted(AdmissionNo,DocID) VALUES (@d1,@d2)"
                cmd = New SqlCommand(cd)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
                cmd.Parameters.AddWithValue("@d2", CInt(ListView1.Items(i).SubItems(1).Text))
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next
            con = New SqlConnection(cs)
            con.Open()
            Dim cqX As String = "delete from Student_Subjects where AdmissionNo=@d1"
            cmd = New SqlCommand(cqX)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    con = New SqlConnection(cs)
                    Dim cd As String = "insert Into Student_Subjects(SubjectCode,AdmissionNo) VALUES (@d1,@d2)"
                    cmd = New SqlCommand(cd)
                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value.ToString())
                    cmd.Parameters.AddWithValue("@d2", txtAdmissionNo.Text)
                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                End If
            Next
            LogFunc(lblUser.Text, "updated student '" & txtStudentName.Text & "' has admission no. '" & txtAdmissionNo.Text & "'")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False

            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        Try
            If Len(Trim(cmbDocumentsSubmitted.Text)) = 0 Then
                MessageBox.Show("Please select doc submitted", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbDocumentsSubmitted.Focus()
                Exit Sub
            End If
            Dim temp As Integer
            temp = ListView1.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(txtDocID.Text)
                lst.SubItems.Add(cmbDocumentsSubmitted.Text)
                ListView1.Items.Add(lst)
                i = i + 1
                cmbDocumentsSubmitted.SelectedIndex = -1
                Exit Sub
            End If
            For j = 0 To temp - 1
                If (ListView1.Items(j).SubItems(2).Text = cmbDocumentsSubmitted.Text) Then
                    MessageBox.Show("Doc already added", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)
            lst1.SubItems.Add(txtDocID.Text)
            lst1.SubItems.Add(cmbDocumentsSubmitted.Text)
            ListView1.Items.Add(lst1)
            k = k + 1
            cmbDocumentsSubmitted.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        Try
            If (ListView1.SelectedItems.Count > 0) Then
                Dim itmCnt, i, t As Integer
                ListView1.FocusedItem.Remove()
                itmCnt = ListView1.Items.Count
                t = 1
                For i = 1 To itmCnt + 1

                    'Dim lst1 As New ListViewItem(i)
                    'ListView1.Items(i).SubItems(0).Text = t
                    t = t + 1
                Next
                btnRemove.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        btnRemove.Enabled = True
    End Sub
    Sub FillSubject()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT distinct RTRIM(SubjectName) FROM Subject where Class=@d1 order by 1", con)
            cmd.Parameters.AddWithValue("@d1", txtClass.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            cmbSubject.Items.Clear()
            While rdr.Read()
                cmbSubject.Items.Add(rdr.GetValue(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmSectionRecord.Reset()
        frmSectionRecord.lblSet.Text = "Student Entry"
        frmSectionRecord.ShowDialog()
    End Sub

    Private Sub cmbResult_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbResult.SelectedIndexChanged
        If cmbResult.Text = "Pass" Then
            txtPercentage.ReadOnly = False
        End If
        If cmbResult.Text = "Fail" Then
            txtPercentage.ReadOnly = True
        End If
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cmbSchoolName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSchoolName.SelectedIndexChanged
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
    Sub fillHouse()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(Housename) FROM House order by 1", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbHouse.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbHouse.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillDocs()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (DocName) FROM Document", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbDocumentsSubmitted.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbDocumentsSubmitted.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillSession()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM Session_Master", con)
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
    Private Sub frmStudent_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSchool()
        fillDocs()
        fillSession()
        fillHouse()
    End Sub

    Private Sub cmbDocumentsSubmitted_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbDocumentsSubmitted.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Doc_ID FROM Document where DocName=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbDocumentsSubmitted.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtDocID.Text = rdr.GetValue(0)
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

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmStudentRecord.Reset()
        frmStudentRecord.lblSet.Text = "Student Entry"
        frmStudentRecord.ShowDialog()
    End Sub

    Private Sub txtPercentage_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPercentage.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtPercentage.Text
            Dim selectionStart = Me.txtPercentage.SelectionStart
            Dim selectionLength = Me.txtPercentage.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
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

    Private Sub btnAddSubject_Click(sender As System.Object, e As System.EventArgs) Handles btnAddSubject.Click
        Try
            If Len(Trim(cmbSubject.Text)) = 0 Then
                MessageBox.Show("Please select subject", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSubject.Focus()
                Exit Sub
            End If
            For Each dr As DataGridViewRow In dgw.Rows
                If cmbSubject.Text = dr.Cells(1).Value.ToString() Then
                    MessageBox.Show("Same subject is already added to grid", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    cmbSubject.SelectedIndex = -1
                    cmbSubject.Focus()
                    Exit Sub
                End If
            Next
            dgw.Rows.Add(txtSubjectCode.Text, cmbSubject.Text)
            cmbSubject.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbSubject_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSubject.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(SubjectCode) FROM Subject where SubjectName=@d1", con)
            cmd.Parameters.AddWithValue("@d1", cmbSubject.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.Read() Then
                txtSubjectCode.Text = rdr.GetValue(0).ToString()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRemoveSubject_Click(sender As System.Object, e As System.EventArgs) Handles btnRemoveSubject.Click
        For Each row As DataGridViewRow In dgw.SelectedRows
            dgw.Rows.Remove(row)
        Next
        btnRemoveSubject.Enabled = False
    End Sub

    Private Sub dgw_MouseClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        btnRemoveSubject.Enabled = True
    End Sub

    Private Sub cmbCategory_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RTRIM(SubCategory) FROM Student where Caste=@d1"
            cmd.Parameters.AddWithValue("@d1", cmbCategory.Text)
            rdr = cmd.ExecuteReader()
            cmbSubCategory.Items.Clear()
            cmbSubCategory.Enabled = True
            While rdr.Read()
                cmbSubCategory.Items.Add(rdr(0))
            End While
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

    Private Sub txtNationality_TextChanged(sender As Object, e As EventArgs) Handles txtNationality.TextChanged

    End Sub
End Class
