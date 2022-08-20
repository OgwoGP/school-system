Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmStudentRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select A_ID,RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(SubCategory) as [Sub Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID Order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub txtStudentName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStudentName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select A_ID,RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(SubCategory) as [Sub Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and StudentName like '%" & txtStudentName.Text & "%' Order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select A_ID,RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(SubCategory) as [Sub Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Session=@d1 and ClassName=@d2 and SectionName=@d3 Order by AdmissionNo", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSection.Text)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select A_ID,RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(SubCategory) as [Sub Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and AdmissionDate between @d1 and @d2 Order by AdmissionNo", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select A_ID,RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(SubCategory) as [Sub Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Classname=@d1 and Caste=@d2 Order by AdmissionNo", con)
            cmd.Parameters.AddWithValue("@d1", cmbClass1.Text)
            cmd.Parameters.AddWithValue("@d2", cmbCategory.Text)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
    Sub fillClass()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (ClassName) FROM Student,Section,Class where Student.SectionID=Section.ID and Section.Class=Class.Classname", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbClass1.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbClass1.Items.Add(drow(0).ToString())
            Next
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

    Private Sub cmbClass_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbClass.SelectedIndexChanged
        Try
            cmbSection.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(SectionName) FROM Student,Section,Class where Student.SectionID=Section.ID and Section.Class=Class.ClassName and Session=@d1 and ClassName=@d2"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            rdr = cmd.ExecuteReader()
            cmbSection.Items.Clear()
            While rdr.Read
                cmbSection.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtAdmissionNo_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtAdmissionNo.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select A_ID,RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(SubCategory) as [Sub Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and AdmissionNo like '%" & txtAdmissionNo.Text & "%' Order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        txtAdmissionNo.Text = ""
        txtStudentName.Text = ""
        cmbCategory.SelectedIndex = -1
        cmbClass.SelectedIndex = -1
        cmbClass1.SelectedIndex = -1
        cmbSection.SelectedIndex = -1
        cmbSession.SelectedIndex = -1
        txtScholarNo.Text = ""
        cmbClass.Enabled = False
        cmbSection.Enabled = False
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillClass()
        fillSession()
        GetData()
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick

        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Student Entry" Then
                    Me.Hide()
                    frmStudent.Show()
                    frmStudent.txtA_ID.Text = dr.Cells(0).Value.ToString()
                    frmStudent.txtAdmissionNo.Text = dr.Cells(1).Value.ToString()
                    frmStudent.txtAdmissionNo1.Text = dr.Cells(1).Value.ToString()
                    frmStudent.txtEnrollmentNo.Text = dr.Cells(2).Value.ToString()
                    frmStudent.txtGRNo.Text = dr.Cells(3).Value.ToString()
                    frmStudent.txtUID.Text = dr.Cells(4).Value.ToString()
                    frmStudent.dtpAdmissionDate.Text = dr.Cells(5).Value.ToString()
                    frmStudent.txtStudentName.Text = dr.Cells(6).Value.ToString()
                    If (dr.Cells(7).Value = "Male") Then
                        frmStudent.rbMale.Checked = True
                    End If
                    If (dr.Cells(7).Value = "Female") Then
                        frmStudent.rbFemale.Checked = True
                    End If
                    frmStudent.dtpDOB.Text = dr.Cells(8).Value.ToString()
                    frmStudent.cmbSession.Text = dr.Cells(9).Value.ToString()
                    frmStudent.cmbCategory.Text = dr.Cells(10).Value.ToString()
                    frmStudent.cmbSubCategory.Text = dr.Cells(11).Value.ToString()
                    frmStudent.cmbReligion.Text = dr.Cells(12).Value.ToString()
                    frmStudent.txtFathername.Text = dr.Cells(13).Value.ToString()
                    frmStudent.txtFatherContactNo.Text = dr.Cells(14).Value.ToString()
                    frmStudent.txtMotherName.Text = dr.Cells(15).Value.ToString()
                    frmStudent.txtPermanentAddress.Text = dr.Cells(16).Value.ToString()
                    frmStudent.txtTemrorayAddress.Text = dr.Cells(17).Value.ToString()
                    frmStudent.txtContactNo.Text = dr.Cells(18).Value.ToString()
                    frmStudent.txtEmailID.Text = dr.Cells(19).Value.ToString()
                    frmStudent.txtSectionID.Text = dr.Cells(20).Value.ToString()
                    frmStudent.txtClass.Text = dr.Cells(21).Value.ToString()
                    frmStudent.txtSection.Text = dr.Cells(22).Value.ToString()
                    frmStudent.txtSchoolID.Text = dr.Cells(23).Value.ToString()
                    frmStudent.cmbSchoolName.Text = dr.Cells(24).Value.ToString()
                    frmStudent.txtLastSchoolAttended.Text = dr.Cells(25).Value.ToString()
                    frmStudent.cmbResult.Text = dr.Cells(26).Value.ToString()
                    frmStudent.txtPercentage.Text = dr.Cells(27).Value.ToString()
                    frmStudent.txtNationality.Text = dr.Cells(28).Value.ToString()
                    frmStudent.cmbStatus.Text = dr.Cells(29).Value.ToString()
                    frmStudent.cmbHouse.Text = dr.Cells(30).Value.ToString()
                    frmStudent.txtSSSM_ID.Text = dr.Cells(31).Value.ToString()
                    frmStudent.txtAccountNo.Text = dr.Cells(32).Value.ToString()
                    frmStudent.txtAccountName.Text = dr.Cells(33).Value.ToString()
                    frmStudent.txtBank.Text = dr.Cells(34).Value.ToString()
                    frmStudent.txtBranch.Text = dr.Cells(35).Value.ToString()
                    frmStudent.txtIFSCcode.Text = dr.Cells(36).Value.ToString()

                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = New SqlCommand("SELECT DocID,DocName from Student,Document,StudentDocSubmitted where Student.AdmissionNo=StudentDocSubmitted.AdmissionNo and Document.Doc_ID=StudentDocSubmitted.DocID and Student.AdmissionNo=@d1", con)
                    cmd.Parameters.AddWithValue("@d1", dr.Cells(1).Value)
                    rdr = cmd.ExecuteReader()
                    frmStudent.ListView1.Items.Clear()
                    While rdr.Read()
                        Dim lst As New ListViewItem()
                        lst.SubItems.Add(rdr(0))
                        lst.SubItems.Add(rdr(1).ToString().Trim())
                        frmStudent.ListView1.Items.Add(lst)
                    End While
                    con.Close()
                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = New SqlCommand("SELECT Photo from Student where AdmissionNo=@d1", con)
                    cmd.Parameters.AddWithValue("@d1", dr.Cells(1).Value)
                    rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        Dim data As Byte() = DirectCast(rdr(0), Byte())
                        Dim ms As New MemoryStream(data)
                        frmStudent.Picture.Image = Image.FromStream(ms)
                    End If
                    con.Close()
                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = New SqlCommand("SELECT distinct RTRIM(Subject.SubjectCode),RTRIM(SubjectName) from Student,Section,Class,Subject,Student_Subjects where  Student.SectionID = Section.Id and Section.Class = Class.ClassName and Student.AdmissionNo = Student_Subjects.AdmissionNo and Class.ClassName=Subject.Class and Student_Subjects.SubjectCode = Subject.SubjectCode and Student.AdmissionNo=@d1 order by 2", con)
                    cmd.Parameters.AddWithValue("@d1", dr.Cells(1).Value)
                    rdr = cmd.ExecuteReader()
                    frmStudent.dgw.Rows.Clear()
                    While rdr.Read()
                        frmStudent.dgw.Rows.Add(rdr(0), rdr(1))
                    End While
                    con.Close()
                    frmStudent.btnDelete.Enabled = True
                    frmStudent.btnUpdate.Enabled = True
                    frmStudent.btnSave.Enabled = False
                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = con.CreateCommand()
                    cmd.CommandText = "SELECT Session FROM Student where AdmissionNo=@d1"
                    cmd.Parameters.AddWithValue("@d1", dr.Cells(1).Value)
                    rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        frmStudent.cmbSession.Text = rdr.GetValue(0)
                    End If
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    If con.State = ConnectionState.Open Then
                        con.Close()
                    End If
                    lblSet.Text = ""
                    frmStudent.FillSubject()
                End If
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

    Private Sub txtGRNo_TextChanged(sender As Object, e As EventArgs) Handles txtScholarNo.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select A_ID,RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(SubCategory) as [Sub Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and GRNo like '%" & txtScholarNo.Text & "%' Order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub
End Class
