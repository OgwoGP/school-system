Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmActiveStudentRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' and StudentName like '%" & txtStudentName.Text & "%' Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' and Session=@d1 and ClassName=@d2 and SectionName=@d3 Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' and AdmissionDate between @d1 and @d2 Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' and Classname=@d1 and Caste=@d2 Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' and AdmissionNo like '%" & txtAdmissionNo.Text & "%' Order by AdmissionNo", con)
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
        txtGRNo.Text = ""
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
                If lblSet.Text = "Hosteler Entry" Then
                    Me.Hide()
                    frmHosteler.Show()
                    frmHosteler.txtAdmissionNo.Text = dr.Cells(0).Value.ToString()
                    frmHosteler.txtStudentName.Text = dr.Cells(5).Value.ToString()
                    frmHosteler.txtClass.Text = dr.Cells(19).Value.ToString()
                    frmHosteler.txtSection.Text = dr.Cells(20).Value.ToString()
                    frmHosteler.txtSchoolName.Text = dr.Cells(22).Value.ToString()
                End If
                If lblSet.Text = "Bus Holder Entry" Then
                    Me.Hide()
                    frmBusCardHolder_Student.Show()
                    frmBusCardHolder_Student.txtAdmissionNo.Text = dr.Cells(0).Value.ToString()
                    frmBusCardHolder_Student.txtStudentName.Text = dr.Cells(5).Value.ToString()
                    frmBusCardHolder_Student.txtClass.Text = dr.Cells(19).Value.ToString()
                    frmBusCardHolder_Student.txtSection.Text = dr.Cells(20).Value.ToString()
                    frmBusCardHolder_Student.txtSchoolName.Text = dr.Cells(22).Value.ToString()
                End If
                If lblSet.Text = "Course Fee Payment" Then
                    Me.Hide()
                    ' frmClassFeePayment.dgw.Rows.Clear()
                    frmClassFeePayment.Show()
                    frmClassFeePayment.txtAdmissionNo.Text = dr.Cells(0).Value.ToString()
                    frmClassFeePayment.txtEnrollmentNo.Text = dr.Cells(1).Value.ToString()
                    frmClassFeePayment.txtStudentName.Text = dr.Cells(5).Value.ToString()
                    frmClassFeePayment.txtClass.Text = dr.Cells(19).Value.ToString()
                    frmClassFeePayment.txtSection.Text = dr.Cells(20).Value.ToString()
                    frmClassFeePayment.txtSchoolName.Text = dr.Cells(22).Value.ToString()
                    frmClassFeePayment.txtSession.Text = dr.Cells(8).Value.ToString()
                    frmClassFeePayment.FillFee()
                    frmClassFeePayment.btnStudentSelection.Enabled = False
                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = con.CreateCommand()
                    cmd.CommandText = "SELECT RTRIM(ContactNo) FROM Student where AdmissionNo=@d1"
                    cmd.Parameters.AddWithValue("@d1", dr.Cells(0).Value)
                    rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        frmClassFeePayment.txtContactNo.Text = rdr.GetValue(0)
                    End If
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    If con.State = ConnectionState.Open Then
                        con.Close()
                    End If
                End If
                If lblSet.Text = "Marks Entry" Then
                    Me.Hide()
                    frmMarksEntry.Show()
                    frmMarksEntry.txtAdmissionNo.Text = dr.Cells(0).Value.ToString()
                    frmMarksEntry.txtEnrollmentNo.Text = dr.Cells(1).Value.ToString()
                    frmMarksEntry.txtStudentName.Text = dr.Cells(5).Value.ToString()
                    frmMarksEntry.txtClass.Text = dr.Cells(19).Value.ToString()
                    frmMarksEntry.txtSchoolName.Text = dr.Cells(22).Value.ToString()
                    frmMarksEntry.txtSession.Text = dr.Cells(8).Value.ToString()
                    frmMarksEntry.FillSubject()
                End If
                If lblSet.Text = "Book Issue" Then
                    Me.Hide()
                    frmBookIssue.Show()
                    frmBookIssue.txtAdmissionNo.Text = dr.Cells(0).Value.ToString()
                    frmBookIssue.txtStudentName.Text = dr.Cells(5).Value.ToString()
                    frmBookIssue.txtClass.Text = dr.Cells(19).Value.ToString()
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

    Private Sub txtGRNo_TextChanged(sender As Object, e As EventArgs) Handles txtGRNo.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.Status='Active' and GRNo like '%" & txtGRNo.Text & "%' Order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
