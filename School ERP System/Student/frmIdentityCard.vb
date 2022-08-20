Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmIdentityCard
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],Photo from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],Photo from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and StudentName like '%" & txtStudentName.Text & "%' Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],Photo from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Session=@d1 and ClassName=@d2 and SectionName=@d3 Order by AdmissionNo", con)
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

    Sub Reset()
        txtStudentName.Text = ""
        cmbClass.SelectedIndex = -1
        cmbSection.SelectedIndex = -1
        cmbSession.SelectedIndex = -1
        cmbClass.Enabled = False
        cmbSection.Enabled = False
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSession()
        GetData()
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

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
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
            If Len(Trim(cmbSection.Text)) = 0 Then
                MessageBox.Show("Please select section", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSection.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select AdmissionNo,StudentName,DOB,FatherName,ClassName,SchoolName,Address,SchoolInfo.ContactNo,Logo,Photo from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Session=@d1 and ClassName=@d2 and SectionName=@d3 Order by AdmissionNo", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSection.Text)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("IdentityCard1.xml")
            Dim rpt As New rptIdentityCard
            rpt.SetDataSource(ds)
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

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select AdmissionNo,StudentName,DOB,FatherName,ClassName,SchoolName,Address,SchoolInfo.ContactNo,Logo,Photo from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and StudentName like '%" & txtStudentName.Text & "%' Order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("IdentityCard1.xml")
            Dim rpt As New rptIdentityCard
            rpt.SetDataSource(ds)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
