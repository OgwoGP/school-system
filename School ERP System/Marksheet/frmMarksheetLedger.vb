Imports System.Data.SqlClient

Public Class frmMarksheetLedger
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
        cmbSession.Enabled = False
        cmbclass.Enabled = False
    End Sub

    Private Sub frmDiscount_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillschool()
    End Sub

    Private Sub cmbSession_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles cmbSession.SelectedIndexChanged
        Try
            cmbClass.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Student_Class) FROM MarksEntry where Session=@d1"
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
    Private Sub cmbSchool_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSchool.SelectedIndexChanged
        Try
            cmbSession.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Session) FROM MarksEntry"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
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

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
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
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT MMPractical, SchoolInfo.S_Id, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName, Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Caste, Student.Religion, Student.SectionID, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result, Student.PassPercentage, Student.Status, MarksEntry.M_Id, MarksEntry.Session, MarksEntry.EntryDate, MarksEntry.Student_Class, MarksEntry_Join.MJ_Id, MarksEntry_Join.SubCode, MarksEntry_Join.MaxMarks, MarksEntry_Join.CreditHour, MarksEntry_Join.OGTheory, MarksEntry_Join.OMTheory, MarksEntry_Join.OGPractical, MarksEntry_Join.OMPractical, MarksEntry_Join.FinalGrade, MarksEntry_Join.GradePoint, MarksEntry_Join.MarksID, Subject.SubjectCode, Subject.SubjectName, Subject.Class,SUM(MarksEntry_Join.OMTheory + MarksEntry_Join.OMPractical) As TotalMarks FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN MarksEntry ON Student.AdmissionNo = MarksEntry.AdmissionNo INNER JOIN MarksEntry_Join ON MarksEntry.M_Id = MarksEntry_Join.MarksID INNER JOIN Subject ON MarksEntry_Join.SubCode = Subject.SubjectCode where SchoolName=@d1 and MarksEntry.Session=@d2 and Student_Class=@d3 group by MMPractical, SchoolInfo.S_Id, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName, Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Caste, Student.Religion, Student.SectionID, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result, Student.PassPercentage, Student.Status, MarksEntry.M_Id, MarksEntry.Session, MarksEntry.EntryDate, MarksEntry.Student_Class, MarksEntry_Join.MJ_Id, MarksEntry_Join.SubCode, MarksEntry_Join.MaxMarks, MarksEntry_Join.CreditHour, MarksEntry_Join.OGTheory, MarksEntry_Join.OMTheory, MarksEntry_Join.OGPractical, MarksEntry_Join.OMPractical, MarksEntry_Join.FinalGrade, MarksEntry_Join.GradePoint, MarksEntry_Join.MarksID, Subject.SubjectCode, Subject.SubjectName, Subject.Class order by TotalMarks desc", con)
            cmd.Parameters.AddWithValue("@d1", cmbSchool.Text)
            cmd.Parameters.AddWithValue("@d2", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d3", cmbClass.Text)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("MarksheetLedger1.xml")
            Dim rpt As New rptMarksheetLedger
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
End Class
