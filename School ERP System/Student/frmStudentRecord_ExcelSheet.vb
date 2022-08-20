Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmStudentRecord_ExcelSheet
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
            txtANo.Text = "A-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and StudentName like '%" & txtStudentName.Text & "%' Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Session=@d1 and ClassName=@d2 and SectionName=@d3 Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and AdmissionDate between @d1 and @d2 Order by AdmissionNo", con)
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
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) as [House],RTRIM(SSSM_ID) as [SSSM ID],RTRIM(AccountNo) as [Account No.],RTRIM(AccountName) as [Account Name],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCCode) as [IFSC Code] from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and AdmissionNo like '%" & txtAdmissionNo.Text & "%' Order by AdmissionNo", con)
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
        cmbClass.SelectedIndex = -1
        cmbSection.SelectedIndex = -1
        cmbSession.SelectedIndex = -1
        txtGRNo.Text = ""
        cmbClass.Enabled = False
        cmbSection.Enabled = False
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        btnUpdate.Enabled = False
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

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub txtGRNo_TextChanged(sender As Object, e As EventArgs) Handles txtGRNo.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AdmissionNo) as [Admission No],RTRIM(EnrollmentNo) as [Enrollment No], RTRIM(GRNo) as [Scholar No.], RTRIM(UID) as [UID],Convert(DateTime,AdmissionDate,103) as [Admission Date], RTRIM(StudentName) as [Student Name], RTRIM(Gender) as [Gender],Convert(DateTime,DOB,103) as [DOB],RTRIM(Session) as Session,RTRIM(Caste) as [Category],RTRIM(Religion) as [Religion],RTRIM(FatherName) as [Father's Name], RTRIM(FatherCN) as [Father's CN], RTRIM(MotherName) as [Mother's Name],RTRIM(PermanentAddress) as [Permanent Address], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(Student.ContactNo) as [Contact No], RTRIM(EmailID) as [Email ID],RTRIM(SectionID) as [Section ID],RTRIM(ClassName) as [Class],RTRIM(SectionName) as Section,RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(LastSchoolAttended) as [Last School Attended],RTRIM(Result) as [Result],RTRIM(PassPerCentage) as [If Pass then %],RTRIM(Nationality) as [Nationality],RTRIM(Status) as [Status],RTRIM(House) from Student,Class,Section,SchoolInfo where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and GRNo like '%" & txtGRNo.Text & "%' Order by AdmissionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Student")
            dgw.DataSource = ds.Tables("Student").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles btnImportExcel.Click
        Try
            Dim OpenFileDialog As New OpenFileDialog
            OpenFileDialog.Filter = "Excel Files | *.xlsx; *.xls;| All Files (*.*)| *.*"
            If OpenFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK AndAlso OpenFileDialog.FileName <> "" Then
                Cursor = Cursors.WaitCursor
                Timer1.Enabled = True
                Dim Pathname As String = OpenFileDialog.FileName
                Dim MyConnection As System.Data.OleDb.OleDbConnection
                Dim DtSet As System.Data.DataSet
                Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Pathname + ";Extended Properties=Excel 8.0;")
                MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
                MyConnection.Open()
                DtSet = New System.Data.DataSet
                MyCommand.Fill(DtSet)
                dgw.Visible = True
                dgw.DataSource = DtSet.Tables(0)
                btnUpdate.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Try
            If dgw.RowCount = Nothing Then
                MessageBox.Show("Sorry nothing to save.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    SqlConnection.ClearAllPools()
                    con = New SqlConnection(cs)
                    con.Open()
                    Dim ct As String = "select AdmissionNo from Student Where AdmissionNo=@d1"
                    cmd = New SqlCommand(ct)
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Connection = con
                    rdr = cmd.ExecuteReader()
                    If Not rdr.Read() Then
                        Cursor = Cursors.WaitCursor
                        Timer1.Enabled = True
                        SqlConnection.ClearAllPools()
                        auto()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb As String = "Insert Into Student(AdmissionNo,EnrollmentNo,GRNo,UID,AdmissionDate,StudentName,Gender,DOB,Session,Caste,Religion,FatherName,FatherCN,MotherName,PermanentAddress,TemporaryAddress,ContactNo,EmailID,SectionID,SchoolID,LastSchoolAttended,Result,PassPerCentage,Nationality,Status,House,Photo,SSSM_ID,AccountNo,Accountname,Bank,Branch,IFSCCode,A_ID) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17,@d18,@d19,@d20,@d21,@d22,@d23,@d24,@d25,@d26,@d27,@d28,@d29,@d30,@d31,@d32,@d33," & Val(txtA_ID.Text) & ")"
                        cmd = New SqlCommand(cb)
                        cmd.Connection = con
                        cmd.Parameters.AddWithValue("@d1", txtANo.Text)
                        cmd.Parameters.AddWithValue("@d2", row.Cells(1).Value)
                        cmd.Parameters.AddWithValue("@d3", row.Cells(2).Value)
                        cmd.Parameters.AddWithValue("@d4", row.Cells(3).Value)
                        cmd.Parameters.AddWithValue("@d5", row.Cells(4).Value)
                        cmd.Parameters.AddWithValue("@d6", row.Cells(5).Value)
                        cmd.Parameters.AddWithValue("@d7", row.Cells(6).Value)
                        cmd.Parameters.AddWithValue("@d8", row.Cells(7).Value)
                        cmd.Parameters.AddWithValue("@d9", row.Cells(8).Value)
                        cmd.Parameters.AddWithValue("@d10", row.Cells(9).Value)
                        cmd.Parameters.AddWithValue("@d11", row.Cells(10).Value)
                        cmd.Parameters.AddWithValue("@d12", row.Cells(11).Value)
                        cmd.Parameters.AddWithValue("@d13", row.Cells(12).Value.ToString())
                        cmd.Parameters.AddWithValue("@d14", row.Cells(13).Value)
                        cmd.Parameters.AddWithValue("@d15", row.Cells(14).Value)
                        cmd.Parameters.AddWithValue("@d16", row.Cells(15).Value)
                        cmd.Parameters.AddWithValue("@d17", row.Cells(16).Value.ToString)
                        cmd.Parameters.AddWithValue("@d18", row.Cells(17).Value)
                        cmd.Parameters.AddWithValue("@d19", Val(row.Cells(18).Value))
                        cmd.Parameters.AddWithValue("@d20", Val(row.Cells(21).Value))
                        cmd.Parameters.AddWithValue("@d21", row.Cells(23).Value)
                        cmd.Parameters.AddWithValue("@d22", row.Cells(24).Value)
                        cmd.Parameters.AddWithValue("@d23", row.Cells(25).Value)
                        cmd.Parameters.AddWithValue("@d24", row.Cells(26).Value)
                        cmd.Parameters.AddWithValue("@d25", row.Cells(27).Value)
                        cmd.Parameters.AddWithValue("@d26", row.Cells(28).Value)
                        cmd.Parameters.AddWithValue("@d28", row.Cells(29).Value.ToString())
                        cmd.Parameters.AddWithValue("@d29", row.Cells(30).Value.ToString())
                        cmd.Parameters.AddWithValue("@d30", row.Cells(31).Value.ToString())
                        cmd.Parameters.AddWithValue("@d31", row.Cells(32).Value.ToString())
                        cmd.Parameters.AddWithValue("@d32", row.Cells(33).Value.ToString())
                        cmd.Parameters.AddWithValue("@d33", row.Cells(34).Value.ToString())
                        Dim ms As New MemoryStream()
                        Dim bmpImage As New Bitmap(My.Resources.photo)
                        bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                        Dim data As Byte() = ms.GetBuffer()
                        Dim p As New SqlParameter("@d27", SqlDbType.Image)
                        p.Value = data
                        cmd.Parameters.Add(p)
                        cmd.ExecuteNonQuery()
                        con.Close()
                        SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb3 As String = "insert into Discount(AdmissionNo,FeeType,Discount) VALUES (@d1,'Class',0.00)"
                        cmd = New SqlCommand(cb3)
                        cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                        cmd.Connection = con
                        cmd.ExecuteReader()
                        con.Close()
                        SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb4 As String = "insert into Discount(AdmissionNo,FeeType,Discount) VALUES (@d1,'Hostel',0.00)"
                        cmd = New SqlCommand(cb4)
                        cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                        cmd.Connection = con
                        cmd.ExecuteReader()
                        con.Close()
                        SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb5 As String = "insert into Discount(AdmissionNo,FeeType,Discount) VALUES (@d1,'Bus',0.00)"
                        cmd = New SqlCommand(cb5)
                        cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                        cmd.Connection = con
                        cmd.ExecuteReader()
                        con.Close()
                    End If
                End If
            Next
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
        Catch ex As SqlException
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        Try
            If dgw.RowCount = Nothing Then
                MessageBox.Show("Sorry nothing to update.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    Cursor = Cursors.WaitCursor
                    Timer1.Enabled = True
                    SqlConnection.ClearAllPools()
                    con = New SqlConnection(cs)
                    con.Open()
                    Dim cb As String = "Update Student set EnrollmentNo=@d2,GRNo=@d3,UID=@d4,AdmissionDate=@d5,StudentName=@d6,Gender=@d7,DOB=@d8,Session=@d9,Caste=@d10,Religion=@d11,FatherName=@d12,FatherCN=@d13,MotherName=@d14,PermanentAddress=@d15,TemporaryAddress=@d16,ContactNo=@d17,EmailID=@d18,SectionID=@d19,SchoolID=@d20,LastSchoolAttended=@d21,Result=@d22,PassPerCentage=@d23,Nationality=@d24,Status=@d25,House=@d26,SSSM_ID=@d27,AccountNo=@d28,Accountname=@d29,Bank=@d30,Branch=@d31,IFSCCode=@d32 where AdmissionNo=@d1"
                    cmd = New SqlCommand(cb)
                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d2", row.Cells(1).Value)
                    cmd.Parameters.AddWithValue("@d3", row.Cells(2).Value)
                    cmd.Parameters.AddWithValue("@d4", row.Cells(3).Value)
                    cmd.Parameters.AddWithValue("@d5", row.Cells(4).Value)
                    cmd.Parameters.AddWithValue("@d6", row.Cells(5).Value)
                    cmd.Parameters.AddWithValue("@d7", row.Cells(6).Value)
                    cmd.Parameters.AddWithValue("@d8", row.Cells(7).Value)
                    cmd.Parameters.AddWithValue("@d9", row.Cells(8).Value)
                    cmd.Parameters.AddWithValue("@d10", row.Cells(9).Value)
                    cmd.Parameters.AddWithValue("@d11", row.Cells(10).Value)
                    cmd.Parameters.AddWithValue("@d12", row.Cells(11).Value)
                    cmd.Parameters.AddWithValue("@d13", row.Cells(12).Value.ToString())
                    cmd.Parameters.AddWithValue("@d14", row.Cells(13).Value)
                    cmd.Parameters.AddWithValue("@d15", row.Cells(14).Value)
                    cmd.Parameters.AddWithValue("@d16", row.Cells(15).Value)
                    cmd.Parameters.AddWithValue("@d17", row.Cells(16).Value.ToString())
                    cmd.Parameters.AddWithValue("@d18", row.Cells(17).Value)
                    cmd.Parameters.AddWithValue("@d19", Val(row.Cells(18).Value))
                    cmd.Parameters.AddWithValue("@d20", Val(row.Cells(21).Value))
                    cmd.Parameters.AddWithValue("@d21", row.Cells(23).Value)
                    cmd.Parameters.AddWithValue("@d22", row.Cells(24).Value)
                    cmd.Parameters.AddWithValue("@d23", row.Cells(25).Value)
                    cmd.Parameters.AddWithValue("@d24", row.Cells(26).Value)
                    cmd.Parameters.AddWithValue("@d25", row.Cells(27).Value)
                    cmd.Parameters.AddWithValue("@d26", row.Cells(28).Value)
                    cmd.Parameters.AddWithValue("@d27", row.Cells(29).Value.ToString())
                    cmd.Parameters.AddWithValue("@d28", row.Cells(30).Value.ToString())
                    cmd.Parameters.AddWithValue("@d29", row.Cells(31).Value.ToString())
                    cmd.Parameters.AddWithValue("@d30", row.Cells(32).Value.ToString())
                    cmd.Parameters.AddWithValue("@d31", row.Cells(33).Value.ToString())
                    cmd.Parameters.AddWithValue("@d32", row.Cells(34).Value.ToString())
                    cmd.ExecuteNonQuery()
                    con.Close()
                End If
            Next
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
        Catch ex As SqlException
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
