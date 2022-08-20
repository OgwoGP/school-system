Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmClassFeePaymentRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select  CFP_ID as [ID],RTRIM(PaymentID) as [Payment ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(CourseFeePayment.Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(CourseFeePayment.Session) as [Session], RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Student,Class,Section,SchoolInfo,CourseFeePayment where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.AdmissionNo=CourseFeePayment.AdmissionNo order by StudentName", con)
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
            cmd = New SqlCommand("Select  CFP_ID as [ID],RTRIM(PaymentID) as [Payment ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(CourseFeePayment.Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(CourseFeePayment.Session) as [Session], RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Student,Class,Section,SchoolInfo,CourseFeePayment where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.AdmissionNo=CourseFeePayment.AdmissionNo and StudentName like '%" & txtStudentName.Text & "%' order by StudentName", con)
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
            If cmbSession.Text = "" Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If cmbClass.Text = "" Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select  CFP_ID as [ID],RTRIM(PaymentID) as [Payment ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(CourseFeePayment.Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(CourseFeePayment.Session) as [Session], RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Student,Class,Section,SchoolInfo,CourseFeePayment where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.AdmissionNo=CourseFeePayment.AdmissionNo and CourseFeePayment.Session=@d1 and CourseFeePayment.Student_Class=@d2 order by StudentName", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
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
            cmd = New SqlCommand("Select  CFP_ID as [ID],RTRIM(PaymentID) as [Payment ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(CourseFeePayment.Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(CourseFeePayment.Session) as [Session], RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Student,Class,Section,SchoolInfo,CourseFeePayment where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.AdmissionNo=CourseFeePayment.AdmissionNo and PaymentDate between @d1 and @d2 order by StudentName", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value
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
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM CourseFeepayment", con)
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
            Dim ct As String = "SELECT distinct RTRIM(Student_Class) from Student,Class,Section,SchoolInfo,CourseFeePayment where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.AdmissionNo=CourseFeePayment.AdmissionNo and CourseFeePayment.Session=@d1"
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


    Private Sub txtAdmissionNo_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtAdmissionNo.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select  CFP_ID as [ID],RTRIM(PaymentID) as [Payment ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(CourseFeePayment.Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(CourseFeePayment.Session) as [Session], RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Student,Class,Section,SchoolInfo,CourseFeePayment where Student.SectionID=Section.ID and Class.ClassName=Section.Class and SchoolInfo.S_ID=Student.SchoolID and Student.AdmissionNo=CourseFeePayment.AdmissionNo and Student.AdmissionNo like '%" & txtAdmissionNo.Text & "%' order by StudentName", con)
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
        cmbSession.SelectedIndex = -1
        cmbClass.Enabled = False
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Now
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSession()
        GetData()
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Course Fee Payment" Then
                    Me.Hide()
                    frmClassFeePayment.Show()
                    frmClassFeePayment.txtCFPId.Text = dr.Cells(0).Value.ToString()
                    frmClassFeePayment.txtFeePaymentID.Text = dr.Cells(1).Value.ToString()
                    frmClassFeePayment.txtAdmissionNo.Text = dr.Cells(2).Value.ToString()
                    frmClassFeePayment.txtStudentName.Text = dr.Cells(3).Value.ToString()
                    frmClassFeePayment.txtEnrollmentNo.Text = dr.Cells(4).Value.ToString() '
                    frmClassFeePayment.txtSchoolName.Text = dr.Cells(5).Value.ToString()
                    frmClassFeePayment.txtClass.Text = dr.Cells(6).Value.ToString()
                    frmClassFeePayment.txtSection.Text = dr.Cells(7).Value.ToString()
                    frmClassFeePayment.txtSession.Text = dr.Cells(8).Value.ToString()
                    frmClassFeePayment.txtCourseFee.Text = dr.Cells(9).Value.ToString()
                    frmClassFeePayment.txtDiscountPer.Text = dr.Cells(10).Value.ToString()
                    frmClassFeePayment.txtDiscount.Text = dr.Cells(11).Value.ToString()
                    frmClassFeePayment.txtPreviousDue.Text = dr.Cells(12).Value.ToString()
                    frmClassFeePayment.txtFine.Text = dr.Cells(13).Value.ToString()
                    frmClassFeePayment.txtGrandTotal.Text = dr.Cells(14).Value.ToString()
                    frmClassFeePayment.txtTotalPaid.Text = dr.Cells(15).Value.ToString()
                    frmClassFeePayment.cmbPaymentMode.Text = dr.Cells(16).Value.ToString()
                    frmClassFeePayment.txtPaymentModeDetails.Text = dr.Cells(17).Value.ToString()
                    frmClassFeePayment.dtpPaymentDate.Text = dr.Cells(18).Value.ToString()
                    frmClassFeePayment.txtBalance.Text = dr.Cells(19).Value.ToString()
                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = New SqlCommand("SELECT RTRIM(FeeName),CourseFeePayment_Join.Fee from CourseFeePayment,CourseFeePayment_Join where CourseFeePayment.CFP_ID=CourseFeePayment_Join.C_PaymentID and CourseFeePayment.CFP_ID=" & dr.Cells(0).Value & "", con)
                    rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                    frmClassFeePayment.dgw.Rows.Clear()
                    While (rdr.Read() = True)
                        frmClassFeePayment.dgw.Rows.Add(rdr(0), rdr(1))
                    End While
                    con.Close()
                    If lblUserType.Text = "Admin" Then
                        frmClassFeePayment.btnDelete.Enabled = True
                        frmClassFeePayment.btnPrint.Enabled = True
                    Else
                        frmClassFeePayment.btnDelete.Enabled = False
                        frmClassFeePayment.btnPrint.Enabled = False
                    End If
                    frmClassFeePayment.btnSave.Enabled = False
                    frmClassFeePayment.btnStudentSelection.Enabled = False
                    frmClassFeePayment.dtpPaymentDate.Enabled = False
                    frmClassFeePayment.btnPrint.Enabled = True
                    frmClassFeePayment.gpbFeeInfo.Enabled = False
                    lblSet.Text = ""
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

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ExportExcel(dgw)
    End Sub

End Class
