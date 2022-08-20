Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmBusFeePayment_StudentRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select  RTRIM(BFP_ID) as [ID], RTRIM(PaymentID) as [Payment ID],RTRIM(BusHolderID) as [Bus Holder ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(BusCardHolder_Student.Location) as [Location],RTRIM(Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(BusFeePayment_Student.Session) as [Session],RTRIM(installment) as [installment],  RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Section,Class,Student,SchoolInfo,BusFeePayment_Student,BusCardHolder_Student where SchoolInfo.S_ID=Student.SchoolID and BusCardHolder_Student.BCH_ID=BusFeePayment_Student.BusHolderID and BusCardHolder_Student.AdmissionNo=Student.AdmissionNo and Class.Classname=Section.Class and Section.ID=Student.SectionID order by StudentName", con)
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
            cmd = New SqlCommand("Select  RTRIM(BFP_ID) as [ID], RTRIM(PaymentID) as [Payment ID],RTRIM(BusHolderID) as [Bus Holder ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(BusCardHolder_Student.Location) as [Location],RTRIM(Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(BusFeePayment_Student.Session) as [Session],RTRIM(installment) as [installment],  RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Section,Class,Student,SchoolInfo,BusFeePayment_Student,BusCardHolder_Student where SchoolInfo.S_ID=Student.SchoolID and BusCardHolder_Student.BCH_ID=BusFeePayment_Student.BusHolderID and BusCardHolder_Student.AdmissionNo=Student.AdmissionNo and Class.Classname=Section.Class and Section.ID=Student.SectionID and StudentName like '%" & txtStudentName.Text & "%' order by StudentName", con)
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
            cmd = New SqlCommand("Select  RTRIM(BFP_ID) as [ID], RTRIM(PaymentID) as [Payment ID],RTRIM(BusHolderID) as [Bus Holder ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(BusCardHolder_Student.Location) as [Location],RTRIM(Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(BusFeePayment_Student.Session) as [Session],RTRIM(installment) as [installment],  RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Section,Class,Student,SchoolInfo,BusFeePayment_Student,BusCardHolder_Student where SchoolInfo.S_ID=Student.SchoolID and BusCardHolder_Student.BCH_ID=BusFeePayment_Student.BusHolderID and BusCardHolder_Student.AdmissionNo=Student.AdmissionNo and Class.Classname=Section.Class and Section.ID=Student.SectionID and BusFeePayment_Student.Session=@d1 and BusFeePayment_Student.Student_Class=@d2 order by StudentName", con)
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
            cmd = New SqlCommand("Select  RTRIM(BFP_ID) as [ID], RTRIM(PaymentID) as [Payment ID],RTRIM(BusHolderID) as [Bus Holder ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(BusCardHolder_Student.Location) as [Location],RTRIM(Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(BusFeePayment_Student.Session) as [Session],RTRIM(installment) as [installment],  RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Section,Class,Student,SchoolInfo,BusFeePayment_Student,BusCardHolder_Student where SchoolInfo.S_ID=Student.SchoolID and BusCardHolder_Student.BCH_ID=BusFeePayment_Student.BusHolderID and BusCardHolder_Student.AdmissionNo=Student.AdmissionNo and Class.Classname=Section.Class and Section.ID=Student.SectionID and PaymentDate between @d1 and @d2 order by StudentName", con)
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
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM BusFeePayment_Student", con)
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
            Dim ct As String = "SELECT distinct RTRIM(Student_Class) as [Class] from Section,Class,Student,SchoolInfo,BusFeePayment_Student,BusCardHolder_Student where SchoolInfo.S_ID=Student.SchoolID and BusCardHolder_Student.BCH_ID=BusFeePayment_Student.BusHolderID and BusCardHolder_Student.AdmissionNo=Student.AdmissionNo and Class.Classname=Section.Class and Section.ID=Student.SectionID and BusFeePayment_Student.Session=@d1 order by 1"
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
            cmd = New SqlCommand("Select  RTRIM(BFP_ID) as [ID], RTRIM(PaymentID) as [Payment ID],RTRIM(BusHolderID) as [Bus Holder ID], RTRIM(Student.AdmissionNo) as [Admission No.],RTRIM(StudentName) as [StudentName],RTRIM(EnrollmentNo) as [Enrollment No.],RTRIM(SchoolName) as [School Name],RTRIM(BusCardHolder_Student.Location) as [Location],RTRIM(Student_Class) as [Class],RTRIM(SectionName) as [Section], RTRIM(BusFeePayment_Student.Session) as [Session],RTRIM(installment) as [installment],  RTRIM(TotalFee) as [Total Fee], RTRIM(DiscountPer) as [Discount %], RTRIM(DiscountAmt) as [Discount], RTRIM(PreviousDue) as [Previous Due], RTRIM(Fine) as [Fine], RTRIM(GrandTotal) as [Grand Total], RTRIM(TotalPaid) as [Total Paid], RTRIM(ModeOfPayment) as [Payment Mode], RTRIM(PaymentModeDetails) as [Payement Mode Info], Convert(Datetime,PaymentDate,131) as [Payement Date], RTRIM(PaymentDue) as [Payement Due] from Section,Class,Student,SchoolInfo,BusFeePayment_Student,BusCardHolder_Student where SchoolInfo.S_ID=Student.SchoolID and BusCardHolder_Student.BCH_ID=BusFeePayment_Student.BusHolderID and BusCardHolder_Student.AdmissionNo=Student.AdmissionNo and Class.Classname=Section.Class and Section.ID=Student.SectionID and Student.AdmissionNo like '%" & txtAdmissionNo.Text & "%' order by StudentName", con)
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
                If lblSet.Text = "Bus Fee Payment" Then
                    Me.Hide()
                    frmBusFeePayment_Student.Show()
                    frmBusFeePayment_Student.cmbInstallment.DropDownStyle = ComboBoxStyle.DropDown
                    frmBusFeePayment_Student.txtBFPId.Text = dr.Cells(0).Value.ToString()
                    frmBusFeePayment_Student.txtFeePaymentID.Text = dr.Cells(1).Value.ToString()
                    frmBusFeePayment_Student.txtBusHolderID.Text = dr.Cells(2).Value.ToString()
                    frmBusFeePayment_Student.txtAdmissionNo.Text = dr.Cells(3).Value.ToString()
                    frmBusFeePayment_Student.txtStudentName.Text = dr.Cells(4).Value.ToString()
                    frmBusFeePayment_Student.txtEnrollmentNo.Text = dr.Cells(5).Value.ToString()
                    frmBusFeePayment_Student.txtLocation.Text = dr.Cells(6).Value.ToString() '
                    frmBusFeePayment_Student.txtSchoolname.Text = dr.Cells(7).Value.ToString()
                    frmBusFeePayment_Student.txtClass.Text = dr.Cells(8).Value.ToString()
                    frmBusFeePayment_Student.txtSection.Text = dr.Cells(9).Value.ToString()
                    frmBusFeePayment_Student.txtSession.Text = dr.Cells(10).Value.ToString()
                    frmBusFeePayment_Student.cmbInstallment.Text = dr.Cells(11).Value.ToString()
                    frmBusFeePayment_Student.txtBusFee.Text = dr.Cells(12).Value.ToString()
                    frmBusFeePayment_Student.txtDiscountPer.Text = dr.Cells(13).Value.ToString()
                    frmBusFeePayment_Student.txtDiscount.Text = dr.Cells(14).Value.ToString()
                    frmBusFeePayment_Student.txtPreviousDue.Text = dr.Cells(15).Value.ToString()
                    frmBusFeePayment_Student.txtFine.Text = dr.Cells(16).Value.ToString()
                    frmBusFeePayment_Student.txtGrandTotal.Text = dr.Cells(17).Value.ToString()
                    frmBusFeePayment_Student.txtTotalPaid.Text = dr.Cells(18).Value.ToString()
                    frmBusFeePayment_Student.cmbPaymentMode.Text = dr.Cells(19).Value.ToString()
                    frmBusFeePayment_Student.txtPaymentModeDetails.Text = dr.Cells(20).Value.ToString()
                    frmBusFeePayment_Student.dtpPaymentDate.Text = dr.Cells(21).Value.ToString()
                    frmBusFeePayment_Student.txtBalance.Text = dr.Cells(22).Value.ToString()
                    If lblUserType.Text = "Admin" Then
                        frmBusFeePayment_Student.btnDelete.Enabled = True
                        frmBusFeePayment_Student.btnPrint.Enabled = True
                    Else
                        frmBusFeePayment_Student.btnDelete.Enabled = False
                        frmBusFeePayment_Student.btnPrint.Enabled = False
                    End If
                    frmBusFeePayment_Student.btnSave.Enabled = False
                    frmBusFeePayment_Student.btnUpdate.Enabled = True
                    frmBusFeePayment_Student.Button2.Enabled = False
                    frmBusFeePayment_Student.dtpPaymentDate.Enabled = False
                    frmBusFeePayment_Student.btnPrint.Enabled = True
                    frmBusFeePayment_Student.cmbInstallment.Enabled = False
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
