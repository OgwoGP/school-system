Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmBusFeeReceipt_Student
    Sub fillPaymentID()
        Try
            Dim CN As New SqlConnection(cs)
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(PaymentID) FROM BusFeePayment_Student", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbPaymentID.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbPaymentID.Items.Add(drow(0).ToString())
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillPaymentID()
    End Sub
    Sub Reset()
        cmbPaymentID.Text = ""
        fillPaymentID()
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If cmbPaymentID.Text = "" Then
                MessageBox.Show("Please select payment id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbPaymentID.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptBusFeeReceipt_Student 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT SchoolInfo.S_Id, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo AS Expr4, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo, SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, BusFeePayment_Student.BFP_ID, BusFeePayment_Student.PaymentID, BusFeePayment_Student.BusHolderID, BusFeePayment_Student.Session, BusFeePayment_Student.Installment,BusFeePayment_Student.TotalFee, BusFeePayment_Student.DiscountPer, BusFeePayment_Student.DiscountAmt, BusFeePayment_Student.PreviousDue, BusFeePayment_Student.Fine,BusFeePayment_Student.GrandTotal, BusFeePayment_Student.TotalPaid, BusFeePayment_Student.ModeOfPayment, BusFeePayment_Student.PaymentModeDetails, BusFeePayment_Student.PaymentDate,BusFeePayment_Student.PaymentDue, BusCardHolder_Student.BCH_Id, BusCardHolder_Student.AdmissionNo, BusCardHolder_Student.BusNo, BusCardHolder_Student.Location, BusCardHolder_Student.JoiningDate, BusCardHolder_Student.Status, Student.AdmissionNo AS Expr1, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName,Student.MotherName, Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.ContactNo, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate,Student.Session AS Expr2, Student.Caste, Student.Religion, Student.SectionID, Student.Photo, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result, Student.PassPercentage,Student.Status AS Expr3, Section.Id, Section.SectionName, Section.Class, Class.ClassName, Class.ClassType FROM BusFeePayment_Student INNER JOIN BusCardHolder_Student ON BusFeePayment_Student.BusHolderID = BusCardHolder_Student.BCH_Id INNER JOIN Student ON BusCardHolder_Student.AdmissionNo = Student.AdmissionNo INNER JOIN  Section ON Student.SectionID = Section.Id INNER JOIN Class ON Section.Class = Class.ClassName INNER JOIN SchoolInfo ON Student.SchoolID = SchoolInfo.S_Id where PaymentID='" & cmbPaymentID.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Student")
            myDA.Fill(myDS, "BusCardHolder_Student")
            myDA.Fill(myDS, "SchoolInfo")
            myDA.Fill(myDS, "BusFeePayment_Student")
            myDA.Fill(myDS, "Class")
            myDA.Fill(myDS, "Section")
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
End Class
