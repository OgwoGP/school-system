Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmHostelFeeReceipt
    Sub fillPaymentID()
        Try
            Dim CN As New SqlConnection(cs)
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(PaymentID) FROM HostelFeePayment", CN)
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
                Dim rpt As New rptHostelFeeReceipt 'The report you created.
                Dim myConnection As SqlConnection
                Dim MyCommand As New SqlCommand()
                Dim myDA As New SqlDataAdapter()
                Dim myDS As New DataSet 'The DataSet you created.
                myConnection = New SqlConnection(cs)
                MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT SchoolInfo.S_Id,HostelFeePayment.Session, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName,Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.ContactNo AS Expr1, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Caste,Student.Religion, Student.SectionID, Student.Photo, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result, Student.PassPercentage, Student.Status, Section.Id, Section.SectionName, Section.Class, Class.ClassName, Class.ClassType, Hosteler.H_Id, Hosteler.AdmissionNo AS Expr2, Hosteler.HostelID, Hosteler.JoiningDate, Hosteler.Status AS Expr3, HostelInfo.HI_Id, HostelInfo.Hostelname,HostelInfo.Address AS Expr4, HostelInfo.ContactNo AS Expr5, HostelInfo.ManagedBy, HostelInfo.Person_ContactNo, HostelFeePayment.HFP_Id, HostelFeePayment.PaymentID, HostelFeePayment.HostelerID, HostelFeePayment.Installment, HostelFeePayment.TotalFee, HostelFeePayment.DiscountPer, HostelFeePayment.DiscountAmt, HostelFeePayment.PreviousDue,HostelFeePayment.Fine, HostelFeePayment.GrandTotal, HostelFeePayment.TotalPaid, HostelFeePayment.ModeOfPayment, HostelFeePayment.PaymentModeDetails, HostelFeePayment.Paymentdate, HostelFeePayment.PaymentDue FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN Section ON Student.SectionID = Section.Id INNER JOIN Class ON Section.Class = Class.ClassName INNER JOIN Hosteler ON Student.AdmissionNo = Hosteler.AdmissionNo INNER JOIN HostelInfo ON Hosteler.HostelID = HostelInfo.HI_Id INNER JOIN HostelFeePayment ON Hosteler.H_Id = HostelFeePayment.HostelerID where PaymentID='" & cmbPaymentID.Text & "'"
                MyCommand.CommandType = CommandType.Text
                myDA.SelectCommand = MyCommand
                myDA.Fill(myDS, "Student")
                myDA.Fill(myDS, "Hosteler")
                myDA.Fill(myDS, "SchoolInfo")
                myDA.Fill(myDS, "HostelFeePayment")
                myDA.Fill(myDS, "HostelInfo")
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
