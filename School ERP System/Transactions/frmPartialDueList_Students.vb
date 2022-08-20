﻿Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmPartialDueList_Students
    Dim st2, st1, st3 As String
  
    Sub fillSession1()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM HostelFeepayment", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSession1.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSession1.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillSession2()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM BusFeepayment_Student", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSession2.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSession2.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
  
    Sub Reset()
       
        cmbSession1.SelectedIndex = -1
        cmbClass1.SelectedIndex = -1
        cmbInstallment.SelectedIndex = -1
        cmbClass1.Enabled = False
        cmbInstallment.Enabled = False
        dgw1.Rows.Clear()
        cmbSession2.SelectedIndex = -1
        cmbClass2.SelectedIndex = -1
        cmbInstallment1.SelectedIndex = -1
        cmbInstallment1.Enabled = False
        cmbClass2.Enabled = False
        dgw2.Rows.Clear()
        fillSession1()
        fillSession2()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmPartialDueList_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSession1()
        fillSession2()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        cmbSession1.SelectedIndex = -1
        cmbClass1.SelectedIndex = -1
        cmbInstallment.SelectedIndex = -1
        cmbClass1.Enabled = False
        cmbInstallment.Enabled = False
        dgw1.Rows.Clear()
    End Sub

    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        cmbSession2.SelectedIndex = -1
        cmbClass2.SelectedIndex = -1
        cmbInstallment1.SelectedIndex = -1
        cmbInstallment1.Enabled = False
        cmbClass2.Enabled = False
        dgw2.Rows.Clear()
    End Sub

    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        ExportExcel(dgw2)
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        ExportExcel(dgw1)
    End Sub

    Private Sub cmbSession1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSession1.SelectedIndexChanged
        Try
            cmbClass1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Student_Class) FROM HostelFeePayment where Session=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession1.Text)
            rdr = cmd.ExecuteReader()
            cmbClass1.Items.Clear()
            While rdr.Read
                cmbClass1.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbSession2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSession2.SelectedIndexChanged
        Try
            cmbClass2.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Student_Class) FROM BusFeePayment_Student where Session=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession2.Text)
            rdr = cmd.ExecuteReader()
            cmbClass2.Items.Clear()
            While rdr.Read
                cmbClass2.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TabControl1_Click(sender As System.Object, e As System.EventArgs) Handles TabControl1.Click
        Reset()
    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        Try
            If Len(Trim(cmbSession1.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession1.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbClass1.Text)) = 0 Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass1.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbInstallment.Text)) = 0 Then
                MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbInstallment.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(Student.AdmissionNo),RTRIM(GRNo),RTRIM(StudentName),RTRIM(HostelName),RTRIM(SchoolName),RTRIM(PaymentDue) FROM HostelFeePayment,Student,Hosteler,HostelInfo,schoolInfo where Student.AdmissionNo=Hosteler.AdmissionNo and HostelInfo.HI_ID=Hosteler.HostelID and HostelFeePayment.HostelerID=Hosteler.H_ID and Student.SchoolID=SchoolInfo.S_ID and HostelFeePayment.Session=@d1 and HostelFeePayment.Student_Class=@d2 and HostelFeePayment.Installment=@d3 and PaymentDue > 0 order by StudentName", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession1.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass1.Text)
            cmd.Parameters.AddWithValue("@d3", cmbInstallment.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw1.Rows.Clear()
            While (rdr.Read() = True)
                dgw1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        Try
            If Len(Trim(cmbSession2.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession2.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbClass2.Text)) = 0 Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass2.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbInstallment1.Text)) = 0 Then
                MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbInstallment1.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(Student.AdmissionNo),RTRIM(GRNo),RTRIM(StudentName),RTRIM(BusCardHolder_Student.Location),RTRIM(SchoolName),RTRIM(PaymentDue) FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN BusCardHolder_Student ON Student.AdmissionNo = BusCardHolder_Student.AdmissionNo INNER JOIN BusFeePayment_Student ON BusCardHolder_Student.BCH_Id = BusFeePayment_Student.BusHolderID where BusFeePayment_Student.Session=@d1 and BusFeePayment_Student.Student_Class=@d2 and Installment=@d3 and PaymentDue > 0 order by StudentName", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession2.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass2.Text)
            cmd.Parameters.AddWithValue("@d3", cmbInstallment1.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw2.Rows.Clear()
            While (rdr.Read() = True)
                dgw2.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbClass1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbClass1.SelectedIndexChanged
        Try
            cmbInstallment.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Installment) FROM HostelFeePayment,Student,Hosteler,HostelInfo where Student.AdmissionNo=Hosteler.AdmissionNo and HostelInfo.HI_ID=Hosteler.HostelID and HostelFeePayment.HostelerID=Hosteler.H_ID and HostelFeePayment.Session=@d1 and HostelFeePayment.Student_Class=@d2"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession1.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass1.Text)
            rdr = cmd.ExecuteReader()
            cmbInstallment.Items.Clear()
            While rdr.Read
                cmbInstallment.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbClass2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbClass2.SelectedIndexChanged
        Try
            cmbInstallment1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Installment) FROM BusFeePayment_Student,Student,BusCardHolder_Student where Student.AdmissionNo=BusCardHolder_Student.AdmissionNo and BusFeePayment_Student.BusHolderID=BusCardHolder_Student.BCH_ID and BusFeePayment_Student.Session=@d1 and BusFeePayment_Student.Student_Class=@d2"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession2.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass2.Text)
            rdr = cmd.ExecuteReader()
            cmbInstallment1.Items.Clear()
            While rdr.Read
                cmbInstallment1.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Try
            If Len(Trim(cmbSession1.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession1.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbClass1.Text)) = 0 Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass1.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbInstallment.Text)) = 0 Then
                MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbInstallment.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptHostelFeePartialDue 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT SchoolInfo.S_Id, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo, SchoolInfo.EstablishedYear, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName,Student.FatherName, Student.MotherName, Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.EmailID, Student.DOB, Student.Gender,Student.AdmissionDate, Student.Session, Student.Caste, Student.Religion, Student.SectionID, Student.Photo, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result,Student.PassPercentage, Hosteler.H_Id, Hosteler.HostelID, Hosteler.JoiningDate, HostelFeePayment.HFP_Id,HostelFeePayment.PaymentID, HostelFeePayment.HostelerID, HostelFeePayment.Installment, HostelFeePayment.TotalFee, HostelFeePayment.DiscountPer,HostelFeePayment.DiscountAmt, HostelFeePayment.PreviousDue, HostelFeePayment.Fine, HostelFeePayment.GrandTotal, HostelFeePayment.TotalPaid, HostelFeePayment.ModeOfPayment,HostelFeePayment.PaymentModeDetails, HostelFeePayment.Paymentdate, HostelFeePayment.PaymentDue, HostelFeePayment.Student_Class , HostelInfo.HI_Id, HostelInfo.Hostelname, HostelInfo.ManagedBy,HostelInfo.Person_ContactNo FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN Hosteler ON Student.AdmissionNo = Hosteler.AdmissionNo INNER JOIN HostelFeePayment ON Hosteler.H_Id = HostelFeePayment.HostelerID INNER JOIN HostelInfo ON Hosteler.HostelID = HostelInfo.HI_Id where HostelFeePayment.Session=@d1 and HostelFeePayment.Student_Class=@d2 and HostelFeePayment.Installment=@d3 and HostelFeePayment.PaymentDue > 0 order by Student.StudentName"
            MyCommand.Parameters.AddWithValue("@d1", cmbSession1.Text)
            MyCommand.Parameters.AddWithValue("@d2", cmbClass1.Text)
            MyCommand.Parameters.AddWithValue("@d3", cmbInstallment.Text)
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Student")
            myDA.Fill(myDS, "Hosteler")
            myDA.Fill(myDS, "SchoolInfo")
            myDA.Fill(myDS, "HostelFeePayment")
            myDA.Fill(myDS, "HostelInfo")
            rpt.SetDataSource(myDS)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        Try
            If Len(Trim(cmbSession2.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession2.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbClass2.Text)) = 0 Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass2.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbInstallment1.Text)) = 0 Then
                MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbInstallment1.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptBusFeePartialDue_Student 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT SchoolInfo.S_Id, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo, SchoolInfo.EstablishedYear, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName,Student.FatherName, Student.MotherName, Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.EmailID, Student.DOB, Student.Gender,Student.AdmissionDate, Student.Caste, Student.Religion, Student.SectionID, Student.Photo, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result,Student.PassPercentage, Student.Status, BusCardHolder_Student.BCH_Id, BusCardHolder_Student.AdmissionNo AS Expr2, BusCardHolder_Student.BusNo, BusCardHolder_Student.Location,BusCardHolder_Student.JoiningDate, BusFeePayment_Student.BFP_ID, BusFeePayment_Student.PaymentID,BusFeePayment_Student.BusHolderID, BusFeePayment_Student.Session , BusFeePayment_Student.Installment, BusFeePayment_Student.TotalFee, BusFeePayment_Student.DiscountPer,BusFeePayment_Student.DiscountAmt, BusFeePayment_Student.PreviousDue, BusFeePayment_Student.Fine, BusFeePayment_Student.GrandTotal, BusFeePayment_Student.TotalPaid,BusFeePayment_Student.ModeOfPayment, BusFeePayment_Student.PaymentModeDetails, BusFeePayment_Student.PaymentDate, BusFeePayment_Student.PaymentDue, BusFeePayment_Student.Student_Class FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN BusCardHolder_Student ON Student.AdmissionNo = BusCardHolder_Student.AdmissionNo INNER JOIN BusFeePayment_Student ON BusCardHolder_Student.BCH_Id = BusFeePayment_Student.BusHolderID where BusFeePayment_Student.Session=@d1 and BusFeePayment_Student.Student_Class=@d2 and Installment=@d3 and PaymentDue > 0 order by StudentName"
            MyCommand.Parameters.AddWithValue("@d1", cmbSession2.Text)
            MyCommand.Parameters.AddWithValue("@d2", cmbClass2.Text)
            MyCommand.Parameters.AddWithValue("@d3", cmbInstallment1.Text)
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Student")
            myDA.Fill(myDS, "BusCardHolder_Student")
            myDA.Fill(myDS, "SchoolInfo")
            myDA.Fill(myDS, "BusFeePayment_Student")
            rpt.SetDataSource(myDS)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Button12_Click(sender As System.Object, e As System.EventArgs) Handles Button12.Click
        Try

            If dgw1.Rows.Count = 0 Then
                MessageBox.Show("Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            If CheckForInternetConnection() = True Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ctn As String = "select RTRIM(APIURL) from SMSSetting where IsDefault='Yes' and IsEnabled='Yes'"
                cmd = New SqlCommand(ctn)
                cmd.Connection = con
                Dim rdr1 As SqlDataReader = cmd.ExecuteReader()
                If rdr1.Read() Then
                    For Each row As DataGridViewRow In dgw1.Rows
                        If Not row.IsNewRow Then
                            GetFatherCN(row.Cells(0).Value)
                            st2 = rdr1.GetValue(0)
                            Dim st3 As String = "Hello Sir/Madam, the due amount of " & cmbInstallment.Text.ToString.Trim() & " installment of hostel Fee of the Student " & row.Cells(2).Value & " is " & row.Cells(5).Value & ""
                            SMSFunc(txtFatherCN.Text, st3, st2)
                        End If
                    Next
                End If
                MessageBox.Show("SMS Successfully sent", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button13_Click(sender As System.Object, e As System.EventArgs) Handles Button13.Click
        Try

            If dgw2.Rows.Count = 0 Then
                MessageBox.Show("Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            If CheckForInternetConnection() = True Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ctn As String = "select RTRIM(APIURL) from SMSSetting where IsDefault='Yes' and IsEnabled='Yes'"
                cmd = New SqlCommand(ctn)
                cmd.Connection = con
                Dim rdr1 As SqlDataReader = cmd.ExecuteReader()
                If rdr1.Read() Then
                    For Each row As DataGridViewRow In dgw2.Rows
                        If Not row.IsNewRow Then
                            GetFatherCN(row.Cells(0).Value)
                            st2 = rdr1.GetValue(0)
                            Dim st3 As String = "Hello Sir/Madam, the due amount of " & cmbInstallment1.Text.ToString.Trim() & " installment of bus Fee of the Student " & row.Cells(2).Value & " is " & row.Cells(5).Value & ""
                            SMSFunc(txtFatherCN.Text, st3, st2)
                        End If
                    Next
                End If
                MessageBox.Show("SMS Successfully sent", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub GetFatherCN(ByVal stX As String)
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RTRIM(FatherCN) FROM Student where AdmissionNo=@d1"
            cmd.Parameters.AddWithValue("@d1", stX)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtFatherCN.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
