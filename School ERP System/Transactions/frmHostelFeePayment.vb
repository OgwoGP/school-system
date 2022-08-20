Imports System.Data.SqlClient
Imports System.IO

Public Class frmHostelFeePayment
    Dim st2 As String
    Sub Calculate()
        Dim num1, num2 As Double
        num1 = CDbl(Val(txthostelFee.Text) + Val(txtFine.Text) + Val(txtPreviousDue.Text) - Val(txtDiscount.Text))
        num1 = Math.Round(num1, 2)
        txtGrandTotal.Text = num1
        num2 = Val(txtGrandTotal.Text) - Val(txtTotalPaid.Text)
        num2 = Math.Round(num2, 2)
        txtBalance.Text = num2
    End Sub
    Private Function GenerateID() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 (HFP_ID) FROM hostelFeePayment ORDER BY HFP_ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("HFP_ID")
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
            txtHFPId.Text = GenerateID()
            Dim a As String = txtAdmissionNo.Text
            txtFeePaymentID.Text = "HFP-" + GenerateID() + "-" + a
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Sub Reset()
        txtAdmissionNo.Text = ""
        txthostelFee.Text = ""
        txtPaymentModeDetails.Text = ""
        txtSchoolName.Text = ""
        txtBalance.Text = ""
        txtClass.Text = ""
        txtDiscount.Text = "0.00"
        txtDiscountPer.Text = "0.00"
        txtEnrollmentNo.Text = ""
        txtFine.Text = "0.00"
        txtGrandTotal.Text = ""
        txtPreviousDue.Text = ""
        txtHostelName.Text = ""
        txtSection.Text = ""
        txtSession.Text = ""
        txtStudentName.Text = ""
        txtTotalPaid.Text = ""
        cmbPaymentMode.SelectedIndex = 0
        cmbInstallment.SelectedIndex = -1
        dtpPaymentDate.Value = Now
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        cmbInstallment.Enabled = True
        btnPrint.Enabled = False
        Button2.Enabled = True
        dtpPaymentDate.Enabled = True
    End Sub
    Sub FillData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT EnrollmentNo,Session FROM Student where AdmissionNo=@d1"
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtEnrollmentNo.Text = rdr.GetValue(0).ToString().Trim()
                txtSession.Text = rdr.GetValue(1).ToString().Trim()
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub Fill()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Discount from Discount where AdmissionNo=@d1 and FeeType='Hostel'"
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtDiscountPer.Text = rdr.GetValue(0)
                Dim num As Double
                num = CDbl((Val(txthostelFee.Text) * Val(txtDiscountPer.Text)) / 100)
                num = Math.Round(num, 2)
                txtDiscount.Text = num
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            Calculate()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmActiveHostelerRecord.lblSet.Text = "Hostel Fee Payment"
        frmActiveHostelerRecord.Reset()
        frmActiveHostelerRecord.ShowDialog()
    End Sub

    Private Sub btnGetFeeList_Click(sender As System.Object, e As System.EventArgs) Handles btnGetFeeList.Click
        Try
            If txtAdmissionNo.Text = "" Then
                MessageBox.Show("Please retrieve hosteler info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAdmissionNo.Focus()
                Return
            End If
            If cmbInstallment.Text = "" Then
                MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbInstallment.Focus()
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Charges from Installment_Hostel,HostelInfo,SchoolInfo,Class where Installment_Hostel.HostelID=HostelInfo.HI_ID and Installment_Hostel.Class=Class.Classname and SchoolInfo.S_ID=Installment_Hostel.SchoolID and Installment=@d1 and HostelName=@d2 and SchoolName=@d3 and Installment_Hostel.Class=@d4"
            cmd.Parameters.AddWithValue("@d1", cmbInstallment.Text)
            cmd.Parameters.AddWithValue("@d2", txtHostelName.Text)
            cmd.Parameters.AddWithValue("@d3", txtSchoolName.Text)
            cmd.Parameters.AddWithValue("@d4", txtClass.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txthostelFee.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Sum(PaymentDue-PreviousDue) from hostelFeePayment where HostelerID=@d1 group by HostelerID"
            cmd.Parameters.AddWithValue("@d1", txtHostelerID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtPreviousDue.Text = rdr.GetValue(0)
            Else
                txtPreviousDue.Text = 0
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            Fill()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
        Reset()
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtAdmissionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve hosteler info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbInstallment.Text)) = 0 Then
            MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbInstallment.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFine.Text)) = 0 Then
            MessageBox.Show("Please enter fine", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFine.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbPaymentMode.Text)) = 0 Then
            MessageBox.Show("Please select Payment Mode", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbPaymentMode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTotalPaid.Text)) = 0 Then
            MessageBox.Show("Please enter total paid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTotalPaid.Focus()
            Exit Sub
        End If
        If Val(txtTotalPaid.Text) < 0 Then
            MessageBox.Show("Total paid must be greater than zero", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTotalPaid.Focus()
            Exit Sub
        End If
        If Val(txtBalance.Text) < 0 Then
            MessageBox.Show("Balance is not possible less than zero", "Input error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select Session,HostelerID,Installment from hostelFeePayment where Session=@d1 and HostelerID=@d2 and Installment=@d3"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSession.Text)
            cmd.Parameters.AddWithValue("@d2", txtHostelerID.Text)
            cmd.Parameters.AddWithValue("@d3", cmbInstallment.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Already paid", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            auto()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into hostelFeePayment(HFP_ID, PaymentID, Hostelerid, Session,installment,TotalFee, DiscountPer, DiscountAmt, PreviousDue, Fine, GrandTotal, TotalPaid, ModeOfPayment, PaymentModeDetails, PaymentDate, PaymentDue,Student_Class) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17)"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtHFPId.Text)
            cmd.Parameters.AddWithValue("@d2", txtFeePaymentID.Text)
            cmd.Parameters.AddWithValue("@d3", txtHostelerID.Text)
            cmd.Parameters.AddWithValue("@d4", txtSession.Text)
            cmd.Parameters.AddWithValue("@d5", cmbInstallment.Text)
            cmd.Parameters.AddWithValue("@d6", txthostelFee.Text)
            cmd.Parameters.AddWithValue("@d7", txtDiscountPer.Text)
            cmd.Parameters.AddWithValue("@d8", txtDiscount.Text)
            cmd.Parameters.AddWithValue("@d9", txtPreviousDue.Text)
            cmd.Parameters.AddWithValue("@d10", txtFine.Text)
            cmd.Parameters.AddWithValue("@d11", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d12", txtTotalPaid.Text)
            cmd.Parameters.AddWithValue("@d13", cmbPaymentMode.Text)
            cmd.Parameters.AddWithValue("@d14", txtPaymentModeDetails.Text)
            cmd.Parameters.AddWithValue("@d15", dtpPaymentDate.Value)
            cmd.Parameters.AddWithValue("@d16", txtBalance.Text)
            cmd.Parameters.AddWithValue("@d17", txtClass.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            LedgerSave(dtpPaymentDate.Value.Date, txtStudentName.Text, txtFeePaymentID.Text, "Hostel Fee Payment", Val(txtGrandTotal.Text), 0, txtAdmissionNo.Text)
            If cmbPaymentMode.Text = "By Cash" Then
                LedgerSave(dtpPaymentDate.Value.Date, "Cash Account", txtFeePaymentID.Text, "Payment", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text)
            End If
            If cmbPaymentMode.Text = "By Cheque" Or cmbPaymentMode.Text = "By DD" Then
                LedgerSave(dtpPaymentDate.Value.Date, "bank Account", txtFeePaymentID.Text, "Payment", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text)
            End If
            Dim st As String = "added the new hostel fee payment entry having payment id '" & txtFeePaymentID.Text & "'"
            LogFunc(lblUser.Text, st)
            If CheckForInternetConnection() = True Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ctn As String = "select RTRIM(APIURL) from SMSSetting where IsDefault='Yes' and IsEnabled='Yes'"
                cmd = New SqlCommand(ctn)
                cmd.Connection = con
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    st2 = rdr.GetValue(0)
                    Dim st3 As String = "Hello, " & txtStudentName.Text & " you have successfully paid " & cmbInstallment.Text & " installment of hostel fee having Fee Payment ID '" & txtFeePaymentID.Text & "'"
                    SMSFunc(txtContactNo.Text, st3, st2)
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                End If
            End If
            MessageBox.Show("Successfully paid", "Fee", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Print()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete the record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
                delete_records()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub delete_records()
        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from hostelFeePayment where HFP_ID= " & txtHFPId.Text & ""
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LedgerDelete(txtFeePaymentID.Text, "Hostel Fee Payment")
                LedgerDelete(txtFeePaymentID.Text, "Payment")
                Dim st As String = "deleted the hostel fee payment entry having payment id '" & txtFeePaymentID.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                Reset()
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtAdmissionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve hosteler info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbInstallment.Text)) = 0 Then
            MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbInstallment.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFine.Text)) = 0 Then
            MessageBox.Show("Please enter fine", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFine.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbPaymentMode.Text)) = 0 Then
            MessageBox.Show("Please select Payment Mode", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbPaymentMode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTotalPaid.Text)) = 0 Then
            MessageBox.Show("Please enter total paid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTotalPaid.Focus()
            Exit Sub
        End If
        If Val(txtTotalPaid.Text) < 0 Then
            MessageBox.Show("Total paid must be greater than zero", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTotalPaid.Focus()
            Exit Sub
        End If
        If Val(txtBalance.Text) < 0 Then
            MessageBox.Show("Balance is not possible less than zero", "Input error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Try
            'con = New SqlConnection(cs)
            'con.Open()
            'Dim ct As String = "select PaymentDate from hostelFeePayment where AdmissionNo=@d1"
            'cmd = New SqlCommand(ct)
            'cmd.Connection = con
            'cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            'Dim da As New SqlDataAdapter(cmd)
            'Dim ds As DataSet = New DataSet()
            'da.Fill(ds)
            'If ds.Tables(0).Rows.Count > 0 Then
            'If dtpPaymentDate.Value.Date < ds.Tables(0).Rows(0)("PaymentDate") Then
            'MessageBox.Show("updating old record is not allowed when student has been already paid fee again", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'End If
            'Exit Sub
            'End If
            'con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update hostelFeePayment set HFP_ID=@d1, PaymentID=@d2, HostelerID=@d3, Session=@d4, Installment=@d5,TotalFee=@d6, DiscountPer=@d7, DiscountAmt=@d8, PreviousDue=@d9, Fine=@d10, GrandTotal=@d11, TotalPaid=@d12, ModeOfPayment=@d13, PaymentModeDetails=@d14, PaymentDue=@d16 where HFP_ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", Val(txtHFPId.Text))
            cmd.Parameters.AddWithValue("@d2", txtFeePaymentID.Text)
            cmd.Parameters.AddWithValue("@d3", txtHostelerID.Text)
            cmd.Parameters.AddWithValue("@d4", txtSession.Text)
            cmd.Parameters.AddWithValue("@d5", cmbInstallment.Text)
            cmd.Parameters.AddWithValue("@d6", txthostelFee.Text)
            cmd.Parameters.AddWithValue("@d7", txtDiscountPer.Text)
            cmd.Parameters.AddWithValue("@d8", txtDiscount.Text)
            cmd.Parameters.AddWithValue("@d9", txtPreviousDue.Text)
            cmd.Parameters.AddWithValue("@d10", txtFine.Text)
            cmd.Parameters.AddWithValue("@d11", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d12", txtTotalPaid.Text)
            cmd.Parameters.AddWithValue("@d13", cmbPaymentMode.Text)
            cmd.Parameters.AddWithValue("@d14", txtPaymentModeDetails.Text)
            cmd.Parameters.AddWithValue("@d16", txtBalance.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            LedgerUpdate(dtpPaymentDate.Value.Date, txtStudentName.Text, Val(txtGrandTotal.Text), 0, txtAdmissionNo.Text, txtFeePaymentID.Text, "Hostel Fee Payment")
            If cmbPaymentMode.Text = "By Cash" Then
                LedgerUpdate(dtpPaymentDate.Value.Date, "Cash Account", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text, txtFeePaymentID.Text, "Payment")
            End If
            If cmbPaymentMode.Text = "By Cheque" Or cmbPaymentMode.Text = "By DD" Then
                LedgerUpdate(dtpPaymentDate.Value.Date, "bank Account", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text, txtFeePaymentID.Text, "Payment")
            End If
            Dim st As String = "updated the hostel fee payment entry having payment id '" & txtFeePaymentID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtFine_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtFine.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtFine.Text
            Dim selectionStart = Me.txtFine.SelectionStart
            Dim selectionLength = Me.txtFine.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtTotalPaid_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtTotalPaid.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtTotalPaid.Text
            Dim selectionStart = Me.txtTotalPaid.SelectionStart
            Dim selectionLength = Me.txtTotalPaid.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtTotalPaid_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles txtTotalPaid.Validating
        If Val(txtTotalPaid.Text) > Val(txtGrandTotal.Text) Then
            MessageBox.Show("Total Pay can not be more than grand total", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotalPaid.Text = ""
            txtTotalPaid.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub txtFine_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtFine.TextChanged
        Fill()
    End Sub

    Private Sub txtTotalPaid_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtTotalPaid.TextChanged
        Fill()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmHostelFeePaymentRecord.lblSet.Text = "Hostel Fee Payment"
        frmHostelFeePaymentRecord.lblUserType.Text = lblUserType.Text
        frmHostelFeePaymentRecord.Reset()
        frmHostelFeePaymentRecord.ShowDialog()
    End Sub
    Sub fillInstallment()
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Installment) FROM Installment_Hostel,HostelInfo where Installment_Hostel.HostelID=HostelInfo.HI_ID and HostelName=@d1"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", txtHostelName.Text)
            cmd.Connection = con
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

    Sub Print()
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptHostelFeeReceipt 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT SchoolInfo.S_Id, HostelFeePayment.Session, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName,Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.ContactNo AS Expr1, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Caste,Student.Religion, Student.SectionID, Student.Photo, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result, Student.PassPercentage, Student.Status, Section.Id, Section.SectionName, Section.Class, Class.ClassName, Class.ClassType, Hosteler.H_Id, Hosteler.AdmissionNo AS Expr2, Hosteler.HostelID, Hosteler.JoiningDate, Hosteler.Status AS Expr3, HostelInfo.HI_Id, HostelInfo.Hostelname,HostelInfo.Address AS Expr4, HostelInfo.ContactNo AS Expr5, HostelInfo.ManagedBy, HostelInfo.Person_ContactNo, HostelFeePayment.HFP_Id, HostelFeePayment.PaymentID, HostelFeePayment.HostelerID, HostelFeePayment.Installment, HostelFeePayment.TotalFee, HostelFeePayment.DiscountPer, HostelFeePayment.DiscountAmt, HostelFeePayment.PreviousDue,HostelFeePayment.Fine, HostelFeePayment.GrandTotal, HostelFeePayment.TotalPaid, HostelFeePayment.ModeOfPayment, HostelFeePayment.PaymentModeDetails, HostelFeePayment.Paymentdate, HostelFeePayment.PaymentDue FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN Section ON Student.SectionID = Section.Id INNER JOIN Class ON Section.Class = Class.ClassName INNER JOIN Hosteler ON Student.AdmissionNo = Hosteler.AdmissionNo INNER JOIN HostelInfo ON Hosteler.HostelID = HostelInfo.HI_Id INNER JOIN HostelFeePayment ON Hosteler.H_Id = HostelFeePayment.HostelerID where PaymentID='" & txtFeePaymentID.Text & "'"
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
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub frmHostelFeePayment_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
