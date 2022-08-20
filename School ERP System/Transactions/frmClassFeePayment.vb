Imports System.Data.SqlClient
Imports System.IO

Public Class frmClassFeePayment
    Dim st2 As String
    Sub Calculate()
        Dim num1, num2 As Double
        num1 = CDbl(Val(txtCourseFee.Text) + Val(txtFine.Text) + Val(txtPreviousDue.Text) - Val(txtDiscount.Text))
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
            cmd = New SqlCommand("SELECT TOP 1 (CFP_ID) FROM CourseFeePayment ORDER BY CFP_ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("CFP_ID")
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
            txtCFPId.Text = GenerateID()
            Dim a As String = txtAdmissionNo.Text
            txtFeePaymentID.Text = "CFP-" + GenerateID() + "-" + a
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Sub Reset()
        txtAdmissionNo.Text = ""
        txtCourseFee.Text = ""
        txtPaymentModeDetails.Text = ""
        txtBalance.Text = ""
        txtClass.Text = ""
        txtDiscount.Text = "0.00"
        txtDiscountPer.Text = "0.00"
        txtEnrollmentNo.Text = ""
        txtFine.Text = "0.00"
        txtGrandTotal.Text = ""
        txtSchoolName.Text = ""
        txtSection.Text = ""
        txtSession.Text = ""
        txtStudentName.Text = ""
        txtContactNo.Text = ""
        txtTotalPaid.Text = ""
        cmbPaymentMode.SelectedIndex = 0
        dtpPaymentDate.Value = Now
        dgw.Rows.Clear()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        btnPrint.Enabled = False
        btnStudentSelection.Enabled = True
        btnRemove.Enabled = False
        dtpPaymentDate.Enabled = True
        btnStudentSelection.Enabled = True
        gpbFeeInfo.Enabled = True
        DataGridView1.Rows.Clear()
    End Sub
    Sub Fill()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Discount from Discount where AdmissionNo=@d1 and FeeType='Class'"
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtDiscountPer.Text = rdr.GetValue(0)
                Dim num As Double
                num = CDbl((Val(txtCourseFee.Text) * Val(txtDiscountPer.Text)) / 100)
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
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles btnStudentSelection.Click
        frmActiveStudentRecord.lblSet.Text = "Course Fee Payment"
        frmActiveStudentRecord.Reset()
        frmActiveStudentRecord.ShowDialog()
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
        Reset()
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click



        'dgw.Rows.Clear()
        'dgw.Refresh()
        'Dim mycount As Integer
        'mycount = dgw.Rows.Count - 1
        'If dgw.Rows.Count < 1 Then
        '    btnRemove.Enabled = True
        'End If

        'If (dgw.Rows.Count <= 1) Then
        '    dgw.AllowUserToAddRows = False
        '    dgw.Rows.RemoveAt(dgw.CurrentRow.Index)
        '    dgw.Rows.RemoveAt(dgw.CurrentRow.Index)
        '    dgw.AllowUserToAddRows = True
        'Else
        '    dgw.Rows.RemoveAt(dgw.CurrentRow.Index)
        'End If

        If Len(Trim(txtAdmissionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve student info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        If dgw.Rows.Count = 0 Then
            MessageBox.Show("Please add fee info in a datagrid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            For Each r As DataGridViewRow In dgw.Rows
                con = New SqlConnection(cs)
                con.Open()
                Dim ct As String = "select Session,AdmissionNo, FeeName from CourseFeePayment,CourseFeePayment_Join where CourseFeePayment.CFP_ID=CourseFeePayment_Join.C_PaymentID and Session=@d1 and AdmissionNo=@d2 and Feename=@d3 and Student_Class=@d4"
                cmd = New SqlCommand(ct)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", txtSession.Text)
                cmd.Parameters.AddWithValue("@d2", txtAdmissionNo.Text)
                cmd.Parameters.AddWithValue("@d3", r.Cells(0).Value.ToString())
                cmd.Parameters.AddWithValue("@d4", txtClass.Text)
                rdr = cmd.ExecuteReader()

            Next
            auto()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into CourseFeePayment(CFP_ID, PaymentID, AdmissionNo, Session,TotalFee, DiscountPer, DiscountAmt,PreviousDue, Fine, GrandTotal, TotalPaid, ModeOfPayment, PaymentModeDetails, PaymentDate, PaymentDue,Student_Class) VALUES (@d1,@d2,@d3,@d4,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17)"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtCFPId.Text)
            cmd.Parameters.AddWithValue("@d2", txtFeePaymentID.Text)
            cmd.Parameters.AddWithValue("@d3", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d4", txtSession.Text)
            cmd.Parameters.AddWithValue("@d6", Val(txtCourseFee.Text))
            cmd.Parameters.AddWithValue("@d7", Val(txtDiscountPer.Text))
            cmd.Parameters.AddWithValue("@d8", Val(txtDiscount.Text))
            cmd.Parameters.AddWithValue("@d9", Val(txtPreviousDue.Text))
            cmd.Parameters.AddWithValue("@d10", Val(txtFine.Text))
            cmd.Parameters.AddWithValue("@d11", Val(txtGrandTotal.Text))
            cmd.Parameters.AddWithValue("@d12", Val(txtTotalPaid.Text))
            cmd.Parameters.AddWithValue("@d13", cmbPaymentMode.Text)
            cmd.Parameters.AddWithValue("@d14", txtPaymentModeDetails.Text)
            cmd.Parameters.AddWithValue("@d15", dtpPaymentDate.Value)
            cmd.Parameters.AddWithValue("@d16", Val(txtBalance.Text))
            cmd.Parameters.AddWithValue("@d17", txtClass.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into CourseFeePayment_Join(C_PaymentID,FeeName, Fee) VALUES (" & txtCFPId.Text & ",@d1,@d2)"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d2", Val(row.Cells(1).Value))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            LedgerSave(dtpPaymentDate.Value.Date, txtStudentName.Text, txtFeePaymentID.Text, "Class Fee Payment", Val(txtGrandTotal.Text), 0, txtAdmissionNo.Text)
            If cmbPaymentMode.Text = "By Cash" Then
                LedgerSave(dtpPaymentDate.Value.Date, "Cash Account", txtFeePaymentID.Text, "Payment", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text)
            End If
            If cmbPaymentMode.Text = "By Cheque" Or cmbPaymentMode.Text = "By DD" Then
                LedgerSave(dtpPaymentDate.Value.Date, "bank Account", txtFeePaymentID.Text, "Payment", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text)
            End If
            Dim st As String = "added the new course fee payment entry having payment id '" & txtFeePaymentID.Text & "'"
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
                    Dim st3 As String = "Hello, " & txtStudentName.Text & " you have successfully paid class fee having Fee Payment ID '" & txtFeePaymentID.Text & """"
                    SMSFunc(txtContactNo.Text, st3, st2)
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                End If
            End If
            MessageBox.Show("Successfully paid", "Fee", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            con.Close()
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
            Dim cq As String = "delete from CourseFeePayment where CFP_ID= " & txtCFPId.Text & ""
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LedgerDelete(txtFeePaymentID.Text, "Class Fee Payment")
                LedgerDelete(txtFeePaymentID.Text, "Payment")
                Dim st As String = "deleted the course fee payment entry having payment id '" & txtFeePaymentID.Text & "'"
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
            MessageBox.Show("Please retrieve student info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        If dgw.Rows.Count = 0 Then
            MessageBox.Show("Please add fee info in a datagrid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            'Dim ct As String = "select PaymentDate from CourseFeePayment where AdmissionNo=@d1"
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
            Dim cb As String = "Update CourseFeePayment set CFP_ID=@d1, PaymentID=@d2, AdmissionNo=@d3, Session=@d4,TotalFee=@d6, DiscountPer=@d7, DiscountAmt=@d8,PreviousDue=@d9, Fine=@d10, GrandTotal=@d11, TotalPaid=@d12, ModeOfPayment=@d13, PaymentModeDetails=@d14, PaymentDue=@d16 where CfP_ID= " & txtCFPId.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", Val(txtCFPId.Text))
            cmd.Parameters.AddWithValue("@d2", txtFeePaymentID.Text)
            cmd.Parameters.AddWithValue("@d3", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d4", txtSession.Text)
            cmd.Parameters.AddWithValue("@d6", Val(txtCourseFee.Text))
            cmd.Parameters.AddWithValue("@d7", Val(txtDiscountPer.Text))
            cmd.Parameters.AddWithValue("@d8", Val(txtDiscount.Text))
            cmd.Parameters.AddWithValue("@d9", Val(txtPreviousDue.Text))
            cmd.Parameters.AddWithValue("@d10", Val(txtFine.Text))
            cmd.Parameters.AddWithValue("@d11", Val(txtGrandTotal.Text))
            cmd.Parameters.AddWithValue("@d12", Val(txtTotalPaid.Text))
            cmd.Parameters.AddWithValue("@d13", cmbPaymentMode.Text)
            cmd.Parameters.AddWithValue("@d14", txtPaymentModeDetails.Text)
            cmd.Parameters.AddWithValue("@d16", Val(txtBalance.Text))
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from CourseFeePayment_Join where C_PaymentID= " & txtCFPId.Text & ""
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con.Open()
            Dim cb1 As String = "insert into CourseFeePayment_Join(C_PaymentID,Month, FeeName, Fee) VALUES (" & txtCFPId.Text & ",@d1,@d2,@d3)"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d2", row.Cells(1).Value)
                    cmd.Parameters.AddWithValue("@d3", Val(row.Cells(2).Value))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            LedgerUpdate(dtpPaymentDate.Value.Date, txtStudentName.Text, Val(txtGrandTotal.Text), 0, txtAdmissionNo.Text, txtFeePaymentID.Text, "Class Fee Payment")
            If cmbPaymentMode.Text = "By Cash" Then
                LedgerUpdate(dtpPaymentDate.Value.Date, "Cash Account", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text, txtFeePaymentID.Text, "Payment")
            End If
            If cmbPaymentMode.Text = "By Cheque" Or cmbPaymentMode.Text = "By DD" Then
                LedgerUpdate(dtpPaymentDate.Value.Date, "bank Account", 0, Val(txtTotalPaid.Text), txtAdmissionNo.Text, txtFeePaymentID.Text, "Payment")
            End If
            Dim st As String = "updated the course fee payment entry having payment id '" & txtFeePaymentID.Text & "'"
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
        frmClassFeePaymentRecord.lblSet.Text = "Course Fee Payment"
        frmClassFeePaymentRecord.lblUserType.Text = lblUserType.Text
        frmClassFeePaymentRecord.Reset()
        frmClassFeePaymentRecord.ShowDialog()
    End Sub
    Sub Print()
        Try
            'Cursor = Cursors.WaitCursor
            'Timer1.Enabled = True
            'Dim rpt As New rptCourseFeeReceipt 'The report you created.
            'Dim myConnection As SqlConnection
            'Dim MyCommand As New SqlCommand()
            'Dim myDA As New SqlDataAdapter()
            'Dim myDS As New DataSet 'The DataSet you created.
            'myConnection = New SqlConnection(cs)
            'MyCommand.Connection = myConnection
            'MyCommand.CommandText =
            'MyCommand.CommandType = CommandType.Text
            'myDA.SelectCommand = MyCommand
            'myDA.Fill(myDS, "SchoolInfo")
            'myDA.Fill(myDS, "Student")
            'myDA.Fill(myDS, "CourseFeePayment")
            'myDA.Fill(myDS, "CourseFeePayment_Join")
            'myDA.Fill(myDS, "Section")
            'myDA.Fill(myDS, "Class")
            'rpt.SetDataSource(myDS)
            'frmReport.CrystalReportViewer1.ReportSource = rpt
            'frmReport.ShowDialog()
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            'old cmd = New SqlCommand("SELECT SchoolInfo.S_Id, CourseFeePayment.Session, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName,Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.ContactNo AS Expr1, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Caste,Student.Religion, Student.SectionID, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result, Student.PassPercentage, Student.Status, Section.Id, Section.SectionName,Section.Class, Class.ClassName, Class.ClassType, CourseFeePayment.CFP_ID, CourseFeePayment.PaymentID, CourseFeePayment.AdmissionNo AS Expr2,CourseFeePayment.TotalFee, CourseFeePayment.DiscountPer, CourseFeePayment.DiscountAmt, CourseFeePayment.PreviousDue, CourseFeePayment.Fine, CourseFeePayment.GrandTotal,CourseFeePayment.TotalPaid, CourseFeePayment.ModeOfPayment, CourseFeePayment.PaymentModeDetails, CourseFeePayment.PaymentDate, CourseFeePayment.PaymentDue,CourseFeePayment_Join.CFPJ_Id, CourseFeePayment_Join.C_PaymentID, CourseFeePayment_Join.FeeName,  CourseFeePayment_Join.Fee FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN Section ON Student.SectionID = Section.Id INNER JOIN Class ON Section.Class = Class.ClassName INNER JOIN CourseFeePayment ON Student.AdmissionNo = CourseFeePayment.AdmissionNo INNER JOIN CourseFeePayment_Join ON CourseFeePayment.CFP_ID = CourseFeePayment_Join.C_PaymentID where CourseFeePayment.PaymentID='" & txtFeePaymentID.Text & "'", con)
            cmd = New SqlCommand("SELECT SchoolInfo.S_Id, CourseFeePayment.Session, SchoolInfo.SchoolName, SchoolInfo.Address, SchoolInfo.ContactNo, SchoolInfo.AltContactNo, SchoolInfo.FaxNo, SchoolInfo.Email, SchoolInfo.Website, SchoolInfo.Logo, SchoolInfo.RegistrationNo,SchoolInfo.EstablishedYear, SchoolInfo.SchoolType, Student.AdmissionNo, Student.EnrollmentNo, Student.GRNo, Student.UID, Student.StudentName, Student.FatherName, Student.MotherName,Student.FatherCN, Student.PermanentAddress, Student.TemporaryAddress, Student.ContactNo AS Expr1, Student.EmailID, Student.DOB, Student.Gender, Student.AdmissionDate, Student.Caste,Student.Religion, Student.SectionID, Student.Nationality, Student.SchoolID, Student.LastSchoolAttended, Student.Result, Student.PassPercentage, Student.Status, Section.Id, Section.SectionName,Section.Class, Class.ClassName, Class.ClassType, CourseFeePayment.CFP_ID, CourseFeePayment.PaymentID, CourseFeePayment.AdmissionNo AS Expr2, CourseFeePayment_Join.Fee as TotalFee, CourseFeePayment.DiscountPer, CourseFeePayment.DiscountAmt, CourseFeePayment.PreviousDue, CourseFeePayment.Fine, CourseFeePayment.GrandTotal,CourseFeePayment.TotalPaid, CourseFeePayment.ModeOfPayment, CourseFeePayment.PaymentModeDetails, CourseFeePayment.PaymentDate, CourseFeePayment.PaymentDue,CourseFeePayment_Join.CFPJ_Id, CourseFeePayment_Join.C_PaymentID, CourseFeePayment_Join.FeeName,  CourseFeePayment_Join.Fee FROM SchoolInfo INNER JOIN Student ON SchoolInfo.S_Id = Student.SchoolID INNER JOIN Section ON Student.SectionID = Section.Id INNER JOIN Class ON Section.Class = Class.ClassName INNER JOIN CourseFeePayment ON Student.AdmissionNo = CourseFeePayment.AdmissionNo INNER JOIN CourseFeePayment_Join ON CourseFeePayment.CFP_ID = CourseFeePayment_Join.C_PaymentID where CourseFeePayment.PaymentID='" & txtFeePaymentID.Text & "'", con)

            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("ClassFeeReceipt.xml")
            Dim rpt As New rptCourseFeeReceipt
            rpt.SetDataSource(ds)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmCourseFeePayment_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Print()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub txtCourseFee_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtCourseFee.TextChanged
        Calculate()
    End Sub
    Sub FillFee()


        Try
            'gpbFeeInfo.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim sql1 As String = "Select distinct RTRIM(Feename),Fee from CourseFeePayment,CourseFeePayment_Join where CourseFeePayment.CFP_ID=CourseFeePayment_Join.C_PaymentID and AdmissionNo=@d1 and Student_Class=@d2 order by 1"
            cmd = New SqlCommand(sql1, con)
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d2", txtClass.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView1.Rows.Clear()

            While rdr.Read
                dgw.Rows.Add(rdr(0), rdr(1))
            End While
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Sum(PaymentDue-PreviousDue) from CourseFeePayment where AdmissionNo=@d1 group by AdmissionNo"
            cmd.Parameters.AddWithValue("@d1", txtAdmissionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtPreviousDue.Text = rdr.GetValue(0)
            Else
                txtPreviousDue.Text = "0.00"
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = "Select distinct RTRIM(Feename) from CourseFee,SchoolInfo where CourseFee.SchoolID=SchoolInfo.S_ID and SchoolName=@d1 and Class=@d2  and CourseFee.FeeName not in (Select distinct Feename from CourseFeePayment,CourseFeePayment_Join where CourseFeePayment.CFP_ID=CourseFeePayment_join.C_PaymentID and AdmissionNo=@d3 and Student_Class=@d2) order by 1"
            cmd = New SqlCommand(sql, con)
            cmd.Parameters.AddWithValue("@d1", txtSchoolName.Text)
            cmd.Parameters.AddWithValue("@d2", txtClass.Text)
            cmd.Parameters.AddWithValue("@d3", txtAdmissionNo.Text)
            ' cmd.Parameters.AddWithValue("@d4", cmbFee.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)


            cmbFee.Items.Clear()
            While (rdr.Read() = True)
                cmbFee.Items.Add(rdr(0))
            End While
            con.Close()
            Fill()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        Try
            For Each row As DataGridViewRow In dgw.SelectedRows
                dgw.Rows.Remove(row)
            Next
            Dim sum As Double = 0
            For Each r As DataGridViewRow In Me.dgw.Rows
                sum = sum + r.Cells(1).Value
            Next
            sum = Math.Round(sum, 2)
            txtCourseFee.Text = sum
            btnRemove.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        If dgw.Rows.Count > 0 Then
            btnRemove.Enabled = True
        End If
    End Sub

    Private Sub cmbFee_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbFee.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = "Select Fee from CourseFee,SchoolInfo where CourseFee.SchoolID=SchoolInfo.S_ID and SchoolName=@d1 and Class=@d2 and FeeName=@d3 order by 1"
            cmd = New SqlCommand(sql, con)
            cmd.Parameters.AddWithValue("@d1", txtSchoolName.Text)
            cmd.Parameters.AddWithValue("@d2", txtClass.Text)
            cmd.Parameters.AddWithValue("@d3", cmbFee.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If (rdr.Read() = True) Then
                txtFee.Text = rdr.GetValue(0)
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        Try
            If cmbFee.Text = "" Then
                MessageBox.Show("Please select fee name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbFee.Focus()
                Exit Sub
            End If
            For Each dr As DataGridViewRow In dgw.Rows
                If cmbFee.Text = dr.Cells(0).Value.ToString() Then
                    MessageBox.Show("Same fee name is already added to grid", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    cmbFee.SelectedIndex = -1
                    txtFee.Text = ""
                    cmbFee.Focus()
                    Exit Sub
                End If
            Next
            dgw.Rows.Add(cmbFee.Text, Val(txtFee.Text))
            Dim sum As Double = 0
            For Each r As DataGridViewRow In Me.dgw.Rows
                sum = sum + r.Cells(1).Value
            Next
            sum = Math.Round(sum, 2)
            txtCourseFee.Text = sum
            cmbFee.SelectedIndex = -1
            txtFee.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
