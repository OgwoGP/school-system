Imports System.Data.SqlClient
Public Class frmBookIssue
    Dim count As Integer
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmBookEntryRecord1.lblSet.Text = "Book Issue"
        frmBookEntryRecord1.Reset()
        frmBookEntryRecord1.ShowDialog()
    End Sub
    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM BookIssue_Student")
            cmd = New SqlCommand(sql)
            cmd.Connection = con
            If (IsDBNull(cmd.ExecuteScalar)) Then
                Num = 1
                txtID.Text = Num.ToString
            Else
                Num = cmd.ExecuteScalar + 1
                txtID.Text = Num.ToString
            End If
            cmd.Dispose()
            con.Close()
            con.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub auto1()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM BookIssue_Staff")
            cmd = New SqlCommand(sql)
            cmd.Connection = con
            If (IsDBNull(cmd.ExecuteScalar)) Then
                Num = 1
                txtID1.Text = Num.ToString
            Else
                Num = cmd.ExecuteScalar + 1
                txtID1.Text = Num.ToString
            End If
            cmd.Dispose()
            con.Close()
            con.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        frmActiveStudentRecord.lblSet.Text = "Book Issue"
        frmActiveStudentRecord.Reset()
        frmActiveStudentRecord.ShowDialog()
    End Sub

    Private Sub dtpIssueDate_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpIssueDate.ValueChanged
        dtpDueDate.Text = dtpIssueDate.Value.Date.AddDays(Val(txtMaxDay_Student.Text))
    End Sub
    Sub FillData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT [Section] FROM Book where AccessionNo=@d1"
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtBookType.Text = rdr.GetValue(0)
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
            cmd.CommandText = "SELECT [Section] FROM Book where AccessionNo=@d1"
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtBookType.Text = rdr.GetValue(0)
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
            cmd.CommandText = "SELECT MaxDays_Staff, MaxDays_Student,MaxBooks_Staff,MaxBooks_Student FROM Setting where BookType=@d1"
            cmd.Parameters.AddWithValue("@d1", txtBookType.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtMaxDay_Staff.Text = rdr.GetValue(0)
                txtMaxDay_Student.Text = rdr.GetValue(1)
                txtMaxBooksAllowedStaff.Text = rdr.GetValue(2)
                txtMaxBooksAllowedStudent.Text = rdr.GetValue(3)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            dtpDueDate.Text = dtpIssueDate.Value.Date.AddDays(Val(txtMaxDay_Student.Text))
            dtpDueDate1.Text = dtpIssueDate1.Value.Date.AddDays(Val(txtMaxDay_Staff.Text))
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub frmBookIssue_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
    Sub Reset()
        txtAccessionNo.Text = ""
        txtAdmissionNo.Text = ""
        txtAuthor.Text = ""
        txtBookTitle.Text = ""
        txtClass.Text = ""
        txtJointAuthor.Text = ""
        txtRemarks.Text = ""
        txtStudentName.Text = ""
        dtpIssueDate.Text = Today
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        Button1.Enabled = True
        Button2.Enabled = True
        auto()
        dtpDueDate.Text = dtpIssueDate.Value.Date.AddDays(Val(txtMaxDay_Student.Text))
    End Sub
    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtAccessionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAdmissionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve admssion no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Count(AccessionNo) FROM BookIssue_Student where AccessionNo=@d1 and Status='Issued'"
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                count = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            If count >= Val(txtMaxBooksAllowedStudent.Text) Then
                MessageBox.Show("Maximum no. of books already issued", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into BookIssue_Student(ID,AccessionNo,AdmissionNo,IssueDate,Duedate, Status, Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,'Issued',@d6)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtAccessionNo.Text)
            cmd.Parameters.AddWithValue("@d3", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d4", dtpIssueDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", dtpDueDate.Value.Date)
            cmd.Parameters.AddWithValue("@d6", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "Update Book Set Status='Issued' where AccessionNo=@d1"
            cmd = New SqlCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "issued the Book for student '" & txtStudentName.Text & "' having Issue ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully issued", "Book", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DeleteRecord()

        Try
            Dim RowsAffected As Integer = 0
            If (txtStatus.Text = "Issued") Then
                MessageBox.Show("Book is issued..Record can not be deleted", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "select BookIssue_Student.ID from BookIssue_Student,Return_Student where BookIssue_Student.ID=Return_Student.IssueID and BookIssue_Student.ID=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Book Return", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from BookIssue_Student where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the issued book record for student '" & txtStudentName.Text & "' having Issue ID='" & txtID.Text & "'"
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            Else
                MessageBox.Show("No Record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtAccessionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAdmissionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve admssion no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAdmissionNo.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "Update BookIssue_Student set AccessionNo=@d2,AdmissionNo=@d3,IssueDate=@d4,Duedate=@d5,Remarks=@d6 where ID=" & txtID.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtAccessionNo.Text)
            cmd.Parameters.AddWithValue("@d3", txtAdmissionNo.Text)
            cmd.Parameters.AddWithValue("@d4", dtpIssueDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", dtpDueDate.Value.Date)
            cmd.Parameters.AddWithValue("@d6", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "updated the issued book record of student '" & txtStudentName.Text & "' having Issue ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmBookIssueRecord_Student.lblSet.Text = "Book Issue"
        frmBookIssueRecord_Student.Reset()
        frmBookIssueRecord_Student.ShowDialog()
    End Sub

    Private Sub btnClose1_Click(sender As System.Object, e As System.EventArgs) Handles btnClose1.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        frmBookEntryRecord1.lblSet.Text = "Book Issue Entry"
        frmBookEntryRecord1.Reset()
        frmBookEntryRecord1.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        frmActiveStaffRecord.lblSet.Text = "Book Issue Entry"
        frmActiveStaffRecord.Reset()
        frmActiveStaffRecord.ShowDialog()
    End Sub
    Sub Reset1()
        txtAccessionNo1.Text = ""
        txtStaffID.Text = ""
        txtAuthor1.Text = ""
        txtBookTitle1.Text = ""
        txtJointAuthors1.Text = ""
        txtRemarks1.Text = ""
        txtStaffName.Text = ""
        btnIssue_Reservation.Enabled = False
        dtpIssueDate1.Text = Today
        btnSave1.Enabled = True
        btnDelete1.Enabled = False
        btnUpdate1.Enabled = False
        Button3.Enabled = True
        Button4.Enabled = True
        auto1()
    End Sub
    Private Sub btnNew1_Click(sender As System.Object, e As System.EventArgs) Handles btnNew1.Click
        Reset1()
    End Sub

    Private Sub dtpIssueDate1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpIssueDate1.ValueChanged
        dtpDueDate1.Text = dtpIssueDate1.Value.Date.AddDays(Val(txtMaxDay_Staff.Text))
    End Sub

    Private Sub btnSave1_Click(sender As System.Object, e As System.EventArgs) Handles btnSave1.Click
        If Len(Trim(txtAccessionNo1.Text)) = 0 Then
            MessageBox.Show("Please retrieve accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo1.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStaffID.Text)) = 0 Then
            MessageBox.Show("Please retrieve staff id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStaffID.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Count(AccessionNo) FROM BookIssue_Staff where AccessionNo=@d1 and Status='Issued'"
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                count = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            If count >= Val(txtMaxBooksAllowedStaff.Text) Then
                MessageBox.Show("Maximum no. of books already issued", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into BookIssue_Staff(ID,AccessionNo,StaffID,IssueDate,Duedate, Status, Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,'Issued',@d6)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID1.Text)
            cmd.Parameters.AddWithValue("@d2", txtAccessionNo1.Text)
            cmd.Parameters.AddWithValue("@d3", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d4", dtpIssueDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d5", dtpDueDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d6", txtRemarks1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "Update Book Set Status='Issued' where AccessionNo=@d1"
            cmd = New SqlCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "issued the Book for staff '" & txtStaffName.Text & "' having Issue ID='" & txtID1.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully issued", "Book", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave1.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnDelete1_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete1.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord1()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DeleteRecord1()

        Try
            Dim RowsAffected As Integer = 0
            'If (txtStatus.Text = "Issued") Then
            '    MessageBox.Show("Book is issued..Record can not be deleted", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    Exit Sub
            'End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "select BookIssue_Staff.ID from BookIssue_Staff,Return_Staff where BookIssue_Staff.ID=Return_Staff.IssueID and BookIssue_Staff.ID=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Book Return", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from BookIssue_Staff where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID1.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the issued book record for staff '" & txtStaffName.Text & "' having Issue ID='" & txtID1.Text & "'"
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset1()
            Else
                MessageBox.Show("No Record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset1()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnUpdate1_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate1.Click
        If Len(Trim(txtAccessionNo1.Text)) = 0 Then
            MessageBox.Show("Please retrieve accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo1.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStaffID.Text)) = 0 Then
            MessageBox.Show("Please retrieve staff id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStaffID.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "Update BookIssue_Staff set AccessionNo=@d2,StaffID=@d3,IssueDate=@d4,Duedate=@d5,Remarks=@d6 where ID=" & txtID.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtAccessionNo1.Text)
            cmd.Parameters.AddWithValue("@d3", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d4", dtpIssueDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d5", dtpDueDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d6", txtRemarks1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "updated the issued book record of staff '" & txtStaffName.Text & "' having Issue ID='" & txtID1.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate1.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnGetData1_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData1.Click
        frmBookIssueRecord_Staff.lblSet.Text = "Book Issue"
        frmBookIssueRecord_Staff.Reset()
        frmBookIssueRecord_Staff.ShowDialog()
    End Sub

    Private Sub btnIssue_Reservation_Click(sender As System.Object, e As System.EventArgs) Handles btnIssue_Reservation.Click
        If Len(Trim(txtAccessionNo1.Text)) = 0 Then
            MessageBox.Show("Please retrieve accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo1.Focus()
            Exit Sub
        End If
        If Len(Trim(txtStaffID.Text)) = 0 Then
            MessageBox.Show("Please retrieve staff id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStaffID.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT Count(AccessionNo) FROM BookIssue_Staff where AccessionNo=@d1 and Status='Issued'"
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                count = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            If count >= Val(txtMaxBooksAllowedStaff.Text) Then
                MessageBox.Show("Maximum no. of books already issued", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into BookIssue_Staff(ID,AccessionNo,StaffID,IssueDate,Duedate, Status, Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,'Issued',@d6)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID1.Text)
            cmd.Parameters.AddWithValue("@d2", txtAccessionNo1.Text)
            cmd.Parameters.AddWithValue("@d3", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d4", dtpIssueDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d5", dtpDueDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d6", txtRemarks1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "Update Book Set Status='Issued' where AccessionNo=@d1"
            cmd = New SqlCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb2 As String = "Update BookReservation Set Status='Issued' where AccessionNo=@d1 and ID=@d2"
            cmd = New SqlCommand(cb2)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            cmd.Parameters.AddWithValue("@d2", txtReservationID.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb3 As String = "Insert into Book_RI(IssueID,ReservationID) values(@d1,@d2)"
            cmd = New SqlCommand(cb3)
            cmd.Parameters.AddWithValue("@d1", txtID1.Text)
            cmd.Parameters.AddWithValue("@d2", txtReservationID.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "issued the Book for staff '" & txtStaffName.Text & "' having Issue ID='" & txtID1.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully issued", "Book", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave1.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        frmBookReservationRecord1.lblSet.Text = "Book Issue"
        frmBookReservationRecord1.Reset()
        frmBookReservationRecord1.ShowDialog()
    End Sub

    Private Sub Button4_MouseHover(sender As System.Object, e As System.EventArgs) Handles Button4.MouseHover
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(Button4, "Retrieve book info from list of books")
    End Sub

    Private Sub Button5_MouseHover(sender As System.Object, e As System.EventArgs) Handles Button5.MouseHover
        ToolTip1.IsBalloon = True
        ToolTip1.UseAnimation = True
        ToolTip1.ToolTipTitle = ""
        ToolTip1.SetToolTip(Button5, "Retrieve book info and staff info from list of books reservation")
    End Sub
End Class
