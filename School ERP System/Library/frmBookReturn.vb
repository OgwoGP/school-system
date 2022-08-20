Imports System.Data.SqlClient
Public Class frmBookReturn
    Dim count As Integer
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmBookIssueRecord_Student1.lblSet.Text = "Book Return"
        frmBookIssueRecord_Student1.Reset()
        frmBookIssueRecord_Student1.ShowDialog()
    End Sub
    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM Return_Student")
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
            Dim sql As String = ("SELECT MAX(ID) FROM Return_Staff")
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
            cmd.CommandText = "SELECT FinePerDay_Staff,FinePerDay_Student FROM Setting where BookType=@d1"
            cmd.Parameters.AddWithValue("@d1", txtBookType.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtFinePerDayStaff.Text = rdr.GetValue(0)
                txtFinePerDayStudent.Text = rdr.GetValue(1)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            If dtpReturnDate.Value.Date < dtpDueDate.Value.Date Then
                txtFine.Text = 0
            Else
                Dim st As String = dtpReturnDate.Value.Date.Subtract(dtpDueDate.Value.Date).ToString()
                txtFine.Text = Val(st) * Val(txtFinePerDayStudent.Text)
            End If
            If dtpReturnDate1.Value.Date < dtpDueDate1.Value.Date Then
                txtFine1.Text = 0
            Else
                Dim st As String = dtpReturnDate1.Value.Date.Subtract(dtpDueDate1.Value.Date).ToString()
                txtFine1.Text = Val(st) * Val(txtFinePerDayStaff.Text)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub frmBookIssue_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
    Sub Reset()
        txtAccessionNo.Text = ""
        dtpReturnDate.Text = Today
        txtFine.Text = ""
        txtAdmissionNo.Text = ""
        txtAuthor.Text = ""
        txtBookTitle.Text = ""
        txtClass.Text = ""
        txtJointAuthor.Text = ""
        txtRemarks.Text = ""
        txtStudentName.Text = ""
        dtpIssueDate.Text = Today
        dtpDueDate.Text = Today
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        Button2.Enabled = True
        auto()
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

            Dim cb As String = "insert into Return_Student(ID,IssueID,ReturnDate,Fine,Remarks) VALUES (@d1,@d2,@d3,@d4,@d5)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtIssueID_Student.Text)
            cmd.Parameters.AddWithValue("@d3", dtpReturnDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtFine.Text)
            cmd.Parameters.AddWithValue("@d5", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "Update Book Set Status='Available' where AccessionNo=@d1"
            cmd = New SqlCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb2 As String = "Update BookIssue_Student Set Status='Returned' where AccessionNo=@d1 and ID=@d2"
            cmd = New SqlCommand(cb2)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Parameters.AddWithValue("@d2", txtIssueID_Student.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            If Val(txtFine.Text) > 0 Then
                LedgerSave(dtpReturnDate.Value.Date, "Cash Account", txtID.Text, "Book Fine(Student)", 0, Val(txtFine.Text), txtAdmissionNo.Text)
            End If
            Dim st As String = "returned the Book of student '" & txtStudentName.Text & "' having ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully returned", "Book", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Return_Student where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LedgerDelete(txtID.Text, "Book Fine(Student)")
                Dim st As String = "deleted the returned book record for student '" & txtStudentName.Text & "' having ID='" & txtID.Text & "'"
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

            Dim cb As String = "Update Return_Student set IssueID=@d2,ReturnDate=@d3,Fine=@d4,Remarks=@d5 where ID=" & txtID.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtIssueID_Student.Text)
            cmd.Parameters.AddWithValue("@d3", dtpReturnDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtFine.Text)
            cmd.Parameters.AddWithValue("@d5", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            If Val(txtFine.Text) > 0 Then
                LedgerUpdate(dtpReturnDate.Value.Date, "Cash Account", 0, Val(txtFine.Text), txtAdmissionNo.Text, txtID.Text, "Book Fine(Student)")
            End If
            Dim st As String = "updated the returned book record of student '" & txtStudentName.Text & "' having ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmBookReturnRecord_Student.lblSet.Text = "Book Return"
        frmBookReturnRecord_Student.Reset()
        frmBookReturnRecord_Student.ShowDialog()
    End Sub

    Private Sub btnClose1_Click(sender As System.Object, e As System.EventArgs) Handles btnClose1.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        frmBookIssueRecord_Staff1.lblSet.Text = "Book Return"
        frmBookIssueRecord_Staff1.Reset()
        frmBookIssueRecord_Staff1.ShowDialog()
    End Sub

    Sub Reset1()
        txtAccessionNo1.Text = ""
        txtReservationID.Text = ""
        txtStaffID.Text = ""
        txtAuthor1.Text = ""
        txtBookTitle1.Text = ""
        txtJointAuthors1.Text = ""
        txtRemarks1.Text = ""
        txtStaffName.Text = ""
        dtpIssueDate1.Text = Today
        btnSave1.Enabled = True
        dtpReturnDate1.Text = Today
        txtFine1.Text = ""
        btnDelete1.Enabled = False
        btnUpdate1.Enabled = False
        dtpDueDate1.Text = Today
        Button4.Enabled = True
        auto1()
    End Sub
    Private Sub btnNew1_Click(sender As System.Object, e As System.EventArgs) Handles btnNew1.Click
        Reset1()
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

            Dim cb As String = "insert into Return_Staff(ID,IssueID,ReturnDate,Fine,Remarks) VALUES (@d1,@d2,@d3,@d4,@d5)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID1.Text)
            cmd.Parameters.AddWithValue("@d2", txtIssueID_Staff.Text)
            cmd.Parameters.AddWithValue("@d3", dtpReturnDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtFine1.Text)
            cmd.Parameters.AddWithValue("@d5", txtRemarks1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "Update Book Set Status='Available' where AccessionNo=@d1"
            cmd = New SqlCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb2 As String = "Update BookIssue_Staff Set Status='Returned' where AccessionNo=@d1 and ID=@d2"
            cmd = New SqlCommand(cb2)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
            cmd.Parameters.AddWithValue("@d2", txtIssueID_Staff.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            If txtReservationID.Text <> "" Then
                con = New SqlConnection(cs)
                con.Open()
                Dim cb3 As String = "Update BookReservation Set Status='Returned' where AccessionNo=@d1 and ID=@d2"
                cmd = New SqlCommand(cb3)
                cmd.Parameters.AddWithValue("@d1", txtAccessionNo1.Text)
                cmd.Parameters.AddWithValue("@d2", txtReservationID.Text)
                cmd.Connection = con
                cmd.ExecuteReader()
                con.Close()
            End If
            If Val(txtFine1.Text) > 0 Then
                LedgerSave(dtpReturnDate1.Value.Date, "Cash Account", txtID1.Text, "Book Fine(Staff)", 0, Val(txtFine1.Text), txtStaffID.Text)
            End If
            Dim st As String = "returned the Book of staff '" & txtStaffName.Text & "' having ID='" & txtID1.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully returned", "Book", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Return_Staff where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID1.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LedgerDelete(txtID1.Text, "Book Fine(Staff)")
                Dim st As String = "deleted the returned book record for staff '" & txtStaffName.Text & "' having Issue ID='" & txtID1.Text & "'"
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

            Dim cb As String = "Update Return_Staff set IssueID=@d2,ReturnDate=@d3,Fine=@d4,Remarks=@d5 where ID=" & txtID1.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtIssueID_Staff.Text)
            cmd.Parameters.AddWithValue("@d3", dtpReturnDate1.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtFine1.Text)
            cmd.Parameters.AddWithValue("@d5", txtRemarks1.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            If Val(txtFine1.Text) > 0 Then
                LedgerUpdate(dtpReturnDate1.Value.Date, "Cash Account", 0, Val(txtFine1.Text), txtStaffID.Text, txtID1.Text, "Book Fine(Staff)")
            End If
            Dim st As String = "updated the returned book record of staff '" & txtStaffName.Text & "' having ID='" & txtID1.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate1.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnGetData1_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData1.Click
        frmBookReturnRecord_Staff.lblSet.Text = "Book Return"
        frmBookReturnRecord_Staff.Reset()
        frmBookReturnRecord_Staff.ShowDialog()
    End Sub

    Private Sub dtpReturnDate_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpReturnDate.ValueChanged
        If dtpReturnDate.Value.Date < dtpDueDate.Value.Date Then
            txtFine.Text = 0
        Else
            Dim st As String = dtpReturnDate.Value.Date.Subtract(dtpDueDate.Value.Date).ToString()
            txtFine.Text = Val(st) * Val(txtFinePerDayStudent.Text)
        End If
    End Sub

    Private Sub dtpReturnDate1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtpReturnDate1.ValueChanged
        If dtpReturnDate1.Value.Date < dtpDueDate1.Value.Date Then
            txtFine1.Text = 0
        Else
            Dim st As String = dtpReturnDate1.Value.Date.Subtract(dtpDueDate1.Value.Date).ToString()
            txtFine1.Text = Val(st) * Val(txtFinePerDayStaff.Text)
        End If
    End Sub
End Class
