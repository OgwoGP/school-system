Imports System.Data.SqlClient
Public Class frmBookReservation
    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM BookReservation")
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
    Sub Reset()
        txtAccessionNo.Text = ""
        txtAuthor.Text = ""
        txtBookTitle.Text = ""
        txtID.Text = ""
        txtJointAuthor.Text = ""
        txtRemarks.Text = ""
        txtS_ID.Text = ""
        txtStaffID.Text = ""
        txtStaffName.Text = ""
        txtStatus.Text = ""
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        btnCancelReservation.Enabled = False
        Button1.Enabled = True
        Button2.Enabled = True
        auto()
    End Sub
    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmBookReservationRecord.lblSet.Text = "Book Reservation Entry"
        frmBookReservationRecord.Reset()
        frmBookReservationRecord.ShowDialog()
    End Sub

    Private Sub btnCancelReservation_Click(sender As System.Object, e As System.EventArgs) Handles btnCancelReservation.Click
        Try
            If (txtStatus.Text = "Cancelled") Then
                MessageBox.Show("Reservation is already cancelled", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If (txtStatus.Text = "Issued") Then
                MessageBox.Show("Reserved book is already issued " & vbCrLf & "so it can not be cancelled", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            If MessageBox.Show("Are you sure want to cancel this reservation?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                con = New SqlConnection(cs)
                con.Open()
                Dim cb As String = "update BookReservation set Status='Cancelled' where ID=" & txtID.Text & ""
                cmd = New SqlCommand(cb)
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
                Dim st As String = "cancelled the reservation having Reservation ID'" & txtID.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully cancelled", "Book Reservation", MessageBoxButtons.OK, MessageBoxIcon.Information)
                btnCancelReservation.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmBookEntryRecord1.lblSet.Text = "Book Reservation Entry"
        frmBookEntryRecord1.Reset()
        frmBookEntryRecord1.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        frmActiveStaffRecord.lblSet.Text = "Book Reservation Entry"
        frmActiveStaffRecord.Reset()
        frmActiveStaffRecord.ShowDialog()
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
        If Len(Trim(txtStaffID.Text)) = 0 Then
            MessageBox.Show("Please retrieve staff id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtStaffID.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select Status from BookReservation where AccessionNo=@d1 and Status='Reserved'"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Book is already reserved", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim ct1 As String = "select Status from BookReservation where AccessionNo=@d1 and Status='Issued'"
            cmd = New SqlCommand(ct1)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("Book is already issued", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into BookReservation(ID,AccessionNo,StaffID,R_Date, Status, Remarks) VALUES (@d1,@d2,@d3,@d4,'Reserved',@d5)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtAccessionNo.Text)
            cmd.Parameters.AddWithValue("@d3", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d4", dtpReservationDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "Update Book Set Status='Reserved' where AccessionNo=@d1"
            cmd = New SqlCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "added the new Book reservation for staff '" & txtStaffName.Text & "' having Reservation ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully reserved", "Book", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            Dim cq As String = "delete from BookReservation where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the book reservation record of staff '" & txtStaffName.Text & "' having Reservation ID='" & txtID.Text & "'")
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

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtAccessionNo.Text)) = 0 Then
            MessageBox.Show("Please retrieve accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo.Focus()
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
            Dim cb As String = "Update BookReservation set AccessionNo=@d2,StaffID=@d3,R_Date=@d4,Remarks=@d5 where ID=" & txtID.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtAccessionNo.Text)
            cmd.Parameters.AddWithValue("@d3", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d4", dtpReservationDate.Value.Date)
            cmd.Parameters.AddWithValue("@d5", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "updated the Book reservation for staff '" & txtStaffName.Text & "' having Reservation ID='" & txtID.Text & "'"
            LogFunc(lblUser.Text, st)
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmBookReservation_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
