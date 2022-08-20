Imports System.Data.SqlClient
Public Class frmSupplier
    Dim s1, s2, s3 As String
    Sub Reset()
        txtLastName.Text = ""
        txtAddress.Text = ""
        txtRemarks.Text = ""
        txtFirstName.Text = ""
        txtSupplierID.Text = ""
        txtContactNo.Text = ""
        txtEmailID.Text = ""
        cbBooks.Checked = False
        cbJM.Checked = False
        cbNewsPaper.Checked = False
        txtLastName.Focus()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        auto()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Function GenerateID() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 ID FROM Supplier ORDER BY ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("ID")
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
    Sub auto()
        Try
            txtID.Text = GenerateID()
            txtSupplierID.Text = "SUP-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtLastName.Text)) = 0 Then
            MessageBox.Show("Please enter last name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtLastName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFirstName.Text)) = 0 Then
            MessageBox.Show("Please enter first name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFirstName.Focus()
            Exit Sub
        End If
        If ((cbBooks.Checked = False) And (cbJM.Checked = False) And (cbNewsPaper.Checked = False)) Then
            MessageBox.Show("Please select Supplier Type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Len(Trim(txtAddress.Text)) = 0 Then
            MessageBox.Show("Please Enter Address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAddress.Focus()
            Exit Sub
        End If
        If Len(Trim(txtContactNo.Text)) = 0 Then
            MessageBox.Show("Please Enter Contact No.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContactNo.Focus()
            Exit Sub
        End If

        Try
            If (cbBooks.Checked = True) Then
                s1 = "Yes"
            Else
                s1 = "No"
            End If
            If (cbJM.Checked = True) Then
                s2 = "Yes"
            Else
                s2 = "No"
            End If
            If (cbNewsPaper.Checked = True) Then
                s3 = "Yes"
            Else
                s3 = "No"
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Supplier(ID,SupplierID,LastName,FirstName,S_Books, S_NewsPaper, S_Magazines, Address, ContactNo, EmailID,Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtSupplierID.Text)
            cmd.Parameters.AddWithValue("@d3", txtLastName.Text)
            cmd.Parameters.AddWithValue("@d4", txtFirstName.Text)
            cmd.Parameters.AddWithValue("@d5", s1)
            cmd.Parameters.AddWithValue("@d6", s3)
            cmd.Parameters.AddWithValue("@d7", s2)
            cmd.Parameters.AddWithValue("@d8", txtAddress.Text)
            cmd.Parameters.AddWithValue("@d9", txtContactNo.Text)
            cmd.Parameters.AddWithValue("@d10", txtEmailID.Text)
            cmd.Parameters.AddWithValue("@d11", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            LogFunc(lblUser.Text, "added the new supplier having supplier id '" & txtSupplierID.Text & "'")
            MessageBox.Show("Successfully saved", "Supplier Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Getdata()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
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
            Dim cl As String = "select Supplier.ID from Supplier,Book where Supplier.ID=Book.SupplierID and Supplier.ID=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Book Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cl1 As String = "select Supplier.ID from JournalandMagazines,Supplier where Supplier.ID=JournalandMagazines.SupplierID and Supplier.ID=@d1"
            cmd = New SqlCommand(cl1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Journals and Magazines Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Supplier where ID =" & txtID.Text & ""
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the supplier record having supplier id '" & txtSupplierID.Text & "'")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Getdata()
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtLastName.Text)) = 0 Then
            MessageBox.Show("Please enter last name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtLastName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtFirstName.Text)) = 0 Then
            MessageBox.Show("Please enter first name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFirstName.Focus()
            Exit Sub
        End If
   
        If ((cbBooks.Checked = False) And (cbJM.Checked = False) And (cbNewsPaper.Checked = False)) Then
            MessageBox.Show("Please select Supplier Type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Len(Trim(txtAddress.Text)) = 0 Then
            MessageBox.Show("Please Enter Address", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAddress.Focus()
            Exit Sub
        End If
        If Len(Trim(txtContactNo.Text)) = 0 Then
            MessageBox.Show("Please Enter Contact No.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContactNo.Focus()
            Exit Sub
        End If

        Try
            If (cbBooks.Checked = True) Then
                s1 = "Yes"
            Else
                s1 = "No"
            End If
            If (cbJM.Checked = True) Then
                s2 = "Yes"
            Else
                s2 = "No"
            End If
            If (cbNewsPaper.Checked = True) Then
                s3 = "Yes"
            Else
                s3 = "No"
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Supplier set SupplierID=@d2,LastName=@d3,FirstName=@d4,S_Books=@d5, S_NewsPaper=@d6, S_Magazines=@d7, Address=@d8, ContactNo=@d9, EmailID=@d10,Remarks=@d11 where ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtSupplierID.Text)
            cmd.Parameters.AddWithValue("@d3", txtLastName.Text)
            cmd.Parameters.AddWithValue("@d4", txtFirstName.Text)
            cmd.Parameters.AddWithValue("@d5", s1)
            cmd.Parameters.AddWithValue("@d6", s3)
            cmd.Parameters.AddWithValue("@d7", s2)
            cmd.Parameters.AddWithValue("@d8", txtAddress.Text)
            cmd.Parameters.AddWithValue("@d9", txtContactNo.Text)
            cmd.Parameters.AddWithValue("@d10", txtEmailID.Text)
            cmd.Parameters.AddWithValue("@d11", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            LogFunc(lblUser.Text, "updated the supplier record having supplier id '" & txtSupplierID.Text & "'")
            MessageBox.Show("Successfully updated", "Supplier Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            Getdata()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(ID),RTRIM(SupplierID),RTRIM(LastName),RTRIM(FirstName),RTRIM(S_Books),RTRIM(S_NewsPaper), RTRIM(S_Magazines), RTRIM(Address), RTRIM(ContactNo), RTRIM(EmailID),RTRIM(Remarks) from supplier order by supplierid", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub


    Private Sub dgw_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub


    Private Sub dgw_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            Dim dr As DataGridViewRow = dgw.SelectedRows(0)
            txtID.Text = dr.Cells(0).Value.ToString()
            txtSupplierID.Text = dr.Cells(1).Value.ToString()
            txtLastName.Text = dr.Cells(2).Value.ToString()
            txtFirstName.Text = dr.Cells(3).Value.ToString()
            If dr.Cells(4).Value = "Yes" Then
                cbBooks.Checked = True
            Else
                cbBooks.Checked = False
            End If
            If dr.Cells(5).Value = "Yes" Then
                cbNewsPaper.Checked = True
            Else
                cbNewsPaper.Checked = False
            End If
            If dr.Cells(6).Value = "Yes" Then
                cbJM.Checked = True
            Else
                cbJM.Checked = False
            End If
            txtAddress.Text = dr.Cells(7).Value.ToString()
            txtContactNo.Text = dr.Cells(8).Value.ToString()
            txtEmailID.Text = dr.Cells(9).Value.ToString()
            txtRemarks.Text = dr.Cells(10).Value.ToString()
            btnUpdate.Enabled = True
            btnDelete.Enabled = True
            btnSave.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmSupplier_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Getdata()
    End Sub
End Class
