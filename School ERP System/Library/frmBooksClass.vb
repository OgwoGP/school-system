Imports System.Data.SqlClient
Public Class frmBooksClass

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Sub Reset()
        txtClass.Text = ""
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        txtClass.Focus()
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtClass.Text = "" Then
            MessageBox.Show("Please enter class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtClass.Focus()
            Return
        End If

        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select Classname from BookClass where Classname=@d1"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", txtClass.Text)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If rdr.Read() Then
                MessageBox.Show("Class Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                txtClass.Text = ""
                txtClass.Focus()
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If

            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into BookClass(ClassName) VALUES (@d1)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtClass.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "added the new book class '" & txtClass.Text & "'")
            MessageBox.Show("Successfully Saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Getdata()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If txtClass.Text = "" Then
            MessageBox.Show("Please enter class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtClass.Focus()
            Return
        End If

        Try

            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "Update BookClass set Classname=@d1 where Classname=@d2"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtClass.Text)
            cmd.Parameters.AddWithValue("@d2", txtClassName.Text)
            cmd.ExecuteReader()
            LogFunc(lblUser.Text, "updated the book class '" & txtClass.Text & "'")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            Getdata()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub DeleteRecord()

        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "select ClassName from BookClass,Category where BookClass.ClassName=Category.Class and ClassName=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtClassName.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Category Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from BookClass where Classname=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtClassName.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the class '" & txtClass.Text & "'")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Getdata()
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
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgw_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                txtClassName.Text = dr.Cells(0).Value.ToString()
                txtClass.Text = dr.Cells(0).Value.ToString()
                btnUpdate.Enabled = True
                btnDelete.Enabled = True
                btnSave.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(Classname) from BookClass order by Classname", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub frmClass_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Getdata()
    End Sub
End Class
