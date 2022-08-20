Imports System.Data.SqlClient
Public Class frmSubCategory
    Sub fillCombo()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(Classname) FROM BookClass,Category where BookClass.ClassName=Category.Class", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbClass.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbClass.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Fill()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID from Category,BookClass where Classname=@d1 and CategoryName=@d2 and BookClass.ClassName=Category.Class"
            cmd.Parameters.AddWithValue("@d1", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d2", cmbCategory.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtCategoryID.Text = rdr.GetValue(0)
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
    Sub Reset()
        cmbClass.SelectedIndex = -1
        txtSearchBySubCategory.Text = ""
        txtSubCategory.Text = ""
        txtCategoryID.Text = ""
        txtSubCategory.Focus()
        btnSave.Enabled = True
        cmbCategory.Enabled = False
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        Getdata()
        auto()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtSubCategory.Text)) = 0 Then
            MessageBox.Show("Please enter sub category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSubCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbClass.Text)) = 0 Then
            MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbClass.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbCategory.Text)) = 0 Then
            MessageBox.Show("Please select category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCategory.Focus()
            Exit Sub
        End If
        Try

            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select CategoryName,ClassName,SubCategoryName from SubCategory,BookClass,Category where Category.Class=BookClass.Classname and Category.ID=SubCategory.CategoryID and CategoryName=@d1 and ClassName=@d2 and SubCategoryName=@d3"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbCategory.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d3", txtSubCategory.Text)
            rdr = cmd.ExecuteReader()

            If rdr.Read() Then
                MessageBox.Show("Record Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                txtSubCategory.Text = ""
                txtSubCategory.Focus()
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            Fill()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into SubCategory(CategoryID,SubCategoryName,ID) VALUES (" & txtCategoryID.Text & ",@d1," & txtID.Text & ")"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubCategory.Text)
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "added the new sub category '" & txtSubCategory.Text & "' having class '" & cmbClass.Text & "' and Category='" & cmbCategory.Text & "'")
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            Getdata()
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
    Private Sub DeleteRecord()

        Try

            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "select SubCategory.ID from SubCategory,Book where SubCategory.ID=Book.SubCategoryID and SubCategory.ID=@d1"
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
            Dim cq As String = "delete from SubCategory where ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the sub category '" & txtSubCategory.Text & "' having class '" & cmbClass.Text & "' and Category='" & cmbCategory.Text & "'")
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

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try
            If Len(Trim(txtSubCategory.Text)) = 0 Then
                MessageBox.Show("Please enter Category name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSubCategory.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbClass.Text)) = 0 Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass.Focus()
                Exit Sub
            End If
            Fill()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "update SubCategory set SubCategoryName=@d1,CategoryID=" & txtCategoryID.Text & " where ID=" & txtID.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtSubCategory.Text)
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "updated the sub category '" & txtSubCategory.Text & "' having class '" & cmbClass.Text & "' and Category='" & cmbCategory.Text & "'")
            MessageBox.Show("Successfully updated", "Sub Category Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            Getdata()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(SubCategory.ID),RTRIM(CategoryID),RTRIM(SubCategoryName), RTRIM(Class),RTRIM(Categoryname) from Category,BookClass,SubCategory where Bookclass.classname=category.Class and SubCategory.CategoryID=Category.ID order by SubCategoryName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4))
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
    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = ("SELECT MAX(ID) FROM SubCategory")
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
    Private Sub frmCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Getdata()
        fillCombo()
    End Sub

    Private Sub dgw_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                txtSubCategory.Text = dr.Cells(2).Value.ToString()
                txtID.Text = dr.Cells(0).Value.ToString()
                txtCategoryID.Text = dr.Cells(1).Value.ToString()
                cmbClass.Text = dr.Cells(3).Value.ToString()
                cmbCategory.Text = dr.Cells(4).Value.ToString()
                btnUpdate.Enabled = True
                btnDelete.Enabled = True
                btnSave.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    Private Sub cmbClass_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbClass.SelectedIndexChanged
        Try
            cmbCategory.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(CategoryName) FROM BookClass,Category where Category.Class=BookClass.ClassName and Classname=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbClass.Text)
            rdr = cmd.ExecuteReader()
            cmbCategory.Items.Clear()
            While rdr.Read
                cmbCategory.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSearchBySubCategory_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSearchBySubCategory.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(SubCategory.ID),RTRIM(CategoryID),RTRIM(SubCategoryName), RTRIM(Class),RTRIM(Categoryname) from Category,BookClass,SubCategory where Bookclass.classname=category.Class and SubCategory.CategoryID=Category.ID and SubCategoryName like '%" & txtSearchBySubCategory.Text & "%' order by SubCategoryName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
