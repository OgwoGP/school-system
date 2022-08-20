Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmStaffRecord_ExcelSheet
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(St_ID) as [ID], RTRIM(StaffID) as [Staff ID], RTRIM(StaffName) as [Staff Name], Convert(DateTime,DateOfJoining,103) as  [Joining Date], RTRIM(Gender) as [Gender], RTRIM(FatherName) as [Father's Name], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(PermanentAddress) as [Permanent Address], RTRIM(Designation) as [Designation], RTRIM(Qualifications) as [Qualifications], Convert(DateTime,DOB,103) as [DOB], RTRIM(PhoneNo) as [Phone No.], RTRIM(MobileNo) as [Mobile No.], RTRIM(Staff.Email) as [Email ID],RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(ClassType) as [Class Type],RTRIM(Salary) as [Basic Salary],RTRIM(AccountName) as [Account Name],RTRIM(AccountNumber) as [Account No.],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCcode) as [IFSC Code] ,RTRIM(Status) as [Status] from Staff,ClassType,SchoolInfo where Staff.ClassType=ClassType.Type and Staff.SchoolID=SchoolInfo.S_ID order by StaffName", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub


    Sub Reset()
        txtStaffName.Text = ""
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        btnUpdate.Enabled = False
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub dgw_RowPostPaint(sender As Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub txtStaffName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(St_ID) as [ID], RTRIM(StaffID) as [Staff ID], RTRIM(StaffName) as [Staff Name], Convert(DateTime,DateOfJoining,103) as  [Joining Date], RTRIM(Gender) as [Gender], RTRIM(FatherName) as [Father's Name], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(PermanentAddress) as [Permanent Address], RTRIM(Designation) as [Designation], RTRIM(Qualifications) as [Qualifications], Convert(DateTime,DOB,103) as [DOB], RTRIM(PhoneNo) as [Phone No.], RTRIM(MobileNo) as [Mobile No.], RTRIM(Staff.Email) as [Email ID],RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(ClassType) as [Class Type],RTRIM(Salary) as [Basic Salary],RTRIM(AccountName) as [Account Name],RTRIM(AccountNumber) as [Account No.],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCcode) as [IFSC Code] ,RTRIM(Status) as [Status] from Staff,ClassType,SchoolInfo where Staff.ClassType=ClassType.Type and Staff.SchoolID=SchoolInfo.S_ID where Staffname like '%" & txtStaffName.Text & "%' order by StaffName", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtpDateTo_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles dtpDateTo.Validating
        If (dtpDateFrom.Value.Date) > (dtpDateTo.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtpDateTo.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(St_ID) as [ID], RTRIM(StaffID) as [Staff ID], RTRIM(StaffName) as [Staff Name], Convert(DateTime,DateOfJoining,103) as  [Joining Date], RTRIM(Gender) as [Gender], RTRIM(FatherName) as [Father's Name], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(PermanentAddress) as [Permanent Address], RTRIM(Designation) as [Designation], RTRIM(Qualifications) as [Qualifications], Convert(DateTime,DOB,103) as [DOB], RTRIM(PhoneNo) as [Phone No.], RTRIM(MobileNo) as [Mobile No.], RTRIM(Staff.Email) as [Email ID],RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(ClassType) as [Class Type],RTRIM(Salary) as [Basic Salary],RTRIM(AccountName) as [Account Name],RTRIM(AccountNumber) as [Account No.],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCcode) as [IFSC Code] ,RTRIM(Status) as [Status] from Staff,ClassType,SchoolInfo where Staff.ClassType=ClassType.Type and Staff.SchoolID=SchoolInfo.S_ID where DateOfJoining between @d1 and @d2 order by StaffName", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub btnImportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnImportExcel.Click
        Try
            Dim OpenFileDialog As New OpenFileDialog
            OpenFileDialog.Filter = "Excel Files | *.xlsx; *.xls;| All Files (*.*)| *.*"
            If OpenFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK AndAlso OpenFileDialog.FileName <> "" Then
                Cursor = Cursors.WaitCursor
                Timer1.Enabled = True
                Dim Pathname As String = OpenFileDialog.FileName
                Dim MyConnection As System.Data.OleDb.OleDbConnection
                Dim DtSet As System.Data.DataSet
                Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Pathname + ";Extended Properties=Excel 8.0;")
                MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
                MyConnection.Open()
                DtSet = New System.Data.DataSet
                MyCommand.Fill(DtSet)
                dgw.Visible = True
                dgw.DataSource = DtSet.Tables(0)
                btnUpdate.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Try
            If dgw.RowCount = Nothing Then
                MessageBox.Show("Sorry nothing to save.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    SqlConnection.ClearAllPools()
                    con = New SqlConnection(cs)
                    con.Open()
                    Dim ct As String = "select ST_ID from Staff Where ST_ID=@d1"
                    cmd = New SqlCommand(ct)
                    cmd.Parameters.AddWithValue("@d1", Val(row.Cells(0).Value))
                    cmd.Connection = con
                    rdr = cmd.ExecuteReader()
                    If Not rdr.Read() Then
                        Cursor = Cursors.WaitCursor
                        Timer1.Enabled = True
                        SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb As String = "Insert Into Staff(St_ID,StaffID,StaffName,DateOfJoining,Gender,FatherName,TemporaryAddress,PermanentAddress,Designation,Qualifications,DOB,PhoneNo,MobileNo,Email,SchoolID,ClassType,Salary,AccountName,AccountNumber,Bank,Branch,IFSCcode ,Status,Photo) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17,@d18,@d19,@d20,@d21,@d22,@d23,@d24)"
                        cmd = New SqlCommand(cb)
                        cmd.Connection = con
                        cmd.Parameters.AddWithValue("@d1", Val(row.Cells(0).Value))
                        cmd.Parameters.AddWithValue("@d2", row.Cells(1).Value)
                        cmd.Parameters.AddWithValue("@d3", row.Cells(2).Value)
                        cmd.Parameters.AddWithValue("@d4", row.Cells(3).Value)
                        cmd.Parameters.AddWithValue("@d5", row.Cells(4).Value)
                        cmd.Parameters.AddWithValue("@d6", row.Cells(5).Value)
                        cmd.Parameters.AddWithValue("@d7", row.Cells(6).Value)
                        cmd.Parameters.AddWithValue("@d8", row.Cells(7).Value)
                        cmd.Parameters.AddWithValue("@d9", row.Cells(8).Value)
                        cmd.Parameters.AddWithValue("@d10", row.Cells(9).Value)
                        cmd.Parameters.AddWithValue("@d11", row.Cells(10).Value)
                        cmd.Parameters.AddWithValue("@d12", row.Cells(11).Value.ToString())
                        cmd.Parameters.AddWithValue("@d13", row.Cells(12).Value.ToString())
                        cmd.Parameters.AddWithValue("@d14", row.Cells(13).Value)
                        cmd.Parameters.AddWithValue("@d15", Val(row.Cells(14).Value))
                        cmd.Parameters.AddWithValue("@d16", row.Cells(16).Value)
                        cmd.Parameters.AddWithValue("@d17", Val(row.Cells(17).Value))
                        cmd.Parameters.AddWithValue("@d18", row.Cells(18).Value.ToString())
                        cmd.Parameters.AddWithValue("@d19", row.Cells(19).Value.ToString())
                        cmd.Parameters.AddWithValue("@d20", row.Cells(20).Value.ToString())
                        cmd.Parameters.AddWithValue("@d21", row.Cells(21).Value.ToString())
                        cmd.Parameters.AddWithValue("@d22", row.Cells(22).Value.ToString())
                        cmd.Parameters.AddWithValue("@d23", row.Cells(23).Value.ToString())
                        Dim ms As New MemoryStream()
                        Dim bmpImage As New Bitmap(My.Resources.photo)
                        bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                        Dim data As Byte() = ms.GetBuffer()
                        Dim p As New SqlParameter("@d24", SqlDbType.Image)
                        p.Value = data
                        cmd.Parameters.Add(p)
                        cmd.ExecuteNonQuery()
                        con.Close()
                        SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb3 As String = "insert into Discount_Staff(StaffID,Discount) VALUES (@d1,0.00)"
                        cmd = New SqlCommand(cb3)
                        cmd.Parameters.AddWithValue("@d1", Val(row.Cells(0).Value))
                        cmd.Connection = con
                        cmd.ExecuteReader()
                        con.Close()
                    End If
                End If
            Next
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
        Catch ex As SqlException
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        Try
            If dgw.RowCount = Nothing Then
                MessageBox.Show("Sorry nothing to update" & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                        Cursor = Cursors.WaitCursor
                    Timer1.Enabled = True
                    SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb As String = "Update Staff Set StaffID=@d2,StaffName=@d3,DateOfJoining=@d4,Gender=@d5,FatherName=@d6,TemporaryAddress=@d7,PermanentAddress=@d8,Designation=@d9,Qualifications=@d10,DOB=@d11,PhoneNo=@d12,MobileNo=@d13,Email=@d14,SchoolID=@d15,ClassType=@d16,Salary=@d17,AccountName=@d18,AccountNumber=@d19,Bank=@d20,Branch=@d21,IFSCcode=@d22,Status=@d23 where ST_ID=@d1"
                        cmd = New SqlCommand(cb)
                        cmd.Connection = con
                        cmd.Parameters.AddWithValue("@d1", Val(row.Cells(0).Value))
                        cmd.Parameters.AddWithValue("@d2", row.Cells(1).Value)
                        cmd.Parameters.AddWithValue("@d3", row.Cells(2).Value)
                        cmd.Parameters.AddWithValue("@d4", row.Cells(3).Value)
                        cmd.Parameters.AddWithValue("@d5", row.Cells(4).Value)
                        cmd.Parameters.AddWithValue("@d6", row.Cells(5).Value)
                        cmd.Parameters.AddWithValue("@d7", row.Cells(6).Value)
                        cmd.Parameters.AddWithValue("@d8", row.Cells(7).Value)
                        cmd.Parameters.AddWithValue("@d9", row.Cells(8).Value)
                        cmd.Parameters.AddWithValue("@d10", row.Cells(9).Value)
                        cmd.Parameters.AddWithValue("@d11", row.Cells(10).Value)
                        cmd.Parameters.AddWithValue("@d12", row.Cells(11).Value.ToString())
                        cmd.Parameters.AddWithValue("@d13", row.Cells(12).Value.ToString())
                        cmd.Parameters.AddWithValue("@d14", row.Cells(13).Value)
                        cmd.Parameters.AddWithValue("@d15", Val(row.Cells(14).Value))
                        cmd.Parameters.AddWithValue("@d16", row.Cells(16).Value)
                        cmd.Parameters.AddWithValue("@d17", Val(row.Cells(17).Value))
                    cmd.Parameters.AddWithValue("@d18", row.Cells(18).Value.ToString())
                    cmd.Parameters.AddWithValue("@d19", row.Cells(19).Value.ToString())
                    cmd.Parameters.AddWithValue("@d20", row.Cells(20).Value.ToString())
                    cmd.Parameters.AddWithValue("@d21", row.Cells(21).Value.ToString())
                    cmd.Parameters.AddWithValue("@d22", row.Cells(22).Value.ToString())
                    cmd.Parameters.AddWithValue("@d23", row.Cells(23).Value.ToString())
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
            Next
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
        Catch ex As SqlException
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
