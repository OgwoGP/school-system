Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmStaffRecord1
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

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ExportExcel(dgw)
    End Sub
End Class
