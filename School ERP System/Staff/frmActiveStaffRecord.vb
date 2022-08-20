Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmActiveStaffRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(St_ID) as [ID], RTRIM(StaffID) as [Staff ID], RTRIM(StaffName) as [Staff Name], Convert(DateTime,DateOfJoining,103) as [Joining Date], RTRIM(Gender) as [Gender], RTRIM(FatherName) as [Father's Name], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(PermanentAddress) as [Permanent Address], RTRIM(Designation) as [Designation], RTRIM(Qualifications) as [Qualifications], Convert(DateTime,DOB,103) as [DOB], RTRIM(PhoneNo) as [Phone No.], RTRIM(MobileNo) as [Mobile No.], RTRIM(Staff.Email) as [Email ID],RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(ClassType) as [Class Type],RTRIM(Salary) as [Basic Salary],RTRIM(AccountName) as [Account Name],RTRIM(AccountNumber) as [Account No.],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCcode) as [IFSC Code], Photo,RTRIM(Status) as [Status] from Staff,ClassType,SchoolInfo where Staff.ClassType=ClassType.Type and Staff.SchoolID=SchoolInfo.S_ID and Staff.Status='Active' order by StaffName", con)
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

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Bus Holder Entry" Then
                    Me.Hide()
                    frmBusCardHolder_Staff.Show()
                    frmBusCardHolder_Staff.txtS_ID.Text = dr.Cells(0).Value.ToString()
                    frmBusCardHolder_Staff.txtStaffID.Text = dr.Cells(1).Value.ToString()
                    frmBusCardHolder_Staff.txtStaffName.Text = dr.Cells(2).Value.ToString()
                    frmBusCardHolder_Staff.txtSchoolName.Text = dr.Cells(15).Value.ToString()
                    lblSet.Text = ""
                End If
                If lblSet.Text = "Book Issue Entry" Then
                    Me.Hide()
                    frmBookIssue.Show()
                    frmBookIssue.txtS_ID.Text = dr.Cells(0).Value.ToString()
                    frmBookIssue.txtStaffID.Text = dr.Cells(1).Value.ToString()
                    frmBookIssue.txtStaffName.Text = dr.Cells(2).Value.ToString()
                    lblSet.Text = ""
                End If
                If lblSet.Text = "Book Reservation Entry" Then
                    Me.Hide()
                    frmBookReservation.Show()
                    frmBookReservation.txtS_ID.Text = dr.Cells(0).Value.ToString()
                    frmBookReservation.txtStaffID.Text = dr.Cells(1).Value.ToString()
                    frmBookReservation.txtStaffName.Text = dr.Cells(2).Value.ToString()
                    lblSet.Text = ""
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
            cmd = New SqlCommand("Select RTRIM(St_ID) as [ID], RTRIM(StaffID) as [Staff ID], RTRIM(StaffName) as [Staff Name], Convert(DateTime,DateOfJoining,103) as [Joining Date], RTRIM(Gender) as [Gender], RTRIM(FatherName) as [Father's Name], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(PermanentAddress) as [Permanent Address], RTRIM(Designation) as [Designation], RTRIM(Qualifications) as [Qualifications], Convert(DateTime,DOB,103) as [DOB], RTRIM(PhoneNo) as [Phone No.], RTRIM(MobileNo) as [Mobile No.], RTRIM(Staff.Email) as [Email ID],RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(ClassType) as [Class Type],RTRIM(Salary) as [Basic Salary],RTRIM(AccountName) as [Account Name],RTRIM(AccountNumber) as [Account No.],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCcode) as [IFSC Code], Photo,RTRIM(Status) as [Status] from Staff,ClassType,SchoolInfo where Staff.ClassType=ClassType.Type and Staff.SchoolID=SchoolInfo.S_ID and Staff.Status='Active' where Staffname like '%" & txtStaffName.Text & "%' order by StaffName", con)
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
            cmd = New SqlCommand("Select RTRIM(St_ID) as [ID], RTRIM(StaffID) as [Staff ID], RTRIM(StaffName) as [Staff Name], Convert(DateTime,DateOfJoining,103) as [Joining Date], RTRIM(Gender) as [Gender], RTRIM(FatherName) as [Father's Name], RTRIM(TemporaryAddress) as [Temporary Address], RTRIM(PermanentAddress) as [Permanent Address], RTRIM(Designation) as [Designation], RTRIM(Qualifications) as [Qualifications], Convert(DateTime,DOB,103) as [DOB], RTRIM(PhoneNo) as [Phone No.], RTRIM(MobileNo) as [Mobile No.], RTRIM(Staff.Email) as [Email ID],RTRIM(SchoolID) as [School ID],RTRIM(SchoolName) as [School Name],RTRIM(ClassType) as [Class Type],RTRIM(Salary) as [Basic Salary],RTRIM(AccountName) as [Account Name],RTRIM(AccountNumber) as [Account No.],RTRIM(Bank) as [Bank],RTRIM(Branch) as [Branch],RTRIM(IFSCcode) as [IFSC Code], Photo,RTRIM(Status) as [Status] from Staff,ClassType,SchoolInfo where Staff.ClassType=ClassType.Type and Staff.SchoolID=SchoolInfo.S_ID and Staff.Status='Active' where DateOfJoining between @d1 and @d2 order by StaffName", con)
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
End Class
