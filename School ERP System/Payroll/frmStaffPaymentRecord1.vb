Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmStaffPaymentRecord1

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(Designation) as [Designation],sum(StaffPayment.salary) as [Basic Salary],sum(Deduction) as [Deduction],sum(netpay) as [Net Pay] from StaffPayment,Staff where Staff.St_ID=StaffPayment.StaffID group by Staff.StaffID,Staffname,designation order by Staffname", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "StaffPayment")
            myDA.Fill(myDataSet, "Staff")
            dgw.DataSource = myDataSet.Tables("StaffPayment").DefaultView
            dgw.DataSource = myDataSet.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub
    Sub Reset()
        txtStaffName.Text = ""
        DateFrom.Value = Today
        DateTo.Value = Now
        GetData()
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnExportExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
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

    Private Sub txtStaffname_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(Designation) as [Designation],sum(StaffPayment.salary) as [Basic Salary],sum(Deduction) as [Deduction],sum(netpay) as [Net Pay] from StaffPayment,Staff where Staff.St_ID=StaffPayment.StaffID and Staffname like '%" & txtStaffName.Text & "%' group by Staff.StaffID,Staffname,designation order by Staffname", con)
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "StaffPayment")
            myDA.Fill(myDataSet, "Staff")
            dgw.DataSource = myDataSet.Tables("StaffPayment").DefaultView
            dgw.DataSource = myDataSet.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(Designation) as [Designation],sum(StaffPayment.salary) as [Basic Salary],sum(Deduction) as [Deduction],sum(netpay) as [Net Pay] from StaffPayment,Staff where Staff.St_ID=StaffPayment.StaffID and PaymentDate between @d1 and @d2 group by Staff.StaffID,Staffname,designation order by Staffname", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "DateIN").Value = DateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "DateIN").Value = DateTo.Value
            Dim myDA As SqlDataAdapter = New SqlDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "StaffPayment")
            myDA.Fill(myDataSet, "Staff")
            dgw.DataSource = myDataSet.Tables("StaffPayment").DefaultView
            dgw.DataSource = myDataSet.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
