Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmStaffPaymentRecord

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(StaffPayment.ID) as [ID],(PaymentID) as [Payment ID],Convert(DateTime,DateFrom,103) as [Date From],Convert(DateTime,dateto,103) as [Date To],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(designation) as [Designation],RTRIM(StaffPayment.salary) as [Salary],RTRIM(presentdays) as [Prsesent Days],RTRIM(advance) as [Advance],RTRIM(deduction) as [Deduction],Convert(DateTime,paymentdate,131) as [Payment Date],RTRIM(modeofpayment) as [Payment Mode],RTRIM(paymentmodedetails) as [Payment Mode Details],RTRIM(netpay) as [Net Pay] from Staffpayment,Staff where Staff.St_ID=StaffPayment.StaffID order by paymentdate", con)
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

    Private Sub dgw_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Payment Entry" Then
                    Me.Hide()
                    frmStaffPayment.txtID.Text = dr.Cells(0).Value.ToString()
                    frmStaffPayment.PaymentID.Text = dr.Cells(1).Value.ToString()
                    frmStaffPayment.DateFrom.Text = dr.Cells(2).Value.ToString()
                    frmStaffPayment.DateTo.Text = dr.Cells(3).Value.ToString()
                    frmStaffPayment.txtStID.Text = dr.Cells(4).Value.ToString()
                    frmStaffPayment.StaffID.Text = dr.Cells(5).Value.ToString()
                    frmStaffPayment.StaffName.Text = dr.Cells(6).Value.ToString()
                    frmStaffPayment.Designation.Text = dr.Cells(7).Value.ToString()
                    con = New SqlConnection(cs)
                    con.Open()
                    Dim cp1 As String = "select Salary from Staff where St_id=" & dr.Cells(4).Value & ""
                    cmd = New SqlCommand(cp1)
                    cmd.Connection = con
                    rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        frmStaffPayment.txtSalary.Text = rdr.GetValue(0)
                    End If
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    If con.State = ConnectionState.Open Then
                        con.Close()
                    End If
                    frmStaffPayment.Salary.Text = dr.Cells(8).Value.ToString()
                    frmStaffPayment.PresentDays.Text = dr.Cells(9).Value.ToString()
                    frmStaffPayment.Advance.Text = dr.Cells(10).Value.ToString()
                    frmStaffPayment.Deduction.Text = dr.Cells(11).Value.ToString()
                    frmStaffPayment.PaymentDate.Text = dr.Cells(12).Value.ToString()
                    frmStaffPayment.paymentmode.Text = dr.Cells(13).Value.ToString()
                    frmStaffPayment.PaymentModeDetails.Text = dr.Cells(14).Value.ToString()
                    frmStaffPayment.NetPay.Text = dr.Cells(15).Value.ToString()
                    frmStaffPayment.btnSave.Enabled = False
                    frmStaffPayment.btnDelete.Enabled = True
                    frmStaffPayment.btnUpdate.Enabled = True
                    frmStaffPayment.btnPrint.Enabled = True
                    frmStaffPayment.DateFrom.Enabled = False
                    frmStaffPayment.DateTo.Enabled = False
                    frmStaffPayment.PaymentDate.Enabled = False
                    frmStaffPayment.Deduction.ReadOnly = True
                    frmStaffPayment.dgw.Enabled = False
                End If
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

    Private Sub txtStaffName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("select RTRIM(StaffPayment.ID) as [ID],(PaymentID) as [Payment ID],Convert(DateTime,DateFrom,103) as [Date From],Convert(DateTime,dateto,103) as [Date To],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(designation) as [Designation],RTRIM(StaffPayment.salary) as [Salary],RTRIM(presentdays) as [Prsesent Days],RTRIM(advance) as [Advance],RTRIM(deduction) as [Deduction],Convert(DateTime,paymentdate,131) as [Payment Date],RTRIM(modeofpayment) as [Payment Mode],RTRIM(paymentmodedetails) as [Payment Mode Details],RTRIM(netpay) as [Net Pay] from Staffpayment,Staff where Staff.St_ID=StaffPayment.StaffID and StaffName like '%" & txtStaffName.Text & "%' order by paymentdate", con)
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
            cmd = New SqlCommand("select RTRIM(StaffPayment.ID) as [ID],(PaymentID) as [Payment ID],Convert(DateTime,DateFrom,103) as [Date From],Convert(DateTime,dateto,103) as [Date To],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID],RTRIM(StaffName) as [Staff Name],RTRIM(designation) as [Designation],RTRIM(StaffPayment.salary) as [Salary],RTRIM(presentdays) as [Prsesent Days],RTRIM(advance) as [Advance],RTRIM(deduction) as [Deduction],Convert(DateTime,paymentdate,131) as [Payment Date],RTRIM(modeofpayment) as [Payment Mode],RTRIM(paymentmodedetails) as [Payment Mode Details],RTRIM(netpay) as [Net Pay] from Staffpayment,Staff where Staff.St_ID=StaffPayment.StaffID and PaymentDate Between @d1 and @d2 order by paymentdate", con)
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
