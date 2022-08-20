Imports System.Data.SqlClient

Public Class frmVoucherReport
    Sub fillPaymentID()
        Try
            Dim CN As New SqlConnection(cs)
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(VoucherNo) FROM Voucher", CN)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbVoucherNo.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbVoucherNo.Items.Add(drow(0).ToString())
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillPaymentID()
    End Sub
    Sub Reset()
        cmbVoucherNo.Text = ""
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        fillPaymentID()
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If cmbVoucherNo.Text = "" Then
                MessageBox.Show("Please select voucher no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbVoucherNo.Focus()
                Exit Sub
            End If
                Cursor = Cursors.WaitCursor
                Timer1.Enabled = True
                Dim rpt As New rptVoucher 'The report you created.
                Dim myConnection As SqlConnection
                Dim MyCommand, MyCommand1 As New SqlCommand()
                Dim myDA, myDA1 As New SqlDataAdapter()
                Dim myDS As New DataSet 'The DataSet you created.
                myConnection = New SqlConnection(cs)
                MyCommand.Connection = myConnection
                MyCommand1.Connection = myConnection
            MyCommand.CommandText = "SELECT Voucher.ID, Voucher.VoucherNo,Voucher.PaymentMode, Voucher.Date, Voucher.Name, Voucher.Details, Voucher.GrandTotal, Voucher_OtherDetails.VD_ID, Voucher_OtherDetails.VoucherID,Voucher_OtherDetails.Particulars, Voucher_OtherDetails.Amount, Voucher_OtherDetails.Note FROM Voucher INNER JOIN Voucher_OtherDetails ON Voucher.ID = Voucher_OtherDetails.VoucherID  where VoucherNo='" & cmbVoucherNo.Text & "'"
                MyCommand1.CommandText = "SELECT * from SchoolInfo"
                MyCommand.CommandType = CommandType.Text
                MyCommand1.CommandType = CommandType.Text
                myDA.SelectCommand = MyCommand
                myDA1.SelectCommand = MyCommand1
                myDA.Fill(myDS, "Voucher")
                myDA.Fill(myDS, "Voucher_OtherDetails")
                myDA1.Fill(myDS, "SchoolInfo")
                rpt.SetDataSource(myDS)
                frmReport.CrystalReportViewer1.ReportSource = rpt
                frmReport.ShowDialog()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub cmbVoucherNo_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbVoucherNo.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptVoucherByDate 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Voucher.ID, Voucher.VoucherNo, Voucher.Date, Voucher.Name, Voucher.Details, Voucher.GrandTotal, Voucher_OtherDetails.VD_ID, Voucher_OtherDetails.VoucherID,Voucher_OtherDetails.Particulars, Voucher_OtherDetails.Amount, Voucher_OtherDetails.Note FROM Voucher INNER JOIN Voucher_OtherDetails ON Voucher.ID = Voucher_OtherDetails.VoucherID where Date >=@d1 and Date < @d2 order by Date"
            MyCommand.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date.AddDays(1)
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Voucher")
            myDA.Fill(myDS, "Voucher_OtherDetails")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", dtpDateFrom.Value.Date)
            rpt.SetParameterValue("p2", dtpDateTo.Value.Date)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
