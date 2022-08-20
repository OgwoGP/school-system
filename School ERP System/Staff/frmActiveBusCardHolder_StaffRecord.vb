Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmActiveBusCardHolder_StaffRecord

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
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Bus Fee Payment" Then
                    Me.Hide()
                    frmBusFeePayment_Staff.Show()
                    frmBusFeePayment_Staff.txtBusHolderID.Text = dr.Cells(0).Value.ToString()
                    frmBusFeePayment_Staff.txtSt_ID.Text = dr.Cells(1).Value.ToString()
                    frmBusFeePayment_Staff.txtStaffID.Text = dr.Cells(2).Value.ToString()
                    frmBusFeePayment_Staff.txtStaffName.Text = dr.Cells(3).Value.ToString()
                    frmBusFeePayment_Staff.txtLocation.Text = dr.Cells(6).Value.ToString()
                    frmBusFeePayment_Staff.fillInstallment()
                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = con.CreateCommand()
                    cmd.CommandText = "SELECT RTRIM(Designation),RTRIM(MobileNo) FROM Staff where ST_ID=@d1"
                    cmd.Parameters.AddWithValue("@d1", dr.Cells(0).Value)
                    rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        frmBusFeePayment_Staff.txtDesignation.Text = rdr.GetValue(0)
                        frmBusFeePayment_Staff.txtContactNo.Text = rdr.GetValue(1)
                    End If
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                    If con.State = ConnectionState.Open Then
                        con.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(BusCardHolder_Staff.BCH_ID) as [ID],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID], RTRIM(StaffName) as [StaffName],RTRIM(SchoolName) as [School Name],RTRIM(BusInfo.BusNo) as [Bus No.],RTRIM(LocationName) as [Location Name], CONVERT(DateTime,JoiningDate,105) as [Joining Date],RTRIM(BusCardHolder_Staff.Status) as [Status] from Staff,BusCardHolder_Staff,Location,BusInfo,schoolInfo where Staff.St_ID=BusCardHolder_Staff.StaffID  and Location.LocationName=BusCardHolder_Staff.Location and Staff.SchoolID=SchoolInfo.S_ID and BusInfo.BusNo=BusCardHolder_Staff.BusNo and BusCardHolder_Staff.Status='Active' order by StaffName", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtStaffName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(BusCardHolder_Staff.BCH_ID) as [ID],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID], RTRIM(StaffName) as [StaffName],RTRIM(SchoolName) as [School Name],RTRIM(BusInfo.BusNo) as [Bus No.],RTRIM(LocationName) as [Location Name], CONVERT(DateTime,JoiningDate,105) as [Joining Date],RTRIM(BusCardHolder_Staff.Status) as [Status] from Staff,BusCardHolder_Staff,Location,BusInfo,schoolInfo where Staff.St_ID=BusCardHolder_Staff.StaffID  and Location.LocationName=BusCardHolder_Staff.Location and Staff.SchoolID=SchoolInfo.S_ID and BusInfo.BusNo=BusCardHolder_Staff.BusNo and BusCardHolder_Staff.Status='Active' and StaffName like '%" & txtStaffName.Text & "%' order by StaffName", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtLocation_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtLocation.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(BusCardHolder_Staff.BCH_ID) as [ID],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID], RTRIM(StaffName) as [StaffName],RTRIM(SchoolName) as [School Name],RTRIM(BusInfo.BusNo) as [Bus No.],RTRIM(LocationName) as [Location Name], CONVERT(DateTime,JoiningDate,105) as [Joining Date],RTRIM(BusCardHolder_Staff.Status) as [Status] from Staff,BusCardHolder_Staff,Location,BusInfo,schoolInfo where Staff.St_ID=BusCardHolder_Staff.StaffID  and Location.LocationName=BusCardHolder_Staff.Location and Staff.SchoolID=SchoolInfo.S_ID and BusInfo.BusNo=BusCardHolder_Staff.BusNo and BusCardHolder_Staff.Status='Active' and LocationName like '%" & txtLocation.Text & "%' order by StaffName", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(BusCardHolder_Staff.BCH_ID) as [ID],RTRIM(Staff.St_ID) as [SID],RTRIM(Staff.StaffID) as [Staff ID], RTRIM(StaffName) as [StaffName],RTRIM(SchoolName) as [School Name],RTRIM(BusInfo.BusNo) as [Bus No.],RTRIM(LocationName) as [Location Name], CONVERT(DateTime,JoiningDate,105) as [Joining Date],RTRIM(BusCardHolder_Staff.Status) as [Status] from Staff,BusCardHolder_Staff,Location,BusInfo,schoolInfo where Staff.St_ID=BusCardHolder_Staff.StaffID  and Location.LocationName=BusCardHolder_Staff.Location and Staff.SchoolID=SchoolInfo.S_ID and BusInfo.BusNo=BusCardHolder_Staff.BusNo and BusCardHolder_Staff.Status='Active' and JoiningDate between @d1 and @d2 order by StaffName", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "JoiningDate").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "JoiningDate").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Staff")
            dgw.DataSource = ds.Tables("Staff").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        txtLocation.Text = ""
        txtStaffName.Text = ""
        GetData()
    End Sub

    Private Sub frmBusCardHolder_StaffRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub

    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub
End Class
