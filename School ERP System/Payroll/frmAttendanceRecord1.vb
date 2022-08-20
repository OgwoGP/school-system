Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmAttendanceEntryRecord1

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Sub Reset()
        DateFrom.Value = Today
        DateTo.Value = Now
        lblWorkingDays.Visible = False
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

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            lblWorkingDays.Visible = True
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = "select RTRIM(Staff.StaffID),RTRIM(StaffName),RTRIM(Designation),Count(StaffAttendance.Status) from StaffAttendance,Staff where Staff.St_ID=StaffAttendance.StaffID and WorkingDate between @d1 and @d2 and StaffAttendance.Status='P' group by Staff.StaffID,StaffName,designation order by StaffName"
            cmd = New SqlCommand(sql, con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "DateIN").Value = DateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "DateIN").Value = DateTo.Value
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim sql1 As String = "select Count(StaffAttendance.Status) from StaffAttendance,Staff where Staff.St_ID=StaffAttendance.StaffID and WorkingDate between @d1 and @d2"
            cmd = New SqlCommand(sql1, con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "DateIN").Value = DateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "DateIN").Value = DateTo.Value
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If (rdr.Read() = True) Then
                lblWorkingDays.Text = rdr.GetValue(0)
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
