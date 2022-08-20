Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmAdvanceEntryRecord1

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = "SELECT RTRIM(Staff.StaffID),RTRIM(Staff.StaffName),RTRIM(Designation),sum(Amount)-sum(Deduction) FROM Staff Inner join AdvanceEntry on Staff.St_ID=AdvanceEntry.StaffID group by StaffName,Staff.StaffID,Designation having (sum(Amount)-sum(Deduction)) > 0 order by StaffName"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
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

    Private Sub txtStaffName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStaffName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim sql As String = "SELECT RTRIM(Staff.StaffID),RTRIM(Staff.StaffName),RTRIM(Designation),sum(Amount)-sum(Deduction) FROM Staff Inner join AdvanceEntry on Staff.St_ID=AdvanceEntry.StaffID and StaffName like '%" & txtStaffName.Text & "%' group by StaffName,Staff.StaffID,Designation having (sum(Amount)-sum(Deduction)) > 0 order by StaffName"
            cmd = New SqlCommand(sql, con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
