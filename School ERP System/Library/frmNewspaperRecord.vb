Imports System.Data.SqlClient
Imports System.IO

Public Class frmNewspaperRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT ID,RTRIM(PaperName) as [Paper Name],N_Date as [Date],RTRIM(Status) as [Status],RTRIM(Remarks) as [Remarks] from NewsPaper order by PaperName, N_Date", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "NewsPaper")
            dgw.DataSource = ds.Tables("NewsPaper").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Sub fillcombo()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(PaperName) FROM NewsPaper", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbPaperName.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbPaperName.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub Reset()
        cmbPaperName.SelectedIndex = -1
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Newspaper Entry" Then
                    Me.Hide()
                    frmNewspaper.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    frmNewspaper.txtID.Text = dr.Cells(0).Value.ToString()
                    frmNewspaper.cmbNewsPaper.Text = dr.Cells(1).Value.ToString()
                    frmNewspaper.dtpDate.Text = dr.Cells(2).Value.ToString()
                    If (dr.Cells(3).Value.ToString = "P") Then
                        frmNewspaper.RadioButton1.Checked = True
                    End If
                    If (dr.Cells(3).Value.ToString = "A") Then
                        frmNewspaper.RadioButton2.Checked = True
                    End If
                    frmNewspaper.txtRemarks.Text = dr.Cells(4).Value.ToString()
                    frmNewspaper.btnUpdate.Enabled = True
                    frmNewspaper.btnDelete.Enabled = True
                    frmNewspaper.btnSave.Enabled = False
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

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            If Len(Trim(cmbPaperName.Text)) = 0 Then
                MessageBox.Show("Please Select Paper Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbPaperName.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT ID,RTRIM(PaperName) as [Paper Name],N_Date as [Date],RTRIM(Status) as [Status],RTRIM(Remarks) as [Remarks] from NewsPaper where N_Date between @d1 and @d2 and PaperName=@d3  order by N_Date", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Parameters.AddWithValue("@d3", cmbPaperName.Text)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "NewsPaper")
            dgw.DataSource = ds.Tables("NewsPaper").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If dgw.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        ExportExcel(dgw)
    End Sub

    Private Sub dtpDateTo_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles dtpDateTo.Validating
        If (dtpDateFrom.Value.Date) > (dtpDateTo.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtpDateTo.Focus()
        End If
    End Sub

    Private Sub frmNewspaperRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillcombo()
    End Sub
End Class
