Imports System.Data.SqlClient

Imports System.IO

Public Class frmSupplierRecord

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(ID),RTRIM(SupplierID),RTRIM(LastName),RTRIM(FirstName),RTRIM(S_Books),RTRIM(S_NewsPaper), RTRIM(S_Magazines), RTRIM(Address), RTRIM(ContactNo), RTRIM(EmailID),RTRIM(Remarks) from supplier order by supplierid", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Getdata()
    End Sub

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dgw_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Book Entry" Then
                    Me.Hide()
                    frmBookEntry.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    frmBookEntry.txtID.Text = dr.Cells(0).Value.ToString()
                    frmBookEntry.txtSupplierID.Text = dr.Cells(1).Value.ToString()
                    frmBookEntry.txtLastName.Text = dr.Cells(2).Value.ToString()
                    frmBookEntry.txtFirstName.Text = dr.Cells(3).Value.ToString()
                    lblSet.Text = ""
                End If
                If lblSet.Text = "JandM Entry" Then
                    Me.Hide()
                    frmJournalsandMagazines.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    frmJournalsandMagazines.txtS_ID.Text = dr.Cells(0).Value.ToString()
                    frmJournalsandMagazines.txtSupplierID.Text = dr.Cells(1).Value.ToString()
                    frmJournalsandMagazines.txtLastName.Text = dr.Cells(2).Value.ToString()
                    frmJournalsandMagazines.txtFirstName.Text = dr.Cells(3).Value.ToString()
                    lblSet.Text = ""
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

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub
End Class
