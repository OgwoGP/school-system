Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmFullDueList_Staff

    Sub fillSession()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM Session_Master", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSession.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSession.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub Reset()
        cmbSession.SelectedIndex = -1
        cmbInstallment.SelectedIndex = -1
        cmbInstallment.Enabled = False
        dgw.Rows.Clear()
        fillSession()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmPartialDueList_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSession()
    End Sub

    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        Reset()
    End Sub

    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        ExportExcel(dgw)
    End Sub

    Private Sub TabControl1_Click(sender As System.Object, e As System.EventArgs) Handles TabControl1.Click
        Reset()
    End Sub

    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        Try
            If Len(Trim(cmbSession.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbInstallment.Text)) = 0 Then
                MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbInstallment.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(Staffname),RTRIM(Designation),RTRIM(BusCardHolder_Staff.Location),RTRIM(SchoolName),SUM(Charges) FROM Staff,BusCardHolder_Staff,SchoolInfo,Location,Installment_Bus,Session_Master where Staff.ST_ID=BusCardHolder_Staff.StaffID and Staff.SchoolID=SchoolInfo.S_ID and Location.LocationName=Installment_Bus.Location and BusCardHolder_Staff.Location=Location.LocationName and Session=@d1 and Installment_Bus.Installment=@d2 group by StaffName,Designation,BusCardHolder_Staff.Location,SchoolName Except SELECT RTRIM(Staffname),RTRIM(Designation),RTRIM(BusCardHolder_Staff.Location),RTRIM(SchoolName),SUM(Charges) FROM Staff,BusCardHolder_Staff,SchoolInfo,Location,Installment_Bus,BusFeePayment_Staff where Staff.ST_ID=BusCardHolder_Staff.StaffID and Staff.SchoolID=SchoolInfo.S_ID and Location.LocationName=Installment_Bus.Location and BusFeePayment_Staff.BusHolderID=BusCardHolder_Staff.BCH_ID and Installment_Bus.Installment=BusFeePayment_Staff.Installment and BusFeePayment_Staff.Session=@d1 and BusCardHolder_Staff.Location=Location.Locationname and BusFeePayment_Staff.Installment=@d2 group by StaffName,Designation,BusCardHolder_Staff.Location,SchoolName order by 1", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbInstallment.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub


    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        Try
            If Len(Trim(cmbSession.Text)) = 0 Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If Len(Trim(cmbInstallment.Text)) = 0 Then
                MessageBox.Show("Please select installment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbInstallment.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(Staffname),RTRIM(Designation),RTRIM(BusCardHolder_Staff.Location),RTRIM(SchoolName),SUM(Charges) FROM Staff,BusCardHolder_Staff,SchoolInfo,Location,Installment_Bus,Session_Master where Staff.ST_ID=BusCardHolder_Staff.StaffID and Staff.SchoolID=SchoolInfo.S_ID and Location.LocationName=Installment_Bus.Location and BusCardHolder_Staff.Location=Location.LocationName and Session=@d1 and Installment_Bus.Installment=@d2 group by StaffName,Designation,BusCardHolder_Staff.Location,SchoolName Except SELECT RTRIM(Staffname),RTRIM(Designation),RTRIM(BusCardHolder_Staff.Location),RTRIM(SchoolName),SUM(Charges) FROM Staff,BusCardHolder_Staff,SchoolInfo,Location,Installment_Bus,BusFeePayment_Staff where Staff.ST_ID=BusCardHolder_Staff.StaffID and Staff.SchoolID=SchoolInfo.S_ID and Location.LocationName=Installment_Bus.Location and BusFeePayment_Staff.BusHolderID=BusCardHolder_Staff.BCH_ID and Installment_Bus.Installment=BusFeePayment_Staff.Installment and BusFeePayment_Staff.Session=@d1 and BusCardHolder_Staff.Location=Location.Locationname and BusFeePayment_Staff.Installment=@d2 group by StaffName,Designation,BusCardHolder_Staff.Location,SchoolName order by 1", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbInstallment.Text)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            DataGridView1.DataSource = dtable
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("BusFeeFullDue_Staff.xml")
            Dim rpt As New rptBusFeeFullDue_Staff
            rpt.SetDataSource(ds)
            rpt.SetParameterValue("p1", cmbSession.Text)
            rpt.SetParameterValue("p2", cmbInstallment.Text)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbSession_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSession.SelectedIndexChanged
        Try
            cmbInstallment.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(Installment) FROM Installment_Bus"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            rdr = cmd.ExecuteReader()
            cmbInstallment.Items.Clear()
            While rdr.Read
                cmbInstallment.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
