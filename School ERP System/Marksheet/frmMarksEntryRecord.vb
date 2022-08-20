Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmMarksEntryRecord
    Public Sub GetData()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select M_ID,RTRIM(Student.AdmissionNo),RTRIM(EnrollmentNo),RTRIM(Student.StudentName),RTRIM(SchoolName),RTRIM(MarksEntry.Student_Class),RTRIM(MarksEntry.Session),EntryDate,RTRIM(MarksEntry.Result) from Student,MarksEntry,SchoolInfo where Student.AdmissionNo=MarksEntry.AdmissionNo and SchoolInfo.S_ID=Student.SchoolID order by StudentName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub txtStudentName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtStudentName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select M_ID,RTRIM(Student.AdmissionNo),RTRIM(EnrollmentNo),RTRIM(Student.StudentName),RTRIM(SchoolName),RTRIM(MarksEntry.Student_Class),RTRIM(MarksEntry.Session),EntryDate,RTRIM(MarksEntry.Result) from Student,MarksEntry,SchoolInfo where Student.AdmissionNo=MarksEntry.AdmissionNo and SchoolInfo.S_ID=Student.SchoolID and StudentName like '%" & txtStudentName.Text & "%' order by StudentName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            If cmbSession.Text = "" Then
                MessageBox.Show("Please select session", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSession.Focus()
                Exit Sub
            End If
            If cmbClass.Text = "" Then
                MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbClass.Focus()
                Exit Sub
            End If
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select M_ID,RTRIM(Student.AdmissionNo),RTRIM(EnrollmentNo),RTRIM(Student.StudentName),RTRIM(SchoolName),RTRIM(MarksEntry.Student_Class),RTRIM(MarksEntry.Session),EntryDate,RTRIM(MarksEntry.Result) from Student,MarksEntry,SchoolInfo where Student.AdmissionNo=MarksEntry.AdmissionNo and SchoolInfo.S_ID=Student.SchoolID and MarksEntry.Session=@d1 and Student_Class=@d2 order by StudentName", con)
            cmd.Parameters.AddWithValue("@d1", cmbSession.Text)
            cmd.Parameters.AddWithValue("@d2", cmbClass.Text)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Sub fillSession()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Session) FROM MarksEntry", con)
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

    Sub fillClass()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct (Student_Class) FROM MarksEntry order by 1", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbClass.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbClass.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtAdmissionNo_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtAdmissionNo.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select M_ID,RTRIM(Student.AdmissionNo),RTRIM(EnrollmentNo),RTRIM(Student.StudentName),RTRIM(SchoolName),RTRIM(MarksEntry.Student_Class),RTRIM(MarksEntry.Session),EntryDate,RTRIM(MarksEntry.Result) from Student,MarksEntry,SchoolInfo where Student.AdmissionNo=MarksEntry.AdmissionNo and SchoolInfo.S_ID=Student.SchoolID and Student.AdmissionNo like '%" & txtAdmissionNo.Text & "%' order by StudentName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        txtAdmissionNo.Text = ""
        txtStudentName.Text = ""
        cmbClass.SelectedIndex = -1
        cmbSession.SelectedIndex = -1
        cmbClass.Enabled = False
        GetData()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub frmStudentRecord_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSession()
        fillClass()
        GetData()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ExportExcel(dgw)
    End Sub

    Private Sub cmbSession_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSession.SelectedIndexChanged
        cmbClass.Enabled = True
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Marks Entry" Then
                    Me.Hide()
                    frmMarksEntry.Show()
                    frmMarksEntry.txtMarksID.Text = dr.Cells(0).Value.ToString()
                    frmMarksEntry.txtAdmissionNo.Text = dr.Cells(1).Value.ToString()
                    frmMarksEntry.txtEnrollmentNo.Text = dr.Cells(2).Value.ToString()
                    frmMarksEntry.txtStudentName.Text = dr.Cells(3).Value.ToString()
                    frmMarksEntry.txtSchoolName.Text = dr.Cells(4).Value.ToString() '
                    frmMarksEntry.txtClass.Text = dr.Cells(5).Value.ToString()
                    frmMarksEntry.txtSession.Text = dr.Cells(6).Value.ToString()
                    If dr.Cells(8).Value = "Pass" Then
                        frmMarksEntry.rbPass.Checked = True
                    Else
                        frmMarksEntry.rbFail.Checked = True
                    End If
                    con = New SqlConnection(cs)
                    con.Open()
                    cmd = New SqlCommand("SELECT RTRIM(SubCode),RTRIM(SubjectName), MaxMarks,MMPractical, CreditHour, OMTheory,OMPractical ,RTRIM(OGTheory), RTRIM(OGPractical), RTRIM(FinalGrade), GradePoint from MarksEntry,MarksEntry_Join,Subject where MarksEntry.M_ID=MarksEntry_Join.MarksID and Subject.SubjectCode=MarksEntry_Join.SubCode and MarksEntry.M_ID=" & dr.Cells(0).Value & "", con)
                    rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                    frmMarksEntry.dgw.Rows.Clear()
                    While (rdr.Read() = True)
                        frmMarksEntry.dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10))
                    End While
                    con.Close()
                    If lblUserType.Text = "Admin" Then
                        frmMarksEntry.btnDelete.Enabled = True
                        frmMarksEntry.btnUpdate.Enabled = True
                    Else
                        frmMarksEntry.btnDelete.Enabled = False
                        frmMarksEntry.btnUpdate.Enabled = False
                    End If
                    frmMarksEntry.btnSave.Enabled = False
                    frmMarksEntry.btnSelection.Enabled = False
                    frmMarksEntry.btnPrint.Enabled = True
                    frmMarksEntry.btnPrint1.Enabled = True
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

End Class
