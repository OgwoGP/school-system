Imports System.Data.SqlClient
Public Class frmLogin
    Dim frm As New frmMain

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If Len(Trim(UserID.Text)) = 0 Then
            MessageBox.Show("Please enter user id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            UserID.Focus()
            Exit Sub
        End If
        If Len(Trim(Password.Text)) = 0 Then
            MessageBox.Show("Please enter password", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Password.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT RTRIM(UserID),RTRIM(Password) FROM Registration where UserID = @d1 and Password=@d2"
            cmd.Parameters.AddWithValue("@d1", UserID.Text)
            cmd.Parameters.AddWithValue("@d2", Encrypt(Password.Text))
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                con = New SqlConnection(cs)
                con.Open()
                cmd = con.CreateCommand()
                cmd.CommandText = "SELECT usertype FROM Registration where UserID=@d3 and Password=@d4"
                cmd.Parameters.AddWithValue("@d3", UserID.Text)
                cmd.Parameters.AddWithValue("@d4", Encrypt(Password.Text))
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    UserType.Text = rdr.GetValue(0).ToString.Trim
                End If
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If

                If UserType.Text = "Admin" Then
                    frm.lblUser.Text = UserID.Text
                    frm.lblUserType.Text = UserType.Text
                    frm.MasterEntryToolStripMenuItem.Enabled = True
                    frm.StudentToolStripMenuItem1.Enabled = True
                    frm.StaffsToolStripMenuItem7.Enabled = True
                    frm.ExamToolStripMenuItem.Enabled = True
                    frm.PayrollToolStripMenuItem.Enabled = True
                    frm.LibraryToolStripMenuItem.Enabled = True
                    frm.TransactionsToolStripMenuItem2.Enabled = True
                    frm.AccountingToolStripMenuItem.Enabled = True
                    frm.UtilitiesToolStripMenuItem.Enabled = True
                    frm.RecordsToolStripMenuItem.Enabled = True
                    frm.ReportToolStripMenuItem.Enabled = True
                    frm.DatabaseToolStripMenuItem1.Enabled = True
                    Dim st As String = "Successfully logged in"
                    LogFunc(UserID.Text, st)
                    Me.Hide()
                    frm.Show()
                End If

                If UserType.Text = "Accountant" Then
                    frm.lblUser.Text = UserID.Text
                    frm.lblUserType.Text = UserType.Text
                    frm.MasterEntryToolStripMenuItem.Enabled = False
                    frm.StudentToolStripMenuItem1.Enabled = False
                    frm.StaffsToolStripMenuItem7.Enabled = False
                    frm.ExamToolStripMenuItem.Enabled = False
                    frm.PayrollToolStripMenuItem.Enabled = True
                    frm.LibraryToolStripMenuItem.Enabled = False
                    frm.TransactionsToolStripMenuItem2.Enabled = True
                    frm.AccountingToolStripMenuItem.Enabled = True
                    frm.UtilitiesToolStripMenuItem.Enabled = False
                    frm.RecordsToolStripMenuItem.Enabled = False
                    frm.ReportToolStripMenuItem.Enabled = False
                    frm.DatabaseToolStripMenuItem1.Enabled = False
                    Dim st As String = "Successfully logged in"
                    LogFunc(UserID.Text, st)
                    Me.Hide()
                    frm.Show()
                End If
                If UserType.Text = "Admission Officer" Then
                    frm.lblUser.Text = UserID.Text
                    frm.lblUserType.Text = UserType.Text
                    frm.MasterEntryToolStripMenuItem.Enabled = False
                    frm.StudentToolStripMenuItem1.Enabled = True
                    frm.StaffsToolStripMenuItem7.Enabled = False
                    frm.ExamToolStripMenuItem.Enabled = False
                    frm.PayrollToolStripMenuItem.Enabled = False
                    frm.LibraryToolStripMenuItem.Enabled = False
                    frm.TransactionsToolStripMenuItem2.Enabled = False
                    frm.AccountingToolStripMenuItem.Enabled = False
                    frm.UtilitiesToolStripMenuItem.Enabled = False
                    frm.RecordsToolStripMenuItem.Enabled = False
                    frm.ReportToolStripMenuItem.Enabled = False
                    frm.DatabaseToolStripMenuItem1.Enabled = False
                    Dim st As String = "Successfully logged in"
                    LogFunc(UserID.Text, st)
                    Me.Hide()
                    frm.Show()
                End If
                If UserType.Text = "Librarian" Then
                    frm.lblUser.Text = UserID.Text
                    frm.lblUserType.Text = UserType.Text
                    frm.MasterEntryToolStripMenuItem.Enabled = False
                    frm.StudentToolStripMenuItem1.Enabled = False
                    frm.StaffsToolStripMenuItem7.Enabled = False
                    frm.ExamToolStripMenuItem.Enabled = False
                    frm.PayrollToolStripMenuItem.Enabled = False
                    frm.LibraryToolStripMenuItem.Enabled = True
                    frm.TransactionsToolStripMenuItem2.Enabled = False
                    frm.AccountingToolStripMenuItem.Enabled = False
                    frm.UtilitiesToolStripMenuItem.Enabled = False
                    frm.RecordsToolStripMenuItem.Enabled = False
                    frm.ReportToolStripMenuItem.Enabled = False
                    frm.DatabaseToolStripMenuItem1.Enabled = False
                    Dim st As String = "Successfully logged in"
                    LogFunc(UserID.Text, st)
                    Me.Hide()
                    frm.Show()
                End If
            Else
                MsgBox("Login is Failed...Try again !", MsgBoxStyle.Critical, "Login Denied")
                UserID.Text = ""
                Password.Text = ""
                UserID.Focus()
            End If
            cmd.Dispose()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        End
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.Hide()
        frmChangePassword.Show()
        frmChangePassword.UserID.Text = ""
        frmChangePassword.OldPassword.Text = ""
        frmChangePassword.NewPassword.Text = ""
        frmChangePassword.ConfirmPassword.Text = ""
        frmChangePassword.UserID.Focus()
    End Sub

    Private Sub LoginForm1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmLogin_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        End
    End Sub

End Class
