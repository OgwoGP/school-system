Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports Microsoft.Win32

Public Class frmSoftwareLicenceRenewal
    Public Shared Function GetUniqueKey(ByVal maxSize As Integer) As String
        Dim chars As Char() = New Char(61) {}
        chars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
        Dim data As Byte() = New Byte(0) {}
        Dim crypto As New RNGCryptoServiceProvider()
        crypto.GetNonZeroBytes(data)
        data = New Byte(maxSize - 1) {}
        crypto.GetNonZeroBytes(data)
        Dim result As New StringBuilder(maxSize)
        For Each b As Byte In data
            result.Append(chars(b Mod (chars.Length)))
        Next
        Return result.ToString()
    End Function
    Private Sub frmActivation_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        txtSerialNo.Text = GetUniqueKey(15)
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnExtendLicence.Click
        Try
            If txtActivationID.Text = "" Then
                MessageBox.Show("Please enter activation id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtActivationID.Focus()
                Exit Sub
            End If
            Dim st As String = (txtSerialNo.Text)
            TextBox1.Text = Encryption.MakePassword(st, 567)
            If txtActivationID.Text = TextBox1.Text Then
                'Declare variables
                Dim oRegKey As RegistryKey

                'Define subkey to work with
                oRegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\WOW6432Node", True)
                'Delete our subkey
                oRegKey.DeleteSubKeyTree("SchoolERP4")

                'Close the RegistryKey
                oRegKey.Close()
                Dim RenewalDate As Date
                RenewalDate = System.DateTime.Today.Date.AddYears(1)
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\SchoolERP4", "Set", RenewalDate)
                MessageBox.Show("Successfully renewal", "Software Licence", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End
            Else
                MessageBox.Show("Invalid activation id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        End
    End Sub
End Class
