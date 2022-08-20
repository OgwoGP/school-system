Imports System.Data.SqlClient
Public Class frmMain

    Dim Filename As String

    Private Sub FeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FeeToolStripMenuItem.Click
        frmFee.lblUser.Text = lblUser.Text
        frmFee.Reset()
        frmFee.ShowDialog()
    End Sub

    Private Sub FeeInfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FeeInfoToolStripMenuItem.Click
        frmClassFeeInfo.lblUser.Text = lblUser.Text
        frmClassFeeInfo.Reset()
        frmClassFeeInfo.ShowDialog()
    End Sub

    Private Sub LocationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LocationToolStripMenuItem.Click
        frmLocation.lblUser.Text = lblUser.Text
        frmLocation.Reset()
        frmLocation.ShowDialog()
    End Sub

    Private Sub HostelInfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HostelInfoToolStripMenuItem.Click
        frmhostelInfo.lblUser.Text = lblUser.Text
        frmhostelInfo.Reset()
        frmhostelInfo.ShowDialog()
    End Sub

    Private Sub SchoolInfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SchoolInfoToolStripMenuItem.Click
        frmSchoolInfo.lblUser.Text = lblUser.Text
        frmSchoolInfo.Reset()
        frmSchoolInfo.ShowDialog()
    End Sub


    Private Sub CalculatorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalculatorToolStripMenuItem.Click
        System.Diagnostics.Process.Start("Calc.exe")
    End Sub

    Private Sub NotepadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotepadToolStripMenuItem.Click
        System.Diagnostics.Process.Start("Notepad.exe")
    End Sub

    Private Sub WordPadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordPadToolStripMenuItem.Click
        System.Diagnostics.Process.Start("WordPad.exe")
    End Sub

    Private Sub MSWordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MSWordToolStripMenuItem.Click
        System.Diagnostics.Process.Start("WinWord.exe")
    End Sub

    Private Sub MSPaintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MSPaintToolStripMenuItem.Click
        System.Diagnostics.Process.Start("MSPaint.exe")
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TaskManagerToolStripMenuItem.Click
        System.Diagnostics.Process.Start("TaskMgr.exe")
    End Sub

    Private Sub SystemInfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SystemInfoToolStripMenuItem.Click
        frmSystemInfo.Show()
    End Sub

    Private Sub BusFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BusFeeToolStripMenuItem.Click
        frmInstallment_Bus.lblUser.Text = lblUser.Text
        frmInstallment_Bus.Reset()
        frmInstallment_Bus.ShowDialog()
    End Sub

    Private Sub HostelFeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HostelFeeToolStripMenuItem.Click
        frmInstallment_Hostel.lblUser.Text = lblUser.Text
        frmInstallment_Hostel.Reset()
        frmInstallment_Hostel.ShowDialog()
    End Sub


    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub

    Sub LogOut()
        Dim st As String = "Successfully logged out"
        LogFunc(lblUser.Text, st)
        Me.Hide()
        frmLogin.Show()
        frmLogin.UserID.Text = ""
        frmLogin.Password.Text = ""
        frmLogin.UserID.Focus()
    End Sub

    Private Sub SessionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SessionToolStripMenuItem.Click
        frmSession.lblUser.Text = lblUser.Text
        frmSession.Reset()
        frmSession.ShowDialog()
    End Sub

    Private Sub BusInfoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BusInfoToolStripMenuItem.Click
        frmBusInfo.lblUser.Text = lblUser.Text
        frmBusInfo.Reset()
        frmBusInfo.ShowDialog()
    End Sub

    Private Sub SchoolTypeToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles SchoolTypeToolStripMenuItem.Click
        frmSchoolType.lblUser.Text = lblUser.Text
        frmSchoolType.Reset()
        frmSchoolType.ShowDialog()
    End Sub

    Private Sub frmMain_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True
    End Sub

    Private Sub StudentsToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles StudentsToolStripMenuItem1.Click
        frmDiscount.lblUser.Text = lblUser.Text
        frmDiscount.Reset()
        frmDiscount.ShowDialog()
    End Sub

    Private Sub StaffsToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles StaffsToolStripMenuItem1.Click
        frmDiscount_Staff.lblUser.Text = lblUser.Text
        frmDiscount_Staff.Reset()
        frmDiscount_Staff.ShowDialog()
    End Sub

    Private Sub DocumentMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DocumentMasterToolStripMenuItem.Click
        frmDocument.lblUser.Text = lblUser.Text
        frmDocument.Reset()
        frmDocument.ShowDialog()
    End Sub

    Private Sub StudentMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StudentMasterToolStripMenuItem.Click
        frmStudent.lblUser.Text = lblUser.Text
        frmStudent.Reset()
        frmStudent.ShowDialog()
    End Sub

    Private Sub HostelerMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HostelerMasterToolStripMenuItem.Click
        frmHosteler.lblUser.Text = lblUser.Text
        frmHosteler.Reset()
        frmHosteler.ShowDialog()
    End Sub

    Private Sub BusHoldersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BusHoldersToolStripMenuItem.Click
        frmBusCardHolder_Student.lblUser.Text = lblUser.Text
        frmBusCardHolder_Student.Reset()
        frmBusCardHolder_Student.ShowDialog()
    End Sub

    Private Sub PromotionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PromotionToolStripMenuItem.Click
        frmClassTransfer.lblUser.Text = lblUser.Text
        frmClassTransfer.Reset()
        frmClassTransfer.ShowDialog()
    End Sub

    Private Sub InactiveEntryToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles InactiveEntryToolStripMenuItem1.Click
        frmInactive.lblUser.Text = lblUser.Text
        frmInactive.Reset()
        frmInactive.ShowDialog()
    End Sub

    Private Sub DepartmentMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DepartmentMasterToolStripMenuItem.Click
        frmDepartment.lblUser.Text = lblUser.Text
        frmDepartment.Reset()
        frmDepartment.ShowDialog()
    End Sub

    Private Sub DesignationMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DesignationMasterToolStripMenuItem.Click
        frmDesignation.lblUser.Text = lblUser.Text
        frmDesignation.Reset()
        frmDesignation.ShowDialog()
    End Sub

    Private Sub StaffMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StaffMasterToolStripMenuItem.Click
        frmStaff.lblUser.Text = lblUser.Text
        frmStaff.Reset()
        frmStaff.ShowDialog()
    End Sub

    Private Sub AdvanceEntryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdvanceEntryToolStripMenuItem.Click
        frmAdvanceEntry.lblUser.Text = lblUser.Text
        frmAdvanceEntry.Reset()
        frmAdvanceEntry.ShowDialog()
    End Sub

    Private Sub AtteanceEntryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AtteanceEntryToolStripMenuItem.Click
        frmAttendance.lblUser.Text = lblUser.Text
        frmAttendance.Reset()
        frmAttendance.ShowDialog()
    End Sub

    Private Sub PaymentToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PaymentToolStripMenuItem.Click
        frmStaffPayment.lblUser.Text = lblUser.Text
        frmStaffPayment.Reset()
        frmStaffPayment.ShowDialog()
    End Sub

    Private Sub AtteandanceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AtteandanceToolStripMenuItem.Click
        frmStudentAttendance.lblUser.Text = lblUser.Text
        frmStudentAttendance.Reset()
        frmStudentAttendance.ShowDialog()
    End Sub

    Private Sub SettingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SettingToolStripMenuItem.Click
        frmSetting.lblUser.Text = lblUser.Text
        frmSetting.Reset()
        frmSetting.ShowDialog()
    End Sub

    Private Sub BookIssueToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BookIssueToolStripMenuItem.Click
        frmBookIssue.lblUser.Text = lblUser.Text
        frmBookIssue.Reset()
        frmBookIssue.Reset1()
        frmBookIssue.ShowDialog()
    End Sub

    Private Sub BookReturnToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BookReturnToolStripMenuItem.Click
        frmBookReturn.lblUser.Text = lblUser.Text
        frmBookReturn.Reset()
        frmBookReturn.Reset1()
        frmBookReturn.ShowDialog()
    End Sub

    Private Sub BookMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BookMasterToolStripMenuItem.Click
        frmBookEntry.lblUser.Text = lblUser.Text
        frmBookEntry.Reset()
        frmBookEntry.ShowDialog()
    End Sub

    Private Sub JournalsAndMagazinesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles JournalsAndMagazinesToolStripMenuItem.Click
        frmJournalsandMagazines.lblUser.Text = lblUser.Text
        frmJournalsandMagazines.Reset()
        frmJournalsandMagazines.ShowDialog()
    End Sub

    Private Sub SupplierMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SupplierMasterToolStripMenuItem.Click
        frmSupplier.lblUser.Text = lblUser.Text
        frmSupplier.Reset()
        frmSupplier.ShowDialog()
    End Sub

    Private Sub ClassToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ClassToolStripMenuItem.Click
        frmBooksClass.lblUser.Text = lblUser.Text
        frmBooksClass.Reset()
        frmBooksClass.ShowDialog()
    End Sub

    Private Sub CategoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CategoryToolStripMenuItem.Click
        frmCategory.lblUser.Text = lblUser.Text
        frmCategory.Reset()
        frmCategory.ShowDialog()
    End Sub

    Private Sub SubCategoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SubCategoryToolStripMenuItem.Click
        frmSubCategory.lblUser.Text = lblUser.Text
        frmSubCategory.Reset()
        frmSubCategory.ShowDialog()
    End Sub

    Private Sub NewspaperMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NewspaperMasterToolStripMenuItem.Click
        frmNewspaper_Master.lblUser.Text = lblUser.Text
        frmNewspaper_Master.Reset()
        frmNewspaper_Master.ShowDialog()
    End Sub

    Private Sub BooksReservationToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BooksReservationToolStripMenuItem.Click
        frmBookReservation.lblUser.Text = lblUser.Text
        frmBookReservation.Reset()
        frmBookReservation.ShowDialog()
    End Sub

    Private Sub QuotationToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles QuotationToolStripMenuItem.Click
        frmQuotation.lblUser.Text = lblUser.Text
        frmQuotation.Reset()
        frmQuotation.ShowDialog()
    End Sub


    Private Sub BooksToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BooksToolStripMenuItem.Click
        frmBookEntryRecord.Reset()
        frmBookEntryRecord.ShowDialog()
    End Sub

    Private Sub StudentsToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles StudentsToolStripMenuItem2.Click
        frmBookIssueRecord_Student.Reset()
        frmBookIssueRecord_Student.ShowDialog()
    End Sub

    Private Sub StaffsToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles StaffsToolStripMenuItem2.Click
        frmBookIssueRecord_Staff.Reset()
        frmBookIssueRecord_Staff1.ShowDialog()
    End Sub

    Private Sub BooksReservationToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles BooksReservationToolStripMenuItem1.Click
        frmBookReservationRecord.Reset()
        frmBookReservationRecord.ShowDialog()
    End Sub

    Private Sub StudentsToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles StudentsToolStripMenuItem3.Click
        frmBookReturnRecord_Student.Reset()
        frmBookReturnRecord_Student.ShowDialog()
    End Sub

    Private Sub StaffsToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles StaffsToolStripMenuItem3.Click
        frmBookReturnRecord_Staff.Reset()
        frmBookReturnRecord_Staff.ShowDialog()
    End Sub

    Private Sub SuppliersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SuppliersToolStripMenuItem.Click
        frmSupplierRecord.ShowDialog()
    End Sub


    Private Sub ExamTypeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExamTypeToolStripMenuItem.Click
        frmExamType.lblUser.Text = lblUser.Text
        frmExamType.Reset()
        frmExamType.ShowDialog()
    End Sub

    Private Sub SubjectMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SubjectMasterToolStripMenuItem.Click
        frmSubjectEntry.lblUser.Text = lblUser.Text
        frmSubjectEntry.Reset()
        frmSubjectEntry.ShowDialog()
    End Sub

    Private Sub AdvanceEntry1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdvanceEntry1ToolStripMenuItem.Click
        frmAdvanceEntryRecord.Reset()
        frmAdvanceEntryRecord.ShowDialog()
    End Sub

    Private Sub ProfileEntryToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ProfileEntryToolStripMenuItem1.Click
        frmStudentRecord1.Reset()
        frmStudentRecord1.ShowDialog()
    End Sub

    Private Sub HostelerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles HostelerToolStripMenuItem1.Click
        frmHostelerRecord.Reset()
        frmHostelerRecord.ShowDialog()
    End Sub

    Private Sub BusHolderToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles BusHolderToolStripMenuItem2.Click
        frmBusCardHolder_StudentRecord.Reset()
        frmBusCardHolder_StudentRecord.ShowDialog()
    End Sub

    Private Sub CurrentAdvanceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CurrentAdvanceToolStripMenuItem.Click
        'frmCurrentAdvance.Reset()
        'frmCurrentAdvance.ShowDialog()
    End Sub

    Private Sub StaffsAttendanceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StaffsAttendanceToolStripMenuItem.Click
        frmAttendanceEntryRecord.Reset()
        frmAttendanceEntryRecord.ShowDialog()
    End Sub

    Private Sub StaffsAttendance2ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StaffsAttendance2ToolStripMenuItem.Click
        frmAttendanceEntryRecord1.Reset()
        frmAttendanceEntryRecord1.ShowDialog()
    End Sub

    Private Sub StaffsPayment1ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StaffsPayment1ToolStripMenuItem.Click
        frmStaffPaymentRecord1.Reset()
        frmStaffPaymentRecord1.ShowDialog()
    End Sub

    Private Sub StaffsPayment2ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StaffsPayment2ToolStripMenuItem.Click
        frmStaffPaymentRecord.Reset()
        frmStaffPaymentRecord.ShowDialog()
    End Sub

    Private Sub ProfileEntryToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles ProfileEntryToolStripMenuItem3.Click
        frmStaffRecord1.Reset()
        frmStaffRecord1.ShowDialog()
    End Sub

    Private Sub BusHolderToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles BusHolderToolStripMenuItem3.Click
        frmBusCardHolder_StaffRecord.Reset()
        frmBusCardHolder_StaffRecord.ShowDialog()
    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim result As Boolean = HandleRegistry()
        If result = False Then 'something went wrong
            frmSoftwareLicenceRenewal.ShowDialog()
        End If
    End Sub


    Private Sub ExamEntryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExamEntryToolStripMenuItem.Click
        frmExamsEntry.lblUser.Text = lblUser.Text
        frmExamsEntry.Reset()
        frmExamsEntry.ShowDialog()
    End Sub

    Private Sub NewspaperEntryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NewspaperEntryToolStripMenuItem.Click
        frmNewspaper.lblUser.Text = lblUser.Text
        frmNewspaper.Reset()
        frmNewspaper.ShowDialog()
    End Sub


    Private Sub JournalsAndMagazinesToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles JournalsAndMagazinesToolStripMenuItem1.Click
        frmJournalsAndMagazinesRecord.Reset()
        frmJournalsAndMagazinesRecord.ShowDialog()
    End Sub

    Private Sub NewspaperToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NewspaperToolStripMenuItem.Click
        frmNewspaperRecord.Reset()
        frmNewspaperRecord.ShowDialog()
    End Sub

    Private Sub StudentsToolStripMenuItem4_Click(sender As System.Object, e As System.EventArgs) Handles StudentsToolStripMenuItem4.Click
        frmBookIssueReport_Student.Reset()
        frmBookIssueReport_Student.ShowDialog()
    End Sub

    Private Sub StaffsToolStripMenuItem4_Click(sender As System.Object, e As System.EventArgs) Handles StaffsToolStripMenuItem4.Click
        frmBookIssueReport_Staff.Reset()
        frmBookIssueReport_Staff.ShowDialog()
    End Sub

    Private Sub BooksReservationToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles BooksReservationToolStripMenuItem2.Click
        frmBookReservationReport.Reset()
        frmBookReservationReport.ShowDialog()
    End Sub

    Private Sub StudentsToolStripMenuItem5_Click(sender As System.Object, e As System.EventArgs) Handles StudentsToolStripMenuItem5.Click
        frmBookReturnReport_Student.Reset()
        frmBookReturnReport_Student.ShowDialog()
    End Sub

    Private Sub StaffsToolStripMenuItem5_Click(sender As System.Object, e As System.EventArgs) Handles StaffsToolStripMenuItem5.Click
        frmBookReturnReport_Staff.Reset()
        frmBookReturnReport_Staff.ShowDialog()
    End Sub

    Private Sub PurchasedBooksToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PurchasedBooksToolStripMenuItem.Click
        frmPurchasedBooksReport.Reset()
        frmPurchasedBooksReport.ShowDialog()
    End Sub


    Private Sub ToolStripMenuItem16_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem16.Click
        frmFineCollectionReport_Student.Reset()
        frmFineCollectionReport_Student.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem17_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem17.Click
        frmFineCollectionReport_Staff.Reset()
        frmFineCollectionReport_Staff.ShowDialog()
    End Sub

    Private Sub PayrollAdvanceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PayrollAdvanceToolStripMenuItem.Click
        frmAdvanceEntryReport.Reset()
        frmAdvanceEntryReport.ShowDialog()
    End Sub

    Private Sub StaffPaymentToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StaffPaymentToolStripMenuItem.Click
        frmStaffPaymentReport.Reset()
        frmStaffPaymentReport.ShowDialog()
    End Sub

    Private Sub SalarySlipToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SalarySlipToolStripMenuItem.Click
        frmSalaryslip.Reset()
        frmSalaryslip.ShowDialog()
    End Sub
    Private Sub UsersRegistrationToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UsersRegistrationToolStripMenuItem.Click
        frmRegistration.lblUser.Text = lblUser.Text
        frmRegistration.Reset()
        frmRegistration.ShowDialog()
    End Sub

    Private Sub SMSSettingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SMSSettingToolStripMenuItem.Click
        frmSMSSetting.Reset()
        frmSMSSetting.ShowDialog()
    End Sub

    Private Sub AttendanceMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AttendanceMasterToolStripMenuItem.Click
        frmAttendanceType.lblUser.Text = lblUser.Text
        frmAttendanceType.Reset()
        frmAttendanceType.ShowDialog()
    End Sub

    Private Sub AttendanceToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles AttendanceToolStripMenuItem1.Click
        frmStudentsAttendanceRecord.Reset()
        frmStudentsAttendanceRecord.ShowDialog()
    End Sub

    Private Sub ExpenseMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExpenseMasterToolStripMenuItem.Click
        Dim frm As New frmExpenseType
        frm.Reset()
        frm.lblUser.Text = lblUser.Text
        frm.ShowDialog()
    End Sub

    Private Sub ExpensesToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ExpensesToolStripMenuItem1.Click
        Dim frm As New frmExpense
        frm.Reset()
        frm.lblUser.Text = lblUser.Text
        frm.ShowDialog()
    End Sub

    Private Sub LogoutToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles LogoutToolStripMenuItem1.Click
        Try
            If MessageBox.Show("Do you really want to logout from application?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                If MessageBox.Show("Do you want backup database before logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    Backup()
                    LogOut()
                Else
                    LogOut()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RestoreToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RestoreToolStripMenuItem.Click
        Try
            With OpenFileDialog1
                .Filter = ("DB Backup File|*.bak;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""

            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Cursor = Cursors.WaitCursor
                Timer4.Enabled = True
                SqlConnection.ClearAllPools()
                con = New SqlConnection(cs)
                con.Open()
                Dim cb As String = "USE Master ALTER DATABASE SchoolDB SET Single_User WITH Rollback Immediate Restore database SchoolDB FROM disk='" & OpenFileDialog1.FileName & "' WITH REPLACE ALTER DATABASE SchoolDB SET Multi_User "
                cmd = New SqlCommand(cb)
                cmd.Connection = con
                cmd.ExecuteReader()
                con.Close()
                Dim st As String = "Sucessfully performed the restore"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully performed", "Database Restore", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Backup()
        Try
            Dim destdir As String = "SchoolDB " & DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss") & ".bak"
            Dim objdlg As New SaveFileDialog
            objdlg.FileName = destdir
            objdlg.ShowDialog()
            Filename = objdlg.FileName
            Cursor = Cursors.WaitCursor
            Timer4.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "backup database SchoolDB to disk='" & Filename & "'with init,stats=10"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer3_Tick(sender As System.Object, e As System.EventArgs) Handles Timer3.Tick
        lblDateTime.Text = Now.ToString("dd/MM/yyyy hh:mm:ss tt")
    End Sub

    Private Sub Timer4_Tick(sender As System.Object, e As System.EventArgs) Handles Timer4.Tick
        Cursor = Cursors.Default
        Timer4.Enabled = False
    End Sub

    Private Sub EmailSettingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EmailSettingToolStripMenuItem.Click
        frmEmailSetting.Reset()
        frmEmailSetting.ShowDialog()
    End Sub

    Private Sub SemesterMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SemesterMasterToolStripMenuItem.Click
        frmSection.lblUser.Text = lblUser.Text
        frmSection.Reset()
        frmSection.ShowDialog()
    End Sub

    Private Sub CourseMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CourseMasterToolStripMenuItem.Click
        frmClassType.lblUser.Text = lblUser.Text
        frmClassType.Reset()
        frmClassType.ShowDialog()
    End Sub

    Private Sub BranchMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BranchMasterToolStripMenuItem.Click
        frmClass.lblUser.Text = lblUser.Text
        frmClass.Reset()
        frmClass.ShowDialog()
    End Sub

    Private Sub DepartmentMasterToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles DepartmentMasterToolStripMenuItem1.Click
        frmDepartment.lblUser.Text = lblUser.Text
        frmDepartment.Reset()
        frmDepartment.ShowDialog()
    End Sub

    Private Sub DesignationMasterToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles DesignationMasterToolStripMenuItem1.Click
        frmDesignation.lblUser.Text = lblUser.Text
        frmDesignation.Reset()
        frmDesignation.ShowDialog()
    End Sub

    Private Sub ProfileEntryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProfileEntryToolStripMenuItem.Click
        frmStaff.lblUser.Text = lblUser.Text
        frmStaff.Reset()
        frmStaff.ShowDialog()
    End Sub

    Private Sub BusHoldersToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles BusHoldersToolStripMenuItem1.Click
        frmBusCardHolder_Staff.lblUser.Text = lblUser.Text
        frmBusCardHolder_Staff.Reset()
        frmBusCardHolder_Staff.ShowDialog()
    End Sub

    Private Sub VoucherToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles VoucherToolStripMenuItem1.Click
        frmVoucher.Reset()
        frmVoucher.lblUser.Text = lblUser.Text
        frmVoucher.ShowDialog()
    End Sub

    Private Sub HostelFeePaymentToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HostelFeePaymentToolStripMenuItem.Click
        frmHostelFeePayment.Reset()
        frmHostelFeePayment.lblUser.Text = lblUser.Text
        frmHostelFeePayment.lblUserType.Text = lblUserType.Text
        frmHostelFeePayment.ShowDialog()
    End Sub

    Private Sub ClassFeePaymentToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ClassFeePaymentToolStripMenuItem.Click
        frmClassFeePayment.Reset()
        frmClassFeePayment.lblUser.Text = lblUser.Text
        frmClassFeePayment.lblUserType.Text = lblUserType.Text
        frmClassFeePayment.ShowDialog()
    End Sub

    Private Sub StaffsToolStripMenuItem8_Click(sender As System.Object, e As System.EventArgs) Handles StaffsToolStripMenuItem8.Click
        frmBusFeePayment_Staff.Reset()
        frmBusFeePayment_Staff.lblUser.Text = lblUser.Text
        frmBusFeePayment_Staff.lblUserType.Text = lblUserType.Text
        frmBusFeePayment_Staff.ShowDialog()
    End Sub

    Private Sub StudentsToolStripMenuItem8_Click(sender As System.Object, e As System.EventArgs) Handles StudentsToolStripMenuItem8.Click
        frmBusFeePayment_Student.Reset()
        frmBusFeePayment_Student.lblUser.Text = lblUser.Text
        frmBusFeePayment_Student.lblUserType.Text = lblUserType.Text
        frmBusFeePayment_Student.ShowDialog()
    End Sub

    Private Sub BusFeePaymentStudentsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BusFeePaymentStudentsToolStripMenuItem.Click
        frmBusFeePayment_StudentRecord.Reset()
        frmBusFeePayment_StudentRecord.ShowDialog()
    End Sub

    Private Sub BusFeePaymentStaffsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BusFeePaymentStaffsToolStripMenuItem.Click
        frmBusFeePayment_StaffRecord.Reset()
        frmBusFeePayment_StaffRecord.ShowDialog()
    End Sub

    Private Sub ClassFeePaymentToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ClassFeePaymentToolStripMenuItem1.Click
        frmClassFeePaymentRecord.Reset()
        frmClassFeePaymentRecord.ShowDialog()
    End Sub

    Private Sub HostelFeePaymentToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles HostelFeePaymentToolStripMenuItem1.Click
        frmHostelFeePaymentRecord.Reset()
        frmHostelFeePaymentRecord.ShowDialog()
    End Sub

    Private Sub StudentsToolStripMenuItem6_Click(sender As System.Object, e As System.EventArgs) Handles StudentsToolStripMenuItem6.Click
        frmBusFeeReceipt_Student.Reset()
        frmBusFeeReceipt_Student.ShowDialog()
    End Sub

    Private Sub StaffsToolStripMenuItem6_Click(sender As System.Object, e As System.EventArgs) Handles StaffsToolStripMenuItem6.Click
        frmBusFeeReceipt_Staff.Reset()
        frmBusFeeReceipt_Staff.ShowDialog()
    End Sub

    Private Sub CourseFeeReceiptToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CourseFeeReceiptToolStripMenuItem.Click
        frmClassFeeReceipt.Reset()
        frmClassFeeReceipt.ShowDialog()
    End Sub

    Private Sub HostelFeeReceiptToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HostelFeeReceiptToolStripMenuItem.Click
        frmHostelFeeReceipt.Reset()
        frmHostelFeeReceipt.ShowDialog()
    End Sub

    Private Sub AttendanceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AttendanceToolStripMenuItem.Click
        frmStudentsAttendanceRecord.Reset()
        frmStudentsAttendanceRecord.ShowDialog()
    End Sub

    Private Sub SQLServerSettingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SQLServerSettingToolStripMenuItem.Click
        frmSqlServerSetting.Reset()
        frmSqlServerSetting.lblSet.Text = "Main Form"
        frmSqlServerSetting.ShowDialog()
    End Sub

    Private Sub DaybookToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DaybookToolStripMenuItem.Click
        frmGeneralDayBook.Reset()
        frmGeneralDayBook.ShowDialog()
    End Sub

    Private Sub GeneralLedgerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GeneralLedgerToolStripMenuItem.Click
        frmGeneralLedger.Reset()
        frmGeneralLedger.ShowDialog()
    End Sub

    Private Sub TrialBalanceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TrialBalanceToolStripMenuItem.Click
        frmTrialBalance.Reset()
        frmTrialBalance.ShowDialog()
    End Sub

    Private Sub BackupToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BackupToolStripMenuItem.Click
        Backup()
    End Sub

  
    Private Sub PartialDueToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PartialDueToolStripMenuItem.Click
        frmPartialDueList_Students.Reset()
        frmPartialDueList_Students.ShowDialog()
    End Sub

    Private Sub FullDueToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FullDueToolStripMenuItem.Click
        frmFullDueList_Students.Reset()
        frmFullDueList_Students.ShowDialog()
    End Sub

    Private Sub PartialDueToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles PartialDueToolStripMenuItem1.Click
        frmPartialDueList_Staff.Reset()
        frmPartialDueList_Staff.ShowDialog()
    End Sub

    Private Sub FullDueToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles FullDueToolStripMenuItem1.Click
        frmFullDueList_Staff.Reset()
        frmFullDueList_Staff.ShowDialog()
    End Sub

    Private Sub GradeMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GradeMasterToolStripMenuItem.Click
        frmGradeMaster.lblUser.Text = lblUser.Text
        frmGradeMaster.Reset()
        frmGradeMaster.ShowDialog()
    End Sub

    Private Sub MarksEntryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MarksEntryToolStripMenuItem.Click
        frmMarksEntry.Reset()
        frmMarksEntry.lblUser.Text = lblUser.Text
        frmMarksEntry.lblUserType.Text = lblUserType.Text
        frmMarksEntry.ShowDialog()
    End Sub

    Private Sub MarksEntryToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles MarksEntryToolStripMenuItem1.Click
        frmMarksEntryRecord.Reset()
        frmMarksEntryRecord.ShowDialog()
    End Sub

    Private Sub MarksheetLedgerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MarksheetLedgerToolStripMenuItem.Click
        frmMarksheetLedger.Reset()
        frmMarksheetLedger.ShowDialog()
    End Sub

   
    Private Sub VoucherToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VoucherToolStripMenuItem.Click
        frmVoucherReport.Reset()
        frmVoucherReport.ShowDialog()
    End Sub

    Private Sub LogsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LogsToolStripMenuItem.Click
        frmLogs.Reset()
        frmLogs.lblUser.Text = lblUser.Text
        frmLogs.ShowDialog()
    End Sub

    Private Sub HouseInfoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HouseInfoToolStripMenuItem.Click
        frmHouse.lblUser.Text = lblUser.Text
        frmHouse.Reset()
        frmHouse.ShowDialog()
    End Sub

    Private Sub StudentsRecordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StudentsRecordToolStripMenuItem.Click
        frmStudentRecord_ExcelSheet.Reset()
        frmStudentRecord_ExcelSheet.ShowDialog()
    End Sub

    Private Sub StaffsRecordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StaffsRecordToolStripMenuItem.Click
        frmStaffRecord_ExcelSheet.Reset()
        frmStaffRecord_ExcelSheet.ShowDialog()
    End Sub

    Private Sub BooksRecordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BooksRecordToolStripMenuItem.Click
        frmBookEntryRecord_ExcelSheet.Reset()
        frmBookEntryRecord_ExcelSheet.ShowDialog()
    End Sub

    Private Sub IdentityCardToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles IdentityCardToolStripMenuItem.Click
        frmIdentityCard.Reset()
        frmIdentityCard.ShowDialog()
    End Sub
End Class