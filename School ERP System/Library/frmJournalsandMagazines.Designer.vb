<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmJournalsandMagazines
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmJournalsandMagazines))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.txtSupplierID = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtNumber = New System.Windows.Forms.TextBox()
        Me.txtRemarks = New System.Windows.Forms.TextBox()
        Me.cmbDepartment = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.dtpPaidOn = New System.Windows.Forms.DateTimePicker()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.dtpBillDate = New System.Windows.Forms.DateTimePicker()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtBillNo = New System.Windows.Forms.TextBox()
        Me.dtpSubDateTo = New System.Windows.Forms.DateTimePicker()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.dtpSubDateFrom = New System.Windows.Forms.DateTimePicker()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtSub = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.dtpSubDate = New System.Windows.Forms.DateTimePicker()
        Me.txtSubNo = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpDateOfReceipt = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtVolume = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.cmbMonth = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.txtIssueNo = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtS_ID = New System.Windows.Forms.TextBox()
        Me.txtDepartmentID = New System.Windows.Forms.TextBox()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Location = New System.Drawing.Point(7, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(787, 702)
        Me.Panel1.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtNumber)
        Me.GroupBox1.Controls.Add(Me.txtRemarks)
        Me.GroupBox1.Controls.Add(Me.cmbDepartment)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Controls.Add(Me.dtpPaidOn)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.txtAmount)
        Me.GroupBox1.Controls.Add(Me.dtpBillDate)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.txtBillNo)
        Me.GroupBox1.Controls.Add(Me.dtpSubDateTo)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.dtpSubDateFrom)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.txtSub)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.dtpSubDate)
        Me.GroupBox1.Controls.Add(Me.txtSubNo)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtName)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.dtpDateOfReceipt)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dtpDate)
        Me.GroupBox1.Controls.Add(Me.Label20)
        Me.GroupBox1.Controls.Add(Me.txtVolume)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.cmbMonth)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtYear)
        Me.GroupBox1.Controls.Add(Me.txtIssueNo)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label22)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 71)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox1.Size = New System.Drawing.Size(647, 618)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Journals And Magazines"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.txtLastName)
        Me.GroupBox2.Controls.Add(Me.txtFirstName)
        Me.GroupBox2.Controls.Add(Me.Label26)
        Me.GroupBox2.Controls.Add(Me.Label27)
        Me.GroupBox2.Controls.Add(Me.Label28)
        Me.GroupBox2.Controls.Add(Me.txtSupplierID)
        Me.GroupBox2.Location = New System.Drawing.Point(39, 389)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(477, 100)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Supplier Info"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(217, 20)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(28, 21)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(301, 46)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "First Name :"
        '
        'txtLastName
        '
        Me.txtLastName.BackColor = System.Drawing.SystemColors.Control
        Me.txtLastName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastName.Location = New System.Drawing.Point(141, 64)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.ReadOnly = True
        Me.txtLastName.Size = New System.Drawing.Size(157, 21)
        Me.txtLastName.TabIndex = 1
        '
        'txtFirstName
        '
        Me.txtFirstName.BackColor = System.Drawing.SystemColors.Control
        Me.txtFirstName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstName.Location = New System.Drawing.Point(304, 64)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.ReadOnly = True
        Me.txtFirstName.Size = New System.Drawing.Size(166, 21)
        Me.txtFirstName.TabIndex = 2
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(138, 46)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(64, 13)
        Me.Label26.TabIndex = 20
        Me.Label26.Text = "Last Name :"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(23, 46)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(82, 13)
        Me.Label27.TabIndex = 19
        Me.Label27.Text = "Supplier Name :"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(23, 23)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(65, 13)
        Me.Label28.TabIndex = 15
        Me.Label28.Text = "Supplier ID :"
        '
        'txtSupplierID
        '
        Me.txtSupplierID.BackColor = System.Drawing.SystemColors.Control
        Me.txtSupplierID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSupplierID.Location = New System.Drawing.Point(141, 20)
        Me.txtSupplierID.Name = "txtSupplierID"
        Me.txtSupplierID.ReadOnly = True
        Me.txtSupplierID.Size = New System.Drawing.Size(70, 21)
        Me.txtSupplierID.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Location = New System.Drawing.Point(36, 503)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 13)
        Me.Label4.TabIndex = 109
        Me.Label4.Text = "Department :"
        '
        'txtNumber
        '
        Me.txtNumber.Location = New System.Drawing.Point(209, 338)
        Me.txtNumber.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtNumber.Name = "txtNumber"
        Me.txtNumber.Size = New System.Drawing.Size(118, 20)
        Me.txtNumber.TabIndex = 15
        '
        'txtRemarks
        '
        Me.txtRemarks.Location = New System.Drawing.Point(209, 532)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtRemarks.Multiline = True
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtRemarks.Size = New System.Drawing.Size(245, 79)
        Me.txtRemarks.TabIndex = 19
        '
        'cmbDepartment
        '
        Me.cmbDepartment.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbDepartment.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbDepartment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDepartment.FormattingEnabled = True
        Me.cmbDepartment.Location = New System.Drawing.Point(209, 505)
        Me.cmbDepartment.Name = "cmbDepartment"
        Me.cmbDepartment.Size = New System.Drawing.Size(228, 21)
        Me.cmbDepartment.TabIndex = 18
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Location = New System.Drawing.Point(36, 532)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(55, 13)
        Me.Label11.TabIndex = 77
        Me.Label11.Text = "Remarks :"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.BackColor = System.Drawing.Color.Transparent
        Me.Label21.Location = New System.Drawing.Point(36, 338)
        Me.Label21.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(50, 13)
        Me.Label21.TabIndex = 106
        Me.Label21.Text = "Number :"
        '
        'dtpPaidOn
        '
        Me.dtpPaidOn.CustomFormat = "dd/MM/yyyy"
        Me.dtpPaidOn.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPaidOn.Location = New System.Drawing.Point(406, 191)
        Me.dtpPaidOn.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtpPaidOn.Name = "dtpPaidOn"
        Me.dtpPaidOn.Size = New System.Drawing.Size(159, 20)
        Me.dtpPaidOn.TabIndex = 10
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Location = New System.Drawing.Point(331, 191)
        Me.Label18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(51, 13)
        Me.Label18.TabIndex = 104
        Me.Label18.Text = "Paid On :"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(209, 181)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(100, 20)
        Me.txtAmount.TabIndex = 7
        '
        'dtpBillDate
        '
        Me.dtpBillDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpBillDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpBillDate.Location = New System.Drawing.Point(406, 164)
        Me.dtpBillDate.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtpBillDate.Name = "dtpBillDate"
        Me.dtpBillDate.Size = New System.Drawing.Size(159, 20)
        Me.dtpBillDate.TabIndex = 8
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.BackColor = System.Drawing.Color.Transparent
        Me.Label17.Location = New System.Drawing.Point(331, 164)
        Me.Label17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(52, 13)
        Me.Label17.TabIndex = 101
        Me.Label17.Text = "Bill Date :"
        '
        'txtBillNo
        '
        Me.txtBillNo.Location = New System.Drawing.Point(209, 156)
        Me.txtBillNo.Name = "txtBillNo"
        Me.txtBillNo.Size = New System.Drawing.Size(100, 20)
        Me.txtBillNo.TabIndex = 6
        '
        'dtpSubDateTo
        '
        Me.dtpSubDateTo.CustomFormat = "dd/MM/yyyy"
        Me.dtpSubDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSubDateTo.Location = New System.Drawing.Point(368, 130)
        Me.dtpSubDateTo.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtpSubDateTo.Name = "dtpSubDateTo"
        Me.dtpSubDateTo.Size = New System.Drawing.Size(159, 20)
        Me.dtpSubDateTo.TabIndex = 5
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Location = New System.Drawing.Point(365, 111)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(26, 13)
        Me.Label16.TabIndex = 98
        Me.Label16.Text = "To :"
        '
        'dtpSubDateFrom
        '
        Me.dtpSubDateFrom.CustomFormat = "dd/MM/yyyy"
        Me.dtpSubDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSubDateFrom.Location = New System.Drawing.Point(209, 130)
        Me.dtpSubDateFrom.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtpSubDateFrom.Name = "dtpSubDateFrom"
        Me.dtpSubDateFrom.Size = New System.Drawing.Size(135, 20)
        Me.dtpSubDateFrom.TabIndex = 4
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Location = New System.Drawing.Point(207, 111)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(36, 13)
        Me.Label15.TabIndex = 96
        Me.Label15.Text = "From :"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Location = New System.Drawing.Point(36, 111)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(126, 13)
        Me.Label14.TabIndex = 95
        Me.Label14.Text = "Subscription Per Annum :"
        '
        'txtSub
        '
        Me.txtSub.Location = New System.Drawing.Point(210, 86)
        Me.txtSub.Name = "txtSub"
        Me.txtSub.Size = New System.Drawing.Size(100, 20)
        Me.txtSub.TabIndex = 3
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Location = New System.Drawing.Point(36, 86)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(71, 13)
        Me.Label13.TabIndex = 93
        Me.Label13.Text = "Subscription :"
        '
        'dtpSubDate
        '
        Me.dtpSubDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpSubDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSubDate.Location = New System.Drawing.Point(461, 62)
        Me.dtpSubDate.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtpSubDate.Name = "dtpSubDate"
        Me.dtpSubDate.Size = New System.Drawing.Size(159, 20)
        Me.dtpSubDate.TabIndex = 2
        '
        'txtSubNo
        '
        Me.txtSubNo.Location = New System.Drawing.Point(210, 61)
        Me.txtSubNo.Name = "txtSubNo"
        Me.txtSubNo.Size = New System.Drawing.Size(100, 20)
        Me.txtSubNo.TabIndex = 1
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Location = New System.Drawing.Point(36, 156)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(46, 13)
        Me.Label12.TabIndex = 90
        Me.Label12.Text = "Bill No. :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Location = New System.Drawing.Point(36, 181)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(49, 13)
        Me.Label9.TabIndex = 89
        Me.Label9.Text = "Amount :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Location = New System.Drawing.Point(331, 64)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(97, 13)
        Me.Label8.TabIndex = 88
        Me.Label8.Text = "Subscription Date :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Location = New System.Drawing.Point(36, 61)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(91, 13)
        Me.Label6.TabIndex = 87
        Me.Label6.Text = "Subscription No. :"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(210, 36)
        Me.txtName.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(245, 20)
        Me.txtName.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Location = New System.Drawing.Point(36, 36)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(33, 13)
        Me.Label3.TabIndex = 84
        Me.Label3.Text = "Title :"
        '
        'dtpDateOfReceipt
        '
        Me.dtpDateOfReceipt.CustomFormat = "dd/MM/yyyy"
        Me.dtpDateOfReceipt.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDateOfReceipt.Location = New System.Drawing.Point(209, 363)
        Me.dtpDateOfReceipt.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtpDateOfReceipt.Name = "dtpDateOfReceipt"
        Me.dtpDateOfReceipt.Size = New System.Drawing.Size(159, 20)
        Me.dtpDateOfReceipt.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(36, 363)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 13)
        Me.Label2.TabIndex = 83
        Me.Label2.Text = "Date Of Receipt :"
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(209, 233)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(135, 20)
        Me.dtpDate.TabIndex = 11
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.Color.Transparent
        Me.Label20.Location = New System.Drawing.Point(36, 233)
        Me.Label20.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(36, 13)
        Me.Label20.TabIndex = 81
        Me.Label20.Text = "Date :"
        '
        'txtVolume
        '
        Me.txtVolume.Location = New System.Drawing.Point(209, 312)
        Me.txtVolume.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtVolume.Name = "txtVolume"
        Me.txtVolume.Size = New System.Drawing.Size(118, 20)
        Me.txtVolume.TabIndex = 14
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Location = New System.Drawing.Point(36, 314)
        Me.Label19.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(48, 13)
        Me.Label19.TabIndex = 79
        Me.Label19.Text = "Volume :"
        '
        'cmbMonth
        '
        Me.cmbMonth.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbMonth.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMonth.FormattingEnabled = True
        Me.cmbMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.cmbMonth.Location = New System.Drawing.Point(209, 259)
        Me.cmbMonth.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.cmbMonth.Name = "cmbMonth"
        Me.cmbMonth.Size = New System.Drawing.Size(159, 21)
        Me.cmbMonth.TabIndex = 12
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Location = New System.Drawing.Point(36, 258)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 13)
        Me.Label7.TabIndex = 76
        Me.Label7.Text = "Month :"
        '
        'txtYear
        '
        Me.txtYear.Location = New System.Drawing.Point(209, 286)
        Me.txtYear.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(99, 20)
        Me.txtYear.TabIndex = 13
        '
        'txtIssueNo
        '
        Me.txtIssueNo.Location = New System.Drawing.Point(209, 207)
        Me.txtIssueNo.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.txtIssueNo.Name = "txtIssueNo"
        Me.txtIssueNo.Size = New System.Drawing.Size(101, 20)
        Me.txtIssueNo.TabIndex = 9
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Location = New System.Drawing.Point(36, 286)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(35, 13)
        Me.Label10.TabIndex = 75
        Me.Label10.Text = "Year :"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.BackColor = System.Drawing.Color.Transparent
        Me.Label22.Location = New System.Drawing.Point(36, 207)
        Me.Label22.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(58, 13)
        Me.Label22.TabIndex = 72
        Me.Label22.Text = "Issue No. :"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.txtS_ID)
        Me.Panel2.Controls.Add(Me.txtDepartmentID)
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Controls.Add(Me.txtID)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(766, 62)
        Me.Panel2.TabIndex = 0
        '
        'txtS_ID
        '
        Me.txtS_ID.Location = New System.Drawing.Point(643, 21)
        Me.txtS_ID.Name = "txtS_ID"
        Me.txtS_ID.Size = New System.Drawing.Size(81, 20)
        Me.txtS_ID.TabIndex = 46
        Me.txtS_ID.Visible = False
        '
        'txtDepartmentID
        '
        Me.txtDepartmentID.Location = New System.Drawing.Point(110, 20)
        Me.txtDepartmentID.Name = "txtDepartmentID"
        Me.txtDepartmentID.Size = New System.Drawing.Size(81, 20)
        Me.txtDepartmentID.TabIndex = 45
        Me.txtDepartmentID.Visible = False
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Location = New System.Drawing.Point(598, 20)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(39, 13)
        Me.lblUser.TabIndex = 43
        Me.lblUser.Text = "Label8"
        Me.lblUser.Visible = False
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(23, 20)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(81, 20)
        Me.txtID.TabIndex = 44
        Me.txtID.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(226, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(290, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Journals and Magazines Entry"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnGetData)
        Me.Panel3.Controls.Add(Me.btnDelete)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Controls.Add(Me.btnNew)
        Me.Panel3.Location = New System.Drawing.Point(664, 76)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(111, 203)
        Me.Panel3.TabIndex = 41
        '
        'btnGetData
        '
        Me.btnGetData.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnGetData.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnGetData.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetData.Location = New System.Drawing.Point(13, 135)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(82, 28)
        Me.btnGetData.TabIndex = 4
        Me.btnGetData.Text = "Get Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDelete.Enabled = False
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(13, 103)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(82, 28)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(13, 167)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(82, 28)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.Location = New System.Drawing.Point(13, 72)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdate.TabIndex = 2
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(13, 41)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(82, 28)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnNew.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew.Location = New System.Drawing.Point(13, 10)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(82, 28)
        Me.btnNew.TabIndex = 0
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmJournalsandMagazines
        '
        Me.AcceptButton = Me.btnSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(801, 716)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmJournalsandMagazines"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents txtDepartmentID As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnGetData As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbDepartment As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents dtpPaidOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents dtpBillDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtBillNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpSubDateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents dtpSubDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtSub As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents dtpSubDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtSubNo As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpDateOfReceipt As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtVolume As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents txtIssueNo As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierID As System.Windows.Forms.TextBox
    Friend WithEvents txtS_ID As System.Windows.Forms.TextBox

End Class
