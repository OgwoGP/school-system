<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStudentsAttendanceRecord
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStudentsAttendanceRecord))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblTotalClasses = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.dgw = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dtpDateTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpDateFrom = New System.Windows.Forms.DateTimePicker()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtSubjectCode = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cmbSubjectName = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmbSchool = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbClass = New System.Windows.Forms.ComboBox()
        Me.cmbSession = New System.Windows.Forms.ComboBox()
        Me.cmbCollegeName = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.lblTotalClasses)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.dgw)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Location = New System.Drawing.Point(7, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1109, 679)
        Me.Panel1.TabIndex = 2
        '
        'lblTotalClasses
        '
        Me.lblTotalClasses.AutoSize = True
        Me.lblTotalClasses.Font = New System.Drawing.Font("Segoe UI Semibold", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalClasses.ForeColor = System.Drawing.Color.MidnightBlue
        Me.lblTotalClasses.Location = New System.Drawing.Point(878, 340)
        Me.lblTotalClasses.Name = "lblTotalClasses"
        Me.lblTotalClasses.Size = New System.Drawing.Size(87, 30)
        Me.lblTotalClasses.TabIndex = 298
        Me.lblTotalClasses.Text = "Label12"
        Me.lblTotalClasses.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI Semibold", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(800, 290)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(255, 50)
        Me.Label11.TabIndex = 297
        Me.Label11.Text = "Total Classes :"
        '
        'dgw
        '
        Me.dgw.AllowUserToAddRows = False
        Me.dgw.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FloralWhite
        Me.dgw.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgw.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgw.BackgroundColor = System.Drawing.Color.White
        Me.dgw.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.LightSteelBlue
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgw.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dgw.ColumnHeadersHeight = 24
        Me.dgw.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4})
        Me.dgw.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgw.EnableHeadersVisualStyles = False
        Me.dgw.GridColor = System.Drawing.Color.White
        Me.dgw.Location = New System.Drawing.Point(9, 252)
        Me.dgw.MultiSelect = False
        Me.dgw.Name = "dgw"
        Me.dgw.ReadOnly = True
        Me.dgw.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgw.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgw.RowHeadersWidth = 25
        Me.dgw.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle7.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.Color.White
        Me.dgw.RowsDefaultCellStyle = DataGridViewCellStyle7
        Me.dgw.RowTemplate.Height = 18
        Me.dgw.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgw.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgw.Size = New System.Drawing.Size(677, 418)
        Me.dgw.TabIndex = 296
        '
        'Column1
        '
        Me.Column1.HeaderText = "Admission No."
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 150
        '
        'Column2
        '
        Me.Column2.HeaderText = "Student Name"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 250
        '
        'Column3
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column3.HeaderText = "Total Attendance"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 150
        '
        'Column4
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column4.HeaderText = "Attendance %"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.dtpDateTo)
        Me.GroupBox3.Controls.Add(Me.dtpDateFrom)
        Me.GroupBox3.Controls.Add(Me.btnSearch)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 165)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(395, 81)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Attendance Date"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(157, 21)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(29, 13)
        Me.Label7.TabIndex = 53
        Me.Label7.Text = "To : "
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(16, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 13)
        Me.Label6.TabIndex = 52
        Me.Label6.Text = "From : "
        '
        'dtpDateTo
        '
        Me.dtpDateTo.CustomFormat = "dd/MM/yyyy"
        Me.dtpDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDateTo.Location = New System.Drawing.Point(159, 41)
        Me.dtpDateTo.Name = "dtpDateTo"
        Me.dtpDateTo.Size = New System.Drawing.Size(130, 20)
        Me.dtpDateTo.TabIndex = 1
        '
        'dtpDateFrom
        '
        Me.dtpDateFrom.CustomFormat = "dd/MM/yyyy"
        Me.dtpDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDateFrom.Location = New System.Drawing.Point(18, 41)
        Me.dtpDateFrom.Name = "dtpDateFrom"
        Me.dtpDateFrom.Size = New System.Drawing.Size(130, 20)
        Me.dtpDateFrom.TabIndex = 0
        '
        'btnSearch
        '
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(295, 40)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(91, 21)
        Me.btnSearch.TabIndex = 5
        Me.btnSearch.Text = "&Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Controls.Add(Me.btnNew)
        Me.Panel3.Location = New System.Drawing.Point(410, 171)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(286, 75)
        Me.Panel3.TabIndex = 66
        '
        'btnClose
        '
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(189, 23)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(82, 28)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(101, 23)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(82, 28)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Export Excel"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnNew.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew.Location = New System.Drawing.Point(13, 23)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(82, 28)
        Me.btnNew.TabIndex = 0
        Me.btnNew.Text = "Reset"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.DateTimePicker1)
        Me.GroupBox1.Controls.Add(Me.DateTimePicker2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.txtSubjectCode)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.cmbSubjectName)
        Me.GroupBox1.Location = New System.Drawing.Point(459, 76)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(643, 81)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Subject Info"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(138, 21)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(29, 13)
        Me.Label12.TabIndex = 71
        Me.Label12.Text = "To : "
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(15, 21)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(39, 13)
        Me.Label13.TabIndex = 70
        Me.Label13.Text = "From : "
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(141, 41)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(113, 20)
        Me.DateTimePicker1.TabIndex = 69
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(18, 41)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(117, 20)
        Me.DateTimePicker2.TabIndex = 68
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(538, 40)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(76, 21)
        Me.Button1.TabIndex = 67
        Me.Button1.Text = "Search"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtSubjectCode
        '
        Me.txtSubjectCode.Location = New System.Drawing.Point(432, 41)
        Me.txtSubjectCode.Name = "txtSubjectCode"
        Me.txtSubjectCode.ReadOnly = True
        Me.txtSubjectCode.Size = New System.Drawing.Size(100, 20)
        Me.txtSubjectCode.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(429, 21)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(77, 13)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Subject Code :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(256, 21)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(83, 13)
        Me.Label9.TabIndex = 33
        Me.Label9.Text = "Subject Name : "
        '
        'cmbSubjectName
        '
        Me.cmbSubjectName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbSubjectName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbSubjectName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSubjectName.FormattingEnabled = True
        Me.cmbSubjectName.Location = New System.Drawing.Point(259, 41)
        Me.cmbSubjectName.Name = "cmbSubjectName"
        Me.cmbSubjectName.Size = New System.Drawing.Size(167, 21)
        Me.cmbSubjectName.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1093, 62)
        Me.Panel2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(410, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(267, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "List Of Students Attendance"
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Location = New System.Drawing.Point(47, 15)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(39, 13)
        Me.lblUser.TabIndex = 43
        Me.lblUser.Text = "Label8"
        Me.lblUser.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox4.Controls.Add(Me.Label10)
        Me.GroupBox4.Controls.Add(Me.cmbSchool)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.cmbClass)
        Me.GroupBox4.Controls.Add(Me.cmbSession)
        Me.GroupBox4.Location = New System.Drawing.Point(9, 76)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox4.Size = New System.Drawing.Size(444, 81)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(14, 17)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(80, 13)
        Me.Label10.TabIndex = 51
        Me.Label10.Text = "School Name : "
        '
        'cmbSchool
        '
        Me.cmbSchool.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbSchool.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbSchool.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSchool.FormattingEnabled = True
        Me.cmbSchool.Location = New System.Drawing.Point(17, 40)
        Me.cmbSchool.Name = "cmbSchool"
        Me.cmbSchool.Size = New System.Drawing.Size(163, 21)
        Me.cmbSchool.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(321, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Class :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(183, 17)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Session : "
        '
        'cmbClass
        '
        Me.cmbClass.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbClass.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbClass.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbClass.Enabled = False
        Me.cmbClass.FormattingEnabled = True
        Me.cmbClass.Location = New System.Drawing.Point(324, 40)
        Me.cmbClass.Name = "cmbClass"
        Me.cmbClass.Size = New System.Drawing.Size(112, 21)
        Me.cmbClass.TabIndex = 2
        '
        'cmbSession
        '
        Me.cmbSession.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbSession.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbSession.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSession.Enabled = False
        Me.cmbSession.FormattingEnabled = True
        Me.cmbSession.Location = New System.Drawing.Point(186, 40)
        Me.cmbSession.Name = "cmbSession"
        Me.cmbSession.Size = New System.Drawing.Size(124, 21)
        Me.cmbSession.TabIndex = 1
        '
        'cmbCollegeName
        '
        Me.cmbCollegeName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbCollegeName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbCollegeName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCollegeName.FormattingEnabled = True
        Me.cmbCollegeName.Location = New System.Drawing.Point(17, 40)
        Me.cmbCollegeName.Name = "cmbCollegeName"
        Me.cmbCollegeName.Size = New System.Drawing.Size(163, 21)
        Me.cmbCollegeName.TabIndex = 0
        '
        'frmStudentsAttendanceRecord
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(1122, 690)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmStudentsAttendanceRecord"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbClass As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSession As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSubjectCode As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmbSubjectName As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSchool As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpDateTo As System.Windows.Forms.DateTimePicker
    Private WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dgw As System.Windows.Forms.DataGridView
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblTotalClasses As System.Windows.Forms.Label
    Private WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cmbCollegeName As System.Windows.Forms.ComboBox

End Class
