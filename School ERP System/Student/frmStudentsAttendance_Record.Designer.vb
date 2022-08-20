<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStudentsAttendance_Record
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStudentsAttendance_Record))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dgw = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.dtpDateTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpDateFrom = New System.Windows.Forms.DateTimePicker()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.btnReset = New System.Windows.Forms.Button()
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
        Me.cmbCollegeName = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.dgw)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(7, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1044, 679)
        Me.Panel1.TabIndex = 2
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
        Me.dgw.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column11, Me.Column2, Me.Column3, Me.Column4, Me.Column7, Me.Column5, Me.Column6, Me.Column8, Me.Column9, Me.Column10})
        Me.dgw.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgw.EnableHeadersVisualStyles = False
        Me.dgw.GridColor = System.Drawing.Color.White
        Me.dgw.Location = New System.Drawing.Point(9, 252)
        Me.dgw.MultiSelect = False
        Me.dgw.Name = "dgw"
        Me.dgw.ReadOnly = True
        Me.dgw.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgw.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgw.RowHeadersWidth = 25
        Me.dgw.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle6.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.Color.White
        Me.dgw.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.dgw.RowTemplate.Height = 18
        Me.dgw.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgw.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgw.Size = New System.Drawing.Size(1028, 418)
        Me.dgw.TabIndex = 296
        '
        'Column1
        '
        Me.Column1.HeaderText = "Attendance ID"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        '
        'Column11
        '
        Me.Column11.HeaderText = "Attendance Type"
        Me.Column11.Name = "Column11"
        Me.Column11.ReadOnly = True
        '
        'Column2
        '
        DataGridViewCellStyle3.Format = "dd/MM/yyyy"
        Me.Column2.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column2.HeaderText = "Date"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        '
        'Column3
        '
        Me.Column3.HeaderText = "Subject Code"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        '
        'Column4
        '
        Me.Column4.HeaderText = "Subject Name"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        '
        'Column7
        '
        Me.Column7.HeaderText = "School"
        Me.Column7.Name = "Column7"
        Me.Column7.ReadOnly = True
        '
        'Column5
        '
        Me.Column5.HeaderText = "Class"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        '
        'Column6
        '
        Me.Column6.HeaderText = "Session"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        '
        'Column8
        '
        Me.Column8.HeaderText = "ST ID"
        Me.Column8.Name = "Column8"
        Me.Column8.ReadOnly = True
        Me.Column8.Visible = False
        '
        'Column9
        '
        Me.Column9.HeaderText = "Staff ID"
        Me.Column9.Name = "Column9"
        Me.Column9.ReadOnly = True
        '
        'Column10
        '
        Me.Column10.HeaderText = "Staff Name"
        Me.Column10.Name = "Column10"
        Me.Column10.ReadOnly = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.dtpDateTo)
        Me.GroupBox3.Controls.Add(Me.dtpDateFrom)
        Me.GroupBox3.Controls.Add(Me.btnSearch)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 75)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(395, 81)
        Me.GroupBox3.TabIndex = 0
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
        Me.btnSearch.TabIndex = 2
        Me.btnSearch.Text = "&Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnExportExcel)
        Me.Panel3.Controls.Add(Me.btnReset)
        Me.Panel3.Location = New System.Drawing.Point(410, 81)
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
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnExportExcel
        '
        Me.btnExportExcel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnExportExcel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnExportExcel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExportExcel.Location = New System.Drawing.Point(101, 23)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(82, 28)
        Me.btnExportExcel.TabIndex = 1
        Me.btnExportExcel.Text = "Export Excel"
        Me.btnExportExcel.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnReset.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnReset.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReset.Location = New System.Drawing.Point(13, 23)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(82, 28)
        Me.btnReset.TabIndex = 0
        Me.btnReset.Text = "Reset"
        Me.btnReset.UseVisualStyleBackColor = True
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
        Me.GroupBox1.Location = New System.Drawing.Point(9, 163)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(692, 81)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Subject Info"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(157, 21)
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
        Me.DateTimePicker1.Location = New System.Drawing.Point(159, 41)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(130, 20)
        Me.DateTimePicker1.TabIndex = 1
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(18, 41)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(130, 20)
        Me.DateTimePicker2.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(592, 40)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(91, 21)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Search"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtSubjectCode
        '
        Me.txtSubjectCode.Location = New System.Drawing.Point(486, 41)
        Me.txtSubjectCode.Name = "txtSubjectCode"
        Me.txtSubjectCode.ReadOnly = True
        Me.txtSubjectCode.Size = New System.Drawing.Size(100, 20)
        Me.txtSubjectCode.TabIndex = 3
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(483, 21)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(77, 13)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Subject Code :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(292, 21)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(83, 13)
        Me.Label9.TabIndex = 33
        Me.Label9.Text = "Subject Name : "
        '
        'cmbSubjectName
        '
        Me.cmbSubjectName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbSubjectName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbSubjectName.FormattingEnabled = True
        Me.cmbSubjectName.Location = New System.Drawing.Point(295, 41)
        Me.cmbSubjectName.Name = "cmbSubjectName"
        Me.cmbSubjectName.Size = New System.Drawing.Size(185, 21)
        Me.cmbSubjectName.TabIndex = 2
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1028, 62)
        Me.Panel2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(397, 15)
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
        'frmStudentsAttendance_Record
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(1057, 690)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmStudentsAttendance_Record"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSubjectCode As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmbSubjectName As System.Windows.Forms.ComboBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnExportExcel As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpDateTo As System.Windows.Forms.DateTimePicker
    Private WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dgw As System.Windows.Forms.DataGridView
    Private WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbCollegeName As System.Windows.Forms.ComboBox
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
