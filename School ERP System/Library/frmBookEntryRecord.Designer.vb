<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBookEntryRecord
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBookEntryRecord))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpDateFrom = New System.Windows.Forms.DateTimePicker()
        Me.dtpDateTo = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.cmbCondition = New System.Windows.Forms.ComboBox()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.txtAccessionNo = New System.Windows.Forms.TextBox()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.txtSubCategory = New System.Windows.Forms.TextBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.txtCategory = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.txtClass = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.txtPublisher = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtLanguage = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtAuthor = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtBookTitle = New System.Windows.Forms.TextBox()
        Me.dgw = New System.Windows.Forms.DataGridView()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblSet = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox11.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel5.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.GroupBox11)
        Me.Panel1.Controls.Add(Me.GroupBox10)
        Me.Panel1.Controls.Add(Me.GroupBox9)
        Me.Panel1.Controls.Add(Me.GroupBox8)
        Me.Panel1.Controls.Add(Me.GroupBox7)
        Me.Panel1.Controls.Add(Me.GroupBox6)
        Me.Panel1.Controls.Add(Me.GroupBox5)
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.dgw)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(8, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1150, 660)
        Me.Panel1.TabIndex = 2
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.Button1)
        Me.GroupBox11.Controls.Add(Me.Label2)
        Me.GroupBox11.Controls.Add(Me.Label3)
        Me.GroupBox11.Controls.Add(Me.dtpDateFrom)
        Me.GroupBox11.Controls.Add(Me.dtpDateTo)
        Me.GroupBox11.Location = New System.Drawing.Point(411, 216)
        Me.GroupBox11.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox11.Size = New System.Drawing.Size(396, 64)
        Me.GroupBox11.TabIndex = 55
        Me.GroupBox11.TabStop = False
        Me.GroupBox11.Text = "By Bill Date"
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(302, 34)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(79, 23)
        Me.Button1.TabIndex = 47
        Me.Button1.Text = "&Get Data"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(180, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "To"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(24, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "From"
        '
        'dtpDateFrom
        '
        Me.dtpDateFrom.CustomFormat = "dd/MM/yyyy"
        Me.dtpDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDateFrom.Location = New System.Drawing.Point(27, 36)
        Me.dtpDateFrom.Name = "dtpDateFrom"
        Me.dtpDateFrom.Size = New System.Drawing.Size(131, 20)
        Me.dtpDateFrom.TabIndex = 0
        '
        'dtpDateTo
        '
        Me.dtpDateTo.CustomFormat = "dd/MM/yyyy"
        Me.dtpDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDateTo.Location = New System.Drawing.Point(173, 36)
        Me.dtpDateTo.Name = "dtpDateTo"
        Me.dtpDateTo.Size = New System.Drawing.Size(123, 20)
        Me.dtpDateTo.TabIndex = 1
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.cmbStatus)
        Me.GroupBox10.Location = New System.Drawing.Point(210, 216)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox10.TabIndex = 54
        Me.GroupBox10.TabStop = False
        Me.GroupBox10.Text = "Search By Status"
        '
        'cmbStatus
        '
        Me.cmbStatus.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbStatus.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Items.AddRange(New Object() {"Available", "Issued", "Reserved"})
        Me.cmbStatus.Location = New System.Drawing.Point(22, 28)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(162, 21)
        Me.cmbStatus.TabIndex = 21
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.cmbCondition)
        Me.GroupBox9.Location = New System.Drawing.Point(9, 216)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox9.TabIndex = 53
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Search By Condition"
        '
        'cmbCondition
        '
        Me.cmbCondition.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbCondition.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbCondition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCondition.FormattingEnabled = True
        Me.cmbCondition.Items.AddRange(New Object() {"Fresh", "Damaged"})
        Me.cmbCondition.Location = New System.Drawing.Point(22, 28)
        Me.cmbCondition.Name = "cmbCondition"
        Me.cmbCondition.Size = New System.Drawing.Size(162, 21)
        Me.cmbCondition.TabIndex = 21
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.txtAccessionNo)
        Me.GroupBox8.Location = New System.Drawing.Point(612, 146)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox8.TabIndex = 52
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Search By Accession No."
        '
        'txtAccessionNo
        '
        Me.txtAccessionNo.Location = New System.Drawing.Point(22, 29)
        Me.txtAccessionNo.Name = "txtAccessionNo"
        Me.txtAccessionNo.Size = New System.Drawing.Size(162, 20)
        Me.txtAccessionNo.TabIndex = 0
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.txtSubCategory)
        Me.GroupBox7.Location = New System.Drawing.Point(411, 146)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox7.TabIndex = 51
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Search By Sub Category"
        '
        'txtSubCategory
        '
        Me.txtSubCategory.Location = New System.Drawing.Point(22, 29)
        Me.txtSubCategory.Name = "txtSubCategory"
        Me.txtSubCategory.Size = New System.Drawing.Size(162, 20)
        Me.txtSubCategory.TabIndex = 0
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.txtCategory)
        Me.GroupBox6.Location = New System.Drawing.Point(210, 146)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox6.TabIndex = 50
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Search By Category"
        '
        'txtCategory
        '
        Me.txtCategory.Location = New System.Drawing.Point(22, 29)
        Me.txtCategory.Name = "txtCategory"
        Me.txtCategory.Size = New System.Drawing.Size(162, 20)
        Me.txtCategory.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.txtClass)
        Me.GroupBox5.Location = New System.Drawing.Point(9, 146)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox5.TabIndex = 49
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Search By Class"
        '
        'txtClass
        '
        Me.txtClass.Location = New System.Drawing.Point(22, 29)
        Me.txtClass.Name = "txtClass"
        Me.txtClass.Size = New System.Drawing.Size(162, 20)
        Me.txtClass.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtPublisher)
        Me.GroupBox4.Location = New System.Drawing.Point(612, 76)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox4.TabIndex = 48
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Search By Publisher"
        '
        'txtPublisher
        '
        Me.txtPublisher.Location = New System.Drawing.Point(22, 29)
        Me.txtPublisher.Name = "txtPublisher"
        Me.txtPublisher.Size = New System.Drawing.Size(162, 20)
        Me.txtPublisher.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtLanguage)
        Me.GroupBox3.Location = New System.Drawing.Point(411, 76)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox3.TabIndex = 47
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Search By Language"
        '
        'txtLanguage
        '
        Me.txtLanguage.Location = New System.Drawing.Point(22, 29)
        Me.txtLanguage.Name = "txtLanguage"
        Me.txtLanguage.Size = New System.Drawing.Size(162, 20)
        Me.txtLanguage.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtAuthor)
        Me.GroupBox2.Location = New System.Drawing.Point(210, 76)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox2.TabIndex = 46
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Search By Author"
        '
        'txtAuthor
        '
        Me.txtAuthor.Location = New System.Drawing.Point(22, 29)
        Me.txtAuthor.Name = "txtAuthor"
        Me.txtAuthor.Size = New System.Drawing.Size(162, 20)
        Me.txtAuthor.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtBookTitle)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 76)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(195, 64)
        Me.GroupBox1.TabIndex = 45
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Search By Book Title"
        '
        'txtBookTitle
        '
        Me.txtBookTitle.Location = New System.Drawing.Point(22, 29)
        Me.txtBookTitle.Name = "txtBookTitle"
        Me.txtBookTitle.Size = New System.Drawing.Size(162, 20)
        Me.txtBookTitle.TabIndex = 0
        '
        'dgw
        '
        Me.dgw.AllowUserToAddRows = False
        Me.dgw.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FloralWhite
        Me.dgw.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
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
        Me.dgw.ColumnHeadersHeight = 30
        Me.dgw.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.DefaultCellStyle = DataGridViewCellStyle3
        Me.dgw.EnableHeadersVisualStyles = False
        Me.dgw.GridColor = System.Drawing.Color.White
        Me.dgw.Location = New System.Drawing.Point(9, 294)
        Me.dgw.MultiSelect = False
        Me.dgw.Name = "dgw"
        Me.dgw.ReadOnly = True
        Me.dgw.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgw.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgw.RowHeadersWidth = 25
        Me.dgw.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.Color.White
        Me.dgw.RowsDefaultCellStyle = DataGridViewCellStyle5
        Me.dgw.RowTemplate.Height = 18
        Me.dgw.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgw.Size = New System.Drawing.Size(1128, 354)
        Me.dgw.TabIndex = 43
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.btnExportExcel)
        Me.Panel5.Controls.Add(Me.btnClose)
        Me.Panel5.Controls.Add(Me.btnReset)
        Me.Panel5.Location = New System.Drawing.Point(819, 81)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(284, 70)
        Me.Panel5.TabIndex = 42
        '
        'btnExportExcel
        '
        Me.btnExportExcel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnExportExcel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnExportExcel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExportExcel.Location = New System.Drawing.Point(101, 23)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(82, 28)
        Me.btnExportExcel.TabIndex = 5
        Me.btnExportExcel.Text = "Export Excel"
        Me.btnExportExcel.UseVisualStyleBackColor = True
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
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.lblSet)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1128, 62)
        Me.Panel2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(483, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "List Of Books"
        '
        'lblSet
        '
        Me.lblSet.AutoSize = True
        Me.lblSet.Location = New System.Drawing.Point(290, 17)
        Me.lblSet.Name = "lblSet"
        Me.lblSet.Size = New System.Drawing.Size(23, 13)
        Me.lblSet.TabIndex = 44
        Me.lblSet.Text = "Set"
        Me.lblSet.Visible = False
        '
        'frmBookEntryRecord
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(1165, 675)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBookEntryRecord"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox11.PerformLayout()
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel5.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents dgw As System.Windows.Forms.DataGridView
    Friend WithEvents btnExportExcel As System.Windows.Forms.Button
    Friend WithEvents lblSet As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPublisher As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtLanguage As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtAuthor As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtBookTitle As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents txtAccessionNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSubCategory As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents txtCategory As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents txtClass As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbStatus As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCondition As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox11 As System.Windows.Forms.GroupBox
    Private WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDateTo As System.Windows.Forms.DateTimePicker

End Class
