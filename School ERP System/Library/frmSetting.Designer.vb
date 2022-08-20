<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetting
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSetting))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.txtMaxBooksAllowedStaff = New System.Windows.Forms.TextBox()
        Me.txtMaxBooksAllowedStudent = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cmbBookType = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtMaxDaysAllowedStaff = New System.Windows.Forms.TextBox()
        Me.txtMaxDaysAllowedStudent = New System.Windows.Forms.TextBox()
        Me.txtFinePerDayStaff = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFinePerDayStudent = New System.Windows.Forms.TextBox()
        Me.dgw = New System.Windows.Forms.DataGridView()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.dgw)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(6, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(748, 377)
        Me.Panel1.TabIndex = 2
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnDelete)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Controls.Add(Me.btnNew)
        Me.Panel3.Location = New System.Drawing.Point(327, 75)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(111, 174)
        Me.Panel3.TabIndex = 41
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
        Me.btnClose.Location = New System.Drawing.Point(13, 134)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(82, 28)
        Me.btnClose.TabIndex = 4
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
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.txtMaxBooksAllowedStaff)
        Me.Panel4.Controls.Add(Me.txtMaxBooksAllowedStudent)
        Me.Panel4.Controls.Add(Me.Label7)
        Me.Panel4.Controls.Add(Me.Label8)
        Me.Panel4.Controls.Add(Me.cmbBookType)
        Me.Panel4.Controls.Add(Me.Label6)
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Controls.Add(Me.Label4)
        Me.Panel4.Controls.Add(Me.Label3)
        Me.Panel4.Controls.Add(Me.txtMaxDaysAllowedStaff)
        Me.Panel4.Controls.Add(Me.txtMaxDaysAllowedStudent)
        Me.Panel4.Controls.Add(Me.txtFinePerDayStaff)
        Me.Panel4.Controls.Add(Me.Label2)
        Me.Panel4.Controls.Add(Me.txtFinePerDayStudent)
        Me.Panel4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel4.Location = New System.Drawing.Point(9, 75)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(312, 218)
        Me.Panel4.TabIndex = 0
        '
        'txtMaxBooksAllowedStaff
        '
        Me.txtMaxBooksAllowedStaff.BackColor = System.Drawing.Color.White
        Me.txtMaxBooksAllowedStaff.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxBooksAllowedStaff.Location = New System.Drawing.Point(202, 155)
        Me.txtMaxBooksAllowedStaff.Name = "txtMaxBooksAllowedStaff"
        Me.txtMaxBooksAllowedStaff.Size = New System.Drawing.Size(99, 21)
        Me.txtMaxBooksAllowedStaff.TabIndex = 5
        Me.txtMaxBooksAllowedStaff.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxBooksAllowedStudent
        '
        Me.txtMaxBooksAllowedStudent.BackColor = System.Drawing.Color.White
        Me.txtMaxBooksAllowedStudent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxBooksAllowedStudent.Location = New System.Drawing.Point(201, 182)
        Me.txtMaxBooksAllowedStudent.Name = "txtMaxBooksAllowedStudent"
        Me.txtMaxBooksAllowedStudent.Size = New System.Drawing.Size(100, 21)
        Me.txtMaxBooksAllowedStudent.TabIndex = 6
        Me.txtMaxBooksAllowedStudent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(17, 155)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(158, 15)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Max. Books Allowed (Staff) :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(17, 182)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(176, 15)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Max. Books Allowed (Student) :"
        '
        'cmbBookType
        '
        Me.cmbBookType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbBookType.FormattingEnabled = True
        Me.cmbBookType.Items.AddRange(New Object() {"Normal", "Reference"})
        Me.cmbBookType.Location = New System.Drawing.Point(202, 16)
        Me.cmbBookType.Name = "cmbBookType"
        Me.cmbBookType.Size = New System.Drawing.Size(100, 23)
        Me.cmbBookType.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(17, 17)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 15)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Book Type :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(17, 99)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(118, 15)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Fine Per Day (Staff) :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(17, 126)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 15)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Fine Per Day (Student) :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 45)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(151, 15)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Max. Days Allowed (Staff) :"
        '
        'txtMaxDaysAllowedStaff
        '
        Me.txtMaxDaysAllowedStaff.BackColor = System.Drawing.Color.White
        Me.txtMaxDaysAllowedStaff.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxDaysAllowedStaff.Location = New System.Drawing.Point(203, 45)
        Me.txtMaxDaysAllowedStaff.Name = "txtMaxDaysAllowedStaff"
        Me.txtMaxDaysAllowedStaff.Size = New System.Drawing.Size(99, 21)
        Me.txtMaxDaysAllowedStaff.TabIndex = 1
        Me.txtMaxDaysAllowedStaff.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxDaysAllowedStudent
        '
        Me.txtMaxDaysAllowedStudent.BackColor = System.Drawing.Color.White
        Me.txtMaxDaysAllowedStudent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxDaysAllowedStudent.Location = New System.Drawing.Point(202, 72)
        Me.txtMaxDaysAllowedStudent.Name = "txtMaxDaysAllowedStudent"
        Me.txtMaxDaysAllowedStudent.Size = New System.Drawing.Size(100, 21)
        Me.txtMaxDaysAllowedStudent.TabIndex = 2
        Me.txtMaxDaysAllowedStudent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFinePerDayStaff
        '
        Me.txtFinePerDayStaff.BackColor = System.Drawing.Color.White
        Me.txtFinePerDayStaff.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFinePerDayStaff.Location = New System.Drawing.Point(202, 99)
        Me.txtFinePerDayStaff.Name = "txtFinePerDayStaff"
        Me.txtFinePerDayStaff.Size = New System.Drawing.Size(99, 21)
        Me.txtFinePerDayStaff.TabIndex = 3
        Me.txtFinePerDayStaff.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(169, 15)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Max. Days Allowed (Student) :"
        '
        'txtFinePerDayStudent
        '
        Me.txtFinePerDayStudent.BackColor = System.Drawing.Color.White
        Me.txtFinePerDayStudent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFinePerDayStudent.Location = New System.Drawing.Point(202, 126)
        Me.txtFinePerDayStudent.Name = "txtFinePerDayStudent"
        Me.txtFinePerDayStudent.Size = New System.Drawing.Size(100, 21)
        Me.txtFinePerDayStudent.TabIndex = 4
        Me.txtFinePerDayStudent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        Me.dgw.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column6, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column1, Me.Column7})
        Me.dgw.Cursor = System.Windows.Forms.Cursors.Hand
        Me.dgw.EnableHeadersVisualStyles = False
        Me.dgw.GridColor = System.Drawing.Color.White
        Me.dgw.Location = New System.Drawing.Point(9, 299)
        Me.dgw.MultiSelect = False
        Me.dgw.Name = "dgw"
        Me.dgw.ReadOnly = True
        Me.dgw.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgw.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgw.RowHeadersWidth = 25
        Me.dgw.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.White
        Me.dgw.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.dgw.RowTemplate.Height = 18
        Me.dgw.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgw.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgw.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgw.Size = New System.Drawing.Size(729, 70)
        Me.dgw.TabIndex = 40
        '
        'Column6
        '
        Me.Column6.HeaderText = "Book Type"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        '
        'Column2
        '
        Me.Column2.HeaderText = "Max. Days Allowed (Staff)"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        '
        'Column3
        '
        Me.Column3.HeaderText = "Max. Days Allowed (Student)"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        '
        'Column4
        '
        Me.Column4.HeaderText = "Fine Per Day (Staff)"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        '
        'Column5
        '
        Me.Column5.HeaderText = "Fine Per Day (Student)"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        '
        'Column1
        '
        Me.Column1.HeaderText = "Max. Books Allowed (Staff)"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        '
        'Column7
        '
        Me.Column7.HeaderText = "Max. Books Allowed (Student)"
        Me.Column7.Name = "Column7"
        Me.Column7.ReadOnly = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(729, 62)
        Me.Panel2.TabIndex = 0
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Location = New System.Drawing.Point(29, 30)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(39, 13)
        Me.lblUser.TabIndex = 43
        Me.lblUser.Text = "Label8"
        Me.lblUser.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(278, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Setting"
        '
        'frmSetting
        '
        Me.AcceptButton = Me.btnSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(761, 390)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.dgw, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtMaxDaysAllowedStaff As System.Windows.Forms.TextBox
    Friend WithEvents dgw As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFinePerDayStudent As System.Windows.Forms.TextBox
    Friend WithEvents txtFinePerDayStaff As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxDaysAllowedStudent As System.Windows.Forms.TextBox
    Friend WithEvents cmbBookType As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtMaxBooksAllowedStaff As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxBooksAllowedStudent As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
