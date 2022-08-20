<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStudentAttendance
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStudentAttendance))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.cmbAttendanceType = New System.Windows.Forms.ComboBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.listView1 = New System.Windows.Forms.ListView()
        Me.columnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.columnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtStaffID = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmbStaffName = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtSubjectCode = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cmbSubjectName = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSt_ID = New System.Windows.Forms.TextBox()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmbSchool = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbClass = New System.Windows.Forms.ComboBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.cmbSession = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.GroupBox5)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.listView1)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Location = New System.Drawing.Point(7, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1133, 666)
        Me.Panel1.TabIndex = 2
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.cmbAttendanceType)
        Me.GroupBox5.Location = New System.Drawing.Point(938, 165)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(183, 81)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Attendance Type"
        '
        'cmbAttendanceType
        '
        Me.cmbAttendanceType.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbAttendanceType.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbAttendanceType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAttendanceType.FormattingEnabled = True
        Me.cmbAttendanceType.Location = New System.Drawing.Point(18, 39)
        Me.cmbAttendanceType.Name = "cmbAttendanceType"
        Me.cmbAttendanceType.Size = New System.Drawing.Size(149, 21)
        Me.cmbAttendanceType.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.dtpDate)
        Me.GroupBox3.Location = New System.Drawing.Point(771, 165)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(161, 81)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Attendance Date"
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(18, 40)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(130, 20)
        Me.dtpDate.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.btnGetData)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnDelete)
        Me.Panel3.Controls.Add(Me.btnUpdate)
        Me.Panel3.Controls.Add(Me.btnSave)
        Me.Panel3.Controls.Add(Me.btnNew)
        Me.Panel3.Location = New System.Drawing.Point(371, 253)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(111, 208)
        Me.Panel3.TabIndex = 66
        '
        'btnGetData
        '
        Me.btnGetData.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnGetData.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnGetData.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetData.Location = New System.Drawing.Point(13, 137)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(82, 28)
        Me.btnGetData.TabIndex = 68
        Me.btnGetData.Text = "Get Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(13, 170)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(82, 28)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDelete.Enabled = False
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(13, 104)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(82, 28)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
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
        'listView1
        '
        Me.listView1.CheckBoxes = True
        Me.listView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.columnHeader1, Me.columnHeader3})
        Me.listView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.listView1.GridLines = True
        Me.listView1.Location = New System.Drawing.Point(9, 253)
        Me.listView1.Name = "listView1"
        Me.listView1.Size = New System.Drawing.Size(356, 406)
        Me.listView1.TabIndex = 65
        Me.listView1.UseCompatibleStateImageBehavior = False
        Me.listView1.View = System.Windows.Forms.View.Details
        '
        'columnHeader1
        '
        Me.columnHeader1.Text = "Admission No."
        Me.columnHeader1.Width = 100
        '
        'columnHeader3
        '
        Me.columnHeader3.Text = "Student Name"
        Me.columnHeader3.Width = 250
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.txtStaffID)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.cmbStaffName)
        Me.GroupBox2.Location = New System.Drawing.Point(390, 165)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(375, 81)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Staff Info"
        '
        'txtStaffID
        '
        Me.txtStaffID.Location = New System.Drawing.Point(266, 40)
        Me.txtStaffID.Name = "txtStaffID"
        Me.txtStaffID.ReadOnly = True
        Me.txtStaffID.Size = New System.Drawing.Size(100, 20)
        Me.txtStaffID.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(263, 17)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 34
        Me.Label6.Text = "Staff ID :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(14, 17)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(69, 13)
        Me.Label7.TabIndex = 33
        Me.Label7.Text = "Staff Name : "
        '
        'cmbStaffName
        '
        Me.cmbStaffName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbStaffName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbStaffName.FormattingEnabled = True
        Me.cmbStaffName.Location = New System.Drawing.Point(17, 40)
        Me.cmbStaffName.Name = "cmbStaffName"
        Me.cmbStaffName.Size = New System.Drawing.Size(235, 21)
        Me.cmbStaffName.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.txtSubjectCode)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.cmbSubjectName)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 165)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(375, 81)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Subject Info"
        '
        'txtSubjectCode
        '
        Me.txtSubjectCode.Location = New System.Drawing.Point(266, 40)
        Me.txtSubjectCode.Name = "txtSubjectCode"
        Me.txtSubjectCode.ReadOnly = True
        Me.txtSubjectCode.Size = New System.Drawing.Size(100, 20)
        Me.txtSubjectCode.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(263, 17)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(77, 13)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Subject Code :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(14, 17)
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
        Me.cmbSubjectName.Location = New System.Drawing.Point(17, 40)
        Me.cmbSubjectName.Name = "cmbSubjectName"
        Me.cmbSubjectName.Size = New System.Drawing.Size(235, 21)
        Me.cmbSubjectName.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.txtID)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtSt_ID)
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1112, 62)
        Me.Panel2.TabIndex = 0
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(115, 31)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(81, 20)
        Me.txtID.TabIndex = 45
        Me.txtID.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(419, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(203, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Students Attendance"
        '
        'txtSt_ID
        '
        Me.txtSt_ID.Location = New System.Drawing.Point(17, 31)
        Me.txtSt_ID.Name = "txtSt_ID"
        Me.txtSt_ID.Size = New System.Drawing.Size(81, 20)
        Me.txtSt_ID.TabIndex = 44
        Me.txtSt_ID.Visible = False
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
        Me.GroupBox4.Controls.Add(Me.btnSearch)
        Me.GroupBox4.Controls.Add(Me.cmbSession)
        Me.GroupBox4.Location = New System.Drawing.Point(9, 76)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox4.Size = New System.Drawing.Size(543, 81)
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
        'btnSearch
        '
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(442, 40)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(91, 21)
        Me.btnSearch.TabIndex = 5
        Me.btnSearch.Text = "&Search"
        Me.btnSearch.UseVisualStyleBackColor = True
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
        'frmStudentAttendance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(1148, 680)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmStudentAttendance"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
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
    Friend WithEvents txtSt_ID As System.Windows.Forms.TextBox
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Private WithEvents btnSearch As System.Windows.Forms.Button
    Private WithEvents columnHeader1 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbClass As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSession As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtStaffID As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmbStaffName As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSubjectCode As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmbSubjectName As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSchool As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbAttendanceType As System.Windows.Forms.ComboBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents btnGetData As System.Windows.Forms.Button
    Friend WithEvents listView1 As System.Windows.Forms.ListView

End Class
