<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBookReservation
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBookReservation))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.btnCancelReservation = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtRemarks = New System.Windows.Forms.RichTextBox()
        Me.dtpReservationDate = New System.Windows.Forms.DateTimePicker()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtStaffID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtStaffName = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.txtS_ID = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.txtJointAuthor = New System.Windows.Forms.TextBox()
        Me.txtAuthor = New System.Windows.Forms.TextBox()
        Me.txtBookTitle = New System.Windows.Forms.TextBox()
        Me.txtAccessionNo = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtStatus)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.txtID)
        Me.Panel1.Controls.Add(Me.txtS_ID)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(7, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(550, 502)
        Me.Panel1.TabIndex = 2
        '
        'txtStatus
        '
        Me.txtStatus.Location = New System.Drawing.Point(295, 353)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(95, 20)
        Me.txtStatus.TabIndex = 9
        Me.txtStatus.Visible = False
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Transparent
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.btnCancelReservation)
        Me.Panel4.Location = New System.Drawing.Point(425, 327)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(111, 79)
        Me.Panel4.TabIndex = 8
        '
        'btnCancelReservation
        '
        Me.btnCancelReservation.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCancelReservation.Enabled = False
        Me.btnCancelReservation.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCancelReservation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancelReservation.Location = New System.Drawing.Point(12, 13)
        Me.btnCancelReservation.Name = "btnCancelReservation"
        Me.btnCancelReservation.Size = New System.Drawing.Size(87, 54)
        Me.btnCancelReservation.TabIndex = 0
        Me.btnCancelReservation.Text = "&Cancel Reservation"
        Me.btnCancelReservation.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.txtRemarks)
        Me.GroupBox3.Controls.Add(Me.dtpReservationDate)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.ForeColor = System.Drawing.Color.Black
        Me.GroupBox3.Location = New System.Drawing.Point(9, 300)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox3.Size = New System.Drawing.Size(280, 191)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Reservation Info"
        '
        'Label7
        '
        Me.Label7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(30, 55)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 13)
        Me.Label7.TabIndex = 322
        Me.Label7.Text = "Remarks :"
        '
        'txtRemarks
        '
        Me.txtRemarks.Location = New System.Drawing.Point(132, 55)
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtRemarks.Size = New System.Drawing.Size(137, 120)
        Me.txtRemarks.TabIndex = 321
        Me.txtRemarks.Text = ""
        '
        'dtpReservationDate
        '
        Me.dtpReservationDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpReservationDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpReservationDate.Location = New System.Drawing.Point(132, 29)
        Me.dtpReservationDate.Name = "dtpReservationDate"
        Me.dtpReservationDate.Size = New System.Drawing.Size(136, 20)
        Me.dtpReservationDate.TabIndex = 320
        '
        'Label6
        '
        Me.Label6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(30, 29)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 13)
        Me.Label6.TabIndex = 319
        Me.Label6.Text = "Reservation Date :"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.txtStaffID)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.txtStaffName)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Black
        Me.GroupBox2.Location = New System.Drawing.Point(9, 207)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(405, 90)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Staff Info"
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(234, 29)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(35, 21)
        Me.Button1.TabIndex = 320
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtStaffID
        '
        Me.txtStaffID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtStaffID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStaffID.Location = New System.Drawing.Point(133, 29)
        Me.txtStaffID.Name = "txtStaffID"
        Me.txtStaffID.ReadOnly = True
        Me.txtStaffID.Size = New System.Drawing.Size(95, 20)
        Me.txtStaffID.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(30, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 319
        Me.Label3.Text = "Staff ID :"
        '
        'txtStaffName
        '
        Me.txtStaffName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStaffName.Location = New System.Drawing.Point(133, 54)
        Me.txtStaffName.Name = "txtStaffName"
        Me.txtStaffName.ReadOnly = True
        Me.txtStaffName.Size = New System.Drawing.Size(253, 20)
        Me.txtStaffName.TabIndex = 1
        '
        'Label16
        '
        Me.Label16.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(30, 57)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(66, 13)
        Me.Label16.TabIndex = 299
        Me.Label16.Text = "Staff Name :"
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(295, 305)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(95, 20)
        Me.txtID.TabIndex = 3
        Me.txtID.Visible = False
        '
        'txtS_ID
        '
        Me.txtS_ID.Location = New System.Drawing.Point(295, 327)
        Me.txtS_ID.Name = "txtS_ID"
        Me.txtS_ID.Size = New System.Drawing.Size(95, 20)
        Me.txtS_ID.TabIndex = 2
        Me.txtS_ID.Visible = False
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
        Me.Panel3.Location = New System.Drawing.Point(425, 82)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(111, 208)
        Me.Panel3.TabIndex = 1
        '
        'btnGetData
        '
        Me.btnGetData.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnGetData.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnGetData.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetData.Location = New System.Drawing.Point(13, 134)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(82, 28)
        Me.btnGetData.TabIndex = 5
        Me.btnGetData.Text = "Get Data"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Hand
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
        Me.btnClose.Location = New System.Drawing.Point(13, 165)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(82, 28)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Cursor = System.Windows.Forms.Cursors.Hand
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
        Me.btnSave.Text = "Reserve"
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
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.txtJointAuthor)
        Me.GroupBox1.Controls.Add(Me.txtAuthor)
        Me.GroupBox1.Controls.Add(Me.txtBookTitle)
        Me.GroupBox1.Controls.Add(Me.txtAccessionNo)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label23)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 75)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(405, 127)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Book Info"
        '
        'Button2
        '
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(234, 23)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(35, 21)
        Me.Button2.TabIndex = 78
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'txtJointAuthor
        '
        Me.txtJointAuthor.Location = New System.Drawing.Point(133, 97)
        Me.txtJointAuthor.Name = "txtJointAuthor"
        Me.txtJointAuthor.ReadOnly = True
        Me.txtJointAuthor.Size = New System.Drawing.Size(253, 21)
        Me.txtJointAuthor.TabIndex = 4
        '
        'txtAuthor
        '
        Me.txtAuthor.Location = New System.Drawing.Point(133, 73)
        Me.txtAuthor.Name = "txtAuthor"
        Me.txtAuthor.ReadOnly = True
        Me.txtAuthor.Size = New System.Drawing.Size(253, 21)
        Me.txtAuthor.TabIndex = 3
        '
        'txtBookTitle
        '
        Me.txtBookTitle.Location = New System.Drawing.Point(133, 47)
        Me.txtBookTitle.Name = "txtBookTitle"
        Me.txtBookTitle.ReadOnly = True
        Me.txtBookTitle.Size = New System.Drawing.Size(253, 21)
        Me.txtBookTitle.TabIndex = 1
        '
        'txtAccessionNo
        '
        Me.txtAccessionNo.Location = New System.Drawing.Point(133, 23)
        Me.txtAccessionNo.Name = "txtAccessionNo"
        Me.txtAccessionNo.ReadOnly = True
        Me.txtAccessionNo.Size = New System.Drawing.Size(95, 21)
        Me.txtAccessionNo.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Location = New System.Drawing.Point(30, 97)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 15)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Joint Authors :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Location = New System.Drawing.Point(30, 73)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 15)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Author(s) :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(30, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 15)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Book Title :"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.BackColor = System.Drawing.Color.Transparent
        Me.Label23.Location = New System.Drawing.Point(30, 23)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(90, 15)
        Me.Label23.TabIndex = 0
        Me.Label23.Text = "Accession No. :"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(527, 62)
        Me.Panel2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(172, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(173, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Book Reservation"
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Location = New System.Drawing.Point(334, 27)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(39, 13)
        Me.lblUser.TabIndex = 5
        Me.lblUser.Text = "Label8"
        Me.lblUser.Visible = False
        '
        'frmBookReservation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(563, 514)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBookReservation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
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
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtJointAuthor As System.Windows.Forms.TextBox
    Friend WithEvents txtAuthor As System.Windows.Forms.TextBox
    Friend WithEvents txtBookTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtAccessionNo As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents txtS_ID As System.Windows.Forms.TextBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents btnGetData As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents txtStaffID As System.Windows.Forms.TextBox
    Private WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents txtStaffName As System.Windows.Forms.TextBox
    Private WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Private WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents dtpReservationDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Public WithEvents btnCancelReservation As System.Windows.Forms.Button
    Private WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.RichTextBox

End Class
