<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMarksheetLedger
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMarksheetLedger))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
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
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Location = New System.Drawing.Point(7, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(754, 138)
        Me.Panel1.TabIndex = 2
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Controls.Add(Me.btnClose)
        Me.Panel3.Controls.Add(Me.btnNew)
        Me.Panel3.Location = New System.Drawing.Point(459, 56)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(286, 75)
        Me.Panel3.TabIndex = 66
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(9, 23)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(82, 28)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "View Report"
        Me.Button1.UseVisualStyleBackColor = True
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
        'btnNew
        '
        Me.btnNew.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnNew.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew.Location = New System.Drawing.Point(99, 23)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(82, 28)
        Me.btnNew.TabIndex = 0
        Me.btnNew.Text = "Reset"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(736, 39)
        Me.Panel2.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(270, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(178, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Marksheet Ledger"
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
        Me.GroupBox4.Location = New System.Drawing.Point(9, 51)
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
        'Timer1
        '
        '
        'frmMarksheetLedger
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(768, 152)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMarksheetLedger"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
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
    Friend WithEvents cmbSchool As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents cmbCollegeName As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
