<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQuotation
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
        Dim Label5 As System.Windows.Forms.Label
        Dim Label4 As System.Windows.Forms.Label
        Dim Label10 As System.Windows.Forms.Label
        Dim Label15 As System.Windows.Forms.Label
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQuotation))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtRemarks = New System.Windows.Forms.RichTextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtReceiverCN = New System.Windows.Forms.TextBox()
        Me.txtReceiverAddress = New System.Windows.Forms.TextBox()
        Me.txtReceiverName = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSenderContactNo = New System.Windows.Forms.TextBox()
        Me.txtSenderAddress = New System.Windows.Forms.TextBox()
        Me.txtSenderName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.cmbDepartment = New System.Windows.Forms.ComboBox()
        Me.dtpOrderDate = New System.Windows.Forms.DateTimePicker()
        Me.txtOrderNo = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.txtQty = New System.Windows.Forms.TextBox()
        Me.txtAccessionNo = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtPublisher = New System.Windows.Forms.TextBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.txtAuthor = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtBookTitle = New System.Windows.Forms.TextBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtDepartmentID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtOrderID = New System.Windows.Forms.TextBox()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Label5 = New System.Windows.Forms.Label()
        Label4 = New System.Windows.Forms.Label()
        Label10 = New System.Windows.Forms.Label()
        Label15 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label5
        '
        Label5.AutoSize = True
        Label5.ForeColor = System.Drawing.Color.Black
        Label5.Location = New System.Drawing.Point(26, 22)
        Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Label5.Name = "Label5"
        Label5.Size = New System.Drawing.Size(59, 13)
        Label5.TabIndex = 268
        Label5.Text = "Order No. :"
        '
        'Label4
        '
        Label4.AutoSize = True
        Label4.ForeColor = System.Drawing.Color.Black
        Label4.Location = New System.Drawing.Point(26, 26)
        Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Label4.Name = "Label4"
        Label4.Size = New System.Drawing.Size(85, 13)
        Label4.TabIndex = 268
        Label4.Text = "Sender's Name :"
        '
        'Label10
        '
        Label10.AutoSize = True
        Label10.ForeColor = System.Drawing.Color.Black
        Label10.Location = New System.Drawing.Point(26, 26)
        Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Label10.Name = "Label10"
        Label10.Size = New System.Drawing.Size(94, 13)
        Label10.TabIndex = 268
        Label10.Text = "Receiver's Name :"
        '
        'Label15
        '
        Label15.AutoSize = True
        Label15.ForeColor = System.Drawing.Color.Black
        Label15.Location = New System.Drawing.Point(26, 71)
        Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Label15.Name = "Label15"
        Label15.Size = New System.Drawing.Size(68, 13)
        Label15.TabIndex = 336
        Label15.Text = "Department :"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.txtRemarks)
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.DataGridView1)
        Me.Panel1.Controls.Add(Me.GroupBox4)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(7, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(648, 685)
        Me.Panel1.TabIndex = 2
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(472, 404)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(55, 13)
        Me.Label16.TabIndex = 96
        Me.Label16.Text = "Remarks :"
        '
        'txtRemarks
        '
        Me.txtRemarks.Location = New System.Drawing.Point(469, 420)
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtRemarks.Size = New System.Drawing.Size(167, 136)
        Me.txtRemarks.TabIndex = 4
        Me.txtRemarks.Text = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.txtReceiverCN)
        Me.GroupBox3.Controls.Add(Me.txtReceiverAddress)
        Me.GroupBox3.Controls.Add(Me.txtReceiverName)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Label10)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 290)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(429, 105)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Receiver Info"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(26, 75)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(70, 13)
        Me.Label7.TabIndex = 338
        Me.Label7.Text = "Contact No. :"
        '
        'txtReceiverCN
        '
        Me.txtReceiverCN.Location = New System.Drawing.Point(121, 75)
        Me.txtReceiverCN.Name = "txtReceiverCN"
        Me.txtReceiverCN.Size = New System.Drawing.Size(188, 20)
        Me.txtReceiverCN.TabIndex = 2
        '
        'txtReceiverAddress
        '
        Me.txtReceiverAddress.Location = New System.Drawing.Point(121, 50)
        Me.txtReceiverAddress.Name = "txtReceiverAddress"
        Me.txtReceiverAddress.Size = New System.Drawing.Size(298, 20)
        Me.txtReceiverAddress.TabIndex = 1
        '
        'txtReceiverName
        '
        Me.txtReceiverName.Location = New System.Drawing.Point(121, 26)
        Me.txtReceiverName.Name = "txtReceiverName"
        Me.txtReceiverName.Size = New System.Drawing.Size(188, 20)
        Me.txtReceiverName.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(26, 50)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(51, 13)
        Me.Label9.TabIndex = 335
        Me.Label9.Text = "Address :"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtSenderContactNo)
        Me.GroupBox1.Controls.Add(Me.txtSenderAddress)
        Me.GroupBox1.Controls.Add(Me.txtSenderName)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 184)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(429, 105)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sender Info"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(26, 75)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 13)
        Me.Label6.TabIndex = 338
        Me.Label6.Text = "Contact No. :"
        '
        'txtSenderContactNo
        '
        Me.txtSenderContactNo.Location = New System.Drawing.Point(121, 75)
        Me.txtSenderContactNo.Name = "txtSenderContactNo"
        Me.txtSenderContactNo.Size = New System.Drawing.Size(188, 20)
        Me.txtSenderContactNo.TabIndex = 2
        '
        'txtSenderAddress
        '
        Me.txtSenderAddress.Location = New System.Drawing.Point(121, 50)
        Me.txtSenderAddress.Name = "txtSenderAddress"
        Me.txtSenderAddress.Size = New System.Drawing.Size(298, 20)
        Me.txtSenderAddress.TabIndex = 1
        '
        'txtSenderName
        '
        Me.txtSenderName.Location = New System.Drawing.Point(121, 26)
        Me.txtSenderName.Name = "txtSenderName"
        Me.txtSenderName.Size = New System.Drawing.Size(188, 20)
        Me.txtSenderName.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 50)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 335
        Me.Label3.Text = "Address :"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FloralWhite
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.LightSteelBlue
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.ColumnHeadersHeight = 24
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5})
        Me.DataGridView1.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView1.EnableHeadersVisualStyles = False
        Me.DataGridView1.GridColor = System.Drawing.Color.White
        Me.DataGridView1.Location = New System.Drawing.Point(9, 562)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.DataGridView1.RowHeadersWidth = 25
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.SteelBlue
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.Color.White
        Me.DataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridView1.RowTemplate.Height = 18
        Me.DataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(627, 115)
        Me.DataGridView1.TabIndex = 2
        '
        'Column1
        '
        Me.Column1.HeaderText = "Accession No."
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        '
        'Column2
        '
        Me.Column2.HeaderText = "Book Title"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 150
        '
        'Column3
        '
        Me.Column3.HeaderText = "Author"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 150
        '
        'Column4
        '
        Me.Column4.HeaderText = "Publisher"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        '
        'Column5
        '
        Me.Column5.HeaderText = "Qty"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cmbDepartment)
        Me.GroupBox4.Controls.Add(Label15)
        Me.GroupBox4.Controls.Add(Me.dtpOrderDate)
        Me.GroupBox4.Controls.Add(Me.txtOrderNo)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Label5)
        Me.GroupBox4.Location = New System.Drawing.Point(9, 79)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(319, 99)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Quotation Info"
        '
        'cmbDepartment
        '
        Me.cmbDepartment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDepartment.FormattingEnabled = True
        Me.cmbDepartment.Location = New System.Drawing.Point(121, 71)
        Me.cmbDepartment.Name = "cmbDepartment"
        Me.cmbDepartment.Size = New System.Drawing.Size(188, 21)
        Me.cmbDepartment.TabIndex = 2
        '
        'dtpOrderDate
        '
        Me.dtpOrderDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpOrderDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOrderDate.Location = New System.Drawing.Point(121, 47)
        Me.dtpOrderDate.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpOrderDate.Name = "dtpOrderDate"
        Me.dtpOrderDate.Size = New System.Drawing.Size(111, 20)
        Me.dtpOrderDate.TabIndex = 1
        '
        'txtOrderNo
        '
        Me.txtOrderNo.Location = New System.Drawing.Point(121, 22)
        Me.txtOrderNo.Name = "txtOrderNo"
        Me.txtOrderNo.ReadOnly = True
        Me.txtOrderNo.Size = New System.Drawing.Size(111, 20)
        Me.txtOrderNo.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 47)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 335
        Me.Label2.Text = "Order Date :"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.btnRemove)
        Me.GroupBox2.Controls.Add(Me.txtQty)
        Me.GroupBox2.Controls.Add(Me.txtAccessionNo)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.txtPublisher)
        Me.GroupBox2.Controls.Add(Me.btnAdd)
        Me.GroupBox2.Controls.Add(Me.txtAuthor)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.txtBookTitle)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(9, 399)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(454, 158)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Book Details"
        '
        'Button2
        '
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(249, 26)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(35, 21)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnRemove.Enabled = False
        Me.btnRemove.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnRemove.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemove.Location = New System.Drawing.Point(375, 26)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(71, 19)
        Me.btnRemove.TabIndex = 7
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'txtQty
        '
        Me.txtQty.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQty.Location = New System.Drawing.Point(121, 127)
        Me.txtQty.Name = "txtQty"
        Me.txtQty.Size = New System.Drawing.Size(76, 20)
        Me.txtQty.TabIndex = 4
        Me.txtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAccessionNo
        '
        Me.txtAccessionNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAccessionNo.Location = New System.Drawing.Point(121, 27)
        Me.txtAccessionNo.Name = "txtAccessionNo"
        Me.txtAccessionNo.ReadOnly = True
        Me.txtAccessionNo.Size = New System.Drawing.Size(122, 20)
        Me.txtAccessionNo.TabIndex = 0
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(26, 128)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(52, 13)
        Me.Label14.TabIndex = 99
        Me.Label14.Text = "Quantity :"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(26, 102)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 13)
        Me.Label13.TabIndex = 98
        Me.Label13.Text = "Publisher :"
        '
        'txtPublisher
        '
        Me.txtPublisher.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPublisher.Location = New System.Drawing.Point(121, 102)
        Me.txtPublisher.Name = "txtPublisher"
        Me.txtPublisher.ReadOnly = True
        Me.txtPublisher.Size = New System.Drawing.Size(122, 20)
        Me.txtPublisher.TabIndex = 3
        '
        'btnAdd
        '
        Me.btnAdd.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.Location = New System.Drawing.Point(290, 26)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(79, 19)
        Me.btnAdd.TabIndex = 6
        Me.btnAdd.Text = "&Add To Grid"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'txtAuthor
        '
        Me.txtAuthor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAuthor.Location = New System.Drawing.Point(121, 76)
        Me.txtAuthor.Name = "txtAuthor"
        Me.txtAuthor.ReadOnly = True
        Me.txtAuthor.Size = New System.Drawing.Size(289, 20)
        Me.txtAuthor.TabIndex = 2
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(26, 27)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(82, 13)
        Me.Label8.TabIndex = 95
        Me.Label8.Text = "Accession No. :"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(26, 77)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(44, 13)
        Me.Label11.TabIndex = 89
        Me.Label11.Text = "Author :"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(26, 51)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 13)
        Me.Label12.TabIndex = 87
        Me.Label12.Text = "Book Title :"
        '
        'txtBookTitle
        '
        Me.txtBookTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBookTitle.Location = New System.Drawing.Point(121, 51)
        Me.txtBookTitle.Name = "txtBookTitle"
        Me.txtBookTitle.ReadOnly = True
        Me.txtBookTitle.Size = New System.Drawing.Size(289, 20)
        Me.txtBookTitle.TabIndex = 1
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.btnPrint)
        Me.Panel4.Controls.Add(Me.btnGetData)
        Me.Panel4.Controls.Add(Me.btnDelete)
        Me.Panel4.Controls.Add(Me.btnClose)
        Me.Panel4.Controls.Add(Me.btnUpdate)
        Me.Panel4.Controls.Add(Me.btnSave)
        Me.Panel4.Controls.Add(Me.btnNew)
        Me.Panel4.Location = New System.Drawing.Point(458, 80)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(111, 232)
        Me.Panel4.TabIndex = 3
        '
        'btnPrint
        '
        Me.btnPrint.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPrint.Enabled = False
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(13, 164)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(82, 28)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnGetData
        '
        Me.btnGetData.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnGetData.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnGetData.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetData.Location = New System.Drawing.Point(13, 133)
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
        Me.btnClose.Location = New System.Drawing.Point(13, 194)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(82, 28)
        Me.btnClose.TabIndex = 6
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
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(13, 41)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(82, 28)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save + Print"
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
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.txtDepartmentID)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtOrderID)
        Me.Panel2.Controls.Add(Me.lblUser)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(627, 62)
        Me.Panel2.TabIndex = 0
        '
        'txtDepartmentID
        '
        Me.txtDepartmentID.Location = New System.Drawing.Point(388, 18)
        Me.txtDepartmentID.Name = "txtDepartmentID"
        Me.txtDepartmentID.Size = New System.Drawing.Size(111, 20)
        Me.txtDepartmentID.TabIndex = 6
        Me.txtDepartmentID.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(229, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Quotation"
        '
        'txtOrderID
        '
        Me.txtOrderID.Location = New System.Drawing.Point(68, 23)
        Me.txtOrderID.Name = "txtOrderID"
        Me.txtOrderID.Size = New System.Drawing.Size(111, 20)
        Me.txtOrderID.TabIndex = 4
        Me.txtOrderID.Visible = False
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Location = New System.Drawing.Point(193, 18)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(39, 13)
        Me.lblUser.TabIndex = 5
        Me.lblUser.Text = "Label8"
        Me.lblUser.Visible = False
        '
        'Timer1
        '
        '
        'frmQuotation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(663, 701)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmQuotation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtOrderID As System.Windows.Forms.TextBox
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnGetData As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpOrderDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtOrderNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtReceiverCN As System.Windows.Forms.TextBox
    Friend WithEvents txtReceiverAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtReceiverName As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSenderContactNo As System.Windows.Forms.TextBox
    Friend WithEvents txtSenderAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtSenderName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtQty As System.Windows.Forms.TextBox
    Friend WithEvents txtAccessionNo As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtPublisher As System.Windows.Forms.TextBox
    Friend WithEvents txtAuthor As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtBookTitle As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmbDepartment As System.Windows.Forms.ComboBox
    Friend WithEvents txtDepartmentID As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtRemarks As System.Windows.Forms.RichTextBox

End Class
