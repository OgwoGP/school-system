<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBookEntry
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBookEntry))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.txtClassificationID = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtRemarks = New System.Windows.Forms.RichTextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.txtSupplierID = New System.Windows.Forms.TextBox()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.dtpBillDate = New System.Windows.Forms.DateTimePicker()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.txtBookPosition = New System.Windows.Forms.TextBox()
        Me.txtNoOfPages = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cmbCondition = New System.Windows.Forms.ComboBox()
        Me.rbReference = New System.Windows.Forms.RadioButton()
        Me.rbNormal = New System.Windows.Forms.RadioButton()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.cmbSubCategory = New System.Windows.Forms.ComboBox()
        Me.cmbCategory = New System.Windows.Forms.ComboBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.cmbClass = New System.Windows.Forms.ComboBox()
        Me.txtBarcode = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.dtpEntryDate = New System.Windows.Forms.DateTimePicker()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtVolume = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtLanguage = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtBillNo = New System.Windows.Forms.TextBox()
        Me.txtPublishingYear = New System.Windows.Forms.TextBox()
        Me.txtPlaceOfPublisher = New System.Windows.Forms.TextBox()
        Me.txtPublisherName = New System.Windows.Forms.TextBox()
        Me.txtEdition = New System.Windows.Forms.TextBox()
        Me.txtISBN = New System.Windows.Forms.TextBox()
        Me.txtJointAuthor = New System.Windows.Forms.TextBox()
        Me.txtAuthor = New System.Windows.Forms.TextBox()
        Me.txtBookTitle = New System.Windows.Forms.TextBox()
        Me.txtAccessionNo = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtAccessionNo1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtID)
        Me.Panel1.Controls.Add(Me.txtClassificationID)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Location = New System.Drawing.Point(7, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(931, 666)
        Me.Panel1.TabIndex = 2
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(805, 404)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(95, 20)
        Me.txtID.TabIndex = 3
        Me.txtID.Visible = False
        '
        'txtClassificationID
        '
        Me.txtClassificationID.Location = New System.Drawing.Point(805, 374)
        Me.txtClassificationID.Name = "txtClassificationID"
        Me.txtClassificationID.Size = New System.Drawing.Size(95, 20)
        Me.txtClassificationID.TabIndex = 2
        Me.txtClassificationID.Visible = False
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
        Me.Panel3.Location = New System.Drawing.Point(808, 54)
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
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.lblUser)
        Me.GroupBox1.Controls.Add(Me.dtpBillDate)
        Me.GroupBox1.Controls.Add(Me.txtPrice)
        Me.GroupBox1.Controls.Add(Me.txtBookPosition)
        Me.GroupBox1.Controls.Add(Me.txtNoOfPages)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.cmbCondition)
        Me.GroupBox1.Controls.Add(Me.rbReference)
        Me.GroupBox1.Controls.Add(Me.rbNormal)
        Me.GroupBox1.Controls.Add(Me.Label25)
        Me.GroupBox1.Controls.Add(Me.Label24)
        Me.GroupBox1.Controls.Add(Me.cmbSubCategory)
        Me.GroupBox1.Controls.Add(Me.cmbCategory)
        Me.GroupBox1.Controls.Add(Me.Label22)
        Me.GroupBox1.Controls.Add(Me.cmbClass)
        Me.GroupBox1.Controls.Add(Me.txtBarcode)
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Controls.Add(Me.dtpEntryDate)
        Me.GroupBox1.Controls.Add(Me.Label20)
        Me.GroupBox1.Controls.Add(Me.txtVolume)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.txtLanguage)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtBillNo)
        Me.GroupBox1.Controls.Add(Me.txtPublishingYear)
        Me.GroupBox1.Controls.Add(Me.txtPlaceOfPublisher)
        Me.GroupBox1.Controls.Add(Me.txtPublisherName)
        Me.GroupBox1.Controls.Add(Me.txtEdition)
        Me.GroupBox1.Controls.Add(Me.txtISBN)
        Me.GroupBox1.Controls.Add(Me.txtJointAuthor)
        Me.GroupBox1.Controls.Add(Me.txtAuthor)
        Me.GroupBox1.Controls.Add(Me.txtBookTitle)
        Me.GroupBox1.Controls.Add(Me.txtAccessionNo)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label23)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 47)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(790, 612)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Book Info"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtRemarks)
        Me.GroupBox3.Location = New System.Drawing.Point(302, 530)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(477, 74)
        Me.GroupBox3.TabIndex = 25
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Remarks"
        '
        'txtRemarks
        '
        Me.txtRemarks.Location = New System.Drawing.Point(19, 21)
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtRemarks.Size = New System.Drawing.Size(451, 44)
        Me.txtRemarks.TabIndex = 0
        Me.txtRemarks.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.txtLastName)
        Me.GroupBox2.Controls.Add(Me.txtFirstName)
        Me.GroupBox2.Controls.Add(Me.Label26)
        Me.GroupBox2.Controls.Add(Me.Label27)
        Me.GroupBox2.Controls.Add(Me.Label28)
        Me.GroupBox2.Controls.Add(Me.txtSupplierID)
        Me.GroupBox2.Location = New System.Drawing.Point(302, 429)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(477, 100)
        Me.GroupBox2.TabIndex = 24
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
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(301, 46)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(73, 15)
        Me.Label16.TabIndex = 21
        Me.Label16.Text = "First Name :"
        '
        'txtLastName
        '
        Me.txtLastName.BackColor = System.Drawing.SystemColors.Control
        Me.txtLastName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastName.Location = New System.Drawing.Point(141, 64)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.ReadOnly = True
        Me.txtLastName.Size = New System.Drawing.Size(157, 21)
        Me.txtLastName.TabIndex = 2
        '
        'txtFirstName
        '
        Me.txtFirstName.BackColor = System.Drawing.SystemColors.Control
        Me.txtFirstName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstName.Location = New System.Drawing.Point(304, 64)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.ReadOnly = True
        Me.txtFirstName.Size = New System.Drawing.Size(166, 21)
        Me.txtFirstName.TabIndex = 3
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(138, 46)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(73, 15)
        Me.Label26.TabIndex = 20
        Me.Label26.Text = "Last Name :"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(23, 46)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(96, 15)
        Me.Label27.TabIndex = 19
        Me.Label27.Text = "Supplier Name :"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(23, 23)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(74, 15)
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
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Location = New System.Drawing.Point(500, 229)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(45, 15)
        Me.lblUser.TabIndex = 5
        Me.lblUser.Text = "Label8"
        Me.lblUser.Visible = False
        '
        'dtpBillDate
        '
        Me.dtpBillDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpBillDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpBillDate.Location = New System.Drawing.Point(186, 554)
        Me.dtpBillDate.Name = "dtpBillDate"
        Me.dtpBillDate.Size = New System.Drawing.Size(110, 21)
        Me.dtpBillDate.TabIndex = 22
        '
        'txtPrice
        '
        Me.txtPrice.Location = New System.Drawing.Point(186, 580)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.Size = New System.Drawing.Size(110, 21)
        Me.txtPrice.TabIndex = 23
        '
        'txtBookPosition
        '
        Me.txtBookPosition.Location = New System.Drawing.Point(186, 477)
        Me.txtBookPosition.Name = "txtBookPosition"
        Me.txtBookPosition.Size = New System.Drawing.Size(110, 21)
        Me.txtBookPosition.TabIndex = 19
        '
        'txtNoOfPages
        '
        Me.txtNoOfPages.Location = New System.Drawing.Point(186, 453)
        Me.txtNoOfPages.Name = "txtNoOfPages"
        Me.txtNoOfPages.Size = New System.Drawing.Size(110, 21)
        Me.txtNoOfPages.TabIndex = 18
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Location = New System.Drawing.Point(30, 477)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(88, 15)
        Me.Label15.TabIndex = 82
        Me.Label15.Text = "Book Position :"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Location = New System.Drawing.Point(30, 453)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(85, 15)
        Me.Label13.TabIndex = 81
        Me.Label13.Text = "No. Of Pages :"
        '
        'cmbCondition
        '
        Me.cmbCondition.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbCondition.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbCondition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCondition.FormattingEnabled = True
        Me.cmbCondition.Items.AddRange(New Object() {"Fresh", "Damaged"})
        Me.cmbCondition.Location = New System.Drawing.Point(186, 501)
        Me.cmbCondition.Name = "cmbCondition"
        Me.cmbCondition.Size = New System.Drawing.Size(110, 23)
        Me.cmbCondition.TabIndex = 20
        '
        'rbReference
        '
        Me.rbReference.AutoSize = True
        Me.rbReference.Location = New System.Drawing.Point(258, 406)
        Me.rbReference.Name = "rbReference"
        Me.rbReference.Size = New System.Drawing.Size(82, 19)
        Me.rbReference.TabIndex = 16
        Me.rbReference.TabStop = True
        Me.rbReference.Text = "Reference"
        Me.rbReference.UseVisualStyleBackColor = True
        '
        'rbNormal
        '
        Me.rbNormal.AutoSize = True
        Me.rbNormal.Location = New System.Drawing.Point(186, 406)
        Me.rbNormal.Name = "rbNormal"
        Me.rbNormal.Size = New System.Drawing.Size(66, 19)
        Me.rbNormal.TabIndex = 15
        Me.rbNormal.TabStop = True
        Me.rbNormal.Text = "Normal"
        Me.rbNormal.UseVisualStyleBackColor = True
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.BackColor = System.Drawing.Color.Transparent
        Me.Label25.Location = New System.Drawing.Point(30, 173)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(61, 15)
        Me.Label25.TabIndex = 77
        Me.Label25.Text = "Category :"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.BackColor = System.Drawing.Color.Transparent
        Me.Label24.Location = New System.Drawing.Point(30, 72)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(69, 15)
        Me.Label24.TabIndex = 76
        Me.Label24.Text = "Entry Date :"
        '
        'cmbSubCategory
        '
        Me.cmbSubCategory.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbSubCategory.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbSubCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSubCategory.Enabled = False
        Me.cmbSubCategory.FormattingEnabled = True
        Me.cmbSubCategory.Location = New System.Drawing.Point(186, 200)
        Me.cmbSubCategory.Name = "cmbSubCategory"
        Me.cmbSubCategory.Size = New System.Drawing.Size(253, 23)
        Me.cmbSubCategory.TabIndex = 7
        '
        'cmbCategory
        '
        Me.cmbCategory.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbCategory.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCategory.Enabled = False
        Me.cmbCategory.FormattingEnabled = True
        Me.cmbCategory.Location = New System.Drawing.Point(186, 173)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(253, 23)
        Me.cmbCategory.TabIndex = 6
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.BackColor = System.Drawing.Color.Transparent
        Me.Label22.Location = New System.Drawing.Point(30, 501)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(65, 15)
        Me.Label22.TabIndex = 73
        Me.Label22.Text = "Condition :"
        '
        'cmbClass
        '
        Me.cmbClass.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbClass.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbClass.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbClass.FormattingEnabled = True
        Me.cmbClass.Location = New System.Drawing.Point(186, 146)
        Me.cmbClass.Name = "cmbClass"
        Me.cmbClass.Size = New System.Drawing.Size(162, 23)
        Me.cmbClass.TabIndex = 5
        '
        'txtBarcode
        '
        Me.txtBarcode.Location = New System.Drawing.Point(186, 253)
        Me.txtBarcode.Name = "txtBarcode"
        Me.txtBarcode.Size = New System.Drawing.Size(156, 21)
        Me.txtBarcode.TabIndex = 9
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.BackColor = System.Drawing.Color.Transparent
        Me.Label21.Location = New System.Drawing.Point(30, 253)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(59, 15)
        Me.Label21.TabIndex = 71
        Me.Label21.Text = "Barcode :"
        '
        'dtpEntryDate
        '
        Me.dtpEntryDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpEntryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEntryDate.Location = New System.Drawing.Point(186, 72)
        Me.dtpEntryDate.Name = "dtpEntryDate"
        Me.dtpEntryDate.Size = New System.Drawing.Size(135, 21)
        Me.dtpEntryDate.TabIndex = 2
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.Color.Transparent
        Me.Label20.Location = New System.Drawing.Point(30, 554)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(59, 15)
        Me.Label20.TabIndex = 69
        Me.Label20.Text = "Bill Date :"
        '
        'txtVolume
        '
        Me.txtVolume.Location = New System.Drawing.Point(186, 304)
        Me.txtVolume.Name = "txtVolume"
        Me.txtVolume.Size = New System.Drawing.Size(110, 21)
        Me.txtVolume.TabIndex = 11
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Location = New System.Drawing.Point(30, 304)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(55, 15)
        Me.Label19.TabIndex = 65
        Me.Label19.Text = "Volume :"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Location = New System.Drawing.Point(30, 429)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(69, 15)
        Me.Label18.TabIndex = 63
        Me.Label18.Text = "Language :"
        '
        'txtLanguage
        '
        Me.txtLanguage.Location = New System.Drawing.Point(186, 429)
        Me.txtLanguage.Name = "txtLanguage"
        Me.txtLanguage.Size = New System.Drawing.Size(110, 21)
        Me.txtLanguage.TabIndex = 17
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Location = New System.Drawing.Point(30, 146)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(43, 15)
        Me.Label12.TabIndex = 58
        Me.Label12.Text = "Class :"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Location = New System.Drawing.Point(30, 580)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(41, 15)
        Me.Label11.TabIndex = 55
        Me.Label11.Text = "Price :"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Location = New System.Drawing.Point(30, 200)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(86, 15)
        Me.Label7.TabIndex = 54
        Me.Label7.Text = "Sub Category :"
        '
        'txtBillNo
        '
        Me.txtBillNo.Location = New System.Drawing.Point(186, 528)
        Me.txtBillNo.Name = "txtBillNo"
        Me.txtBillNo.Size = New System.Drawing.Size(110, 21)
        Me.txtBillNo.TabIndex = 21
        '
        'txtPublishingYear
        '
        Me.txtPublishingYear.Location = New System.Drawing.Point(186, 382)
        Me.txtPublishingYear.Name = "txtPublishingYear"
        Me.txtPublishingYear.Size = New System.Drawing.Size(110, 21)
        Me.txtPublishingYear.TabIndex = 14
        '
        'txtPlaceOfPublisher
        '
        Me.txtPlaceOfPublisher.Location = New System.Drawing.Point(186, 356)
        Me.txtPlaceOfPublisher.Name = "txtPlaceOfPublisher"
        Me.txtPlaceOfPublisher.Size = New System.Drawing.Size(253, 21)
        Me.txtPlaceOfPublisher.TabIndex = 13
        '
        'txtPublisherName
        '
        Me.txtPublisherName.Location = New System.Drawing.Point(186, 330)
        Me.txtPublisherName.Name = "txtPublisherName"
        Me.txtPublisherName.Size = New System.Drawing.Size(253, 21)
        Me.txtPublisherName.TabIndex = 12
        '
        'txtEdition
        '
        Me.txtEdition.Location = New System.Drawing.Point(186, 278)
        Me.txtEdition.Name = "txtEdition"
        Me.txtEdition.Size = New System.Drawing.Size(110, 21)
        Me.txtEdition.TabIndex = 10
        '
        'txtISBN
        '
        Me.txtISBN.Location = New System.Drawing.Point(186, 227)
        Me.txtISBN.Name = "txtISBN"
        Me.txtISBN.Size = New System.Drawing.Size(156, 21)
        Me.txtISBN.TabIndex = 8
        '
        'txtJointAuthor
        '
        Me.txtJointAuthor.Location = New System.Drawing.Point(186, 121)
        Me.txtJointAuthor.Name = "txtJointAuthor"
        Me.txtJointAuthor.Size = New System.Drawing.Size(253, 21)
        Me.txtJointAuthor.TabIndex = 4
        '
        'txtAuthor
        '
        Me.txtAuthor.Location = New System.Drawing.Point(186, 97)
        Me.txtAuthor.Name = "txtAuthor"
        Me.txtAuthor.Size = New System.Drawing.Size(253, 21)
        Me.txtAuthor.TabIndex = 3
        '
        'txtBookTitle
        '
        Me.txtBookTitle.Location = New System.Drawing.Point(186, 47)
        Me.txtBookTitle.Name = "txtBookTitle"
        Me.txtBookTitle.Size = New System.Drawing.Size(253, 21)
        Me.txtBookTitle.TabIndex = 1
        '
        'txtAccessionNo
        '
        Me.txtAccessionNo.Location = New System.Drawing.Point(186, 23)
        Me.txtAccessionNo.Name = "txtAccessionNo"
        Me.txtAccessionNo.Size = New System.Drawing.Size(95, 21)
        Me.txtAccessionNo.TabIndex = 0
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.BackColor = System.Drawing.Color.Transparent
        Me.Label17.Location = New System.Drawing.Point(30, 227)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(41, 15)
        Me.Label17.TabIndex = 25
        Me.Label17.Text = "ISBN :"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Location = New System.Drawing.Point(30, 527)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(52, 15)
        Me.Label14.TabIndex = 22
        Me.Label14.Text = "Bill No. :"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Location = New System.Drawing.Point(30, 382)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(99, 15)
        Me.Label10.TabIndex = 18
        Me.Label10.Text = "Publishing Year :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Location = New System.Drawing.Point(30, 356)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(114, 15)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "Place Of Publisher :"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Location = New System.Drawing.Point(30, 330)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(65, 15)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Publisher :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Location = New System.Drawing.Point(30, 278)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 15)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Edition :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Location = New System.Drawing.Point(30, 121)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 15)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Joint Authors :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Location = New System.Drawing.Point(30, 97)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 15)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Author(s) :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Location = New System.Drawing.Point(30, 406)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 15)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Section :"
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
        Me.Panel2.Controls.Add(Me.txtAccessionNo1)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Location = New System.Drawing.Point(9, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(910, 40)
        Me.Panel2.TabIndex = 0
        '
        'txtAccessionNo1
        '
        Me.txtAccessionNo1.Location = New System.Drawing.Point(565, 19)
        Me.txtAccessionNo1.Name = "txtAccessionNo1"
        Me.txtAccessionNo1.Size = New System.Drawing.Size(95, 20)
        Me.txtAccessionNo1.TabIndex = 1
        Me.txtAccessionNo1.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(376, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Book Entry"
        '
        'frmBookEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(945, 678)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBookEntry"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
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
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents cmbClass As System.Windows.Forms.ComboBox
    Friend WithEvents txtBarcode As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents dtpEntryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtVolume As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtLanguage As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtBillNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPublishingYear As System.Windows.Forms.TextBox
    Friend WithEvents txtPlaceOfPublisher As System.Windows.Forms.TextBox
    Friend WithEvents txtPublisherName As System.Windows.Forms.TextBox
    Friend WithEvents txtEdition As System.Windows.Forms.TextBox
    Friend WithEvents txtISBN As System.Windows.Forms.TextBox
    Friend WithEvents txtJointAuthor As System.Windows.Forms.TextBox
    Friend WithEvents txtAuthor As System.Windows.Forms.TextBox
    Friend WithEvents txtBookTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtAccessionNo As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents dtpBillDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtBookPosition As System.Windows.Forms.TextBox
    Friend WithEvents txtNoOfPages As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cmbCondition As System.Windows.Forms.ComboBox
    Friend WithEvents rbReference As System.Windows.Forms.RadioButton
    Friend WithEvents rbNormal As System.Windows.Forms.RadioButton
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents cmbSubCategory As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCategory As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents txtSupplierID As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtRemarks As System.Windows.Forms.RichTextBox
    Public WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents txtClassificationID As System.Windows.Forms.TextBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents btnGetData As System.Windows.Forms.Button
    Friend WithEvents txtAccessionNo1 As System.Windows.Forms.TextBox

End Class
