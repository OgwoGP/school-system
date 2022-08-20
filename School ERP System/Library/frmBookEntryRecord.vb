Imports System.Data.SqlClient

Imports System.IO

Public Class frmBookEntryRecord

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetData()
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub
    Sub Reset()
        txtAccessionNo.Text = ""
        txtAuthor.Text = ""
        txtBookTitle.Text = ""
        txtCategory.Text = ""
        txtClass.Text = ""
        txtLanguage.Text = ""
        txtPublisher.Text = ""
        txtSubCategory.Text = ""
        cmbCondition.SelectedIndex = -1
        cmbStatus.SelectedIndex = -1
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
        Getdata()
    End Sub
    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub dgw_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Book Entry" Then
                    Me.Hide()
                    frmBookEntry.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    frmBookEntry.txtAccessionNo.Text = dr.Cells(0).Value.ToString()
                    frmBookEntry.txtAccessionNo1.Text = dr.Cells(0).Value.ToString()
                    frmBookEntry.txtBookTitle.Text = dr.Cells(1).Value.ToString()
                    frmBookEntry.dtpEntryDate.Text = dr.Cells(2).Value.ToString()
                    frmBookEntry.txtAuthor.Text = dr.Cells(3).Value.ToString()
                    frmBookEntry.txtJointAuthor.Text = dr.Cells(4).Value.ToString()
                    frmBookEntry.txtClassificationID.Text = dr.Cells(5).Value.ToString()
                    frmBookEntry.cmbClass.Text = dr.Cells(6).Value.ToString()
                    frmBookEntry.cmbCategory.Text = dr.Cells(7).Value.ToString()
                    frmBookEntry.cmbSubCategory.Text = dr.Cells(8).Value.ToString()
                    frmBookEntry.txtBarcode.Text = dr.Cells(9).Value.ToString()
                    frmBookEntry.txtISBN.Text = dr.Cells(10).Value.ToString()
                    frmBookEntry.txtVolume.Text = dr.Cells(11).Value.ToString()
                    frmBookEntry.txtEdition.Text = dr.Cells(12).Value.ToString()
                    frmBookEntry.txtPublisherName.Text = dr.Cells(13).Value.ToString()
                    frmBookEntry.txtPlaceOfPublisher.Text = dr.Cells(14).Value.ToString()
                    frmBookEntry.txtPublishingYear.Text = dr.Cells(15).Value.ToString()
                    If dr.Cells(16).Value = "Normal" Then
                        frmBookEntry.rbNormal.Checked = True
                    End If
                    If dr.Cells(16).Value = "Reference" Then
                        frmBookEntry.rbReference.Checked = True
                    End If
                    frmBookEntry.txtLanguage.Text = dr.Cells(17).Value.ToString()
                    frmBookEntry.txtBookPosition.Text = dr.Cells(18).Value.ToString()
                    frmBookEntry.txtPrice.Text = dr.Cells(19).Value.ToString()
                    frmBookEntry.txtID.Text = dr.Cells(20).Value.ToString()
                    frmBookEntry.txtSupplierID.Text = dr.Cells(21).Value.ToString()
                    frmBookEntry.txtLastName.Text = dr.Cells(22).Value.ToString()
                    frmBookEntry.txtFirstName.Text = dr.Cells(23).Value.ToString()
                    frmBookEntry.txtBillNo.Text = dr.Cells(24).Value.ToString()
                    frmBookEntry.dtpBillDate.Text = dr.Cells(25).Value.ToString()
                    frmBookEntry.txtNoOfPages.Text = dr.Cells(26).Value.ToString()
                    frmBookEntry.cmbCondition.Text = dr.Cells(27).Value.ToString()
                    frmBookEntry.txtRemarks.Text = dr.Cells(29).Value.ToString()
                    frmBookEntry.btnSave.Enabled = False
                    frmBookEntry.btnUpdate.Enabled = True
                    frmBookEntry.btnDelete.Enabled = True
                    lblSet.Text = ""
                End If
                If lblSet.Text = "Quotation" Then
                    Me.Hide()
                    frmQuotation.Show()
                    ' or simply use column name instead of index
                    'dr.Cells["id"].Value.ToString();
                    frmQuotation.txtAccessionNo.Text = dr.Cells(0).Value.ToString()
                    frmQuotation.txtBookTitle.Text = dr.Cells(1).Value.ToString()
                    frmQuotation.txtAuthor.Text = dr.Cells(3).Value.ToString()
                    frmQuotation.txtPublisher.Text = dr.Cells(13).Value.ToString()
                    lblSet.Text = ""
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub dgw_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub txtBookTitle_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtBookTitle.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and BookTitle like '%" & txtBookTitle.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtAuthor_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtAuthor.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and Author like '%" & txtAuthor.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtLanguage_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtLanguage.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and [Language] like '%" & txtLanguage.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtPublisher_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtPublisher.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and Publisher like '%" & txtPublisher.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtClass_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtClass.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and Classname like '%" & txtClass.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtCategory_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtCategory.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and CategoryName like '%" & txtCategory.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtSubCategory_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSubCategory.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and SubCategoryName like '%" & txtSubCategory.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtAccessionNo_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtAccessionNo.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and AccessionNo like '%" & txtAccessionNo.Text & "%' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbCondition_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbCondition.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and Condition= '" & cmbCondition.Text & "' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbStatus.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID  and Status= '" & cmbStatus.Text & "' order by AccessionNo", con)
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dtpDateTo_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles dtpDateTo.Validating
        If (dtpDateFrom.Value.Date) > (dtpDateTo.Value.Date) Then
            MessageBox.Show("Invalid Selection", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dtpDateTo.Focus()
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select RTRIM(AccessionNo) as [Accession No], RTRIM(BookTitle) as [Book Title], EntryDate as [Entry Date], RTRIM(Author) as [Author], RTRIM(JointAuthors) as [Joint Authors], SubCategoryID as [Classification ID],RTRIM(ClassName) as [Class],RTRIM(CategoryName) as [Category],RTRIM(SubCategoryName) as [Sub Category], RTRIM(Barcode) as [Barcode], RTRIM(ISBN) as [ISBN], RTRIM(Volume) as [Volume], RTRIM(Edition) as [Edition], RTRIM(Publisher) as [Publisher], RTRIM(PlaceOfPublisher) as [Publisher Place], RTRIM(PublishingYear) as [Publishing Year], RTRIM(Section) as [Section], RTRIM(Language) as [Language], RTRIM(BookPosition) as [Book Position], Price,Supplier.ID as [SID], RTRIM(Supplier.SupplierID) as [Supplier ID],RTRIM(LastName) as [Last Name],RTRIM(FirstName) as [First Name],RTRIM(BillNo) as [Bill No], BillDate as [Bill Date], RTRIM(NoOfPages) as [No of Pages], RTRIM(Condition) as [Condition], RTRIM(Status) as [Status], RTRIM(Book.Remarks) as [Remarks] from Book,Supplier,BookClass,Category,SubCategory where Book.SupplierID=Supplier.ID and Book.SubCategoryID=SubCategory.ID and BookClass.Classname=Category.Class and SubCategory.CategoryID=Category.ID and BillDate between @d1 and @d2 order by AccessionNo,BillDate", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            ds = New DataSet()
            adp.Fill(ds, "Book")
            dgw.DataSource = ds.Tables("Book").DefaultView
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
