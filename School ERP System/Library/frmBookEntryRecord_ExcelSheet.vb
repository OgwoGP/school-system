Imports System.Data.SqlClient

Imports System.IO

Public Class frmBookEntryRecord_ExcelSheet

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

    Private Sub dgw_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

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

    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        ExportExcel(dgw)
    End Sub

    Private Sub btnImportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnImportExcel.Click
        Try
            Dim OpenFileDialog As New OpenFileDialog
            OpenFileDialog.Filter = "Excel Files | *.xlsx; *.xls;| All Files (*.*)| *.*"
            If OpenFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK AndAlso OpenFileDialog.FileName <> "" Then
                Cursor = Cursors.WaitCursor
                Timer1.Enabled = True
                Dim Pathname As String = OpenFileDialog.FileName
                Dim MyConnection As System.Data.OleDb.OleDbConnection
                Dim DtSet As System.Data.DataSet
                Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Pathname + ";Extended Properties=Excel 8.0;")
                MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
                MyConnection.Open()
                DtSet = New System.Data.DataSet
                MyCommand.Fill(DtSet)
                dgw.Visible = True
                dgw.DataSource = DtSet.Tables(0)
                btnUpdate.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        Try
            If dgw.RowCount = Nothing Then
                MessageBox.Show("Sorry nothing to save.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    SqlConnection.ClearAllPools()
                    con = New SqlConnection(cs)
                    con.Open()
                    Dim ct As String = "select AccessionNo from Book Where AccessionNo=@d1"
                    cmd = New SqlCommand(ct)
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Connection = con
                    rdr = cmd.ExecuteReader()
                    If Not rdr.Read() Then
                        Cursor = Cursors.WaitCursor
                        Timer1.Enabled = True
                        SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb As String = "Insert Into Book(AccessionNo,BookTitle, EntryDate,Author,JointAuthors, SubCategoryID,Barcode,ISBN,Volume,Edition,Publisher,PlaceOfPublisher,PublishingYear,Section,Language,BookPosition, Price,SupplierID,BillNo, BillDate,NoOfPages,Condition,Status,Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17,@d18,@d19,@d20,@d21,@d22,'Available',@d23)"
                        cmd = New SqlCommand(cb)
                        cmd.Connection = con
                        cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value.ToString())
                        cmd.Parameters.AddWithValue("@d2", row.Cells(1).Value.ToString())
                        cmd.Parameters.AddWithValue("@d3", row.Cells(2).Value.ToString())
                        cmd.Parameters.AddWithValue("@d4", row.Cells(3).Value.ToString())
                        cmd.Parameters.AddWithValue("@d5", row.Cells(4).Value.ToString())
                        cmd.Parameters.AddWithValue("@d6", Val(row.Cells(5).Value))
                        cmd.Parameters.AddWithValue("@d7", row.Cells(9).Value.ToString())
                        cmd.Parameters.AddWithValue("@d8", row.Cells(10).Value.ToString())
                        cmd.Parameters.AddWithValue("@d9", row.Cells(11).Value.ToString())
                        cmd.Parameters.AddWithValue("@d10", row.Cells(12).Value.ToString())
                        cmd.Parameters.AddWithValue("@d11", row.Cells(13).Value.ToString())
                        cmd.Parameters.AddWithValue("@d12", row.Cells(14).Value.ToString())
                        cmd.Parameters.AddWithValue("@d13", row.Cells(15).Value.ToString())
                        cmd.Parameters.AddWithValue("@d14", row.Cells(16).Value.ToString())
                        cmd.Parameters.AddWithValue("@d15", row.Cells(17).Value.ToString())
                        cmd.Parameters.AddWithValue("@d16", row.Cells(18).Value.ToString())
                        cmd.Parameters.AddWithValue("@d17", Val(row.Cells(19).Value))
                        cmd.Parameters.AddWithValue("@d18", Val(row.Cells(20).Value))
                        cmd.Parameters.AddWithValue("@d19", row.Cells(24).Value.ToString())
                        cmd.Parameters.AddWithValue("@d20", row.Cells(25).Value.ToString())
                        cmd.Parameters.AddWithValue("@d21", Val(row.Cells(26).Value))
                        cmd.Parameters.AddWithValue("@d22", row.Cells(27).Value.ToString())
                        cmd.Parameters.AddWithValue("@d23", row.Cells(29).Value.ToString())
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                End If
            Next
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        Try
            If dgw.RowCount = Nothing Then
                MessageBox.Show("Sorry nothing to update.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                        Cursor = Cursors.WaitCursor
                    Timer1.Enabled = True
                    SqlConnection.ClearAllPools()
                        con = New SqlConnection(cs)
                        con.Open()
                    Dim cb As String = "Update Book Set BookTitle=@d2, EntryDate=@d3,Author=@d4,JointAuthors=@d5, SubCategoryID=@d6,Barcode=@d7,ISBN=@d8,Volume=@d9,Edition=@d10,Publisher=@d11,PlaceOfPublisher=@d12,PublishingYear=@d13,Section=@d14,Language=@d15,BookPosition=@d16, Price=@d17,SupplierID=@d18,BillNo=@d19, BillDate=@d20,NoOfPages=@d21,Condition=@d22,Remarks=@d23 where AccessionNo=@d1"
                    cmd = New SqlCommand(cb)
                        cmd.Connection = con
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value.ToString())
                    cmd.Parameters.AddWithValue("@d2", row.Cells(1).Value.ToString())
                    cmd.Parameters.AddWithValue("@d3", row.Cells(2).Value.ToString())
                    cmd.Parameters.AddWithValue("@d4", row.Cells(3).Value.ToString())
                    cmd.Parameters.AddWithValue("@d5", row.Cells(4).Value.ToString())
                        cmd.Parameters.AddWithValue("@d6", Val(row.Cells(5).Value))
                    cmd.Parameters.AddWithValue("@d7", row.Cells(9).Value.ToString())
                    cmd.Parameters.AddWithValue("@d8", row.Cells(10).Value.ToString())
                    cmd.Parameters.AddWithValue("@d9", row.Cells(11).Value.ToString())
                    cmd.Parameters.AddWithValue("@d10", row.Cells(12).Value.ToString())
                    cmd.Parameters.AddWithValue("@d11", row.Cells(13).Value.ToString())
                    cmd.Parameters.AddWithValue("@d12", row.Cells(14).Value.ToString())
                    cmd.Parameters.AddWithValue("@d13", row.Cells(15).Value.ToString())
                    cmd.Parameters.AddWithValue("@d14", row.Cells(16).Value.ToString())
                    cmd.Parameters.AddWithValue("@d15", row.Cells(17).Value.ToString())
                    cmd.Parameters.AddWithValue("@d16", row.Cells(18).Value.ToString())
                        cmd.Parameters.AddWithValue("@d17", Val(row.Cells(19).Value))
                        cmd.Parameters.AddWithValue("@d18", Val(row.Cells(20).Value))
                    cmd.Parameters.AddWithValue("@d19", row.Cells(24).Value.ToString())
                    cmd.Parameters.AddWithValue("@d20", row.Cells(25).Value.ToString())
                        cmd.Parameters.AddWithValue("@d21", Val(row.Cells(26).Value))
                    cmd.Parameters.AddWithValue("@d22", row.Cells(27).Value.ToString())
                    cmd.Parameters.AddWithValue("@d23", row.Cells(29).Value.ToString())
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
            Next
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Reset()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
