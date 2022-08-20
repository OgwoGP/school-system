Imports System.Data.SqlClient
Public Class frmBookEntry
    Dim s As String
    Sub Reset()
        txtAccessionNo.Text = ""
        txtAuthor.Text = ""
        txtBarcode.Text = ""
        txtBillNo.Text = ""
        txtBookPosition.Text = ""
        txtBookTitle.Text = ""
        txtEdition.Text = ""
        txtFirstName.Text = ""
        txtISBN.Text = ""
        txtJointAuthor.Text = ""
        txtLanguage.Text = ""
        txtLastName.Text = ""
        txtNoOfPages.Text = ""
        txtPlaceOfPublisher.Text = ""
        txtPrice.Text = ""
        txtPublisherName.Text = ""
        txtPublishingYear.Text = ""
        txtRemarks.Text = ""
        txtSupplierID.Text = ""
        txtVolume.Text = ""
        cmbCategory.SelectedIndex = -1
        cmbClass.SelectedIndex = -1
        cmbSubCategory.SelectedIndex = -1
        cmbCondition.SelectedIndex = -1
        dtpBillDate.Text = Today
        dtpEntryDate.Text = Today
        cmbCategory.Enabled = False
        cmbSubCategory.Enabled = False
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        rbNormal.Checked = False
        rbReference.Checked = False
        txtAccessionNo.Focus()
    End Sub
    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DeleteRecord()

        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "select Book.AccessionNo from Book,Quotation_Join where Book.AccessionNo=Quotation_Join.AccessionNo and Book.AccessionNo=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Quotation", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cl1 As String = "select Book.AccessionNo from Book,BookIssue_Student where Book.AccessionNo=BookIssue_Student.AccessionNo and Book.AccessionNo=@d1"
            cmd = New SqlCommand(cl1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Book Issue [Students]", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cl2 As String = "select Book.AccessionNo from Book,BookIssue_Staff where Book.AccessionNo=BookIssue_Staff.AccessionNo and Book.AccessionNo=@d1"
            cmd = New SqlCommand(cl2)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Book Issue [Staffs]", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cl3 As String = "select Book.AccessionNo from Book,BookReservation where Book.AccessionNo=BookReservation.AccessionNo and Book.AccessionNo=@d1"
            cmd = New SqlCommand(cl3)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Book Reservation", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Book where AccessionNo=@d1"
            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the book '" & txtBookTitle.Text & "' having Accession No. '" & txtAccessionNo.Text & "'")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                AutocompleteAuthor()
                AutocompleteBookTitle()
                AutocompleteJointAuthors()
                AutocompleteLanguage()
                AutocompletePublisherPlace()
                AutocompletePublisher()
                AutocompleteBookPosition()
            Else
                MessageBox.Show("No Record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub AutocompleteBookTitle()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Distinct (BookTitle) FROM Book", con)
            ds = New DataSet()
            adp = New SqlDataAdapter(cmd)
            adp.Fill(ds, "Book")
            Dim col As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("BookTitle").ToString())
            Next
            txtBookTitle.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtBookTitle.AutoCompleteCustomSource = col
            txtBookTitle.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub AutocompleteAuthor()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Distinct Author FROM Book", con)
            ds = New DataSet()
            adp = New SqlDataAdapter(cmd)
            adp.Fill(ds, "Book")
            Dim col As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("Author").ToString())
            Next
            txtAuthor.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtAuthor.AutoCompleteCustomSource = col
            txtAuthor.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub AutocompleteJointAuthors()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Distinct JointAuthors FROM Book", con)
            ds = New DataSet()
            adp = New SqlDataAdapter(cmd)
            adp.Fill(ds, "Book")
            Dim col As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("JointAuthors").ToString())
            Next
            txtJointAuthor.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtJointAuthor.AutoCompleteCustomSource = col
            txtJointAuthor.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub AutocompletePublisherPlace()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Distinct PlaceOfPublisher FROM Book", con)
            ds = New DataSet()
            adp = New SqlDataAdapter(cmd)
            adp.Fill(ds, "Book")
            Dim col As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("PlaceOfPublisher").ToString())
            Next
            txtPlaceOfPublisher.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtPlaceOfPublisher.AutoCompleteCustomSource = col
            txtPlaceOfPublisher.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub AutocompleteLanguage()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Distinct [language] FROM Book", con)
            ds = New DataSet()
            adp = New SqlDataAdapter(cmd)
            adp.Fill(ds, "Book")
            Dim col As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("language").ToString())
            Next
            txtLanguage.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtLanguage.AutoCompleteCustomSource = col
            txtLanguage.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub AutocompletePublisher()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Distinct Publisher FROM Book", con)
            ds = New DataSet()
            adp = New SqlDataAdapter(cmd)
            adp.Fill(ds, "Book")
            Dim col As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("Publisher").ToString())
            Next
            txtPublisherName.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtPublisherName.AutoCompleteCustomSource = col
            txtPublisherName.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub AutocompleteBookPosition()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Distinct BookPosition FROM Book", con)
            ds = New DataSet()
            adp = New SqlDataAdapter(cmd)
            adp.Fill(ds, "Book")
            Dim col As AutoCompleteStringCollection = New AutoCompleteStringCollection()
            Dim i As Integer = 0
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("BookPosition").ToString())
            Next
            txtBookPosition.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtBookPosition.AutoCompleteCustomSource = col
            txtBookPosition.AutoCompleteMode = AutoCompleteMode.Suggest
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub fillCombo()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(Classname) FROM BookClass,Category,SubCategory where BookClass.Classname=Category.Class and Category.ID=SubCategory.CategoryID", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbClass.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbClass.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Fill()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT SubCategory.ID from Category,BookClass,SubCategory where Classname=@d1 and CategoryName=@d2 and SubCategoryName=@d3 and BookClass.ClassName=Category.Class and SubCategory.CategoryID=Category.ID"
            cmd.Parameters.AddWithValue("@d1", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d2", cmbCategory.Text)
            cmd.Parameters.AddWithValue("@d3", cmbSubCategory.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtClassificationID.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub frmBookEntry_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        AutocompleteAuthor()
        AutocompleteBookTitle()
        AutocompleteJointAuthors()
        AutocompleteLanguage()
        AutocompletePublisherPlace()
        AutocompletePublisher()
        AutocompleteBookPosition()
        fillCombo()
    End Sub

    Private Sub cmbClass_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbClass.SelectedIndexChanged
        Try
            cmbCategory.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(CategoryName) FROM BookClass,Category,SubCategory where Category.Class=BookClass.ClassName and Category.ID=SubCategory.CategoryID and Classname=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbClass.Text)
            rdr = cmd.ExecuteReader()
            cmbCategory.Items.Clear()
            While rdr.Read
                cmbCategory.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbCategory_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try
            cmbSubCategory.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(SubCategoryName) FROM BookClass,Category,SubCategory where Category.Class=BookClass.ClassName and Category.ID=SubCategory.CategoryID and Classname=@d1 and CategoryName=@d2"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbClass.Text)
            cmd.Parameters.AddWithValue("@d2", cmbCategory.Text)
            rdr = cmd.ExecuteReader()
            cmbSubCategory.Items.Clear()
            While rdr.Read
                cmbSubCategory.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtAccessionNo.Text)) = 0 Then
            MessageBox.Show("Please enter accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo.Focus()
            Exit Sub
        End If
        If Len(Trim(txtBookTitle.Text)) = 0 Then
            MessageBox.Show("Please enter book title", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtBookTitle.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAuthor.Text)) = 0 Then
            MessageBox.Show("Please enter author", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAuthor.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbClass.Text)) = 0 Then
            MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbClass.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbCategory.Text)) = 0 Then
            MessageBox.Show("Please select category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbSubCategory.Text)) = 0 Then
            MessageBox.Show("Please select sub category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbSubCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPublisherName.Text)) = 0 Then
            MessageBox.Show("Please enter publisher", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPublisherName.Focus()
            Exit Sub
        End If
        If ((rbNormal.Checked = False) And (rbReference.Checked = False)) Then
            MessageBox.Show("Please select section", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
     
        If Len(Trim(txtNoOfPages.Text)) = 0 Then
            MessageBox.Show("Please enter no. of pages", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNoOfPages.Focus()
            Exit Sub
        End If
       
        If Len(Trim(cmbCondition.Text)) = 0 Then
            MessageBox.Show("Please select condition", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCondition.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPrice.Text)) = 0 Then
            MessageBox.Show("Please enter price", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPrice.Focus()
            Exit Sub
        End If
        If Len(Trim(txtSupplierID.Text)) = 0 Then
            MessageBox.Show("Please retrieve supplier info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSupplierID.Focus()
            Exit Sub
        End If
        Try

            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select AccessionNo from Book where AccessionNo=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            rdr = cmd.ExecuteReader()

            If rdr.Read() Then
                MessageBox.Show("Accession No. Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                txtAccessionNo.Text = ""
                txtAccessionNo.Focus()
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            If (rbNormal.Checked = True) Then
                s = rbNormal.Text
            End If
            If (rbReference.Checked = True) Then
                s = rbReference.Text
            End If
            Fill()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Book(AccessionNo, BookTitle, EntryDate, Author, JointAuthors, SubCategoryID, Barcode, ISBN, Volume, Edition, Publisher, PlaceOfPublisher, PublishingYear, [Section], [Language], BookPosition, Price, SupplierID,BillNo, BillDate, NoOfPages, Condition, Status, Remarks) VALUES (@d1,@d2,@EntryDate,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14,@d15,@d16,@d17,@d18,@BillDate,@d19,@d20,'Available',@d21)"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Parameters.AddWithValue("@d2", txtBookTitle.Text)
            cmd.Parameters.AddWithValue("@d3", txtAuthor.Text)
            cmd.Parameters.AddWithValue("@d4", txtJointAuthor.Text)
            cmd.Parameters.AddWithValue("@d5", txtClassificationID.Text)
            cmd.Parameters.AddWithValue("@d6", txtBarcode.Text)
            cmd.Parameters.AddWithValue("@d7", txtISBN.Text)
            cmd.Parameters.AddWithValue("@d8", txtVolume.Text)
            cmd.Parameters.AddWithValue("@d9", txtEdition.Text)
            cmd.Parameters.AddWithValue("@d10", txtPublisherName.Text)
            cmd.Parameters.AddWithValue("@d11", txtPlaceOfPublisher.Text)
            cmd.Parameters.AddWithValue("@d12", txtPublishingYear.Text)
            cmd.Parameters.AddWithValue("@d13", s)
            cmd.Parameters.AddWithValue("@d14", txtLanguage.Text)
            cmd.Parameters.AddWithValue("@d15", txtBookPosition.Text)
            cmd.Parameters.AddWithValue("@d16", Val(txtPrice.Text))
            cmd.Parameters.AddWithValue("@d17", Val(txtID.Text))
            cmd.Parameters.AddWithValue("@d18", txtBillNo.Text)
            cmd.Parameters.AddWithValue("@d19", Val(txtNoOfPages.Text))
            cmd.Parameters.AddWithValue("@d20", cmbCondition.Text)
            cmd.Parameters.AddWithValue("@d21", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@EntryDate", dtpEntryDate.Value.Date)
            cmd.Parameters.AddWithValue("@BillDate", dtpBillDate.Value.Date)
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "added the new book '" & txtBookTitle.Text & "' having Accession No. '" & txtAccessionNo.Text & "'")
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            AutocompleteAuthor()
            AutocompleteBookTitle()
            AutocompleteJointAuthors()
            AutocompleteLanguage()
            AutocompletePublisherPlace()
            AutocompletePublisher()
            AutocompleteBookPosition()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtAccessionNo.Text)) = 0 Then
            MessageBox.Show("Please enter accession no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAccessionNo.Focus()
            Exit Sub
        End If
        If Len(Trim(txtBookTitle.Text)) = 0 Then
            MessageBox.Show("Please enter book title", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtBookTitle.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAuthor.Text)) = 0 Then
            MessageBox.Show("Please enter author", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtAuthor.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbClass.Text)) = 0 Then
            MessageBox.Show("Please select class", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbClass.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbCategory.Text)) = 0 Then
            MessageBox.Show("Please select category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbSubCategory.Text)) = 0 Then
            MessageBox.Show("Please select sub category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbSubCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPublisherName.Text)) = 0 Then
            MessageBox.Show("Please enter publisher", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPublisherName.Focus()
            Exit Sub
        End If
        If ((rbNormal.Checked = False) And (rbReference.Checked = False)) Then
            MessageBox.Show("Please select section", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
      
        If Len(Trim(txtNoOfPages.Text)) = 0 Then
            MessageBox.Show("Please enter no. of pages", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNoOfPages.Focus()
            Exit Sub
        End If
       
        If Len(Trim(cmbCondition.Text)) = 0 Then
            MessageBox.Show("Please select condition", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCondition.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPrice.Text)) = 0 Then
            MessageBox.Show("Please enter price", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPrice.Focus()
            Exit Sub
        End If
        If Len(Trim(txtSupplierID.Text)) = 0 Then
            MessageBox.Show("Please retrieve supplier info", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSupplierID.Focus()
            Exit Sub
        End If
        Try
            If (rbNormal.Checked = True) Then
                s = rbNormal.Text
            End If
            If (rbReference.Checked = True) Then
                s = rbReference.Text
            End If
            Fill()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Book set AccessionNo=@d1, BookTitle=@d2, EntryDate=@EntryDate, Author=@d3, JointAuthors=@d4, SubCategoryID=@d5, Barcode=@d6, ISBN=@d7, Volume=@d8, Edition=@d9, Publisher=@d10, PlaceOfPublisher=@d11, PublishingYear=@d12, [Section]=@d13, [Language]=@d14, BookPosition=@d15, Price=@d16, SupplierID=@d17,BillNo=@d18, BillDate=@BillDate, NoOfPages=@d19, Condition=@d20,Remarks=@d21 where AccessionNo=@d22"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtAccessionNo.Text)
            cmd.Parameters.AddWithValue("@d2", txtBookTitle.Text)
            cmd.Parameters.AddWithValue("@d3", txtAuthor.Text)
            cmd.Parameters.AddWithValue("@d4", txtJointAuthor.Text)
            cmd.Parameters.AddWithValue("@d5", txtClassificationID.Text)
            cmd.Parameters.AddWithValue("@d6", txtBarcode.Text)
            cmd.Parameters.AddWithValue("@d7", txtISBN.Text)
            cmd.Parameters.AddWithValue("@d8", txtVolume.Text)
            cmd.Parameters.AddWithValue("@d9", txtEdition.Text)
            cmd.Parameters.AddWithValue("@d10", txtPublisherName.Text)
            cmd.Parameters.AddWithValue("@d11", txtPlaceOfPublisher.Text)
            cmd.Parameters.AddWithValue("@d12", txtPublishingYear.Text)
            cmd.Parameters.AddWithValue("@d13", s)
            cmd.Parameters.AddWithValue("@d14", txtLanguage.Text)
            cmd.Parameters.AddWithValue("@d15", txtBookPosition.Text)
            cmd.Parameters.AddWithValue("@d16", Val(txtPrice.Text))
            cmd.Parameters.AddWithValue("@d17", Val(txtID.Text))
            cmd.Parameters.AddWithValue("@d18", txtBillNo.Text)
            cmd.Parameters.AddWithValue("@d19", Val(txtNoOfPages.Text))
            cmd.Parameters.AddWithValue("@d20", cmbCondition.Text)
            cmd.Parameters.AddWithValue("@d21", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d22", txtAccessionNo1.Text)
            cmd.Parameters.AddWithValue("@EntryDate", dtpEntryDate.Value.Date)
            cmd.Parameters.AddWithValue("@BillDate", dtpBillDate.Value.Date)
            cmd.ExecuteReader()
            con.Close()
            LogFunc(lblUser.Text, "updated the book '" & txtBookTitle.Text & "' record having Accession No. '" & txtAccessionNo.Text & "'")
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            AutocompleteAuthor()
            AutocompleteBookTitle()
            AutocompleteJointAuthors()
            AutocompleteLanguage()
            AutocompletePublisherPlace()
            AutocompletePublisher()
            AutocompleteBookPosition()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub txtNoOfPages_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtNoOfPages.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPrice_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtPrice.Text
            Dim selectionStart = Me.txtPrice.SelectionStart
            Dim selectionLength = Me.txtPrice.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        frmSupplierRecord.lblSet.Text = "Book Entry"
        frmSupplierRecord.Getdata()
        frmSupplierRecord.ShowDialog()
    End Sub

    Private Sub txtPublishingYear_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPublishingYear.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmBookEntryRecord.lblSet.Text = "Book Entry"
        frmBookEntryRecord.Reset()
        frmBookEntryRecord.ShowDialog()
    End Sub

    Private Sub txtPublisherName_LostFocus(sender As Object, e As System.EventArgs) Handles txtPublisherName.LostFocus
        txtPublisherName.Text = txtPublisherName.Text.Trim()
    End Sub

    Private Sub txtPlaceOfPublisher_LostFocus(sender As Object, e As System.EventArgs) Handles txtPlaceOfPublisher.LostFocus
        txtPlaceOfPublisher.Text = txtPlaceOfPublisher.Text.Trim()
    End Sub

    Private Sub txtBookTitle_LostFocus(sender As Object, e As System.EventArgs) Handles txtBookTitle.LostFocus
        txtBookTitle.Text = txtBookTitle.Text.Trim()
    End Sub

    Private Sub txtJointAuthor_LostFocus(sender As Object, e As System.EventArgs) Handles txtJointAuthor.LostFocus
        txtJointAuthor.Text = txtJointAuthor.Text.Trim()
    End Sub

    Private Sub txtAuthor_LostFocus(sender As Object, e As System.EventArgs) Handles txtAuthor.LostFocus
        txtAuthor.Text = txtAuthor.Text.Trim()
    End Sub

    Private Sub txtLanguage_LostFocus(sender As Object, e As System.EventArgs) Handles txtLanguage.LostFocus
        txtLanguage.Text = txtLanguage.Text.Trim()
    End Sub
End Class
