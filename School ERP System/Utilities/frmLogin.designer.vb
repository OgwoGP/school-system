<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
<Global.System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1726")> _
Partial Class frmLogin
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
    Friend WithEvents UsernameLabel As System.Windows.Forms.Label
    Friend WithEvents PasswordLabel As System.Windows.Forms.Label
    Friend WithEvents UserID As System.Windows.Forms.TextBox
    Friend WithEvents Password As System.Windows.Forms.TextBox
    Friend WithEvents OK As System.Windows.Forms.Button
    Friend WithEvents Cancel As System.Windows.Forms.Button

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
        Me.UsernameLabel = New System.Windows.Forms.Label()
        Me.PasswordLabel = New System.Windows.Forms.Label()
        Me.UserID = New System.Windows.Forms.TextBox()
        Me.Password = New System.Windows.Forms.TextBox()
        Me.UserType = New System.Windows.Forms.TextBox()
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.OK = New System.Windows.Forms.Button()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'UsernameLabel
        '
        Me.UsernameLabel.BackColor = System.Drawing.Color.Transparent
        Me.UsernameLabel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UsernameLabel.ForeColor = System.Drawing.Color.Black
        Me.UsernameLabel.Location = New System.Drawing.Point(7, 5)
        Me.UsernameLabel.Name = "UsernameLabel"
        Me.UsernameLabel.Size = New System.Drawing.Size(69, 23)
        Me.UsernameLabel.TabIndex = 0
        Me.UsernameLabel.Text = "&User ID : "
        Me.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PasswordLabel
        '
        Me.PasswordLabel.BackColor = System.Drawing.Color.Transparent
        Me.PasswordLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PasswordLabel.ForeColor = System.Drawing.Color.Black
        Me.PasswordLabel.Location = New System.Drawing.Point(7, 42)
        Me.PasswordLabel.Name = "PasswordLabel"
        Me.PasswordLabel.Size = New System.Drawing.Size(80, 23)
        Me.PasswordLabel.TabIndex = 2
        Me.PasswordLabel.Text = "&Password :"
        Me.PasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UserID
        '
        Me.UserID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserID.Location = New System.Drawing.Point(82, 6)
        Me.UserID.Name = "UserID"
        Me.UserID.Size = New System.Drawing.Size(166, 22)
        Me.UserID.TabIndex = 0
        '
        'Password
        '
        Me.Password.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Password.Location = New System.Drawing.Point(82, 43)
        Me.Password.Name = "Password"
        Me.Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.Password.Size = New System.Drawing.Size(166, 22)
        Me.Password.TabIndex = 1
        '
        'UserType
        '
        Me.UserType.Location = New System.Drawing.Point(101, 196)
        Me.UserType.Name = "UserType"
        Me.UserType.Size = New System.Drawing.Size(55, 20)
        Me.UserType.TabIndex = 10
        Me.UserType.Visible = False
        '
        'LinkLabel2
        '
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.BackColor = System.Drawing.Color.Transparent
        Me.LinkLabel2.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel2.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.LinkLabel2.LinkColor = System.Drawing.Color.Black
        Me.LinkLabel2.Location = New System.Drawing.Point(58, 147)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(138, 21)
        Me.LinkLabel2.TabIndex = 4
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Change Password"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.YellowGreen
        Me.Panel1.Controls.Add(Me.LinkLabel2)
        Me.Panel1.Controls.Add(Me.OK)
        Me.Panel1.Controls.Add(Me.UsernameLabel)
        Me.Panel1.Controls.Add(Me.Password)
        Me.Panel1.Controls.Add(Me.UserType)
        Me.Panel1.Controls.Add(Me.PasswordLabel)
        Me.Panel1.Controls.Add(Me.Cancel)
        Me.Panel1.Controls.Add(Me.UserID)
        Me.Panel1.Location = New System.Drawing.Point(6, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(274, 177)
        Me.Panel1.TabIndex = 0
        '
        'OK
        '
        Me.OK.BackColor = System.Drawing.Color.White
        Me.OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.OK.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OK.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.OK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.OK.Location = New System.Drawing.Point(10, 93)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(109, 34)
        Me.OK.TabIndex = 2
        Me.OK.Text = "&Login"
        Me.OK.UseVisualStyleBackColor = False
        '
        'Cancel
        '
        Me.Cancel.BackColor = System.Drawing.Color.White
        Me.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancel.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Cancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Cancel.Location = New System.Drawing.Point(139, 93)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(109, 34)
        Me.Cancel.TabIndex = 3
        Me.Cancel.Text = "&Cancel"
        Me.Cancel.UseVisualStyleBackColor = False
        '
        'frmLogin
        '
        Me.AcceptButton = Me.OK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.CancelButton = Me.Cancel
        Me.ClientSize = New System.Drawing.Size(286, 190)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login Form"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UserType As System.Windows.Forms.TextBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel

End Class
