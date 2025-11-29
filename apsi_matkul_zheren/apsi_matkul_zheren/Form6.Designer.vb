<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form6
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form6))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnhome = New System.Windows.Forms.Button()
        Me.btnback = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnsearch = New System.Windows.Forms.Button()
        Me.txtsearch = New System.Windows.Forms.TextBox()
        Me.cmbjenisbb = New System.Windows.Forms.ComboBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.btnedit = New System.Windows.Forms.Button()
        Me.btninput = New System.Windows.Forms.Button()
        Me.btndelete = New System.Windows.Forms.Button()
        Me.txtnamabb = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.btndeleteall = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txttotbb = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtnoid = New System.Windows.Forms.TextBox()
        Me.txtkodebb = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.OldLace
        Me.Panel1.Controls.Add(Me.btnhome)
        Me.Panel1.Controls.Add(Me.btnback)
        Me.Panel1.Location = New System.Drawing.Point(699, 433)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(113, 40)
        Me.Panel1.TabIndex = 18
        '
        'btnhome
        '
        Me.btnhome.BackColor = System.Drawing.Color.Bisque
        Me.btnhome.BackgroundImage = CType(resources.GetObject("btnhome.BackgroundImage"), System.Drawing.Image)
        Me.btnhome.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnhome.Location = New System.Drawing.Point(64, 6)
        Me.btnhome.Margin = New System.Windows.Forms.Padding(2)
        Me.btnhome.Name = "btnhome"
        Me.btnhome.Size = New System.Drawing.Size(37, 30)
        Me.btnhome.TabIndex = 2
        Me.btnhome.UseVisualStyleBackColor = False
        '
        'btnback
        '
        Me.btnback.BackColor = System.Drawing.Color.Bisque
        Me.btnback.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnback.ForeColor = System.Drawing.Color.Black
        Me.btnback.Location = New System.Drawing.Point(6, 12)
        Me.btnback.Margin = New System.Windows.Forms.Padding(2)
        Me.btnback.Name = "btnback"
        Me.btnback.Size = New System.Drawing.Size(55, 20)
        Me.btnback.TabIndex = 1
        Me.btnback.Text = "BACK"
        Me.btnback.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.OldLace
        Me.GroupBox1.Controls.Add(Me.btnsearch)
        Me.GroupBox1.Controls.Add(Me.txtsearch)
        Me.GroupBox1.Controls.Add(Me.cmbjenisbb)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.btnedit)
        Me.GroupBox1.Controls.Add(Me.btninput)
        Me.GroupBox1.Controls.Add(Me.btndelete)
        Me.GroupBox1.Controls.Add(Me.txtnamabb)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.btndeleteall)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txttotbb)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtnoid)
        Me.GroupBox1.Controls.Add(Me.txtkodebb)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(14, 72)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(798, 346)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Data Bahan Baku"
        '
        'btnsearch
        '
        Me.btnsearch.BackColor = System.Drawing.Color.Bisque
        Me.btnsearch.BackgroundImage = CType(resources.GetObject("btnsearch.BackgroundImage"), System.Drawing.Image)
        Me.btnsearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnsearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnsearch.ForeColor = System.Drawing.Color.Black
        Me.btnsearch.Location = New System.Drawing.Point(318, 22)
        Me.btnsearch.Margin = New System.Windows.Forms.Padding(2)
        Me.btnsearch.Name = "btnsearch"
        Me.btnsearch.Size = New System.Drawing.Size(39, 34)
        Me.btnsearch.TabIndex = 119
        Me.btnsearch.UseVisualStyleBackColor = False
        '
        'txtsearch
        '
        Me.txtsearch.Location = New System.Drawing.Point(10, 30)
        Me.txtsearch.Margin = New System.Windows.Forms.Padding(2)
        Me.txtsearch.Name = "txtsearch"
        Me.txtsearch.Size = New System.Drawing.Size(306, 20)
        Me.txtsearch.TabIndex = 120
        '
        'cmbjenisbb
        '
        Me.cmbjenisbb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbjenisbb.FormattingEnabled = True
        Me.cmbjenisbb.Location = New System.Drawing.Point(526, 138)
        Me.cmbjenisbb.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbjenisbb.Name = "cmbjenisbb"
        Me.cmbjenisbb.Size = New System.Drawing.Size(242, 21)
        Me.cmbjenisbb.TabIndex = 84
        '
        'DataGridView1
        '
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(10, 61)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 72
        Me.DataGridView1.RowTemplate.Height = 31
        Me.DataGridView1.Size = New System.Drawing.Size(487, 216)
        Me.DataGridView1.TabIndex = 83
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(455, 327)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(44, 13)
        Me.Label12.TabIndex = 82
        Me.Label12.Text = "Delete"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(406, 327)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(29, 13)
        Me.Label11.TabIndex = 81
        Me.Label11.Text = "Edit"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(346, 327)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(36, 13)
        Me.Label10.TabIndex = 80
        Me.Label10.Text = "Input"
        '
        'btnedit
        '
        Me.btnedit.BackColor = System.Drawing.Color.Bisque
        Me.btnedit.BackgroundImage = CType(resources.GetObject("btnedit.BackgroundImage"), System.Drawing.Image)
        Me.btnedit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnedit.Location = New System.Drawing.Point(399, 289)
        Me.btnedit.Margin = New System.Windows.Forms.Padding(2)
        Me.btnedit.Name = "btnedit"
        Me.btnedit.Size = New System.Drawing.Size(44, 36)
        Me.btnedit.TabIndex = 78
        Me.btnedit.UseVisualStyleBackColor = False
        '
        'btninput
        '
        Me.btninput.BackColor = System.Drawing.Color.Bisque
        Me.btninput.BackgroundImage = CType(resources.GetObject("btninput.BackgroundImage"), System.Drawing.Image)
        Me.btninput.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btninput.Location = New System.Drawing.Point(343, 289)
        Me.btninput.Margin = New System.Windows.Forms.Padding(2)
        Me.btninput.Name = "btninput"
        Me.btninput.Size = New System.Drawing.Size(42, 36)
        Me.btninput.TabIndex = 77
        Me.btninput.UseVisualStyleBackColor = False
        '
        'btndelete
        '
        Me.btndelete.BackColor = System.Drawing.Color.Bisque
        Me.btndelete.BackgroundImage = CType(resources.GetObject("btndelete.BackgroundImage"), System.Drawing.Image)
        Me.btndelete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btndelete.Location = New System.Drawing.Point(457, 289)
        Me.btndelete.Margin = New System.Windows.Forms.Padding(2)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(40, 36)
        Me.btndelete.TabIndex = 79
        Me.btndelete.UseVisualStyleBackColor = False
        '
        'txtnamabb
        '
        Me.txtnamabb.Location = New System.Drawing.Point(525, 180)
        Me.txtnamabb.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtnamabb.Name = "txtnamabb"
        Me.txtnamabb.Size = New System.Drawing.Size(244, 20)
        Me.txtnamabb.TabIndex = 76
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(634, 292)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(44, 13)
        Me.Label13.TabIndex = 75
        Me.Label13.Text = "Delete"
        '
        'btndeleteall
        '
        Me.btndeleteall.BackColor = System.Drawing.Color.Bisque
        Me.btndeleteall.BackgroundImage = CType(resources.GetObject("btndeleteall.BackgroundImage"), System.Drawing.Image)
        Me.btndeleteall.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btndeleteall.Location = New System.Drawing.Point(634, 254)
        Me.btndeleteall.Margin = New System.Windows.Forms.Padding(2)
        Me.btndeleteall.Name = "btndeleteall"
        Me.btndeleteall.Size = New System.Drawing.Size(39, 36)
        Me.btndeleteall.TabIndex = 74
        Me.btndeleteall.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(524, 204)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(109, 13)
        Me.Label5.TabIndex = 72
        Me.Label5.Text = "Total Bahan Baku"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(523, 163)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 13)
        Me.Label8.TabIndex = 71
        Me.Label8.Text = "Nama Bahan Baku"
        '
        'txttotbb
        '
        Me.txttotbb.Location = New System.Drawing.Point(526, 219)
        Me.txttotbb.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txttotbb.Name = "txttotbb"
        Me.txttotbb.Size = New System.Drawing.Size(244, 20)
        Me.txttotbb.TabIndex = 70
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(523, 122)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 13)
        Me.Label2.TabIndex = 69
        Me.Label2.Text = "Jenis Bahan Baku"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(521, 43)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(109, 13)
        Me.Label7.TabIndex = 68
        Me.Label7.Text = "Kode Bahan Baku"
        '
        'txtnoid
        '
        Me.txtnoid.Enabled = False
        Me.txtnoid.Location = New System.Drawing.Point(525, 100)
        Me.txtnoid.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtnoid.Name = "txtnoid"
        Me.txtnoid.Size = New System.Drawing.Size(244, 20)
        Me.txtnoid.TabIndex = 66
        '
        'txtkodebb
        '
        Me.txtkodebb.Location = New System.Drawing.Point(525, 61)
        Me.txtkodebb.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtkodebb.Name = "txtkodebb"
        Me.txtkodebb.Size = New System.Drawing.Size(244, 20)
        Me.txtkodebb.TabIndex = 65
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.142858!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(523, 81)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 64
        Me.Label6.Text = "No. ID"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.apsi_matkul_zheren.My.Resources.Resources.logo_kia
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.ErrorImage = CType(resources.GetObject("PictureBox1.ErrorImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(720, 11)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(107, 50)
        Me.PictureBox1.TabIndex = 126
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Snow
        Me.Label1.Font = New System.Drawing.Font("Cooper Black", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Peru
        Me.Label1.Location = New System.Drawing.Point(187, 19)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(432, 42)
        Me.Label1.TabIndex = 125
        Me.Label1.Text = "PT KERAMIKA INDONESIA ASSOCIATION" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "TBK (KIA)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Form6
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.apsi_matkul_zheren.My.Resources.Resources.background_keramik
        Me.ClientSize = New System.Drawing.Size(831, 486)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form6"
        Me.Text = "Designer 2"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnhome As Button
    Friend WithEvents btnback As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnsearch As Button
    Friend WithEvents txtsearch As TextBox
    Friend WithEvents cmbjenisbb As ComboBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label12 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents btnedit As Button
    Friend WithEvents btninput As Button
    Friend WithEvents btndelete As Button
    Friend WithEvents txtnamabb As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents btndeleteall As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents txttotbb As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents txtnoid As TextBox
    Friend WithEvents txtkodebb As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
End Class
