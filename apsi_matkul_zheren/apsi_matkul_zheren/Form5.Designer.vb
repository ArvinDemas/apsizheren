<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form5))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnhome = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btncekdatap = New System.Windows.Forms.Button()
        Me.btndatarc = New System.Windows.Forms.Button()
        Me.btndatabb = New System.Windows.Forms.Button()
        Me.btndatapp = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.OldLace
        Me.Panel2.Controls.Add(Me.btnhome)
        Me.Panel2.Location = New System.Drawing.Point(787, 372)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(47, 47)
        Me.Panel2.TabIndex = 21
        '
        'btnhome
        '
        Me.btnhome.BackColor = System.Drawing.Color.Bisque
        Me.btnhome.BackgroundImage = CType(resources.GetObject("btnhome.BackgroundImage"), System.Drawing.Image)
        Me.btnhome.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnhome.Location = New System.Drawing.Point(4, 7)
        Me.btnhome.Margin = New System.Windows.Forms.Padding(2)
        Me.btnhome.Name = "btnhome"
        Me.btnhome.Size = New System.Drawing.Size(37, 30)
        Me.btnhome.TabIndex = 2
        Me.btnhome.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.OldLace
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.btncekdatap)
        Me.Panel1.Controls.Add(Me.btndatarc)
        Me.Panel1.Controls.Add(Me.btndatabb)
        Me.Panel1.Controls.Add(Me.btndatapp)
        Me.Panel1.Location = New System.Drawing.Point(11, 100)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(757, 276)
        Me.Panel1.TabIndex = 20
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.AntiqueWhite
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Location = New System.Drawing.Point(280, 28)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(207, 35)
        Me.Panel3.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Cooper Black", 14.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(11, 7)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(189, 21)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "STAFF DESIGNER"
        '
        'btncekdatap
        '
        Me.btncekdatap.BackColor = System.Drawing.Color.Bisque
        Me.btncekdatap.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncekdatap.ForeColor = System.Drawing.Color.Black
        Me.btncekdatap.Location = New System.Drawing.Point(153, 193)
        Me.btncekdatap.Margin = New System.Windows.Forms.Padding(2)
        Me.btncekdatap.Name = "btncekdatap"
        Me.btncekdatap.Size = New System.Drawing.Size(227, 73)
        Me.btncekdatap.TabIndex = 8
        Me.btncekdatap.Text = "CEK DATA PESANAN"
        Me.btncekdatap.UseVisualStyleBackColor = False
        '
        'btndatarc
        '
        Me.btndatarc.BackColor = System.Drawing.Color.Bisque
        Me.btndatarc.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndatarc.ForeColor = System.Drawing.Color.Black
        Me.btndatarc.Location = New System.Drawing.Point(400, 193)
        Me.btndatarc.Margin = New System.Windows.Forms.Padding(2)
        Me.btndatarc.Name = "btndatarc"
        Me.btndatarc.Size = New System.Drawing.Size(227, 73)
        Me.btndatarc.TabIndex = 7
        Me.btndatarc.Text = "DATA REQUEST CUSTOMER"
        Me.btndatarc.UseVisualStyleBackColor = False
        '
        'btndatabb
        '
        Me.btndatabb.BackColor = System.Drawing.Color.Bisque
        Me.btndatabb.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndatabb.ForeColor = System.Drawing.Color.Black
        Me.btndatabb.Location = New System.Drawing.Point(500, 91)
        Me.btndatabb.Margin = New System.Windows.Forms.Padding(2)
        Me.btndatabb.Name = "btndatabb"
        Me.btndatabb.Size = New System.Drawing.Size(227, 73)
        Me.btndatabb.TabIndex = 6
        Me.btndatabb.Text = "DATA BAHAN BAKU"
        Me.btndatabb.UseVisualStyleBackColor = False
        '
        'btndatapp
        '
        Me.btndatapp.BackColor = System.Drawing.Color.Bisque
        Me.btndatapp.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.85714!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndatapp.ForeColor = System.Drawing.Color.Black
        Me.btndatapp.Location = New System.Drawing.Point(32, 91)
        Me.btndatapp.Margin = New System.Windows.Forms.Padding(2)
        Me.btndatapp.Name = "btndatapp"
        Me.btndatapp.Size = New System.Drawing.Size(227, 73)
        Me.btndatapp.TabIndex = 2
        Me.btndatapp.Text = "DATA PERENCANAAN PRODUKSI"
        Me.btndatapp.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.apsi_matkul_zheren.My.Resources.Resources.logo_kia
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.ErrorImage = CType(resources.GetObject("PictureBox1.ErrorImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(739, 7)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(107, 50)
        Me.PictureBox1.TabIndex = 124
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Snow
        Me.Label1.Font = New System.Drawing.Font("Cooper Black", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Peru
        Me.Label1.Location = New System.Drawing.Point(206, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(432, 42)
        Me.Label1.TabIndex = 123
        Me.Label1.Text = "PT KERAMIKA INDONESIA ASSOCIATION" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "TBK (KIA)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.apsi_matkul_zheren.My.Resources.Resources.background_keramik
        Me.ClientSize = New System.Drawing.Size(855, 427)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "Form5"
        Me.Text = "Designer 1"
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnhome As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents btncekdatap As Button
    Friend WithEvents btndatarc As Button
    Friend WithEvents btndatabb As Button
    Friend WithEvents btndatapp As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
End Class
