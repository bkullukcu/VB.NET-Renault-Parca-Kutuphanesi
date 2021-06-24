<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Appui
    Inherits System.Windows.Forms.Form

    'Form, bileşen listesini temizlemeyi bırakmayı geçersiz kılar.
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

    'Windows Form Tasarımcısı tarafından gerektirilir
    Private components As System.ComponentModel.IContainer

    'NOT: Aşağıdaki yordam Windows Form Tasarımcısı için gereklidir
    'Windows Form Tasarımcısı kullanılarak değiştirilebilir.  
    'Kod düzenleyicisini kullanarak değiştirmeyin.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Appui))
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.PictureBox13 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Image1
        '
        Me.Image1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Image1.Location = New System.Drawing.Point(599, 55)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(313, 318)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.Image1.TabIndex = 42
        Me.Image1.TabStop = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Gold
        Me.Button2.Location = New System.Drawing.Point(68, 74)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 88)
        Me.Button2.TabIndex = 41
        Me.Button2.Text = "Önceki Sayfa"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(82, 249)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 13)
        Me.Label3.TabIndex = 40
        Me.Label3.Text = "Parça Tipini Seçiniz"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(234, 249)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(237, 21)
        Me.ComboBox1.TabIndex = 39
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Gold
        Me.Button1.Location = New System.Drawing.Point(85, 410)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(193, 48)
        Me.Button1.TabIndex = 38
        Me.Button1.Text = "Parçayı CATIA'da Aç"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(82, 321)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 13)
        Me.Label1.TabIndex = 44
        Me.Label1.Text = "Orta Delik Çapını Seçiniz"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(234, 283)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(237, 21)
        Me.ComboBox2.TabIndex = 43
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(82, 283)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 13)
        Me.Label2.TabIndex = 46
        Me.Label2.Text = "Parça Uzunluğunu Seçiniz"
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(234, 321)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(237, 21)
        Me.ComboBox3.TabIndex = 45
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Gold
        Me.Button3.Location = New System.Drawing.Point(85, 481)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(193, 48)
        Me.Button3.TabIndex = 47
        Me.Button3.Text = "Uyarı"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Gold
        Me.Button4.Location = New System.Drawing.Point(719, 477)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(193, 48)
        Me.Button4.TabIndex = 49
        Me.Button4.Text = "Kapat"
        Me.Button4.UseVisualStyleBackColor = False
        Me.Button4.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(-4, 553)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(130, 17)
        Me.Label5.TabIndex = 51
        Me.Label5.Text = "Created by Renault"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(896, 553)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(108, 17)
        Me.Label4.TabIndex = 50
        Me.Label4.Text = "Renault Internal"
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button5.BackgroundImage = Global.ToolingKutuphane.My.Resources.Resources.home_button
        Me.Button5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button5.Location = New System.Drawing.Point(210, 79)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(84, 80)
        Me.Button5.TabIndex = 52
        Me.Button5.UseVisualStyleBackColor = False
        '
        'PictureBox13
        '
        Me.PictureBox13.BackgroundImage = Global.ToolingKutuphane.My.Resources.Resources.logo
        Me.PictureBox13.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox13.Location = New System.Drawing.Point(934, -15)
        Me.PictureBox13.Name = "PictureBox13"
        Me.PictureBox13.Size = New System.Drawing.Size(70, 96)
        Me.PictureBox13.TabIndex = 55
        Me.PictureBox13.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.ToolingKutuphane.My.Resources.Resources.ÖNEMLİ_OKU_____________________________________
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.Location = New System.Drawing.Point(-1, -12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1005, 470)
        Me.PictureBox1.TabIndex = 56
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button6.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.Button6.ForeColor = System.Drawing.Color.Gold
        Me.Button6.Location = New System.Drawing.Point(618, 478)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(60, 50)
        Me.Button6.TabIndex = 73
        Me.Button6.Text = "Notu Kaydet"
        Me.Button6.UseVisualStyleBackColor = False
        Me.Button6.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.ForeColor = System.Drawing.Color.Red
        Me.TextBox1.Location = New System.Drawing.Point(327, 464)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(285, 81)
        Me.TextBox1.TabIndex = 72
        Me.TextBox1.Visible = False
        '
        'Appui
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.BackgroundImage = Global.ToolingKutuphane.My.Resources.Resources.renault_seeklogocom_removebg_preview1
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(1001, 570)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.PictureBox13)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Appui"
        Me.Text = "Appui"
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Image1 As PictureBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Button5 As Button
    Friend WithEvents PictureBox13 As PictureBox
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Button6 As Button
    Friend WithEvents TextBox1 As TextBox
End Class
