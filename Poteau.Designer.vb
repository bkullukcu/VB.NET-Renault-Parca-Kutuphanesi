<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Poteau
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Poteau))
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.PictureBox13 = New System.Windows.Forms.PictureBox()
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox13, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Gold
        Me.Button2.Location = New System.Drawing.Point(70, 54)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 88)
        Me.Button2.TabIndex = 36
        Me.Button2.Text = "Önceki Sayfa"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(100, 234)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(111, 13)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "Parça Boyunu Seçiniz"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(100, 266)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(110, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Bağlantı Tipini Seçiniz"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(47, 296)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(163, 13)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Parçada İzolasyon Cale'i Var mı ?"
        '
        'Image1
        '
        Me.Image1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Image1.Location = New System.Drawing.Point(599, 137)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(313, 318)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.Image1.TabIndex = 32
        Me.Image1.TabStop = False
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(236, 288)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(237, 21)
        Me.ComboBox3.TabIndex = 31
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(236, 258)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(237, 21)
        Me.ComboBox2.TabIndex = 30
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(236, 226)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(237, 21)
        Me.ComboBox1.TabIndex = 29
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Gold
        Me.Button1.Location = New System.Drawing.Point(116, 354)
        Me.Button1.Name = "Button1"
        Me.Button1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Button1.Size = New System.Drawing.Size(193, 48)
        Me.Button1.TabIndex = 28
        Me.Button1.Text = "Parçayı CATIA'da Aç"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(-2, 557)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(130, 17)
        Me.Label5.TabIndex = 53
        Me.Label5.Text = "Created by Renault"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(880, 557)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(108, 17)
        Me.Label4.TabIndex = 52
        Me.Label4.Text = "Renault Internal"
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button5.BackgroundImage = Global.ToolingKutuphane.My.Resources.Resources.home_button
        Me.Button5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button5.Location = New System.Drawing.Point(196, 59)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(84, 80)
        Me.Button5.TabIndex = 54
        Me.Button5.UseVisualStyleBackColor = False
        '
        'PictureBox13
        '
        Me.PictureBox13.BackgroundImage = Global.ToolingKutuphane.My.Resources.Resources.logo
        Me.PictureBox13.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox13.Location = New System.Drawing.Point(918, -10)
        Me.PictureBox13.Name = "PictureBox13"
        Me.PictureBox13.Size = New System.Drawing.Size(70, 96)
        Me.PictureBox13.TabIndex = 69
        Me.PictureBox13.TabStop = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Gold
        Me.Button3.Location = New System.Drawing.Point(116, 408)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(193, 48)
        Me.Button3.TabIndex = 71
        Me.Button3.Text = "Isolation Poteau Aç "
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Poteau
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.BackgroundImage = Global.ToolingKutuphane.My.Resources.Resources.renault_seeklogocom_removebg_preview1
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(986, 583)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.PictureBox13)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Poteau"
        Me.Text = "Poteau"
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox13, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button2 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Image1 As PictureBox
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Button5 As Button
    Friend WithEvents PictureBox13 As PictureBox
    Friend WithEvents Button3 As Button
End Class
