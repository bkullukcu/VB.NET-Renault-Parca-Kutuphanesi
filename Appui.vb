Public Class Appui
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image
    Private Sub Appui_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Gauge")
        ComboBox1.Items.Add("Nylon Gauge")
        ComboBox1.Items.Add("Serumlu Kesit")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Me.ComboBox2.ResetText()
        Me.ComboBox3.ResetText()
        Me.ComboBox2.Items.Clear()
        Me.ComboBox3.Items.Clear()
        ComboBox1.Focus()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5104_GAUGE(12)_20_L60.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_20x20_nylon\5223_GAUGE(20)_20_L50_nylon.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_Serumlu Kesit\Serumlu_Kesit.jpg") '3
            With Me.ComboBox2
                Select Case ComboBox1.SelectedItem
                    Case "Gauge"
                        Image1.Image = picture1
                    Case "Nylon Gauge"
                        Image1.Image = picture2
                    Case "Serumlu Kesit"
                        Image1.Image = picture3
                End Select
            End With
            If Me.ComboBox1.SelectedItem.ToString = "Gauge" Then
                Me.ComboBox2.Items.Add("20 [mm]")
                Me.ComboBox2.Items.Add("30 [mm]")
                Me.ComboBox2.Items.Add("40 [mm]")
                Me.ComboBox2.Items.Add("50 [mm]")
                Me.ComboBox2.Items.Add("60 [mm]")

                Label2.Show()
                ComboBox2.Show()
                Label1.Show()
                ComboBox3.Show()
            ElseIf ComboBox1.SelectedItem.ToString = "Nylon Gauge" Then
                Me.ComboBox2.Items.Add("20 [mm]")
                Me.ComboBox2.Items.Add("30 [mm]")
                Me.ComboBox2.Items.Add("40 [mm]")
                Me.ComboBox2.Items.Add("50 [mm]")

                Label2.Show()
                ComboBox2.Show()
                Label1.Show()
                ComboBox3.Show()
            ElseIf ComboBox1.SelectedItem.ToString = "Serumlu Kesit" Then
                Label2.Hide()
                ComboBox2.Hide()
                Label1.Hide()
                ComboBox3.Hide()
                Image1.Image = picture3
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        'ComboBox3.Focus()
        'Try
        '    'Horizontal
        '    Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2000H.png") '1
        '    Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2001H.png") '2
        '    Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2002H.png") '3
        '    Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2003H.png") '4
        '    Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2010H.png") '5
        '    Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2011H.png") '6
        '    Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2012H.png") '7
        '    Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2013H.png") '8
        '    Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2020H.png") '9
        '    Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2021H.png") '10
        '    Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2022H.png") '11
        '    Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2023H.png") '12
        '    'Vertical
        '    Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2050.png") '13
        '    Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2051.jpg") '14
        '    Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2052.png") '15
        '    Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2053.png") '16
        '    Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2056.png") '17
        '    Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2057.jpg") '18
        '    Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2058.png") '19
        '    Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2059.png") '20
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox2.Focus()
        ComboBox3.ResetText()
        Me.ComboBox3.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5104_GAUGE(12)_20_L60.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_20x20_nylon\5223_GAUGE(20)_20_L50_nylon.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_Serumlu Kesit\Serumlu_Kesit.jpg") '3
            With Me.ComboBox2
                If ComboBox1.SelectedItem = "Gauge" Then
                    If ComboBox2.SelectedItem = "20 [mm]" Then
                        Image1.Image = picture1
                        Me.ComboBox3.Items.Add("7 [mm]")
                    ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
                        Image1.Image = picture1
                        Me.ComboBox3.Items.Add("7 [mm]")
                    ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
                        Image1.Image = picture1
                        Me.ComboBox3.Items.Add("7 [mm]")
                    ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
                        Image1.Image = picture1
                        Me.ComboBox3.Items.Add("7 [mm]")
                    ElseIf ComboBox2.SelectedItem = "60 [mm]" Then
                        Image1.Image = picture1
                        Me.ComboBox3.Items.Add("7 [mm]")
                    End If
                End If
                If ComboBox1.SelectedItem = "Nylon Gauge" Then
                    If ComboBox2.SelectedItem = "20 [mm]" Then
                        Me.ComboBox3.Items.Add("8 [mm]")
                        Image1.Image = picture2
                    ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
                        Me.ComboBox3.Items.Add("8 [mm]")
                        Image1.Image = picture2
                    ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
                        Image1.Image = picture2
                        Me.ComboBox3.Items.Add("8 [mm]")
                    ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
                        Me.ComboBox3.Items.Add("8 [mm]")
                        Image1.Image = picture2
                    End If
                End If
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi001.Show()
        Me.Close()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        With Me.ComboBox2
            If ComboBox1.SelectedItem = "Gauge" Then
                If ComboBox2.SelectedItem = "20 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5100_GAUGE(12)_20_L20.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5101_GAUGE(12)_20_L30.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5102_GAUGE(12)_20_L40.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5103_GAUGE(12)_20_L50.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "60 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5104_GAUGE(12)_20_L60.CATPart")

                    On Error GoTo 0
                End If

            ElseIf ComboBox1.SelectedItem = "Nylon Gauge" Then
                If ComboBox2.SelectedItem = "20 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_20x20_nylon\#5220_SKIN_GAUGE(20)x20_L20_nylon.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_20x20_nylon\#5221_SKIN_GAUGE(20)x20_L30_nylon.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_20x20_nylon\#5222_SKIN_GAUGE(20)x20_L40_nylon.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_20x20_nylon\#5223_SKIN_GAUGE(20)x20_L50_nylon.CATProduct")

                    On Error GoTo 0
                End If
            ElseIf ComboBox1.SelectedItem = "Serumlu Kesit" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_Serumlu Kesit\#Kordonlu Dis Yuzey Kesit.CATProduct")
                On Error GoTo 0

            End If

        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PictureBox1.Visible = True
        Button4.Visible = True
        'Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_Serumlu Kesit\Serumlu_Kesit.jpg") '3
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        PictureBox1.Visible = False
        Button4.Visible = False
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class