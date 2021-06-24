Public Class SopapEquipment
    Private Sub SopapEquipment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("E140002100=PT_40degré_2_Moyen_A")
        ComboBox1.Items.Add("E140002110=PT_Marbre_Groupe_01")
        ComboBox1.Items.Add("E140002500_Retourneur_01")
    End Sub
    Dim picture1 As Image
    Dim picture2 As Image
    'Dim picture3 As Image


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Me.ComboBox2.ResetText()
        'Me.'ComboBox3.ResetText()
        Me.ComboBox2.Items.Clear()
        'Me.'ComboBox3.Items.Clear()
        ComboBox1.Focus()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_02100_SOPAP_EQUIPMENT\E140002100=PT_40degré_2_Moyen_A\E140002110.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_02100_SOPAP_EQUIPMENT\E140002110=PT_Marbre_Groupe_01\E140002110.jpg") '2
            'Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_Serumlu Kesit\Serumlu_Kesit.jpg") '3
            With Me.ComboBox2
                Select Case ComboBox1.SelectedItem
                    Case "E140002100=PT_40degré_2_Moyen_A"
                        Image1.Image = picture1
                    Case "E140002110=PT_Marbre_Groupe_01"
                        Image1.Image = picture2
                    Case "E140002500_Retourneur_01"
                        Image1.Image = picture1
                End Select
            End With
            If Me.ComboBox1.SelectedItem.ToString = "E140002100=PT_40degré_2_Moyen_A" Then
                Me.ComboBox2.Items.Add("7000_ANG_BKT_axb_a")
                Me.ComboBox2.Items.Add("7001_ANG_BKT_70x70_th20_a")
                Me.ComboBox2.Items.Add("7002_ANG_BKT_80x70_th20_a")
                Me.ComboBox2.Items.Add("7003_ANG_BKT_90x70_th20_a")
                Me.ComboBox2.Items.Add("7004_ANG_BKT_80x80_th20_a")
                Me.ComboBox2.Items.Add("7005_ANG_BKT_90x90_th20_a")
                Me.ComboBox2.Items.Add("7006_ANG_BKT_70x100_th20_a")
                Me.ComboBox2.Items.Add("7007_ANG_BKT_80x100_th20_a")
                Me.ComboBox2.Items.Add("7008_ANG_BKT_90x100_th20_a")
                Me.ComboBox2.Items.Add("7009_ANG_BKT_100x100_th20_a")

                'Label2.Show()
                'ComboBox2.Show()
                'Label3.Show()
                'ComboBox3.Show()
            ElseIf ComboBox1.SelectedItem.ToString = "7100_type_b" Then
                Me.ComboBox2.Items.Add("7100_ANG_BKT_axb_b")
                Me.ComboBox2.Items.Add("7101_ANG_BKT_70x70_th20_b")
                Me.ComboBox2.Items.Add("7102_ANG_BKT_80x70_th20_b")
                Me.ComboBox2.Items.Add("7103_ANG_BKT_90x70_th20_b")
                Me.ComboBox2.Items.Add("7104_ANG_BKT_80x80_th20_b")
                Me.ComboBox2.Items.Add("7105_ANG_BKT_90x90_th20_b")
                Me.ComboBox2.Items.Add("7106_ANG_BKT_70x100_th20_b")
                Me.ComboBox2.Items.Add("7107_ANG_BKT_80x100_th20_b")
                Me.ComboBox2.Items.Add("7108_ANG_BKT_90x100_th20_b")
                Me.ComboBox2.Items.Add("7109_ANG_BKT_100x100_th20_b")
                'Label2.Show()
                'ComboBox2.Show()
                'Label3.Show()
                'ComboBox3.Show()
            ElseIf ComboBox1.SelectedItem.ToString = "7200_type_c" Then
                Me.ComboBox2.Items.Add("7200_ANG_BKT_axb_c")
                Me.ComboBox2.Items.Add("7201_ANG_BKT_70x70_th20_c")
                Me.ComboBox2.Items.Add("7202_ANG_BKT_80x70_th20_c")
                Me.ComboBox2.Items.Add("7203_ANG_BKT_90x70_th20_c")
                Me.ComboBox2.Items.Add("7204_ANG_BKT_80x80_th20_c")
                Me.ComboBox2.Items.Add("7205_ANG_BKT_90x90_th20_c")
                Me.ComboBox2.Items.Add("7206_ANG_BKT_70x100_th20_c")
                Me.ComboBox2.Items.Add("7207_ANG_BKT_80x100_th20_c")
                Me.ComboBox2.Items.Add("7208_ANG_BKT_90x100_th20_c")
                Me.ComboBox2.Items.Add("7209_ANG_BKT_100x100_th20_c")
                Label2.Show()
                ComboBox2.Show()
                Label3.Show()
                'ComboBox3.Hide()
                Image1.Image = picture1
            ElseIf ComboBox1.SelectedItem.ToString = "7500_type_DIN" Then
                Me.ComboBox2.Items.Add("7500_DIN_BRACKET_70x70")
                Me.ComboBox2.Items.Add("7501_DIN_BRACKET_70x90")
                Me.ComboBox2.Items.Add("7502_DIN_BRACKET_70x110")
                Me.ComboBox2.Items.Add("7503_DIN_BRACKET_70x130")
                'Label2.Show()
                'ComboBox2.Show()
                Image1.Image = picture1
            ElseIf ComboBox1.SelectedItem.ToString = "7300" Then
                Label2.Hide()
                ComboBox2.Hide()
                'Label3.Show()
                'ComboBox3.Hide()
                Image1.Image = picture1
            End If
        Catch ex As Exception
        End Try
    End Sub

    'Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
    '    'ComboBox3.SelectedIndexChanged
    '    ''ComboBox3.Focus()
    '    'Try
    '    '    'Horizontal
    '    '    Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2000H.png") '1
    '    '    Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2001H.png") '2
    '    '    Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2002H.png") '3
    '    '    Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2003H.png") '4
    '    '    Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2010H.png") '5
    '    '    Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2011H.png") '6
    '    '    Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2012H.png") '7
    '    '    Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2013H.png") '8
    '    '    Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2020H.png") '9
    '    '    Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2021H.png") '10
    '    '    Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2022H.png") '11
    '    '    Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2023H.png") '12
    '    '    'Vertical
    '    '    Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2050.png") '13
    '    '    Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2051.jpg") '14
    '    '    Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2052.png") '15
    '    '    Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2053.png") '16
    '    '    Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2056.png") '17
    '    '    Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2057.jpg") '18
    '    '    Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2058.png") '19
    '    '    Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2059.png") '20
    '    'Catch ex As Exception

    '    'End Try
    'End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'ComboBox2.Focus()
        ''ComboBox3.ResetText()
        ''Me.'ComboBox3.Items.Clear()
        'Try
        '    Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_12\5104_GAUGE(12)_20_L60.jpg") '1
        '    Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_ep_20x20_nylon\5223_GAUGE(20)_20_L50_nylon.jpg") '2
        '    Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_Serumlu Kesit\Serumlu_Kesit.jpg") '3
        '    With Me.ComboBox2
        '        If ComboBox1.SelectedItem = "Gauge" Then
        '            If ComboBox2.SelectedItem = "20 [mm]" Then
        '                Image1.Image = picture1
        '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '            ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
        '                Image1.Image = picture1
        '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '            ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
        '                Image1.Image = picture1
        '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '            ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
        '                Image1.Image = picture1
        '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '            ElseIf ComboBox2.SelectedItem = "60 [mm]" Then
        '                Image1.Image = picture1
        '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '            End If
        '        End If
        '        If ComboBox1.SelectedItem = "Nylon Gauge" Then
        '            If ComboBox2.SelectedItem = "20 [mm]" Then
        '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '                Image1.Image = picture2
        '            ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
        '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '                Image1.Image = picture2
        '            ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
        '                Image1.Image = picture2
        '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '            ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
        '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '                Image1.Image = picture2
        '            End If
        '        End If
        '    End With
        'Catch ex As Exception
        'End Try
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        If Me.ComboBox1.SelectedItem.ToString = "E140002100=PT_40degré_2_Moyen_A" Then
            'Get CATIA or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_02100_SOPAP_EQUIPMENT\E140002100=PT_40degré_2_Moyen_A\3D\Vis_Moyen_2509_02.CATProduct")

            On Error GoTo 0
        ElseIf ComboBox1.SelectedItem.ToString = "E140002110=PT_Marbre_Groupe_01" Then
            'Get CATIA or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_02100_SOPAP_EQUIPMENT\E140002110=PT_Marbre_Groupe_01\3D\PT_Marbre_Groupe_01.CATProduct")

            On Error GoTo 0
        ElseIf ComboBox1.SelectedItem.ToString = "E140002500_Retourneur_01" Then
            'Get CATIA or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_02100_SOPAP_EQUIPMENT\E140002500_Retourneur_01\3D\Vis_RT_STRUCTURE_2520.CATProduct")

            On Error GoTo 0
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class