Public Class Detector
    Private Sub Detector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("00_Dedecteur_Equerre_01")
        ComboBox1.Items.Add("01_Detecteur_Equerre_Laser_01")
        ComboBox1.Items.Add("10_Conception_de_reference_")
        ComboBox1.Items.Add("11_Detecteur_Laser_Accessoires")
        ComboBox1.Items.Add("12_Detecteur_Laser_Support_01")
        ComboBox1.Items.Add("13_Detecteur_Laser_Pipe")
        ComboBox1.Items.Add("14_Groupe_fixation_50x50_tube_26,9_pipe")
    End Sub
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image
    Dim picture5 As Image
    Dim picture6 As Image
    Dim picture7 As Image
    Dim picture8 As Image
    Dim picture9 As Image

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Me.ComboBox2.ResetText()
        'Me.'ComboBox3.ResetText()
        Me.ComboBox2.Items.Clear()
        'Me.'ComboBox3.Items.Clear()
        ComboBox1.Focus()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\00_Dedecteur_Equerre_01.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\01_Detecteur_Equerre_Laser_01.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\10_Conception_de_reference.jpg") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\12_Detecteur_Laser_Support_01.jpg") '1
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\13_Detecteur_Laser_Pipe.jpg") '2
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\14_Groupe_fixation_50x50_tube_26,9_pipe.jpg") '3
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Protection_Laser_E1400009928.jpg") '1
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Support_11_E140000926.jpg") '2
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Support_12_E140000927.jpg") '3
            With Me.ComboBox1
                Select Case ComboBox1.SelectedItem
                    Case "00_Dedecteur_Equerre_01"
                        Image1.Image = picture1
                        Label2.Visible = False
                        ComboBox2.Visible = False
                    Case "01_Detecteur_Equerre_Laser_01"
                        Image1.Image = picture2
                        Label2.Visible = False
                        ComboBox2.Visible = False
                    Case "10_Conception_de_reference_"
                        Image1.Image = picture3
                        Label2.Visible = False
                        ComboBox2.Visible = False
                    Case "11_Detecteur_Laser_Accessoires"
                        Image1.Image = picture7
                        Label2.Visible = True
                        ComboBox2.Visible = True
                    Case "12_Detecteur_Laser_Support_01"
                        Image1.Image = picture4
                        Label2.Visible = True
                        ComboBox2.Visible = True
                    Case "13_Detecteur_Laser_Pipe"
                        Image1.Image = picture5
                        Label2.Visible = True
                        ComboBox2.Visible = True
                    Case "14_Groupe_fixation_50x50_tube_26,9_pipe"
                        Label2.Visible = False
                        ComboBox2.Visible = False
                End Select
            End With
            'If Me.ComboBox1.SelectedItem.ToString = "00_Dedecteur_Equerre_01" Then
            '    Me.ComboBox2.Items.Add("Dedecteur_Groupe_050918")
            '    Me.ComboBox2.Items.Add("dedecteur_groupe_01")
            '    'Label2.Show()
            '    'ComboBox2.Show()
            '    'Label3.Show()
            ''    ''ComboBox3.Show()
            'ElseIf ComboBox1.SelectedItem.ToString = "01_Detecteur_Equerre_Laser_01" Then
            '    Me.ComboBox2.Items.Add("Laser_dedecteur_groupe_01_A")
            'ElseIf ComboBox1.SelectedItem.ToString = "7200_type_c" Then
            '    Me.ComboBox2.Items.Add("7200_ANG_BKT_axb_c")
            '    Me.ComboBox2.Items.Add("7201_ANG_BKT_70x70_th20_c")
            '    Me.ComboBox2.Items.Add("7202_ANG_BKT_80x70_th20_c")
            '    Me.ComboBox2.Items.Add("7203_ANG_BKT_90x70_th20_c")
            '    Me.ComboBox2.Items.Add("7204_ANG_BKT_80x80_th20_c")
            '    Me.ComboBox2.Items.Add("7205_ANG_BKT_90x90_th20_c")
            '    Me.ComboBox2.Items.Add("7206_ANG_BKT_70x100_th20_c")
            '    Me.ComboBox2.Items.Add("7207_ANG_BKT_80x100_th20_c")
            '    Me.ComboBox2.Items.Add("7208_ANG_BKT_90x100_th20_c")
            '    Me.ComboBox2.Items.Add("7209_ANG_BKT_100x100_th20_c")
            'Label2.Show()
            'ComboBox2.Show()
            'Label3.Show()
            ''ComboBox3.Hide()
            'Image1.Image = picture1
            'If ComboBox1.SelectedItem.ToString = "10_Conception_de_reference_" Then
            '    Me.ComboBox2.Items.Add("DEQ_Laser_Conception_de_reference")
            '    Me.ComboBox2.Items.Add("Groupe_fixation_50x50_tube_26,9_pipe")
            '    Me.ComboBox2.Items.Add("IFM_Equipment_standard")
            '    Me.ComboBox2.Items.Add("V_CHC08_30_RK__")
            '    Me.ComboBox2.Items.Add("Vis_sans_tete_11_")

            '    'Label2.Show()
            '    'ComboBox2.Show()
            '    'Image1.Image = picture1
        Catch ex As Exception
        End Try
        If ComboBox1.SelectedItem.ToString = "11_Detecteur_Laser_Accessoires" Then
            Me.ComboBox2.Items.Add("DEQ_Protection_Laser_01")
            Me.ComboBox2.Items.Add("DEQ_Support_11")
            Me.ComboBox2.Items.Add("DEQ_Support_12")
        ElseIf ComboBox1.SelectedItem.ToString = "12_Detecteur_Laser_Support_01" Then
            Me.ComboBox2.Items.Add("DEQ_Support_01_L100")
            Me.ComboBox2.Items.Add("DEQ_Support_02_L200")
            Me.ComboBox2.Items.Add("DEQ_Support_03_L300")
            Me.ComboBox2.Items.Add("DEQ_Support_04_L400")
            Me.ComboBox2.Items.Add("DEQ_Support_05_L500")
        ElseIf ComboBox1.SelectedItem.ToString = "13_Detecteur_Laser_Pipe" Then
            Me.ComboBox2.Items.Add("DEQ_26,9_pipe_01_L200")
            Me.ComboBox2.Items.Add("DEQ_26,9_pipe_02_L350")
            Me.ComboBox2.Items.Add("DEQ_26,9_pipe_03_L500")
            Me.ComboBox2.Items.Add("DEQ_26,9_pipe_04_L650")
            Me.ComboBox2.Items.Add("DEQ_26,9_pipe_05_L800")
            'ElseIf ComboBox1.SelectedItem.ToString = "14_Groupe_fixation_50x50_tube_26,9_pipe" Then
            '    Me.ComboBox2.Items.Add("Groupe_fixation_50x50_tube_26,9_pipe")
        End If

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
        ComboBox2.Focus()
        ''ComboBox3.ResetText()
        '''Me.'ComboBox3.Items.Clear()
        'Try
        '    Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\00_Dedecteur_Equerre_01.jpg") '1
        '    Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\01_Detecteur_Equerre_Laser_01.jpg") '2
        '    Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\10_Conception_de_reference.jpg") '3
        '    Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\12_Detecteur_Laser_Support_01.jpg") '1
        '    Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\13_Detecteur_Laser_Pipe.jpg") '2
        '    Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\14_Groupe_fixation_50x50_tube_26,9_pipe.jpg") '3
        '    Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Protection_Laser_E1400009928.jpg") '1
        '    Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Support_11_E140000926.jpg") '2
        '    Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Support_12_E140000927.jpg") '3
        '    '    With Me.ComboBox2
        '    '        If ComboBox1.SelectedItem = "Gauge" Then
        '    '            If ComboBox2.SelectedItem = "20 [mm]" Then
        '    '                Image1.Image = picture1
        '    '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '    '            ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
        '    '                Image1.Image = picture1
        '    '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '    '            ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
        '    '                Image1.Image = picture1
        '    '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '    '            ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
        '    '                Image1.Image = picture1
        '    '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '    '            ElseIf ComboBox2.SelectedItem = "60 [mm]" Then
        '    '                Image1.Image = picture1
        '    '                'Me.'ComboBox3.Items.Add("7 [mm]")
        '    '            End If
        '    '        End If
        '    '        If ComboBox1.SelectedItem = "Nylon Gauge" Then
        '    '            If ComboBox2.SelectedItem = "20 [mm]" Then
        '    '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '    '                Image1.Image = picture2
        '    '            ElseIf ComboBox2.SelectedItem = "30 [mm]" Then
        '    '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '    '                Image1.Image = picture2
        '    '            ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
        '    '                Image1.Image = picture2
        '    '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '    '            ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
        '    '                'Me.'ComboBox3.Items.Add("8 [mm]")
        '    '                Image1.Image = picture2
        '    '            End If
        '    '        End If
        '    '    End With
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
        If Me.ComboBox1.SelectedItem.ToString = "00_Dedecteur_Equerre_01" Then
            'Get CATIA or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\00_Dedecteur_Equerre_01\3D\Dedecteur_Groupe_050918.CATProduct")

            On Error GoTo 0
        ElseIf Me.ComboBox1.SelectedItem.ToString = "01_Detecteur_Equerre_Laser_01" Then
            'Get CATIA or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\01_Detecteur_Equerre_Laser_01\3D\Laser_dedecteur_groupe_01_A.CATProduct")

            On Error GoTo 0
        ElseIf Me.ComboBox1.SelectedItem.ToString = "10_Conception_de_reference_" Then
            'Get CATIA or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\10_Conception_de_reference_\3D\Groupe_fixation_50x50_tube_26,9_pipe.CATProduct")

            On Error GoTo 0
        ElseIf Me.ComboBox1.SelectedItem.ToString = "11_Detecteur_Laser_Accessoires" Then

            If ComboBox2.SelectedItem.ToString = "DEQ_Protection_Laser_01" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Protection_Laser_01.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_Support_11" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Support_11.CATPart")


                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_Support_12" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\11_Detecteur_Laser_Accessoires\DEQ_Support_12.CATPart")

                On Error GoTo 0
            End If
        ElseIf Me.ComboBox1.SelectedItem.ToString = "12_Detecteur_Laser_Support_01" Then

            If ComboBox2.SelectedItem.ToString = "DEQ_Support_01_L100" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\12_Detecteur_Laser_Support_01\DEQ_Support_01_L100.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_Support_02_L200" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\12_Detecteur_Laser_Support_01\DEQ_Support_02_L200.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_Support_03_L300" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\12_Detecteur_Laser_Support_01\DEQ_Support_03_L300.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_Support_04_L400" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\12_Detecteur_Laser_Support_01\DEQ_Support_04_L400.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_Support_05_L500" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\12_Detecteur_Laser_Support_01\DEQ_Support_05_L500.CATPart")

                On Error GoTo 0
            End If

        ElseIf Me.ComboBox1.SelectedItem.ToString = "13_Detecteur_Laser_Pipe" Then
            If ComboBox2.SelectedItem.ToString = "DEQ_26,9_pipe_01_L200" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\13_Detecteur_Laser_Pipe\3D\DEQ_26,9_pipe_01_L200.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_26,9_pipe_02_L350" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\13_Detecteur_Laser_Pipe\3D\DEQ_26,9_pipe_02_L350.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_26,9_pipe_03_L500" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\13_Detecteur_Laser_Pipe\3D\DEQ_26,9_pipe_03_L500.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_26,9_pipe_04_L650" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\13_Detecteur_Laser_Pipe\3D\DEQ_26,9_pipe_04_L650.CATPart")

                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem.ToString = "DEQ_26,9_pipe_05_L800" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\13_Detecteur_Laser_Pipe\3D\DEQ_26,9_pipe_05_L800.CATPart")

                On Error GoTo 0
            End If
        ElseIf Me.ComboBox1.SelectedItem.ToString = "14_Groupe_fixation_50x50_tube_26,9_pipe" Then
            'Get CATIA or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00900_DETECTOR_RFID_EQUIPMENT\14_Groupe_fixation_50x50_tube_26,9_pipe\3D\Groupe_fixation_50x50_tube_26,9_pipe.CATProduct")

            On Error GoTo 0
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '    PictureBox1.Visible = True
    '    Button4.Visible = True
    '    'Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\APPUI_TOUCHE\3D_Serumlu Kesit\Serumlu_Kesit.jpg") '3
    'End Sub

    'Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
    '    PictureBox1.Visible = False
    '    Button4.Visible = False
    'End Sub
End Class
