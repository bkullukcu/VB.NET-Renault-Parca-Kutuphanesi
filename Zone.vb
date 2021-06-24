Public Class Zone
    Private Sub Zone_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Goujon_Support_Group_01")
        ComboBox1.Items.Add("Pince_Cible_Groupe")
        ComboBox1.Items.Add("Goujon_Brush_Support_Groupe")
        ComboBox1.Items.Add("Goujon_Connection_Plate-1")
        ComboBox1.Items.Add("Goujon_Connection_Plate-2")
        ComboBox1.Items.Add("Goujon_Point_0_Groupe")
        ComboBox1.Items.Add("Goujon_Cable_Support_Group_01")
        ComboBox1.Items.Add("NC_Locator_Support_Group_01")
        ComboBox1.Items.Add("Robot_Cabinets_Support_Group_01")
        ComboBox1.Items.Add("Goujon_Cible_Support_01")
        ComboBox1.Items.Add("OPB_Group_01")
        ComboBox1.Items.Add("Pupitre_01")
        ComboBox1.Items.Add("AXE-0_Cable_Support_Group_01")
        ComboBox1.Items.Add("Diversity_Detector_Posts-1")
        ComboBox1.Items.Add("Diversity_Detector_Posts-2")
        ComboBox1.Items.Add("Laser_Support_Vertical_01")
        ComboBox1.Items.Add("Laser_Support_Horizontal_01")
        ComboBox1.Items.Add("Pince_Support_Group_01")
        ComboBox1.Items.Add("Support_Afficheur_01")
        Image1.Image = picture1
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
    Dim picture10 As Image
    Dim picture11 As Image
    Dim picture12 As Image
    Dim picture13 As Image
    Dim picture14 As Image
    Dim picture15 As Image
    Dim picture16 As Image
    Dim picture17 As Image
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\01_Goujon_Support_Group_01.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\02_Pince_Cible_Groupe1.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\03_Goujon_Brush_Support_Groupe1.jpg") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\04_Goujon_Connection_Plates.jpg") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\05_Goujon_Point_0_Groupe_01.jpg") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\06_Goujon_Cable_Support_Group_01.jpg") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\07_NC_Locator_Support_Group_01.jpg") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\08_Robot_Cabinets_Support_Group_01.jpg") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\10_Goujon_Cible_Support_01.jpg") '6
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\11_OPB_Group_01.jpg") '7
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\12_Pupitre_01.jpg") '8
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\14_AXE-0_Cable_Support_Group_01.jpg") '12
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\15_Diversity_Detector_Posts.jpg") '13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\16_Laser_Support_Vertical_01.jpg") '14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\17_Laser_Support_Horizontal_01.jpg") '15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\18_Pince_Support_Group_01.jpg") '16
            Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\19_Support_Afficheur_01.jpg") '17
            Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\04_Goujon_Connection_Plates\3D\Goujon_Connection_Plate_02.jpg") '17

            Select Case ComboBox1.SelectedItem.ToString
                Case "Goujon_Support_Group_01"
                    Image1.Image = picture1

                Case "Pince_Cible_Groupe"
                    Image1.Image = picture2

                Case "Goujon_Brush_Support_Groupe"
                    Image1.Image = picture3

                Case "Goujon_Connection_Plate-1"
                    Image1.Image = picture4
                Case "Goujon_Connection_Plate-2"
                    Image1.Image = picture18

                Case "Goujon_Point_0_Groupe"
                    Image1.Image = picture5

                Case "Goujon_Cable_Support_Group_01"
                    Image1.Image = picture6

                Case "NC_Locator_Support_Group_01"
                    Image1.Image = picture7

                Case "Robot_Cabinets_Support_Group_01"
                    Image1.Image = picture8

                Case "Goujon_Cible_Support_01"
                    Image1.Image = picture9

                Case "OPB_Group_01"
                    Image1.Image = picture10

                Case "Pupitre_01"
                    Image1.Image = picture11

                Case "AXE-0_Cable_Support_Group_01"
                    Image1.Image = picture12

                Case "Diversity_Detector_Posts-1"
                    Image1.Image = picture13
                Case "Diversity_Detector_Posts-2"
                    Image1.Image = picture13

                Case "Laser_Support_Vertical_01"
                    Image1.Image = picture14
                Case "Laser_Support_Horizontal_01"
                    Image1.Image = picture15

                Case "Pince_Support_Group_01"
                    Image1.Image = picture16

                Case "Support_Afficheur_01"
                    Image1.Image = picture17

            End Select

        Catch ex As Exception
            'MsgBox("Bir ya da birkaç fotoğraf eksik")
        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Select Case ComboBox1.SelectedItem.ToString
            Case "Goujon_Support_Group_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\01_Goujon_Support_Group_01\3D\Goujon_Support_Group_01.CATProduct")

                On Error GoTo 0

            Case "Pince_Cible_Groupe"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\02_Pince_Cible_Groupe\3D\PC_Pince_Cible_01.CATProduct")

                On Error GoTo 0
            Case "Goujon_Brush_Support_Groupe"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\03_Goujon_Brush_Support_Groupe\3D\Goujon_Brush_Support_Groupe.CATProduct")

                On Error GoTo 0
            Case "Goujon_Connection_Plate-1"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\04_Goujon_Connection_Plates\3D\Goujon_Connection_Plate_01\Goujon_Connection_Plate_01.CATPart")

                On Error GoTo 0
            Case "Goujon_Connection_Plate-2"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\04_Goujon_Connection_Plates\3D\Goujon_Connection_Plate_02\Goujon_Connection_Plate_02.CATProduct")

                On Error GoTo 0
            Case "Goujon_Point_0_Groupe"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\05_Goujon_Point_0_Groupe\3D\Goujon_Point_0_Groupe_01.CATProduct")

                On Error GoTo 0
            Case "Goujon_Cable_Support_Group_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\06_Goujon_Cable_Support_Group_01\3D\Goujon_Cable_Support_Group_01.CATProduct")

                On Error GoTo 0
            Case "NC_Locator_Support_Group_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\07_NC_Locator_Support_Group_01\3D\NC_Locator_Support_Group_01.CATProduct")

                On Error GoTo 0
            Case "Robot_Cabinets_Support_Group_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\08_Robot_Cabinets_Support_Group_01\3D\Robot_Cabinets_Support_Group_01.CATProduct")

                On Error GoTo 0
            Case "Goujon_Cible_Support_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\10_Goujon_Cible_Support_01\3D\Goujon_Cible_Support_01.CATProduct")

                On Error GoTo 0
            Case "OPB_Group_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\11_OPB_Group_01\3D\OPB_Group_01.CATProduct")

                On Error GoTo 0
            Case "Pupitre_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\12_Pupitre_01\3D\Pupitre_01.CATProduct")

                On Error GoTo 0
            Case "AXE-0_Cable_Support_Group_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\14_AXE-0_Cable_Support_Group_01\3D\AXE-0_Cable_Support_Group_01.CATProduct")

                On Error GoTo 0
            Case "Diversity_Detector_Posts-1"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\15_Diversity_Detector_Posts\3D\Diversity_Detector_Support_01\Diversity_Detector_Support_01.CATProduct")

                On Error GoTo 0
            Case "Diversity_Detector_Posts-2"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\15_Diversity_Detector_Posts\3D\Diversity_Detector_Support_02\Diversity_Detector_Support_02.CATProduct")

                On Error GoTo 0
            Case "Laser_Support_Vertical_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\16_Laser_Support_Vertical_01\3D\Laser_Support_Vertical_Groupe_01.CATProduct")

                On Error GoTo 0
            Case "Laser_Support_Horizontal_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\17_Laser_Support_Horizontal_01\3D\Laser_Support_H_Group_01.CATProduct")

                On Error GoTo 0
            Case "Pince_Support_Group_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\18_Pince_Support_Group_01\3D\Pince_Support_Group_01.CATProduct")

                On Error GoTo 0
            Case "Support_Afficheur_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_03100_ZONE_EQUIPMENT\19_Support_Afficheur_01\3D\Support_Afficheur_01.CATProduct")

                On Error GoTo 0
        End Select
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class