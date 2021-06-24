Public Class Hardware
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
    Dim picture18 As Image
    Dim picture19 As Image
    Dim picture20 As Image
    Dim picture21 As Image
    Dim picture22 As Image
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        ComboBox4.ResetText()
        Me.ComboBox2.Items.Clear()
        Me.ComboBox3.Items.Clear()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985\nut_din_985-M4x0_7-10.png") ' 1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\H\E_H_04.png") ' 2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\Hm\Hm M04.png") ' 3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x10.png") ' 4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_06_NM__.png") ' 5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_08_NM____.png") ' 6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\3699_06_10\3699_06_10.png") ' 7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7060_06_10\7060_06_10.png") ' 8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36010610\3601_06_10.png") ' 9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36040600\3604_06_00.png") ' 10
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 04.png") ' 11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D04.png") ' 12
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\PE\PE_D06-18.png") ' 13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D04.png") ' 14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 04.png") ' 15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Broche\Broche Fixe 06150-116x110.png") ' 16
            Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.png") ' 17
            Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_08__.png") ' 18
            Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-08.png") ' 19
            Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\HC poussoir bille 06x15.png") ' 20
            Dim picture21 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\VIS_HC_05x10_plat.png") ' 21
            Dim picture22 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HCP bout plat\HCP_06_10.png") ' 22
            If ComboBox1.SelectedItem.ToString = "Ecrou" Then
                ComboBox2.Items.Add("DIN 985")
                ComboBox2.Items.Add("H")
                ComboBox2.Items.Add("Hm")
                'ComboBox2.Items.Add("80 [mm]")
                Image1.Image = picture1
                Label4.Visible = False
                ComboBox4.Visible = False
                'Label1.Text = "Ekipman Tipi Seçiniz"
                Label2.Text = "Parça Tipi Seçiniz"
                Label3.Text = "Metrik"
                'Label4.Text = "Ekipman Seçiniz"
                Label3.Visible = True
                ComboBox3.Visible = True
            ElseIf ComboBox1.SelectedItem.ToString = "PIN G_GT" Then
                ComboBox2.Items.Add("4 [mm]")
                ComboBox2.Items.Add("5 [mm]")
                ComboBox2.Items.Add("6 [mm]")
                ComboBox2.Items.Add("8 [mm]")
                ComboBox2.Items.Add("10 [mm]")
                ComboBox2.Items.Add("12 [mm]")
                Label4.Visible = False
                ComboBox4.Visible = False
                Image1.Image = picture4
                'Label1.Text = "Ekipman Tipi Seçiniz"
                Label2.Text = "Çap Seçiniz"
                Label3.Text = "Boy Seçiniz"
                ' Label4.Text = "Ekipman Seçiniz"
                Label3.Visible = True
                ComboBox3.Visible = True
            ElseIf ComboBox1.SelectedItem.ToString = "PRODUCT VIS + RONDELLE" Then
                ComboBox2.Items.Add("M4")
                ComboBox2.Items.Add("M5")
                ComboBox2.Items.Add("M6")
                ComboBox2.Items.Add("M8")
                ComboBox2.Items.Add("M10")
                ComboBox2.Items.Add("M12")
                ComboBox2.Items.Add("M16")
                Image1.Image = picture5
                Label4.Visible = True
                ComboBox4.Visible = True
                ' Label1.Text = "Ekipman Tipi Seçiniz"
                Label2.Text = "Metrik"
                Label3.Text = "Ekipman Tipi Seçiniz"
                Label4.Text = "Ekipman Seçiniz"
            ElseIf ComboBox1.SelectedItem.ToString = "Rakor" Then
                ComboBox2.Items.Add("3699_06_10")
                ComboBox2.Items.Add("3699_06_13")
                ComboBox2.Items.Add("7060_06_10")
                ComboBox2.Items.Add("7060_06_13")
                ComboBox2.Items.Add("7100_08_10")
                ComboBox2.Items.Add("7100_08_13")
                ComboBox2.Items.Add("7100_08_17")
                ComboBox2.Items.Add("36010610")
                ComboBox2.Items.Add("36040600")
                Image1.Image = picture7
                Label4.Visible = False
                ComboBox4.Visible = False
                Label3.Visible = False
                ComboBox3.Visible = False
                ' Label1.Text = "Ekipman Tipi Seçiniz"
                Label2.Text = "Ekipman Seçiniz"
                'Label3.Text = "Metrik"
                'Label4.Text = "Ekipman Seçiniz"
            ElseIf ComboBox1.SelectedItem.ToString = "Rondelle" Then
                ComboBox2.Items.Add("L")
                ComboBox2.Items.Add("M")
                ComboBox2.Items.Add("PE")
                ComboBox2.Items.Add("RK")
                ComboBox2.Items.Add("Z")
                Image1.Image = picture11
                Label4.Visible = False
                ComboBox4.Visible = False
                'Label1.Text = "Ekipman Tipi Seçiniz"
                Label2.Text = "Ekipman Tipi Seçiniz"
                Label3.Text = "Çap Seçiniz"
                ' Label4.Text = "Ekipman Seçiniz"
                Label3.Visible = True
                ComboBox3.Visible = True
            ElseIf ComboBox1.SelectedItem.ToString = "Visserie" Then
                ComboBox2.Items.Add("Broche")
                ComboBox2.Items.Add("Chc")
                ComboBox2.Items.Add("FHC")
                ComboBox2.Items.Add("H")
                ComboBox2.Items.Add("HC Poussoir")
                ComboBox2.Items.Add("HCP Bout Plat")
                Image1.Image = picture16
                Label4.Visible = False
                ComboBox4.Visible = False
                'Label1.Text = "Ekipman Tipi Seçiniz"
                Label2.Text = "Ekipman Tipi Seçiniz"
                Label3.Text = "Metrik"
                ' Label4.Text = "Ekipman Seçiniz"
                Label3.Visible = True
                ComboBox3.Visible = True
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub PilotesMobile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Ecrou")
        ComboBox1.Items.Add("PIN G_GT")
        ComboBox1.Items.Add("PRODUCT VIS + RONDELLE")
        ComboBox1.Items.Add("Rakor")
        ComboBox1.Items.Add("Rondelle")
        ComboBox1.Items.Add("Visserie")
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        ComboBox3.ResetText()
        ComboBox4.ResetText()

        Me.ComboBox3.Items.Clear()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985\nut_din_985-M4x0_7-10.png") ' 1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\H\E_H_04.png") ' 2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\Hm\Hm M04.png") ' 3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x10.png") ' 4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_06_NM__.png") ' 5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_08_NM____.png") ' 6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\3699_06_10\3699_06_10.png") ' 7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7060_06_10\7060_06_10.png") ' 8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36010610\3601_06_10.png") ' 9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36040600\3604_06_00.png") ' 10
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 04.png") ' 11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D04.png") ' 12
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\PE\PE_D06-18.png") ' 13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D04.png") ' 14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 04.png") ' 15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Broche\Broche Fixe 06150-116x110.png") ' 16
            Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.png") ' 17
            Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_08__.png") ' 18
            Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-08.png") ' 19
            Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\HC poussoir bille 06x15.png") ' 20
            Dim picture21 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\VIS_HC_05x10_plat.png") ' 21
            Dim picture22 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HCP bout plat\HCP_06_10.png") ' 22
            If ComboBox1.SelectedItem.ToString = "Ecrou" Then


                If ComboBox2.SelectedItem.ToString = "DIN 985" Then
                    Label3.Text = "Metrik"

                    Image1.Image = picture1
                    ComboBox3.Items.Add("M4")
                    ComboBox3.Items.Add("M5")
                    ComboBox3.Items.Add("M6")
                    ComboBox3.Items.Add("M8")
                    ComboBox3.Items.Add("M10")
                    ComboBox3.Items.Add("M12")
                    ComboBox3.Items.Add("M14")
                    ComboBox3.Items.Add("M16")
                ElseIf ComboBox2.SelectedItem.ToString = "H" Then
                    Label3.Text = "Metrik"

                    Image1.Image = picture2
                    ComboBox3.Items.Add("M4")
                    ComboBox3.Items.Add("M5")
                    ComboBox3.Items.Add("M6")
                    ComboBox3.Items.Add("M8")
                    ComboBox3.Items.Add("M10")
                    ComboBox3.Items.Add("M12")
                    'ComboBox3.Items.Add("M14")
                    ComboBox3.Items.Add("M16")
                ElseIf ComboBox2.SelectedItem.ToString = "Hm" Then
                    Label3.Text = "Metrik"

                    'ElseIf ComboBox2.SelectedItem.ToString ="80 [mm]"
                    Image1.Image = picture3
                    ComboBox3.Items.Add("M4")
                    ComboBox3.Items.Add("M5")
                    ComboBox3.Items.Add("M6")
                    ComboBox3.Items.Add("M8")
                End If
            ElseIf ComboBox1.SelectedItem.ToString = "PIN G_GT" Then
                If ComboBox2.SelectedItem.ToString = "4 [mm]" Then
                    Label3.Text = "Boy Seçiniz"

                    Image1.Image = picture4
                    For i As Integer = 10 To 30 Step 5
                        ComboBox3.Items.Add(i.ToString & " [mm]")
                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "5 [mm]" Then
                    Label3.Text = "Boy Seçiniz"

                    Image1.Image = picture4
                    For i As Integer = 16 To 25 Step 9
                        ComboBox3.Items.Add(i.ToString & " [mm]")
                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "6 [mm]" Then
                    Label3.Text = "Boy Seçiniz"

                    Image1.Image = picture4
                    For i As Integer = 20 To 50 Step 5
                        ComboBox3.Items.Add(i.ToString & " [mm]")
                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "8 [mm]" Then
                    Label3.Text = "Boy Seçiniz"

                    Image1.Image = picture4
                    For i As Integer = 10 To 80 Step 5
                        If i = 15 Or i = 65 Or i = 75 Then

                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")

                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "10 [mm]" Then
                    Label3.Text = "Boy Seçiniz"

                    Image1.Image = picture4
                    For i As Integer = 20 To 80 Step 5
                        If i = 55 Or i = 65 Or i = 60 Or i = 70 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")

                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "12 [mm]" Then
                    Label3.Text = "Boy Seçiniz"

                    Image1.Image = picture4
                    For i As Integer = 40 To 80 Step 10
                        If i = 70 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")

                    Next
                End If
            ElseIf ComboBox1.SelectedItem.ToString = "PRODUCT VIS + RONDELLE" Then
                ComboBox3.Items.Add("CHC")
                ComboBox3.Items.Add("H")
                If ComboBox2.SelectedItem.ToString = "M4" Then
                    Label3.Text = "Ekipman Tipi Seçiniz"

                    Image1.Image = picture5
                ElseIf ComboBox2.SelectedItem.ToString = "M5" Then
                    Image1.Image = picture5
                    Label3.Text = "Ekipman Tipi Seçiniz"


                ElseIf ComboBox2.SelectedItem.ToString = "M6" Then
                    Image1.Image = picture5
                    Label3.Text = "Ekipman Tipi Seçiniz"

                ElseIf ComboBox2.SelectedItem.ToString = "M8" Then
                    Image1.Image = picture5
                    Label3.Text = "Ekipman Tipi Seçiniz"

                ElseIf ComboBox2.SelectedItem.ToString = "M10" Then
                    Image1.Image = picture5
                    Label3.Text = "Ekipman Tipi Seçiniz"

                ElseIf ComboBox2.SelectedItem.ToString = "M12" Then
                    Image1.Image = picture5
                    Label3.Text = "Ekipman Tipi Seçiniz"

                ElseIf ComboBox2.SelectedItem.ToString = "M16" Then
                    Image1.Image = picture5
                    Label3.Text = "Ekipman Tipi Seçiniz"

                End If
            ElseIf ComboBox1.SelectedItem.ToString = "Rakor" Then

                If ComboBox2.SelectedItem.ToString = "3699_06_10" Then
                    Image1.Image = picture7

                ElseIf ComboBox2.SelectedItem.ToString = "3699_06_13" Then
                    Image1.Image = picture7

                ElseIf ComboBox2.SelectedItem.ToString = "7060_06_10" Then
                    Image1.Image = picture8

                ElseIf ComboBox2.SelectedItem.ToString = "7060_06_13" Then
                    Image1.Image = picture8

                ElseIf ComboBox2.SelectedItem.ToString = "7100_08_10" Then
                    Image1.Image = picture8

                ElseIf ComboBox2.SelectedItem.ToString = "7100_08_13" Then
                    Image1.Image = picture8

                ElseIf ComboBox2.SelectedItem.ToString = "7100_08_17" Then
                    Image1.Image = picture8

                ElseIf ComboBox2.SelectedItem.ToString = "36010610" Then
                    Image1.Image = picture9

                ElseIf ComboBox2.SelectedItem.ToString = "36040600" Then
                    Image1.Image = picture10

                End If
            ElseIf ComboBox1.SelectedItem.ToString = "Rondelle" Then

                If ComboBox2.SelectedItem.ToString = "L" Then
                    Label3.Text = "Çap Seçiniz"

                    Image1.Image = picture11
                    For i As Integer = 4 To 12 Step 1
                        If i = 7 Or i = 9 Or i = 11 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")

                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "M" Then
                    Label3.Text = "Çap Seçiniz"

                    Image1.Image = picture12
                    For i As Integer = 4 To 12 Step 1
                        If i = 7 Or i = 9 Or i = 11 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")
                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "PE" Then
                    Label3.Text = "Çap Seçiniz"

                    Image1.Image = picture13
                    For i As Integer = 6 To 10 Step 2
                        If i = 7 Or i = 9 Or i = 11 Or i = 8 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")
                    Next
                    ComboBox3.Items.Add("08-25 [mm]")
                    ComboBox3.Items.Add("08-40 [mm]")

                ElseIf ComboBox2.SelectedItem.ToString = "RK" Then
                    Label3.Text = "Çap Seçiniz"

                    Image1.Image = picture14
                    For i As Integer = 4 To 12 Step 1
                        If i = 7 Or i = 9 Or i = 11 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")
                    Next
                    ComboBox3.Items.Add("16 [mm]")

                ElseIf ComboBox2.SelectedItem.ToString = "Z" Then
                    Label3.Text = "Çap Seçiniz"

                    Image1.Image = picture15
                    For i As Integer = 4 To 12 Step 1
                        If i = 7 Or i = 9 Or i = 11 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add(i.ToString & " [mm]")
                    Next
                End If
            ElseIf ComboBox1.SelectedItem.ToString = "Visserie" Then
                If ComboBox2.SelectedItem.ToString = "Broche" Then
                    Image1.Image = picture16
                    Label3.Text = "Parça Seçin"
                    ComboBox3.Items.Add("Broche Fixe 06150-116x110")
                    ComboBox3.Items.Add("Patin _07140-108")
                    ComboBox3.Items.Add("Patin_07140-116")
                ElseIf ComboBox2.SelectedItem.ToString = "Chc" Then
                    Label3.Text = "Metrik"

                    Image1.Image = picture17
                    For i As Integer = 6 To 20 Step 2
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("4-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 40 Step 5
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("4-0" & i.ToString)
                    Next
                    For i As Integer = 6 To 20 Step 2
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("5-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 50 Step 5
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("5-0" & i.ToString)
                    Next
                    For i As Integer = 10 To 20 Step 2
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("6-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 60 Step 5
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("6-0" & i.ToString)
                    Next
                    For i As Integer = 12 To 20 Step 2
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("8-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 100 Step 5
                        If i = 85 Or i = 95 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("8-0" & i.ToString)
                    Next

                    For i As Integer = 25 To 80 Step 5
                        If i = 75 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("10-0" & i.ToString)
                    Next
                    For i As Integer = 30 To 50 Step 5
                        If i = 85 Or i = 95 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("12-0" & i.ToString)
                    Next
                    For i As Integer = 30 To 50 Step 5
                        If i = 85 Or i = 95 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("16-0" & i.ToString)
                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "FHC" Then
                    Label3.Text = "Metrik"

                    Image1.Image = picture18
                    For i As Integer = 8 To 20 Step 2
                        If i = 14 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("4-0" & i.ToString)
                        ComboBox3.Items.Add("5-0" & i.ToString)
                        ComboBox3.Items.Add("6-0" & i.ToString)

                    Next
                    For i As Integer = 25 To 30 Step 5
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("4-0" & i.ToString)
                        ComboBox3.Items.Add("5-0" & i.ToString)
                        ComboBox3.Items.Add("6-0" & i.ToString)
                    Next
                    For i As Integer = 10 To 20 Step 2
                        If i = 14 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("8-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 30 Step 5
                        'If i = 85 Or i = 95 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("8-0" & i.ToString)
                    Next
                    For i As Integer = 16 To 20 Step 2
                        'If i = 75 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("10-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 35 Step 5
                        'If i = 85 Or i = 95 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("10-0" & i.ToString)
                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "H" Then
                    Label3.Text = "Metrik"

                    Image1.Image = picture19
                    For i As Integer = 8 To 20 Step 4
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("4-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 40 Step 5
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("4-0" & i.ToString)
                    Next
                    For i As Integer = 10 To 20 Step 2
                        If i = 14 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("5-0" & i.ToString)

                    Next
                    For i As Integer = 25 To 50 Step 5
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("5-0" & i.ToString)
                    Next
                    For i As Integer = 12 To 20 Step 4
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("6-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 60 Step 5
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("6-0" & i.ToString)
                    Next
                    For i As Integer = 12 To 20 Step 4
                        'If i = 7 Or i = 9 Or i = 11 Or i = 13 Or i = 14 Or i = 15 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("8-0" & i.ToString)
                    Next
                    For i As Integer = 25 To 100 Step 5
                        If i = 85 Or i = 95 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("8-0" & i.ToString)
                    Next

                    For i As Integer = 20 To 70 Step 5
                        'If i = 75 Then
                        '    Continue For
                        'End If
                        ComboBox3.Items.Add("10-0" & i.ToString)
                    Next
                    For i As Integer = 30 To 50 Step 5
                        If i = 85 Or i = 95 Then
                            Continue For
                        End If
                        ComboBox3.Items.Add("12-0" & i.ToString)
                    Next
                    For i As Integer = 30 To 50 Step 5
                        'If i = 85 Or i = 95 Then
                        '    Continue For   
                        'End If
                        ComboBox3.Items.Add("16-0" & i.ToString)
                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "HC Poussoir" Then
                    Label3.Text = "Metrik"

                    Image1.Image = picture20
                    ComboBox3.Items.Add("5x10")
                    ComboBox3.Items.Add("6x15 Poussoir Bille")
                    ComboBox3.Items.Add("6x15 VIS_HC_Bille")

                ElseIf ComboBox2.SelectedItem.ToString = "HCP Bout Plat" Then
                    Label3.Text = "Metrik"

                    Image1.Image = picture21
                    ComboBox3.Items.Add("6_10")
                    ComboBox3.Items.Add("6_12")
                    ComboBox3.Items.Add("6_16")
                    ComboBox3.Items.Add("6_20")
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox4.ResetText()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985\nut_din_985-M4x0_7-10.png") ' 1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\H\E_H_04.png") ' 2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\Hm\Hm M04.png") ' 3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x10.png") ' 4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_06_NM__.png") ' 5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_08_NM____.png") ' 6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\3699_06_10\3699_06_10.png") ' 7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7060_06_10\7060_06_10.png") ' 8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36010610\3601_06_10.png") ' 9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36040600\3604_06_00.png") ' 10
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 04.png") ' 11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D04.png") ' 12
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\PE\PE_D06-18.png") ' 13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D04.png") ' 14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 04.png") ' 15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Broche\Broche Fixe 06150-116x110.png") ' 16
            Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.png") ' 17
            Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_08__.png") ' 18
            Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-08.png") ' 19
            Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\HC poussoir bille 06x15.png") ' 20
            Dim picture21 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\VIS_HC_05x10_plat.png") ' 21
            Dim picture22 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HCP bout plat\HCP_06_10.png") ' 22
            If ComboBox1.SelectedItem.ToString = "PRODUCT VIS + RONDELLE" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then
                    Image1.Image = picture5
                ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                    Image1.Image = picture6

                End If
                If ComboBox2.SelectedItem.ToString = "M4" Then



                    For i As Integer = 8 To 18 Step 2
                        If i = 14 Or i = 18 Then
                            Continue For
                        End If
                        ComboBox4.Items.Add("0" & i.ToString & "_NM")
                    Next
                    For i As Integer = 20 To 40 Step 5

                        ComboBox4.Items.Add("0" & i.ToString & "_NM")
                    Next
                    For i As Integer = 8 To 18 Step 2
                        If i = 14 Or i = 18 Then
                            Continue For
                        End If
                        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                    Next
                    For i As Integer = 20 To 40 Step 5

                        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                    Next


                ElseIf ComboBox2.SelectedItem.ToString = "M5" Then
                    If ComboBox3.SelectedItem.ToString = "CHC" Then

                        For i As Integer = 8 To 18 Step 2
                            If i = 14 Or i = 18 Then
                                Continue For
                            End If
                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")


                        Next
                        For i As Integer = 20 To 50 Step 5

                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")

                        Next
                        'For i As Integer = 8 To 18 Step 2
                        '    If i = 14 Or i = 18 Then
                        '        Continue For
                        '    End If
                        '    '  ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        'Next
                        'For i As Integer = 20 To 50 Step 5

                        '    'ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        'Next
                    ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                        For i As Integer = 10 To 18 Step 2
                            If i = 14 Or i = 18 Then
                                Continue For
                            End If
                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")

                        Next
                        For i As Integer = 20 To 50 Step 5

                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")

                        Next
                        'For i As Integer = 8 To 18 Step 2
                        '    If i = 14 Or i = 18 Then
                        '        Continue For
                        '    End If
                        '    ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        'Next
                        'For i As Integer = 20 To 50 Step 5

                        '    ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        'Next
                    End If
                ElseIf ComboBox2.SelectedItem.ToString = "M6" Then
                    For i As Integer = 10 To 18 Step 2
                        If i = 14 Or i = 18 Then
                            Continue For
                        End If
                        ComboBox4.Items.Add("0" & i.ToString & "_NM")
                        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                        ComboBox4.Items.Add("0" & i.ToString & "_RKPE")

                    Next
                    For i As Integer = 20 To 60 Step 5

                        ComboBox4.Items.Add("0" & i.ToString & "_NM")
                        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                        ComboBox4.Items.Add("0" & i.ToString & "_RKPE")

                    Next
                    'For i As Integer = 6 To 18 Step 2
                    '    If i = 14 Or i = 18 Then
                    '        Continue For
                    '    End If
                    '    ComboBox4.Items.Add("0" & i.ToString & "_RK")
                    'Next
                    'For i As Integer = 20 To 40 Step 5

                    '    ComboBox4.Items.Add("0" & i.ToString & "_RK")
                    'Next
                ElseIf ComboBox2.SelectedItem.ToString = "M8" Then
                    If ComboBox3.SelectedItem.ToString = "CHC" Then

                        For i As Integer = 12 To 18 Step 2
                            If i = 14 Or i = 18 Then
                                Continue For
                            End If
                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")
                            ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                            ComboBox4.Items.Add("0" & i.ToString & "_RKPE")
                        Next
                        For i As Integer = 20 To 100 Step 5
                            If i = 85 Or i = 95 Then
                                Continue For
                            End If
                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")
                            ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                            ComboBox4.Items.Add("0" & i.ToString & "_RKPE")
                        Next
                        'For i As Integer = 6 To 18 Step 2
                        '    If i = 14 Or i = 18 Then
                        '        Continue For
                        '    End If
                        '    ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        'Next
                        'For i As Integer = 20 To 40 Step 5

                        '    ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        'Next
                    ElseIf ComboBox3.SelectedItem.ToString = "H" Then

                        For i As Integer = 12 To 18 Step 2
                            If i = 14 Or i = 18 Then
                                Continue For
                            End If
                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")
                            ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                            ComboBox4.Items.Add("0" & i.ToString & "_RKPE")
                        Next
                        For i As Integer = 20 To 100 Step 5
                            If i = 85 Or i = 95 Then
                                Continue For
                            End If
                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")
                            ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                            ComboBox4.Items.Add("0" & i.ToString & "_RKPE")
                        Next
                    End If

                ElseIf ComboBox2.SelectedItem.ToString = "M10" Then
                    If ComboBox3.SelectedItem.ToString = "CHC" Then

                        For i As Integer = 20 To 80 Step 5
                            If i = 75 Then
                                Continue For
                            End If
                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")
                            ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                            ComboBox4.Items.Add("0" & i.ToString & "_RKPE")

                        Next
                        'For i As Integer = 6 To 18 Step 2
                        '        If i = 14 Or i = 18 Then
                        '            Continue For
                        '        End If
                        '        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        '    Next
                        '    For i As Integer = 20 To 40 Step 5

                        '        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        '    Next
                    ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                        For i As Integer = 20 To 70 Step 5

                            ComboBox4.Items.Add("0" & i.ToString & "_NM")
                            ComboBox4.Items.Add("0" & i.ToString & "_RK")
                            ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                            ComboBox4.Items.Add("0" & i.ToString & "_RKPE")

                        Next
                    End If
                ElseIf ComboBox2.SelectedItem.ToString = "M12" Then
                    For i As Integer = 30 To 50 Step 5

                        ComboBox4.Items.Add("0" & i.ToString & "_NM")
                        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                        ComboBox4.Items.Add("0" & i.ToString & "_RKPE")

                    Next
                ElseIf ComboBox2.SelectedItem.ToString = "M16" Then
                    For i As Integer = 30 To 50 Step 5

                        ComboBox4.Items.Add("0" & i.ToString & "_NM")
                        ComboBox4.Items.Add("0" & i.ToString & "_RK")
                        ComboBox4.Items.Add("0" & i.ToString & "_NMPE")
                        ComboBox4.Items.Add("0" & i.ToString & "_RKPE")

                    Next
                End If
            End If

        Catch ex As Exception
        End Try


    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi200.Show()
        Me.Close()
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc

        If ComboBox1.SelectedItem.ToString = "Ecrou" Then
            If ComboBox2.SelectedItem.ToString = "DIN 985" Then
                If ComboBox3.SelectedItem.ToString = "M4" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985\nut_din_985-M4x0_7-10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M5" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985\nut_din_985-M5x0_8-10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M6" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985\nut_din_985-M6x1-10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M8" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985\nut_din_985-M8x1_25-10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M10" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985nut_din_985-M10x1_5-10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985nut_din_985-M12x1_75-10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M14" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985nut_din_985-M14x2-10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M16" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\DIN 985nut_din_985-M16x2-10.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "H" Then
                If ComboBox3.SelectedItem.ToString = "M4" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M5" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_05.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M6" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M8" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_08.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M10" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M14" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_14.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M16" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\E_H_16.CATPart")

                    On Error GoTo 0

                End If
            ElseIf ComboBox2.SelectedItem.ToString = "Hm" Then

                If ComboBox3.SelectedItem.ToString = "M4" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\Hm\Hm M04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M5" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\Hm\Hm M05.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M6" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\Hm\Hm M06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "M8" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Ecrou\Hm\Hm M08.CATPart")

                    On Error GoTo 0
                End If
            End If

        ElseIf ComboBox1.SelectedItem.ToString = "PIN G_GT" Then
            If ComboBox2.SelectedItem.ToString = "4 [mm]" Then


                If ComboBox3.SelectedItem.ToString = "10 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "15 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x15.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "20 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x20.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "25 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x25.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "30 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_4\PIN_G_4x30.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "5 [mm]" Then


                If ComboBox3.SelectedItem.ToString = "16 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_5\PIN_GT_5x16.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "25 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_5\PIN_GT_5x25.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "6 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "20 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_6\PIN_GT_6x20.CATPart")

                    On Error GoTo 0

                ElseIf ComboBox3.SelectedItem.ToString = "25 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_6\PIN_GT_6x25.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "30 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_6\PIN_GT_6x30.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "35 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_6\PIN_GT_6x35.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_6\PIN_GT_6x40.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "45 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_6\PIN_GT_6x45.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "50 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_6\PIN_GT_6x50.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "8 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "10 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_8x10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "20 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x20.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "25 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x25.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "30 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x30.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "35 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x35.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x40.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "45 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x45.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "50 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x50.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "55 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x55.CATPartt")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x60.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "70 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x70.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "80 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_8\PIN_GT_8x80.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "10 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "20 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x20.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "25 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x25.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "30 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x30.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "35 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x35.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x40.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "45 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x45.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "50 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x50.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "75 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x75.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "80 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_10\PIN_GT_10x80.CATPart")

                    On Error GoTo 0
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "12 [mm]" Then

                If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_12\PIN_GT_12x40.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "50 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_12\PIN_GT_12x50.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_12\PIN_GT_12x60.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "80 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PIN G_GT\diam_12\PIN_GT_12x80.CATPart")

                    On Error GoTo 0
                End If
            End If

        ElseIf ComboBox1.SelectedItem.ToString = "PRODUCT VIS + RONDELLE" Then
            If ComboBox2.SelectedItem.ToString = "M4" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then
                    If ComboBox4.SelectedItem.ToString = "06_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_06_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "08_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_08_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_10_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_12_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_16_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_20_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_25_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_30_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_35_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_40_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_06_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_08_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_10_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_12_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_16_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_20_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_25_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_30_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\CHC\V_CHC04_40_RK__.CATProduct")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                    If ComboBox4.SelectedItem.ToString = "08_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_08_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_12_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_16_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_20_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_25_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_30_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_35_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_40_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "08_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_08_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_12_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_16_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_20_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_25_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_30_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_35_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M4\H\V_H04_40_NM____.CATProduct")

                        On Error GoTo 0
                    End If
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "M5" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then
                    'If ComboBox4.SelectedItem.ToString = "06_NM" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    If ComboBox4.SelectedItem.ToString = "08_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_08_NM__   .CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_10_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_12_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_16_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_20_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_25_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_30_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_35_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_40_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_45_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_50_NM__.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_06_NM__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_08_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_10_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_12_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_16_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_20_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_25_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_30_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_CHC05_50_RK__.CATProduct")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                    'If ComboBox4.SelectedItem.ToString = "06_NM" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    'If ComboBox4.SelectedItem.ToString = "08_NM" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_CHC05_08_NM__.CATProduct")

                    '    End If
                    '    On Error GoTo 0
                    If ComboBox4.SelectedItem.ToString = "010_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_10_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_12_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_16_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_20_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_25_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_30_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_35_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_40_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_45_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_50_NM____.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\CHC\V_H05_06_NM____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_08_RK____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_10_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_12_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_16_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_20_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_25_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_30_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_35_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_40_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_45_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M5\H\V_H05_50_RK____.CATProduct")

                        On Error GoTo 0
                    End If
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "M6" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then
                    If ComboBox4.SelectedItem.ToString = "010_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_10_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_12_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_16_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_20_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_25_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_30_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_35_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_40_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_45_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_50_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_55_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_60_NM__.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_06_NM__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_08_RK__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_10_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_12_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_16_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_20_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_25_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_30_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_50_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_55_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_60_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_10_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_12_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_16_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_20_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_25_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_30_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_35_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_40_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_45_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_50_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_55_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_60_NMPE.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_06_NM__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_08_RK__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_10_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_12_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_16_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_20_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_25_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_30_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_35_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_40_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_45_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_50_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_55_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_60_RKPE.CATProduct")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                    'If ComboBox4.SelectedItem.ToString = "06_NM" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    'ElseIf ComboBox4.SelectedItem.ToString = "08_NM" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    'ElseIf ComboBox4.SelectedItem.ToString = "010_NM" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    'If ComboBox4.SelectedItem.ToString = "010_NM" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\CHC\V_CHC06_10_NM__.CATProduct")

                    '    End If
                    '    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "012_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_12_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_16_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_20_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_25_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_30_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_35_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_40_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_45_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_50_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_55_NM____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_60_NM____.CATProduct")

                    On Error GoTo 0
                    'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_06_NM____.CATProduct")

                    '    End If
                    '    On Error GoTo 0
                    'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_08_RK____.CATProduct")

                    '    End If
                    '    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_10_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_12_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_16_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_20_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_25_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_30_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_35_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_40_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_45_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_50_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_55_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_60_RK____.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "010_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_10_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "012_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_12_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "016_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_16_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "020_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_20_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "025_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_25_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "030_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_30_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "035_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_35_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "040_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_40_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "045_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_45_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "050_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_50_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "055_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_55_NMPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "060_NMPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_60_NMPE__.CATProduct")

                    On Error GoTo 0
                    'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_06_NM____.CATProduct")

                    '    End If
                    '    On Error GoTo 0
                    'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_08_RK____.CATProduct")

                    '    End If
                    '    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "010_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_10_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "012_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_12_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "016_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_16_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "020_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_20_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "025_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_25_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "030_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_30_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_35_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_40_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_45_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_50_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_55_RKPE__.CATProduct")

                    On Error GoTo 0
                ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M6\H\V_H06_60_RKPE__.CATProduct")

                    On Error GoTo 0
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "M8" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then

                    If ComboBox4.SelectedItem.ToString = "012_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_12_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_16_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_20_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_25_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_30_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_35_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_40_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_45_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_50_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_55_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_60_NM__.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_06_NM__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_08_RK__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_10_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_12_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_16_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_20_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_25_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_30_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_50_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_55_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_60_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_10_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_12_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_16_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_20_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_25_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_30_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_35_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_40_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_45_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_50_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_55_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_60_NMPE.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_06_NM__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_08_RK__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_10_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_12_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_16_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_20_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_25_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_30_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_35_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_40_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_45_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_50_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_55_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\CHC\V_CHC08_60_RKPE.CATProduct")

                        On Error GoTo 0
                    End If

                ElseIf ComboBox3.SelectedItem.ToString = "H" Then

                    If ComboBox4.SelectedItem.ToString = "012_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_12_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_16_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_20_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_25_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_30_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_35_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_40_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_45_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_50_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_55_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_60_NM____.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_06_NM____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_08_RK____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_10_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_12_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_16_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_20_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_25_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_30_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_35_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_40_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_45_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_50_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_55_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_60_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_10_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_12_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_16_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_20_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_25_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_30_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_35_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_40_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_45_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_50_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_55_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_60_NMPE__.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_06_NM____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_08_RK____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_10_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_12_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_16_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_20_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_25_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_30_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_35_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_40_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_45_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_50_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_55_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_60_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_35_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_40_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_45_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_50_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_55_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_60_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_65_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_65_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_65_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_65_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_70_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_70_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_70_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_70_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_75_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_75_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_75_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_75_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_80_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_80_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_80_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_80_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "090_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_90_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "090_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_90_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "090_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_90_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "090_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_90_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "0100_RK" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_100_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "0100_NM" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_100_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "0100_RKPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_100_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "0100_NMPE" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M8\H\V_H08_100_NMPE__.CATProduct")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "M10" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then
                    If ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_20_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_25_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_30_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_35_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_40_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_45_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_50_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_55_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_60_NM__.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" 
                        '    'Get CATIA Or LauncCHC it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_06_NM__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" 
                        '    'Get CATIA Or LauncCHC it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_08_RK__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_10_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_12_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_16_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_20_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_25_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_30_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_50_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_55_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_60_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_10_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_12_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_16_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_20_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_25_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_30_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_35_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_40_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_45_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_50_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_55_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_60_NMPE.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" 
                        '    'Get CATIA Or LauncCHC it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_06_NM__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" 
                        '    'Get CATIA Or LauncCHC it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_08_RK__.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_10_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_12_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_16_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_20_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_25_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_30_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_35_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_40_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_45_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_50_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_55_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_60_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_35_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_40_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_45_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_50_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_55_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_60_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_65_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_65_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_65_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_65_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_70_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_70_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_70_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_70_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_75_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_75_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_75_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "075_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_75_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_RKPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_80_RKPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_NMPE" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_80_NMPE.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_80_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "080_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\CHC\V_CHC10_80_NM__.CATProduct")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem.ToString = "H" Then

                    If ComboBox4.SelectedItem.ToString = "020_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_20_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_25_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_30_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_35_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_40_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_45_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_50_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_55_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_60_NM____.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" 
                        '    'Get CATIA Or LauncH it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_06_NM____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" 
                        '    'Get CATIA Or LauncH it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_08_RK____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_10_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_12_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_16_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_20_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_25_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_30_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_35_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_40_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_45_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_50_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_55_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_60_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_10_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_12_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_16_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_20_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_25_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_30_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_35_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_40_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_45_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_50_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_55_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_60_NMPE__.CATProduct")

                        On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "06_RK" 
                        '    'Get CATIA Or LauncH it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_06_NM____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                        'ElseIf ComboBox4.SelectedItem.ToString = "08_RK" 
                        '    'Get CATIA Or LauncH it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    If CATIA Is Nothing 
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_08_RK____.CATProduct")

                        '    End If
                        '    On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "010_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_10_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "012_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_12_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "016_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_16_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "020_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_20_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "025_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_25_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_30_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_35_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_40_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_45_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_50_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_55_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_60_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_35_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_40_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_45_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_50_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_55_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_60_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_65_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_65_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_65_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "065_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_65_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_RKPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_70_RKPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_NMPE" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_70_NMPE__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_70_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "070_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M10\H\V_H10_70_NM____.CATProduct")

                        On Error GoTo 0

                    End If
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "M12" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then
                    If ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_30_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_50_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_55_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_60_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_50_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_30_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_35_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_40_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_45_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_50_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_55_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_60_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_35_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_40_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_45_NM__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\CHC\V_CHC12_50_NM__.CATProduct")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                    If ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_30_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_35_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_40_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_45_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_50_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_55_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_60_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_35_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_40_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_45_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_50_RK____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "030_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_30_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_35_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_40_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_45_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_50_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "055_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_55_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "060_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_60_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "035_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_35_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_40_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_45_NM____.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_NM" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M12\H\V_H12_50_NM____.CATProduct")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "M16" Then
                If ComboBox3.SelectedItem.ToString = "CHC" Then
                    If ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\CHC\V_CHC16_30_RK__.CATProduct")

                        On Error GoTo 0


                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\CHC\V_CHC16_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\CHC\V_CHC16_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\CHC\V_CHC16_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncCHC it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\CHC\V_CHC16_50_RK__.CATProduct")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem.ToString = "H" Then
                    If ComboBox4.SelectedItem.ToString = "030_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\H\V_H16_30_RK__.CATProduct")

                        On Error GoTo 0


                    ElseIf ComboBox4.SelectedItem.ToString = "035_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\H\V_H16_35_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "040_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\H\V_H16_40_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "045_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\H\V_H16_45_RK__.CATProduct")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "050_RK" Then
                        'Get CATIA Or LauncH it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\PRODUCT VIS+RONDELLE\M16\H\V_H16_50_RK__.CATProduct")

                        On Error GoTo 0
                    End If
                End If
            End If
        ElseIf ComboBox1.SelectedItem.ToString = "Rakor" Then

            If ComboBox2.SelectedItem.ToString = "3699_06_10" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\3699_06_10\3699_06_10.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "3699_06_13" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\3699_06_13\3699_06_13.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "7060_06_10" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7060_06_10\7060_06_10.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "7060_06_13" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7060_06_13\7060_06_13.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "7100_08_10" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7100_08_10\7100_08_10.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "7100_08_13" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7100_08_13\7100_08_13.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "7100_08_17" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\7100_08_17\7100_08_17.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "36010610" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36010610\3601_06_10.CATPart")

                On Error GoTo 0

            ElseIf ComboBox2.SelectedItem.ToString = "36040600" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rakor\36040600\3604_06_00.CATPart")

                On Error GoTo 0

            End If
        ElseIf ComboBox1.SelectedItem.ToString = "Rondelle" Then

            If ComboBox2.SelectedItem.ToString = "L" Then


                If ComboBox3.SelectedItem.ToString = "4 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 05.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 08.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\L\L 12.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "M" Then


                If ComboBox3.SelectedItem.ToString = "4 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D05.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D08.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\M\NOMEL_D12.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "PE" Then


                If ComboBox3.SelectedItem.ToString = "6 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\PE\PE_D06-18.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "08-25 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\PE\PE_D08-25.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\PE\PE_D10-28.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "08-40[mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\PE\PE_D08-40.CATPart")

                    On Error GoTo 0
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "RK" Then

                If ComboBox3.SelectedItem.ToString = "4 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D05.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D08.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D12.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "13 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "14 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "15 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\RK\RK_D16.CATPart")

                    On Error GoTo 0
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "Z" Then


                If ComboBox3.SelectedItem.ToString = "4 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 05.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 08.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Rondelle\Z\Z 12.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "13 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "14 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "15 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                    On Error GoTo 0


                End If
            End If
        ElseIf ComboBox1.SelectedItem.ToString = "Visserie" Then
            If ComboBox2.SelectedItem.ToString = "Broche" Then

                If ComboBox3.SelectedItem.ToString = "Broche Fixe 06150-116x110" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Broche\Broche Fixe 06150-116x110.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "Patin _07140-108" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Broche\Patin pour vis broche\Patin _07140-108.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "Patin_07140-116" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Broche\Patin pour vis broche\Patin _07140-116.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "Chc" Then
                If ComboBox3.SelectedItem.ToString = "4-06" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-08" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-014" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-06" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-08" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-014" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-014" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-055" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-060" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-014" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-055" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-060" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-065" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-070" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-075" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-080" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-090" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-0100" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-055" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-060" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-065" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-070" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-080" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\Chc\CHC_04-06.CATPart")

                    On Error GoTo 0

                End If
            ElseIf ComboBox2.SelectedItem.ToString = "FHC" Then

                If ComboBox3.SelectedItem.ToString = "4-08" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_08__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-08" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_08__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-08" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_08__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_10__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_10__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_10__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_12__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_12__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_12__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_16__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_16__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_16__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_18__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_18__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_18__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_20__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_20__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_20__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_08__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_25__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_25__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M4\V_TFHC_M04_30__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M5\V_TFHC_M05_30__.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M6\V_TFHC_M06_30__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M8\V_TFHC_M08_10__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M8\V_TFHC_M08_12__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M8\V_TFHC_M08_16__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M8\V_TFHC_M08_18__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M8\V_TFHC_M08_20__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M8\V_TFHC_M08_25__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M8\V_TFHC_M08_30__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M10\V_TFHC_M10_16__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M10\V_TFHC_M10_18__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M10\V_TFHC_M10_20__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M10\V_TFHC_M10_25__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M10\V_TFHC_M10_30__.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\FHC\M10\V_TFHC_M10_35__.CATPart")
                    On Error GoTo 0
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "H" Then

                If ComboBox3.SelectedItem.ToString = "4-08" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-08.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-12.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-16.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-20.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-25.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-30.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-35.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "4-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-40.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-010" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_04-10.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-12.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-16.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-018" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-18.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-20.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-25.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-30.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-35.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-40.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-45.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "5-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_05-50.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-12.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-16.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-20.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-25.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-30.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-35.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-40.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-45.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-50.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-055" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-55.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6-060" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_06-60.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-012" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-12.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-016" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-16.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-20.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-25.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-30.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-35.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-40.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-45.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-50.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-055" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-55.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-060" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-60.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-065" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-65.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-070" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-70.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-075" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-75.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-080" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-80.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-090" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-90.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "8-0100" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_08-100.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-020" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-20.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-025" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-25.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-30.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-35.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-40.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-45.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-50.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-055" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-55.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-060" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-60.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-065" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-65.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "10-070" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_10-70.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_12-30.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_12-35.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_12-40.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_12-45.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "12-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_12-50.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-030" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_16-30.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-035" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_16-35.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-040" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_16-40.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-045" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_16-45.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "16-050" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\H\H_16-50.CATPart")
                    On Error GoTo 0
                End If

            ElseIf ComboBox2.SelectedItem.ToString = "HCP Bout Plat" Then

                If ComboBox3.SelectedItem.ToString = "6_10" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HCP bout plat\HCP_06_10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6_12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HCP bout plat\HCP_06_12.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6_16" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HCP bout plat\HCP_06_16.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6_20" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HCP bout plat\HCP_06_20.CATPart")

                    On Error GoTo 0

                End If

            ElseIf ComboBox2.SelectedItem.ToString = "HC Poussoir" Then
                If ComboBox3.SelectedItem.ToString = "5x10" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\VIS_HC_bille_M5x10.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6x15 Poussoir Bille" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\HC poussoir bille 06x15.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "6x15 VIS_HC_Bille" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\HARDWARE (PIM-VIS)\Visserie\HC poussoir\A bille\VIS_HC_05x10_plat.CATPart")

                    On Error GoTo 0
                End If


        End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class