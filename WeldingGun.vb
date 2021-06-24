Public Class WeldingGun
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
    Dim picture23 As Image
    Dim picture24 As Image
    Dim picture25 As Image
    Dim picture26 As Image
    Dim picture27 As Image
    Dim picture28 As Image
    Dim picture29 As Image
    Dim picture30 As Image
    Dim picture31 As Image
    Dim picture32 As Image
    Dim picture33 As Image
    Dim picture34 As Image
    Dim picture35 As Image
    Dim picture36 As Image
    Dim picture37 As Image
    Dim picture38 As Image
    Dim picture39 As Image
    Dim picture40 As Image
    Dim picture41 As Image
    Dim picture42 As Image
    Dim picture43 As Image
    Dim picture44 As Image
    Private Sub WeldingGun_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Manuel")
        ComboBox1.Items.Add("Robotize")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi000.Show()
        Me.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.ResetText()

        Try
            'Manuel
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X20X095046XZA_ALLCATPART.png") ' 1
            'Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X21X037005XMA_ALLCATPART.png") ' 2
            'Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X27X045004XMA_ALLCATPART.png") ' 3
            'Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X14X055037XPA.png") ' 4
            'Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X03X030017XMA.png") ' 5
            'Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X15X065038XPA_ALLCATPART.png") ' 6
            'Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\J09X019041CMA_ALLCATPART.png") ' 7
            ''Robotise
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J01_R3C10006X01R20_V5.png") ' 8
            'Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J02_R3C10028X02R20_V5.png") ' 9
            'Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J03_R3C10023X01R20_V5.png") ' 10
            'Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J04_R3C10028X01R20_V5.png") ' 11
            'Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J05_R3C10015X01R20_V5.png") ' 12
            'Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J06_R3C10015X02R20_V5.png") ' 13
            'Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J07_R3J10050X01R20_V5.png") ' 14
            'Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J08_R3G07017X01R20_V5.png") ' 15
            'Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X01_R3X90028X01R20_V5.png") ' 16
            'Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X02_R3X10028X03R20_V5.png") ' 17
            'Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X03_R3X10028X01R20_V5.png") ' 18
            'Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X04_R3X10032X01R20_V5.png") ' 19
            'Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X05_R3X10032X02R20_V5.png") ' 20
            'Dim picture21 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X06_R3X10032X03R20_V5.png") ' 21
            'Dim picture22 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X07_R3X10032X06R20_V5.png") ' 22
            'Dim picture23 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X08_R3ZJA045X01R20_V5.png") ' 23
            'Dim picture24 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X09_R3X10040X01R20_V5.png") ' 24
            'Dim picture25 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X10_R3X10040X02R20_V5.png") ' 25
            'Dim picture26 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X11_R3X10040X03R20_V5.png") ' 26
            'Dim picture27 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X12_R3X10040X04R20_V5.png") ' 27
            'Dim picture28 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X13_R3Z10045X01R20_V5.png") ' 28
            'Dim picture29 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X14_R3Z10045X02R20_V5.png") ' 29
            'Dim picture30 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X15_R3Z10055X01R20_V5.png") ' 30
            'Dim picture31 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X16_R3Z10053X01R20_V5.png") ' 31
            'Dim picture32 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X17_R3Z10053X02R20_V5.png") ' 32
            'Dim picture33 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X18_R3Z10045X06R20_V5.png") ' 33
            'Dim picture34 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X19_R3Z10045X07R20_V5.png") ' 34
            'Dim picture35 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X20_R3Z10055X04R20_V5.png") ' 35
            'Dim picture36 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X21_R3Z10050X01R20_V5.png") ' 36
            'Dim picture37 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X22_R3Z10065X01R20_V5.png") ' 37
            'Dim picture38 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X23_R3Z10050X02R20_V5.png") ' 38
            'Dim picture39 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X24_R3Z10050X03R20_V5.png") ' 39
            'Dim picture40 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X25_RWG07073X01R20_V5.png") ' 40
            'Dim picture41 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X26_RWG07090X03R20_V5.png") ' 41
            'Dim picture42 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X27_RWG07090X02R20_V5.png") ' 42
            'Dim picture43 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X28_RWG07090X01R20_V5.png") ' 43
            'Dim picture44 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X29_RWG07120X01R20_V5.png") ' 44



            Me.ComboBox2.Items.Clear()
            If Me.ComboBox1.SelectedItem.ToString = "Manuel" Then
                Me.ComboBox2.Items.Add("X20X095046XZA_ALLCATPART")
                Me.ComboBox2.Items.Add("X21X037005XMA_ALLCATPART")
                Me.ComboBox2.Items.Add("X27X045004XMA_ALLCATPART")
                Me.ComboBox2.Items.Add("X14X055037XPA")
                Me.ComboBox2.Items.Add("X03X030017XMA")
                Me.ComboBox2.Items.Add("X15X065038XPA_ALLCATPART")
                Me.ComboBox2.Items.Add("J09X019041CMA_ALLCATPART")


            ElseIf Me.ComboBox1.SelectedItem.ToString = "Robotize" Then
                Me.ComboBox2.Items.Add("J01_R3C10006X01R20_V5")
                Me.ComboBox2.Items.Add("J02_R3C10028X02R20_V5")
                Me.ComboBox2.Items.Add("J03_R3C10023X01R20_V5")
                Me.ComboBox2.Items.Add("J04_R3C10028X01R20_V5")
                Me.ComboBox2.Items.Add("J05_R3C10015X01R20_V5")
                Me.ComboBox2.Items.Add("J06_R3C10015X02R20_V5")
                Me.ComboBox2.Items.Add("J07_R3J10050X01R20_V5")
                Me.ComboBox2.Items.Add("J08_R3G07017X01R20_V5")
                Me.ComboBox2.Items.Add("X01_R3X90028X01R20_V5")
                Me.ComboBox2.Items.Add("X02_R3X10028X03R20_V5")
                Me.ComboBox2.Items.Add("X03_R3X10028X01R20_V5")
                Me.ComboBox2.Items.Add("X04_R3X10032X01R20_V5")
                Me.ComboBox2.Items.Add("X05_R3X10032X02R20_V5")
                Me.ComboBox2.Items.Add("X06_R3X10032X03R20_V5")
                Me.ComboBox2.Items.Add("X07_R3X10032X06R20_V5")
                Me.ComboBox2.Items.Add("X08_R3ZJA045X01R20_V5")
                Me.ComboBox2.Items.Add("X09_R3X10040X01R20_V5")
                Me.ComboBox2.Items.Add("X10_R3X10040X02R20_V5")
                Me.ComboBox2.Items.Add("X11_R3X10040X03R20_V5")
                Me.ComboBox2.Items.Add("X12_R3X10040X04R20_V5")
                Me.ComboBox2.Items.Add("X13_R3Z10045X01R20_V5")
                Me.ComboBox2.Items.Add("X14_R3Z10045X02R20_V5")
                Me.ComboBox2.Items.Add("X15_R3Z10055X01R20_V5")
                Me.ComboBox2.Items.Add("X16_R3Z10053X01R20_V5")
                Me.ComboBox2.Items.Add("X17_R3Z10053X02R20_V5")
                Me.ComboBox2.Items.Add("X18_R3Z10045X06R20_V5")
                Me.ComboBox2.Items.Add("X19_R3Z10045X07R20_V5")
                Me.ComboBox2.Items.Add("X20_R3Z10055X04R20_V5")
                Me.ComboBox2.Items.Add("X21_R3Z10050X01R20_V5")
                Me.ComboBox2.Items.Add("X22_R3Z10065X01R20_V5")
                Me.ComboBox2.Items.Add("X23_R3Z10050X02R20_V5")
                Me.ComboBox2.Items.Add("X24_R3Z10050X03R20_V5")
                Me.ComboBox2.Items.Add("X25_RWG07073X01R20_V5")
                Me.ComboBox2.Items.Add("X26_RWG07090X03R20_V5")
                Me.ComboBox2.Items.Add("X27_RWG07090X02R20_V5")
                Me.ComboBox2.Items.Add("X28_RWG07090X01R20_V5")
                Me.ComboBox2.Items.Add("X29_RWG07120X01R20_V5")


            End If
            ComboBox1.Focus()
            With Me.ComboBox1
                Select Case ComboBox1.SelectedItem
                    Case "Manuel"
                        Image1.Image = picture1
                    Case "Robotize"
                        Image1.Image = picture8

                End Select
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            'Manuel
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X20X095046XZA_ALLCATPART.png") ' 1
            'Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X21X037005XMA_ALLCATPART.png") ' 2
            'Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X27X045004XMA_ALLCATPART.png") ' 3
            'Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X14X055037XPA.png") ' 4
            'Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X03X030017XMA.png") ' 5
            'Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X15X065038XPA_ALLCATPART.png") ' 6
            'Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\J09X019041CMA_ALLCATPART.png") ' 7
            ''Robotise
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J01_R3C10006X01R20_V5.png") ' 8
            'Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J02_R3C10028X02R20_V5.png") ' 9
            'Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J03_R3C10023X01R20_V5.png") ' 10
            'Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J04_R3C10028X01R20_V5.png") ' 11
            'Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J05_R3C10015X01R20_V5.png") ' 12
            'Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J06_R3C10015X02R20_V5.png") ' 13
            'Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J07_R3J10050X01R20_V5.png") ' 14
            'Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J08_R3G07017X01R20_V5.png") ' 15
            'Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X01_R3X90028X01R20_V5.png") ' 16
            'Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X02_R3X10028X03R20_V5.png") ' 17
            'Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X03_R3X10028X01R20_V5.png") ' 18
            'Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X04_R3X10032X01R20_V5.png") ' 19
            'Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X05_R3X10032X02R20_V5.png") ' 20
            'Dim picture21 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X06_R3X10032X03R20_V5.png") ' 21
            'Dim picture22 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X07_R3X10032X06R20_V5.png") ' 22
            'Dim picture23 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X08_R3ZJA045X01R20_V5.png") ' 23
            'Dim picture24 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X09_R3X10040X01R20_V5.png") ' 24
            'Dim picture25 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X10_R3X10040X02R20_V5.png") ' 25
            'Dim picture26 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X11_R3X10040X03R20_V5.png") ' 26
            'Dim picture27 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X12_R3X10040X04R20_V5.png") ' 27
            'Dim picture28 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X13_R3Z10045X01R20_V5.png") ' 28
            'Dim picture29 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X14_R3Z10045X02R20_V5.png") ' 29
            'Dim picture30 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X15_R3Z10055X01R20_V5.png") ' 30
            'Dim picture31 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X16_R3Z10053X01R20_V5.png") ' 31
            'Dim picture32 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X17_R3Z10053X02R20_V5.png") ' 32
            'Dim picture33 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X18_R3Z10045X06R20_V5.png") ' 33
            'Dim picture34 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X19_R3Z10045X07R20_V5.png") ' 34
            'Dim picture35 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X20_R3Z10055X04R20_V5.png") ' 35
            'Dim picture36 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X21_R3Z10050X01R20_V5.png") ' 36
            'Dim picture37 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X22_R3Z10065X01R20_V5.png") ' 37
            'Dim picture38 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X23_R3Z10050X02R20_V5.png") ' 38
            'Dim picture39 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X24_R3Z10050X03R20_V5.png") ' 39
            'Dim picture40 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X25_RWG07073X01R20_V5.png") ' 40
            'Dim picture41 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X26_RWG07090X03R20_V5.png") ' 41
            'Dim picture42 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X27_RWG07090X02R20_V5.png") ' 42
            'Dim picture43 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X28_RWG07090X01R20_V5.png") ' 43
            'Dim picture44 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X29_RWG07120X01R20_V5.png") ' 44
            ComboBox2.Focus()
            With Me.ComboBox2
                Select Case ComboBox1.SelectedItem
                    Case "Manuel"
                        Select Case ComboBox2.SelectedItem
                            Case "X20X095046XZA_ALLCATPART"
                                Image1.Image = picture1
                            Case "X21X037005XMA_ALLCATPART"
                                'Image1.Image = picture2
                            Case "X27X045004XMA_ALLCATPART"
                                'Image1.Image = picture3
                            Case "X14X055037XPA"
                                'Image1.Image = picture4
                            Case "X03X030017XMA"
                                'Image1.Image = picture5
                            Case "X15X065038XPA_ALLCATPART"
                                'Image1.Image = picture6
                            Case "J09X019041CMA_ALLCATPART"
                                'Image1.Image = picture7
                        End Select

                    Case "Robotize"
                        Select Case ComboBox2.SelectedItem
                            Case "J01_R3C10006X01R20_V5"
                                Image1.Image = picture8
                            Case "J02_R3C10028X02R20_V5"
                                'Image1.Image = picture9
                            Case "J03_R3C10023X01R20_V5"
                                'Image1.Image = picture10
                            Case "J04_R3C10028X01R20_V5"
                                'Image1.Image = picture11
                            Case "J05_R3C10015X01R20_V5"
                                'Image1.Image = picture12
                            Case "J06_R3C10015X02R20_V5"
                                'Image1.Image = picture13
                            Case "J07_R3J10050X01R20_V5"
                                'Image1.Image = picture14
                            Case "J08_R3G07017X01R20_V5"
                                'Image1.Image = picture15
                            Case "X01_R3X90028X01R20_V5"
                                'Image1.Image = picture16
                            Case "X02_R3X10028X03R20_V5"
                                'Image1.Image = picture17
                            Case "X03_R3X10028X01R20_V5"
                                'Image1.Image = picture18
                            Case "X04_R3X10032X01R20_V5"
                                'Image1.Image = picture19
                            Case "X05_R3X10032X02R20_V5"
                                'Image1.Image = picture20
                            Case "X06_R3X10032X03R20_V5"
                                'Image1.Image = picture21
                            Case "X07_R3X10032X06R20_V5"
                                'Image1.Image = picture22
                            Case "X08_R3ZJA045X01R20_V5"
                                'Image1.Image = picture23
                            Case "X09_R3X10040X01R20_V5"
                                'Image1.Image = picture24
                            Case "X10_R3X10040X02R20_V5"
                                'Image1.Image = picture25
                            Case "X11_R3X10040X03R20_V5"
                                'Image1.Image = picture26
                            Case "X12_R3X10040X04R20_V5"
                                'Image1.Image = picture27
                            Case "X13_R3Z10045X01R20_V5"
                                'Image1.Image = picture28
                            Case "X14_R3Z10045X02R20_V5"
                                'Image1.Image = picture29
                            Case "X15_R3Z10055X01R20_V5"
                                'Image1.Image = picture30
                            Case "X16_R3Z10053X01R20_V5"
                                'Image1.Image = picture30
                            Case "X17_R3Z10053X02R20_V5"
                                'Image1.Image = picture31
                            Case "X18_R3Z10045X06R20_V5"
                                'Image1.Image = picture32
                            Case "X19_R3Z10045X07R20_V5"
                                'Image1.Image = picture33
                            Case "X20_R3Z10055X04R20_V5"
                                'Image1.Image = picture34
                            Case "X21_R3Z10050X01R20_V5"
                                'Image1.Image = picture35
                            Case "X22_R3Z10065X01R20_V5"
                                'Image1.Image = picture36
                            Case "X23_R3Z10050X02R20_V5"
                                'Image1.Image = picture37
                            Case "X24_R3Z10050X03R20_V5"
                                'Image1.Image = picture38
                            Case "X25_RWG07073X01R20_V5"
                                'Image1.Image = picture39
                            Case "X26_RWG07090X03R20_V5"
                                'Image1.Image = picture40
                            Case "X27_RWG07090X02R20_V5"
                                'Image1.Image = picture41
                            Case "X28_RWG07090X01R20_V5"
                                'Image1.Image = picture42
                            Case "X29_RWG07120X01R20_V5"
                                'Image1.Image = picture43

                        End Select

                End Select


            End With


        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Select Case ComboBox1.SelectedItem
            Case "Manuel"
                Select Case ComboBox2.SelectedItem
                    Case "X20X095046XZA_ALLCATPART"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X20X095046XZA_ALLCATPART.png")

                        On Error GoTo 0
                    Case "X21X037005XMA_ALLCATPART"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X21X037005XMA_ALLCATPART.png")

                        On Error GoTo 0
                    Case "X27X045004XMA_ALLCATPART"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X27X045004XMA_ALLCATPART.png")

                        On Error GoTo 0
                    Case "X14X055037XPA"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X14X055037XPA.png")

                        On Error GoTo 0
                    Case "X03X030017XMA"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X03X030017XMA.png")

                        On Error GoTo 0
                    Case "X15X065038XPA_ALLCATPART"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\X15X065038XPA_ALLCATPART.png")

                        On Error GoTo 0
                    Case "J09X019041CMA_ALLCATPART"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Manuel\J09X019041CMA_ALLCATPART.png")

                        On Error GoTo 0
                End Select

            Case "Robotize"
                Select Case ComboBox2.SelectedItem
                    Case "J01_R3C10006X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J01_R3C10006X01R20_V5.png")

                        On Error GoTo 0
                    Case "J02_R3C10028X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J02_R3C10028X02R20_V5.png")

                        On Error GoTo 0
                    Case "J03_R3C10023X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J03_R3C10023X01R20_V5.png")

                        On Error GoTo 0
                    Case "J04_R3C10028X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J04_R3C10028X01R20_V5.png")

                        On Error GoTo 0
                    Case "J05_R3C10015X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J05_R3C10015X01R20_V5.png")

                        On Error GoTo 0
                    Case "J06_R3C10015X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J06_R3C10015X02R20_V5.png")

                        On Error GoTo 0
                    Case "J07_R3J10050X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")

                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J07_R3J10050X01R20_V5.png")



                        On Error GoTo 0
                    Case "J08_R3G07017X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\J08_R3G07017X01R20_V5.png")

                        On Error GoTo 0
                    Case "X01_R3X90028X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X01_R3X90028X01R20_V5.png")

                        On Error GoTo 0
                    Case "X02_R3X10028X03R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X02_R3X10028X03R20_V5.png")

                        On Error GoTo 0
                    Case "X03_R3X10028X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X03_R3X10028X01R20_V5.png") ' 18

                        On Error GoTo 0
                    Case "X04_R3X10032X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X04_R3X10032X01R20_V5.png") ' 19
                        On Error GoTo 0
                    Case "X05_R3X10032X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X05_R3X10032X02R20_V5.png") ' 20
                        On Error GoTo 0
                    Case "X06_R3X10032X03R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X06_R3X10032X03R20_V5.png") ' 21


                        On Error GoTo 0
                    Case "X07_R3X10032X06R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X07_R3X10032X06R20_V5.png") ' 22

                        On Error GoTo 0
                    Case "X08_R3ZJA045X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X08_R3ZJA045X01R20_V5.png") ' 23

                        On Error GoTo 0
                    Case "X09_R3X10040X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X09_R3X10040X01R20_V5.png") ' 24

                        On Error GoTo 0
                    Case "X10_R3X10040X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X10_R3X10040X02R20_V5.png") ' 25

                        On Error GoTo 0
                    Case "X11_R3X10040X03R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X11_R3X10040X03R20_V5.png") ' 26

                        On Error GoTo 0
                    Case "X12_R3X10040X04R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X12_R3X10040X04R20_V5.png") ' 27

                        On Error GoTo 0
                    Case "X13_R3Z10045X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X13_R3Z10045X01R20_V5.png") ' 28

                        On Error GoTo 0
                    Case "X14_R3Z10045X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X14_R3Z10045X02R20_V5.png") ' 29

                        On Error GoTo 0
                    Case "X15_R3Z10055X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X15_R3Z10055X01R20_V5.png") ' 30

                        On Error GoTo 0
                    Case "X16_R3Z10053X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X16_R3Z10053X01R20_V5.png") ' 31

                        On Error GoTo 0
                    Case "X17_R3Z10053X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X17_R3Z10053X02R20_V5.png") ' 32

                        On Error GoTo 0
                    Case "X18_R3Z10045X06R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X18_R3Z10045X06R20_V5.png") ' 33

                        On Error GoTo 0
                    Case "X19_R3Z10045X07R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X19_R3Z10045X07R20_V5.png") ' 34

                        On Error GoTo 0
                    Case "X20_R3Z10055X04R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X20_R3Z10055X04R20_V5.png") ' 35

                        On Error GoTo 0
                    Case "X21_R3Z10050X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X21_R3Z10050X01R20_V5.png") ' 36

                        On Error GoTo 0
                    Case "X22_R3Z10065X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X22_R3Z10065X01R20_V5.png") ' 37

                        On Error GoTo 0
                    Case "X23_R3Z10050X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X23_R3Z10050X02R20_V5.png") ' 38

                        On Error GoTo 0
                    Case "X24_R3Z10050X03R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X24_R3Z10050X03R20_V5.png") ' 39

                        On Error GoTo 0
                    Case "X25_RWG07073X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X25_RWG07073X01R20_V5.png") ' 40

                        On Error GoTo 0
                    Case "X26_RWG07090X03R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X26_RWG07090X03R20_V5.png") ' 41

                        On Error GoTo 0
                    Case "X27_RWG07090X02R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X27_RWG07090X02R20_V5.png") ' 42

                        On Error GoTo 0
                    Case "X28_RWG07090X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X28_RWG07090X01R20_V5.png") ' 43

                        On Error GoTo 0
                    Case "X29_RWG07120X01R20_V5"
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\WELDING_GUN (PINCES)\ARO_Pince_Robotise\X29_RWG07120X01R20_V5.png") ' 44

                        On Error GoTo 0

                End Select

        End Select

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class