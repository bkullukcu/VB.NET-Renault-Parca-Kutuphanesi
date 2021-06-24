Public Class Cylinder
    'Accessory
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
    'Cylinder
    Dim picture13 As Image
    Dim picture14 As Image
    Dim picture15 As Image
    Dim picture16 As Image
    Private Sub Cylinder_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("SMC")
        ComboBox2.Items.Add("Accessory")
        ComboBox2.Items.Add("Cylinder")
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            'Accessory
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS.png") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E.png") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES.png") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F.png") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JA.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ.png") '9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L.png") '10
            'Cylinder
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32.PNG") '11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40.PNG") '12
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50.PNG") '13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63.PNG") '14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80.PNG") '15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100.PNG") '16

            ComboBox2.Focus()
            ComboBox4.ResetText()
            ComboBox3.ResetText()
            Me.ComboBox4.Items.Clear()
            Me.ComboBox3.Items.Clear()
            Select Case ComboBox2.SelectedItem
                Case "Cylinder"
                    Image1.Image = picture13
                    Label1.Text = "Parça Çapını Seçiniz"
                    Label4.Text = "Strok Seçiniz"
                    ComboBox3.Items.Add("32 [mm]")
                    ComboBox3.Items.Add("40 [mm]")
                    ComboBox3.Items.Add("50 [mm]")
                    ComboBox3.Items.Add("63 [mm]")
                    ComboBox3.Items.Add("80 [mm]")
                    ComboBox3.Items.Add("100 [mm]")
                Case "Accessory"
                    Image1.Image = picture1
                    ComboBox3.Items.Add("C")
                    ComboBox3.Items.Add("D")
                    ComboBox3.Items.Add("DS")
                    ComboBox3.Items.Add("E")
                    ComboBox3.Items.Add("ES")
                    ComboBox3.Items.Add("F")
                    ComboBox3.Items.Add("GKM")
                    ComboBox3.Items.Add("JA")
                    ComboBox3.Items.Add("KJ")
                    ComboBox3.Items.Add("L")
                    Label1.Text = "Ekipman Tipini Seçiniz"
                    Label4.Text = "Parça Çapını Seçiniz"
            End Select
            With Me.ComboBox2
                Select Case ComboBox2.SelectedItem
                    Case "Accessory"
                        Image1.Image = picture1
                    Case "Cylinder"
                        If ComboBox3.SelectedItem = "32 [mm]" Then
                            Image1.Image = picture11
                        ElseIf ComboBox3.SelectedItem = "40 [mm]" Then
                            Image1.Image = picture12
                        ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox3.SelectedItem = "63 [mm]" Then
                            Image1.Image = picture14
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture15
                        ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                            Image1.Image = picture16

                        End If
                End Select

            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox4.ResetText()
        Me.ComboBox4.Items.Clear()
        Try
            'Accessory
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS.png") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E.png") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES.png") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F.png") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JA.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ.png") '9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L.png") '10
            'Cylinder
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32.PNG") '11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40.PNG") '12
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50.PNG") '13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63.PNG") '14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80.PNG") '15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100.PNG") '16



            ComboBox3.Focus()
            With Me.ComboBox3
                Select Case ComboBox2.SelectedItem
                    Case "Accessory"


                        If ComboBox3.SelectedItem = "C" Then
                            Image1.Image = picture1

                            ComboBox4.Items.Add("32 [mm]")
                            ComboBox4.Items.Add("40 [mm]")

                            ComboBox4.Items.Add("50 [mm]")

                            ComboBox4.Items.Add("63 [mm]")

                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")
                        ElseIf ComboBox3.SelectedItem = "D" Then
                            Image1.Image = picture2

                            ComboBox4.Items.Add("32 [mm]")
                            ComboBox4.Items.Add("40 [mm]")

                            ComboBox4.Items.Add("50 [mm]")

                            ComboBox4.Items.Add("63 [mm]")

                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")
                        ElseIf ComboBox3.SelectedItem = "DS" Then
                            Image1.Image = picture3

                            ComboBox4.Items.Add("32 [mm]")
                            ComboBox4.Items.Add("40 [mm]")

                            ComboBox4.Items.Add("50 [mm]")

                            ComboBox4.Items.Add("63 [mm]")

                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")
                        ElseIf ComboBox3.SelectedItem = "E" Then
                            Image1.Image = picture4

                            ComboBox4.Items.Add("32 [mm]")
                            ComboBox4.Items.Add("40 [mm]")

                            ComboBox4.Items.Add("50 [mm]")

                            ComboBox4.Items.Add("63 [mm]")

                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")
                        ElseIf ComboBox3.SelectedItem = "ES" Then
                            Image1.Image = picture5

                            ComboBox4.Items.Add("32 [mm]")
                            ComboBox4.Items.Add("40 [mm]")

                            ComboBox4.Items.Add("50 [mm]")

                            ComboBox4.Items.Add("63 [mm]")

                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")
                        ElseIf ComboBox3.SelectedItem = "F" Then
                            Image1.Image = picture6

                            ComboBox4.Items.Add("32 [mm]")
                            ComboBox4.Items.Add("40 [mm]")

                            ComboBox4.Items.Add("50 [mm]")

                            ComboBox4.Items.Add("63 [mm]")

                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")
                        ElseIf ComboBox3.SelectedItem = "GKM" Then
                            Image1.Image = picture7
                            ComboBox4.Items.Add("10-20")

                            ComboBox4.Items.Add("12-24")

                            ComboBox4.Items.Add("16-32")

                            ComboBox4.Items.Add("20-40")

                            ComboBox4.Items.Add("30-54")
                        ElseIf ComboBox3.SelectedItem = "JA" Then
                            Image1.Image = picture8
                            ComboBox4.Items.Add("10-125")

                            ComboBox4.Items.Add("12-125")

                            ComboBox4.Items.Add("16-150")

                            ComboBox4.Items.Add("27-200")

                            ComboBox4.Items.Add("20-150")
                        ElseIf ComboBox3.SelectedItem = "KJ" Then
                            ComboBox4.Items.Add("10")

                            ComboBox4.Items.Add("12")

                            ComboBox4.Items.Add("16")
                            ComboBox4.Items.Add("20")
                            ComboBox4.Items.Add("27")

                            Image1.Image = picture9
                        ElseIf ComboBox3.SelectedItem = "L" Then
                            Image1.Image = picture10

                            ComboBox4.Items.Add("32 [mm]")
                            ComboBox4.Items.Add("40 [mm]")

                            ComboBox4.Items.Add("50 [mm]")

                            ComboBox4.Items.Add("63 [mm]")

                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")
                        End If
                    Case "Cylinder"
                        If ComboBox3.SelectedItem = "32 [mm]" Then
                            ComboBox4.Items.Add("50 [mm]")
                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")

                            ComboBox4.Items.Add("160 [mm]")
                        ElseIf ComboBox3.SelectedItem = "40 [mm]" Then
                            ComboBox4.Items.Add("50 [mm]")
                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")

                            ComboBox4.Items.Add("160 [mm]")
                        ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                            ComboBox4.Items.Add("50 [mm]")
                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")

                            ComboBox4.Items.Add("160 [mm]")

                            ComboBox4.Items.Add("200 [mm]")

                            ComboBox4.Items.Add("250 [mm]")

                            ComboBox4.Items.Add("320 [mm]")
                        ElseIf ComboBox3.SelectedItem = "63 [mm]" Then
                            ComboBox4.Items.Add("50 [mm]")
                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")

                            ComboBox4.Items.Add("160 [mm]")

                            ComboBox4.Items.Add("200 [mm]")

                            ComboBox4.Items.Add("250 [mm]")

                            ComboBox4.Items.Add("300 [mm]")
                            ComboBox4.Items.Add("320 [mm]")

                            ComboBox4.Items.Add("350 [mm]")

                            ComboBox4.Items.Add("400 [mm]")
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            ComboBox4.Items.Add("50 [mm]")
                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")

                            ComboBox4.Items.Add("160 [mm]")

                            ComboBox4.Items.Add("200 [mm]")

                            ComboBox4.Items.Add("250 [mm]")


                            ComboBox4.Items.Add("320 [mm]")

                            ComboBox4.Items.Add("400 [mm]")
                            ComboBox4.Items.Add("500 [mm]")
                            ComboBox4.Items.Add("600 [mm]")
                            ComboBox4.Items.Add("700 [mm]")
                        ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                            ComboBox4.Items.Add("50 [mm]")
                            ComboBox4.Items.Add("80 [mm]")

                            ComboBox4.Items.Add("100 [mm]")

                            ComboBox4.Items.Add("125 [mm]")

                            ComboBox4.Items.Add("160 [mm]")

                            ComboBox4.Items.Add("200 [mm]")

                            ComboBox4.Items.Add("250 [mm]")


                            ComboBox4.Items.Add("320 [mm]")

                            ComboBox4.Items.Add("400 [mm]")
                            ComboBox4.Items.Add("500 [mm]")
                            ComboBox4.Items.Add("600 [mm]")
                            ComboBox4.Items.Add("650 [mm]")
                            ComboBox4.Items.Add("700 [mm]")
                            ComboBox4.Items.Add("800 [mm]")
                            ComboBox4.Items.Add("900 [mm]")
                        End If



                End Select
            End With

        Catch ex As Exception
        End Try

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C.png") '1

            'Me.ComboBox2.Items.Clear()
            Image1.Image = picture1
            Me.ComboBox3.Items.Clear()
            Me.ComboBox4.Items.Clear()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            'Accessory
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS.png") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E.png") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES.png") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F.png") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JA.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ.png") '9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L.png") '10
            'Cylinder
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32.PNG") '11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40.PNG") '12
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50.PNG") '13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63.PNG") '14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80.PNG") '15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100.PNG") '16
        Catch ex As Exception
            ComboBox3.Focus()

            With Me.ComboBox2
                Select Case ComboBox2.SelectedItem
                    Case "Accesory"
                        If ComboBox4.SelectedItem = "32 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                            Image1.Image = picture13
                        End If

                    Case "Cylinder"
                        If ComboBox4.SelectedItem = "50 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "200 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "250 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "300 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "320 [mm]" Then
                            Image1.Image = picture13

                        ElseIf ComboBox4.SelectedItem = "350 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "400 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "500 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "600 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "700 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "800 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox4.SelectedItem = "900 [mm]" Then
                            Image1.Image = picture13
                        End If
                End Select
            End With
        End Try
    End Sub

    Private Sub Image1_Click(sender As Object, e As EventArgs) Handles Image1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi200.Show()
        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Select Case ComboBox2.SelectedItem
            Case "Accessory"
                If ComboBox3.SelectedItem = "C" Then
                    If ComboBox4.SelectedItem = "32 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C5040.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C5050.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C5063.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C5080.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C5100.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\C\C5125.stp")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "D" Then
                    If ComboBox4.SelectedItem = "32 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D5040.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D5050.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D5063.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D5080.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D5100.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\D\D5032.stp")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "DS" Then
                    If ComboBox4.SelectedItem = "32 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS5040.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS5050.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS5063.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS5080.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS5100.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\DS\DS5125.stp")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "E" Then
                    If ComboBox4.SelectedItem = "32 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E5040.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E5050.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E5063.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E5080.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E5100.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\E\E5125.stp")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "ES" Then
                    If ComboBox4.SelectedItem = "32 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES5040.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES5050.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES5063.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES5080.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES5100.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\ES\ES5125.stp")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "F" Then
                    If ComboBox4.SelectedItem = "32 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\F\F5032.stp")

                        On Error GoTo 0
                    End If

                ElseIf ComboBox3.SelectedItem = "GKM" Then
                    If ComboBox4.SelectedItem = "10-20" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM10-20.stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "12-24" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM12-24.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "16-32" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM16-32.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "20-40" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM20-40.stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "30-54" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\GKM\GKM30-54.stp")


                        On Error GoTo 0

                    End If

                ElseIf ComboBox3.SelectedItem = "JA" Then
                    If ComboBox4.SelectedItem = "10-125" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JA30-10-125(0).stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "12-125" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JA40-12-125(0).stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "16-150" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JA50-16-150(0).stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "27-200" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JA125-27-200(0).stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "20-150" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\JA\JAH50-20-150(0).stp")


                        On Error GoTo 0

                    End If
                ElseIf ComboBox3.SelectedItem = "KJ" Then
                    If ComboBox4.SelectedItem = "10" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ10D_D.stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "12" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ12D_D.stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "16" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ16D_D.stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "20" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ20D_D.stp")


                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "27" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\KJ\KJ27D_D.stp")


                        On Error GoTo 0

                    End If
                ElseIf ComboBox3.SelectedItem = "L" Then
                    If ComboBox4.SelectedItem = "32 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L5032.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "40 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L5040.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L5050.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "63 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L5063.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L5080.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L5100.stp")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\ACCESSORY\L\L5125.stp")

                        On Error GoTo 0
                    End If
                End If

                'If ComboBox3.SelectedItem = "32 [mm]" Then
                '    ComboBox4.Items.Add("50 [mm]")
                '    ComboBox4.Items.Add("80 [mm]")

                '    ComboBox4.Items.Add("100 [mm]")

                '    ComboBox4.Items.Add("125 [mm]")

                '    ComboBox4.Items.Add("160 [mm]")
                'ElseIf ComboBox3.SelectedItem = "40 [mm]" Then
                '    ComboBox4.Items.Add("50 [mm]")
                '    ComboBox4.Items.Add("80 [mm]")

                '    ComboBox4.Items.Add("100 [mm]")

                '    ComboBox4.Items.Add("125 [mm]")

                '    ComboBox4.Items.Add("160 [mm]")
                'ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                '    ComboBox4.Items.Add("50 [mm]")
                '    ComboBox4.Items.Add("80 [mm]")

                '    ComboBox4.Items.Add("100 [mm]")

                '    ComboBox4.Items.Add("125 [mm]")

                '    ComboBox4.Items.Add("160 [mm]")

                '    ComboBox4.Items.Add("200 [mm]")

                '    ComboBox4.Items.Add("250 [mm]")

                '    ComboBox4.Items.Add("320 [mm]")
                'ElseIf ComboBox3.SelectedItem = "63 [mm]" Then
                '    ComboBox4.Items.Add("50 [mm]")
                '    ComboBox4.Items.Add("80 [mm]")

                '    ComboBox4.Items.Add("100 [mm]")

                '    ComboBox4.Items.Add("125 [mm]")

                '    ComboBox4.Items.Add("160 [mm]")

                '    ComboBox4.Items.Add("200 [mm]")

                '    ComboBox4.Items.Add("250 [mm]")

                '    ComboBox4.Items.Add("300 [mm]")
                '    ComboBox4.Items.Add("320 [mm]")

                '    ComboBox4.Items.Add("350 [mm]")

                '    ComboBox4.Items.Add("400 [mm]")
                'ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                '    ComboBox4.Items.Add("50 [mm]")
                '    ComboBox4.Items.Add("80 [mm]")

                '    ComboBox4.Items.Add("100 [mm]")

                '    ComboBox4.Items.Add("125 [mm]")

                '    ComboBox4.Items.Add("160 [mm]")

                '    ComboBox4.Items.Add("200 [mm]")

                '    ComboBox4.Items.Add("250 [mm]")


                '    ComboBox4.Items.Add("320 [mm]")

                '    ComboBox4.Items.Add("400 [mm]")
                '    ComboBox4.Items.Add("500 [mm]")
                '    ComboBox4.Items.Add("600 [mm]")
                '    ComboBox4.Items.Add("700 [mm]")
                'ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                '    ComboBox4.Items.Add("50 [mm]")
                '    ComboBox4.Items.Add("80 [mm]")

                '    ComboBox4.Items.Add("100 [mm]")

                '    ComboBox4.Items.Add("125 [mm]")

                '    ComboBox4.Items.Add("160 [mm]")

                '    ComboBox4.Items.Add("200 [mm]")

                '    ComboBox4.Items.Add("250 [mm]")


                '    ComboBox4.Items.Add("320 [mm]")

                '    ComboBox4.Items.Add("400 [mm]")
                '    ComboBox4.Items.Add("500 [mm]")
                '    ComboBox4.Items.Add("600 [mm]")
                '    ComboBox4.Items.Add("650 [mm]")
                '    ComboBox4.Items.Add("700 [mm]")
                '    ComboBox4.Items.Add("800 [mm]")
                '    ComboBox4.Items.Add("900 [mm]")
                'End If
            Case "Cylinder"
                If ComboBox3.SelectedItem = "32 [mm]" Then
                    If ComboBox4.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32_50_QAN0123.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32_80_QAN0124.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32_100_QAN0125.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32_125_QAN0126.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB32\CP95NDB32_160_QAN0127.CATPart")

                        On Error GoTo 0

                    ElseIf ComboBox3.SelectedItem = "40 [mm]" Then
                        If ComboBox4.SelectedItem = "50 [mm]" Then
                            'Get CATIA Or Launch it if necessary.
                            On Error Resume Next
                            CATIA = GetObject(, "CATIA.Application")

                            CATIA = CreateObject("CATIA.Application")
                            CATIA.Visible = True
                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40_50_QAN0128.CATPart")

                            On Error GoTo 0
                        ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                            'Get CATIA Or Launch it if necessary.
                            On Error Resume Next
                            CATIA = GetObject(, "CATIA.Application")

                            CATIA = CreateObject("CATIA.Application")
                            CATIA.Visible = True
                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40_80_QAN0129.CATPart")

                            On Error GoTo 0
                        ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                            'Get CATIA Or Launch it if necessary.
                            On Error Resume Next
                            CATIA = GetObject(, "CATIA.Application")

                            CATIA = CreateObject("CATIA.Application")
                            CATIA.Visible = True
                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40_100_QAN0130.CATPart")

                            On Error GoTo 0
                        ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                            'Get CATIA Or Launch it if necessary.
                            On Error Resume Next
                            CATIA = GetObject(, "CATIA.Application")

                            CATIA = CreateObject("CATIA.Application")
                            CATIA.Visible = True
                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40_125_QAN0131.CATPart")

                            On Error GoTo 0
                        ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                            'Get CATIA Or Launch it if necessary.
                            On Error Resume Next
                            CATIA = GetObject(, "CATIA.Application")

                            CATIA = CreateObject("CATIA.Application")
                            CATIA.Visible = True
                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB40\CP95NDB40_160_QAN0132.CATPart")

                            On Error GoTo 0

                        ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                            If ComboBox4.SelectedItem = "50 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_50_QAN0134.CATPart")

                                On Error GoTo 0
                            ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_80_QAN0137.CATPart")

                                On Error GoTo 0
                            ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_100_QAN0138.CATPart")

                                On Error GoTo 0
                            ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_125_QAN0139.CATPart")

                                On Error GoTo 0
                            ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_160_QAN0140.CATPart")

                                On Error GoTo 0

                            ElseIf ComboBox4.SelectedItem = "200 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_200_QAN0192.CATPart")

                                On Error GoTo 0
                            ElseIf ComboBox4.SelectedItem = "250 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_250_QAN0193.CATPart")

                                On Error GoTo 0

                            ElseIf ComboBox4.SelectedItem = "320 [mm]" Then
                                'Get CATIA Or Launch it if necessary.
                                On Error Resume Next
                                CATIA = GetObject(, "CATIA.Application")

                                CATIA = CreateObject("CATIA.Application")
                                CATIA.Visible = True
                                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB50\CP95NDB50_320_QAN0194.CATPart")

                                On Error GoTo 0
                            ElseIf ComboBox3.SelectedItem = "63 [mm]" Then
                                If ComboBox4.SelectedItem = "50 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_50_QAN0141.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_80_QAN0142.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_100_QAN0143.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_125_QAN0144.CATPart")

                                    On Error GoTo 0

                                ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_160_QAN0145.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "200 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_200_QAN0146.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "250 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_250_QAN0195.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "300 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_300_QAN0195.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "320 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_320_QAN0196.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "350 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_350_QAN0196.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox4.SelectedItem = "400 [mm]" Then
                                    'Get CATIA Or Launch it if necessary.
                                    On Error Resume Next
                                    CATIA = GetObject(, "CATIA.Application")

                                    CATIA = CreateObject("CATIA.Application")
                                    CATIA.Visible = True
                                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB63\CP95NDB63_400_QAN0197.CATPart")

                                    On Error GoTo 0
                                ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                                    If ComboBox4.SelectedItem = "50 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_50_QAN0147.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "80 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_80_QAN0148.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "100 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_100_QAN0149.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_125_QAN0150.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_160_QAN0151.CATPart")

                                        On Error GoTo 0

                                    ElseIf ComboBox4.SelectedItem = "200 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_200_QAN0152.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "250 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_250_QAN0198.CATPart")

                                        On Error GoTo 0

                                    ElseIf ComboBox4.SelectedItem = "320 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_320_QAN0199.CATPart")

                                        On Error GoTo 0

                                    ElseIf ComboBox4.SelectedItem = "400 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_400_QAN0200.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "500 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_500_QAN0201.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "600 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_600_QAN0202.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox4.SelectedItem = "700 [mm]" Then
                                        'Get CATIA Or Launch it if necessary.
                                        On Error Resume Next
                                        CATIA = GetObject(, "CATIA.Application")

                                        CATIA = CreateObject("CATIA.Application")
                                        CATIA.Visible = True
                                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB80\CP95NDB80_700_QAN0202.CATPart")

                                        On Error GoTo 0
                                    ElseIf ComboBox3.SelectedItem = "100 [mm]" Then

                                        If ComboBox4.SelectedItem = "100 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_100_QAN0153.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "125 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_125_QAN0154.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "160 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_160_QAN0155.CATPart")

                                            On Error GoTo 0

                                        ElseIf ComboBox4.SelectedItem = "200 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_200_QAN0156.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "250 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_250_QAN0157.CATPart")

                                            On Error GoTo 0

                                        ElseIf ComboBox4.SelectedItem = "320 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_320_QAN0158.CATPart")

                                            On Error GoTo 0

                                        ElseIf ComboBox4.SelectedItem = "400 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_400_QAN0203.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "500 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_500_QAN0204.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "600 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_600_QAN0205.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "650 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_650_QAN0205.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "700 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_700_QAN0206.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "800 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_800_QAN0207.CATPart")

                                            On Error GoTo 0
                                        ElseIf ComboBox4.SelectedItem = "900 [mm]" Then
                                            'Get CATIA Or Launch it if necessary.
                                            On Error Resume Next
                                            CATIA = GetObject(, "CATIA.Application")

                                            CATIA = CreateObject("CATIA.Application")
                                            CATIA.Visible = True
                                            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\CYLINDER\SMC\CYLINDER\CP95NDB100\CP95NDB100_900_QAN0208.CATPart")

                                            On Error GoTo 0
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
        End Select

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class