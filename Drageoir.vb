Public Class Drageoir
    'Horizontal
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
    'Vertical
    Dim picture13 As Image
    Dim picture14 As Image
    Dim picture15 As Image
    Dim picture16 As Image
    Dim picture17 As Image
    Dim picture18 As Image
    Dim picture19 As Image
    Dim picture20 As Image
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            'Horizontal
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2000H.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2001H.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2002H.png") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2003H.png") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2010H.png") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2011H.png") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2012H.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2013H.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2020H.png") '9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2021H.png") '10
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2022H.png") '11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2023H.png") '12
            'Vertical
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2050.png") '13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2051.jpg") '14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2052.png") '15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2053.png") '16
            Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2056.png") '17
            Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2057.jpg") '18
            Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2058.png") '19
            Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2059.png") '20
            ComboBox2.ResetText()
            ComboBox3.ResetText()

            Me.ComboBox2.Items.Clear()
            Me.ComboBox3.Items.Clear()
            If Me.ComboBox1.SelectedItem.ToString = "Horizontal" Then
                Me.ComboBox2.Items.Add("85 [mm]")
                Me.ComboBox2.Items.Add("110 [mm]")
                Me.ComboBox2.Items.Add("135 [mm]")
                Me.ComboBox2.Items.Add("160 [mm]")
                Me.ComboBox3.Items.Add("55 [mm]")
                Me.ComboBox3.Items.Add("80 [mm]")
                Me.ComboBox3.Items.Add("105 [mm]")
            Else
                Me.ComboBox2.Items.Add("20 [mm]")
                Me.ComboBox2.Items.Add("20 [mm]/ 4 Delik")
                Me.ComboBox3.Items.Add("40 [mm]")
                Me.ComboBox3.Items.Add("60 [mm]")
                Me.ComboBox3.Items.Add("80 [mm]")
                Me.ComboBox3.Items.Add("100 [mm]")
            End If
            ComboBox1.Focus()
            With Me.ComboBox1
                Select Case ComboBox1.SelectedItem
                    Case "Horizontal"
                        Image1.Image = picture1
                    Case "Vertical"
                        Image1.Image = picture13
                End Select
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            'Horizontal
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2000H.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2001H.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2002H.png") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2003H.png") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2010H.png") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2011H.png") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2012H.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2013H.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2020H.png") '9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2021H.png") '10
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2022H.png") '11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2023H.png") '12
            'Vertical
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2050.png") '13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2051.jpg") '14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2052.png") '15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2053.png") '16
            Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2056.png") '17
            Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2057.jpg") '18
            Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2058.png") '19
            Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2059.png") '20
            ComboBox2.Focus()
            ComboBox3.ResetText()

            With Me.ComboBox2
                Select Case ComboBox1.SelectedItem
                    Case "Horizontal"
                        If ComboBox2.SelectedItem = "85 [mm]" Then
                            Image1.Image = picture1
                        ElseIf ComboBox2.SelectedItem = "110 [mm]" Then
                            Image1.Image = picture2
                        ElseIf ComboBox2.SelectedItem = "135 [mm]" Then
                            Image1.Image = picture3
                        ElseIf ComboBox2.SelectedItem = "160 [mm]" Then
                            Image1.Image = picture4
                        End If
                    Case "Vertical"
                        Image1.Image = picture13
                End Select
            End With
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            'Horizontal
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2000H.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2001H.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2002H.png") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2003H.png") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2010H.png") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2011H.png") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2012H.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2013H.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2020H.png") '9
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2021H.png") '10
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2022H.png") '11
            Dim picture12 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2023H.png") '12
            'Vertical
            Dim picture13 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2050.png") '13
            Dim picture14 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2051.jpg") '14
            Dim picture15 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2052.png") '15
            Dim picture16 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2053.png") '16
            Dim picture17 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2056.png") '17
            Dim picture18 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2057.jpg") '18
            Dim picture19 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2058.png") '19
            Dim picture20 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2059.png") '20
        Catch ex As Exception
            ComboBox3.Focus()
            With Me.ComboBox3
                If ComboBox1.SelectedItem = "Horizontal" Then
                    If ComboBox2.SelectedItem = "85 [mm]" Then
                        If ComboBox3.SelectedItem = "55 [mm]" Then
                            Image1.Image = picture1
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture5
                        ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                            Image1.Image = picture9
                        End If
                    End If
                    If ComboBox2.SelectedItem = "110 [mm]" Then
                        If ComboBox3.SelectedItem = "55 [mm]" Then
                            Image1.Image = picture2
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture6
                        ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                            Image1.Image = picture10
                        End If
                    End If
                    If ComboBox2.SelectedItem = "135 [mm]" Then
                        If ComboBox3.SelectedItem = "55 [mm]" Then
                            Image1.Image = picture3
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture7
                        ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                            Image1.Image = picture11
                        End If
                    End If
                    If ComboBox2.SelectedItem = "160 [mm]" Then
                        If ComboBox3.SelectedItem = "55 [mm]" Then
                            Image1.Image = picture4
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture8
                        ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                            Image1.Image = picture12
                        End If
                    End If
                End If
                If ComboBox1.SelectedItem = "Vertical" Then
                    If ComboBox2.SelectedItem = "20 [mm]" Then
                        If ComboBox3.SelectedItem = "40 [mm]" Then
                            Image1.Image = picture13
                        ElseIf ComboBox3.SelectedItem = "60 [mm]" Then
                            Image1.Image = picture14
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture15
                        ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                            Image1.Image = picture16
                        End If
                    End If
                    If ComboBox2.SelectedItem = "20 [mm]/ 4 Delik" Then
                        If ComboBox3.SelectedItem = "40 [mm]" Then
                            Image1.Image = picture17
                        ElseIf ComboBox3.SelectedItem = "60 [mm]" Then
                            Image1.Image = picture18
                        ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture19
                        ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                            Image1.Image = picture20
                        End If
                    End If

                End If
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        If ComboBox1.SelectedItem = "Horizontal" Then
            If ComboBox2.SelectedItem = "85 [mm]" Then
                If ComboBox3.SelectedItem = "55 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2000_INTERFACE PLATE_L85_H55.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2010_INTERFACE PLATE_L85_H80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2020_INTERFACE PLATE_L85_H105.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                End If
            End If
            If ComboBox2.SelectedItem = "110 [mm]" Then
                If ComboBox3.SelectedItem = "55 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2001_INTERFACE PLATE_L110_H55.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2011_INTERFACE PLATE_L110_H80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2021_INTERFACE PLATE_L110_H105.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                End If
            End If
            If ComboBox2.SelectedItem = "135 [mm]" Then
                If ComboBox3.SelectedItem = "55 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2002_INTERFACE PLATE_L135_H55.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2012_INTERFACE PLATE_L135_H80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2022_INTERFACE PLATE_L135_H105.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                End If
            End If
            If ComboBox2.SelectedItem = "160 [mm]" Then
                If ComboBox3.SelectedItem = "55 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2003_INTERFACE PLATE_L160_H55.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2013_INTERFACE PLATE_L160_H80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "105 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2000_H\2023_INTERFACE PLATE_L160_H105.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            End If
        End If
        If ComboBox1.SelectedItem = "Vertical" Then
            If ComboBox2.SelectedItem = "20 [mm]" Then
                If ComboBox3.SelectedItem = "40 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2050_INTERFACE_PLATE_V_L20_H40.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "60 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2051_INTERFACE_PLATE_V_L20_H60.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2052_INTERFACE_PLATE_V_L20_H80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2053_INTERFACE_PLATE_V_L20_H100.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            End If
            If ComboBox2.SelectedItem = "20 [mm]/ 4 Delik" Then
                If ComboBox3.SelectedItem = "40 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2056_INTERFACE_PLATE_V_L20_H40.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "60 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2057_INTERFACE_PLATE_V_L20_H60.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "80 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2058_INTERFACE_PLATE_V_L20_H80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_02000_INTERFACE PLATE_drageoir\3D\2050_V\2059_INTERFACE_PLATE_V_L20_H100.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            End If
        End If
    End Sub
    Private Sub Drageoir_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Horizontal")
        ComboBox1.Items.Add("Vertical")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    'Private Sub Image1_Click(sender As Object, e As EventArgs) Handles Image1.Click

    'End Sub
End Class