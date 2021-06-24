Public Class Serrage
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    'D40
    Dim picture4 As Image
    Dim picture5 As Image
    Dim picture6 As Image
    Dim picture7 As Image
    Dim picture8 As Image
    Dim picture9 As Image

    Private Sub Serrage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("50 [mm],63 [mm]")
        ComboBox1.Items.Add("40 [mm]")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4001.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4003.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4004.png") '3
            'D40
            'Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4006.png") '4
            'Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4007.png") '5
            'Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4008.png") '6
            'Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4009.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4010.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4010D40.png") '9
            Me.ComboBox2.Items.Clear()
            ComboBox2.ResetText()

            If Me.ComboBox1.SelectedItem.ToString = "50 [mm],63 [mm]" Then
                Me.ComboBox2.Items.Add("60 [mm]")
                Me.ComboBox2.Items.Add("100 [mm]")
                Me.ComboBox2.Items.Add("120 [mm]")
                Me.ComboBox2.Items.Add("140 [mm]")
                Me.ComboBox2.Items.Add("Offset")
                Me.ComboBox2.Items.Add("80 [mm]")
            Else
                Me.ComboBox2.Items.Add("40 [mm]")
                Me.ComboBox2.Items.Add("80 [mm]")
                Me.ComboBox2.Items.Add("100 [mm]")
                Me.ComboBox2.Items.Add("120 [mm]")
                Me.ComboBox2.Items.Add("Offset D40")
            End If
            ComboBox1.Focus()
            With Me.ComboBox1
                Select Case ComboBox1.SelectedItem
                    Case "50 [mm],63 [mm]"
                        Image1.Image = picture1
                    Case "40 [mm]"
                        Image1.Image = picture8
                End Select
            End With
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4001.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4003.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4004.png") '3
            'D40
            'Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4006.png") '4
            'Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4007.png") '5
            'Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4008.png") '6
            'Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4009.png") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4010.png") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4010D40.png") '9
            ComboBox2.Focus()
            With Me.ComboBox2
                Select Case ComboBox1.SelectedItem
                    Case "50 [mm],63 [mm]"
                        If ComboBox2.SelectedItem = "60 [mm]" Then
                            Image1.Image = picture1
                        ElseIf ComboBox2.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture1
                        ElseIf ComboBox2.SelectedItem = "100 [mm]" Then
                            Image1.Image = picture3
                        ElseIf ComboBox2.SelectedItem = "120 [mm]" Then
                            Image1.Image = picture2
                        ElseIf ComboBox2.SelectedItem = "140 [mm]" Then
                            Image1.Image = picture3
                        ElseIf ComboBox2.SelectedItem = "Offset" Then
                            Image1.Image = picture8
                        End If
                    Case "40 [mm]"
                        If ComboBox2.SelectedItem = "40 [mm]" Then
                            Image1.Image = picture8
                        ElseIf ComboBox2.SelectedItem = "80 [mm]" Then
                            Image1.Image = picture8
                        ElseIf ComboBox2.SelectedItem = "100 [mm]" Then
                            Image1.Image = picture8
                        ElseIf ComboBox2.SelectedItem = "120 [mm]" Then
                            Image1.Image = picture8
                        ElseIf ComboBox2.SelectedItem = "Offset D40" Then
                            Image1.Image = picture9
                        End If

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
        Dim Msg
        Select Case ComboBox1.SelectedItem
            Case "50 [mm],63 [mm]"
                If ComboBox2.SelectedItem = "60 [mm]" Then
                    Msg = MsgBox("Parça Bulunamadı",, "Hata")
                ElseIf ComboBox2.SelectedItem = "80 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4001_CLAMP_BRACKET_DIN_h80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox2.SelectedItem = "100 [mm]" Then
                    MsgBox("Parça Bulunamadı",, "Hata")
                ElseIf ComboBox2.SelectedItem = "120 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4003_CLAMP_BRACKET_DIN_h120.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox2.SelectedItem = "140 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\4004_CLAMP_BRACKET_DIN_h140.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox2.SelectedItem = "Offset" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4010_CLAMP_BRACKET_DIN_OFFSET.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            Case "40 [mm]"
                If ComboBox2.SelectedItem = "40 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4006_CLAMP_BRACKET_DIN_D40_h40.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox2.SelectedItem = "80 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4007_CLAMP_BRACKET_DIN_D40_h80.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox2.SelectedItem = "100 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4008_CLAMP_BRACKET_DIN_D40_h100.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox2.SelectedItem = "120 [mm]" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4009_CLAMP_BRACKET_DIN_D40_h120.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox2.SelectedItem = "Offset D40" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_04000_CLAMP_BRACKET_support_serrage\3D_new_2\D40_3D_pour_Destaco\4010_CLAMP_BRACKET_DIN_D40_OFFSET.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                End If

        End Select
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class