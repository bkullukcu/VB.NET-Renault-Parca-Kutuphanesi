Public Class PinBracket
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image
    Dim picture5 As Image
    Dim picture6 As Image
    Dim picture7 As Image
    Dim picture8 As Image

    'Private Sub Image1_Click(sender As Object, e As EventArgs) Handles Image1.Click
    ' End Sub

    Private Sub PinBracket_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Dikey")
        ComboBox1.Items.Add("Yatay")
        ComboBox2.Items.Add("Yan")
        ComboBox2.Items.Add("Ön")
        ComboBox3.Items.Add("6 [mm]")
        ComboBox3.Items.Add("10 [mm]")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.Focus()
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001114.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001120.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001115.jpg") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001121.jpg") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001116.jpg") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001117.jpg") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001123.jpg") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001122.jpg") '8

            With Me.ComboBox1

                Select Case ComboBox1.SelectedItem
                    Case "Dikey"
                        Image1.Image = picture1
                    Case "Yatay"
                        Image1.Image = picture2
                End Select
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox2.Focus()
        ComboBox3.ResetText()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001114.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001120.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001115.jpg") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001121.jpg") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001116.jpg") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001117.jpg") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001123.jpg") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001122.jpg") '8

            With Me.ComboBox2
                If ComboBox1.SelectedItem = "Dikey" Then
                    If ComboBox2.SelectedItem = "Yan" Then
                        Image1.Image = picture3
                    ElseIf ComboBox2.SelectedItem = "Ön" Then
                        Image1.Image = picture1
                    End If
                End If
                If ComboBox1.SelectedItem = "Yatay" Then
                    If ComboBox2.SelectedItem = "Yan" Then
                        Image1.Image = picture4
                    ElseIf ComboBox2.SelectedItem = "Ön" Then
                        Image1.Image = picture4
                    End If
                End If
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001114.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001120.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001115.jpg") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001121.jpg") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001116.jpg") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001117.jpg") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001123.jpg") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001122.jpg") '8
            ComboBox3.Focus()
            With Me.ComboBox3
                If ComboBox1.SelectedItem = "Dikey" Then
                    If ComboBox2.SelectedItem = "Yan" Then
                        If ComboBox3.SelectedItem = "6 [mm]" Then
                            Image1.Image = picture3
                        Else
                            Image1.Image = picture6
                        End If
                    End If
                    If ComboBox2.SelectedItem = "Ön" Then
                        If ComboBox3.SelectedItem = "6 [mm]" Then
                            Image1.Image = picture1
                        Else
                            Image1.Image = picture5
                        End If
                    End If
                End If
                If ComboBox1.SelectedItem = "Yatay" Then
                    If ComboBox2.SelectedItem = "Yan" Then
                        If ComboBox3.SelectedItem = "6 [mm]" Then
                            Image1.Image = picture4
                        Else
                            Image1.Image = picture7
                        End If
                    End If
                    If ComboBox2.SelectedItem = "Ön" Then
                        If ComboBox3.SelectedItem = "6 [mm]" Then
                            Image1.Image = picture2
                        Else
                            Image1.Image = picture8
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
        If ComboBox1.SelectedItem = "Dikey" Then
            If ComboBox2.SelectedItem = "Yan" Then
                If ComboBox3.SelectedItem = "6 [mm]" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001115.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                Else
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001117.CATPart")

                    On Error GoTo 0
                    'Get CATIA or Launch it if necessary.
                    ' Add a new Part

                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            End If
            If ComboBox2.SelectedItem = "Ön" Then
                If ComboBox3.SelectedItem = "6 [mm]" Then
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001114.CATPart")

                    On Error GoTo 0
                    ' Add a new Part

                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                Else
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001116.CATPart")

                    On Error GoTo 0
                    ' Add a new Part

                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            End If
        End If
        If ComboBox1.SelectedItem = "Yatay" Then
            If ComboBox2.SelectedItem = "Yan" Then
                If ComboBox3.SelectedItem = "6 [mm]" Then
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001121.CATPart")

                    On Error GoTo 0
                    ' Add a new Part

                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                Else
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001123.CATPart")

                    On Error GoTo 0
                    ' Add a new Part

                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            End If
            If ComboBox2.SelectedItem = "Ön" Then
                If ComboBox3.SelectedItem = "6 [mm]" Then
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001120.CATPart")

                    On Error GoTo 0
                    ' Add a new Part

                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                Else
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_03000_PIN BRACKET_support pilote fixe\3D\E140001122.CATPart")

                    On Error GoTo 0
                    ' Add a new Part

                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                End If
            End If
        End If
        'CATIA.ActiveDocument.Part.InWorkObject = MyBody1 ' Activate "PartDesign"
        'On Error GoTo 0
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

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click

    End Sub
End Class