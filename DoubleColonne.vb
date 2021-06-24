Public Class DoubleColonne
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image
    Dim picture5 As Image
    Dim picture6 As Image
    Dim picture7 As Image
    Private Sub DoubleColonne_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("SMC")
        ComboBox1.Items.Add("Tuenkers")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi200.Show()
        Me.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA32TF-XC35W\MGPA32TF.PNG") ' 1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA40TF-XC35W\MGPA40TF.PNG") ' 2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF.PNG") ' 3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF.PNG") ' 4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF.PNG") ' 5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF.PNG") ' 6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_TUENKERS.PNG") ' 7

            Me.ComboBox2.Items.Clear()
            Me.ComboBox3.Items.Clear()
            ComboBox1.Focus()
            ComboBox2.ResetText()
            ComboBox3.ResetText()
            If Me.ComboBox1.SelectedItem.ToString = "SMC" Then
                Me.ComboBox2.Items.Add("32 [mm]")
                Me.ComboBox2.Items.Add("40 [mm]")
                Me.ComboBox2.Items.Add("50 [mm]")
                Me.ComboBox2.Items.Add("63 [mm]")
                Me.ComboBox2.Items.Add("80 [mm]")
                Me.ComboBox2.Items.Add("100 [mm]")
                Label1.Text = "Strok Seçiniz"
                Image1.Image = picture1

            ElseIf Me.ComboBox1.SelectedItem.ToString = "Tuenkers" Then
                Me.ComboBox2.Items.Add("40 [mm]")
                Me.ComboBox2.Items.Add("63 [mm]")
                Image1.Image = picture7
                Label1.Text = "Ekipman Tipini Seçiniz"

            End If
            With Me.ComboBox1
                Select Case ComboBox1.SelectedItem
                    'Case "3000_E1166943.."
                    '    Image1.Image = picture1
                    'Case "3000_E1166944.."
                    '    Image1.Image = picture3
                    'Case "3000_E1166945.."
                    '    Image1.Image = picture6
                    'Case "3000_E1166946.."
                    '    Image1.Image = picture71
                    'Case "3000_E1166947.."
                    '    Image1.Image = picture77
                    'Case "3000_E1166948.."
                    '    Image1.Image = picture80
                    'Case "3000_E1166949.."
                    '    Image1.Image = picture83
                    'Case "3000_E1166951.. (Oblong)"
                    '    Image1.Image = picture91
                    'Case "3000_E1166952.. (Oblong)"
                    '    Image1.Image = picture125
                    'Case "3000_E1166953.. (Oblong)"
                    '    Image1.Image = picture137
                End Select
            End With
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA32TF-XC35W\MGPA32TF.PNG") ' 1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA40TF-XC35W\MGPA40TF.PNG") ' 2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF.PNG") ' 3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF.PNG") ' 4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF.PNG") ' 5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF.PNG") ' 6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_TUENKERS.PNG") ' 7        

            ComboBox3.ResetText()
            ComboBox2.Focus()
            ComboBox3.Items.Clear()
            If Me.ComboBox1.SelectedItem.ToString = "SMC" Then
                If Me.ComboBox2.SelectedItem.ToString = "32 [mm]" Then
                    Me.ComboBox3.Items.Add("25 [mm]")
                    Me.ComboBox3.Items.Add("50 [mm]")
                    Me.ComboBox3.Items.Add("75 [mm]")
                    Me.ComboBox3.Items.Add("100 [mm]")
                    Me.ComboBox3.Items.Add("125 [mm]")
                    Image1.Image = picture1
                ElseIf Me.ComboBox2.SelectedItem.ToString = "40 [mm]" Then
                    Me.ComboBox3.Items.Add("25 [mm]")
                    Me.ComboBox3.Items.Add("50 [mm]")
                    Me.ComboBox3.Items.Add("75 [mm]")
                    Me.ComboBox3.Items.Add("100 [mm]")
                    Me.ComboBox3.Items.Add("125 [mm]")
                    Image1.Image = picture2
                ElseIf Me.ComboBox2.SelectedItem.ToString = "50 [mm]" Then
                    Me.ComboBox3.Items.Add("25 [mm]")
                    Me.ComboBox3.Items.Add("50 [mm]")
                    Me.ComboBox3.Items.Add("75 [mm]")
                    Me.ComboBox3.Items.Add("100 [mm]")
                    Me.ComboBox3.Items.Add("125 [mm]")
                    Me.ComboBox3.Items.Add("150 [mm]")
                    Image1.Image = picture3
                ElseIf Me.ComboBox2.SelectedItem.ToString = "63 [mm]" Then
                    Me.ComboBox3.Items.Add("25 [mm]")
                    Me.ComboBox3.Items.Add("50 [mm]")
                    Me.ComboBox3.Items.Add("75 [mm]")
                    Me.ComboBox3.Items.Add("100 [mm]")
                    Me.ComboBox3.Items.Add("125 [mm]")
                    Me.ComboBox3.Items.Add("150 [mm]")
                    Image1.Image = picture4
                ElseIf Me.ComboBox2.SelectedItem.ToString = "80 [mm]" Then
                    Me.ComboBox3.Items.Add("25 [mm]")
                    Me.ComboBox3.Items.Add("50 [mm]")
                    Me.ComboBox3.Items.Add("75 [mm]")
                    Me.ComboBox3.Items.Add("100 [mm]")
                    Me.ComboBox3.Items.Add("125 [mm]")
                    Me.ComboBox3.Items.Add("150 [mm]")
                    Me.ComboBox3.Items.Add("175 [mm]")
                    Me.ComboBox3.Items.Add("200 [mm]")
                    Image1.Image = picture5
                ElseIf Me.ComboBox2.SelectedItem.ToString = "100 [mm]" Then
                    Me.ComboBox3.Items.Add("25 [mm]")
                    Me.ComboBox3.Items.Add("50 [mm]")
                    Me.ComboBox3.Items.Add("75 [mm]")
                    Me.ComboBox3.Items.Add("100 [mm]")
                    Me.ComboBox3.Items.Add("125 [mm]")
                    Me.ComboBox3.Items.Add("150 [mm]")
                    Me.ComboBox3.Items.Add("175 [mm]")
                    Me.ComboBox3.Items.Add("200 [mm]")
                    Image1.Image = picture6
                End If
            End If
            If Me.ComboBox1.SelectedItem.ToString = "Tuenkers" Then

                If Me.ComboBox2.SelectedItem.ToString = "40 [mm]" Then
                    Me.ComboBox3.Items.Add("A13_T12")
                    Me.ComboBox3.Items.Add("A18_T12")
                    Me.ComboBox3.Items.Add("A23_T12")
                    Me.ComboBox3.Items.Add("A28_T12")
                    Image1.Image = picture7
                ElseIf Me.ComboBox2.SelectedItem.ToString = "63 [mm]" Then
                    Me.ComboBox3.Items.Add("63_5_A13_T12")
                    Me.ComboBox3.Items.Add("63_5_Z_A13_T12")
                    Image1.Image = picture7
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Select Case ComboBox1.SelectedItem
            Case "SMC"
                If ComboBox2.SelectedItem = "32 [mm]" Then
                    If ComboBox3.SelectedItem = "25 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA32TF-XC35W\MGPA32TF-25Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA32TF-XC35W\MGPA32TF-50Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "75 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA32TF-XC35W\MGPA32TF-75Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA32TF-XC35W\MGPA32TF-100Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA32TF-XC35W\MGPA32TF-125Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    End If
                ElseIf ComboBox2.SelectedItem = "40 [mm]" Then
                    If ComboBox3.SelectedItem = "25 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA40TF-XC35W\MGPA40TF-25Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA40TF-XC35W\MGPA40TF-50Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "75 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA40TF-XC35W\MGPA40TF-75Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA40TF-XC35W\MGPA40TF-100Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA40TF-XC35W\MGPA40TF-125Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    End If
                ElseIf ComboBox2.SelectedItem = "50 [mm]" Then
                    If ComboBox3.SelectedItem = "25 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF-25Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF-50Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "75 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF-75Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF-100Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF-125Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "150 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA50TF-XC35W\MGPA50TF-150Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    End If
                ElseIf ComboBox2.SelectedItem = "63 [mm]" Then
                    If ComboBox3.SelectedItem = "25 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF-25Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF-50Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "75 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF-75Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF-100Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF-125Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "150 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA63TF-XC35W\MGPA63TF-150Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    End If
                ElseIf ComboBox2.SelectedItem = "80 [mm]" Then
                    If ComboBox3.SelectedItem = "25 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-25Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-50Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "75 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-75Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-100Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-125Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "150 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-150Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "175 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-175Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "200 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA80TF-XC35W\MGPA80TF-200Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    End If
                ElseIf ComboBox2.SelectedItem = "100 [mm]" Then
                    If ComboBox3.SelectedItem = "25 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-25Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "50 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-50Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "75 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-75Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "100 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-100Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "125 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-125Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    ElseIf ComboBox3.SelectedItem = "150 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-150Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "175 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-175Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                    ElseIf ComboBox3.SelectedItem = "200 [mm]" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\SMC\MGPA100TF-XC35W\MGPA100TF-200Z-XC35W.CATPart")

                        On Error GoTo 0
                        ' Add a new Part


                        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                    End If
                End If
            Case "Tuenkers"
                'Me.ComboBox3.Items.Add("63_5_A13_T12")
                'Me.ComboBox3.Items.Add("63_5_Z_A13_T12")
                If ComboBox3.SelectedItem = "A13_T12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_40_A13_A18_T12\SZKD_40_A13_T12_TUENKERS.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "A18_T12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_40_A13_A18_T12\SZKD_40_A18_T12_TUENKERS.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

                ElseIf ComboBox3.SelectedItem = "A23_T12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_40_A23_A28_T12\SZKD_40_A23_T12_TUENKERS.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "A28_T12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_40_A23_A28_T12\SZKD_40_A28_T12_TUENKERS.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.


                ElseIf ComboBox3.SelectedItem = "63_5_A13_T12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_63_5_A13_T12\SZKD_63_5_A13_T12_TUENKERS.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                ElseIf ComboBox3.SelectedItem = "63_5_Z_A13_T12" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\DOUBLE COLONNE\TUENKERS\SZKD_63_5_Z_A13_T12\SZKD_63_5_Z_A13_T12_TUENKERS.CATPart")

                    On Error GoTo 0
                    ' Add a new Part


                    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.



                End If
        End Select
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class