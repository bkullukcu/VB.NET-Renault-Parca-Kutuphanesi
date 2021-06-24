Public Class Serrages
    Dim picture1 As Image
    Dim picture2 As Image
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        ComboBox4.ResetText()
        Me.ComboBox2.Items.Clear()
        Me.ComboBox3.Items.Clear()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\Serrages DESTACO.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\Serrages TUENKERS.PNG") '2

            If ComboBox1.SelectedItem.ToString = "DESTACO" Then
                ComboBox2.Items.Add("40 [mm]")
                ComboBox2.Items.Add("50 [mm]")
                ComboBox2.Items.Add("63 [mm]")
                ComboBox2.Items.Add("80 [mm]")
                ComboBox3.Items.Add("R")
                ComboBox3.Items.Add("M")
                ComboBox3.Items.Add("L")
                ComboBox4.Items.Add("Horizontal")
                ComboBox4.Items.Add("Vertical")
                Image1.Image = picture1
            ElseIf ComboBox1.SelectedItem.ToString = "TUENKERS" Then
                ComboBox2.Items.Add("40 [mm]")
                ComboBox2.Items.Add("50 [mm]")
                ComboBox2.Items.Add("63 [mm]")
                ComboBox2.Items.Add("80 [mm]")
                ComboBox3.Items.Add("A10")
                ComboBox3.Items.Add("A11")
                ComboBox3.Items.Add("A12")
                ComboBox4.Items.Add("Horizontal")
                ComboBox4.Items.Add("Vertical")
                Image1.Image = picture2
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PilotesMobile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("DESTACO")
        ComboBox1.Items.Add("TUENKERS")
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        ComboBox3.ResetText()
        ComboBox4.ResetText()

        Me.ComboBox3.Items.Clear()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\Serrages DESTACO.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\Serrages TUENKERS.PNG") '2
            'If ComboBox1.SelectedItem.ToString = "SMC" Then
            '    If ComboBox2.SelectedItem.ToString = "40 [mm]" Then

            '        If ComboBox1.SelectedItem.ToString = "Tuenkers" Then
            '            If ComboBox2.SelectedItem.ToString = "40 [mm]" Then


            '            ElseIf ComboBox2.SelectedItem.ToString = "40_1 [mm]" Then

            '            ElseIf ComboBox2.SelectedItem.ToString = "63 [mm]" Then

            '            ElseIf ComboBox2.SelectedItem.ToString = "63_1 [mm]" Then

            '            ElseIf ComboBox2.SelectedItem.ToString = "40 [mm]" Then
            '            ElseIf ComboBox2.SelectedItem.ToString = "60 [mm]" Then
            '            ElseIf ComboBox2.SelectedItem.ToString = "R12" Then
            '            ElseIf ComboBox2.SelectedItem.ToString = "R20" Then
            '            End If
            '        End If
            '    End If
            'End If
            'If ComboBox1.SelectedItem.ToString = "SMC" Then


            '    ComboBox3.Items.Add("40 [mm]")
            '    ComboBox3.Items.Add("60 [mm]")

            'Else

            '    ComboBox3.Items.Add("40 [mm]")
            '    'ComboBox3.Items.Add("60 [mm]")

            'End If
            If ComboBox1.SelectedItem.ToString = "DESTACO" Then
                If ComboBox2.SelectedItem.ToString = "80 [mm]" Then
                    ComboBox3.Items.Add("M")
                    ComboBox4.Items.Add("Horizontal")
                    ComboBox4.Items.Add("Vertical")
                    Image1.Image = picture1
                    'ComboBox2.Items.Add("40 [mm]")
                    'ComboBox2.Items.Add("50 [mm]")
                    'ComboBox2.Items.Add("63 [mm]")
                    'ComboBox2.Items.Add("80 [mm]")
                Else
                    ComboBox3.Items.Add("R")
                    ComboBox3.Items.Add("M")
                    ComboBox3.Items.Add("L")
                    ComboBox4.Items.Add("Horizontal")
                    ComboBox4.Items.Add("Vertical")
                    Image1.Image = picture1
                End If

            ElseIf ComboBox1.SelectedItem.ToString = "TUENKERS" Then
                'ComboBox2.Items.Add("40 [mm]")
                'ComboBox2.Items.Add("50 [mm]")
                'ComboBox2.Items.Add("63 [mm]")
                'ComboBox2.Items.Add("80 [mm]")
                ComboBox3.Items.Add("A10")
                ComboBox3.Items.Add("A11")
                ComboBox3.Items.Add("A12")
                ComboBox4.Items.Add("Horizontal")
                ComboBox4.Items.Add("Vertical")
                Image1.Image = picture2
            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox4.ResetText()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\Serrages DESTACO.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\Serrages TUENKERS.PNG") '2
            If ComboBox1.SelectedItem.ToString = "DESTACO" Then
                'ComboBox2.Items.Add("40 [mm]")
                'ComboBox2.Items.Add("50 [mm]")
                'ComboBox2.Items.Add("63 [mm]")
                'ComboBox2.Items.Add("80 [mm]")
                'ComboBox3.Items.Add("R")
                'ComboBox3.Items.Add("M")
                'ComboBox3.Items.Add("L")
                ComboBox4.Items.Add("Horizontal")
                ComboBox4.Items.Add("Vertical")
                Image1.Image = picture1
            ElseIf ComboBox1.SelectedItem.ToString = "TUENKERS" Then
                'ComboBox2.Items.Add("40 [mm]")
                'ComboBox2.Items.Add("50 [mm]")
                'ComboBox2.Items.Add("63 [mm]")
                'ComboBox2.Items.Add("80 [mm]")
                'ComboBox3.Items.Add("A10")
                'ComboBox3.Items.Add("A11")
                'ComboBox3.Items.Add("A12")
                ComboBox4.Items.Add("Horizontal")
                ComboBox4.Items.Add("Vertical")
                Image1.Image = picture2
            End If
            'If ComboBox1.SelectedItem.ToString = "SMC" Then
            '    If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
            '        ComboBox4.Items.Add("12")
            '        ComboBox4.Items.Add("14")
            '    ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
            '        ComboBox4.Items.Add("11")
            '        ComboBox4.Items.Add("13")
            '    End If
            'ElseIf ComboBox1.SelectedItem.ToString = "Tuenkers" Then
            '    If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
            '        ComboBox4.Items.Add("R12")
            '        ComboBox4.Items.Add("R20")
            '        'ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
            '        '    ComboBox4.Items.Add("R12")
            '        '    ComboBox4.Items.Add("R20")
            '    End If

            'ComboBox4.Items.Add("R12")
            'ComboBox4.Items.Add("R20")
            'Label4.Text = "Ekipman Tipi"
            'Image1.Image = picture2
            'End If
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

        If ComboBox1.SelectedItem.ToString = "DESTACO" Then
            If ComboBox2.SelectedItem = "40 [mm]" Then
                If ComboBox3.SelectedItem = "R" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UR404-15-117_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "M" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8\82M-3E030040L8+Arm^8UM404-15-117_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8+Arm^8UM404-15-117_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "L" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8+Arm^8UL404-15-117_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030040L8+Arm^8UL404-15-117_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem = "50 [mm]" Then

                If ComboBox3.SelectedItem = "R" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030050L8\82M-3E030050L8+Arm^8UR504-15-144_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030050L8\82M-3E030050L8+Arm^8UR504-15-144_Vertical.CATPart")

                        On Error GoTo 0
                    End If

                ElseIf ComboBox3.SelectedItem = "M" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030050L8\82M-3E030050L8+Arm^8UM504-15-144_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030050L8\82M-3E030050L8+Arm^8UM504-15-144_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "L" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030050L8\82M-3E030050L8+Arm^8UL504-15-144_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030050L8\82M-3E030050L8+Arm^8UL504-15-144_Horizontal.CATPart")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem = "63 [mm]" Then
                If ComboBox3.SelectedItem = "R" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030063L8\82M-3E030063L8+Arm^8UR634-15-144_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030063L8\82M-3E030063L8+Arm^8UR634-15-144_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "M" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030063L8\82M-3E030063L8+Arm^8UM634-15-144_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030063L8\82M-3E030063L8+Arm^8UM634-15-144_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "L" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030063L8\82M-3E030063L8+Arm^8UL634-15-144_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E030063L8\82M-3E030063L8+Arm^8UL634-15-144_Veritical.CATPart")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem = "80 [mm]" Then
                '    If ComboBox3.SelectedItem = "R" Then
                '        If ComboBox4.SelectedItem = "Horizontal" Then
                '            'Get CATIA Or Launch it if necessary.
                '            On Error Resume Next
                '            CATIA = GetObject(, "CATIA.Application")
                '            
                '                CATIA = CreateObject("CATIA.Application")
                '                CATIA.Visible = True
                '                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-60-DCN711NN.step")

                '            End If
                '            On Error GoTo 0
                '        ElseIf ComboBox4.SelectedItem = "Vertical" Then
                '            'Get CATIA Or Launch it if necessary.
                '            On Error Resume Next
                '            CATIA = GetObject(, "CATIA.Application")
                '            
                '                CATIA = CreateObject("CATIA.Application")
                '                CATIA.Visible = True
                '                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-60-DCN711NN.step")

                '            End If
                '            On Error GoTo 0
                '        End If
                If ComboBox3.SelectedItem = "M" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E230080L8\82M-3E230080L8+Arm^8UM801-45-204_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\DESTACO\82M-3E230080L8\82M-3E230080L8+Arm^8UM801-45-204_Vertical.CATPart")

                        On Error GoTo 0
                        'ElseIf ComboBox3.SelectedItem = "L" Then

                        '    If ComboBox4.SelectedItem = "Horizontal" Then
                        '        'Get CATIA Or Launch it if necessary.
                        '        On Error Resume Next
                        '        CATIA = GetObject(, "CATIA.Application")
                        '        
                        '            CATIA = CreateObject("CATIA.Application")
                        '            CATIA.Visible = True
                        '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-60-DCN711NN.step")

                        '        End If
                        '        On Error GoTo 0
                        '    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        '        'Get CATIA Or Launch it if necessary.
                        '        On Error Resume Next
                        '        CATIA = GetObject(, "CATIA.Application")
                        '        
                        '            CATIA = CreateObject("CATIA.Application")
                        '            CATIA.Visible = True
                        '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-60-DCN711NN.step")

                        '        End If
                        '        On Error GoTo 0
                        '    End If
                        'End If
                    End If
                End If
            End If

        ElseIf ComboBox1.SelectedItem.ToString = "TUENKERS" Then
            If ComboBox2.SelectedItem = "40 [mm]" Then
                If ComboBox3.SelectedItem = "A10" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_40_BR2_A00_T12\V_V2_40_BR2_A00_T12+Arm^SPANNARM_KPL_40_A10_219528_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_40_BR2_A00_T12\V_V2_40_BR2_A00_T12+Arm^SPANNARM_KPL_40_A10_219528_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "A12" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_40_BR2_A00_T12\V_V2_40_BR2_A00_T12+Arm^SPANNARM_KPL_40_A12_235943_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_40_BR2_A00_T12\V_V2_40_BR2_A00_T12+Arm^SPANNARM_KPL_40_A12_235943_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "A11" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_40_BR2_A00_T12\V_V2_40_BR2_A00_T12+Arm^SPANNARM_KPL_40_A11_235942_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_40_BR2_A00_T12\V_V2_40_BR2_A00_T12+Arm^SPANNARM_KPL_40_A11_235942_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem = "50 [mm]" Then

                If ComboBox3.SelectedItem = "A10" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_50_1_BR2_A00_T12\V_V2_50_1_BR2_A00_T12+Arm^SPANNARM_KPL_50_A10_216884_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_50_1_BR2_A00_T12\V_V2_50_1_BR2_A00_T12+Arm^SPANNARM_KPL_50_A10_216884_Horizontal.CATPart")

                        On Error GoTo 0
                    End If

                ElseIf ComboBox3.SelectedItem = "A12" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_50_1_BR2_A00_T12\V_V2_50_1_BR2_A00_T12+Arm^SPANNARM_KPL_50_A12_216890_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_50_1_BR2_A00_T12\V_V2_50_1_BR2_A00_T12+Arm^SPANNARM_KPL_50_A12_216890_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "A11" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_50_1_BR2_A00_T12\V_V2_50_1_BR2_A00_T12+Arm^SPANNARM_KPL_50_A11_216886_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_50_1_BR2_A00_T12\V_V2_50_1_BR2_A00_T12+Arm^SPANNARM_KPL_50_A11_216886_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem = "63 [mm]" Then
                If ComboBox3.SelectedItem = "A10" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_63_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_63_A10_217498_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_63_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_63_A10_217498_Vertical.CATPart")

                        On Error GoTo 0
                    End If

                ElseIf ComboBox3.SelectedItem = "A12" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_63_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_63_A12_217500_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_63_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_63_A12_217500_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "A11" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_63_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_63_A11_217499_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_63_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_63_A11_217499_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem = "80 [mm]" Then
                If ComboBox3.SelectedItem = "A10" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_80_1_BR2_A00_T12\V_V2_80_1_BR2_A00_T12+Arm^SPANNARM_KPL_80_A10_227660_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_80_1_BR2_A00_T12\V_V2_80_1_BR2_A00_T12+Arm^SPANNARM_KPL_80_A10_227660_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "A12" Then
                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_80_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_80_A12_230858_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_80_1_BR2_A00_T12\V_V2_63_1_BR2_A00_T12+Arm^SPANNARM_KPL_80_A12_230858_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "A11" Then

                    If ComboBox4.SelectedItem = "Horizontal" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_80_1_BR2_A00_T12\V_V2_80_1_BR2_A00_T12+Arm^SSPANNARM_KPL_80_A11_227661_Horizontal.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem = "Vertical" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SERRAGES\TUENKERS\V_V2_80_1_BR2_A00_T12\V_V2_80_1_BR2_A00_T12+Arm^SPANNARM_KPL_80_A11_227661_Vertical.CATPart")

                        On Error GoTo 0
                    End If
                End If
            End If
        End If



    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class