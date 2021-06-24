Public Class PilotesMobile
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
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP_SMC.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR_TUENKERS.PNG") '2

            If ComboBox1.SelectedItem.ToString = "SMC" Then

                ComboBox2.Items.Add("40 [mm]")
                ComboBox3.Items.Add("40 [mm]")
                ComboBox3.Items.Add("60 [mm]")
                Image1.Image = picture1
                Label4.Text = "Pim Oturma Çapı (12-20)"
            ElseIf ComboBox1.SelectedItem.ToString = "Tuenkers" Then
                ComboBox2.Items.Add("40 [mm]")
                ComboBox2.Items.Add("40_1 [mm]")
                ComboBox2.Items.Add("63 [mm]")
                ComboBox2.Items.Add("63_1 [mm]")
                ComboBox3.Items.Add("40 [mm]")
                'ComboBox3.Items.Add("60 [mm]")
                ComboBox4.Items.Add("R12")
                ComboBox4.Items.Add("R20")
                Label4.Text = "Ekipman Tipi"
                Image1.Image = picture2


            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PilotesMobile_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Items.Add("SMC")
        ComboBox1.Items.Add("Tuenkers")


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        ComboBox3.ResetText()
        ComboBox4.ResetText()

        Me.ComboBox3.Items.Clear()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP_SMC.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR_TUENKERS.PNG") '2

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
            If ComboBox1.SelectedItem.ToString = "SMC" Then


                ComboBox3.Items.Add("40 [mm]")
                ComboBox3.Items.Add("60 [mm]")

            Else

                ComboBox3.Items.Add("40 [mm]")
                'ComboBox3.Items.Add("60 [mm]")

            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox4.ResetText()
        Me.ComboBox4.Items.Clear()
        If ComboBox1.SelectedItem.ToString = "SMC" Then
            If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                ComboBox4.Items.Add("12")
                ComboBox4.Items.Add("14")
            ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                ComboBox4.Items.Add("11")
                ComboBox4.Items.Add("13")
            End If
        ElseIf ComboBox1.SelectedItem.ToString = "Tuenkers" Then
            If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                ComboBox4.Items.Add("R12")
                ComboBox4.Items.Add("R20")
                'ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                '    ComboBox4.Items.Add("R12")
                '    ComboBox4.Items.Add("R20")
            End If

            'ComboBox4.Items.Add("R12")
            'ComboBox4.Items.Add("R20")
            'Label4.Text = "Ekipman Tipi"
            'Image1.Image = picture2
        End If


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
        If ComboBox1.SelectedItem.ToString = "SMC" Then
            If ComboBox2.SelectedItem.ToString = "40 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    If ComboBox4.SelectedItem.ToString = "12" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-40-DCN712NN.IGS")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "14" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-40-DCN714NN.IGS")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then

                    If ComboBox4.SelectedItem.ToString = "11" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-60-DCN711NN.step")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "13" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\SMC\CKZP\CKZP40-60-DCN713NN.IGS")

                        On Error GoTo 0
                    End If
                End If
            End If
        ElseIf ComboBox1.SelectedItem.ToString = "Tuenkers" Then
            If ComboBox2.SelectedItem.ToString = "40 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    If ComboBox4.SelectedItem.ToString = "R12" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_40_R12_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_40_R20_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    End If
                    'ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                    '    If ComboBox4.SelectedItem.ToString = "R12" Then
                    '        'Get CATIA Or Launch it if necessary.
                    '        On Error Resume Next
                    '        CATIA = GetObject(, "CATIA.Application")
                    '        
                    '            CATIA = CreateObject("CATIA.Application")
                    '            CATIA.Visible = True
                    '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                    '        End If
                    '        On Error GoTo 0
                    '    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                    '        'Get CATIA Or Launch it if necessary.
                    '        On Error Resume Next
                    '        CATIA = GetObject(, "CATIA.Application")
                    '        
                    '            CATIA = CreateObject("CATIA.Application")
                    '            CATIA.Visible = True
                    '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                    '        End If
                    '        On Error GoTo 0
                    '    End If
                    'End If
                End If
            ElseIf ComboBox2.SelectedItem.ToString = "40_1 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    If ComboBox4.SelectedItem.ToString = "R12" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_40_1_R12_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_40_1_R20_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    End If
                End If
                'ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                '    If ComboBox4.SelectedItem.ToString = "R12" Then
                '        'Get CATIA Or Launch it if necessary.
                '        On Error Resume Next
                '        CATIA = GetObject(, "CATIA.Application")
                '        
                '            CATIA = CreateObject("CATIA.Application")
                '            CATIA.Visible = True
                '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                '        End If
                '        On Error GoTo 0
                '    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                '        'Get CATIA Or Launch it if necessary.
                '        On Error Resume Next
                '        CATIA = GetObject(, "CATIA.Application")
                '        
                '            CATIA = CreateObject("CATIA.Application")
                '            CATIA.Visible = True
                '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                '        End If
                '        On Error GoTo 0
                '    End If
                'End If
            ElseIf ComboBox2.SelectedItem.ToString = "63 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    If ComboBox4.SelectedItem.ToString = "R12" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_63_R12_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_63_R20_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    End If
                End If
                'ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                '    If ComboBox4.SelectedItem.ToString = "R12" Then
                '        'Get CATIA Or Launch it if necessary.
                '        On Error Resume Next
                '        CATIA = GetObject(, "CATIA.Application")
                '        
                '            CATIA = CreateObject("CATIA.Application")
                '            CATIA.Visible = True
                '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                '        End If
                '        On Error GoTo 0
                '    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                '        'Get CATIA Or Launch it if necessary.
                '        On Error Resume Next
                '        CATIA = GetObject(, "CATIA.Application")
                '        
                '            CATIA = CreateObject("CATIA.Application")
                '            CATIA.Visible = True
                '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                '        End If
                '        On Error GoTo 0
                '    End If
                'End If
            ElseIf ComboBox2.SelectedItem.ToString = "63_1 [mm]" Then
                If ComboBox3.SelectedItem.ToString = "40 [mm]" Then
                    If ComboBox4.SelectedItem.ToString = "R12" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_63_1_R12_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\SZKR_63_1_R20_T12_TUENKERS.CATPart")

                        On Error GoTo 0
                    End If
                    'ElseIf ComboBox3.SelectedItem.ToString = "60 [mm]" Then
                    '    If ComboBox4.SelectedItem.ToString = "R12" Then
                    '        'Get CATIA Or Launch it if necessary.
                    '        On Error Resume Next
                    '        CATIA = GetObject(, "CATIA.Application")
                    '        
                    '            CATIA = CreateObject("CATIA.Application")
                    '            CATIA.Visible = True
                    '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                    '        End If
                    '        On Error GoTo 0
                    '    ElseIf ComboBox4.SelectedItem.ToString = "R20" Then
                    '        'Get CATIA Or Launch it if necessary.
                    '        On Error Resume Next
                    '        CATIA = GetObject(, "CATIA.Application")
                    '        
                    '            CATIA = CreateObject("CATIA.Application")
                    '            CATIA.Visible = True
                    '            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MOBILE\TUENKERS\SZKR\")

                    '        End If
                    '        On Error GoTo 0
                    '    End If
                    'End If
                End If
                        End If
                    End If


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class