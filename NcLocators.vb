Public Class NcLocators
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image
    Dim picture5 As Image
    Dim picture6 As Image
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        ComboBox4.ResetText()

        Me.ComboBox2.Items.Clear()
        Me.ComboBox3.Items.Clear()
        Me.ComboBox4.Items.Clear()
        Me.ComboBox5.Items.Clear()
        ComboBox5.ResetText()
        ComboBox2.Items.Add("AGK")
        ComboBox2.Items.Add("CKK")
        ComboBox2.Items.Add("NC Accessory")
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_BOSCH-REXROTH.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_BOSCH-REXROTH.PNG") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R037551034.PNG") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R039660545.PNG") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R341703809.PNG") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R344702001.PNG") '6
            Image1.Image = picture1
        Catch ex As Exception
        End Try
    End Sub

    Private Sub NcLocators_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("BOSCH-REXROTH")
        'ComboBox2.Items.Add("AGK")
        'ComboBox2.Items.Add("CKK")
        'ComboBox2.Items.Add("NC Accessory")
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        ComboBox3.ResetText()
        ComboBox4.ResetText()
        Me.ComboBox5.Items.Clear()
        ComboBox5.ResetText()
        Me.ComboBox3.Items.Clear()
        Me.ComboBox4.Items.Clear()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_BOSCH-REXROTH.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_BOSCH-REXROTH.PNG") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R037551034.PNG") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R039660545.PNG") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R341703809.PNG") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R344702001.PNG") '6

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

            If ComboBox2.SelectedItem.ToString = "NC Accessory" Then
                Label1.Text = "Parça Kodu Seçiniz"
                Image1.Image = picture3
                ComboBox3.Items.Add("R037551034")
                ComboBox3.Items.Add("R039660545")
                ComboBox3.Items.Add("R341703809")
                ComboBox3.Items.Add("R344702001")
                'Label2.Text = "Parça Kodu Seçiniz"
                'Label1.Visible = False
                Label4.Visible = False
                Label5.Visible = False
                'ComboBox3.Visible = False
                ComboBox4.Visible = False
                ComboBox5.Visible = False
            ElseIf ComboBox2.SelectedItem.ToString = "AGK" Then
                ComboBox3.Items.Add("420 [mm]")
                ComboBox3.Items.Add("570 [mm]")
                ComboBox3.Items.Add("1020 [mm]")

                Label1.Text = "Strok Seçiniz"
                'Label1.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                'ComboBox3.Visible = True
                ComboBox4.Visible = True
                ComboBox5.Visible = True
                Image1.Image = picture1
            ElseIf ComboBox2.SelectedItem.ToString = "CKK" Then
                ComboBox3.Items.Add("170 [mm]")
                ComboBox3.Items.Add("210 [mm]")
                ComboBox3.Items.Add("250 [mm]")
                ComboBox3.Items.Add("330 [mm]")
                ComboBox3.Items.Add("410 [mm]")
                ComboBox3.Items.Add("530 [mm]")
                ComboBox3.Items.Add("610 [mm]")
                Label1.Text = "Strok Seçiniz"
                'Label1.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                ' ComboBox3.Visible = True
                ComboBox4.Visible = True
                ComboBox5.Visible = True
                Image1.Image = picture2
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox4.ResetText()
        Me.ComboBox4.Items.Clear()
        Me.ComboBox5.Items.Clear()
        ComboBox5.ResetText()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_BOSCH-REXROTH.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_BOSCH-REXROTH.PNG") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R037551034.PNG") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R039660545.PNG") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R341703809.PNG") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R344702001.PNG") '6
            If ComboBox2.SelectedItem.ToString = "AGK" Then
                ComboBox4.Items.Add("MF01")
                ComboBox4.Items.Add("RV01")
            ElseIf ComboBox2.SelectedItem.ToString = "CKK" Then
                ComboBox4.Items.Add("MF01")
                ComboBox4.Items.Add("RV01")
                ComboBox4.Items.Add("RV02")
                ComboBox4.Items.Add("RV03")
                ComboBox4.Items.Add("RV04")
            ElseIf ComboBox2.SelectedItem.ToString = "NC Accessory" Then
                Label1.Text = "Parça Kodu Seçiniz"
                If ComboBox3.SelectedItem.ToString = "R037551034" Then

                    Image1.Image = picture3
                ElseIf ComboBox3.SelectedItem.ToString = "R039660545" Then
                    Image1.Image = picture4
                ElseIf ComboBox3.SelectedItem.ToString = "R341703809" Then
                    Image1.Image = picture5
                ElseIf ComboBox3.SelectedItem.ToString = "R344702001" Then
                    Image1.Image = picture6
                End If
            End If
            If ComboBox2.SelectedItem.ToString = "AGK" Then
                ComboBox5.Items.Add("MA01")
                ComboBox5.Items.Add("RV01")

            ElseIf ComboBox2.SelectedItem.ToString = "CKK" Then
                ComboBox5.Items.Add("MF01")
                ComboBox5.Items.Add("RV01")
                ComboBox5.Items.Add("RV02")
                ComboBox5.Items.Add("RV03")
                ComboBox5.Items.Add("RV04")
            End If
            'ComboBox4.Items.Add("R12")
            'ComboBox4.Items.Add("R20")
            'Label4.Text = "Ekipman Tipi"
            'Image1.Image = picture2
        Catch ex As Exception
        End Try

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Me.ComboBox5.Items.Clear()
        ComboBox5.ResetText()
        If ComboBox2.SelectedItem.ToString = "AGK" Then
            If ComboBox4.SelectedItem.ToString = "MF01" Then
                ComboBox5.Items.Add("MA01")
                ComboBox5.Items.Add("MA02")
                ComboBox5.Items.Add("MA03")
            Else
                ComboBox5.Items.Add("MA02")
            End If
        ElseIf ComboBox2.SelectedItem.ToString = "CKK" Then
            ComboBox5.Items.Add("MF01")
            ComboBox5.Items.Add("RV01")
            ComboBox5.Items.Add("RV02")
            ComboBox5.Items.Add("RV03")
            ComboBox5.Items.Add("RV04")
        End If
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

        'If ComboBox1.SelectedItem.ToString = "AGK" Then
        If ComboBox2.SelectedItem.ToString = "AGK" Then
            If ComboBox3.SelectedItem.ToString = "420 [mm]" Then
                If ComboBox4.SelectedItem.ToString = "MF01" Then
                    If ComboBox5.SelectedItem.ToString = "MA01" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA01.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "MA02" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA02.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "MA03" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA03.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox4.SelectedItem.ToString = "RV01" Then
                    'If ComboBox5.SelectedItem.ToString = "MA01" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C570_L860_MF01_MA01.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    If ComboBox5.SelectedItem.ToString = "MA02" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_RV01_MA02.CATPart")

                        On Error GoTo 0
                        'ElseIf ComboBox5.SelectedItem.ToString = "MA03" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA01.CATPart")

                        '    End If
                        '    On Error GoTo 0
                        'End If
                    End If
                End If
            ElseIf ComboBox3.SelectedItem.ToString = "570 [mm]" Then

                If ComboBox4.SelectedItem.ToString = "MF01" Then
                    If ComboBox5.SelectedItem.ToString = "MA01" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C570_L860_MF01_MA01.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "MA02" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C570_L860_RV01_MA02.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "MA03" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C570_L860_MF01_MA03.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox4.SelectedItem.ToString = "RV01" Then
                    'If ComboBox5.SelectedItem.ToString = "MA01" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA01.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    If ComboBox5.SelectedItem.ToString = "MA02" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA01.CATPart")

                        On Error GoTo 0
                        'ElseIf ComboBox5.SelectedItem.ToString = "MA03" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA01.CATPart")

                        '    On Error GoTo 0
                    End If
                End If

            ElseIf ComboBox3.SelectedItem.ToString = "1020 [mm]" Then

                If ComboBox4.SelectedItem.ToString = "MF01" Then
                    If ComboBox5.SelectedItem.ToString = "MA01" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C1020_L1310_MF01_MA01.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "MA02" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C1020_L1310_MF01_MA02.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "MA03" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C1020_L1310_MF01_MA03.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox4.SelectedItem.ToString = "RV01" Then
                    'If ComboBox5.SelectedItem.ToString = "MA01" Then
                    '    'Get CATIA Or Launch it if necessary.
                    '    On Error Resume Next
                    '    CATIA = GetObject(, "CATIA.Application")
                    '    
                    '        CATIA = CreateObject("CATIA.Application")
                    '        CATIA.Visible = True
                    '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA01.CATPart")

                    '    End If
                    '    On Error GoTo 0
                    If ComboBox5.SelectedItem.ToString = "MA02" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C1020_L1310_RV01_MA02.CATPart")

                        On Error GoTo 0
                        'ElseIf ComboBox5.SelectedItem.ToString = "MA03" Then
                        '    'Get CATIA Or Launch it if necessary.
                        '    On Error Resume Next
                        '    CATIA = GetObject(, "CATIA.Application")
                        '    
                        '        CATIA = CreateObject("CATIA.Application")
                        '        CATIA.Visible = True
                        '        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\AGK\AGK_32-10_C420_L710_MF01_MA01.CATPart")

                        '    On Error GoTo 0
                    End If
                End If

                'ComboBox3.Items.Add("170 [mm]")
                'ComboBox3.Items.Add("210 [mm]")
                'ComboBox3.Items.Add("250 [mm]")
                'ComboBox3.Items.Add("330 [mm]")
                'ComboBox3.Items.Add("410 [mm]")
                'ComboBox3.Items.Add("530 [mm]")
                'ComboBox3.Items.Add("610 [mm]")

                'ElseIf ComboBox1.SelectedItem.ToString = "Tuenkers" Then
            End If
        ElseIf ComboBox2.SelectedItem.ToString = "CKK" Then
            If ComboBox3.SelectedItem.ToString = "170 [mm]" Then
                If ComboBox5.SelectedItem.ToString = "MF01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C170_L400_MF01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C170_L400_RV01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV02" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C170_L400_RV02.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV03" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C170_L400_RV03.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV04" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C170_L400_RV04.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox3.SelectedItem.ToString = "210 [mm]" Then
                If ComboBox5.SelectedItem.ToString = "MF01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C210_L440_MF01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C210_L440_RV01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV02" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C210_L440_RV02.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV03" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C210_L440_RV03.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV04" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C210_L440_RV04.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox3.SelectedItem.ToString = "250 [mm]" Then
                If ComboBox5.SelectedItem.ToString = "MF01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C250_L480_MF01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C250_L480_RV01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV02" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C250_L480_RV02.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV03" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C250_L480_RV03.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV04" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C250_L480_RV04.CATPart")

                    On Error GoTo 0
                End If
            ElseIf ComboBox3.SelectedItem.ToString = "330 [mm]" Then
                If ComboBox5.SelectedItem.ToString = "MF01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C330_L560_MF01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C330_L560_RV01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV02" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C330_L560_RV02.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV03" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C330_L560_RV03.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV04" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C330_L560_RV04.CATPart")

                    On Error GoTo 0
                End If



            ElseIf ComboBox3.SelectedItem.ToString = "410 [mm]" Then
                If ComboBox5.SelectedItem.ToString = "MF01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C410_L640_MF01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C410_L640_RV01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV02" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C410_L640_RV02.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV03" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C410_L640_RV03.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV04" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C410_L640_RV04.CATPart")

                    On Error GoTo 0
                End If



            ElseIf ComboBox3.SelectedItem.ToString = "530 [mm]" Then
                If ComboBox5.SelectedItem.ToString = "MF01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C530_L760_MF01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV01" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C530_L760_RV01.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV02" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C530_L760_RV02.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV03" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C530_L760_RV03.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox5.SelectedItem.ToString = "RV04" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C530_L760_RV04.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "610 [mm]" Then
                    If ComboBox5.SelectedItem.ToString = "MF01" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C610_L840_MF01.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "RV01" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C610_L840_RV01.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "RV02" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C610_L840_RV02.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "RV03" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C610_L840_RV03.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox5.SelectedItem.ToString = "RV04" Then
                        'Get CATIA Or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\CKK\20-145\CKK_20-145_C610_L840_RV04.CATPart")

                        On Error GoTo 0
                    End If

                End If


            End If
        ElseIf ComboBox2.SelectedItem.ToString = "NC Accessory" Then

            If ComboBox3.SelectedItem.ToString = "R037551034" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R037551034.CATPart")

                On Error GoTo 0
            ElseIf ComboBox3.SelectedItem.ToString = "R039660545" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R039660545.CATPart")

                On Error GoTo 0
            ElseIf ComboBox3.SelectedItem.ToString = "R341703809" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R341703809.CATPart")

                On Error GoTo 0
            ElseIf ComboBox3.SelectedItem.ToString = "R344702001" Then
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\NC LOCATORS\BOSCH-REXROTH\NC Accessory\R344702001.CATPart")

                On Error GoTo 0
            End If
        End If

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class