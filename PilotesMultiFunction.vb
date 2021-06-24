Public Class PilotesMultiFunction
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image

    'Private Sub Image1_Click(sender As Object, e As EventArgs) Handles Image1.Click
    ' End Sub

    Private Sub PinBracket_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("DELTA")
        ComboBox1.Items.Add("DESTACO")
        ComboBox1.Items.Add("SMC")
        'ComboBox1.Items.Add("Tuenkers")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.Focus()
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        Me.ComboBox3.Items.Clear()
        Me.ComboBox2.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox4.ResetText()


        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML3163.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\82P50-800.PNG") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P63-805\82P63-805.PNG") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63.PNG") '4
            With Me.ComboBox1

                Select Case ComboBox1.SelectedItem
                    Case "DELTA"
                        ComboBox4.Visible = False
                        Label4.Visible = False
                        ComboBox2.Visible = True
                        Image1.Image = picture1
                        Label2.Text = "Pilot Tipi Seçiniz"
                        Label2.Visible = True
                        ComboBox3.Visible = False
                        Label3.Visible = False
                        'ComboBox2.Items.Add("")
                        ComboBox2.Items.Add("ML316317-99")
                        ComboBox2.Items.Add("ML316318-99")
                        ComboBox2.Items.Add("ML316319-99")
                        ComboBox2.Items.Add("ML316320-99")
                        ComboBox2.Items.Add("ML316321-99")
                    Case "DESTACO"
                        ComboBox4.Visible = True
                        Label4.Visible = True
                        Label2.Text = "Ekipman Tipi Seçiniz"
                        Label3.Text = "Çap Seçiniz"
                        ComboBox2.Visible = True
                        ComboBox3.Visible = True
                        Label2.Visible = True
                        Label3.Visible = True
                        Image1.Image = picture2
                        ComboBox2.Items.Add("82P50-800")
                        ComboBox2.Items.Add("82P63-805")
                    Case "SMC"
                        'Label2.Text = "Delik Çapı Seçiniz"
                        Label2.Text = "Pilot Tipi Seçiniz"
                        ComboBox4.Visible = False
                        Label4.Visible = False
                        Image1.Image = picture4
                        ComboBox2.Visible = True
                        ComboBox2.Items.Add("CKQRM63 - 298.0R - X2309")
                        ComboBox2.Items.Add("CKQRM63 - 298.0R - X2308")
                        ComboBox2.Items.Add("CKQRM63 -298D-X2309")
                        ComboBox2.Items.Add("CKQRM63 - 203.0R - X2309")
                        ComboBox2.Items.Add("CKQRM63 - 203.0R - X2308")
                        ComboBox2.Items.Add("CKQRM63 -203D-X2309")
                        ComboBox2.Items.Add("CKQRM63 - 198.0R - X2309")
                        ComboBox2.Items.Add("CKQRM63 - 198.0R - X2308")
                        ComboBox2.Items.Add("CKQRM63 -198D-X2309")
                        ComboBox2.Items.Add("CKQRM63 -198D-X2308")
                        ComboBox2.Items.Add("CKQRM63 - 158.0R - X2309")
                        ComboBox2.Items.Add("CKQRM63 - 158.0R - X2308")
                        Label2.Visible = True
                        ComboBox3.Visible = False
                        Label3.Visible = False
                    Case "Tuenkers"
                        Label2.Text = "Çap Seçiniz"
                        Label3.Text = "Pilot Tipi Seçiniz"
                        Image1.Image = picture3
                        ComboBox2.Visible = True
                        ComboBox3.Visible = True
                        Label2.Visible = True
                        Label3.Visible = True
                End Select
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox3.ResetText()
        Me.ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox4.ResetText()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML3163.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\82P50-800.PNG") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P63-805\82P63-805.PNG") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63.PNG") '4
            ComboBox2.Focus()
            ComboBox3.ResetText()
            With Me.ComboBox2
                'If ComboBox1.SelectedItem = "DELTA" Then
                '    If ComboBox2.SelectedItem = "Yan" Then
                '        Image1.Image = picture3
                '    ElseIf ComboBox2.SelectedItem = "Ön" Then
                '        Image1.Image = picture1
                '    End If

                If ComboBox1.SelectedItem = "DESTACO" Then
                    If ComboBox2.SelectedItem = "82P50-800" Then
                        Label3.Text = "Çap Seçiniz"

                        Image1.Image = picture2
                        ComboBox4.Visible = True
                        Label4.Visible = True
                        ComboBox3.Items.Add("15.8")
                        ComboBox3.Items.Add("19.8")
                        ComboBox3.Items.Add("20.3")
                        ComboBox3.Items.Add("29")
                        ComboBox3.Items.Add("29.8")

                    ElseIf ComboBox2.SelectedItem = "82P63-805" Then
                        ComboBox4.Visible = False
                        Label3.Text = "Pilot Tipi Seçiniz"
                        Label4.Visible = False
                        Image1.Image = picture3
                        ComboBox3.Items.Add("82P63-805-L1R0101000_1")
                        ComboBox3.Items.Add("82P63-805-L1R0101000_2")
                        ComboBox3.Items.Add("82P63-805-L1R0401000_1")
                        ComboBox3.Items.Add("82P63-805-L1R0401000_2")
                    End If

                ElseIf ComboBox1.SelectedItem = "SMC" Then
                    Image1.Image = picture4
                    ComboBox4.Visible = False
                    Label4.Visible = False
                    If ComboBox2.SelectedItem = "CKQRM63 - 298.0R - X2309" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 - 298.0R - X2308" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 -298D-X2309" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 - 203.0R - X2309" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 - 203.0R - X2308" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 -203D-X2309" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 - 198.0R - X2309" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 - 198.0R - X2308" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 -198D-X2309" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 -198D-X2308" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 - 158.0R - X2309" Then
                    ElseIf ComboBox2.SelectedItem = "CKQRM63 - 158.0R - X2308" Then
                    End If

                    'ElseIf ComboBox1.SelectedItem = "Tuenkers" Then
                    '    If ComboBox2.SelectedItem = "Yan" Then
                    '        Image1.Image = picture4
                    '    ElseIf ComboBox2.SelectedItem = "Ön" Then
                    '        Image1.Image = picture4
                    '    End If
                End If
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox4.Items.Clear()
        ComboBox4.ResetText()

        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML3163.PNG") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\82P50-800.PNG") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P63-805\82P63-805.PNG") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63.PNG") '4
            ComboBox3.Focus()
            With Me.ComboBox3
                If ComboBox1.SelectedItem = "DESTACO" Then
                    If ComboBox2.SelectedItem = "82P50-800" Then
                        If ComboBox3.SelectedItem = "15.8" Then
                            ComboBox4.Items.Add("82P50-800L10A01000")
                            ComboBox4.Items.Add("82P50-800L10A01001")
                            ComboBox4.Items.Add("82P50-800L10A01010")
                            ComboBox4.Items.Add("82P50-800L10A01011")
                            ComboBox4.Items.Add("82P50-800L10A01020")
                            ComboBox4.Items.Add("82P50-800L10A01021")
                            ComboBox4.Items.Add("82P50-800L10A01030")
                            ComboBox4.Items.Add("82P50-800L10A01031")
                            ComboBox4.Items.Add("82P50-800L10B01000")
                            ComboBox4.Items.Add("82P50-800L10B01010")
                            ComboBox4.Items.Add("82P50-800L10B01020")
                            ComboBox4.Items.Add("82P50-800L10B01030")
                        ElseIf ComboBox3.SelectedItem = "19.8" Then
                            ComboBox4.Items.Add("82P50-800L10B02000")
                            ComboBox4.Items.Add("82P50-800L10B02001")
                            ComboBox4.Items.Add("82P50-800L10B02010")
                            ComboBox4.Items.Add("82P50-800L10B02011")
                            ComboBox4.Items.Add("82P50-800L10B02020")
                            ComboBox4.Items.Add("82P50-800L10B02021")
                            ComboBox4.Items.Add("82P50-800L10B02030")
                            ComboBox4.Items.Add("82P50-800L10B02031")
                            ComboBox4.Items.Add("82P50-800L10C02000")
                            ComboBox4.Items.Add("82P50-800L10C02001")
                            ComboBox4.Items.Add("82P50-800L10C02010")
                            ComboBox4.Items.Add("82P50-800L10C02011")
                            ComboBox4.Items.Add("82P50-800L10C02020")
                            ComboBox4.Items.Add("82P50-800L10C02021")
                            ComboBox4.Items.Add("82P50-800L10C02030")
                            ComboBox4.Items.Add("82P50-800L10C02031")
                        ElseIf ComboBox3.SelectedItem = "20.3" Then
                            ComboBox4.Items.Add("82P50-800L10B03000")
                            ComboBox4.Items.Add("82P50-800L10B03001")
                            ComboBox4.Items.Add("82P50-800L10B03010")
                            ComboBox4.Items.Add("82P50-800L10B03011")
                            ComboBox4.Items.Add("82P50-800L10B03020")
                            ComboBox4.Items.Add("82P50-800L10B03021")
                            ComboBox4.Items.Add("82P50-800L10B03030")
                            ComboBox4.Items.Add("82P50-800L10B03031")
                            ComboBox4.Items.Add("82P50-800L10C03000")
                            ComboBox4.Items.Add("82P50-800L10C03001")
                            ComboBox4.Items.Add("82P50-800L10C03010")
                            ComboBox4.Items.Add("82P50-800L10C03011")
                            ComboBox4.Items.Add("82P50-800L10C03020")
                            ComboBox4.Items.Add("82P50-800L10C03021")
                            ComboBox4.Items.Add("82P50-800L10C03030")
                            ComboBox4.Items.Add("82P50-800L10C03031")
                        ElseIf ComboBox3.SelectedItem = "29" Then
                            ComboBox4.Items.Add("82P50-800L10C04040")
                        ElseIf ComboBox3.SelectedItem = "29.8" Then
                            ComboBox4.Items.Add("82P50-800L10B04000")
                            ComboBox4.Items.Add("82P50-800L10B04001")
                            ComboBox4.Items.Add("82P50-800L10B04010")
                            ComboBox4.Items.Add("82P50-800L10B04011")
                            ComboBox4.Items.Add("82P50-800L10B04020")
                            ComboBox4.Items.Add("82P50-800L10B04021")
                            ComboBox4.Items.Add("82P50-800L10B04030")
                            ComboBox4.Items.Add("82P50-800L10B04031")
                            ComboBox4.Items.Add("82P50-800L10C04000")
                            ComboBox4.Items.Add("82P50-800L10C04001")
                            ComboBox4.Items.Add("82P50-800L10C04010")
                            ComboBox4.Items.Add("82P50-800L10C04011")
                            ComboBox4.Items.Add("82P50-800L10C04020")
                            ComboBox4.Items.Add("82P50-800L10C04021")
                            ComboBox4.Items.Add("82P50-800L10C04030")
                            ComboBox4.Items.Add("82P50-800L10C04031")
                        End If
                    ElseIf ComboBox2.SelectedItem = "82P63-805" Then
                        ComboBox3.Items.Add("82P63-805L1R0101000")
                        ComboBox3.Items.Add("82P63-805L1R0101000")
                        ComboBox3.Items.Add("82P63-805L1R0401000")
                        ComboBox3.Items.Add("82P63-805L1R0401000")
                    End If
                End If
                Select Case ComboBox1.SelectedItem
                    Case "DELTA"
                        ComboBox2.Items.Add("ML316317-99")
                        ComboBox2.Items.Add("ML316318-99")
                        ComboBox2.Items.Add("ML316319-99")
                        ComboBox2.Items.Add("ML316320-99")
                        ComboBox2.Items.Add("ML316321-99")
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
        If ComboBox1.SelectedItem = "DESTACO" Then
            If ComboBox2.SelectedItem = "82P50-800" Then
                If ComboBox3.SelectedItem = "15.8" Then
                    If ComboBox4.SelectedItem.ToString = "82P50-800L10A01000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10A01001" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01001.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10A01010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10A01011" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01011.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10A01020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10A01021" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01021.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10A01030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01030.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10A01031" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10A01031.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B01000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10B01000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B01010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10B01010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B01020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10B01020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B01030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\15,8\82P50-800L10B01030.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "19.8" Then
                    If ComboBox4.SelectedItem.ToString = "82P50-800L10B02000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B02001" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02001.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B02010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B02011" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02011.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B02020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B02021" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02021.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B02030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02030.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B02031" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10B02031.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02001" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02001.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02011" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02011.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02021" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02021.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02030.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C02031" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\19,8\82P50-800L10C02031.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "20.3" Then
                    If ComboBox4.SelectedItem.ToString = "82P50-800L10B03000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20.3\82P50-800L10B03000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B03001" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10B03001.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B03010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10B03010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B03011" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10B03011.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B03020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10B03020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B03021" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10B03021.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B03030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10B03030.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B03031" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10B03031.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10BC03000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03001" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10C03001.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10C03010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03011" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10C03011.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10C03020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03021" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10C03021.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10C03030.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C03031" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\20,3\82P50-800L10C03031.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "29" Then
                    If ComboBox4.SelectedItem.ToString = "82P50-800L10C04040" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29\82P50-800L10C04040.CATPart")

                        On Error GoTo 0
                    End If
                ElseIf ComboBox3.SelectedItem = "29.8" Then
                    If ComboBox4.SelectedItem.ToString = "82P50-800L10B04000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B04001" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04001.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B04010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B04011" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04011.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B04020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B04021" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04021.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B04030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04030.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10B04031" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10B04031.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04000" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04000.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04001" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04001.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04010" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04010.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04011" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04011.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04020" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04020.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04021" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04021.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04030" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04030.CATPart")

                        On Error GoTo 0
                    ElseIf ComboBox4.SelectedItem.ToString = "82P50-800L10C04031" Then
                        'Get CATIA or Launch it if necessary.
                        On Error Resume Next
                        CATIA = GetObject(, "CATIA.Application")

                        CATIA = CreateObject("CATIA.Application")
                        CATIA.Visible = True
                        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P50-800\29,8\82P50-800L10C04031.CATPart")

                        On Error GoTo 0
                    End If
                End If
            ElseIf ComboBox2.SelectedItem = "82P63-805" Then
                If ComboBox3.SelectedItem.ToString = "82P63-805-L1R0101000_1" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P63-805\82P63-805-L1R0101000_1\82P63-805L1R0101000.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "82P63-805-L1R0101000_2" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P63-805\82P63-805-L1R0101000_2\82P63-805L1R0101000.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "82P63-805-L1R0401000_1" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P63-805\82P63-805-L1R0401000_1\82P63-805L1R0401000.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox3.SelectedItem.ToString = "82P63-805-L1R0401000_2" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DESTACO\82P63-805\82P63-805-L1R0401000_2\82P63-805L1R0401000.CATPart")
                    On Error GoTo 0
                End If
            End If
        ElseIf ComboBox1.SelectedItem.ToString = "SMC" Then
            If ComboBox2.SelectedItem = "CKQRM63 - 298.0R - X2309" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63-298R-X2309.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 - 298.0R - X2308" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 - 298.0R - X2308.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 -298D-X2309" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 -298D-X2309.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 - 203.0R - X2309" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 - 203.0R - X2309.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 - 203.0R - X2308" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 - 203.0R - X2308.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 -203D-X2309" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 -203D-X2309.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 - 198.0R - X2309" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 - 198.0R - X2309.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 - 198.0R - X2308" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 - 198.0R - X2308.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 -198D-X2309" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 -198D-X2309.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 -198D-X2308" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 -198D-X2308.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 - 158.0R - X2309" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 - 158.0R - X2309.CATPart")
                On Error GoTo 0
            ElseIf ComboBox2.SelectedItem = "CKQRM63 - 158.0R - X2308" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\SMC\CKQRM63 - 158.0R - X2308.CATPart")
                On Error GoTo 0
            End If
        End If

        Select Case ComboBox1.SelectedItem.ToString
            Case "DELTA"
                If ComboBox2.SelectedItem = "ML316317-99" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML316317-99.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "ML316318-99" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML316318-99.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "ML316319-99" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML316319-99.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "ML316320-99" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML316320-99.CATPart")
                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "ML316321-99" Then
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\PILOTES MULTI FONCTION\DELTA\ML3163\ML316321-99.CATPart")
                    On Error GoTo 0
                End If
        End Select
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi200.Show()
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class