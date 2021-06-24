Public Class Pneumatique
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
    Private Sub Pneumatique_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("E140001000_Pneu_Equip_type_01")
        ComboBox1.Items.Add("E140001005_Pneu_Equip_type_02")
        ComboBox1.Items.Add("E140001010_Pneu_Equip_type_11")
        ComboBox1.Items.Add("E140001015_Pneu_Equip_type_12")
        ComboBox1.Items.Add("E140001019_Pneu_Equip_Mini_Support")
        ComboBox1.Items.Add("E140001020_Pneu_Equip_Mini_Type_01")
        ComboBox1.Items.Add("E140001025_Pneu_Equip_Mini_Type_02")
        ComboBox1.Items.Add("E140001030_Pneu_Equip_Mini_Type_03")
        ComboBox1.Items.Add("E140001080_Pneu_Equip_Collon_Type_01")
        ComboBox1.Items.Add("E140001090_GJN_Barrette_de_Masse")
        ComboBox1.Items.Add("E140001095_GJN_Plot_de_Masse")
        Image1.Image = picture1

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001000_Festo_Module_type_01.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001019_Pneumatique_Equip_Mini_Support.jpg") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001020_Pneumatique_Equip_Mini_Type_01.jpg") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001025_Pneumatique_Equip_Mini_Type_02.jpg") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001030_Pneumatique_Equip_Mini_Type_03.jpg") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001080_Pneu_Equip_Collon_Type_01.jpg") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001090_GJN_Barrette_de_Masses.jpg") '7
            Dim picture8 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001095_GJN_Plot_de_Masse.jpg") '8
            Dim picture9 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001005_Pneu_Equip_type_02\3d_\E140001005.jpg") '6
            Dim picture10 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001010_Pneu_Equip_type_11\E140001010.jpg") '7
            Dim picture11 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001015_Pneu_Equip_type_12\E140001015.jpg") '8
            Select Case ComboBox1.SelectedItem.ToString
                Case "E140001000_Pneu_Equip_type_01"
                    Image1.Image = picture1
                Case "E140001005_Pneu_Equip_type_02"
                    Image1.Image = picture9
                Case "E140001010_Pneu_Equip_type_11"
                    Image1.Image = picture10
                Case "E140001015_Pneu_Equip_type_12"
                    Image1.Image = picture11

                Case "E140001019_Pneu_Equip_Mini_Support"
                    Image1.Image = picture2

                Case "E140001020_Pneu_Equip_Mini_Type_01"
                    Image1.Image = picture3

                Case "E140001025_Pneu_Equip_Mini_Type_02"
                    Image1.Image = picture4

                Case "E140001030_Pneu_Equip_Mini_Type_03"
                    Image1.Image = picture5

                Case "E140001080_Pneu_Equip_Collon_Type_01"
                    Image1.Image = picture6

                Case "E140001090_GJN_Barrette_de_Masse"
                    Image1.Image = picture7

                Case "E140001095_GJN_Plot_de_Masse"
                    Image1.Image = picture8
            End Select

        Catch ex As Exception
            'MsgBox("Bir ya da birkaç fotoğraf eksik")
        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Select Case ComboBox1.SelectedItem
            Case "E140001000_Pneu_Equip_type_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001000_Pneu_Equip_type_01\3d_\Festo_Groupe_Type_01.CATProduct")

                On Error GoTo 0
            Case "E140001005_Pneu_Equip_type_02"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001005_Pneu_Equip_type_02\3d_\Festo_Groupe_Type_02 .CATProduct")

                On Error GoTo 0
            Case "E140001010_Pneu_Equip_type_11"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001010_Pneu_Equip_type_11\3D\Festo_Groupe_Type_11.CATProduct")

                On Error GoTo 0
            Case "E140001015_Pneu_Equip_type_12"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001015_Pneu_Equip_type_12\3d\Festo_Groupe_Type_12 .CATProduct")

                On Error GoTo 0
            Case "E140001019_Pneu_Equip_Mini_Support"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001019_Pneu_Equip_Mini_Support\3D\Pneu_Equip_Mini_Ref_conception.CATProduct")

                On Error GoTo 0

            Case "E140001020_Pneu_Equip_Mini_Type_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001020_Pneu_Equip_Mini_Type_01\3D\Pneu_Equip_Mini_Type_01.CATProduct")

                On Error GoTo 0
            Case "E140001025_Pneu_Equip_Mini_Type_02"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001025_Pneu_Equip_Mini_Type_02\3D\Pneu_Equip_Mini_Type_02.CATProduct")

                On Error GoTo 0
            Case "E140001030_Pneu_Equip_Mini_Type_03"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001030_Pneu_Equip_Mini_Type_03\3D\Pneu_Equip_Mini_Type_03.CATProduct")

                On Error GoTo 0
            Case "E140001080_Pneu_Equip_Collon_Type_01"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001080_Pneu_Equip_Collon_Type_01\3D\Pneumatique_Equip_Collon_Type_01.CATProduct")

                On Error GoTo 0
            Case "E140001090_GJN_Barrette_de_Masse"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001090_GJN_Barrette_de_Masse\3d\GJN_Barrette_de_Masses.CATProduct")

                On Error GoTo 0
            Case "E140001095_GJN_Plot_de_Masse"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\E140001095_GJN_Plot_de_Masse\3d\GJN_Plot_de_Masse_01.CATProduct")

                On Error GoTo 0
        End Select
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class