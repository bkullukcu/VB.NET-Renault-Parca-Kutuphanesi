Public Class Profiles
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image
    Dim picture5 As Image
    Dim picture6 As Image
    Dim picture7 As Image
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try

            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\HEA\HEA.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\IPE\IPE.png") '2
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\L_egal\L_egal.png") '3
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\L_inegal\L_inegal.png") '4
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\TUBE_carre_rect\TUBE_Carre_rect.png") '5
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\TUBE_rond_acier\TUBE_rond_acier.png") '6
            Dim picture7 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\UPN\UPN.png") '7

            ComboBox1.Focus()
            With Me.ComboBox1
                Select Case ComboBox1.SelectedItem
                    Case "HEA"
                        Image1.Image = picture1
                    Case "IPE"
                        Image1.Image = picture2
                    Case "L_EGAL"
                        Image1.Image = picture3
                    Case "L_INEGAL"
                        Image1.Image = picture4
                    Case "TUBE_CARRE_RECT"
                        Image1.Image = picture5
                    Case "TUBE_ROND_ACIER"
                        Image1.Image = picture6
                    Case "UPN"
                        Image1.Image = picture7
                End Select
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Profiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("HEA")
        ComboBox1.Items.Add("IPE")
        ComboBox1.Items.Add("L_EGAL")
        ComboBox1.Items.Add("L_INEGAL")
        ComboBox1.Items.Add("TUBE_CARRE_RECT")
        ComboBox1.Items.Add("TUBE_ROND_ACIER")
        ComboBox1.Items.Add("UPN")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi001.Show()
        Me.Hide()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        With Me.ComboBox1
            Select Case ComboBox1.SelectedItem
                Case "HEA"
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\HEA\profil_HEA.CATPart")

                    On Error GoTo 0
                Case "IPE"
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\IPE\profil_IPE.CATPart")

                    On Error GoTo 0
                Case "L_EGAL"
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\L_egal\L_egal.CATPart")

                    On Error GoTo 0
                Case "L_INEGAL"
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\L_inegal\profil_L_inegal.CATPart")

                    On Error GoTo 0
                Case "TUBE_CARRE_RECT"
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\TUBE_carre_rect\TUBE_carre.CATPart")

                    On Error GoTo 0
                Case "TUBE_ROND_ACIER"
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\TUBE_rond_acier\tube_acier.CATPart")

                    On Error GoTo 0
                Case "UPN"
                    'Get CATIA or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PROFILES\UPN\profile_UPN.CATPart")

                    On Error GoTo 0
            End Select
        End With

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class