Public Class Parcatipi000
    Dim CATIA As Object
    Dim Mydocument
    Dim PartFactory
    Dim iPartDoc
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        'Welding Gun
        WeldingGun.Show()
        Me.Close()

    End Sub

    Private Sub Parcatipi000_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        'Axe Vehicle
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\AXE VEHICLE\Axe_vehicule.CATPart")

        On Error GoTo 0
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        'Guijon
        Guijon.Show()
        Me.Close()

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        'Human
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\HUMAN_MODEL\ERGO_WINDOW_HUMAN\ERGO_WINDOW_with_HUMAN2.CATProduct")

        On Error GoTo 0
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'Rivet
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\RIVET_GUN\TAURUS_RIVETAGE.CATPart")

        On Error GoTo 0

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()

    End Sub
End Class