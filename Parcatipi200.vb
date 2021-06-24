Public Class Parcatipi200
    Dim CATIA As Object
    Dim Mydocument
    Dim PartFactory
    Dim iPartDoc
    Private Sub Parcatipi200_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Anneau
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\ANNEAU DE LEVAGE\07710-3035.CATPart")

        On Error GoTo 0
        ' Add a new Part


        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Cylinder
        Cylinder.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Double Colonne
        DoubleColonne.Show()
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Hardware
        Hardware.Show()
        Me.Close()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'Nc Locators
        NcLocators.Show()
        Me.Close()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        'Pilotes Mobile
        PilotesMobile.Show()
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Pilotes Multi Function
        PilotesMultiFunction.Show()
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Rail Chariot
        RailChariot.Show()

        Me.Close()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'Sensor
        Sensor.Show()
        Me.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Serrages
        Serrages.Show()
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

    End Sub
End Class