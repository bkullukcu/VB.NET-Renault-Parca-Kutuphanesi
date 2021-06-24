Public Class Sensor
    Dim picture1 As Image
    Dim picture2 As Image

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi200.Show()
        Me.Close()
    End Sub

    Private Sub StoolHightener_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("IFM")
        ComboBox1.Items.Add("Senstronic")

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.Focus()
        ComboBox2.ResetText()
        ComboBox2.Items.Clear()

        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\IFM\Laser_IFM.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\SENSTRONIC\Inductive_SENSTRONIC.PNG") '2
            Select Case ComboBox1.SelectedItem
                Case "IFM"
                    Image1.Image = picture1
                    ComboBox2.Items.Add("E2d101")
                    ComboBox2.Items.Add("E20938")
                    ComboBox2.Items.Add("E21171")
                    ComboBox2.Items.Add("O1D100")
                Case "Senstronic"
                    ComboBox2.Items.Add("A120212l181011KTK")
                    ComboBox2.Items.Add("SDPL26-18PL")
                    Image1.Image = picture2
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
            Case "IFM"
                'Image1.Image = picture1
                If ComboBox2.SelectedItem = "E2d101" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\IFM\Laser\E2d101.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "E20938" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\IFM\Laser\E20938.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "E21171" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\IFM\Laser\E21171.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "O1D100" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\IFM\Laser\IFM_DETECTEUR_PHOTOELEC_O1D100.CATPart")

                    On Error GoTo 0
                End If
            Case "Senstronic"
                If ComboBox2.SelectedItem = "A120212l181011KTK" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\SENSTRONIC\Inductive\A120212l181011KTK.CATPart")

                    On Error GoTo 0
                ElseIf ComboBox2.SelectedItem = "SDPL26-18PL" Then
                    'Get CATIA Or Launch it if necessary.
                    On Error Resume Next
                    CATIA = GetObject(, "CATIA.Application")

                    CATIA = CreateObject("CATIA.Application")
                    CATIA.Visible = True
                    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\SENSTRONIC\Inductive\SDPL26-18PL.CATPart")

                    On Error GoTo 0
                End If


        End Select

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\IFM\Laser_IFM.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\SENSOR\SENSTRONIC\Inductive_SENSTRONIC.PNG") '2
            Select Case ComboBox1.SelectedItem

                Case "IFM"
                    Image1.Image = picture1

                Case "Senstronic"

                    Image1.Image = picture2
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    'Private Sub ''Image1_Click(sender As Object, e As EventArgs)

    'End Sub
End Class