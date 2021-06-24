Public Class RailChariot
    Dim picture1 As Image
    Dim picture2 As Image
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Dim Msg
        If ComboBox2.SelectedItem.ToString = "Rail" Then
            Dim urlFile As String = "I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\RAIL_CHARIOT\BOSCH-REXROTH\RAIL\Bosch Rexroth Rail.url"
            Process.Start(urlFile)
        Else
            'Get CATIA Or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\RAIL_CHARIOT\BOSCH-REXROTH\CHARIOT\KWD_035_FNS_CS.CATPart")

            On Error GoTo 0
        End If
    End Sub

    Private Sub RailChariot_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("BOSCH-REXROTH")
        ComboBox2.Items.Add("Chariot")
        ComboBox2.Items.Add("Rail")


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Me.ComboBox2.Items.Clear()
        ComboBox1.Focus()
        ComboBox2.ResetText()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\RAIL_CHARIOT\BOSCH-REXROTH\CHARIOT\Chariot.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\RAIL_CHARIOT\BOSCH-REXROTH\Rail_Chariot BOSCH-REXROTH.PNG") '2

            Image1.Image = picture1
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox2.Focus()
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\RAIL_CHARIOT\BOSCH-REXROTH\CHARIOT\Chariot.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\201_PURCHASED EQUIPMENTS\RAIL_CHARIOT\BOSCH-REXROTH\Rail_Chariot BOSCH-REXROTH.PNG") '2

            If ComboBox2.SelectedItem.ToString = "Rail" Then
                Button1.Text = "Url'ye Git"
                Image1.Image = picture2
            Else
                Button1.Text = "Parçayı CATIA'da Aç"
                Image1.Image = picture1

            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi200.Show()
        Me.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click

    End Sub
End Class