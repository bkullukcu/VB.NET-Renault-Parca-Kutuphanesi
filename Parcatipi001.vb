Public Class Parcatipi001
    Dim CATIA As Object
    Dim Mydocument
    Dim PartFactory
    Dim iPartDoc
    Private Sub Parcatipi001_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Appui
        Appui.Show()
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Pilote Spesific
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\PILOTE SPESIFIC\EXXXXXXXXX.CATPart")

        On Error GoTo 0
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Profiles
        Profiles.Show()
        Me.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\001_CHANGEABLE TEMPLATES\MARBRE")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub
End Class