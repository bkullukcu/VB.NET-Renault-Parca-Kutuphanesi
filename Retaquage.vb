Public Class Retaquage
    Private Sub Retaquage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("RTQ_ROUE")
        ComboBox1.Items.Add("Groupe_Fix_Retaquage")

    End Sub
    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

        End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00100_RETAQUAGE_EQUIPMENT\E140000100.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00100_RETAQUAGE_EQUIPMENT\E140000110=Groupe_Fix_Retq_120et80tube\E140000110.jpg") '1
            If ComboBox1.SelectedItem.ToString = "RTQ_ROUE" Then
                Image1.Image = picture1

            Else
                Image1.Image = picture2

            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Select Case ComboBox1.SelectedItem
            Case "RTQ_ROUE"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00100_RETAQUAGE_EQUIPMENT\E140000100=RTQ_ROUE\3D\RTQ_ROUE_01.CATProduct")

                On Error GoTo 0
                ' Add a new Part

                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

            Case "Groupe_Fix_Retaquage"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00100_RETAQUAGE_EQUIPMENT\E140000110=Groupe_Fix_Retq_120et80tube\3D\Groupe_Fix_Retaquage_1.CATProduct")

                On Error GoTo 0
                ' Add a new Part


                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.

        End Select
        End Sub

        Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
            RenaultStandardi.Show()
            Me.Close()
        End Sub

        'Private Sub ''Image1_Click(sender As Object, e As EventArgs)

        'End Sub

    End Class