Public Class Table_Depose_Equipment
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image

    Private Sub Table_Depose_Equipment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("TDE_Retaquage_Pilot_Surface_003")
        ComboBox1.Items.Add("TDE_Retaquage_Pilot_Surface_XYZ_001")
        ComboBox1.Items.Add("TDE_Retaquage_Pilot_Surface_XZ_002")

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()
    End Sub

    Private Sub StoolHightener_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'For i As Integer = 150 To 600 Step 50
        '    ComboBox1.Items.Add(i.ToString & " [mm]")
        'Next

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00500_TABLE_DEPOSE_EQUIPMENT\01_Retaquage_Pilot_Pieces\3D\TDE_Retaquage_Pilot_Surface_003.jpg") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00500_TABLE_DEPOSE_EQUIPMENT\01_Retaquage_Pilot_Pieces\3D\TDE_Retaquage_Pilot_XYZ_001.jpg") '1
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00500_TABLE_DEPOSE_EQUIPMENT\01_Retaquage_Pilot_Pieces\3D\TDE_Retaquage_Pilot_XZ_002.jpg") '1

            Select Case ComboBox1.SelectedItem
                Case "TDE_Retaquage_Pilot_Surface_003"
                    Image1.Image = picture1

                Case "TDE_Retaquage_Pilot_Surface_XYZ_001"
                    Image1.Image = picture2

                Case "TDE_Retaquage_Pilot_Surface_XZ_002"
                    Image1.Image = picture3

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
            Case "TDE_Retaquage_Pilot_Surface_003"
                'Image1.Image = picture1
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00500_TABLE_DEPOSE_EQUIPMENT\01_Retaquage_Pilot_Pieces\3D\TDE_Retaquage_Pilot_Surface_003.CATPart")

                On Error GoTo 0
            Case "TDE_Retaquage_Pilot_Surface_XYZ_001"
                'Image1.Image = picture2
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00500_TABLE_DEPOSE_EQUIPMENT\01_Retaquage_Pilot_Pieces\3D\TDE_Retaquage_Pilot_XYZ_001.CATPart")

                On Error GoTo 0
            Case "TDE_Retaquage_Pilot_Surface_XZ_002"
                'Image1.Image = picture3
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00500_TABLE_DEPOSE_EQUIPMENT\01_Retaquage_Pilot_Pieces\3D\TDE_Retaquage_Pilot_XZ_002.CATPart")

                On Error GoTo 0
        End Select
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    'Private Sub ''Image1_Click(sender As Object, e As EventArgs)

    'End Sub
End Class