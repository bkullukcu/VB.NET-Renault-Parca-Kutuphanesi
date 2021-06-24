Public Class Two_Face_Stool_Poteau
    Private Sub Two_Face_Stool_Poteau_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i As Integer = 150 To 500 Step 50
            ComboBox1.Items.Add(i.ToString & " [mm]")
        Next
        Try
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\ScreenHunter_031.jpg") '1
            Image1.Image = picture1
        Catch ex As Exception
            'MsgBox("Bir ya da birkaç fotoğraf eksik")
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi.Show()
        Me.Close()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        Select Case ComboBox1.SelectedItem
            Case "150 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1501_2FACES_STOOL_H150\1501_2FACES_STOOL_H150.CATProduct")

                On Error GoTo 0
                ' Add a new Part

                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
            Case "200 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1502_2FACES_STOOL_H200\1502_2FACES_STOOL_H200.CATProduct")

                On Error GoTo 0
                ' Add a new Part
                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
            Case "250 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1503_2FACES_STOOL_H250\1503_2FACES_STOOL_H250.CATProduct")

                On Error GoTo 0
                ' Add a new Part


                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
            Case "300 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1504_2FACES_STOOL_H300\1504_2FACES_STOOL_H300.CATProduct")

                On Error GoTo 0
                ' Add a new Part


                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
            Case "350 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1505_2FACES_STOOL_H350\1505_2FACES_STOOL_H350.CATProduct")

                On Error GoTo 0
                ' Add a new Part


                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
            Case "400 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1506_2FACES_STOOL_H400\1506_2FACES_STOOL_H400.CATProduct")

                On Error GoTo 0
                ' Add a new Part


                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
            Case "450 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1507_2FACES_STOOL_H450\1507_2FACES_STOOL_H450.CATProduct")

                On Error GoTo 0
                ' Add a new Part


                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
            Case "500 [mm]"
                'Get CATIA Or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01500_2FACES_STOOL_poteau_2_faces\3D\1508_2FACES_STOOL_H500\1508_2FACES_STOOL_H500.CATProduct")

                On Error GoTo 0
                ' Add a new Part


                'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                'Case "550 [mm]"
                '    'Get CATIA Or Launch it if necessary.
                '    On Error Resume Next
                '    CATIA = GetObject(, "CATIA.Application")

                '    CATIA = CreateObject("CATIA.Application")
                '    CATIA.Visible = True
                '    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01900_STOOL_HIGHTENER_rehausse_poteau\3D\1908_H550\1908_STOOL_HIGHTENER_H550.CATProduct")

                '    On Error GoTo 0
                '    ' Add a new Part


                '    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
                'Case "600 [mm]"
                '    'Get CATIA Or Launch it if necessary.
                '    On Error Resume Next
                '    CATIA = GetObject(, "CATIA.Application")

                '    CATIA = CreateObject("CATIA.Application")
                '    CATIA.Visible = True
                '    iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_01900_STOOL_HIGHTENER_rehausse_poteau\3D\1909_H600\1909_STOOL_HIGHTENER_H600.CATProduct")

                '    On Error GoTo 0
                '    ' Add a new Part


                '    'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
        End Select
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    'Private Sub ''Image1_Click(sender As Object, e As EventArgs)

    'End Sub
End Class