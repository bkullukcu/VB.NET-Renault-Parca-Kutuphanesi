Public Class Guijon
    Dim picture1 As Image
    Dim picture2 As Image
    Dim picture3 As Image
    Dim picture4 As Image
    Dim picture5 As Image
    Dim picture6 As Image
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Guijon_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Manuel")
        ComboBox1.Items.Add("Robotize")

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            'Horizontal
            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm205e_m8.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\lm310-bsb612-da.png") '1
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200-bsb514-ok.png") '1
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200-bsb618f83-ok.png") '1
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200-bsb12517bm-c.png") '1
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Robotize\lm310-bsb14522bm-smb-wr+m100579-adapterplatte.png") '17

            ComboBox2.ResetText()


            Me.ComboBox2.Items.Clear()

            If Me.ComboBox1.SelectedItem.ToString = "Manuel" Then
                Me.ComboBox2.Items.Add("plm205e_m6")
                Me.ComboBox2.Items.Add("plm205e_m8")
                Me.ComboBox2.Items.Add("plm200g2_6")
                Me.ComboBox2.Items.Add("plm200gk_fl5")
                Me.ComboBox2.Items.Add("plm200g5_6")
                Me.ComboBox2.Items.Add("plm200gk_fl11")
                Me.ComboBox2.Items.Add("plm560gk_fl11")
                Me.ComboBox2.Items.Add("plm560g2_6")
                Me.ComboBox2.Items.Add("plm560gk_fl13")
                Me.ComboBox2.Items.Add("plm560g5_6")
                Me.ComboBox2.Items.Add("pse100gmasse")
                Me.ComboBox2.Items.Add("pse100g5_6")
                Me.ComboBox2.Items.Add("lm310-bsb612-da")
                Me.ComboBox2.Items.Add("plm200-bsb618f83-ok")
                Me.ComboBox2.Items.Add("plm200-bsb12517bm-c")
                Me.ComboBox2.Items.Add("plm200-bsb514-ok")
                Me.ComboBox2.Visible = True
            ElseIf Me.ComboBox1.SelectedItem.ToString = "Robotize" Then
                Me.ComboBox2.Visible = False


            End If
            ComboBox1.Focus()
            With Me.ComboBox1
                Select Case ComboBox1.SelectedItem
                    Case "Manuel"
                        Image1.Image = picture1
                        ComboBox2.Show()
                        Label1.Show()
                    Case "Robotize"
                        ComboBox2.Hide()
                        Label1.Hide()
                        Image1.Image = picture6
                End Select
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try

            Dim picture1 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm205e_m8.png") '1
            Dim picture2 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\lm310-bsb612-da.png") '1
            Dim picture3 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200-bsb514-ok.png") '1
            Dim picture4 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200-bsb618f83-ok.png") '1
            Dim picture5 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200-bsb12517bm-c.png") '1
            Dim picture6 As Image = Image.FromFile("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Robotize\lm310-bsb14522bm-smb-wr+m100579-adapterplatte.png") '17

            ComboBox2.Focus()
            With Me.ComboBox2
                Select Case ComboBox1.SelectedItem
                    Case "Manuel"
                        If ComboBox2.SelectedItem = "plm205e_m6" Then
                            Image1.Image = picture1
                        ElseIf ComboBox2.SelectedItem = "lm310-bsb612-da" Then
                            Image1.Image = picture2
                        ElseIf ComboBox2.SelectedItem = "plm200-bsb618f83-ok" Then
                            Image1.Image = picture3
                        ElseIf ComboBox2.SelectedItem = "plm200-bsb12517bm-c" Then
                            Image1.Image = picture4
                        ElseIf ComboBox2.SelectedItem = "plm200-bsb514-ok" Then
                            Image1.Image = picture5
                        Else
                            Image1.Image = picture1
                        End If
                    Case "Robotize"
                        Image1.Image = picture6
                End Select
            End With
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi000.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc

        If ComboBox1.SelectedItem = "Manuel" Then
            If ComboBox2.SelectedItem = "plm205e_m6" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm205e_m6.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm205e_m8" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm205e_m8.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm200g2_6" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200g2_6.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm200gk_fl5" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200gk_fl5.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm200g5_6" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200g5_6.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm200gk_fl11" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm200gk_fl11.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm560gk_fl11" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm560gk_fl11.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm560g2_6" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm560g2_6.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm560gk_fl13" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm560gk_fl13.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "plm560g5_6" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\plm560g5_6.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "pse100gmasse" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\pse100gmasse.igs")

                On Error GoTo 0
            End If
            If ComboBox2.SelectedItem = "pse100g5_6" Then
                'Get CATIA or Launch it if necessary.
                On Error Resume Next
                CATIA = GetObject(, "CATIA.Application")

                CATIA = CreateObject("CATIA.Application")
                CATIA.Visible = True
                iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Manuel\pse100g5_6.igs")

                On Error GoTo 0
            End If
        End If

        If ComboBox1.SelectedItem = "Robotize" Then
            'Get CATIA Or Launch it if necessary.
            On Error Resume Next
            CATIA = GetObject(, "CATIA.Application")

            CATIA = CreateObject("CATIA.Application")
            CATIA.Visible = True
            iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\000_STUDY SUPPORT EQUIPMENTS\GOUJON_GUN\Goujon Robotize\lm310-bsb14522bm-smb-wr+m100579-adapterplatte.CATPart")

            On Error GoTo 0
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub
End Class