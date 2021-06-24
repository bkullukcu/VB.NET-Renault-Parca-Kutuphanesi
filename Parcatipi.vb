Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office


Public Class Parcatipi
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PinBracket.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cale.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Drageoir.Show()
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Serrage.Show()
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        RenaultStandardi.Show()
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Poteau.Show()
        Me.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        StoolHightener.Show()
        Me.Close()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
        Dim xlApp As New Excel.Application
        xlApp.Visible = True
        Try
            xlApp.Workbooks.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01000_PNEUMATIQUE_AUTOMATION_EQUIPMENTS\Valf Adası Seçme Macrosu.xlsx")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
            xlApp.Visible = False
        End Try
        If System.IO.File.Exists("xlApp") Then
            Me.Close()
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        xlApp = Nothing


        GC.Collect()
        GC.WaitForPendingFinalizers()


    End Sub


    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_00500_BASE_marbre\CHECKING_BALL & ACCESSORIES\CHECKING_BALL_CAP_SET_HIGHTENER\3D\CHECKING_BALL_CAP_SET_HIGHTENER_01.CATProduct")

        On Error GoTo 0
        ' Add a new Part

        'Mydocument = CATIA.Documents.Add("Part")
        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
        If System.IO.File.Exists("iPartDoc") Then
            Me.Close()
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E11_00700_ADJUSTMENT PILLAR_pied de reglage\ADJUST_PILLAR\3D\0701_ADJUST_PILLAR.CATProduct")

        On Error GoTo 0
        ' Add a new Part

        'Mydocument = CATIA.Documents.Add("Part")
        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
        If System.IO.File.Exists("iPartDoc") Then
            Me.Close()
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")
        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_01100_CONNECTION_EQUIPMENT\01_V_Bute_groupe\3D\V_Bute_groupe.CATProduct")

        On Error GoTo 0
        ' Add a new Part

        'Mydocument = CATIA.Documents.Add("Part")
        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
        If System.IO.File.Exists("iPartDoc") Then
            Me.Close()
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Retaquage.Show()
        Me.Close()
    End Sub

    Private Sub Parcatipi_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")

        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00100_RETAQUAGE_EQUIPMENT\E140000100=RTQ_ROUE\3D\RTQ_ROUE_01.CATProduct")

        On Error GoTo 0
        ' Add a new Part

        'Mydocument = CATIA.Documents.Add("Part")
        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
        If System.IO.File.Exists("iPartDoc") Then
            Me.Close()
        End If

    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        Pilote.Show()
        Me.Close()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Detector.Show()
        Me.Close()

    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Two_Face_Stool_Poteau.Show()
        Me.Close()
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        AppuiTouche.Show()
        Me.Close()

    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Equerre.Show()
        Me.Close()

    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        SopapEquipment.Show()
        Me.Close()

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Table_Depose_Equipment.Show()
        Me.Close()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")
        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_04200_MINOMI_PES_PS_EQUIPMENT\PES-R_Middle\3D\PES_Middle_Carcasse_01.CATProduct")

        On Error GoTo 0
        ' Add a new Part

        'Mydocument = CATIA.Documents.Add("Part")
        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
        If System.IO.File.Exists("iPartDoc") Then
            Me.Close()
        End If
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Pneumatique.Show()
        Me.Close()
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Zone.Show()
        Me.Hide()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        'Get CATIA Or Launch it if necessary.
        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")
        CATIA = CreateObject("CATIA.Application")
        CATIA.Visible = True
        iPartDoc = CATIA.Documents.Open("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_10000_SAS_AGV\01_SAS_Conception_de_reference\3D\General_SAS_AGV_Chariot.CATProduct")

        On Error GoTo 0
        ' Add a new Part

        'Mydocument = CATIA.Documents.Add("Part")
        'PartFactory = Mydocument.Part.ShapeFactory  ' Retrieve the Part Factory.
        If System.IO.File.Exists("iPartDoc") Then
            Me.Close()
        End If
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs)
        Try
            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_10300_GOUJON SUPPORT")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Try
            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00200_ACCROCHAGE_EQUIPMENT")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Try

            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00300_PLATFORM_EQUIPMENT")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Try

            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_10400_REHAUSSE")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Try

            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00700_GABARIT_EQUIPMENT")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Try

            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00600_ROBOT_INTERFACE")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Try

            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_10100_COLLE_SUPPORT")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Try

            Process.Start("I:\_sharediao\DIV\tooling_ass_oyk\uet_tooling_oyk\ORTAK\UET_MEKANIK\19-)Catalogue\BIBLIOTHEQUE_2020\101_REMONTAGE_std_RENAULT\E14_00400_PALAN_HANDLING_EQUIPMENT")
        Catch ex As Exception
            MsgBox("Dosya Bulunamadı")
        End Try
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

    End Sub

    Private Sub PictureBox30_Click(sender As Object, e As EventArgs) Handles PictureBox30.Click

    End Sub
End Class