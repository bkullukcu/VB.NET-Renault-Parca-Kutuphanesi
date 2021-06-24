Public Class RenaultStandardi
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim CATIA As Object
        Dim Mydocument
        Dim PartFactory
        Dim iPartDoc
        ' Get CATIA Or Launch it if necessary.

        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")
        If Err.Number Then
            Label4.Visible = True
        Else
            Label4.Visible = False
        End If
        On Error GoTo 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Parcatipi.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Parcatipi200.Show()
        Me.Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Bilgiler.Show()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Parcatipi000.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Parcatipi001.Show()
        Me.Close()
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Güncelleme.Show()
        Me.Close()
    End Sub
End Class
