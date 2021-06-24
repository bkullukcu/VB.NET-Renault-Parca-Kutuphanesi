Public Class Güncelleme
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim value1 As String = "\_My_Site"
        Console.WriteLine(value1)
        Dim value2 As String = value1.Replace("\_My_Site", "")
        Console.WriteLine(value2)
    End Sub
End Class