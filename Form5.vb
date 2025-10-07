Public Class Form5
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        IO.File.WriteAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\DevLatests.txt", TextBox1.Text)
        DialogResult = DialogResult.OK
        Me.Close()
    End Sub
End Class