Public Class Files
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        IO.File.WriteAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & lblFile.Text, TextBox1.Text)
        DialogResult = DialogResult.OK
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DialogResult = DialogResult.Cancel
    End Sub
End Class