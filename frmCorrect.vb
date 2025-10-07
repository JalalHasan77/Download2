Public Class frmCorrect
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim lcComponent() As String
        Dim clLine As New Collection
        Dim clLink As New Collection
        Dim link As String = ""

        Dim inFile As New IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory & "\done.txt")
        Dim outFile As New IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\done2.txt")

        Dim oneLine As String
        Dim count As Integer = 0
        While Not inFile.Peek
            oneLine = inFile.ReadLine
            lcComponent = Split(oneLine, vbTab)

            If clLine.Contains(oneLine.Trim.ToUpper) = True Then
                TextBox1.Text = TextBox1.Text & vbCrLf & "Whole: " & vbTab & oneLine.Trim.ToUpper
            Else
                clLine.Add(oneLine.Trim.ToUpper, oneLine.Trim.ToUpper)
            End If

            link = lcComponent(0)

            If clLink.Contains(link.Trim.ToUpper) = True Then
                TextBox1.Text = TextBox1.Text & vbCrLf & "Link: " & vbTab & link.Trim.ToUpper
            Else
                outFile.WriteLine(oneLine)
            End If

            Label1.Text = count
            Application.DoEvents()
            count += 1
        End While
        'MsgBox("Done")
        'MsgBox(clLine.Count & vbTab & clLink.Count)

        inFile.Close()
        outFile.Flush()
        outFile.Close()




    End Sub
End Class