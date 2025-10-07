Public Class Form2
    Public PictureBoxSize As Size
    Public PictureBoxLocation As Point
    Public lcShown As Boolean = False
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.Width = 1
        Me.Height = 1
        Me.Size = New Size(Me.Label1.Location.X * 1 + Me.Label1.Width + Me.Label1.Location.X * 1, Me.Label1.Location.Y * 1 + Me.Label1.Height + Me.Label1.Location.Y * 1)





        Dim LocX As Integer
        Dim LocY As Integer
        LocX = PictureBoxLocation.X + CInt(PictureBoxSize.Width / 2) - CInt(Me.Size.Width / 2)
        LocY = PictureBoxLocation.Y + CInt(PictureBoxSize.Height / 2) - CInt(Me.Size.Height / 2)
        Me.Location = New Point(LocX, LocY)






        Application.DoEvents()

    End Sub
End Class