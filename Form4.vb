Imports System.Text.RegularExpressions
Imports System.Net

Public Class Form4
    Public WithEvents download As WebClient
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim arrText() As String = TextBox1.Text.Split(vbCrLf)
        Dim intArr() As String
        Dim Country As String = ""
        Dim Link As String = ""

        TextBox1.Text = ""
        For Each ABC As String In arrText
            Country = ""
            Link = ""

            intArr = ABC.Split({"Flag of ", " src="""}, StringSplitOptions.None)
            '  TextBox1.Text = TextBox1.Text & vbCrLf & Join(intArr, vbTab)
            Try

                Country = intArr(1)
                Country = Country.Substring(0, -1 + InStr(Country, """"))
                If InStr(Country, ",") > 0 Then
                    Country = Country.Substring(0, -1 + InStr(Country, ","))
                End If

                Link = intArr(3)
                Link = Link.Substring(0, -1 + InStr(Link, """"))

                TextBox1.Text = TextBox1.Text & vbCrLf & Country & vbTab & Link


                DownLoadFile(Link,
                System.AppDomain.CurrentDomain.BaseDirectory & "/flags/flg" & Country.Replace(" ", "_") & "SQR.png")
                Application.DoEvents()

            Catch ex As Exception

            End Try

        Next

    End Sub
    Private Sub DownLoadFile(ByVal URL As String, ByVal FileName As String)
        URL = URL.Replace("&amp;", "&")
        download = New WebClient
        download.DownloadFileAsync(
            New Uri(URL),
            FileName)


        While download.IsBusy
            '    ToolStripStatusLabel1.Text = "Downloading " & lcListItem.Word
            Application.DoEvents()
        End While
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim mainDir As New IO.DirectoryInfo("\\HQPROFILES\HQUP$\2271\Desktop\screens\")

        Dim allFiles() As IO.FileInfo = mainDir.GetFiles("*.pn*")





        For Each onefile As IO.FileInfo In allFiles
            'IO.File.Move(onefile.FullName, onefile.DirectoryName & "\" & IO.Path.GetFileNameWithoutExtension(onefile.Name) & "Logo.png")
            Dim Img As Image = Image.FromFile(onefile.FullName)
            Dim bmp As New Bitmap(Img.Width, Img.Height)

            Dim GR As Graphics = Graphics.FromImage(bmp)
            '  Dim R As NewRectangle(0, 0, Img.Width, Img.Height)

            '=================================
            'Draw on picture and save it 
            GR.DrawImage(Img, 0, 0, Img.Width, Img.Height)



            bmp.Save("\\HQPROFILES\HQUP$\2271\Desktop\screens\" & IO.Path.GetFileNameWithoutExtension(onefile.Name) & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)


        Next

        MsgBox("done")


    End Sub
End Class