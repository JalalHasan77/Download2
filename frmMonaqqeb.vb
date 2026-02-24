Imports System.Net
Imports HtmlAgilityPack
Imports System.Text
Imports System.Net.Http
Imports System.Threading.Tasks

Imports System.IO
Imports System.Text.RegularExpressions

Public Class frmMonaqqeb
    Public WithEvents download As WebClient
    Dim encoding As Encoding
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        IO.File.WriteAllText("result.html", "", Encoding.UTF8)
        DowloadMonuqqeb(TextBox1.Text)
        Process.Start("result.html")
    End Sub

    Sub DowloadMonuqqeb(ByVal URL As String)

        While URL <> ""
            DownLoadFile(URL, "Monaqqeb.html")
            'Process.Start("result.html")
            '
            list2()

            URL = getNextResult()
        End While


    End Sub

    Private Function getNextResult() As String
        Dim fileBytes As Byte() = IO.File.ReadAllBytes("Monaqqeb.html")
        Dim html As String = Encoding.GetEncoding(1256).GetString(fileBytes)

        ' Load HTML
        Dim doc As New HtmlDocument()
        doc.LoadHtml(html)

        Dim hrefValue As String = Nothing

        ' Look for the specific <img> that indicates "no more results"
        Dim noMoreResultsImg = doc.DocumentNode.SelectSingleNode("//img[@src='../images/newimages/nomoreresults.gif' and @alt='نهائية النتائج']")

        If noMoreResultsImg IsNot Nothing Then
            ' Tag found → return empty string
            Return ""
        Else
            ' Otherwise, proceed to get enclosing <a> href for other images (example)
            Dim imgNode = doc.DocumentNode.SelectSingleNode("//img[@alt='النتائج التالية']") ' your other img condition
            If imgNode IsNot Nothing Then
                Dim parentLink = imgNode.ParentNode
                If parentLink IsNot Nothing AndAlso parentLink.Name.ToLower() = "a" Then
                    hrefValue = parentLink.GetAttributeValue("href", "")
                Else
                    hrefValue = ""
                End If
            Else
                hrefValue = ""
            End If
        End If
        Return "http://holyquran.net/cgi-bin/" & hrefValue
    End Function

    Sub list2()

        Dim fileBytes As Byte() = IO.File.ReadAllBytes("Monaqqeb.html")

        ' Dim html As String = IO.File.ReadAllText("Monaqqeb.html", encoding)

        Dim html As String = Encoding.GetEncoding(1256).GetString(fileBytes)
        RemoveWhatIsBefore(html, "</form>", False)
        RemoveWhatIsBefore(html, "<tr>", True)

        Dim doc As New HtmlDocument()
        doc.LoadHtml(html)

        Dim resultDict As New Dictionary(Of Integer, Tuple(Of String, String))

        Dim rows = doc.DocumentNode.SelectNodes("//tr")

        Dim thirdCellValue As String = Nothing
        Dim singleCellValue As String = Nothing
        Dim counter As Integer = 1

        If rows IsNot Nothing Then
            For Each row In rows

                Dim cells = row.SelectNodes("td")

                If cells IsNot Nothing Then

                    ' Row with 4 cells → get third cell (index 2)
                    If cells.Count = 4 Then
                        thirdCellValue = cells(1).InnerText.Trim()
                    End If

                    ' Row with 1 cell → get that cell
                    If cells.Count = 1 Then
                        'singleCellValue = cells(0).InnerText.Trim()
                        Dim text As String = cells(0).InnerText.Trim()

                        ' Check if it contains at least one Arabic character
                        If Regex.IsMatch(text, "[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]") Then
                            singleCellValue = text
                        End If
                    End If

                    ' If both values are available, store tuple
                    If thirdCellValue IsNot Nothing AndAlso singleCellValue IsNot Nothing Then
                        resultDict.Add(counter, Tuple.Create(thirdCellValue, singleCellValue))
                        counter += 1

                        ' Reset for next pair
                        thirdCellValue = Nothing
                        singleCellValue = Nothing
                    End If

                End If

            Next
        End If

        ' Example: Print dictionary

        For Each kvp In resultDict
            IO.File.AppendAllText("result.html", kvp.Key & "<br><br>" & kvp.Value.Item1 & "<br><br>" & kvp.Value.Item2 & "<br><br><br><br>", Encoding.UTF8)

        Next


    End Sub

















    Private Sub ListResults()

        Dim html As String = IO.File.ReadAllText("Monaqqeb.html")

        RemoveWhatIsBefore(html, "</form>", False)
        RemoveWhatIsBefore(html, "<tr>", True)
        'RemoveWhatIsBefore(allContentArr(I), ">", False)
        'OpenningTag = allContentArr(I).Substring(0, InStr(allContentArr(I), " "))
        'ClosingTag = OpenningTag.Replace("<", "</")
        'RemoveWhatIsAfter(allContentArr(I), ClosingTag.Trim, False)

        Dim ALLTR() As String
        ALLTR = Split(html, "<tr>")

        For I As Integer = 0 To ALLTR.Count - 1 Step 2
            MsgBox(ALLTR(I) & vbCrLf & vbCrLf & ALLTR(I + 1))



        Next
        '
    End Sub

    Private Sub DownLoadFile(ByVal URL As String, ByVal FileName As String)
        'Exit Sub

        URL = URL.Replace("&amp;", "&")
        download = New WebClient
        Try
            download.DownloadFileAsync(
            New Uri(URL),
            FileName)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        While download.IsBusy

            Application.DoEvents()
        End While
    End Sub

    Sub RemoveWhatIsBefore(ByRef Text As String,
                       ByVal MarkingText As String,
                       ByVal KeepMarkingText As Boolean)


        If InStr(Text, MarkingText) <= 0 Then Exit Sub

        Dim Position As Integer

        Position = Text.IndexOf(MarkingText)

        If KeepMarkingText = True Then

        Else
            Position = Position + MarkingText.Length
        End If

        Try
            Text = Text.Substring(Position)
        Catch ex As Exception

        End Try


    End Sub

    Sub RemoveWhatIsAfter(ByRef Text As String,
                           ByVal MarkingText As String,
                           ByVal KeepMarkingText As Boolean)

        If InStr(Text, MarkingText) <= 0 Then Exit Sub

        Dim Position As Integer

        Position = Text.IndexOf(MarkingText)

        If KeepMarkingText = False Then

        Else
            Position = Position + MarkingText.Length
        End If

        Try
            Text = Text.Substring(0, Position)
        Catch ex As Exception

        End Try


    End Sub

    Sub RemoveEmptyORNull(ByRef SourceArray() As String,
                          ByVal RemoveWhiteSpaces As Boolean,
                          ByVal RemoveEmptyCells As Boolean)

        Dim L As New List(Of String)
        L.AddRange(SourceArray)

        If RemoveWhiteSpaces = True Then
            L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        End If

        If RemoveEmptyCells = True Then
            L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
        End If

        SourceArray = L.ToArray
    End Sub

    Sub RemoveFirstCell(ByRef lcArray() As String)
        lcArray(0) = String.Empty

        RemoveEmptyORNull(SourceArray:=lcArray,
        RemoveWhiteSpaces:=True,
        RemoveEmptyCells:=True)

    End Sub

    Private Sub FrmMonaqqeb_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "http://holyquran.net/cgi-bin/qsearch.pl?st=%C2%ED%C7%CA&sc=1&sv=1&ec=114&ev=7&ae=&mw=p&alef=ON"
    End Sub
End Class