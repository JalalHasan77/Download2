Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO



Public Class Form1
    Public WithEvents download As WebClient
    Dim Dic As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Dim NewDic As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
    Dim PhotoDic As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    Dim enumerator As IEnumerator
    Dim L As New List(Of links)
    Dim index As Integer = 0
    Dim previous As Integer = -1

    Dim pURL As New Process
    Dim pFile As New Process

    Dim DevLastTime As Date
    Private glTitle As String



    Dim filePath As String = "D:\VB Projects\Downloads\Downloads\bin\Debug\Guaridian.txt"
    Dim prevDatTime As New DateTime(1990, 1, 1, 0, 0, 0)
    Dim GuardianIndex As Integer = 53

    Private OriginalNodes As New List(Of TreeNode)() ' stores original root-level copies

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' LoadLinksInotCollection()
        LoadTreeFromFile("C:\Users\2271\Downloads\EssaysClassifications_Updated.txt", TreeView2)
        TreeView2.ExpandAll()

        ' Backup the original tree
        For Each n As TreeNode In TreeView2.Nodes
            OriginalNodes.Add(CType(n.Clone(), TreeNode))
        Next


        prevDatTime = File.GetLastWriteTime(filePath)
        pURL = Nothing
        pFile = Nothing

    End Sub

    ''' <summary>
    ''' Load a TreeView from a file of "index - item" lines (e.g., "1.2.3 - Title").
    ''' Accepts separators: -, –, :, or TAB. Ignores empty / comment lines.
    ''' </summary>
    Private Sub LoadTreeFromFile(filePath As String, tv As TreeView)
        If Not File.Exists(filePath) Then
            MessageBox.Show("File not found: " & filePath)
            Return
        End If

        tv.BeginUpdate()
        tv.Nodes.Clear()

        ' Cache nodes by full index ("1.2.3") so we can attach children quickly
        Dim nodeByIndex As New Dictionary(Of String, TreeNode)(StringComparer.OrdinalIgnoreCase)

        ' Regex: capture hierarchical index and the item text.
        '   Group 1: digits with dot segments (e.g., 1.2.3)
        '   Separator: -, en dash, :, or tab (with optional spaces)
        '   Group 2: the item text
        Dim rx As New Regex("^\s*(\d+(?:\.\d+)*)\s*(?:-|–|:|\t)\s*(.+?)\s*$", RegexOptions.Compiled)

        For Each rawLine In File.ReadLines(filePath)
            Dim line As String = rawLine.Trim()
            If line.Length = 0 OrElse line.StartsWith("#") Then Continue For

            Dim m = rx.Match(line)
            If Not m.Success Then
                ' Fallback: allow "index item" (space) if no separator was used
                Dim firstSpace = line.IndexOf(" "c)
                If firstSpace > 0 AndAlso Regex.IsMatch(line.Substring(0, firstSpace), "^\d+(?:\.\d+)*$") Then
                    m = Regex.Match(line, "^\s*(\d+(?:\.\d+)*)\s+(.+?)\s*$")
                    If Not m.Success Then Continue For
                Else
                    Continue For
                End If
            End If

            Dim fullIndex = m.Groups(1).Value
            Dim itemText = m.Groups(2).Value

            ' Ensure all ancestor nodes exist: "1" -> "1.2" -> "1.2.3"
            Dim parts = fullIndex.Split("."c)
            Dim path = ""
            Dim parentNode As TreeNode = Nothing

            For i = 0 To parts.Length - 1
                path = If(i = 0, parts(0), path & "." & parts(i))

                Dim currentNode As TreeNode = Nothing
                If nodeByIndex.TryGetValue(path, currentNode) Then
                    parentNode = currentNode
                    Continue For
                End If

                ' Create a new node for this index level

                Dim newNode As New TreeNode(parts(i)) With {.Tag = path}


                If parentNode Is Nothing Then
                    tv.Nodes.Add(newNode)

                Else
                    parentNode.Nodes.Add(newNode)
                End If

                nodeByIndex(path) = newNode
                parentNode = newNode
            Next

            ' Set the final node's display text to "index - item"
            Dim leaf As TreeNode = nodeByIndex(fullIndex)
            leaf.Text = fullIndex & " - " & itemText
            leaf.ToolTipText = itemText ' optional
            leaf.Tag = fullIndex
        Next

        tv.EndUpdate()
        'tv.ExpandAll() ' optional
    End Sub


    Private Sub LoadLinksInotCollection()
        Dim TExt As String = ""
        Dim allContent As String
        Dim allContentArr() As String
        Try
            allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\done.txt")
        Catch ex As Exception
            Exit Sub
        End Try

        allContentArr = Split(allContent, vbCrLf)
        Dim subArr() As String
        'Dim text As String = ""
        'Clipboard.Clear()

        For Each Str As String In allContentArr
            'If Not Str.Contains("quantamagazine") Then Continue For
            If Str <> "" Then
                subArr = Split(Str, vbTab, 2)
                Try
                    Dic.Add(subArr(0), subArr(1))
                Catch ex As Exception
                    TExt = TExt & vbCrLf & subArr(0)
                    'Clipboard.SetText(TExt & vbTab & ex.Message)
                End Try
            End If
        Next
        MsgBox("Done")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'LoadLinksInotCollection()

        CheckedListBox1.Items.Clear()

        DownLoadGuardianFile()
        DownloadDogoNewsScience()
        DownLoadScienceNewsFile()
        DowloadBBCFuture()
        DowloadBBCWorklife()
        DownLoadBBCTravel()
        DownlaodBBCScience()
        DownLoadBBCInDepth()
        DownLoadCNNScience()
        DownLoadLiveScience()
        DownLoadScienceAlert()
        DownloadGlobalResearch()
        DownlaodNature()
        DownlaodBigThink()
        DownlaodQuantaMagazine()
        DownloadAmricanScience()
        DownloadScienceFocus()
        DownloadTheConversation()
        DownloadNationalInterset()
        downloadWM()
        DownloadPopSci()
        DownloadLosAngelesTimes()
        DownloadNewYorkTimes()
        DownloadAP()
        DownLoadNewScientes()
        DownloadEconomist()
        Download_FT()
        DownLoadIndepenednet()
        downloadFreecodecamp()
        downloadRoundTheCode()
        DownloadFreeThing()
        DownloadCostOfEveryThing()



        L.Add(New links(URL:="https://www.dev.to/latest", FileName:="Div.txt"))
        L.Add(New links(URL:="https://www.theguardian.com/science", FileName:="GuardianScience.txt"))
        L.Add(New links(URL:="https://www.reuters.com/science/", FileName:="reuters.txt"))
        L.Add(New links(URL:="https://www.forbes.com/science/", FileName:="forbes.txt"))
        L.Add(New links(URL:="https://katehon.com/en/analitics/trendy?page=1", FileName:="Katehonetrendy.txt"))
        L.Add(New links(URL:="https://katehon.com/en/articles?page=0", FileName:="KatehoneArticles.txt"))
        L.Add(New links(URL:="https://katehon.com/en/analytics?page=0", FileName:="KatehoneAnalytics.txt"))
        L.Add(New links(URL:="https://www.theguardian.com/uk/commentisfree?page=1", FileName:="Guaridian.txt"))
        L.Add(New links(URL:="https://www.theguardian.com/world/iran", FileName:="GuaridianIran.txt"))
        Timer1.Start()

        Label1.Text = "Finish"
        Application.DoEvents()
    End Sub

    Private Sub DownloadFreeThing()
        Label1.Text = "Downloading Free Think.."
        Application.DoEvents()
        DownLoadFile("https: //www.freethink.com/articles/",
            System.AppDomain.CurrentDomain.BaseDirectory & "/FreeThink.txt")

        ListFreeThink()
    End Sub

    Private Sub ListFreeThink()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/FreeThink.txt")


        allContentArr = Split(allContent, "<div class=""mb-10"">")
        RemoveFirstCell(allContentArr)

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty
        For I As Integer = 0 To allContentArr.Length - 1
            RemoveWhatIsBefore(allContentArr(I), "class=""inline eyebrow__field-link"">", False)
            RemoveWhatIsBefore(allContentArr(I), "<a href=""", False)

            Link = allContentArr(I).Substring(0, -1 + InStr(allContentArr(I), """"))

            RemoveWhatIsBefore(allContentArr(I), """>", False)

            RemoveWhatIsAfter(allContentArr(I), "<", False)
            Title = allContentArr(I)
            refineTitle(Text:=Title)
            If AddLinktoCheckedListBox1(Title:="Free Think: " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Free Think: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub DownLoadGuardianFile()
        Label1.Text = "Downloading Guardian.."
        Application.DoEvents()
        DownLoadFile("https://www.theguardian.com/news/series/ten-best-photographs-of-the-day",
            System.AppDomain.CurrentDomain.BaseDirectory & "/file.txt")

        ListGuardianPhotos()
    End Sub

    Sub DowloadBBCFuture()
        Label1.Text = "Downloading BBC Future.."
        Application.DoEvents()

        DownLoadFile("https://bbc.co.uk/future",
    System.AppDomain.CurrentDomain.BaseDirectory & "/BBCfuture.txt")

        ListBBCFuture()
    End Sub

    Sub DowloadBBCWorklife()
        Label1.Text = "Downloading BBC Worklife.."
        Application.DoEvents()

        DownLoadFile("https://bbc.co.uk/worklife",
    System.AppDomain.CurrentDomain.BaseDirectory & "/BBCworklife.txt")

        ListBBCWorkLife()
    End Sub

    Sub DownLoadScienceNewsFile()
        Label1.Text = "Downloading Science News.."
        Application.DoEvents()

        'https://www.sciencenews.org/all-stories
        DownLoadFile("https://www.sciencenews.org/all-stories",
            System.AppDomain.CurrentDomain.BaseDirectory & "/science.txt")



        ListScienceNews()
    End Sub

    Sub DownLoadBBCTravel()
        Label1.Text = "Downloading BBC Travel.."
        Application.DoEvents()

        DownLoadFile("https://www.bbc.co.uk/travel/",
    System.AppDomain.CurrentDomain.BaseDirectory & "/BBCtravel.txt")


        ListBBCTravel()
    End Sub


    Sub DownLoadLiveScience()
        Label1.Text = "Downloading LiveScience.."
        Application.DoEvents()

        DownLoadFile("https://www.livescience.com/news",
        System.AppDomain.CurrentDomain.BaseDirectory & "/LiveScience.txt")

        ListLiveScience()
    End Sub

    Sub DownLoadScienceAlert()
        Label1.Text = "Downloading Science Alert.."
        Application.DoEvents()

        DownLoadFile("https://www.sciencealert.com/latest",
                     System.AppDomain.CurrentDomain.BaseDirectory & "/ScienceAlert.txt")
        'DownLoadFile("https://www.sciencealert.com/",


        ListScienceAlert()
    End Sub

    Sub DownlaodBigThink()

        Label1.Text = "Downloading Big Think.."
        Application.DoEvents()

        DownLoadFile("https://bigthink.com/articles/",
            System.AppDomain.CurrentDomain.BaseDirectory & "/BigThink.txt")

        ListBigThink()

    End Sub

    Sub DownlaodQuantaMagazine()
        Label1.Text = "Downloading Quanta Magazine.."
        Application.DoEvents()

        DownLoadFile("https://www.quantamagazine.org/archive/page/1/",
        System.AppDomain.CurrentDomain.BaseDirectory & "/quantamagazine.txt")

        ListQuantaMagazine()

    End Sub

    Sub DownlaodGuardianScience()

        Label1.Text = "Downloading Guardian Science.."
        Application.DoEvents()

        DownLoadFile("https://www.theguardian.com/science",
        System.AppDomain.CurrentDomain.BaseDirectory & "/GuardianScience.txt")

        ListGuardianScience()

    End Sub

    Sub DownlaodForbes()

        Label1.Text = "Downloading Forbes... "
        Application.DoEvents()
        DownLoadFile("https://www.forbes.com/science/",
        System.AppDomain.CurrentDomain.BaseDirectory & "/Forbes.txt")
        Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "/Forbes.txt")
        ListForbes()
    End Sub
    Sub DownloadScienceFocus()

        Label1.Text = "Downloading Science Focus.."
        Application.DoEvents()

        DownLoadFile("https://www.sciencefocus.com/",
System.AppDomain.CurrentDomain.BaseDirectory & "/ScienceFocus.txt")

        ListScienceFocus()

    End Sub

    Sub DownloadTheConversation()
        Dim FileName As String
        FileName = System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation.txt"
        IO.File.WriteAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation.txt", "")

        Label1.Text = "Downloading Conversation/Global "
        Application.DoEvents()
        DownLoadFile("https://theconversation.com/global", FileName)


        Label1.Text = "Downloading Conversation/Au "
        Application.DoEvents()

        IO.File.AppendAllText(FileName,
        DownLoadFileAndBringContent("https://theconversation.com/au",
System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation_au.txt"))

        Label1.Text = "Downloading Conversation/Africa "
        Application.DoEvents()
        IO.File.AppendAllText(FileName, DownLoadFileAndBringContent("https://theconversation.com/africa",
System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation_africa.txt"))

        Label1.Text = "Downloading Conversation/Canada "
        Application.DoEvents()
        IO.File.AppendAllText(FileName, DownLoadFileAndBringContent("https://theconversation.com/ca",
System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation_ca.txt"))

        Label1.Text = "Downloading Conversation/New Zelanda "
        Application.DoEvents()
        IO.File.AppendAllText(FileName, DownLoadFileAndBringContent("https://theconversation.com/nz",
System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation_nz.txt"))

        Label1.Text = "Downloading Conversation/UK "
        Application.DoEvents()
        IO.File.AppendAllText(FileName, DownLoadFileAndBringContent("https://theconversation.com/uk",
System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation_uk.txt"))

        Label1.Text = "Downloading Conversation/USA "
        Application.DoEvents()
        IO.File.AppendAllText(FileName, DownLoadFileAndBringContent("https://theconversation.com/us",
System.AppDomain.CurrentDomain.BaseDirectory & "/TheConversation_us.txt"))


        ListTheConversation()
    End Sub

    Function DownLoadFileAndBringContent(ByVal URL As String, ByVal FileName As String) As String
        DownLoadFile(URL, FileName)


        Dim AllContent As String = IO.File.ReadAllText(FileName)

        DownLoadFileAndBringContent = AllContent


    End Function

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
    Sub ListScienceNews()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\science.txt")

        RemoveWhatIsBefore(allContent, "<ol class=""river-with-sidebar__list___1EfmS"">", False)
        RemoveWhatIsAfter(allContent, "class=""pagination__wrapper___2qtdg""", False)

        Dim allContentArr() As String
        allContentArr = Split(allContent, "<h3")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            If InStr(Str, "Readers") Then
                Dim A As String = ""
            Else
                Dim A As String = ""
            End If
            If Str Like "*<a href=""//www.sciencenews.org/article/*" Or Str Like "*<a href=""https://www.sciencenews.org/article/*" Then

                RemoveWhatIsAfter(Str, "</a>", True)
                RemoveWhatIsBefore(Str, "<a href=""", True)
            Else
                Continue For
            End If

            Dim Link As String = Str
            Dim Title As String = Str



            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            RemoveWhatIsAfter(Title, "</a>", False)
            RemoveWhatIsBefore(Title, ">", False)

            RemoveWhatIsBefore(Link, "<a href=""", False)


            If InStr(Link.ToLower, "https://".ToLower) = False Then Link = "https://" + Link

            refineTitle(Text:=Title)

            If AddLinktoCheckedListBox1(Title:="Science News: " & Title, Link:=Link) = True Then
                Files = Files + 1
            End If

        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Science News: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub
    Sub refineTitle(ByRef Text As String)
        Text = Trim(StrReverse(Text))
        Text = Text.Replace(vbCr, "")
        Text = Text.Replace(vbLf, "")
        Text = Text.Replace(vbCrLf, "")
        Text = Text.Replace(vbTab, "")
        Text = Trim(StrReverse(Text))

        Text = Text.Replace("&#x27;", "'")
        Text = Text.Replace("&#8217;", "'")
        Text = Text.Replace("&#8216;", "'")
        Text = Text.Replace("&#8221;", """")
        Text = Text.Replace("&#8220;", """")
        Text = Text.Replace("&nbsp;", " ")
        Text = Text.Replace("&amp;", "&")
        Text = Text.Replace("&lsquo;", """")
        Text = Text.Replace("&rsquo;", """")


        Dim arrText As String()
        arrText = Split(Text, "<")

        For I As Integer = LBound(arrText) To UBound(arrText)
            RemoveWhatIsBefore(arrText(I), ">", False)
        Next

        Text = Join(arrText)
        Text = Text.Trim


    End Sub

    Function AddLinktoCheckedListBox1(ByVal Title As String,
                                 ByVal Link As String) As Boolean
        Link = Link.Replace("https://", "#YYYY#")
        Link = Link.Replace("\\", "")
        Link = Link.Replace("""", "")
        Link = Link.Replace("//", "")
        Link = Link.Replace("#YYYY#", "https://")

        AddLinktoCheckedListBox1 = True
        If Dic.ContainsKey(Link) Then
            AddLinktoCheckedListBox1 = False
        Else
            Dim tempTitle As String = Title
            refineTitle(tempTitle)
            Dic.Add(Link, Now.ToString("yyyyMMdd") & vbTab & Title)
            CheckedListBox1.Items.Add(New Item(Title:=tempTitle, Link:=Link))
        End If

        Application.DoEvents()
    End Function

    Sub ListBBCFuture()
        ListBBC(Section:="future")
    End Sub

    Sub ListBBCTravel()
        ListBBC(Section:="travel")
    End Sub

    Sub ListBBCWorkLife()
        ListBBC(Section:="worklife")
    End Sub

    Sub ListBigThink()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\BigThink.txt")
        allContentArr = Split(allContent, "<div class=""card-headline"">")
        RemoveFirstCell(allContentArr)

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty
        For I As Integer = 0 To allContentArr.Length - 1
            RemoveWhatIsBefore(allContentArr(I), "<a href=""", False)
            Link = allContentArr(I).Substring(0, -1 + InStr(allContentArr(I), """>"))
            RemoveWhatIsBefore(allContentArr(I), """>", False)
            RemoveWhatIsAfter(allContentArr(I), "<", False)
            Title = allContentArr(I)
            refineTitle(Text:=Title)
            If AddLinktoCheckedListBox1(Title:="Big Think: " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Big Think: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub ListQuantaMagazine()
        Dim Link As String = String.Empty
        Dim Title As String = String.Empty



        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\quantamagazine.txt")
        allContentArr = Regex.Split(allContent, "<div class=\'card__content|<section class=\""outer\"">")

        RemoveFirstCell(allContentArr)

        Dim lcTEXT As String = ""

        For J As Integer = LBound(allContentArr) To UBound(allContentArr)
            RemoveWhatIsBefore(allContentArr(J), "<a href=", False)
            allContentArr(J) = allContentArr(J).Substring(1)

            Link = allContentArr(J)
            RemoveWhatIsAfter(Link, " ", False)
            'lcTEXT = lcTEXT & vbCrLf & allContentArr(J).Replace(" ", "[]").Replace(vbCr, "[]").Replace(vbLf, "[]").Replace(vbTab, "[]")

            Title = allContentArr(J)
            RemoveWhatIsBefore(Title, Link, False)
            RemoveWhatIsAfter(Title, "</a>", False)
            RemoveWhatIsAfter(Title, "</h3>", False)
            RemoveWhatIsBeforeLast(Title, ">", False)
            refineTitle(Title)
            If AddLinktoCheckedListBox1(Title:="Quanta Magazine" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next


        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Quanta Magazine" & ": No Downloads ", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If


    End Sub

    Private Sub RemoveWhatIsBeforeLast(ByRef Text As String,
                           ByVal MarkingText As String,
                           ByVal KeepMarkingText As Boolean)

        Text = Strings.StrReverse(Text)
        RemoveWhatIsAfter(Text, Strings.StrReverse(MarkingText), KeepMarkingText)
        Text = Strings.StrReverse(Text)
    End Sub

    Private Sub RemoveWhatIsAfterLast(ByRef Text As String,
                           ByVal MarkingText As String,
                           ByVal KeepMarkingText As Boolean)

        Text = Strings.StrReverse(Text)
        RemoveWhatIsBefore(Text, Strings.StrReverse(MarkingText), KeepMarkingText)
        Text = Strings.StrReverse(Text)
    End Sub


    Sub ListGuardianScience()
        Dim Files As Integer = 0

        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\GuardianScience.txt")
        'get the class name 'it is everchanging by guardian
        Dim ClassName As String = ""
        allContentArr = Split(allContent, """><a href=""") '"<div class=""dcr-f9aim1")
        ClassName = allContentArr(0)
        ClassName = StrReverse(ClassName)
        RemoveWhatIsAfter(ClassName, """", False)
        ClassName = StrReverse(ClassName)
        'end of: get the class name 'it is everchanging by guardian
        ClassName = "<div class=""" & ClassName

        allContentArr = Split(allContent, ClassName) '"<div class=""dcr-f9aim1")

        RemoveFirstCell(allContentArr)

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty
        Dim show As Boolean = False
        For i As Integer = 0 To allContentArr.Length - 1
            Link = allContentArr(i)

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            Title = allContentArr(i) 'StrReverse(allContentArr(i))
            RemoveWhatIsAfter(Title, "</h3>", False)
            Title = Title.Replace("</span>", "")

            RemoveWhatIsBefore(Title, "<h3 class=""card-headline", False)
            RemoveWhatIsBeforeLast(Title, ">", False)


            If AddLinktoCheckedListBox1(Title:="Guardian Science" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Guardian Science: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub

    Sub ListForbes()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\Forbes.txt")

        '
        While allContent.Contains("class=""LzN-")
            allContent = allContent.Replace("class=""LzN-", "####$$$$$#####@@@")
        End While

        While allContent.Contains("class=""_1-FLFW4R")
            allContent = allContent.Replace("class=""_1-FLFW4R", "####$$$$$#####!!!")
        End While
        'allContent = allContent.Replace("class=""LzN-|class=""_1-FLFW4R", "####$$$$$#####!!!")

        'allContentArr = Regex.Split(allContent, "class=""LzN-|class=""_1-FLFW4R")

        allContentArr = Split(allContent, "####$$$$$#####")

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty

        Dim arrTitle As String()

        RemoveFirstCell(allContentArr)
        Dim text As String = ""
        For I As Integer = LBound(allContentArr) To UBound(allContentArr)
            Link = allContentArr(I)
            Title = allContentArr(I)

            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            RemoveWhatIsBefore(Title, "href=""", False)
            RemoveWhatIsBefore(Title, ">", False)

            RemoveWhatIsAfter(Title, "</h2></a>", False)
            RemoveWhatIsAfter(Title, "</a></h3>", False)

            arrTitle = Split(Title.Trim, "<")

            For J As Integer = LBound(arrTitle) To UBound(arrTitle)
                If arrTitle(J).Contains(">") Then
                    RemoveWhatIsBefore(arrTitle(J), ">", False)
                End If
                arrTitle(J) = Trim(arrTitle(J))
            Next

            Title = Join(arrTitle, " ")
            If AddLinktoCheckedListBox1(Title:="Forbes" & ": " & Title, Link:=Link) Then Files = Files + 1

        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Forbes: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub ListScienceFocus()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\ScienceFocus.txt")
        allContentArr = Split(allContent, "<a class=""issue-data""")

        RemoveFirstCell(allContentArr)

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty
        Dim Category As String = String.Empty
        For i As Integer = 0 To allContentArr.Length - 1
            Title = allContentArr(i)
            RemoveWhatIsAfter(Title, "</span>", False)
            refineTitle(Title)

            Link = allContentArr(i)
            RemoveWhatIsBefore(Link, "standard-card-new__article-title qa-card-link", False)
            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.sciencefocus.com/" & Link

            refineTitle(Title)
            If AddLinktoCheckedListBox1(Title:="Science Focus: " & Title, Link:=Link) Then Files = Files + 1

        Next
        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Science Focus: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub ListTheConversation()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\TheConversation.txt")
        allContentArr = Split(allContent, "<div class=""group relative")
        RemoveFirstCell(allContentArr)

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty

        'Dim SubArr() As String
        For i As Integer = 0 To allContentArr.Length - 1
            Link = allContentArr(i)
            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            '
            Title = allContentArr(i)
            Dim Sep As String = ""
            If Title.Contains("xl:text-3xl""><span>") Then
                Sep = "xl:text-3xl""><span>"
            ElseIf Title.Contains("lg:leading-6""><span>") Then
                Sep = "lg:leading-6""><span>"
            ElseIf Title.Contains("lg:text-lg/6""><span>") Then
                Sep = "lg:text-lg/6""><span>"
            End If
            RemoveWhatIsBefore(Title, Sep, False)
            RemoveWhatIsAfter(Title, "<", False)
            refineTitle(Title)
            If AddLinktoCheckedListBox1(Title:="The Conversation" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next
        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="The Conversation: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub
    Sub ListScienceAlert()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\ScienceAlert.txt")

        allContentArr = Regex.Split(allContent, "<div class=""entry-teaser"">|<div class=""trending-news__inner")
        RemoveFirstCell(allContentArr)

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty
        For i As Integer = 0 To allContentArr.Length - 1
            RemoveWhatIsBefore(allContentArr(i), "<a class=""text-", True)
            RemoveWhatIsAfter(allContentArr(i), "</a>", True)

            Link = allContentArr(i)
            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            Title = allContentArr(i)
            RemoveWhatIsBefore(Title, Link, False)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "</a>", False)
            If AddLinktoCheckedListBox1(Title:="Science Alert" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Science Alert: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub

    Sub ListLiveScience()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\Livescience.txt")

        allContentArr = Split(allContent, "<a href=""https://www.livescience")

        For I As Integer = 0 To allContentArr.Length - 1
            If InStr(allContentArr(I), "class=""article-link""") Then
            Else
                allContentArr(I) = ""
            End If
        Next
        RemoveEmptyORNull(allContentArr, True, True)

        Dim Link As String = ""
        Dim Title As String = ""
        For I As Integer = 0 To allContentArr.Length - 1
            Link = "https://www.livescience" & Mid(allContentArr(I), 1, -1 + InStr(allContentArr(I), """"))
            RemoveWhatIsBefore(allContentArr(I), "h3 class=""article-name"">", False)
            RemoveWhatIsAfter(allContentArr(I), "</h3>", False)
            Title = allContentArr(I)

            If AddLinktoCheckedListBox1(Title:="Live Science" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Live Science : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub


    Sub ListBBC(ByVal Section As String)
        Dim Files As Integer = 0

        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\BBC" & Section & ".txt")


        'Dim pattern As String = "(<div Class=""article-hero__content-title b-font-family-serif b-font-weight-300|" &
        '    "<a Class=""article-title-card-rectangle__link article-title-card-rectangle__text-container"" target="""" rel="""" id="""" |" &
        '    "<a Class=""rectangle-story-item__title""|<a Class=""full-width-image-article__link"")Then"

        'Dim pattern As String = "(href= ""|<div Class=""full-width-image-article__image""><picture>)"
        Dim pattern As String = "href=""/" & Section

        allContentArr = Regex.Split(allContent, pattern)
        RemoveFirstCell(allContentArr)

        For I As Integer = 0 To allContentArr.Length - 1 'Step 2
            If Not allContentArr(I) Like "/article/########*" Then
                allContentArr(I) = ""
            End If
        Next

        RemoveEmptyORNull(SourceArray:=allContentArr, RemoveEmptyCells:=True, RemoveWhiteSpaces:=True)

        For I As Integer = 0 To allContentArr.Length - 1 'Step 2
            If Not (allContentArr(I) Like "*<!-- -->*" Or allContentArr(I) Like "*<h2 class=*" Or allContentArr(I) Like "*<span>*") Then
                allContentArr(I) = ""
            End If
        Next
        RemoveEmptyORNull(SourceArray:=allContentArr, RemoveEmptyCells:=True, RemoveWhiteSpaces:=True)


        Dim Link As String = ""
        Dim Title As String = ""
        Dim OpenningTag As String = ""
        Dim ClosingTag As String = ""
        For I As Integer = 0 To allContentArr.Length - 1
            Link = ""
            Title = ""


            If I = 2 Then 'allContentArr.Length - 2 Then
                Dim B As String = String.Empty
                B = ""
            End If

            Link = Mid(allContentArr(I), 1, -1 + InStr(allContentArr(I), """"))
            Link = "https://www.bbc.com/" & Section & "" & Link

            RemoveWhatIsBefore(allContentArr(I), "><", True)
            RemoveWhatIsBefore(allContentArr(I), ">", False)
            OpenningTag = allContentArr(I).Substring(0, InStr(allContentArr(I), " "))
            ClosingTag = OpenningTag.Replace("<", "</")
            RemoveWhatIsAfter(allContentArr(I), ClosingTag.Trim, False)


            'Title = allContentArr(I)
            '====================================
            Dim arrTitle As String()
            arrTitle = Split(allContentArr(I), "<")

            For J As Integer = LBound(arrTitle) To UBound(arrTitle)
                RemoveWhatIsBefore(arrTitle(J), ">", False)
                arrTitle(J) = Trim(arrTitle(J))
            Next
            RemoveEmptyORNull(SourceArray:=arrTitle, RemoveEmptyCells:=True, RemoveWhiteSpaces:=True)
            '====================================
            Title = Join(arrTitle, " ")


            If AddLinktoCheckedListBox1(Title:="BBC " & Section & ": " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="BBC " & Section & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub


    Function DecideWhichTextComeFirst(ByVal Delimiters As String(), ByVal Text As String) As String
        Dim Before As Integer = Text.Length - 1
        Dim ArrayIndex As Integer = 0

        For I As Integer = 0 To Delimiters.Length - 1
            If InStr(Text, Delimiters(I)) > 0 AndAlso InStr(Text, Delimiters(I)) < Before Then
                Before = InStr(Text, Delimiters(I))
                ArrayIndex = I
            End If
        Next

        Return Delimiters(ArrayIndex)
    End Function


    Sub ListGuardianPhotos()
        Exit Sub
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\file.txt")

        'allContentArr = Split(allContent, "<div class=""fc-item__container""> ")
        allContentArr = Split(allContent, "<div class=""u-cf index-page"" data-link-name=""Front | /news/series/ten-best-photographs-of-the-day"" role=""main"">")

        allContent = allContentArr(1)
        allContentArr = Split(allContent, "<h3 class=""fc-item__title"">")
        'data-link-name="article"

        Dim LinkEnd As Integer
        Dim Link As String
        Dim Title As String

        For I = 1 To allContentArr.Length - 1

            Try
                'Find Link ===========================
                LinkEnd = allContentArr(I).IndexOf("</a>")
                Link = allContentArr(I).Substring(0, LinkEnd + 4)
                Link = Link.Replace("<a href=""", "")
                LinkEnd = Link.IndexOf("""")
                Link = Link.Substring(0, LinkEnd)
                '=====================================

                Dim lcDate As String = Link
                lcDate = lcDate.Replace(IO.Path.GetFileNameWithoutExtension(lcDate), "")
                RemoveWhatIsBefore(lcDate, "gallery/", False)
                lcDate = replaceMonthByNum(lcDate)

                'Find Title ===========================
                LinkEnd = allContentArr(I).IndexOf("class=""js-headline-text"">")
                Title = allContentArr(I).Substring("class=""js-headline-text"">".Length + LinkEnd)
                LinkEnd = Title.IndexOf("<")
                Title = Title.Substring(0, LinkEnd)
                '=====================================

                If AddLinktoCheckedListBox1(Title:="Guardian Photos(" & lcDate & "): " & Title, Link:=Link) Then Files = Files + 1

            Catch ex As Exception

            End Try
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Guardian Photos : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub


    Sub DownLoadNewScientes()
        Label1.Text = "Downloading New Scientist.."
        Application.DoEvents()
        DownLoadFile("https://www.newscientist.com/",
            System.AppDomain.CurrentDomain.BaseDirectory & "/NewScientist.txt")
        '=================================================================================
        ListNewScientest()
    End Sub

    Sub DownloadAmricanScience()
        Label1.Text = "Downloading Amrican Science.."
        Application.DoEvents()


        DownLoadFile("https://www.scientificamerican.com/section/lateststories/",
            System.AppDomain.CurrentDomain.BaseDirectory & "/AmricanScience.txt")


        ListAmricanNews()

    End Sub

    Sub ListAmricanNews()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/AmricanScience.html")
        Dim Pattern As String = ""
        Pattern = ""
        Dim allContentArr() = Split(allContent, "<h2 class=""articleTitle-") ' Regex.Split(allContent, pattern:="<h2 class=""articleTitle-")

        ''allContentArr = Split(components(1), "<h2 class=""css-")
        RemoveFirstCell(allContentArr)
        ''RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr

            Dim Link As String = Str
            Link = Str
            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.scientificamerican.com" + Link





            Dim Title As String = Str

            RemoveWhatIsBefore(Title, "mtY5p""><p>", False)
            RemoveWhatIsAfter(Title, "</p>", False)


            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="Amrican Science : " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Amrican Science : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub

    Private Sub ListNewScientest()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/NewScientist.txt")
        Dim Pattern As String = ""
        Pattern = ""
        Dim allContentArr() = Regex.Split(allContent, pattern:="class=""CardLink")

        'allContentArr = Split(components(1), "<h2 class=""css-")
        RemoveFirstCell(allContentArr)
        'RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            If Str.Trim = "-container""><h3 class=""card__legal-heading""></h3><h2".Trim Then

            Else
                Dim Link As String = Str
                Dim Title As String = Str

                Title = Str
                RemoveWhatIsBefore(Title, "<h3 class=""Card__Title"">", False)
                RemoveWhatIsAfter(Title, "<", False)
                refineTitle(Text:=Title)

                Link = Str
                RemoveWhatIsBefore(Link, "href=""", False)
                RemoveWhatIsAfter(Link, """", False)
                Link = "https://www.newscientist.com/" + Link
                If Trim(Title) <> "" Then
                    If AddLinktoCheckedListBox1(Title:="New Scientist : " & Title, Link:=Link) Then Files = Files + 1
                End If
            End If
        Next
        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="New Scientist: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub DownLoadIndepenednet()
        Label1.Text = "Downloading independent science.."
        Application.DoEvents()
        DownLoadFile("https://www.independent.co.uk/news/science",
            System.AppDomain.CurrentDomain.BaseDirectory & "/independentScience.txt")

        '=================================================================================
        ListIndepenedent()

    End Sub

    Sub ListIndepenedent()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/independentScience.txt")
        Dim Pattern As String = ""
        Pattern = "<a class=""title"" href="""
        Dim allContentArr() = Regex.Split(allContent, pattern:=Pattern)


        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            Dim Link As String = Str
            Dim Title As String = Str

            RemoveWhatIsAfter(Link, """>", False)

            RemoveWhatIsBefore(Title, """>", False)
            RemoveWhatIsAfter(Title, "<", False)

            refineTitle(Text:=Title)
            If AddLinktoCheckedListBox1(Title:="The Independent" & ": " & Title, Link:="https://www.independent.co.uk/" & Link) Then Files = Files + 1

        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="The Independent: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub
    Sub DownloadEconomist()
        Dim Files As Integer = 0
        Files = Files + DownloadAndListEconomist_Middle_East_and_Africa()
        Files = Files + DownloadAndListEconomist_business()
        Files = Files + DownloadAndListEconomist_Finance()
        Files = Files + DownloadAndListEconomist_Science()
        Files = Files + DownloadAndListEconomist_culture()
        Files = Files + DownloadAndListEconomist_Essay()
        Files = Files + DownloadAndListEconomist_WorldAhead()
        Files = Files + DownloadAndListEconomist_OpenFuture()
        Files = Files + DownloadAndListEconomist_WhatIf()
        Files = Files + DownloadAndListEconomist_EconomistExplains()
    End Sub

    Sub Download_FT()
        Label1.Text = "Downloading FT.."
        Application.DoEvents()
        DownLoadFile("https://www.ft.com/science?page=1",
            System.AppDomain.CurrentDomain.BaseDirectory & "/" & "FT.txt")
        List_FT("FT.txt")

    End Sub

    Sub List_FT(ByVal FileName As String)
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)

        ' RemoveWhatIsBefore(allContent, "<div class=""column floatAll"">", False)

        'RemoveWhatIsAfter(allContent, "class=""pagination__wrapper___2qtdg""", False)

        Dim allContentArr() As String
        allContentArr = Split(allContent, "<div class=""o-teaser__heading"">")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            Dim Link As String = Str
            Dim Title As String = Str



            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.ft.com/" & Link
            RemoveWhatIsBefore(Title, "class=""js-teaser-heading-link"">", False)
            RemoveWhatIsAfter(Title, "<", False)
            refineTitle(Text:=Title)

            If AddLinktoCheckedListBox1(Title:="FT: " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="FT: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Function DownloadAndListEconomist_Middle_East_and_Africa() As Integer
        Label1.Text = "Downloading: Economist Middle East and Africa"
        Application.DoEvents()

        Call DownloadEconomist("https://www.economist.com/middle-east-and-africa", "Economist_Middle_East_and_Africa.txt")
        DownloadAndListEconomist_Middle_East_and_Africa = ListEconomist("The Econimics: Middle East & Africa ", "Economist_Middle_East_and_Africa.txt")
    End Function

    Function DownloadAndListEconomist_business() As Integer
        Label1.Text = "Downloading: Economist Business"
        Application.DoEvents()

        Dim FileName As String = "business.txt"
        Call DownloadEconomist("https://www.economist.com/business", FileName)
        DownloadAndListEconomist_business = ListEconomist("The Econimics: Business", FileName)
    End Function

    Function DownloadAndListEconomist_Finance() As Integer
        Label1.Text = "Downloading: Economist Finance"
        Application.DoEvents()

        Dim FileName As String = "Finance.txt"
        Call DownloadEconomist("https://www.economist.com/finance-and-economics", FileName)
        DownloadAndListEconomist_Finance = ListEconomist("The Econimics: Finance and Economist", FileName)
    End Function

    Function DownloadAndListEconomist_Science() As Integer
        Label1.Text = "Downloading: Economist Science"
        Application.DoEvents()

        Dim FileName As String = "Economist_Science.txt"
        Call DownloadEconomist("https://www.economist.com/science-and-technology", FileName)
        DownloadAndListEconomist_Science = ListEconomist("The Econimics: Science-and-Technology", FileName)
    End Function

    Function DownloadAndListEconomist_culture() As Integer
        Label1.Text = "Downloading: Economist Culture"
        Application.DoEvents()

        Dim FileName As String = "Economist_Culture.txt"
        Call DownloadEconomist("https://www.economist.com/culture", FileName)
        DownloadAndListEconomist_culture = ListEconomist("The Econimics: Culture", FileName)
    End Function

    Function DownloadAndListEconomist_Essay() As Integer
        Label1.Text = "Downloading: Economist Essay"
        Application.DoEvents()

        Dim FileName As String = "Economist_Essay.txt"
        Call DownloadEconomist("https://www.economist.com/essay", FileName)
        DownloadAndListEconomist_Essay = ListEconomist("The Econimics: Essay", FileName)
    End Function

    Function DownloadAndListEconomist_WorldAhead() As Integer
        Label1.Text = "Downloading: Economist World Ahead"
        Application.DoEvents()

        Dim FileName As String = "Economist_World_Ahead.txt"
        Call DownloadEconomist("https://www.economist.com/the-world-ahead-2023", FileName)
        'DownloadAndListEconomist_WorldAhead = ListEconomist_wordAheads("The Econimics: World Ahead", FileName)
        DownloadAndListEconomist_WorldAhead = ListEconomist("The Econimics: World Ahead", FileName)
    End Function

    Function DownloadAndListEconomist_OpenFuture() As Integer
        Label1.Text = "Downloading: Economist Open Future"
        Application.DoEvents()

        Dim FileName As String = "Economist_open_future.txt"
        Call DownloadEconomist("https://www.economist.com/open-future", FileName)
        DownloadAndListEconomist_OpenFuture = ListEconomist("The Econimics: Open Future", FileName)
    End Function

    Function DownloadAndListEconomist_WhatIf() As Integer
        Label1.Text = "Downloading: Economist What If"
        Application.DoEvents()

        Dim FileName As String = "Economist_WhatIf.txt"
        Call DownloadEconomist("https://www.economist.com/what-if-2021", FileName)

        DownloadAndListEconomist_WhatIf = ListEconomistWhatIf("The Econimics: What IF", FileName)
    End Function

    Function DownloadAndListEconomist_EconomistExplains() As Integer
        Label1.Text = "Downloading: Economist Explains"
        Application.DoEvents()

        Dim FileName As String = "Economist_the_economist_explains.txt"
        Call DownloadEconomist("https://www.economist.com/the-economist-explains", FileName)
        DownloadAndListEconomist_EconomistExplains = ListEconomist("The Econimics: The Economist Explains", FileName)
    End Function


    Function ListEconomistWhatIf(ByVal Category As String, ByVal FileName As String) As Integer
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)

        Dim allContentArr() As String
        ' allContentArr = Split(allContent,)
        allContentArr = Regex.Split(allContent, "<h3 class=""css-")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr

            RemoveWhatIsBefore(Str, "<a href=""", False)

            Dim Link As String = Str
            Dim Title As String = Str

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)

            refineTitle(Text:=Title)
            If AddLinktoCheckedListBox1(Title:=Category & ":  " & Title, Link:="https://www.economist.com/" & Link) Then Files = Files + 1
        Next

        ListEconomistWhatIf = Files

    End Function
    Sub DownloadEconomist(ByVal Link As String, ByVal FileName As String)
        DownLoadFile(Link,
            System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName)
    End Sub
    Function ListEconomist_wordAheads(ByVal Category As String, ByVal FileName As String)
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)

        Dim allContentArr() As String
        allContentArr = Split(allContent, "<h3 type=""world-ahead"" class=""")

        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            Dim Link As String = Str
            Dim Title As String = Str

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            RemoveWhatIsBefore(Title, "<span class=""_divider""> | </span></span>", False)
            RemoveWhatIsAfter(Title, "<", False)

            refineTitle(Text:=Title)
            If AddLinktoCheckedListBox1(Title:=Category & ": " & Title, Link:="https://www.economist.com/" & Link) Then Files = Files + 1

        Next
        ListEconomist_wordAheads = Files
    End Function

    Function ListEconomist(ByVal Category As String, ByVal FileName As String) As Integer
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)

        Dim allContentArr() As String
        allContentArr = Split(allContent, "<h3 class=""css-")

        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            Dim Link As String = Str
            Dim Title As String = Str

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            RemoveWhatIsBefore(Title, "><", False)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)

            refineTitle(Text:=Title)
            If AddLinktoCheckedListBox1(Title:=Category & ": " & Title, Link:="https://www.economist.com/" & Link) Then Files = Files + 1
        Next
        ListEconomist = Files
    End Function

    Sub DownloadAP()
        Label1.Text = "Downloading AP.."
        Application.DoEvents()
        DownLoadFile("https://apnews.com/science",
            System.AppDomain.CurrentDomain.BaseDirectory & "/AP.txt")

        ListAP()
    End Sub

    Sub ListGlobalSearchScience()
        'ListGlobalResearch("GlobalResearch.txt")
        ListGlobalResearchCategory(Section:="Global Search: Science", FileName:="GlobalSearchScience.txt")
        ListGlobalResearchCategory(Section:="Global Search: Economy", FileName:="GlobalSearchEconomy.txt")
        ListGlobalResearchCategory(Section:="Global Search: US-Nato War", FileName:="GlobalSearchUSNatoWar.txt")
        ListGlobalResearchCategory(Section:="Global-Search Civil Rights", FileName:="GlobalSearchCivilRights.txt")
        ListGlobalResearchCategory(Section:="Global Search: Environment ", FileName:="GlobalSearchEenvironment.txt")
        ListGlobalResearchCategory(Section:="Global Search: Poverty", FileName:="GlobalSearchPoverty.txt")
        ListGlobalResearchCategory(Section:="Global Search: Media", FileName:="GlobalSearchMedia.txt")
        ListGlobalResearchCategory(Section:="Global Search: 911", FileName:="GlobalSearch911.txt")
        ListGlobalResearchCategory(Section:="Global Search: War-Crime", FileName:="GlobalSearchWarCrime.txt")
        ListGlobalResearchCategory(Section:="Global Search: Militry", FileName:="GlobalSearchMilitry.txt")
        ListGlobalResearchCategory(Section:="Global Search: History", FileName:="GlobalSearchHistory.txt")
    End Sub

    Sub ListGlobalResearchCategory(ByVal Section As String, ByVal FileName As String)
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)

        Dim allContentArr() As String
        allContentArr = Split(allContent, "<strong class=""title"">")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            Dim Link As String = Str
            Dim Title As String = Str

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            RemoveWhatIsAfter(Title, "</a>", False)

            RemoveWhatIsBefore(Link, "<a href=""", False)

            Title = StrReverse(Title)
            RemoveWhatIsAfter(Title, ">", False)

            Title = StrReverse(Title)
            refineTitle(Text:=Title)


            AddLinktoCheckedListBox1(Title:=Section & ": " & Title, Link:=Link)



        Next
    End Sub


    Sub ListAP()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/AP.txt")
        '<h2 class=""css-
        Dim allContentArr() As String = Split(allContent, "<h3 class=""PagePromo-title"">")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr

            Dim Link As String = Str
            Dim Title As String = Str

            Title = Str
            RemoveWhatIsBefore(Title, "><", False)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)


            Link = Str
            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://apnews.com/" + Link



            'MsgBox(Link & vbCrLf & vbCrLf & Title)
            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="AP: " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="AP: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub



    Sub DownloadReuters()
        Label1.Text = "Downloading Reuters.."
        Application.DoEvents()
        DownLoadFile("http://www.reuters.com/science/",
           System.AppDomain.CurrentDomain.BaseDirectory & "/reuters.txt")
        Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "/reuters.txt")
        ' ListReuters()
    End Sub

    Sub ListReuters()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/reuters.txt")
        '<h2 class=""css-
        Dim allContentArr() As String = Split(allContent, "<a data-testid=""Heading""")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr

            Dim Link As String = Str
            Dim Title As String = Str


            Link = Str
            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.reuters.com/" + Link

            Title = Str

            RemoveWhatIsBefore(Title, "<span>", False)
            'RemoveWhatIsBefore(Title, "</span>", False)

            'RemoveWhatIsBefore(Title, "<p class="" Text__text", False)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)

            'MsgBox(Link & vbCrLf & vbCrLf & Title)
            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="Reuters : " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Reuters: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub DownloadNewYorkTimes()
        Label1.Text = "Downloading New York Times.."
        Application.DoEvents()
        DownLoadFile("https://www.nytimes.com/international/section/science/",
            System.AppDomain.CurrentDomain.BaseDirectory & "/NYTimes.txt")

        'ListNewYorkTimes()

    End Sub

    Sub ListNewYorkTimes()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/NYTimes.txt")
        Dim Components() As String = Split(allContent, "<section id=""stream-panel""")

        Dim allContentArr() As String = Split(Components(1), "<h3 class=""css-")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr

            Dim Link As String = Str
            Dim Title As String = Str


            Link = Str
            RemoveWhatIsBefore(Link, " href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = " https://www.nytimes.com/" + Link

            Title = Str
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)
            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="New York Times : " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next


        allContentArr = Split(Components(1), "<li class=""css-112uytv"">")
        'allContentArr = Regex.Split(Components(1), "(<a href="")")
        RemoveFirstCell(allContentArr)
        'RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            Dim Link As String = Str
            Dim Title As String = Str
            Title = Str
            RemoveWhatIsBefore(Title, "<h2 class=""css-1", False)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)

            Link = Str
            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.nytimes.com/" + Link
            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="New York Times : " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="New York Times: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub



    Sub DownloadLosAngelesTimes()
        Label1.Text = "Downloading Los Angeles Times.."
        Application.DoEvents()
        DownLoadFile("https://www.latimes.com/science",
            System.AppDomain.CurrentDomain.BaseDirectory & "/LATimes.txt")


        ListLosAngelesTimes()

    End Sub




    Sub ListLosAngelesTimes()
        Dim Files As Integer = 0

        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/LATimes.txt")

        Dim allContentArr() As String = Split(allContent, "<h2 class=""promo-")
        RemoveFirstCell(allContentArr)


        For Each Str As String In allContentArr

            Dim Link As String = Str
            Dim Title As String = Str

            If InStr(Str, "/story/") Then
                Link = Str
                RemoveWhatIsBefore(Link, "<a class=""link"" href=""", False)
                RemoveWhatIsAfter(Link, """", False)

                Title = Str
                RemoveWhatIsBefore(Title, ">", False)
                RemoveWhatIsBefore(Title, ">", False)
                RemoveWhatIsAfter(Title, "<", False)

                If AddLinktoCheckedListBox1(Title:="Los Angeles Times : " & Title, Link:=Link) Then Files = Files + 1

            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Los Angeles Times: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub



    Sub DownloadPopSci()
        Dim Files As Integer = 0
        Files = Files + DownLoadPopSci_Science()
        Files = Files + DownLoadPopSci_technology()
        Files = Files + DownLoadPopSci_diy()
        Files = Files + DownLoadPopSci_reviews()

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Sci Pop: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub

    Function DownLoadPopSci_Science() As Integer

        Dim Category As String = "Science"
        SetPopSciDownloadingParameters(Category:=Category)
        DownLoadPopSci_Science = ListPopSci(Category)
    End Function

    Function DownLoadPopSci_technology() As Integer
        Dim Category As String = "technology"
        SetPopSciDownloadingParameters(Category:=Category)
        DownLoadPopSci_technology = ListPopSci(Category)
    End Function

    Function DownLoadPopSci_diy() As Integer
        Dim Category As String = "diy"
        SetPopSciDownloadingParameters(Category:=Category)
        DownLoadPopSci_diy = ListPopSci(Category)
    End Function

    Function DownLoadPopSci_reviews() As Integer
        Dim Category As String = "reviews"
        SetPopSciDownloadingParameters(Category:=Category)
        DownLoadPopSci_reviews = ListPopSci(Category)
    End Function

    Function ListPopSci(ByVal Category As String) As Integer
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\PopScience_" & Category & ".txt")

        Dim allContentArr() As String
        allContentArr = Regex.Split(allContent, "<a class=""card-post-title-link"" ")

        RemoveFirstCell(allContentArr)

        For I As Integer = LBound(allContentArr) To UBound(allContentArr)

            Dim Link As String = allContentArr(I)
            Dim Title As String = allContentArr(I)
            Dim str = allContentArr(I)

            Link = allContentArr(I)
            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            RemoveWhatIsBefore(Title, "<span class=""mobile lg:hidden"">", False)
            RemoveWhatIsAfter(Title, "<", False)

            If AddLinktoCheckedListBox1(Title:="Sci Pop (" & Category & "): " & Title, Link:=Link) Then Files = Files + 1

        Next

        ListPopSci = Files
    End Function


    Sub SetPopSciDownloadingParameters(ByVal Category As String)
        Exit Sub
        ImplementDownloading(Message:="Downloading Pop Science - " & Category & "..",
                             Link:="https://www.popsci.com/" & Category & "/",
                             Category:=Category)
    End Sub


    Sub ImplementDownloading(ByVal Message As String,
                             ByVal Link As String,
                             ByVal Category As String)

        Label1.Text = Message
        Application.DoEvents()
        DownLoadFile(Link,
            System.AppDomain.CurrentDomain.BaseDirectory & "/PopScience_" & Category & ".txt")
    End Sub




    Sub DownloadGlobalResearch()
        Dim Files As Integer = 0
        Files = Files + downloadLatestNews_and_TopStories()
        Files = Files + download_US_Nato_War()
        Files = Files + download_Economy()
        Files = Files + download_PoliceStateCivilRights()
        Files = Files + download_Environment()
        Files = Files + download_PovertyAndSocialInequalit()
        Files = Files + download_MediaDisInformation()
        Files = Files + download_Terrorism()
        Files = Files + download_CrimesAgainstHumanity()
        Files = Files + download_MilitarizationAndWMD()
        Files = Files + download_History()
        Files = Files + download_ScienceAndMedicine()


        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Global Researches: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

        IO.File.WriteAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\Logs_20260127_1000.txt", "")

    End Sub

    Function downloadLatestNews_and_TopStories() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/latest-news-and-top-stories",
                                          msg:="Downloading Global Research: Latest News and Top Stories",
                                          FileName:="Latest", CategoryName:="Latest News and Top Stories")


    End Function
    Function download_US_Nato_War() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/us-nato-war-agenda",
                                          msg:="Downloading Global Research: US NATO War Agenda",
                                          FileName:="us-nato-war-agenda", CategoryName:="US NATO War Agenda")
    End Function

    Function download_Economy() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/global-economy",
                                  msg:="Downloading Global Research: Global Economy",
                                  FileName:="GlobalEconomy", CategoryName:="Global Economy")
    End Function

    Function download_PoliceStateCivilRights() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/police-state-civil-rights",
                                  msg:="Downloading Global Research: Police State & Civil Rights",
                                  FileName:="PoliceStateCivilRights", CategoryName:="Police State & Civil Rights")
    End Function

    Function download_Environment() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/environment",
                                  msg:="Downloading Global Research: Environment",
                                  FileName:="Environment", CategoryName:="Environment")
    End Function

    Function download_PovertyAndSocialInequalit() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/poverty-social-inequality",
                                  msg:="Downloading Global Research: Poverty & Social Inequality",
                                  FileName:="PovertyAndSocialInequality", CategoryName:="Poverty & Social Inequality")
    End Function

    Function download_LawAndJustice() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/law-and-justice",
                                  msg:="Downloading Global Research: Media Disinformation",
                                  FileName:="LawAndJustice", CategoryName:="Media Disinformation")
    End Function

    Function download_MediaDisInformation() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/media-disinformation",
                                  msg:="Downloading Global Research: Law and Justice",
                                  FileName:="media-disinformation", CategoryName:="Media Disinformation")
    End Function

    Function download_Terrorism() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/9-11-war-on-terrorism",
                                  msg:="Downloading Global Research: Terrorism",
                                  FileName:="Terrorism", CategoryName:="Terrorism")
    End Function
    Function download_CrimesAgainstHumanity() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/crimes-against-humanity",
                                  msg:="Downloading Global Research: Crimes against Humanity",
                                  FileName:="CrimesAgainstHumanity", CategoryName:="Crimes against Humanity")
    End Function
    Function download_MilitarizationAndWMD() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/militarization-and-wmd",
                                  msg:="Downloading Global Research: Militarization and WMD",
                                  FileName:="MilitarizationAndWMD", CategoryName:="Militarization And WMD")
    End Function
    Function download_History() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/culture-society-history",
                                  msg:="Downloading Global Research: History",
                                  FileName:="History", CategoryName:="History")
    End Function
    Function download_ScienceAndMedicine() As Integer
        Return DownLoadGlobalResearchSubCategory(Site:="https://www.globalresearch.ca/theme/science-and-medicine",
                                  msg:="Downloading Global Research: Science And Medicine",
                                  FileName:="ScienceAndMedicine", CategoryName:="Science And Medicine")
    End Function

    Function DownLoadGlobalResearchSubCategory(ByVal Site As String,
                 ByVal msg As String,
                 ByVal FileName As String,
                 ByVal CategoryName As String) As Integer

        Label1.Text = msg
        Application.DoEvents()
        DownLoadFile(Site,
        System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName & ".txt")


        Return ListGlobalResearch(CategoryName:=CategoryName, FileName:=FileName & ".txt")
    End Function

    Function ListGlobalResearch(ByVal CategoryName As String, ByVal FileName As String) As Integer
        Dim Files As Integer = 0

        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)

        RemoveWhatIsBefore(allContent, "<div class=""column floatAll"">", False)

        Dim allContentArr() As String
        allContentArr = Split(allContent, "<strong class=""title"">")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr
            Dim Link As String = Str
            Dim Title As String = Str



            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            RemoveWhatIsAfter(Title, "</a>", False)
            RemoveWhatIsBefore(Link, "<a href=""", False)

            Title = StrReverse(Title)
            RemoveWhatIsAfter(Title, ">", False)
            Title = StrReverse(Title)
            refineTitle(Text:=Title)

            If AddLinktoCheckedListBox1(Title:="Global Research (" & CategoryName & "): " & Title, Link:=Link) Then Files = Files + 1
        Next

        ListGlobalResearch = Files
    End Function

    Sub DownloadNationalInterset()
        Dim Files As Integer = 0
        NewDic.Clear()
        Dim lcDic As New Dictionary(Of String, String)

        DownloadNationalInterset_Iran(lcDic)
        DownloadNationalInterset_SaudiArabia(lcDic)
        DownloadNationalInterset_Bahrain(lcDic)
        DownloadNationalInterset_Yemen(lcDic)
        DownloadNationalInterset_recentstories(lcDic)

        Dim Category As String = ""
        For Each KV As KeyValuePair(Of String, String) In lcDic

            If NewDic.ContainsKey(KV.Key) = True Then
                Category = Join((NewDic(KV.Key).ToList).ToArray, ",")
                Category = "(" & Category & ")"
            End If

            If AddLinktoCheckedListBox1(Title:="The National Interest " & Category & " : " & KV.Value, Link:=KV.Key) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="The National Interest " & Category & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub
    Sub DownloadNationalInterset_Iran(ByRef lcNewDic As Dictionary(Of String, String))
        ImplementDownloading(NewDic, lcNewDic, "Iran")
    End Sub
    Sub DownloadNationalInterset_Yemen(ByRef lcNewDic As Dictionary(Of String, String))
        ImplementDownloading(NewDic, lcNewDic, "Yemen")
    End Sub
    Sub DownloadNationalInterset_Bahrain(ByRef lcNewDic As Dictionary(Of String, String))
        ImplementDownloading(NewDic, lcNewDic, "Bahrain")
    End Sub

    Sub DownloadNationalInterset_SaudiArabia(ByRef lcNewDic As Dictionary(Of String, String))
        ImplementDownloading(NewDic, lcNewDic, "saudi-arabia")
    End Sub

    Sub DownloadNationalInterset_recentstories(ByRef lcNewDic As Dictionary(Of String, String))
        ImplementDownloading(NewDic, lcNewDic, "recent-stories")
    End Sub



    Sub ImplementDownloading(ByRef NewDic As Dictionary(Of String, List(Of String)),
                             ByRef lcNewDic As Dictionary(Of String, String),
                             ByVal Category As String)
        Dim Link As String = "https://nationalinterest.org/"
        If Category = "recent-stories" Then
            Link = Link & Category
        Else
            Link = Link & "topic/" & Category
        End If

        DownloadNationalInterset(Link:=Link,
                                 Msg:="Downloading National Interest/" & Category & "..",
                                 FileName:="NationalInterest-" & Category & ".txt")
        ListTheNationalInterset(FileName:="NationalInterest-" & Category & ".txt", NewDic:=NewDic, lcNewDic:=lcNewDic, Category:=Category)
    End Sub


    Sub DownloadNationalInterset(ByVal Link As String,
                                 ByVal Msg As String,
                                 ByVal FileName As String)
        Label1.Text = Msg
        Application.DoEvents()

        DownLoadFile(Link, System.AppDomain.CurrentDomain.BaseDirectory & FileName)
    End Sub


    Sub ListTheNationalInterset(ByVal FileName As String,
                                ByRef NewDic As Dictionary(Of String, List(Of String)),
                                ByRef lcNewDic As Dictionary(Of String, String),
                                ByVal Category As String)

        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & FileName)
        allContentArr = Split(allContent, "<h3 class=""article__title"">")
        Dim A As String = ""
        RemoveFirstCell(allContentArr)

        Dim Link As String = String.Empty
        Dim Title As String = String.Empty

        'Dim SubArr() As String
        For i As Integer = 0 To allContentArr.Length - 1
            Link = allContentArr(i)
            RemoveWhatIsBefore(Link, """", False)
            Link = "https://nationalinterest.org" & Link.Substring(0, -1 + InStr(Link, """"))

            Title = allContentArr(i)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)

            refineTitle(Title)

            'AddLinktoCheckedListBox1(Title:="The National Interest" & "-" & Category & ": " & Title, Link:=Link)
            If lcNewDic.ContainsKey(Link) = False Then
                lcNewDic.Add(Link, Title)
            End If


            Dim L As New List(Of String)
            If NewDic.ContainsKey(Link) Then
                L = NewDic(Link)
                NewDic.Remove(Link)
                'Else
                '    L.Add(Category)
            End If
            L.Add(Category)
            NewDic.Add(Link, L)
        Next

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'add title to Clipboard to let Chatgpt classfy them
        'Dim Title As String = ""
        'Dim TitlesList As New List(Of String)
        'For Each ListItem As Object In CheckedListBox1.CheckedItems
        '    Title = CType(ListItem, Item).Title
        '    Title = Mid(Title, InStr(Title, ":") + 1)
        '    If Not Title Like "*No Downloads*" Then TitlesList.Add(Title)
        'Next
        'Dim CHATGPTCommand As String = "using the classifications in the attached file, classify and categorize the following titles into their best fit according to the full hierarchy. "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " only show the category index prefix (like 1.1.2.3 Title...) without listing or showing any category names. "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " omit empty categories. "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " each line should start with its full classification index followed by the title. "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " make sure no title is skipped. "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " you can assign a title to more than one category if applicable. "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " if a title doesn’t fit any category, put it under ""7. UNCATEGORIZED / MISCELLANEOUS"". "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " make a downloadable txt file named " & Now.ToString("yyyyMMdd") & "_indexed.txt "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & " Preserve totles exactly as written: "
        'CHATGPTCommand = CHATGPTCommand & vbCrLf & ""


        'Clipboard.SetText(CHATGPTCommand & vbCrLf & Join(TitlesList.ToArray, vbCrLf))
        'end of: add title to Clipboard to let Chatgpt classfy them ===================================================
        'Exit Sub
        'get the classified titles ===============
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Dim ofd As New OpenFileDialog()
        ' Set dialog properties
        ofd.Title = "Select a Text File"
        ofd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        ofd.Multiselect = False

        '' Show dialog
        'If ofd.ShowDialog() = DialogResult.OK Then
        '    dict = BringClassifiedTitles(ofd.FileName)
        'Else
        '    Exit Sub
        'End If
        ''End of: get the classified titles ===============


        'Dim listOfDirectories As New List(Of String)
        'Dim listOfLinks As New List(Of String)

        'getDirsAndLinks(listOfDirectories, listOfLinks)

        'For I As Integer = 0 To listOfDirectories.Count - 1
        '    analysFile(listOfDirectories(I))
        'Next


        For Each ListItem As Object In CheckedListBox1.CheckedItems
            Dim LcItem As Item = CType(ListItem, Item)

            Dim Category As String = ""
            Try
                Dim L As New List(Of String)
                L = NewDic(LcItem.Link)
                Category = vbTab & Join(L.ToArray, ",")
            Catch ex As Exception

            End Try
            Dim xTITLE As String = Mid(LcItem.Title, InStr(LcItem.Title, ":") + 1).Trim()

            If (LcItem.Title & vbTab & Category) Like "*No Downloads*" Then
                IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\done.txt",
                    LcItem.Link & vbTab & Now.Date.ToString("yyyyMMdd") & vbTab & LcItem.Title & vbTab & Category & vbTab & "8" & vbCrLf)
            Else
                Try
                    IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\done.txt",
                          LcItem.Link & vbTab & Now.Date.ToString("yyyyMMdd") & vbTab & LcItem.Title & vbTab & Category & vbTab & dict(xTITLE) & vbCrLf)
                Catch ex As Exception
                    IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\done.txt",
                                          LcItem.Link & vbTab & Now.Date.ToString("yyyyMMdd") & vbTab & LcItem.Title & vbTab & Category & vbTab & "7" & vbCrLf)
                End Try
            End If
        Next

        Button3.Enabled = False
    End Sub

    Function BringClassifiedTitles(lcfilePath As String) As Dictionary(Of String, String)

        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)



        For Each line As String In File.ReadLines(lcfilePath)
            If String.IsNullOrWhiteSpace(line) Then Continue For

            ' Split the line into index and title parts
            Dim firstSpace As Integer = line.IndexOf(" "c)
            If firstSpace = -1 Then Continue For

            Dim index As String = line.Substring(0, firstSpace).Trim()
            Dim title As String = line.Substring(firstSpace + 1).Trim()

            ' Add to dictionary or append index if already exists
            If dict.ContainsKey(title) Then
                dict(title) &= "," & index
            Else
                dict.Add(title, index)
            End If
        Next

        Return dict

    End Function

    Function replaceMonthByNum(ByRef FolderName As String)
        For I As Integer = 1 To 12
            Dim DT As New Date(2020, I, 1)
            If InStr(FolderName, DT.ToString("MMM").ToLower) Then
                FolderName = FolderName.Replace(DT.ToString("MMM").ToLower, DT.Month.ToString("00"))
                FolderName = FolderName.Replace("/", "")
                Exit For
            End If
        Next
        replaceMonthByNum = FolderName
    End Function
    '============

    Sub getDirsAndLinks(ByRef listOfDirectories As List(Of String), ByRef listOfLinks As List(Of String))


        For Each ListItem As Object In CheckedListBox1.CheckedItems
            Dim LcItem As Item = CType(ListItem, Item)
            If InStr(LcItem.Link, "https://www.theguardian.com/") > 0 And InStr(LcItem.Title, "Guardian Photos") > 0 Then
                'Get Folder Name from Link
                '   A-get date from link
                Dim FolderName As String = LcItem.Link
                FolderName = FolderName.Replace(IO.Path.GetFileNameWithoutExtension(FolderName), "")
                RemoveWhatIsBefore(FolderName, "gallery/", False)

                '   B-substitute month name by its number ==========================
                FolderName = replaceMonthByNum(FolderName)
                'For I As Integer = 1 To 12
                '    Dim DT As New Date(2020, I, 1)
                '    If InStr(FolderName, DT.ToString("MMM").ToLower) Then
                '        FolderName = FolderName.Replace(DT.ToString("MMM").ToLower, DT.Month.ToString("00"))
                '        FolderName = FolderName.Replace("/", "")
                '        Exit For
                '    End If
                'Next
                '============================================================

                '   C-create the folder================================
                Dim Counter As String = ""
                'If IO.Directory.Exists(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FolderName & Counter) = False Then
                '    IO.Directory.CreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FolderName & Counter)
                'End If

                While IO.Directory.Exists(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FolderName & Counter)
                    If Counter = "" Then
                        Counter = "_001"
                    Else
                        Counter = Counter.Replace("_", "")
                        Counter = Format(CInt(Counter) + 1, "000")
                        Counter = "_" & Counter
                    End If

                End While
                IO.Directory.CreateDirectory(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FolderName & Counter)
                '======================================================

                DownLoadFile(URL:=LcItem.Link, FileName:=FolderName & "\Photos.txt")
                listOfDirectories.Add(FolderName)
                listOfLinks.Add(LcItem.Link)


            End If

        Next
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



    Sub analysFile(ByVal FolderPath As String)
        IO.File.WriteAllText(FolderPath & "\Entries.txt", "")

        Dim allContent As String
        Dim FilePath As String
        FilePath = FolderPath & "\Photos.txt"
        allContent = IO.File.ReadAllText(FilePath)

        '<div class="content__main-column content__main-column--gallery">
        RemoveWhatIsBefore(Text:=allContent, MarkingText:="<div class=""content__main-column content__main-column--gallery"">", KeepMarkingText:=False)
        Dim allContentArr() As String
        allContentArr = Split(allContent, "<li id=""img-")

        'Remove First element=======================
        RemoveFirstCell(allContentArr)
        '===========================================
        Dim Dic As New Dictionary(Of String, String)
        Dim L As New List(Of Integer)

        For I As Integer = 1 To allContentArr.Length - 1
            Dim PhotoText As String = ""
            PhotoText = allContentArr(I)
            RemoveWhatIsBefore(PhotoText, "<!--[if IE 9]></video><![endif]-->", False)
            RemoveWhatIsBefore(PhotoText, "alt=""", False)
            RemoveWhatIsAfter(PhotoText, """", False)

            RemoveWhatIsBefore(allContentArr(I), "<!--[if IE 9]><video style=""display: none;""><![endif]-->", False)
            RemoveWhatIsAfter(allContentArr(I), "<!--[if IE 9]></video><![endif]-->", False)

            Dim allPhotos() As String
            allPhotos = Split(allContentArr(I), "srcset=""")

            RemoveFirstCell(allPhotos)

            For J As Integer = 0 To allPhotos.Length - 1
                RemoveWhatIsAfter(allPhotos(J), " ", False)
            Next

            'MsgBox(allPhotos(0))
            Dim FileName As String
            FileName = allPhotos(0)
            RemoveWhatIsAfter(FileName, "?", False)
            FileName = IO.Path.GetFileName(FileName)

            Dim Count As Integer = 1
            Dim TempFileName = FileName
            While IO.File.Exists(FolderPath & "\" & TempFileName)
                TempFileName = IO.Path.GetFileNameWithoutExtension(FileName) & "" & Count & IO.Path.GetExtension(FileName)
                Count = Count + 1
            End While
            FileName = TempFileName

            DownLoadFile(allPhotos(0), FolderPath & "\" & IO.Path.GetFileName(FileName))

            'IO.File.AppendAllText(FolderPath & "\" & "Entrries.txt", IO.Path.GetFileName(FileName) & vbTab & PhotoText & vbCrLf)
            Dic.Add(IO.Path.GetFileNameWithoutExtension(FileName), PhotoText)
            L.Add(CInt(IO.Path.GetFileNameWithoutExtension(FileName)))

        Next

        L.Sort()
        For I As Integer = 0 To L.Count - 1
            IO.File.AppendAllText(FolderPath & "\" & "Entries.txt", L(I) & ".jpg" & vbTab & Dic(L(I)) & vbCrLf)
        Next


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Name As Integer
        Name = TreeView1.Nodes.Count + 1


        Dim TN As New TreeNode()
        TN.Name = Name
        TN.Text = TextBox3.Text

        TreeView1.Nodes.Add(TN)
        TextBox3.Text = ""
    End Sub

    Function getMatch(ByVal OrgText As String, ByVal lcMatchingText As String)
        getMatch = ""
        For J As Integer = 0 To OrgText.Length - lcMatchingText.Length - 1
            If OrgText.Substring(J, lcMatchingText.Length) Like lcMatchingText Then
                getMatch = OrgText.Substring(J, lcMatchingText.Length)
                Exit Function
            End If
        Next

    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Process.Start(System.AppDomain.CurrentDomain.BaseDirectory)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        For I As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(I, True)
        Next
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim lstTitle As New List(Of String)
        glTitle = ""

        CheckedListBox1.Items.Clear()
        Dim Temp() As String

        Dim Match As Boolean = False
        If Dic.Count = 0 Then LoadLinksInotCollection()

        Dim indxTitle As Integer = 2
        Dim indxDate As Integer = 1
        Dim indxIndex As Integer = 4

        Dim SearchField As New List(Of Integer)
        Dim SearchValues As New List(Of String)
        glTitle = "All Entries "
        Select Case True
            Case TextBox2.Enabled = False And DateTimePicker1.Enabled = False
                'this means, the search field will be date and for today date...
                SearchField.Add(indxDate)
                SearchValues.Add("*" & DateTimePicker1.Value.ToString("yyyyMMdd") & "*")
                glTitle = "All Entries of " & DateTimePicker1.Value.ToString("yyyy-MM-dd")
            Case TextBox2.Enabled = True And DateTimePicker1.Enabled = False
                SearchField.Add(indxTitle)
                SearchValues.Add("*" & TextBox2.Text.ToLower.Trim & "*")

                SearchField.Add(indxIndex)
                SearchValues.Add("*" & TextBox2.Text.ToLower.Trim & "*")
                glTitle = "All Entries containing '" & TextBox2.Text.ToLower.Trim & "'"
            Case TextBox2.Enabled = False And DateTimePicker1.Enabled = True
                SearchField.Add(indxDate)
                SearchValues.Add("*" & DateTimePicker1.Value.ToString("yyyyMMdd") & "*")
                glTitle = "All Entries of " & DateTimePicker1.Value.ToString("yyyy-MM-dd")
            Case TextBox2.Enabled = True And DateTimePicker1.Enabled = True
                SearchField.Add(indxDate)
                SearchValues.Add("*" & DateTimePicker1.Value.ToString("yyyyMMdd") & "*")

                SearchField.Add(indxTitle)
                SearchValues.Add("*" & TextBox2.Text.ToLower.Trim & "*")

                SearchField.Add(indxIndex)
                SearchValues.Add("*" & TextBox2.Text.ToLower.Trim & "*")

                glTitle = "All Entries of " & DateTimePicker1.Value.ToString("yyyy-MM-dd") & " containing '" & TextBox2.Text.ToLower.Trim & "'"
        End Select


        Dim LineComp() As String
        For Each KV As KeyValuePair(Of String, String) In Dic
            Match = False

            Temp = Split(KV.Key & vbTab & KV.Value, vbTab)
            If Temp.Length < 5 Then
                While Temp.Length < 5
                    Dim l As New List(Of String)
                    l = Temp.ToList
                    l.Add("7")
                    Temp = l.ToArray
                End While
            End If



            For I As Integer = 0 To SearchField.Count - 1
                Dim A As String = Temp(SearchField(I)).ToLower
                Dim B As String = SearchValues(I)
                If SearchField(I) <> 4 Then
                    If Temp(SearchField(I)).ToLower Like SearchValues(I) Then
                        Match = True
                        Exit For
                    Else
                        Match = False
                    End If
                Else
                    Dim tempArr As String()
                    tempArr = Split(Temp(SearchField(I)), ",")
                    For Each index As String In tempArr
                        If index = SearchValues(I).Replace("*", "") Then
                            Match = True
                            Exit For
                        End If
                    Next
                End If


            Next I

            If Match Then
                LineComp = Split(KV.Value, vbTab)
                Try
                    CheckedListBox1.Items.Add(New Item(Title:=LineComp(1), Link:=KV.Key, Indexes:=LineComp(3)))
                Catch ex As Exception
                    CheckedListBox1.Items.Add(New Item(Title:=LineComp(1), Link:=KV.Key, Indexes:="8"))
                End Try
            End If
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim Site As String = ""
        Dim Color As String = "FFFFF"
        Dim FileIsEmpty As Boolean = True


        IO.File.WriteAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "")
        addUpperPartOfHTMLpage()

        Dim TitlePhrase As String = IO.File.ReadAllText("TitlePhrase.txt")
        Dim UnorderListPhrase As String = IO.File.ReadAllText("UnorderListPhrase.txt")

        Dim myDic As New Dictionary(Of String, List(Of Object))


        Select Case CheckBox4.Checked
            Case False
                'iterate throw checkboxlist to get the link and title ========================================
                For Each item As Item In CheckedListBox1.Items
                    Site = getTitleWithoutPranthesesOrColon(item.Title).Trim
                    Dim L As New List(Of Object)

                    If myDic.ContainsKey(Site) Then
                        L = CType(myDic(Site), Object)
                        myDic.Remove(Site)
                    End If

                    L.Add(item)
                    myDic.Add(Site, L)
                Next
                'end of: iterate throw checkboxlist to get the link and title ===============================




                For Each KV As KeyValuePair(Of String, List(Of Object)) In myDic
                    Dim L As New List(Of Object)
                    L = CType(KV.Value, List(Of Object))

                    ' refine the list, if it contains > 1 item, then delete any item with No Downloads "
                    While L.Count > 1 And L.FirstOrDefault(Function(x) x.Title.Contains("No Downloads")) IsNot Nothing
                        Dim itemToRemove = L.FirstOrDefault(Function(x) x.Title.Contains("No Downloads"))
                        If itemToRemove IsNot Nothing Then
                            L.Remove(itemToRemove)
                        End If
                    End While
                    'end of refine================================

                    For Each item As Item In L ' CheckedListBox1.Items

                        If Site <> getTitleWithoutPranthesesOrColon(item.Title) Then
                            Site = getTitleWithoutPranthesesOrColon(item.Title)


                            If FileIsEmpty = False Then
                                IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "</ul>" & vbCrLf)
                                IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "<br>" & vbCrLf)
                            End If

                            'Write Title
                            Dim ColorPhrase As String = ""
                            If InStr(item.Title.ToUpper, "No Downloads".ToUpper) Then
                                ColorPhrase = "Color:red"
                            End If

                            IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", TitlePhrase.Replace("@SITE", Site).Replace("@ColorPhrase", ColorPhrase) & vbCrLf)
                            IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", UnorderListPhrase & vbCrLf)

                            FileIsEmpty = False

                            If Color = "color:#FFFFFF" Then
                                Color = "color:#DEEAF6"
                            Else
                                Color = "#FFFFFF"
                            End If
                        End If

                        'IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "<li><div style="" background-color: " & Color & ";""><a href=""" & item.Link & """>" & item.Title & "</a></div></li>" & vbCrLf)
                        If InStr(item.Title.ToUpper, "No Downloads".ToUpper) Then

                        Else
                            IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "<li style=""margin: 12px 0""><a href=""" & item.Link & """  target=""_blank"" >" & item.Title & "</a></li>" & vbCrLf)
                        End If
                    Next
                Next
                IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "</ul>" & vbCrLf)
            Case Else
                Dim index As New List(Of String)
                'iterate throw checkboxlist to get the link and title ========================================
                For Each item As Item In CheckedListBox1.Items
                    index = item.indices

                    Dim L As New List(Of Object)
                    For Each SingleIndex In index
                        L.Clear()

                        If myDic.ContainsKey(SingleIndex) Then
                            L = CType(myDic(SingleIndex), Object)
                            myDic.Remove(SingleIndex)
                        End If

                        L.Add(item)
                        myDic.Add(SingleIndex, L)
                    Next
                Next
                'end of: iterate throw checkboxlist to get the link and title ===============================

                'Sort the dictionary ==========================================
                Dim L1 As New List(Of String)
                L1.AddRange(myDic.Keys)
                L1.Sort()
                Dim tempDic As New Dictionary(Of String, List(Of Object))
                For Each lcIndex In L1
                    tempDic.Add(lcIndex, myDic(lcIndex))
                Next
                myDic = tempDic
                'end of: Sort the dictionary ==================================

                'Build the required Titles Hierarchy=================
                Dim titlesDict As New Dictionary(Of String, String)
                titlesDict = GetHierarchyTitles(MyDic:=myDic, titlesDict:=titlesDict)
                'end of ====================================================

                '===========================================================
                Dim dt As New DataTable("Hierarchy")
                dt.Columns.Add("INDEX", GetType(String))
                dt.Columns.Add("TITLE", GetType(String))
                dt.Columns.Add("lcPARENT", GetType(String))

                For Each KV As KeyValuePair(Of String, String) In titlesDict
                    Dim Parent As String = ""
                    If KV.Key.Contains(".") Then
                        PArent = KV.Key.Substring(0, KV.Key.LastIndexOf("."c))
                    End If

                    dt.Rows.Add(KV.Key, KV.Value, Parent)
                    dt.AcceptChanges()
                Next
                DataGridView1.DataSource = dt.DefaultView
                'end of=========================================================


                IterateThroughTable(DT:=dt, CurrentRow:=Nothing, Titles:=myDic)
                ' MsgBox(Text1)
        End Select

        addLowerPartOfHTMLpage()
        Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "links.html")
        'WebBrowser1.Url = New Uri(System.AppDomain.CurrentDomain.BaseDirectory & "links.html")
    End Sub
    Dim Text1 As String = ""

    Private Sub IterateThroughTable(DT As DataTable, CurrentRow As DataRow, Titles As Dictionary(Of String, List(Of Object)))
        Dim index As String
        If CurrentRow Is Nothing Then
            index = ""
        Else
            index = CurrentRow("INDEX")
            ' Text1 = Text1 & vbCrLf & Space() & index & CurrentRow("TITLE")
        End If

        Dim DR() As DataRow
        DR = DT.Select("lcPARENT='" & index & "'")

        If DR.Count > 0 Then

            For Each singleDR As DataRow In DR
                IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "<ul>")
                IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "<li>" & singleDR("TITLE") & "</li>")

                If Titles.ContainsKey(singleDR("INDEX")) Then
                    IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "<ul>")

                    Dim Items As List(Of Object) = Titles(singleDR("INDEX"))
                    For Each lcItem In Items
                        IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "<li style=""margin: 14px 0""><a href=""" & lcItem.Link & """  target=""_blank"" >" & lcItem.Title & "</a></li>")
                    Next
                    IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "</ul>")
                End If
                IterateThroughTable(DT, singleDR, Titles)
            Next
        End If
        IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "</ul>")


    End Sub

    Function LoadIndexTitlesDictionary(filePath As String) As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ' Read all lines from the text file
        For Each line As String In File.ReadAllLines(filePath)
            Dim trimmedLine As String = line.Trim()
            If String.IsNullOrEmpty(trimmedLine) Then Continue For


            Dim idx As String = Split(trimmedLine, " ", 2)(0)
            idx = idx.Trim().TrimEnd("."c)

            Dim title As String = Split(trimmedLine, " ", 2)(1)

            ' Avoid duplicates, if necessary
            If Not dict.ContainsKey(idx) Then
                dict.Add(idx, title)
            End If
        Next

        Return dict
    End Function

    Function GetHierarchyTitles(ByVal MyDic As Dictionary(Of String, List(Of Object)),
                                titlesDict As Dictionary(Of String, String)) As Dictionary(Of String, String)


        'load all titles into MainDic ===================================
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ' Read all lines from the text file
        For Each line As String In File.ReadAllLines("C:\Users\2271\Downloads\EssaysClassifications_Updated.txt")
            Dim trimmedLine As String = line.Trim()
            If String.IsNullOrEmpty(trimmedLine) Then Continue For


            Dim idx As String = Split(trimmedLine, " ", 2)(0)
            idx = idx.Trim().TrimEnd("."c)

            Dim title As String = Split(trimmedLine, " ", 2)(1)

            ' Avoid duplicates, if necessary
            If Not dict.ContainsKey(idx) Then
                dict.Add(idx, title)
            End If
        Next
        'End of: load all titles into MainDic ===================================


        For Each Key As String In MyDic.Keys

            Key = TrimTrailingDot(Key)

            Dim parts = Key.Split("."c)
            Dim current As String = ""
            For i As Integer = 0 To parts.Length - 1
                ' Build the progressive index (5 → 5.3 → 5.3.1 → 5.3.1.2)
                current = If(current = "", parts(i), current & "." & parts(i))

                If titlesDict.ContainsKey(current) = False Then
                    titlesDict.Add(current, dict(current))
                End If
            Next
        Next
        Return titlesDict
    End Function

    Function TrimTrailingDot(index As String) As String
        If index IsNot Nothing AndAlso index.EndsWith(".") Then
            Return index.TrimEnd("."c)
        End If
        Return index
    End Function


    Private Sub addLowerPartOfHTMLpage()

        IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", "</body></html>")
    End Sub

    Private Sub addUpperPartOfHTMLpage()
        Dim html As String = IO.File.ReadAllText("Template.htm")
        html = html.Replace("@TITLE@", glTitle)
        IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory & "links.html", html)
    End Sub

    Function getTitleWithoutPranthesesOrColon(ByVal Title As String) As String


        Title = Title.Substring(0, -1 + InStr(Title, ":"))
        If InStr(Title, "(") > 0 Then
            Title = Title.Substring(0, -1 + InStr(Title, "("))
        End If

        getTitleWithoutPranthesesOrColon = Title

    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            DateTimePicker1.Enabled = True
        Else
            DateTimePicker1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.Enabled = True
        Else
            TextBox2.Enabled = False
        End If
    End Sub

    'Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
    '    If Not CheckedListBox1.SelectedItem Is Nothing Then
    '        If InStr(CType(CheckedListBox1.SelectedItem, Item).Link, "gallery/") > 0 Then

    '        Else
    '            Exit Sub
    '        End If
    '        PhotoDic.Clear()
    '            'Try
    '            Dim Folder As String = CType(CheckedListBox1.SelectedItem, Item).Link
    '            Folder = Folder.Replace(IO.Path.GetFileNameWithoutExtension(Folder), "")
    '            RemoveWhatIsBefore(Folder, "gallery/", False)
    '            Folder = replaceMonthByNum(Folder)
    '            Folder = System.AppDomain.CurrentDomain.BaseDirectory & Folder

    '        If Not IO.Directory.Exists(Folder) Then
    '            Exit Sub
    '        End If

    '        Dim inFile As New IO.StreamReader(Folder & "\Entries.txt")

    '            Dim oneLine As String = ""
    '            Dim oneLineArr() As String

    '            While Not inFile.Peek
    '                oneLine = inFile.ReadLine
    '                oneLineArr = Split(oneLine, vbTab)
    '                PhotoDic.Add(Folder & "\" & oneLineArr(0), oneLineArr(1))
    '            End While

    '            If PhotoDic.Count > 0 Then
    '                SetPictureTagToZero()
    '                SetImageToTag()
    '                ReflectChangesOnButtons()
    '            End If
    '        End If
    'End Sub

    'Sub SetPictureTagToZero()
    '    SetObjectValue(PictureBox1.Tag, 0)
    'End Sub

    'Sub IncreaseTageByOne()
    '    ChangeObjectValue(PictureBox1.Tag, 1)
    'End Sub
    'Sub DecreaseTageByOne()
    '    ChangeObjectValue(PictureBox1.Tag, -1)
    'End Sub

    'Sub SetImageToTag()
    '    ChangePictureBoxImage(PictureBox1.Image, PhotoDic.Keys(PictureBox1.Tag).ToString)
    'End Sub

    'Sub ChangePictureBoxImage(ByRef Picture As Image, ByVal ImagePath As String)
    '    Picture = Image.FromFile(ImagePath)
    'End Sub



    Sub ChangeObjectValue(ByRef obj As Object, ByVal Number As Integer)


        Number = CInt(obj) + Number

        SetObjectValue(obj, Number)
    End Sub

    Sub SetObjectValue(ByRef obj As Object, ByVal Number As Integer)
        obj = CObj(Number)
    End Sub

    'Sub ReflectChangesOnButtons()

    '    Form2.Close()
    '    If CInt(PictureBox1.Tag) = PhotoDic.Count - 1 Then
    '        Button9.Enabled = False
    '        Button8.Enabled = True
    '    ElseIf CInt(PictureBox1.Tag) = 0 Then
    '        Button9.Enabled = True
    '        Button8.Enabled = False
    '    Else
    '        Button9.Enabled = True
    '        Button8.Enabled = True
    '    End If
    'End Sub

    'Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
    '    IncreaseTageByOne()
    '    SetImageToTag()
    '    ReflectChangesOnButtons()
    'End Sub

    'Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
    '    DecreaseTageByOne()
    '    SetImageToTag()
    '    ReflectChangesOnButtons()
    'End Sub

    'Private Sub PictureBox1_Click(sender As Object, e As EventArgs)
    '    Try

    '    Catch ex As Exception

    '    End Try
    '    Form2.PictureBoxSize = PictureBox1.Size
    '    Form2.PictureBoxLocation = New Point(Me.Location.X + TabControl1.Location.X + PictureBox1.Location.X, Me.Location.Y + TabControl1.Location.Y + PictureBox1.Location.Y)


    '    Form2.Label1.Text = PhotoDic.Values(CInt(PictureBox1.Tag))
    '    If Form2.lcShown = False Then
    '        Form2.lcShown = True
    '        Form2.Show()
    '    Else
    '        Form2.lcShown = False
    '        Form2.Close()
    '    End If

    'End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim TN As New TreeNode
        TN.Name = TreeView1.SelectedNode.Name & "." & CStr(TreeView1.SelectedNode.Nodes.Count + 1)
        TN.Text = TextBox3.Text


        TreeView1.SelectedNode.Nodes.Add(TN)

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory & "\Tree.txt", "")
        For Each TR In TreeView1.Nodes
            StartListTree(TR)
        Next

    End Sub


    Sub StartListTree(ByVal TR As TreeNode)
        If IsNothing(TR.Parent) Then
            IO.File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory & "\Tree.txt", TR.Text & vbTab & TR.Name & vbTab & "" & vbCrLf)
        Else
            IO.File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory & "\Tree.txt", TR.Text & vbTab & TR.Name & vbTab & TR.Parent.Name & vbCrLf)
        End If

        For Each TR In TR.Nodes
            StartListTree(TR)
        Next




    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim inFile As New IO.StreamReader(AppDomain.CurrentDomain.BaseDirectory & "\Tree.txt")
        Dim oneLine As String
        Dim oneLineArr() As String

        While Not inFile.Peek
            oneLine = inFile.ReadLine
            oneLineArr = Split(oneLine, vbTab)

            Dim TR As New TreeNode
            TR.Text = oneLineArr(0)
            TR.Name = oneLineArr(1)

            If oneLineArr(2) = "" Then
                TreeView1.Nodes.Add(TR)
            Else
                Dim TR1 As New TreeNode
                TR1 = TreeView1.Nodes.Find(oneLineArr(2), True)(0)
                TR1.Nodes.Add(TR)
            End If
        End While

        inFile.Close()
        inFile.Dispose()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim AllText As String = TextBox1.Text
        Dim arrText() As String = AllText.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
        Dim tempArr() As String
        Dim headers() As String = arrText(0).Split(vbTab)
        Dim oneRecord As String = ""
        Dim AllRecords As String = ""

        AllRecords = AllRecords + "{" & vbCrLf
        AllRecords = AllRecords + "  ""/Countries"": {" & vbCrLf



        For I As Integer = 1 To arrText.Length - 1
            oneRecord = ""
            tempArr = arrText(I).Split(vbTab)

            oneRecord = "    """ & tempArr(0) & """: {" & vbCrLf
            For J = 0 To tempArr.Length - 1
                If Not (headers(J) = "Cities" Or headers(J) = "BelongsTo") Then
                    oneRecord = oneRecord & "      """ & headers(J) & """: " & """" & tempArr(J) & """," & vbCrLf
                Else
                    oneRecord = oneRecord & "      """ & headers(J) & """: " & returnITArray(tempArr(J)) & "," & vbCrLf
                End If

            Next
            oneRecord = oneRecord & "      ""subCollection"": {}" & vbCrLf
            oneRecord = oneRecord & "    }"

            If I = 1 Then
                AllRecords = AllRecords & oneRecord
            Else
                AllRecords = AllRecords & "," & vbCrLf & oneRecord
            End If

        Next

        AllRecords = AllRecords & vbCrLf & "  }"
        AllRecords = AllRecords & vbCrLf & "}"

        TextBox1.Text = AllRecords

        IO.File.WriteAllText("\\HQPROFILES\HQUP$\2271\Desktop\all_countries.json", AllRecords)



    End Sub


    Function returnITArray(ByVal componenes As String) As String
        Dim tempArr As String() = Split(componenes, "|")

        For I As Integer = 0 To tempArr.Length - 1
            tempArr(I) = """" & Trim(tempArr(I)) & """"
        Next
        returnITArray = "[" & Join(tempArr, ",") & "]"

    End Function

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        ListGuardianIran(FileName:="GuaridianIran.txt")

        'DownlaodGuaridianOpinion()

        'DownloadTheConversation()
        'ListTheConversation()

    End Sub

    Private Sub DownlaodGuaridianIran()
        Dim URL As String = "https://www.theguardian.com/world/iran"
        Label1.Text = "Downloading guardian Opinion .."
        Application.DoEvents()
        DownLoadFile(URL, System.AppDomain.CurrentDomain.BaseDirectory & "/GuaridianIran.txt")
        ListGuardianIran(FileName:="GuaridianIran")
    End Sub

    Private Sub ListGuardianIran(ByVal FileName As String)
        Dim Files As Integer = 0
        FileName = FileName.Replace(".txt.txt", ".txt")

        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName)
        Dim allcontentArr As String() = Split(allContent, "<li class=""")

        Dim Link As String = ""
        Dim Title As String
        Dim Auther As String
        RemoveFirstCell(allcontentArr)
        For Each entry As String In allcontentArr
            Link = entry
            Title = entry

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.theguardian.com/" & Link

            RemoveWhatIsBefore(Title, "headline-text", False)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)

            refineTitle(Title)

            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="Guardian Iran" & ": " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Guardian Iran" & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If


        'Dim L As New List(Of String)
        'L = allcontentArr.ToList
        'L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        'L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
        'allcontentArr = L.ToArray

        'RemoveFirstCell(allcontentArr)
        'L.Clear()

        'Dim Link As String = ""
        'Dim Title As String
        'Dim Auther As String

        'For Each entry As String In allcontentArr
        '    Link = entry
        '    Title = entry
        '    Auther = entry


        '    RemoveWhatIsBefore(Link, "<a href=""", False)
        '    RemoveWhatIsAfter(Link, """", False)


        '    Link = "https://www.theguardian.com/" & Link

        '    RemoveWhatIsBefore(Title, "<h", False)
        '    RemoveWhatIsBefore(Title, ">", False)
        '    RemoveWhatIsAfter(Title, "</h", False)

        '    Dim Tags As String()
        '    Tags = Split(Title, "<")

        '    For I As Integer = 0 To Tags.Count - 1
        '        Try

        '            Tags(I) = Tags(I).Substring(Tags(I).IndexOf(">"))
        '            Tags(I) = Tags(I).Replace(">", "")

        '            L.Clear()
        '            L = Tags.ToList
        '            L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        '            L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
        '            Title = Join(L.ToArray)
        '        Catch ex As Exception

        '        End Try

        '    Next

        '    If Auther.Contains("<span class=""dcr-1tmbud3") Or Auther.Contains("<span class=""dcr-2n3kt1"">") Then
        '        RemoveWhatIsBefore(Auther, "<span class=""dcr-1tmbud3"">", False)
        '        RemoveWhatIsBefore(Auther, "<span class=""dcr-2n3kt1"">", False)
        '        RemoveWhatIsAfter(Auther, "</span>", False)
        '        Auther = "(" & Auther & ") "
        '    Else
        '        Auther = ""
        '    End If

        '    Title = Auther & Title
        '    refineTitle(Title)

        '    If Trim(Title) <> "" Then
        '        If AddLinktoCheckedListBox1(Title:="Guardian Iran" & ": " & Title, Link:=Link) Then Files = Files + 1
        '    End If
        'Next

        'If Files = 0 Then
        '    AddLinktoCheckedListBox1(Title:="Guardian Iran" & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        'End If
    End Sub


    Private Sub DownlaodGuaridianOpinion()
        Dim URL As String = "https://www.theguardian.com/uk/commentisfree"



        Label1.Text = "Downloading guardian Opinion .."

        Application.DoEvents()
        DownLoadFile(URL, System.AppDomain.CurrentDomain.BaseDirectory & "/Guaridian.txt")

        ListGuardianOpinion(FileName:="Guaridian")

        'Process.Start("https://www.theguardian.com/commentisfree?page=" & GuardianIndex)

        'Process.Start("D:\VB Projects\Downloads\Downloads\bin\Debug\Guaridian.txt")


        'Timer1.Start()

    End Sub

    Private Sub ListGuardianOpinion(ByVal FileName As String)
        Dim Files As Integer = 0

        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName)
        Dim allcontentArr As String() = Split(allContent, "<ul class=""dcr-")

        Dim L As New List(Of String)
        L = allcontentArr.ToList
        L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
        allcontentArr = L.ToArray

        RemoveFirstCell(allcontentArr)
        L.Clear()

        For Each Str As String In allcontentArr
            L.AddRange(Split(Str, "<li ").ToList)
        Next
        L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))

        For I As Integer = 0 To L.Count - 1
            If L(I).Contains("<a href=""/commentisfree/") Then
            Else
                L(I) = ""
            End If
        Next
        L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
        allcontentArr = L.ToArray



        Dim Link As String = ""
        Dim Title As String
        Dim Auther As String

        For Each entry As String In allcontentArr
            Link = entry
            Title = entry
            Auther = entry


            RemoveWhatIsBefore(Link, "<a href=""/commentisfree/", False)
            RemoveWhatIsAfter(Link, """", False)


            Link = "https://www.theguardian.com/commentisfree/" & Link

            RemoveWhatIsBefore(Title, "<h", False)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "</h", False)

            Dim Tags As String()
            Tags = Split(Title, "<")

            For I As Integer = 0 To Tags.Count - 1
                Try

                    Tags(I) = Tags(I).Substring(Tags(I).IndexOf(">"))
                    Tags(I) = Tags(I).Replace(">", "")

                    L.Clear()
                    L = Tags.ToList
                    L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
                    L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
                    Title = Join(L.ToArray)
                Catch ex As Exception

                End Try

            Next



            'Extract Auther 
            '"</h3><span class=""dcr-"
            If Auther.Contains("<span class=""dcr-1tmbud3") Or Auther.Contains("<span class=""dcr-2n3kt1"">") Then
                RemoveWhatIsBefore(Auther, "<span class=""dcr-1tmbud3"">", False)
                RemoveWhatIsBefore(Auther, "<span class=""dcr-2n3kt1"">", False)
                RemoveWhatIsAfter(Auther, "</span>", False)
                Auther = "(" & Auther & ") "
            Else
                Auther = ""
            End If

            Title = Auther & Title
            refineTitle(Title)

            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="Guardian Opinion" & ": " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Guardian Opinion" & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub



    Private Sub DownlaodBBCScience()
        Dim URL As String = "https://www.bbc.com/news/science_and_environment"
        Label1.Text = "Downloading BBC Science .."
        Application.DoEvents()
        DownLoadFile(URL,
            System.AppDomain.CurrentDomain.BaseDirectory & "/BBC_Science.txt")

        ListBBCScience(FileName:="BBC_Science")
    End Sub

    Private Sub ListBBCScience(FileName As String)
        Dim Files As Integer = 0

        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName & ".txt")

        Dim allContentArr() As String = Split(allContent, "{\""url\"":\""")

        For I As Integer = LBound(allContentArr) To UBound(allContentArr)
            If Not allContentArr(I).Contains("/news/articles/") Then
                allContentArr(I) = ""
            End If
        Next

        Dim L As New List(Of String)
        L = allContentArr.ToList
        L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
        allContentArr = L.ToArray

        RemoveFirstCell(allContentArr)

        Dim Link As String
        Dim Title As String

        For Each entry As String In allContentArr
            Link = entry
            Title = entry

            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.bbc.com/" & "/" & Link

            RemoveWhatIsBefore(Title, "\""headline\"":\""", False)
            RemoveWhatIsAfter(Title, "\""", False)
            refineTitle(Title)

            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:=FileName.Replace("_", " ") & ": " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:=FileName.Replace("_", " ") & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub DownloadDogoNewsScience()
        Dim URL As String = "https://www.dogonews.com/page/1"
        Label1.Text = "Downloading Dogo .."
        Application.DoEvents()
        DownLoadFile(URL,
            System.AppDomain.CurrentDomain.BaseDirectory & "/Dogo.txt")

        ListDogo(FileName:="Dogo")
    End Sub

    Sub ListDogo(ByVal FileName As String)
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName & ".txt")

        Dim allContentArr() As String = Split(allContent, "<a class=""flex article-name")

        RemoveFirstCell(allContentArr)

        Dim Link As String
        Dim Title As String

        For i As Integer = 0 To allContentArr.Length - 1
            Link = allContentArr(i)
            Title = allContentArr(i)

            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.dogonews.com/" & "/" & Link

            RemoveWhatIsBefore(Title, "lang=""en"">", False)
            RemoveWhatIsAfter(Title, "<", False)
            refineTitle(Title)

            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:=FileName.Replace("_", " ") & ": " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:=FileName.Replace("_", " ") & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub
    Sub DownLoadBBCInDepth()
        Dim URL As String = "https://www.bbc.com/news/bbcindepth"
        Label1.Text = "Downloading BBC In Depth..."
        Application.DoEvents()
        DownLoadFile(URL,
        System.AppDomain.CurrentDomain.BaseDirectory & "/BBCIndEpth.txt")

        ListBBCInDepth(FileName:="BBCIndEpth")
    End Sub


    Sub DownLoadCNNScience()
        Dim CNNLinks() As String = {"https://edition.cnn.com/science",
            "https://edition.cnn.com/science/space",
            "https://edition.cnn.com/science/life",
            "https://edition.cnn.com/science/unearthed"}
        Dim FileName As String
        For I As Integer = 0 To CNNLinks.Count - 1
            FileName = CNNLinks(I)
            RemoveWhatIsBeforeLast(FileName, "/", False)

            Label1.Text = "Downloading CNN " & FileName & ".."
            Application.DoEvents()
            DownLoadFile(CNNLinks(I),
            System.AppDomain.CurrentDomain.BaseDirectory & "/CNN_" & FileName & ".txt")

            ListCNNScience("CNN_" & FileName)


        Next
    End Sub

    Sub ListCNNScience(ByVal FileName As String)
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName & ".txt")

        Dim allContentArr() As String = Split(allContent, "<a href=""")

        For i As Integer = 0 To allContentArr.Length - 1
            If allContentArr(i) Like "/2024/*/*/*" And allContentArr(i) Like "*data-editable=""headline"">*" Then

            Else
                allContentArr(i) = ""
            End If
        Next

        Dim L As New List(Of String)
        L = allContentArr.ToList
        L.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        L.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))
        allContentArr = L.ToArray

        Dim Link As String
        Dim Title As String

        For i As Integer = 0 To allContentArr.Length - 1
            Link = allContentArr(i)
            Title = allContentArr(i)

            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.cnn.com" & Link
            RemoveWhatIsBefore(Title, "data-editable=""headline"">", False)
            RemoveWhatIsAfter(Title, "<", False)

            refineTitle(Title)

            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:=FileName.Replace("_", " ") & ": " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:=FileName.Replace("_", " ") & ": No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub


    Private Sub ListBBCInDepth(FileName As String)
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/" & FileName & ".txt")
        'Dim Components() As String = Split(allContent, "<div class=""Grid-block  Grid-is2to3-xs_is1to1 "">")
        'RemoveFirstCell(Components)

        Dim allContentArr() As String = Split(allContent, "{""title"":""")
        RemoveFirstCell(allContentArr)
        Dim Link As String
        Dim Title As String
        For i As Integer = 0 To allContentArr.Length - 1
            Link = allContentArr(i)
            Title = allContentArr(i)

            RemoveWhatIsBefore(Link, "href"":""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.bbc.com" & Link

            RemoveWhatIsAfter(Title, """", False)
            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="BBC InDepth: " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="BBC InDepth: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If


    End Sub

    Sub DownloadCostOfEveryThing()
        Label1.Text = "Downloading Cost Of EveryThing.."
        Application.DoEvents()
        DownLoadFile("https://www.rt.com/shows/cost-of-everything/",
            System.AppDomain.CurrentDomain.BaseDirectory & "/CostOfEverything.txt")

        ListCostOfEveryThing()

    End Sub

    Private Sub ListCostOfEveryThing()
        Dim Files As Integer = 0
        Dim allContent As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "/CostOfEverything.txt")
        'Dim Components() As String = Split(allContent, "<div class=""Grid-block  Grid-is2to3-xs_is1to1 "">")
        'RemoveFirstCell(Components)


        Dim allContentArr() As String = Split(allContent, "<a class=""link link_hover"" href=""/shows/cost-of-everything/")
        RemoveFirstCell(allContentArr)

        For Each Str As String In allContentArr

            Dim Link As String = Str
            Dim Title As String = Str


            Link = Str
            'RemoveWhatIsBefore(Link, " href=""", False)
            RemoveWhatIsAfter(Link, """>", False)
            Link = "https://www.rt.com/shows/cost-of-everything/" + Link

            Title = Str
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "</a>", False)
            refineTitle(Title)

            If Trim(Title) <> "" Then
                If AddLinktoCheckedListBox1(Title:="RT-Cost Of Everything: " & Title, Link:=Link) Then Files = Files + 1
            End If
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="RT-Cost Of Everything: No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Private Sub DownLoadKateHone()
        DownlaodArticles()
        DownlaodAnalytics()
        DownlaodTrend()

        'https://katehon.com/en/analitics/trendy?page=1
        'https://katehon.com/en/articles?page=0
        ''https://katehon.com/en/analytics?page=0


    End Sub

    Sub DownlaodArticles()
        DownLoadKateHone(Section:="articles?page=0", fileName:="Articles")
    End Sub

    Sub DownlaodAnalytics()
        DownLoadKateHone(Section:="analytics?page=0", fileName:="Analytics")
    End Sub

    Sub DownlaodTrend()
        DownLoadKateHone(Section:="analitics/trendy?page=0", fileName:="trendy")
    End Sub
    Private Sub DownLoadKateHone(ByVal Section As String, ByVal fileName As String)
        Label1.Text = "Downloading Kathone..."
        Application.DoEvents()

        DownLoadFile("https://katehon.com/en/" & Section,
        System.AppDomain.CurrentDomain.BaseDirectory & "/Katehone" & fileName & ".txt")


        ListKatehone(fileName)
        Application.DoEvents()

        Exit Sub






    End Sub

    Private Sub ListKatehone(ByVal FileName As String)
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)


        allContentArr = Split(allContent, "<h2 class=""title"">")

        RemoveFirstCell(allContentArr)
        Dim Link As String = ""
        Dim Title As String = ""
        For I As Integer = 0 To allContentArr.Length - 1
            RemoveWhatIsAfter(allContentArr(I), "</a></h2>", False)
            Link = allContentArr(I)
            RemoveWhatIsBefore(Link, """", False)
            RemoveWhatIsAfter(Link, """", False)


            Title = allContentArr(I)
            RemoveWhatIsBefore(Title, ">", False)
            RemoveWhatIsAfter(Title, "<", False)

            If AddLinktoCheckedListBox1(Title:="Katehon" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Katehon : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If

    End Sub

    'Sub DownLoadDiv()
    '    Dim DevFile As New IO.FileInfo(System.AppDomain.CurrentDomain.BaseDirectory & "\DevLatests.txt")
    '    DevLastTime = DevFile.LastWriteTime
    '    Timer1.Start()
    '    Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "\DevLatests.txt")
    'End Sub

    Sub ListDev()
        Dim Files As Integer
        Dim allContent As String
        Dim allContentArr() As String

        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\DevLatests.txt")
        Dim Pattern As String = "<h2 class=""crayons-story__title"">"
        allContentArr = Regex.Split(allContent, Pattern)
        RemoveFirstCell(allContentArr)
        Dim Link As String = ""
        Dim Title As String = ""
        For I As Integer = 0 To allContentArr.Length - 1

            RemoveWhatIsBefore(allContentArr(I), "<a href=""", False)
            RemoveWhatIsAfter(allContentArr(I), "</a>", False)

            Link = allContentArr(I)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://dev.to//" & Link

            Title = allContentArr(I)
            RemoveWhatIsBeforeLast(Title, ">", False)
            refineTitle(Title)

            If AddLinktoCheckedListBox1(Title:="Div" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next


        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Div : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If



    End Sub



    Sub downloadRoundTheCode()
        Label1.Text = "Downloading RoundTheCode..."
        Application.DoEvents()

        DownLoadFile("https://www.roundthecode.com/",
        System.AppDomain.CurrentDomain.BaseDirectory & "/RoundTheCode.txt")

        ListRoundTheCode()
        Application.DoEvents()
    End Sub

    Sub ListRoundTheCode()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\RoundTheCode.txt")
        Dim Pattern As String = "<div class=""article-listing-item"">|home-code-example-listing-container"">|<div class=""video-item-container"">"
        '
        'allContentArr = Regex.Split(allContent, Pattern)
        allContentArr = Split(allContent, "<div class=""content"">")
        RemoveFirstCell(allContentArr)

        Dim Link As String = ""
        Dim Title As String = ""
        Dim Text As String = ""
        For I As Integer = 0 To allContentArr.Length - 2

            Title = allContentArr(I)
            Link = allContentArr(I)


            RemoveWhatIsBefore(Link, "href=""", False)
            RemoveWhatIsAfter(Link, """", False)
            Link = "https://www.roundthecode.com/" & Link

            RemoveWhatIsBefore(Title, "<h3>", False)
            RemoveWhatIsAfter(Title, "</h3>", False)

            If AddLinktoCheckedListBox1(Title:="Round The Code" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next


        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Round The Code : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub


    Sub downloadFreecodecamp()
        Label1.Text = "Downloading freecodecamp..."
        Application.DoEvents()

        DownLoadFile("https://www.freecodecamp.org/news/",
        System.AppDomain.CurrentDomain.BaseDirectory & "/freecodecamp.txt")

        ListFreecodecamp()
        Application.DoEvents()
    End Sub

    Sub ListFreecodecamp()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\freecodecamp.txt")
        Dim Pattern As String = "<h2 class=""post-card-title"">"
        '
        allContentArr = Regex.Split(allContent, Pattern)
        RemoveFirstCell(allContentArr)

        Dim Link As String = ""
        Dim Title As String = ""
        Dim Text As String = ""
        For I As Integer = 0 To allContentArr.Length - 1
            RemoveWhatIsBefore(allContentArr(I), "<a href=""", False)
            RemoveWhatIsAfter(allContentArr(I), "</a", False)

            Link = allContentArr(I)
            Title = allContentArr(I)

            RemoveWhatIsAfter(Link, """>", False)
            RemoveWhatIsBefore(Title, """>", False)
            refineTitle(Title)
            Link = "https://www.freecodecamp.org" & Link
            If AddLinktoCheckedListBox1(Title:="Free Code Camp" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next


        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Free Code Camp : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub


    Sub downloadWM()
        Label1.Text = "Downloading WM..."
        Application.DoEvents()

        DownLoadFile("https://wn.com/technology/",
        System.AppDomain.CurrentDomain.BaseDirectory & "/WM.txt")

        ListWM()
        Application.DoEvents()
    End Sub


    Sub ListWM()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\WM.txt")
        Dim Pattern As String = "<div class=""nl-item section"">|<div class=""tvl-item"">"
        '
        allContentArr = Regex.Split(allContent, Pattern)
        RemoveFirstCell(allContentArr)

        Dim Link As String = ""
        Dim Title As String = ""
        Dim Text As String = ""
        For I As Integer = 0 To allContentArr.Length - 1
            RemoveWhatIsBefore(allContentArr(I), "<h", False)
            RemoveWhatIsAfter(allContentArr(I), "</h", False)

            Link = allContentArr(I)
            Title = allContentArr(I)

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)


            Title = allContentArr(I)
            RemoveWhatIsBefore(Title, Link, False)
            Dim S As String = "<span dir=""ltr"" lang=""en"">"
            RemoveWhatIsBefore(Title, S, False)
            RemoveWhatIsAfter(Title, "<", False)
            Text = Text & vbCrLf & Title

            If AddLinktoCheckedListBox1(Title:="WN" & ": " & Title, Link:=Link) Then Files = Files + 1
        Next


        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="WN : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Sub DownlaodNature()
        '
        Label1.Text = "Downloading Nature..."
        Application.DoEvents()

        DownLoadFile("https://www.nature.com/nature/articles?searchType=journalSearch&sort=PubDate&page=1",
        System.AppDomain.CurrentDomain.BaseDirectory & "/Nature.txt")

        ListNature()
        Application.DoEvents()
    End Sub


    Sub ListNature()
        Dim Files As Integer = 0
        Dim allContent As String
        Dim allContentArr() As String
        allContent = IO.File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory & "\Nature.txt")

        allContentArr = Split(allContent, "<h3 class=""c-card__title")
        RemoveFirstCell(allContentArr)

        Dim Link As String = ""
        Dim Title As String = ""
        For I As Integer = 0 To allContentArr.Length - 1
            Link = allContentArr(I)

            RemoveWhatIsBefore(Link, "<a href=""", False)
            RemoveWhatIsAfter(Link, """", False)

            Title = allContentArr(I)
            RemoveWhatIsAfter(Title, "</a>", False)
            RemoveWhatIsBeforeLast(Title, ">", False)

            If AddLinktoCheckedListBox1(Title:="Nature" & ": " & Title, Link:="https://www.nature.com/" & Link) Then Files = Files + 1
        Next

        If Files = 0 Then
            AddLinktoCheckedListBox1(Title:="Nature : No Downloads", Link:=Guid.NewGuid.ToString.Replace("-", ""))
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim Text As String = TextBox4.Text
        Dim arrText As New List(Of String)
        arrText = Split(Text, "<div class=""noteText"">").ToList

        arrText = arrText.Select(Of String)(Function(X) X.Replace(vbCrLf, "")).ToList

        arrText = arrText.Select(Of String)(Function(X) Strings.Left(X, InStr(X, "<") - 1)).ToList

        arrText.RemoveAll(Function(str As String) String.IsNullOrWhiteSpace(str))
        arrText.RemoveAll(Function(str As String) String.IsNullOrEmpty(str))


        Clipboard.SetText(Join(arrText.ToArray, vbCrLf))
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim ofd As New OpenFileDialog()
        ' Set dialog properties
        ofd.Title = "Select a Text File"
        ofd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        ofd.InitialDirectory = "D:\VB Projects\Downloads\bin\Debug\File1000LinesEach\" 'Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        ofd.Multiselect = False

        ' Show dialog
        Dim L As New List(Of String)

        If ofd.ShowDialog() = DialogResult.OK Then
            Using reader As New StreamReader(ofd.FileName)
                Label2.Text = ofd.FileName
                Dim line As String
                While Not reader.Peek
                    ' Split the line by vbTab into an array
                    line = reader.ReadLine

                    If line.Contains("No Downloads") Then Continue While

                    Dim fields() As String = line.Split(New String() {vbTab}, StringSplitOptions.None)
                    line = fields(fields.Length - 1)
                    L.Add(Mid(line, InStr(line, ":") + 1))
                End While
            End Using
        Else
            Exit Sub
        End If

        Dim CHATGPTCommand As String = "using the classifications in the attached file, classify and categorize the following titles into their best fit according to the full hierarchy. "
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " only show the category index prefix (like 1.1.2.3 Title...) without listing or showing any category names. "
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " omit empty categories. "
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " each line should start with its full classification index followed by the title. "
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " make sure no title is skipped. "
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " you can assign a title to more than one category if applicable. "
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " if a title doesn’t fit any category, put it under ""7. UNCATEGORIZED / MISCELLANEOUS ""."
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " Preserve titles exactly as written "
        CHATGPTCommand = CHATGPTCommand & vbCrLf & " make the result into a downloadable txt file named '" & IO.Path.GetFileNameWithoutExtension(Label2.Text) & "_INDEXED.txt" & "' " & vbCrLf


        Clipboard.SetText(CHATGPTCommand & vbCrLf & Join(L.ToArray, vbCrLf))

        MsgBox("(" & L.Count & ") Title are in Clipbpard...")
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        'get the classified titles ===============
        Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Dim ofd As New OpenFileDialog()
        ' Set dialog properties
        ofd.Title = "Select a Text File"
        ofd.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        ofd.Multiselect = False

        ' Show dialog
        If ofd.ShowDialog() = DialogResult.OK Then
            dict = BringClassifiedTitles(ofd.FileName)
        Else
            Exit Sub
        End If
        'End of: get the classified titles ===============


        Dim inFile As New IO.StreamReader(Label2.Text)

        Dim outFileName As String = IO.Path.GetDirectoryName(Label2.Text) & "\" & IO.Path.GetFileNameWithoutExtension(Label2.Text) & "_INDEXED" & IO.Path.GetExtension(Label2.Text)

        Dim outFile As New IO.StreamWriter(outFileName) 'Label2.Text)
        Dim oneLine As String = ""
        Dim arrOneLine() As String = Nothing
        Dim TITLE As String = ""
        Dim index As String = ""
        While Not inFile.Peek
            index = "8"

            oneLine = inFile.ReadLine
            arrOneLine = Split(oneLine, vbTab)
            TITLE = arrOneLine(2).Trim()
            TITLE = Mid(TITLE, InStr(TITLE, ":") + 1).Trim()
            Try
                index = dict(TITLE)
            Catch ex As Exception

            End Try
            outFile.WriteLine(oneLine & vbTab & index)

        End While



        inFile.Close()
        outFile.Flush()
        outFile.Close()


        IO.File.Move(Label2.Text, IO.Path.GetDirectoryName(Label2.Text) & "\Done\" & IO.Path.GetFileName(Label2.Text))


        MsgBox("Done")

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ListGlobalSearchScience()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        MsgBox(TreeView2.SelectedNode.Tag)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'DownloadAndListEconomist_OpenFuture()
        'DownloadEconomist()
        'Download_FT()
        GetUser()

    End Sub

    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (
    ByVal lpBuffer As String, ByRef nSize As Integer) As Integer


    Sub GetUser()
        Dim buffer As String = New String(CChar(" "), 25)
        Dim retVal As Integer = GetUserName(buffer, 25)
        Dim userName As String = Strings.Left(buffer, InStr(buffer, Chr(0)) - 1)
        MsgBox(userName)
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        MsgBox(WebBrowser1.DocumentText)
    End Sub

    Private Sub Button6_ClientSizeChanged(sender As Object, e As EventArgs) Handles Button6.ClientSizeChanged

    End Sub




    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If index <> previous Then
            previous = index
            pURL = Process.Start(L(index).URL)
        End If

        If pFile Is Nothing Then
            pFile = Process.Start(System.AppDomain.CurrentDomain.BaseDirectory & "\" & L(index).FileName)
        End If



        Dim lcFileInfo As New IO.FileInfo(System.AppDomain.CurrentDomain.BaseDirectory & "\" & L(index).FileName)
        Dim lcLastModified As Date = lcFileInfo.LastWriteTime
        If L(index).lastModified < lcLastModified Then
            Select Case L(index).URL
                Case "https://www.dev.to/latest"
                    ListDev()
                Case "https://www.theguardian.com/science"
                    ListGuardianScience()
                Case "https://www.reuters.com/science/"
                    ListReuters()
                Case "https://www.forbes.com/science/"
                    ListForbes()
                Case "https://katehon.com/en/analitics/trendy?page=1", "https://katehon.com/en/articles?page=0", "https://katehon.com/en/analytics?page=0"
                    ListKatehone(L(index).FileName)
                Case "https://www.theguardian.com/uk/commentisfree?page=1"
                    ListGuardianOpinion(L(index).FileName)
                Case "https://www.theguardian.com/world/iran"
                    ListGuardianIran(L(index).FileName)

            End Select
            pURL = Nothing
            pFile = Nothing
            index = index + 1

            'If index = L.Count Then
            '    Dim A As String = ""
            'End If
        End If

        If index = L.Count Then
            Timer1.Stop()
            Exit Sub
        End If

    End Sub



    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Exit Sub

        ' ===== CONFIGURATION =====
        Dim inputFile As String = "D:\VB Projects\Downloads\bin\Debug\CorrectedDone.txt"
        Dim outputFolder As String = "D:\VB Projects\Downloads\bin\Debug\File1000LinesEach\"
        Dim linesPerFile As Integer = 1000
        ' ==========================
        Try
            Kill("D:\VB Projects\Downloads\bin\Debug\File1000LinesEach\*.txt")
        Catch ex As Exception

        End Try


        If Not Directory.Exists(outputFolder) Then
            Directory.CreateDirectory(outputFolder)
        End If

        Dim fileCounter As Integer = 1
        Dim lineCounter As Integer = 0
        Dim writer As StreamWriter = Nothing

        Dim previousDate As String = ""

        Dim line As String
        Dim arrline() As String
        Dim outputFile As String = ""
        Dim Lines As String = ""
        Try
            Using reader As New StreamReader(inputFile)
                While Not reader.Peek
                    line = reader.ReadLine()
                    arrline = Split(line, vbTab)


                    'Try
                    '    Dim A = arrline(1).Trim
                    'Catch ex As Exception
                    '    Lines = Lines & vbCrLf & line
                    'End Try

                    If previousDate.Trim <> arrline(1).Trim Then

                        If writer IsNot Nothing Then
                            writer.Close()
                            writer.Dispose()

                            previousDate = arrline(1).Trim

                            outputFile = Path.Combine(outputFolder, $"{arrline(1)}.txt")
                            writer = New StreamWriter(outputFile, False)
                        Else
                            outputFile = Path.Combine(outputFolder, $"{arrline(1)}.txt")
                            writer = New StreamWriter(outputFile, False)
                        End If


                    End If

                    writer.WriteLine(line)

                End While
            End Using

        Finally

        End Try
        ' Clipboard.SetText(Lines)
        MsgBox("done")

    End Sub

    Private Function MakeNode(text As String, children() As String) As TreeNode
        Dim root As New TreeNode(text)
        For Each c In children
            root.Nodes.Add(New TreeNode(c))
        Next
        Return root
    End Function

    ' Search as you type (you can also trigger on Button click)
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Dim term As String = TextBox5.Text.Trim()
        FilterTree(TreeView2, term)
    End Sub

    Private Sub FilterTree(tv As TreeView, searchText As String)
        tv.BeginUpdate()
        tv.Nodes.Clear()

        If String.IsNullOrEmpty(searchText) Then
            ' Restore full tree
            For Each n As TreeNode In OriginalNodes
                tv.Nodes.Add(CType(n.Clone(), TreeNode))
            Next
            tv.EndUpdate()
            tv.ExpandAll()
            Return
        End If

        ' Filtered view: keep only matching nodes (and ancestors)
        Dim filtered As New List(Of TreeNode)
        For Each root In OriginalNodes
            Dim match As TreeNode = FilterNode(root, searchText)
            If match IsNot Nothing Then
                filtered.Add(match)
            End If
        Next

        tv.Nodes.AddRange(filtered.ToArray())
        tv.ExpandAll()
        tv.EndUpdate()
    End Sub

    ' Recursive filter function — returns a copy of node if it or children match
    Private Function FilterNode(node As TreeNode, term As String) As TreeNode
        Dim copy As TreeNode = CType(node.Clone(), TreeNode)
        copy.Nodes.Clear()
        Dim matched As Boolean = node.Text.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0

        For Each child As TreeNode In node.Nodes
            Dim filteredChild As TreeNode = FilterNode(child, term)
            If filteredChild IsNot Nothing Then
                copy.Nodes.Add(filteredChild)
                matched = True
            End If
        Next

        Return If(matched, copy, Nothing)
    End Function

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        frmMonaqqeb.Show()
    End Sub
End Class

Class Item
    Sub New(ByVal Title As String, ByVal Link As String)
        _Title = Title
        _Link = Link
    End Sub

    Sub New(ByVal Title As String, ByVal Link As String, ByVal Indexes As String)
        _Title = Title
        _Link = Link
        Me.indexes = Indexes
    End Sub

    Private _Title As String
    Public Property Title As String
        Get
            Return _Title
        End Get
        Set(value As String)
            _Title = value
        End Set
    End Property


    Private _Link As String
    Public Property Link As String
        Get
            Return _Link
        End Get
        Set(value As String)
            _Link = value
        End Set
    End Property

    Private _indexes As String
    Public WriteOnly Property indexes As String
        Set(value As String)
            _indexes = value
        End Set
    End Property

    Public ReadOnly Property indices As List(Of String)
        Get
            Return Split(_indexes, ",").ToList
        End Get
    End Property


    Public Overrides Function ToString() As String 'this is the heart of the mission
        Return _Title
    End Function
End Class

Public Class links
    Property URL As String
    Property FileName As String
    Property lastModified As Date


    Sub New(ByVal URL As String, ByVal FileName As String)
        Me.URL = URL
        Me.FileName = FileName

        Dim DevFile As New IO.FileInfo(System.AppDomain.CurrentDomain.BaseDirectory & "\" & FileName)
        Me.lastModified = DevFile.LastWriteTime


    End Sub
End Class



Public Class CategoriesAndLinks
    Property Categories As List(Of String)
    Property Link As String
End Class