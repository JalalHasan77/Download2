Imports System.Text.RegularExpressions
Imports System.Net
Imports System.IO
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports VBIDE = Microsoft.Vbe.Interop


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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' ====== EDIT THESE PATHS ======
        Dim folderPath As String = "C:\Your\Folder\With\Xlsx"            ' e.g., C:\Data\Invoices
        Dim moduleBasPath As String = "C:\Your\Folder\Module1.bas"       ' e.g., C:\Data\Module1.bas
        Dim arabicMapPath As String = "C:\Your\Folder\ArabicMap.txt"     ' e.g., C:\Data\ArabicMap.txt
        ' ==============================

        If Not Directory.Exists(folderPath) Then
            Console.WriteLine("Folder not found: " & folderPath)
            Return
        End If
        If Not File.Exists(moduleBasPath) Then
            Console.WriteLine("Module .bas not found: " & moduleBasPath)
            Return
        End If
        If Not File.Exists(arabicMapPath) Then
            Console.WriteLine("ArabicMap .txt not found: " & arabicMapPath)
            Return
        End If

        Dim app As Excel.Application = Nothing
        Try
            app = New Excel.Application()
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            app.Visible = False

            For Each xlsx In Directory.EnumerateFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                ProcessWorkbook(app, xlsx, moduleBasPath, arabicMapPath)
            Next

        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
        Finally
            If app IsNot Nothing Then
                app.DisplayAlerts = True
                app.ScreenUpdating = True
                app.Quit()
                ReleaseCom(app)
            End If
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

        Console.WriteLine("Done.")
    End Sub


    Private Sub ProcessWorkbook(app As Excel.Application, xlsxPath As String, moduleBasPath As String, arabicMapPath As String)
        Dim wb As Excel.Workbook = Nothing
        Try
            wb = app.Workbooks.Open(xlsxPath)

            ' 1) Ensure macro-enabled format on SAVE-AS
            Dim xlsmPath As String = Path.ChangeExtension(xlsxPath, ".xlsm")

            ' 2) Import the VBA module (.bas) into the project
            ImportVbaModule(wb, moduleBasPath)

            ' 3) Create/replace ArabicMap sheet and fill from txt
            UpsertArabicMapSheet(wb, arabicMapPath)

            ' 4) Update "unpaid" sheet
            UpdateUnpaidSheet(wb)

            ' 5) Save as .xlsm (macro-enabled)
            wb.SaveAs(Filename:=xlsmPath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)

            ' (Optional) Delete original .xlsx after successful save
            ' File.Delete(xlsxPath)

            Console.WriteLine($"Processed → {Path.GetFileName(xlsxPath)}  →  {Path.GetFileName(xlsmPath)}")

        Catch ex As Exception
            Console.WriteLine($"FAILED: {Path.GetFileName(xlsxPath)} → {ex.Message}")
        Finally
            If wb IsNot Nothing Then
                wb.Close(SaveChanges:=False)
                ReleaseCom(wb)
            End If
        End Try
    End Sub

    Private Sub ImportVbaModule(wb As Excel.Workbook, moduleBasPath As String)
        ' Requires Excel Trust Center: "Trust access to the VBA project object model"
        Dim vbProj As VBIDE.VBProject = Nothing
        Dim vbComp As VBIDE.VBComponent = Nothing
        Try
            vbProj = CType(wb.VBProject, VBIDE.VBProject)

            ' Optionally remove any existing module named "Module1" to avoid duplicates
            Dim existing As VBIDE.VBComponent = Nothing
            For Each comp As VBIDE.VBComponent In vbProj.VBComponents
                If String.Equals(comp.Name, "Module1", StringComparison.OrdinalIgnoreCase) Then
                    existing = comp
                    Exit For
                End If
            Next
            If existing IsNot Nothing Then
                vbProj.VBComponents.Remove(existing)
            End If

            ' Import .bas (this keeps the module name from inside the .bas, usually "Module1")
            vbComp = vbProj.VBComponents.Import(moduleBasPath)

        Catch ex As Exception
            Throw New Exception("VBA import failed. Ensure Excel setting 'Trust access to the VBA project object model' is enabled. " & ex.Message)
        Finally
            If vbComp IsNot Nothing Then ReleaseCom(vbComp)
            If vbProj IsNot Nothing Then ReleaseCom(vbProj)
        End Try
    End Sub

    Private Sub UpsertArabicMapSheet(wb As Excel.Workbook, arabicMapPath As String)
        Dim ws As Excel.Worksheet = Nothing
        Dim target As Excel.Worksheet = Nothing
        Try
            ' Delete if exists
            For Each s As Excel.Worksheet In wb.Worksheets
                If String.Equals(s.Name, "ArabicMap", StringComparison.OrdinalIgnoreCase) Then
                    s.Delete()
                    Exit For
                End If
            Next

            ' Add new sheet at end
            target = CType(wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)), Excel.Worksheet)
            target.Name = "ArabicMap"

            ' Headers
            target.Cells(1, 1).Value = "Key"
            target.Cells(1, 2).Value = "Text"

            ' Fill from the tab-separated file (Key<TAB>Text)
            Dim row As Integer = 2
            For Each line In File.ReadLines(arabicMapPath)
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim parts = line.Split(vbTab)
                If parts.Length >= 2 Then
                    target.Cells(row, 1).Value = parts(0).Trim()
                    target.Cells(row, 2).Value = parts(1).Trim()
                    row += 1
                End If
            Next

            ' Auto-fit
            target.Columns("A:B").EntireColumn.AutoFit()

        Catch ex As Exception
            Throw New Exception("ArabicMap creation failed: " & ex.Message)
        Finally
            If ws IsNot Nothing Then ReleaseCom(ws)
            If target IsNot Nothing Then ReleaseCom(target)
        End Try
    End Sub

    Private Sub UpdateUnpaidSheet(wb As Excel.Workbook)
        Dim ws As Excel.Worksheet = Nothing
        Try
            ' Find sheet "unpaid" (case-insensitive)
            For Each s As Excel.Worksheet In wb.Worksheets
                If String.Equals(s.Name, "unpaid", StringComparison.OrdinalIgnoreCase) Then
                    ws = s
                    Exit For
                End If
            Next
            If ws Is Nothing Then
                Throw New Exception("Sheet 'unpaid' not found.")
            End If

            ' Delete first row
            ws.Rows("1:1").Delete()

            ' Find last used row/col
            Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim lastCol As Integer = ws.Cells(1, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            ' Add new last column
            Dim newCol As Integer = lastCol + 1
            ws.Cells(1, newCol).Value = "TOTALInWORDS"

            ' Apply =ArabicDinars(Jrow) starting row 2 downward
            If lastRow >= 2 Then
                Dim firstDataCell As Excel.Range = ws.Cells(2, newCol)
                firstDataCell.Formula = "=ArabicDinars(J2)"
                Dim fillRange As Excel.Range = ws.Range(firstDataCell, ws.Cells(lastRow, newCol))
                firstDataCell.AutoFill(Destination:=fillRange, Type:=Excel.XlAutoFillType.xlFillDefault)

                ReleaseCom(fillRange)
                ReleaseCom(firstDataCell)
            End If

            ' Optional: format column as General/Text
            ws.Columns(newCol).NumberFormat = "@"

        Catch ex As Exception
            Throw New Exception("Updating 'unpaid' failed: " & ex.Message)
        Finally
            If ws IsNot Nothing Then ReleaseCom(ws)
        End Try
    End Sub

    Private Sub ReleaseCom(ByVal o As Object)
        Try
            If o IsNot Nothing AndAlso Marshal.IsComObject(o) Then
                Marshal.FinalReleaseComObject(o)
            End If
        Catch
        Finally
            o = Nothing
        End Try
    End Sub



End Class





