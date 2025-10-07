Public Class Form6

    Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long,
   ByVal szURL As String,
   ByVal szFileName As String,
   ByVal dwReserved As Long,
   ByVal lpfnCB As Long) As Long

    Private Const ERROR_SUCCESS As Long = 0
    Private Const BINDF_GETNEWESTVERSION As Long = &H10
    Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sSourceUrl As String
        Dim sLocalFile As String
        Dim hfile As Long

        sSourceUrl = "https://www.usatoday.com/"
        sLocalFile = "D:\deleteme.htm"

        Label1.Text = sSourceUrl
        Label2.Text = sLocalFile

        If DownloadFile(sSourceUrl, sLocalFile) Then

            '   hfile = FreeFile()
            '   FileOpen sLocalFile For Input As #hfile
            'TextBox1.Text = Input$(LOF(hfile), hfile)
            '   Close #hfile

        End If
    End Sub

    Public Function DownloadFile(sSourceUrl As String,
                             sLocalFile As String) As Boolean

        'Download the file. BINDF_GETNEWESTVERSION forces 
        'the API to download from the specified source. 
        'Passing 0& as dwReserved causes the locally-cached 
        'copy to be downloaded, if available. If the API 
        'returns ERROR_SUCCESS (0), DownloadFile returns True.
        DownloadFile = URLDownloadToFile(0&,
                                         sSourceUrl,
                                         sLocalFile,
                                         BINDF_GETNEWESTVERSION,
                                         0&) = ERROR_SUCCESS

    End Function

End Class