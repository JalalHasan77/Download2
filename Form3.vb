Public Class Form3
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ICBS_CS As String
        Dim EBDB_CS As String

        'ICBS_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBSTRN3)));User Id=ICBS;Password=SBCI;"
        'ICBS_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.20.16)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ICBS)));User Id=ICBS;Password=icbs01;"

        'EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=s0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=EBDB)));User Id=INAP;Password=inap;"

        'EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=s0230)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=EBDB)));User Id=inap;Password=inap;"

        'EBDB_CS = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=s0230)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=EBDB)));User Id=inap;Password=inap;" 'live


        EBDB_CS = "Provider = OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=EBDB)));User Id=inap;Password=inap;"
        'Dim Dcon As New OleDb.OleDbConnection(ICBS_CS)
        Dim Dcon As New OleDb.OleDbConnection(EBDB_CS)
        Dim DT0 As New DataTable
        Dim DT As New DataTable
        Dcon.Open()
        DT0 = Dcon.GetSchema("TABLES")
        Dcon.Close()
        DT = DT0.Clone

        For Each dr As DataRow In DT0.Rows
            If dr("TABLE_NAME").ToString Like "TRSRY_*" Then 'Or dr("COLUMN_NAME").ToString Like "*_NATIONAL_NBR*" Then
                ComboBox1.Items.Add(dr("TABLE_NAME").ToString)
            End If
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ConnectionString As String
        If RadioButton1.Checked Then
            ConnectionString = "Provider = OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0320)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=EBDB)));User Id=inap;Password=inap;"
        Else
            ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=S0344)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=EBDB)));User Id=INAP;Password=inap;"
        End If

        Dim Dcon As New OleDb.OleDbConnection(connectionString:=ConnectionString)
        'Dcon.ConnectionString = ConnectionString
        Dim DT0 As New DataTable
        Dim DT As New DataTable

        Dcon.Open()
        DT0 = Dcon.GetSchema("COLUMNS", {Nothing, Nothing, ComboBox1.SelectedItem.ToString, Nothing})
        'DT0 = Dcon.GetSchema("COLUMNS", {Nothing, Nothing, Nothing, Nothing})
        Dcon.Close()

        TextBox1.Text = ComboBox1.SelectedItem.ToString & "@" & IIf(RadioButton1.Checked, "Test_Unit", "Live_Unit")
        For Each DR As DataRow In DT0.Rows
            'vbCrLf & DR("TABLE_NAME").ToString _
            TextBox1.Text = TextBox1.Text & vbCrLf _
            & DR("COLUMN_NAME").ToString _
            & vbTab & DR("DATA_TYPE").ToString _
            & vbTab & DR("CHARACTER_MAXIMUM_LENGTH").ToString
        Next

        DataGridView1.DataSource = DT0.DefaultView

    End Sub
End Class