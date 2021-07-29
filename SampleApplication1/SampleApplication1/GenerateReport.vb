Imports Spreadsheet = Microsoft.Office.Interop.Excel

Public Class GenerateReport
    Public storageDirectory
    Public duplicates

    Dim ThisMoment As Date


    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim file As System.IO.StreamWriter

        ThisMoment = Now

        file = My.Computer.FileSystem.OpenTextFileWriter(storageDirectory + "\SummaryReport" + ".txt", True)
        file.WriteLine("Summary Report: " + ThisMoment)
        file.WriteLine("")
        file.WriteLine("Number of Duplicate Files: " + Label10.Text)
        'insert code for list of file duplicates and their paths'
        file.WriteLine("")
        file.WriteLine("Non-Matching Files: " + Lbl_NEF.Text)
        file.WriteLine("Non-Matching Directories: " + Label8.Text)
        file.WriteLine("Non-Matching Files & Directories: " + Label9.Text)
        file.WriteLine("")
        file.WriteLine("Matching Files: " + Label11.Text)
        file.WriteLine("Matching Directories: " + Label12.Text)
        file.WriteLine("Matching Files & Directories: " + Label14.Text)

        MessageBox.Show("Summary Report Successfully generated!")
        Process.Start(storageDirectory)

        file.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim xlApp As New Spreadsheet.Application
        Dim xlWorkbook As Spreadsheet.Workbook = xlApp.Workbooks.Add
        Dim xlWorksheet As Spreadsheet.Worksheet

        xlWorksheet = CType(xlWorkbook.Worksheets(1), Spreadsheet.Worksheet)

        Dim columns = Me.ListView1.Columns.Count - 1
        Dim rows = Me.ListView1.Items.Count - 1

        For i = 0 To columns
            xlWorksheet.Cells(1, i + 1) = Me.ListView1.Columns(i).Text
        Next

        For i = 0 To rows
            For j = 0 To Me.ListView1.Items(i).SubItems.Count - 1
                xlWorksheet.Cells(i + 2, j + 1) = Me.ListView1.Items(i).SubItems(j).Text
            Next
        Next

        Dim processString = RandString(10)

        Dim filePath = storageDirectory + "\TabularReport_" + processString + ".xlsx"

        xlWorkbook.SaveAs(filePath)

        MessageBox.Show("Tabular Report Successfully generated!")
        Process.Start(storageDirectory)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Application.Exit()
        End
    End Sub

    Function RandString(ByVal n As Long) As String
        Dim i As Long
        Dim j As Long
        Dim m As Long
        Dim s As String
        Dim pool As String
        pool = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        m = Len(pool)
        For i = 1 To n
            j = 1 + Int(m * Rnd())
            s = s & Mid(pool, j, 1)
        Next i
        RandString = s
    End Function

End Class