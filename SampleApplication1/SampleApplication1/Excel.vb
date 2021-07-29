Imports System
Imports System.IO
Imports Scripting
Imports Spreadsheet = Microsoft.Office.Interop.Excel

Public Class Excel
    Dim actualPath As String
    Dim idealPath As String
    Dim fileMark = 0
    Dim pathMark = 0

    'spreadsheet variable declaration'
    Dim app
    Dim workbook As Spreadsheet.Workbook
    Dim worksheet As Spreadsheet.Worksheet
    Dim usedRange
    Dim usedRange2DArray As Object(,)

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog1.Filter = "Excel Files (*.xls*)|*.xls*|CSV files (*.csv)|*.csv"
        OpenFileDialog1.RestoreDirectory = True

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            idealPath = OpenFileDialog1.FileName
            TextBox2.Text = idealPath

            app = New Spreadsheet.Application()
            workbook = app.Workbooks.Open(idealPath)
            worksheet = workbook.Sheets(1)

            usedRange = worksheet.UsedRange

            fileMark = 1

            Dim ExcelObject As Object

            For rows = 2 To usedRange.Rows.Count
                Dim filePath As String = ""
                For columns = 1 To usedRange.Columns.Count

                    ExcelObject = CType(usedRange.Cells(rows, columns), Spreadsheet.Range)

                    If (ExcelObject.value = "" And usedRange.Columns.Count = columns) Then
                        filePath = filePath + "\"
                        Exit For
                    ElseIf (ExcelObject.value = "" And usedRange.Columns.Count <> columns) Then
                        Continue For
                    Else
                        filePath = filePath + ExcelObject.value.ToString + "\"
                    End If

                Next
                ListView2.Items.Add(filePath.Substring(0, filePath.Length() - 1))
            Next

        End If

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            ActualPathTB.Text = FolderBrowserDialog1.SelectedPath
            actualPath = FolderBrowserDialog1.SelectedPath

            Dim path = FolderBrowserDialog1.SelectedPath

            Dim files() As String = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories)

            For Each file As String In files
                ListView1.Items.Add(file.Replace(actualPath + "\", ""))
            Next

        End If

        pathMark = 1

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If (pathMark <> 1 And fileMark <> 1) Then
            MessageBox.Show("There are no files or folders selected.")
        ElseIf (pathMark <> 1) Then
            MessageBox.Show("Please select the folder path containing the directory/directories to be evaluated.")
        ElseIf (fileMark <> 1) Then
            MessageBox.Show("Please select the file containing the ideal path structures.")
        Else

            Dim reportForm = New GenerateReport()
            Dim fso As New FileSystemObject

            Dim offset = 0

            Dim idealData = ListView2.Items
            Dim actualData = ListView1.Items

            For iterator = 0 To idealData.Count - 1
                Dim idealCaseFlag = 0
                Dim array(3)
                Dim fileTest As String
                Dim fileName = actualPath + "\" + idealData(iterator).Text.ToString

                array(0) = idealData(iterator).Text.ToString

                If (fileName.Substring(fileName.Length - 1) = "\") Then
                    'MessageBox.Show("Ideal Path for this is only a Folder")'
                    array(3) = "X"

                    Dim actualEval = actualPath + "\" + actualData(iterator - offset).Text.ToString
                    Dim actualFileName = fso.GetFileName(actualEval)
                    Dim actualFolderPath = actualEval.Replace(actualFileName, "")

                    If (fileName = actualEval) Then
                        array(2) = "O"

                        Dim idealEval = fileName

                        If (idealEval.Replace(actualPath + "\", "") = actualFolderPath) Then
                            array(1) = actualEval.Replace(actualPath + "\", "")
                        Else

                            While (idealEval <> "")
                                idealEval = RecursePath(0, (idealEval.Length - 1), idealEval)

                                If (My.Computer.FileSystem.DirectoryExists(idealEval)) Then
                                    array(2) = "X"
                                    array(1) = actualData(iterator - offset).Text.ToString
                                    offset = offset + 1
                                    Exit While
                                End If

                            End While

                        End If
                    Else
                        If (My.Computer.FileSystem.DirectoryExists(fileName)) Then
                            array(2) = "O"
                            array(1) = fileName.Replace(actualPath + "\", "")
                            offset = offset + 1
                        Else
                            While (fileName <> "")
                                fileName = RecursePath(0, (fileName.Length - 1), fileName)

                                If (My.Computer.FileSystem.DirectoryExists(fileName)) Then
                                    array(2) = "X"
                                    array(1) = actualData(iterator - offset).Text.ToString
                                    Exit While
                                End If

                            End While

                            If (fileName = "") Then
                                array(2) = "X"
                                array(1) = "Does not exist"
                            End If
                        End If

                        

                    End If

                Else
                        'MessageBox.Show("Ideal Path contains a File")'
                        fileTest = fso.GetFileName(fileName)

                        Dim actualEval = actualPath + "\" + actualData(iterator - offset).Text.ToString
                        Dim actualFileName = fso.GetFileName(actualEval)
                        Dim actualFolderPath = actualEval.Replace(actualFileName, "")

                        Dim idealEval = actualPath + "\" + idealData(iterator).Text.ToString
                        Dim idealFileName = fso.GetFileName(idealEval)
                        Dim idealFolderPath = idealEval.Replace(idealFileName, "")

                        If (idealEval = actualEval) Then
                            'the foldder and  file structure are the same'
                            array(1) = actualData(iterator - offset).Text.ToString
                            array(2) = "O"
                            array(3) = "O"
                        Else
                            'there is something different with the two paths'

                            'filename comparison'
                            If (idealFileName = actualFileName) Then
                                array(3) = "O"
                            Else
                                array(3) = "X"
                            End If


                            'folder structure comparison'
                            If (idealFolderPath = actualFolderPath) Then
                                array(1) = actualData(iterator - offset).Text.ToString
                                array(2) = "O"
                            Else
                                Dim idealHolder = idealFolderPath
                                Dim idealHolder2 = idealFolderPath

                                While (idealFolderPath <> "")
                                    idealFolderPath = RecursePath(0, (idealFolderPath.Length - 1), idealFolderPath)

                                    If (idealFolderPath = actualFolderPath.Substring(0, actualFolderPath.Length - 2)) Then
                                        array(2) = "X"
                                        array(1) = actualEval
                                        idealCaseFlag = 1
                                        Exit While
                                    End If

                                End While

                                If (idealCaseFlag = 0) Then
                                    array(2) = "X"

                                    While (idealHolder <> "")
                                        idealHolder = RecursePath(0, (idealHolder.Length - 1), idealHolder)

                                        If (idealHolder = actualFolderPath.Substring(0, actualFolderPath.Length - 1)) Then
                                            array(1) = actualEval.Replace(actualPath + "\", "")
                                            idealCaseFlag = 1
                                            Exit While
                                        End If

                                    End While

                                    If (idealCaseFlag = 0) Then
                                        While (idealHolder2 <> "")
                                            idealHolder2 = RecursePath(0, (idealFolderPath.Length - 1), idealHolder2)

                                            If (My.Computer.FileSystem.DirectoryExists(idealHolder2)) Then
                                                array(1) = idealHolder2.Replace(actualPath + "\", "")
                                                Exit While
                                            End If
                                        End While

                                        offset = offset + 1
                                    End If

                                End If


                            End If
                        End If
                End If
                reportForm.ListView1.Items.Add(New ListViewItem(New String() {array(0), array(1), array(2), array(3)}))
            Next
            Me.Hide()
            MessageBox.Show("Proceed")
            reportForm.storageDirectory = RecursePath(0, idealPath.Length - 1, idealPath)
            reportForm.Show()
        End If

    End Sub

    Public Function RecursePath(ByVal a As Integer, ByVal b As Integer, ByVal path As String) As String

        If (path.Length = 1) Then
            Return ""
        ElseIf (path.Substring(path.Length - 1) <> "\") Then
            path = path.Substring(0, (path.Length - 1))
            Return RecursePath(0, (path.Length) - 1, path)
        Else
            path = path.Substring(0, (path.Length - 1))
            Return path
        End If

    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If (pathMark <> 1 And fileMark <> 1) Then
            MessageBox.Show("There are no files or folders selected.")
        ElseIf (pathMark <> 1) Then
            MessageBox.Show("Please select the folder path containing the directory/directories to be evaluated.")
        ElseIf (fileMark <> 1) Then
            MessageBox.Show("Please select the file containing the ideal path structures.")
        Else

            Dim reportForm = New GenerateReport()
            Dim fso As New FileSystemObject

            Dim offset = 0

            Dim idealData = ListView2.Items
            Dim actualData = ListView1.Items

            For iterator = 0 To idealData.Count - 1
                Dim idealCaseFlag = 0
                Dim array(3)
                Dim fileTest As String
                Dim fileName = actualPath + "\" + idealData(iterator).Text.ToString

                array(0) = idealData(iterator).Text.ToString

                If (fileName.Substring(fileName.Length - 1) = "\") Then
                    'MessageBox.Show("Ideal Path for this is only a Folder")'

                    array(3) = "X"

                    Dim actualEval = actualPath + "\" + actualData(iterator - offset).Text.ToString
                    Dim actualFileName = fso.GetFileName(actualEval)
                    Dim actualFolderPath = actualEval.Replace(actualFileName, "")


                    'folder and file check'
                    If (fileName = actualFolderPath) Then
                        array(2) = "O"
                        array(1) = actualEval.Replace(actualPath + "\", "")
                    End If

                    'folder and folder check'
                    If (fileName <> actualFolderPath) Then
                        Dim check = fileName

                        If (My.Computer.FileSystem.DirectoryExists(fileName)) Then
                            array(2) = "O"
                            array(1) = fileName.Replace(actualPath + "\", "")
                            offset = offset + 1
                        Else
                            While (fileName <> "")
                                fileName = RecursePath(0, (fileName.Length - 1), fileName)

                                If (My.Computer.FileSystem.DirectoryExists(fileName)) Then
                                    If (fileName = actualFolderPath) Then
                                        array(2) = "X"
                                        array(1) = actualData(iterator - offset).Text.ToString
                                        Exit While
                                    Else
                                        array(2) = "X"
                                        array(1) = fileName.Replace(actualPath + "\", "")
                                        offset = offset + 1
                                        Exit While
                                    End If


                                End If

                            End While

                            If (fileName = "") Then
                                array(2) = "X"
                                array(1) = "Does not exist in directory"
                                offset = offset + 1
                            End If
                        End If
                    End If

                Else
                    'MessageBox.Show("Ideal Path contains a File")'
                    fileTest = fso.GetFileName(fileName)

                    Dim actualEval = actualPath + "\" + actualData(iterator - offset).Text.ToString
                    Dim actualFileName = fso.GetFileName(actualEval)
                    Dim actualFolderPath = actualEval.Replace(actualFileName, "")

                    Dim idealEval = actualPath + "\" + idealData(iterator).Text.ToString
                    Dim idealFileName = fso.GetFileName(idealEval)
                    Dim idealFolderPath = idealEval.Replace(idealFileName, "")

                    If (idealEval = actualEval) Then
                        'the foldder and  file structure are the same'
                        array(1) = actualData(iterator - offset).Text.ToString
                        array(2) = "O"
                        array(3) = "O"
                    Else
                        'there is something different with the two paths'

                        'filename comparison'
                        If (idealFileName = actualFileName) Then
                            array(3) = "O"
                        Else
                            array(3) = "X"
                        End If


                        'folder structure comparison'
                        If (idealFolderPath = actualFolderPath) Then
                            array(1) = actualData(iterator - offset).Text.ToString
                            array(2) = "O"
                        Else
                            Dim idealHolder = idealFolderPath
                            Dim idealHolder2 = idealFolderPath

                            While (idealFolderPath <> "")
                                idealFolderPath = RecursePath(0, (idealFolderPath.Length - 1), idealFolderPath)

                                If (idealFolderPath = actualFolderPath.Substring(0, actualFolderPath.Length - 2)) Then
                                    array(2) = "X"
                                    array(1) = actualEval
                                    idealCaseFlag = 1
                                    Exit While
                                End If

                            End While

                            If (idealCaseFlag = 0) Then
                                array(2) = "X"

                                While (idealHolder <> "")
                                    idealHolder = RecursePath(0, (idealHolder.Length - 1), idealHolder)

                                    If (idealHolder = actualFolderPath.Substring(0, actualFolderPath.Length - 1)) Then
                                        array(1) = actualEval.Replace(actualPath + "\", "")
                                        idealCaseFlag = 1
                                        Exit While
                                    End If

                                End While

                                If (idealCaseFlag = 0) Then
                                    While (idealHolder2 <> "")
                                        idealHolder2 = RecursePath(0, (idealFolderPath.Length - 1), idealHolder2)

                                        If (My.Computer.FileSystem.DirectoryExists(idealHolder2)) Then
                                            array(1) = idealHolder2.Replace(actualPath + "\", "")
                                            Exit While
                                        End If
                                    End While

                                    offset = offset + 1
                                End If

                            End If


                        End If
                    End If
                End If
                reportForm.ListView1.Items.Add(New ListViewItem(New String() {array(0), array(1), array(2), array(3)}))
            Next
            Me.Hide()
            MessageBox.Show("Proceed")
            reportForm.storageDirectory = RecursePath(0, idealPath.Length - 1, idealPath)
            reportForm.Show()
        End If

    End Sub
End Class