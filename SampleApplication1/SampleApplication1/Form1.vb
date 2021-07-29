Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Form1_Proceed.Click
        Me.Hide()
        If (String.IsNullOrEmpty(ComboBox1.SelectedItem)) Then
            MessageBox.Show("Please Select an Option")
            Me.Show()
        ElseIf (ComboBox1.SelectedItem = "Batch Process Filenames") Then
            Dim csvForm = New CSV()
            csvForm.Show()
        ElseIf (ComboBox1.SelectedItem = "File and Path Structure Evaluation") Then
            Dim excelForm = New Excel()
            excelForm.Show()
        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class
