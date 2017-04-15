Public Class MultiFile

    Private tempCSV As cls_CSVFileReader

    Private Sub MultiFile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UpdateOpen()
    End Sub

    Sub UpdateOpen()

        DataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect
        tempCSV = New cls_CSVFileReader(FAE.filePath, TextBoxD.Text, CheckBox1.Checked)
        DataGridView1.DataSource = tempCSV.GetDataTable
        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns()(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        UpdateOpen()
    End Sub

    Private Sub NewFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewFileToolStripMenuItem.Click
        OpenFileDialog1.ShowDialog()
        FAE.filePath = OpenFileDialog1.FileName
        If FAE.filePath = Nothing Then
            Exit Sub
        End If
        UpdateOpen()
    End Sub

    Private Sub TextBoxD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxD.TextChanged
        UpdateOpen()
    End Sub

    Private Sub ButtonExecute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExecute.Click

        ' checking if at least one column is selected
        Try
            Dim test = DataGridView1.SelectedColumns.Item(0)
        Catch ex As Exception
            MsgBox("Please select at least one column")
            Exit Sub
        End Try

        ' checking if data is descending
        Dim i As Byte
        If CheckDirection.Checked Then
            i = 1
        Else
            i = 0
        End If

        FAE.series = New ArrayList
        FAE.headerList = New List(Of String)
        Dim hd As List(Of String) = tempCSV.HeaderList
        For z = 0 To Integer.Parse(tempCSV.ColumnsCount) - 1
            If DataGridView1.Columns(z).Selected Then

                Dim data = tempCSV.GetDataTableColumn(z + 1)
                Dim pseudo As Double
                For n As Integer = 0 To data.Rows.Count - 1

                    If Not Double.TryParse(data.Rows(n).Item(0).ToString, pseudo) Then
                        MsgBox("Data is not continuously numerical")
                        Exit Sub
                    End If

                Next

                FAE.series.Add(New cls_TimeSeries(data, i))
                FAE.headerList.Add(hd(z))

            End If
        Next

        Close()

    End Sub

End Class