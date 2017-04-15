Public Class SingleFile
    Dim tempCSV As cls_CSVFileReader

    Private Sub Form_File_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UpdateOpen()
    End Sub

    Sub UpdateOpen()
        ErrorProvider1.Clear()
        Try
            tempCSV = New cls_CSVFileReader(Professional.filePath, TextBoxD.Text, CheckBox1.Checked)
            DataGridView1.DataSource = tempCSV.GetDataTable
            ComboBox1.Items.Clear()
            Dim i As Integer
            For i = 0 To tempCSV.ColumnsCount - 1
                ComboBox1.Items.Add(tempCSV.HeaderList(i))
            Next i
            i = Nothing
            ComboBox1.SelectedIndex = -1
            ComboBox1.Text = Nothing
        Catch ex As Exception
            ErrorProvider1.SetError(ButtonOpen, "Invalid file")
        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        UpdateOpen()
    End Sub

    Private Sub ButtonOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOpen.Click
        OpenFileDialog1.ShowDialog()
        Professional.filePath = OpenFileDialog1.FileName
        If Professional.filePath = Nothing Then
            Exit Sub
        End If
        UpdateOpen()
    End Sub

    Private Sub TextBoxD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxD.TextChanged
        UpdateOpen()
    End Sub

    Private Sub ButtonExecute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExecute.Click

        Try
            If Not fileChecking() Then
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Something went wrong")
            Exit Sub
        End Try

        ' checking if data is descending, if true, data must be reversed
        If CheckDirection.Checked Then
            Professional._qrmc.timeSeries = New cls_TimeSeries(tempCSV.GetDataTableColumn(ComboBox1.SelectedIndex + 1), 1)
        Else
            Professional._qrmc.timeSeries = New cls_TimeSeries(tempCSV.GetDataTableColumn(ComboBox1.SelectedIndex + 1), 0)
        End If
        Close()
    End Sub

    Private Function fileChecking()

        If ComboBox1.SelectedIndex = -1 Then
            MsgBox("Select a column")
            Return False
        End If

        ' checking if table contains double
        Dim data As DataRowCollection = tempCSV.GetDataTableColumn(ComboBox1.SelectedIndex + 1).Rows
        Dim pseudo As Double
        For i As Integer = 0 To data.Count - 1

            If Not Double.TryParse(data(i).Item(0).ToString, pseudo) Then
                MsgBox("Data is not continuously numerical")
                Return False
            End If

        Next

        Return True

    End Function

End Class