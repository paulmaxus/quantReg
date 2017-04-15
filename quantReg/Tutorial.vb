Public Class Tutorial

    Private tempCSV As cls_CSVFileReader
    Public _filepath As String

    Private Sub Wizard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBoxTau.Text = Produce(TrackBar2.Value)
        ComboBoxType.SelectedIndex = 1
        OpFlag.SelectedIndex = 0
    End Sub

    ' choose file section
    Private Sub ButtonOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewFileButton.Click
        OpenFileDialog1.ShowDialog()
        _filepath = OpenFileDialog1.FileName
        If _filepath = Nothing Then
            Exit Sub
        End If
        UpdateOpen()
    End Sub

    Sub UpdateOpen()

        ErrorProvider1.Clear()

        If String.IsNullOrEmpty(_filepath) Then
            Exit Sub
        Else
            Try
                tempCSV = New cls_CSVFileReader(_filepath, DelimiterTextBox.Text, CheckBoxHeader.Checked)
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
                ErrorProvider1.SetError(NewFileButton, "Invalid file")
            End Try
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxHeader.CheckedChanged
        UpdateOpen()
    End Sub

    Private Sub TextBoxD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelimiterTextBox.TextChanged
        UpdateOpen()
    End Sub

    Private Sub Continue1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Continue1.Click
        TabControl1.SelectTab(1)
    End Sub

    Private Sub Continue2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Continue2.Click
        TabControl1.SelectTab(2)
    End Sub

    Private Sub Finish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Finish.Click

        If String.IsNullOrEmpty(_filepath) Then
            TabControl1.SelectedTab = TabPage1
            ErrorProvider1.SetError(NewFileButton, "Load file")
            Exit Sub
        End If

        Try
            If Not fileChecking() Then
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ComboBox1, "File reading went wrong")
            TabControl1.SelectedTab = TabPage1
            Exit Sub
        End Try

        Professional.Show()

        If DescendingCheckBox.Checked Then
            Professional._qrmc.timeSeries = New cls_TimeSeries(tempCSV.GetDataTableColumn(ComboBox1.SelectedIndex + 1), 1)
        Else
            Professional._qrmc.timeSeries = New cls_TimeSeries(tempCSV.GetDataTableColumn(ComboBox1.SelectedIndex + 1), 0)
        End If

        Professional.NumericUpDownInvestment.Value = NumericUpDownInvestment.Value

        Professional.ComboBoxType.SelectedIndex = ComboBoxType.SelectedIndex

        If TextBox1.Visible Then
            Professional.TextBoxN.Text = TextBox1.Text
            Professional.TextBoxN.Visible = True
        End If

        Professional.Alpha.Text = NumericUpDownPrecision.Value
        Professional.Alpha.Visible = True
        Professional.TrackBar1.Value = NumericUpDownPrecision.Value
        Professional.TrackBar1.Visible = True

        Professional.TextBoxTau.Text = TextBoxTau.Text

        Professional.OpFlag.SelectedIndex = OpFlag.SelectedIndex
        Professional.TabControl1.SelectedIndex = TabControl2.SelectedIndex
        If Not String.IsNullOrEmpty(MuValue.Text) Then
            Professional.MuValue.Text = Double.Parse(MuValue.Text)
        End If

        Professional.Execute()
        Professional.TrackBar1.Show()
        Professional.Alpha.Show()

        Close()

    End Sub

    Private Sub ComboBoxType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxType.SelectedIndexChanged
        If ComboBoxType.SelectedIndex = 0 Then
            TrackBar1.Visible = True
            TextBox1.Visible = True
        Else
            TrackBar1.Visible = False
            TextBox1.Visible = False
        End If
    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        TextBox1.Text = TrackBar1.Value * 100
    End Sub

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        TextBoxTau.Text = Produce(TrackBar2.Value)
    End Sub

    ' generates line of tau
    Private Function Produce(ByVal i As Integer) As String
        Dim chain = New System.Text.StringBuilder
        For a As Integer = 0 To 255 Step i
            chain.Append(a & ", ")
        Next
        ' delete last ", "
        chain.Remove(chain.Length - 2, 2)
        ' fill textbox
        Produce = chain.ToString
    End Function

    Private Function fileChecking() As Boolean

        If ComboBox1.SelectedIndex = -1 Then
            TabControl1.SelectedTab = TabPage1
            ErrorProvider1.SetError(ComboBox1, "Select a column")
            Return False
        End If

        ' checking if table contains double
        Dim data As DataRowCollection = tempCSV.GetDataTableColumn(ComboBox1.SelectedIndex + 1).Rows
        Dim pseudo As Double
        For i As Integer = 0 To data.Count - 1

            If Not Double.TryParse(data(i).Item(0).ToString, pseudo) Then
                TabControl1.SelectedTab = TabPage1
                MsgBox("Data is not continuously numerical")
                Return False
            End If

        Next

        Return True

    End Function

End Class