Imports System.Data
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Professional
    Public filePath As String
    Public _qrmc As QuantRegMainControl
    Public _result() As Double

    Private Sub Form_Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If _qrmc Is Nothing Then
            _qrmc = New QuantRegMainControl
        End If

        TextBoxTau.Text = Produce(TrackBar2.Value)
        ComboBoxType.SelectedIndex = 1
        OpFlag.SelectedIndex = 0

    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        OpenFileDialog1.ShowDialog()
        filePath = OpenFileDialog1.FileName
        If filePath = Nothing Then
            Exit Sub
        End If
        ErrorProvider1.Clear()
        TrackBar1.Visible = False
        Alpha.Visible = False
        SingleFile.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Start.Click

        If String.IsNullOrEmpty(filePath) Then
            MsgBox("No file loaded")
            Exit Sub
        End If

        Execute()
        TrackBar1.Show()
        Alpha.Show()
    End Sub

    Public Sub Execute()

        _qrmc.start = NumericUpDownInvestment.Value
        _qrmc.algo = ComboBoxType.SelectedIndex
        If TextBoxN.Visible Then
            _qrmc.n = Integer.Parse(TextBoxN.Text)
        End If
        _qrmc.alpha = Double.Parse(Alpha.Text)


        TextBoxTau.Text.Trim()
        Dim storage As String() = TextBoxTau.Text.Split(New Char() {","})
        Dim tau As ArrayList = New ArrayList
        For i = 0 To storage.Count - 1
            tau.Add(Long.Parse(storage(i)))
        Next
        _qrmc.tau = tau


        _qrmc.ofg = OpFlag.SelectedIndex
        _qrmc.mf = TabControl1.SelectedIndex
        If TabControl1.SelectedIndex = 1 Then
            _qrmc.mu = Double.Parse(MuValue.Text)
        End If
        _result = _qrmc.estimateParameter()

        UpdateGraph(_qrmc.alpha)

    End Sub

    Public Sub UpdateGraph(ByVal alpha As Double)

        Dim elementsInside As Integer = _qrmc.qla.GetCountElementsLine.Count

        Chart1.Series(0).Points.Clear()
        Chart1.Series(1).Points.Clear()
        Chart1.Series(2).Points.Clear()
        Chart1.Series(3).Points.Clear()
        Chart1.Series(4).Points.Clear()
        Chart1.Series(5).Points.Clear()

        Dim avg As List(Of cls_Output) = _qrmc.qla.GetAverageLine
        Dim aqtl As List(Of cls_Output) = _qrmc.qla.GetQuantileLine(alpha)
        Dim omaqtl As List(Of cls_Output) = _qrmc.qla.GetQuantileLine(1 - alpha)

        For i = 0 To elementsInside - 1

            If _qrmc.mf = 0 Or _qrmc.mf = 1 Or _qrmc.mf = 2 Then
                Chart1.Series(3).Points.AddXY(avg(i).Name, cls_GBMQuantile.GetAverage(_result(0), Long.Parse(avg(i).Name), _qrmc.start))
                If _qrmc.mf = 1 Or _qrmc.mf = 2 Then
                    Chart1.Series(4).Points.AddXY(aqtl(i).Name, cls_GBMQuantile.GetQuantile(_result(0), _result(1), alpha, Long.Parse(aqtl(i).Name), _qrmc.start))
                    Chart1.Series(5).Points.AddXY(omaqtl(i).Name, cls_GBMQuantile.GetQuantile(_result(0), _result(1), 1 - alpha, Long.Parse(omaqtl(i).Name), _qrmc.start))
                End If
            End If

            Chart1.Series(0).Points.AddXY(avg(i).Name, avg(i).Value)
            Chart1.Series(1).Points.AddXY(aqtl(i).Name, aqtl(i).Value)
            Chart1.Series(2).Points.AddXY(omaqtl(i).Name, omaqtl(i).Value)

        Next i

    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Alpha.Text = TrackBar1.Value / 100
        Execute()
    End Sub

    Private Sub ComboBoxType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBoxType.SelectedIndexChanged
        If ComboBoxType.SelectedIndex = 0 Then
            TextBoxN.Visible = True
            TrackBar3.Visible = True
        Else
            TextBoxN.Visible = False
            TrackBar3.Visible = False
        End If
    End Sub

    Private Sub TrackBar3_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar3.Scroll
        TextBoxN.Text = TrackBar3.Value * 100
    End Sub

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        TextBox2.Text = TrackBar2.Value
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

End Class