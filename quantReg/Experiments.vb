
Public Module Experiments


    Function LasVegas(ByVal series As cls_TimeSeries, ByVal times As ArrayList, ByVal start As Double) As QuantileLinesAnalysis

        Dim quantileLines As New QuantileLinesAnalysis
        For i As Integer = 0 To times.Count - 1
            quantileLines.AddVerticalVector(times(i))
            For a As Integer = 0 To series.GetPriceList.Count - times(i) - 1
                Dim value As Double = start * series.GetReturn(a, a + times(i))
                quantileLines.AddData(times(i), value)
            Next
        Next
        LasVegas = quantileLines

    End Function


    Function MonteCarlo(ByVal n As Integer, ByVal series As cls_TimeSeries, ByVal times As ArrayList, ByVal start As Integer) As QuantileLinesAnalysis

        Dim quantileLines As New QuantileLinesAnalysis()
        For i As Integer = 0 To times.Count - 1
            quantileLines.AddVerticalVector(times(i))
            For z As Integer = 1 To n
                Dim a As Integer = Rnd() * (series.GetPriceList.Count - times(i) - 1)
                Dim value As Double = start * series.GetReturn(a, a + times(i))
                'Debug.Print(value)
                quantileLines.AddData(times(i), value)
            Next
        Next
        MonteCarlo = quantileLines
    End Function


    Sub BootsTrap()

    End Sub

End Module
