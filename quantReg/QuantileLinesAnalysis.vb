Public Class QuantileLinesAnalysis
    Dim _dataSet As New List(Of cls_VerticalVector)


    Public Sub New()

    End Sub

    Public Property Count() As Integer
        Get
            Return _dataSet.Count
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property


    Sub AddVerticalVector(ByVal vectorName As String)
        If idNamePosition(vectorName) = -1 Then
            Dim a As New cls_VerticalVector
            a.IdName = vectorName
            _dataSet.Add(a)
            a = Nothing
        Else

        End If
    End Sub

    Sub AddData(ByVal vectorName As String, ByVal value As Double)
        Dim currentPos As Integer
        currentPos = idNamePosition(vectorName)
        If currentPos = -1 Then
            AddVerticalVector(vectorName)
            currentPos = idNamePosition(vectorName)
        End If
        _dataSet(currentPos).addItem(value)
    End Sub

    Private Function idNamePosition(ByVal idName As String) As Integer
        idNamePosition = _dataSet.FindIndex((Function(c As cls_VerticalVector) c.IdName = idName))
    End Function

    Sub Clear()
        _dataSet.Clear()
    End Sub

    Function GetAverageLine() As List(Of cls_Output)
        Dim temp As New List(Of cls_Output)
        For Each element In _dataSet
            Dim a As New cls_Output(element.IdName, element.GetAverage)
            temp.Add(a)
            a = Nothing
        Next
        GetAverageLine = temp
        temp = Nothing
    End Function

    Function GetQuantileLine(ByVal alpha As Double) As List(Of cls_Output)
        Dim temp As New List(Of cls_Output)
        For Each element In _dataSet
            Dim a As New cls_Output(element.IdName, element.GetQuantile(alpha))
            temp.Add(a)
            a = Nothing
        Next
        GetQuantileLine = temp
        temp = Nothing
    End Function

    Function GetCountElementsLine() As List(Of cls_Output)
        Dim temp As New List(Of cls_Output)
        For Each element In _dataSet
            Dim a As New cls_Output(element.IdName, element.Count)
            temp.Add(a)
            a = Nothing
        Next
        GetCountElementsLine = temp
        temp = Nothing
    End Function

    Function GetMinLine() As List(Of cls_Output)
        Dim temp As New List(Of cls_Output)
        For Each element In _dataSet
            Dim a As New cls_Output(element.IdName, element.GetMin)
            temp.Add(a)
            a = Nothing
        Next
        GetMinLine = temp
        temp = Nothing
    End Function

    Function GetMaxLine() As List(Of cls_Output)
        Dim temp As New List(Of cls_Output)
        For Each element In _dataSet
            Dim a As New cls_Output(element.IdName, element.GetMax)
            temp.Add(a)
            a = Nothing
        Next
        GetMaxLine = temp
        temp = Nothing
    End Function

    Function GetQuantileForValueLine(ByVal xValue As Double) As List(Of cls_Output)
        Dim temp As New List(Of cls_Output)
        For Each element In _dataSet
            Dim a As New cls_Output(element.IdName, element.GetQuantileForValue(xValue))
            temp.Add(a)
            a = Nothing
        Next
        GetQuantileForValueLine = temp
        temp = Nothing
    End Function

End Class