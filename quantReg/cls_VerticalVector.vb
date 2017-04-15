Public Class cls_VerticalVector
    Private _data As New List(Of Double)
    Private _name As String
    Private _isSorted As Boolean
    Private _isAvgCorrect As Boolean
    Private _average As Double
    Private _itemsCount As Boolean
    Private _vectorName As String

    Public Property IdName() As String
        Get
            Return _vectorName
        End Get
        Set(ByVal value As String)
            _vectorName = value
        End Set
    End Property

    Public Property Count() As Long
        Get
            Return _data.Count
        End Get
        Set(ByVal value As Long)

        End Set
    End Property

    Public Sub New()
        _isSorted = False
        _isAvgCorrect = False
        _name = Nothing
        _vectorName = "_default"
    End Sub

    Public Sub addItem(ByVal ItemValue As Double)
        _data.Add(ItemValue)
        _isSorted = False
        _isAvgCorrect = False
    End Sub

    Function GetAverage()
        If Count() > 0 Then
            If _isAvgCorrect = False Then
                _average = _data.Average()
                _isAvgCorrect = True
            End If

            GetAverage = _average

        Else
            GetAverage = Nothing
        End If
    End Function

    Function GetQuantile(ByVal alpha As Double)
        If Count() > 0 And alpha >= 0 And alpha <= 1 Then
            SortData()
            Dim alphaPosition As Integer
            alphaPosition = ((_data.Count - 1) * alpha)
            GetQuantile = _data(alphaPosition)
        Else
            GetQuantile = Nothing
        End If
    End Function


    Function GetMin()
        If Count() > 0 Then
            SortData()
            GetMin = _data(0)
        Else
            GetMin = Nothing
        End If
    End Function

    Function GetMax()
        If Count() > 0 Then
            SortData()
            GetMax = _data(_data.Count - 1)
        Else
            GetMax = Nothing
        End If
    End Function

    Private Sub SortData()
        If _isSorted = False Then
            _data.Sort()
            _isSorted = True
        End If
    End Sub

    Function GetQuantileForValue(ByVal xValue As Double) As  double
        Dim counter As Long
        If xValue >= GetMax() Then
            counter = _data.Count
        ElseIf xValue <= GetMin() Then
            counter = 0
        Else
            counter = 1
            While _data(counter - 1) < xValue
                counter = counter + 1
            End While
        End If



        If Count > 0 Then
            GetQuantileForValue = (counter) / Count
        Else
            GetQuantileForValue = Nothing
        End If
    End Function

End Class
