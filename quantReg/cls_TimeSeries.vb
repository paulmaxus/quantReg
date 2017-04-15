Public Class cls_TimeSeries
    Private data As New List(Of Double)
    Private _dt As New DataTable



    Public Sub New(ByVal DataTableWithOneColumn As DataTable, ByVal directionFlag As Byte)
        _dt = DataTableWithOneColumn
        FillData(directionFlag)
    End Sub

    Private Sub FillData(ByVal directionFlag As Byte)
        Dim i As Long
        For i = 0 To _dt.Rows.Count - 1
            If IsNumeric(_dt.Rows(i)(0)) Then
                If directionFlag = 0 Then
                    data.Add(CDbl(_dt.Rows(i)(0)))
                Else
                    data.Insert(0, (CDbl(_dt.Rows(i)(0))))
                End If
            End If
        Next
    End Sub



    Private Function ConvertToDataTable(ByRef arraylist As List(Of Double)) As DataTable
        Dim dataTable As New DataTable
        dataTable.Columns.Add("Column1")

        For count As Integer = 0 To arraylist.Count - 1
            Dim drow As Data.DataRow = dataTable.NewRow
            drow(0) = CDbl(arraylist(count))
            dataTable.Rows.Add(drow)
        Next
        Return dataTable
    End Function

    Function GetDataColumn() As DataTable
        GetDataColumn = ConvertToDataTable(data)
    End Function

    Function GetPriceList() As List(Of Double)
        GetPriceList = data
    End Function

    Function RandomInvestment(ByVal BlockSize As Integer) As Double
        Dim BuyPosition As Long
        Dim SellPosition As Long
        BuyPosition = ((data.Count - 1) - BlockSize - 1) * Rnd() + BlockSize + 1
        SellPosition = BuyPosition - BlockSize
        RandomInvestment = data(SellPosition) / data(BuyPosition)
    End Function

    Function GetReturn(ByVal StartNo As Integer, ByVal EndNo As Integer) As Double
        GetReturn = data(EndNo) / data(StartNo)
    End Function

End Class
