Public Class cls_CSVFileReader

    Private _fileLineArray() As String = Nothing
    Private _columns As Integer
    Private _rows As Integer
    Private _ContentArray As String()()
    Private _headerInside As Boolean
    Public HeaderList As New List(Of String)

    Public Sub New(ByVal path As String, ByVal Delimiter As String, ByVal HeaderInside As Boolean)
        _columns = 0
        _rows = 0
        _headerInside = HeaderInside

        Try
            _fileLineArray = System.IO.File.ReadAllLines(path)
            _ContentArray = GetContentArray(Delimiter)
            UpdateRowsCount()
            UpdateColumnsCount()
            UpdateHeader(HeaderInside)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub

    Private Function GetContentArray(ByVal Delimiter As String) As String()()
        Dim fileContentArray(_fileLineArray.Length - 1)() As String

        Try

            Dim i As Integer = 0
            For i = 0 To _fileLineArray.Length - 1
                Dim line As String = _fileLineArray(i)
                fileContentArray(i) = line.Split(Delimiter)
            Next

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try

        Return fileContentArray
    End Function


    Public Property RowsCount() As Long
        Get
            Return _rows
        End Get
        Set(ByVal value As Long)

        End Set
    End Property

    Public Property ColumnsCount() As Long
        Get
            Return _columns
        End Get
        Set(ByVal value As Long)

        End Set
    End Property

    Public Property GetDataArray() As String()()
        Get
            Return _ContentArray
        End Get
        Set(ByVal value As String()())

        End Set
    End Property

    Private Sub UpdateRowsCount()
        Try
            _rows = UBound(_ContentArray, 1) + 1
        Catch ex As Exception
            _rows = 0
        End Try
    End Sub

    Private Sub UpdateColumnsCount()
        Try
            _columns = UBound(_ContentArray(0), 1) + 1
        Catch ex As Exception
            _columns = 0
        End Try
    End Sub

    Sub UpdateHeader(ByVal HeaderInside As Boolean)
        Dim i As Integer
        For i = 0 To _columns - 1
            If HeaderInside Then HeaderList.Add(GetDataArray(0)(i)) Else HeaderList.Add("C" & i + 1)
        Next i
        i = Nothing
    End Sub


    Function GetDataColumn(ByVal columnIndex As Integer, ByVal columnIsNumeric As Boolean) As List(Of String)
        If (columnIndex > _columns) Or (columnIndex < 1) Then columnIndex = 1
        Dim i As Long
        Dim temp As New List(Of String)
        Dim startValue As Byte
        If _headerInside = True Then startValue = 1 Else startValue = 0
        For i = startValue To _rows - 1
            If columnIsNumeric = True Then
                If IsNumeric(_ContentArray(i)(columnIndex - 1)) Then temp.Add(_ContentArray(i)(columnIndex - 1).ToString)
            Else
                temp.Add(_ContentArray(i)(columnIndex - 1).ToString)
            End If
        Next i
        GetDataColumn = temp
        temp = Nothing
    End Function

    Private Function ConvertToDataTable(ByRef arraylist As List(Of String)) As Data.DataTable
        Dim dataTable As New Data.DataTable
        dataTable.Columns.Add("Column1")
        dataTable.Columns.Add("Column2")
        Dim startValue As Integer

        If _headerInside = True Then startValue = 1 Else startValue = 0
        For count As Integer = 0 To arraylist.Count - 1
            Dim drow As Data.DataRow = dataTable.NewRow
            drow(0) = arraylist(count)
            drow(1) = arraylist(count)
            dataTable.Rows.Add(drow)
        Next
        Return dataTable
    End Function

    Public Function GetDataTableColumn(ByVal columnIndex As Integer) As DataTable
        If (columnIndex > _columns) Or (columnIndex < 1) Then columnIndex = 1
        Dim temp As New DataTable
        temp = GetDataTable()
        While temp.Columns.Count > columnIndex
            temp.Columns.RemoveAt(columnIndex)
        End While
        While temp.Columns.Count > 1
            temp.Columns.RemoveAt(0)
        End While
        Return temp
    End Function


    Public Function GetDataTable() As Data.DataTable
        Dim dataTable As New Data.DataTable
        For j = 1 To ColumnsCount
            dataTable.Columns.Add(HeaderList(j - 1))
        Next j

        Dim startValue As Integer
        If _headerInside = True Then startValue = 1 Else startValue = 0

        For cur_row As Integer = startValue To RowsCount - 1 'change endValue to length - 1
            Dim drow As Data.DataRow = dataTable.NewRow
            For cur_col As Integer = 1 To ColumnsCount
                drow(cur_col - 1) = _ContentArray(cur_row)(cur_col - 1)
            Next cur_col
            'drow(0) = _ContentArray(cur_row)
            dataTable.Rows.Add(drow)
        Next
        Return dataTable
    End Function


    Function GetHeaderTable() As DataTable
        Dim datatable As New Data.DataTable
        datatable.Columns.Add("Headers")
        For count As Integer = 0 To HeaderList.Count - 1
            Dim drow As Data.DataRow = datatable.NewRow
            drow(0) = HeaderList(count)
            datatable.Rows.Add(drow)
        Next
        Return datatable
    End Function

End Class
