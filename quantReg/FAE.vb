Public Class FAE

    Public filePath As String
    Public savePath As String
    Public series As ArrayList
    Public headerList As List(Of String)

    Private Sub Fast_Access_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBoxTau.Text = Produce(TrackBar2.Value)
    End Sub

    Private Sub NewFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewFileToolStripMenuItem.Click
        OpenFileDialog1.ShowDialog()
        filePath = OpenFileDialog1.FileName
        If filePath = Nothing Then
            Exit Sub
        End If
        Try
            MultiFile.Show()
        Catch ex As Exception
            MsgBox("invalid file")
            Exit Sub
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If filePath = Nothing Then
            MsgBox("Please load file")
            Exit Sub
        End If

        SaveFileDialog1.ShowDialog()

        savePath = SaveFileDialog1.FileName
        If savePath = Nothing Then
            Exit Sub
        End If

        TextBoxTau.Text.Trim()
        Dim storage As String() = TextBoxTau.Text.Split(New Char() {","})
        Dim tau = New ArrayList
        For i = 0 To storage.Count - 1
            tau.Add(Integer.Parse(storage(i)))
        Next

        ProgressBar1.Increment(5)

        Dim qrmcList = New List(Of QuantRegMainControl)
        For z = 0 To series.Count - 1
            Dim qrmc = New QuantRegMainControl
            qrmc.algo = 1
            qrmc.alpha = 0.05
            qrmc.ofg = 0
            qrmc.start = 1000
            qrmc.tau = tau
            qrmc.timeSeries = series(z)
            qrmcList.Add(qrmc)
        Next
        Dim dt = New DataTable
        dt.Columns.Add("Feature", Type.GetType("System.String"))
        For i = 0 To headerList.Count - 1
            dt.Columns.Add(headerList(i), Type.GetType("System.Double"))
        Next

        ProgressBar1.Increment(5)

        Dim rowOne = dt.NewRow()
        rowOne(0) = "1: Rendite"
        For z As Integer = 0 To qrmcList.Count - 1
            qrmcList(z).mf = 0
            Dim result = qrmcList(z).estimateParameter(0)
            rowOne(z + 1) = result
            qrmcList(z).mu = result
        Next
        dt.Rows.Add(rowOne)

        ProgressBar1.Increment(20)

        Dim rowTwo = dt.NewRow()
        rowTwo(0) = "2: Risiko"
        For z As Integer = 0 To qrmcList.Count - 1
            qrmcList(z).mf = 1
            rowTwo(z + 1) = qrmcList(z).estimateParameter(1)
        Next
        dt.Rows.Add(rowTwo)

        ProgressBar1.Increment(30)

        Dim rowThree = dt.NewRow()
        rowThree(0) = "3: Rendite"
        Dim rowFour = dt.NewRow()
        rowFour(0) = "3: Risiko"
        For z As Integer = 0 To qrmcList.Count - 1
            qrmcList(z).mf = 2
            Dim result = qrmcList(z).estimateParameter
            rowThree(z + 1) = result(0)
            rowFour(z + 1) = result(1)
        Next
        dt.Rows.Add(rowThree)
        dt.Rows.Add(rowFour)

        ProgressBar1.Increment(20)

        ExportDataGridContentToCSVFile(savePath, dt)

        ProgressBar1.Increment(20)
        MsgBox("Computation was successful")
        ProgressBar1.Value = 0

    End Sub

    Private Function ExportDataGridContentToCSVFile(ByVal Filename As String, ByVal dt As DataTable) As Boolean
        ' Die Variable Created übernimmt den Kontrollwert,
        ' ob die Datei angelegt wurde.
        Dim Created As Boolean = False
        ' Fehlerüberwachung einschalten
        Try
            ' StreamWriter initialisieren
            Using sw As IO.StreamWriter = New IO.StreamWriter( _
              Filename, False, System.Text.Encoding.Unicode)

                ' Spalten anlegen
                For n As Integer = 0 To dt.Columns.Count - 1
                    sw.Write(dt.Columns(n))
                    If (n < dt.Columns.Count - 1) Then
                        sw.Write(",")
                    End If
                Next

                ' Neue Zeile schreiben und...
                sw.Write(sw.NewLine())

                ' ... den Inhalt des Grids in eine Komma 
                ' getrennte Datei speichern.
                For Each dr As DataRow In dt.Rows()
                    For n As Integer = 0 To dt.Columns.Count - 1
                        If Not Convert.IsDBNull(dr(n)) Then
                            sw.Write(dr(n).ToString())
                        End If
                        If (n < dt.Columns.Count - 1) Then
                            sw.Write(",")
                        End If
                    Next

                    ' Neue Zeile anlegen.
                    sw.Write(sw.NewLine())
                Next
            End Using

            ' Wurde die Datei angelegt wird die Kontrollvariable 
            ' Created mit True initialisiert
            If IO.File.Exists(Filename) Then Created = True

        Catch ex As IO.IOException
            ' Eventuell auftretenden Fehler abfangen
            MessageBox.Show(ex.Message(), "Info - IOException")
        Catch ex As Exception
            MessageBox.Show(ex.Message(), "Info - Exception")
        End Try

        ' Funktionsrückgabe
        Return Created
    End Function

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
End Class