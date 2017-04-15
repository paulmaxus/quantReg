Imports alglib

Public Class ErrorTerm
    Public _alpha As Double
    Public _times As ArrayList
    Public _of As Byte
    Public _mf As Byte
    Public _mu As Double
    Public _start As Double


    Private _AverageLineExp As List(Of cls_Output)
    Public Property AverageLineExp() As List(Of cls_Output)
        Get
            Return _AverageLineExp
        End Get
        Set(ByVal value As List(Of cls_Output))
            _AverageLineExp = value
        End Set
    End Property

    Private _aQuantilesExp As List(Of cls_Output)
    Public Property aQuantilesExp() As List(Of cls_Output)
        Get
            Return _aQuantilesExp
        End Get
        Set(ByVal value As List(Of cls_Output))
            _aQuantilesExp = value
        End Set
    End Property


    Private _omaQuantilesExp As List(Of cls_Output)
    Public Property omaQuantilesExp() As List(Of cls_Output)
        Get
            Return _omaQuantilesExp
        End Get
        Set(ByVal value As List(Of cls_Output))
            _omaQuantilesExp = value
        End Set
    End Property




    Sub New(ByVal alpha As Double, ByVal times As ArrayList, ByVal start As Double)
        _alpha = alpha
        _times = times
        _start = start
    End Sub

    Function errorFunction(ByVal a As Double, ByVal b As Double, ByVal flag As Byte)
        Select Case flag
            Case 0
                errorFunction = (a - b) ^ 2
            Case 1
                errorFunction = System.Math.Abs((a / b) - 1) ^ 2
            Case Else
                Throw New ArgumentException("Flag nicht belegt")
        End Select
    End Function

    Function ErrorTerm(ByVal mu As Double, ByVal sigma As Double)
        Dim result As Double = 0
        Select Case _mf
            Case 0
                For i As Integer = 0 To _times.Count - 1
                    result = result + errorFunction(_AverageLineExp(i).Value, cls_GBMQuantile.GetAverage(mu, Long.Parse(_times(i)), _start), _of)
                Next
            Case 1
                For i As Integer = 0 To _times.Count - 1
                    result = result + errorFunction(_aQuantilesExp(i).Value, cls_GBMQuantile.GetQuantile(mu, sigma, _alpha, Long.Parse(_times(i)), _start), _of)
                    result = result + errorFunction(_omaQuantilesExp(i).Value, cls_GBMQuantile.GetQuantile(mu, sigma, 1 - _alpha, Long.Parse(_times(i)), _start), _of)
                Next
            Case Else
                For i As Integer = 0 To _times.Count - 1
                    result = result + errorFunction(_AverageLineExp(i).Value, cls_GBMQuantile.GetAverage(mu, Long.Parse(_times(i)), _start), _of)
                    result = result + errorFunction(_aQuantilesExp(i).Value, cls_GBMQuantile.GetQuantile(mu, sigma, _alpha, Long.Parse(_times(i)), _start), _of)
                    result = result + errorFunction(_omaQuantilesExp(i).Value, cls_GBMQuantile.GetQuantile(mu, sigma, 1 - _alpha, Long.Parse(_times(i)), _start), _of)
                Next
        End Select
        'Debug.Print(mu & " " & sigma)
        ErrorTerm = result
    End Function

    Public Function Compute(ByVal mu As Double, ByVal sigma As Double, ByVal opf As Byte, ByVal mf As Byte) As Double()
        ' 
        '  This example demonstrates minimization of error term
        '  using numerical differentiation to calculate gradient.
        '
        Dim x() As Double = New Double() {mu, sigma}
        Dim epsg As Double = 0.000001
        Dim epsf As Double = 0
        Dim epsx As Double = 0
        Dim diffstep As Double = 0.0001
        Dim maxits As Integer = 0
        Dim state As mincgstate = New XAlglib.mincgstate() ' initializer can be dropped, but compiler will issue warning
        Dim rep As mincgreport = New XAlglib.mincgreport() ' initializer can be dropped, but compiler will issue warning

        _of = opf
        _mf = mf
        If _mf = 1 Then
            _mu = mu
        End If
        XAlglib.mincgcreatef(x, diffstep, state)
        XAlglib.mincgsetcond(state, epsg, epsf, epsx, maxits)
        XAlglib.mincgoptimize(state, AddressOf function1_func, Nothing, Nothing)
        XAlglib.mincgresults(state, x, rep)
        'Debug.Print("{0}", alglib.ap.format(x, 5))
        Compute = x

    End Function

    'funktion (error term) prepared for algorithm
    Public Sub function1_func(ByVal x As Double(), ByRef func As Double, ByVal obj As Object)
        'different handling with different models
        If _mf = 0 Then
            func = ErrorTerm(x(0), 0)
        ElseIf _mf = 1 Then
            func = ErrorTerm(_mu, x(1))
        Else
            func = ErrorTerm(x(0), x(1))
        End If
    End Sub

    'Sub update(ByRef modelFlag As Byte, ByVal mu As Double, ByVal sigma As Double)
    '    _model.Mu = mu
    '    _model.Sigma = sigma
    '    AverageLineMod = _model.GetAverageLine
    '    If modelFlag = 0 Then Return
    '    aQuantilesMod = _model.GetQuantileLine(_alpha)
    '    If modelFlag = 1 Then Return
    '    omaQuantilesMod = _model.GetQuantileLine(1 - _alpha)
    'End Sub


    'Sub updateModelOne(ByVal mu As Double)
    '    _model.Mu = mu
    '    AverageLineMod = _model.GetAverageLine
    'End Sub


    'Function gradientDescent(ByVal opFlag As Byte, ByVal modelFlag As Byte) As Double
    '    Dim previous As Double = 0
    '    Dim actual As Double = math.Round(Compute(opFlag, modelFlag), 3)
    '    Dim muValue As Double
    '    Dim sigmaValue As Double
    '    Dim mu As Double = _model.Mu
    '    Dim sigma As Double = _model.Sigma

    '    While previous <> actual Or actual <> 0
    '        Debug.Print(actual)
    '        previous = actual
    '        update(modelFlag, mu + 0.001, sigma)
    '        muValue = math.Round(Compute(opFlag, modelFlag), 3)
    '        Debug.Print(muValue)
    '        update(modelFlag, mu, sigma + 0.01)
    '        'sigmaValue = math.Round(Compute(opFlag, modelFlag), 3)
    '        Debug.Print(sigmaValue)
    '        If math.Abs(actual - muValue) <= math.Abs(actual - sigmaValue) Then
    '            If actual < muValue Then
    '                mu = mu - 0.001
    '            ElseIf actual > muValue Then
    '                mu = mu + 0.001
    '            Else
    '                '           End If
    '      ElseIf actual < sigmaValue Then
    '                sigma = sigma - 0.01
    '    ElseIf actual > sigmaValue Then
    '                sigma = sigma + 0.01
    '  Else
    '            End If
    '            update(modelFlag, mu, sigma)
    '            actual = math.Round(Compute(opFlag, modelFlag), 3)
    '            Debug.Print(actual)
    '  End While
    '    Return mu
    'End Function


    'Function gDModelOne(ByVal opFlag As Byte) As Double
    '    Dim previous As Double = 0
    '    Dim actual As Double = math.Round(Compute(opFlag, 0), 3)
    '    Dim mu As Double = _model.Mu
    '    Dim value As Double

    '    While previous <> actual Or actual <> 0
    '        previous = actual
    '        updateModelOne(mu + 0.001)
    '        value = math.Round(Compute(opFlag, 0), 3)

    '        If actual < value Then
    '            mu = mu - 0.001
    '            updateModelOne(mu)
    '            actual = math.Round(Compute(opFlag, 0), 3)
    '        Else
    '            mu = mu + 0.001
    '            actual = value
    '        End If
    '    End While
    '    Return mu
    'End Function

End Class
