Public Class QuantRegMainControl
    Private _timeSeries As cls_TimeSeries
    Private _algo As Byte
    Private _n As Integer
    Private _of As Byte
    Private _mf As Byte = 5
    Private _mu As Double = 0.08
    Private _tau As ArrayList
    Private _start As Double
    Private _alpha As Double
    Private _qla As QuantileLinesAnalysis

    Public Property timeSeries() As cls_TimeSeries
        Get
            Return _timeSeries
        End Get
        Set(ByVal value As cls_TimeSeries)
            _timeSeries = value
        End Set
    End Property

    Public Property algo() As Byte
        Get
            Return _algo
        End Get
        Set(ByVal value As Byte)
            _algo = value
        End Set
    End Property

    Public Property n() As Integer
        Get
            Return _n
        End Get
        Set(ByVal value As Integer)
            _n = value
        End Set
    End Property

    Public Property ofg() As Byte
        Get
            Return _of
        End Get
        Set(ByVal value As Byte)
            _of = value
        End Set
    End Property

    Public Property mf() As Byte
        Get
            Return _mf
        End Get
        Set(ByVal value As Byte)
            _mf = value
        End Set
    End Property

    Public Property mu() As Double
        Get
            Return _mu
        End Get
        Set(ByVal value As Double)
            _mu = value
        End Set
    End Property

    Public Property tau() As ArrayList
        Get
            Return _tau
        End Get
        Set(ByVal value As ArrayList)
            _tau = value
        End Set
    End Property

    Public Property start() As Double
        Get
            Return _start
        End Get
        Set(ByVal value As Double)
            _start = value
        End Set
    End Property

    Public Property alpha() As Double
        Get
            Return _alpha
        End Get
        Set(ByVal value As Double)
            _alpha = value
        End Set
    End Property

    Public Property qla() As QuantileLinesAnalysis
        Get
            Return _qla
        End Get
        Set(ByVal value As QuantileLinesAnalysis)
            _qla = value
        End Set
    End Property


    Public Sub simpleQLA()

        Dim qla As QuantileLinesAnalysis
        If _algo = 0 Then
            qla = Experiments.MonteCarlo(_n, _timeSeries, _tau, _start)
        Else
            qla = Experiments.LasVegas(_timeSeries, _tau, _start)
        End If

        _qla = qla

    End Sub

    Public Function estimateParameter() As Double()
        simpleQLA()
        Dim _et As New ErrorTerm(_alpha, _tau, _start)
        _et.AverageLineExp = _qla.GetAverageLine
        _et.aQuantilesExp = _qla.GetQuantileLine(_alpha)
        _et.omaQuantilesExp = _qla.GetQuantileLine(1 - _alpha)
        estimateParameter = _et.Compute(_mu, 0.2, _of, _mf)
    End Function



End Class
