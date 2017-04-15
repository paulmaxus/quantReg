Public Class cls_Output
    Private _name As String
    Private _value As Double

    Public Sub New(ByVal Name As String, ByVal value As Double)
        _name = Name
        _value = value
    End Sub


    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Value() As Double
        Get
            Return _value
        End Get
        Set(ByVal value As Double)
            _value = value
        End Set
    End Property


End Class
