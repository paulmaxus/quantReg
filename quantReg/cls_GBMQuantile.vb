Public Module cls_GBMQuantile

    Public Function GetAverage(ByVal mu As Double, ByVal tau As Long, ByVal start As Double) As Double
        GetAverage = start * Math.Exp(mu * tau / 256)
    End Function

    Function GetQuantile(ByVal mu As Double, ByVal sigma As Double, ByVal alpha As Double, ByVal tau As Long, ByVal start As Double) As Double

        If alpha = 0 Then alpha = 0.001
        If alpha = 1 Then alpha = 0.999

        Dim X As Double = NormInv(alpha, 0, 1)

        GetQuantile = GeoBrownMotion(start, mu, sigma, tau / 256, X)

    End Function

    ''' <remarks>T in years, but t given as days, transformation is t/256</remarks>
    Function GeoBrownMotion(ByVal S_0 As Double, ByVal Mu As Double, ByVal Sigma As Double, ByVal T As Double, ByVal X As Double) As Double
        GeoBrownMotion = S_0 * Math.Exp((Mu - (Sigma ^ 2) / 2) * T + Sigma * Math.Sqrt(T) * X)
    End Function

    Function NormInv(ByVal Probability As Double, ByVal Mu As Double, ByVal Sigma As Double)

        Dim x As Double
        Dim p As Double
        Dim c0 As Double, c1 As Double, c2 As Double
        Dim d1 As Double, d2 As Double, d3 As Double
        Dim t As Double
        Dim q As Double

        q = Probability
        If (q = 0.5) Then
            NormInv = Mu
        Else
            q = 1.0 - q

            If ((q > 0) And (q < 0.5)) Then
                p = q
            Else
                If (q = 1) Then
                    p = 1 - 0.9999999
                Else
                    p = 1.0 - q
                End If
            End If

            t = Math.Sqrt(Math.Log(1.0 / (p * p)))

            c0 = 2.515517
            c1 = 0.802853
            c2 = 0.010328

            d1 = 1.432788
            d2 = 0.189269
            d3 = 0.001308

            x = t - (c0 + c1 * t + c2 * (t * t)) / (1.0 + d1 * t + d2 * (t * t) + d3 * (t ^ 3))

            If (q > 0.5) Then
                x = -1.0 * x
            End If
        End If

        NormInv = (x * Sigma) + Mu

    End Function
End Module
