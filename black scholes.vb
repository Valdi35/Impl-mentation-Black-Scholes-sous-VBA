'Fonction de répartition de la loi normale centrée réduite

Function FRNCR(x As Double) As Double

FRNCR = Application.NormSDist(x)

End Function

'Fonction de densité de la loi normale centrée réduite

Function FDNCR(x As Double) As Double

FDNCR = Exp((-x) ^ 2 / 2) / Sqr(2 * Application.Pi())

End Function
'Pricing d'une option par Black et Scholes

Function BSCallPut(S As Double, K As Double, r As Double, T As Double, v As Double, q As Double, PutCall As String) As Double

Dim d1 As Double
Dim d2 As Double

d1 = (Application.Ln(S / K) + (r - q + v ^ 2 / 2) * T) / v * Sqr(T)
d2 = d1 - v * Sqr(T)

'd1 = (Application.Ln(S / K) + (r - q + v ^ 2 / 2) * T) / v * Sqr(T)
Select Case PutCall
Case "Call"
BSCallPut = S * Exp(-q * T) * FRNCR(d1) - K * Exp(-r * T) * FRNCR(d2)
Case "Put"
BSCallPut = K * Exp(-r * T) * FRNCR(-d2) - S * Exp(-q * T) * FRNCR(-d1)
End Select

End Function
