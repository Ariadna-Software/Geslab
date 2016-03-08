Attribute VB_Name = "HorasDec"
Option Explicit


'Dada una hora devuelve un numero en formato Single
Public Function DevuelveValorHora(vHora As Date) As Single
Dim AUx As Single
Dim C

    C = Minute(vHora)
    AUx = C / 60
    DevuelveValorHora = Hour(vHora) + Round(AUx, 2)
End Function




'Dada una hora en centesimal la pasamos a formato hora
Public Function DevuelveHora(vHora As Single) As Date
Dim X
Dim Y
Dim Cad As String

    vHora = Abs(Round(vHora, 2))
    X = Int(vHora)
    Y = vHora - X
    'En y esta la parte centesimal de una hora
    Y = Round(Y * 60, 0)
    X = X Mod 24
    Cad = X & ":" & Y & ":00"
    DevuelveHora = CDate(Cad)
End Function
