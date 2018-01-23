Function DifEnElTiempo(Inicio As Date, Final As Date, TimeUnit As String)
‘Devuelve la diferencia total en segundos
Select Case TimeUnit

Case Is = “segundos”
‘Devolver la diferencia en segundos
DifEnElTiempo = Abs(Final – Inicio) * 86400

Case Is = “minutos”
‘Devolver la diferencia en minutos
DifEnElTiempo = Abs(Final – Inicio) * 1440

Case Is = “horas”
‘Devolver la diferencia en horas
DifEnElTiempo = Abs(Final – Inicio) * 24

Case Is = “dias”
‘Devolver la diferencia en días
DifEnElTiempo = Abs(Final – Inicio)

Case Is = “meses”
‘Devolver la diferencia en meses
DifEnElTiempo = (Abs(Final – Inicio)) / 30

Case Else

DifEnElTiempo = “Revisar argumentos”

End Select

End Function
