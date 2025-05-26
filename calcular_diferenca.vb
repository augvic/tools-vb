Function calcularDiferenca(ByVal datadeEntrega As Date, ByVal datadePrevisao As Date) As Integer

Dim diasUteis As Integer
Dim dia As Date

diasUteis = 0

If Not datadeEntrega = datadePrevisao Then
    For dia = IIf(datadeEntrega < datadePrevisao, datadeEntrega, datadePrevisao) To IIf(datadeEntrega < datadePrevisao, datadePrevisao, datadeEntrega)
        If Weekday(dia) <> vbSaturday And Weekday(dia) <> vbSunday And Not IsFeriado(inicioTransporte, feriados) Then
            diasUteis = diasUteis + 1
        End If
    Next
Else
    calcularDiferenca = 0
    Exit Function
End If

If datadeEntrega < datadePrevisao Then
    calcularDiferenca = -diasUteis
Else
    calcularDiferenca = diasUteis
End If

End Function