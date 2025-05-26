Function IsFeriado(ByVal dataVerificar As Date, ByVal feriados As Variant) As Boolean

Dim i As Integer
For i = LBound(feriados) To UBound(feriados)
    If dataVerificar = feriados(i) Then
        IsFeriado = True
        Exit Function
    End If
Next i

IsFeriado = False
End Function