Attribute VB_Name = "libCCBanco"
Public Function CodigoDeControl(ByVal strBanOfiCuenta As String) As String

Dim conPesos
Dim lngPrimerCodigo As Long, lngSegundoCodigo As Long
Dim I As Long, J As Long
conPesos = "06030709100508040201"
J = 1
lngPrimerCodigo = 0
lngSegundoCodigo = 0

' Banco(4) + Oficina(4) nos dará el primer dígito de control
For I = 8 To 1 Step -1
  lngPrimerCodigo = lngPrimerCodigo + (Mid$(strBanOfiCuenta, I, 1) * Mid$(conPesos, J, 2))
  J = J + 2
Next I

J = 1 ' reiniciar el contador de pesos

' Número de cuenta nos dará el segundo digito de control
For I = 18 To 9 Step -1
  lngSegundoCodigo = lngSegundoCodigo + (Mid$(strBanOfiCuenta, I, 1) * Mid$(conPesos, J, 2))
  J = J + 2
Next I


' ajustar el primer dígito de control
lngPrimerCodigo = 11 - (lngPrimerCodigo Mod 11)
If lngPrimerCodigo = 11 Then
    lngPrimerCodigo = 0
ElseIf lngPrimerCodigo = 10 Then
    lngPrimerCodigo = 1
End If


' ajustar el segundo dígito de control
lngSegundoCodigo = 11 - (lngSegundoCodigo Mod 11)
If lngSegundoCodigo = 11 Then
    lngSegundoCodigo = 0
ElseIf lngSegundoCodigo = 10 Then
    lngSegundoCodigo = 1
End If

' convertirlos en cadenas y concatenarlos
CodigoDeControl = Format(lngPrimerCodigo) & Format(lngSegundoCodigo)

End Function


