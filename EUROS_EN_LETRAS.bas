Attribute VB_Name = "Módulo1"
Option Explicit
' ==============================================
'  EUROS_EN_LETRAS (español – formato factura)
'  - 100      -> "CIEN EUROS"
'  - 1234,50  -> "MIL DOSCIENTOS TREINTA Y CUATRO EUROS CON CINCUENTA CÉNTIMOS"
'  - 1        -> "UN EURO"
'  - -5,3     -> "MENOS CINCO EUROS CON TREINTA CÉNTIMOS"

' ==============================================

Public Function EUROS_EN_LETRAS(ByVal Importe As Variant) As String
    On Error GoTo ErrHandler
    If Not IsNumeric(Importe) Then EUROS_EN_LETRAS = "": Exit Function

    Dim total As Currency, euros As Currency, cent As Integer, neg As Boolean
    total = CCur(Importe)

    ' Negativos
    If total < 0 Then
        neg = True
        total = -total
    End If

    ' Redondeo a 2 decimales
    total = Redondeo2(total)

    euros = Fix(total)
    cent = CInt((total - euros) * 100)

    ' Por si un borde deja 100 céntimos
    If cent = 100 Then
        euros = euros + 1
        cent = 0
    End If

    Dim txt As String
    txt = NumeroEnLetrasCLng(CLng(euros), True)
    txt = txt & IIf(euros = 1, " euro", " euros")

    ' Solo añadimos céntimos si > 0
    If cent > 0 Then
        txt = txt & " con " & NumeroEnLetrasCLng(CLng(cent), True)
        txt = txt & IIf(cent = 1, " céntimo", " céntimos")
    End If

    If neg Then txt = "menos " & txt
    EUROS_EN_LETRAS = UCase$(txt)
    Exit Function

ErrHandler:
    EUROS_EN_LETRAS = "#ERROR"
End Function

' Redondeo a 2 decimales
Private Function Redondeo2(ByVal x As Currency) As Currency
    If x >= 0 Then
        Redondeo2 = CCur(Int(x * 100 + 0.5) / 100)
    Else
        Redondeo2 = CCur(-Int(-x * 100 + 0.5) / 100)
    End If
End Function

' Convierte enteros 0..2.147.483.647 a letras (español)
' apocope=True usa "un" / "veintiún" ante sustantivo
Private Function NumeroEnLetrasCLng(ByVal n As Long, ByVal apocope As Boolean) As String
    Dim resto As Long, texto As String

    If n = 0 Then
        NumeroEnLetrasCLng = "cero"
        Exit Function
    End If

    ' Millones
    If n >= 1000000 Then
        Dim millones As Long
        millones = n \ 1000000
        resto = n Mod 1000000

        If millones = 1 Then
            texto = "un millón"
        Else
            texto = NumeroEnLetrasCLng(millones, True) & " millones"
        End If

        If resto > 0 Then texto = texto & " " & NumeroEnLetrasCLng(resto, apocope)
        NumeroEnLetrasCLng = texto
        Exit Function
    End If

    ' Miles
    If n >= 1000 Then
        Dim miles As Long
        miles = n \ 1000
        resto = n Mod 1000

        If miles = 1 Then
            texto = "mil"
        Else
            texto = NumeroEnLetrasCLng(miles, True) & " mil"
        End If

        If resto > 0 Then texto = texto & " " & NumeroEnLetrasCLng(resto, apocope)
        NumeroEnLetrasCLng = texto
        Exit Function
    End If

    ' Cientos
    If n >= 100 Then
        Dim cientos As Long
        cientos = n \ 100
        resto = n Mod 100

        Select Case cientos
            Case 1: If resto = 0 Then texto = "cien" Else texto = "ciento"
            Case 2: texto = "doscientos"
            Case 3: texto = "trescientos"
            Case 4: texto = "cuatrocientos"
            Case 5: texto = "quinientos"
            Case 6: texto = "seiscientos"
            Case 7: texto = "setecientos"
            Case 8: texto = "ochocientos"
            Case 9: texto = "novecientos"
        End Select

        If resto > 0 Then texto = texto & " " & NumeroEnLetrasCLng(resto, apocope)
        NumeroEnLetrasCLng = texto
        Exit Function
    End If

    ' Decenas y unidades
    NumeroEnLetrasCLng = DecenasYUnidades(n, apocope)
End Function

Private Function DecenasYUnidades(ByVal n As Long, ByVal apocope As Boolean) As String
    Dim u As Variant
    u = Array("", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve")

    Select Case n
        Case 0: DecenasYUnidades = "cero"
        Case 1 To 9
            DecenasYUnidades = IIf(apocope And n = 1, "un", u(n))
        Case 10: DecenasYUnidades = "diez"
        Case 11: DecenasYUnidades = "once"
        Case 12: DecenasYUnidades = "doce"
        Case 13: DecenasYUnidades = "trece"
        Case 14: DecenasYUnidades = "catorce"
        Case 15: DecenasYUnidades = "quince"
        Case 16: DecenasYUnidades = "dieciséis"
        Case 17: DecenasYUnidades = "diecisiete"
        Case 18: DecenasYUnidades = "dieciocho"
        Case 19: DecenasYUnidades = "diecinueve"
        Case 20: DecenasYUnidades = "veinte"
        Case 21: DecenasYUnidades = IIf(apocope, "veintiún", "veintiuno")
        Case 22: DecenasYUnidades = "veintidós"
        Case 23: DecenasYUnidades = "veintitrés"
        Case 24: DecenasYUnidades = "veinticuatro"
        Case 25: DecenasYUnidades = "veinticinco"
        Case 26: DecenasYUnidades = "veintiséis"
        Case 27: DecenasYUnidades = "veintisiete"
        Case 28: DecenasYUnidades = "veintiocho"
        Case 29: DecenasYUnidades = "veintinueve"
        Case 30: DecenasYUnidades = "treinta"
        Case 40: DecenasYUnidades = "cuarenta"
        Case 50: DecenasYUnidades = "cincuenta"
        Case 60: DecenasYUnidades = "sesenta"
        Case 70: DecenasYUnidades = "setenta"
        Case 80: DecenasYUnidades = "ochenta"
        Case 90: DecenasYUnidades = "noventa"
        Case Else
            Dim dec As Long, uni As Long
            dec = (n \ 10) * 10
            uni = n Mod 10
            DecenasYUnidades = DecenasYUnidades(dec, apocope) & " y " & IIf(apocope And uni = 1, "un", u(uni))
    End Select
End Function

