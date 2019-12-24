Attribute VB_Name = "modBoletos"
Public Function CODBOLETO(CODBARRAS)

CODBOLETO = Linha_Digitavel(CODBARRAS)

End Function

Public Function CODBARRAS(ValorDoFaturamento, DatadeVencimento, cdmoeda, banco, AGENCIA, CONTA, conta_dac, CARTEIRA, NumDocumento)

CODBARRAS = Monta_CodBarras(ValorDoFaturamento, DatadeVencimento, cdmoeda, banco, AGENCIA, CONTA, conta_dac, CARTEIRA, NumDocumento, Calculo_DV10(Formata(AGENCIA, 4) & Formata(CONTA, 5) & Formata(CARTEIRA, 3) & Formata(NumDocumento, 8)))

End Function

Public Function NOSSONUMERO(CARTEIRA, NumDocumento, AGENCIA, CONTA)

NOSSONUMERO = CARTEIRA & "/" & NumDocumento & "-" & Calculo_DV10(Formata(AGENCIA, 4) & Formata(CONTA, 5) & Formata(CARTEIRA, 3) & Formata(NumDocumento, 8))

End Function

Public Function NumDocumento(codCadastro, DatadeVencimento)

NumDocumento = Formata(FormatNumber(codCadastro, 0) & Format(DatadeVencimento, "ddmmyy"), 8)

End Function

Public Function FormatarValor(strValor)

FormatarValor = FormatNumber(strValor, 2)

End Function

Public Function FormatarData(strData)

FormatarData = Format(strData, "dd/mm/yyyy")

End Function

Public Function Formata(VALOR, Tam, Optional Dec)
Dim aux As String
Dim a As Integer

aux = Trim(VALOR)
If Not IsMissing(Dec) Then
   aux = FormatNumber(VALOR, Dec)
End If
aux = Replace(aux, ",", "")
aux = Replace(aux, ".", "")
If Len(aux) > Tam Then
   aux = Right(aux, Tam)
ElseIf Len(aux) < Tam Then
   For a = 1 To Tam - Len(aux)
       aux = "0" + aux
   Next a
End If
Formata = aux

End Function
    
Public Function Monta_CodBarras(valor1, vencimento1, MOEDA, banco, AGENCIA, CONTA, conta_dac, CARTEIRA, NOSSONUMERO, dv_nossonumero)
    
    Dim database, fator, VALOR, dvcb, codigo_sequencia As String
    
    database = CDate("7/10/1997")
    fator = DateDiff("d", database, vencimento1)
    VALOR = Int(valor1 * 100)
    
    Do While Len(VALOR) < 10
        VALOR = "0" & VALOR
    Loop
    
    'codigo_sequencia = Formata(Banco, 3) & Formata(Moeda, 1) & Formata(fator, 4) & Formata(Valor, 10) & Formata(Carteira, 3) & Formata(nossonumero, 8) & Formata(dv_nossonumero, 1) & Formata(Agencia, 4) & Formata(Conta, 5) & Formata(Conta_dac, 1) & "000"
    codigo_sequencia = Formata(banco, 3) & Formata(MOEDA, 1) & Formata(fator, 4) & Formata(VALOR, 10) & Formata(AGENCIA, 4) & Formata(CONTA, 7) & Formata(conta_dac, 1) & Formata(NOSSONUMERO, 13)
    
    dvcb = calcula_DV_CodBarras(codigo_sequencia)
    
    Monta_CodBarras = Left(codigo_sequencia, 4) & dvcb & Right(codigo_sequencia, 39)
    
End Function
    
Function Linha_Digitavel(sequencia_codigo_barra)
    
    Dim seq1, seq2, seq3, seq4, dvcb, dv1, dv2, dv3
    
    '         10        20        30        40
    '12345678901234567890123456789012345678901234
    '3419 7233 5 00000059 00175 02280204 4 2923 05456 9 000
    '3419 1233 2 00000059 00175 01250204 2 2923 05456 9 000
    'codigo_sequencia = Formata(Banco, 3) & Formata(Moeda, 1) & Formata(fator, 4) & Formata(Valor, 10) & Formata(Agencia, 4) & Formata(Conta, 7) & Formata(Conta_dac, 1) & Formata(nossonumero, 13)
    
    '
    'Boleto Real
    '
    seq1 = Left(sequencia_codigo_barra, 4) & Mid(sequencia_codigo_barra, 20, 5)
    seq2 = Mid(sequencia_codigo_barra, 25, 10)
    seq3 = Mid(sequencia_codigo_barra, 35, 10)
    seq4 = Mid(sequencia_codigo_barra, 6, 14)
    
    'seq1 = banco & moeda & agencia & left(cc,1)
    'seq2 = right( cc, 6 ) & mid( dv_conta,31, 1) & left( nossonumero, 3 )
    'seq3 = right( nossonumero, 10 )
    'seq4 = right( sequencia_codigo_barra,14)
    
    '
    'Boleto Itau
    '
    'seq1 = Left(sequencia_codigo_barra, 4) & Mid(sequencia_codigo_barra, 20, 5)
    'seq2 = Mid(sequencia_codigo_barra, 25, 10)
    'seq3 = Mid(sequencia_codigo_barra, 35, 10)
    'seq4 = Mid(sequencia_codigo_barra, 6, 14)
    
    'seq1 = banco & moeda & carteira & left(nossonumero,2)
    'seq2 = right( nossonumero, 6 ) & dv_nossonumero & left( agencia , 3 )
    'seq3 = right( agencia, 1 ) & conta & dv_conta & "000"
    'seq4 = mid(sequencia_codigo_barra,6,14)
    
    dvcb = Mid(sequencia_codigo_barra, 5, 1)
    
    dv1 = Calculo_DV10(seq1)
    dv2 = Calculo_DV10(seq2)
    dv3 = Calculo_DV10(seq3)
    
    seq1 = Left(seq1 & dv1, 5) & "." & Mid(seq1 & dv1, 6, 5)
    seq2 = Left(seq2 & dv2, 5) & "." & Mid(seq2 & dv2, 6, 6)
    seq3 = Left(seq3 & dv3, 5) & "." & Mid(seq3 & dv3, 6, 6)
    
    Linha_Digitavel = seq1 & " " & seq2 & " " & seq3 & " " & dvcb & " " & seq4
    
End Function
    
Function Calculo_DV10(strNumero)
    
    Dim fator, total, numero, resto, i As Integer
    
    fator = 2
    total = 0
    For i = Len(strNumero) To 1 Step -1
        numero = Mid(strNumero, i, 1) * fator
        If numero > 9 Then
            numero = CInt(Left(numero, 1)) + CInt(Right(numero, 1))
        End If
        total = total + numero
        If fator = 2 Then
            fator = 1
        Else
            fator = 2
        End If
    Next
    resto = total Mod 10
    resto = 10 - resto
    If resto = 10 Then
        Calculo_DV10 = 0
    Else
        Calculo_DV10 = resto
    End If
    
End Function
    
Function calcula_DV_CodBarras(sequencia)
    
    Dim fator, total, numero, resto, i, resultado As Integer
    
    fator = 2
    total = 0
    For i = 43 To 1 Step -1
        numero = Val(Mid(sequencia, i, 1))
        If fator > 9 Then
            fator = 2
        End If
        numero = numero * fator
        total = total + numero
        fator = fator + 1
    Next
    resto = total Mod 11
    resultado = 11 - resto
    If resultado = 10 Or resultado = 0 Or resultado = 11 Then
        calcula_DV_CodBarras = 1
    Else
        calcula_DV_CodBarras = resultado
    End If
    
End Function

Public Function WBarCode(VALOR)
    Dim f, f1, f2, i
    Dim texto
    Const fino = 1
    Const largo = 3
    Const altura = 50
    Dim BarCodes(99)
    Dim retorno As String
    
    
    If BarCodes(0) = "" Then
      BarCodes(0) = "00110"
      BarCodes(1) = "10001"
      BarCodes(2) = "01001"
      BarCodes(3) = "11000"
      BarCodes(4) = "00101"
      BarCodes(5) = "10100"
      BarCodes(6) = "01100"
      BarCodes(7) = "00011"
      BarCodes(8) = "10010"
      BarCodes(9) = "01010"
      For f1 = 9 To 0 Step -1
        For f2 = 9 To 0 Step -1
          f = f1 * 10 + f2
          texto = ""
          For i = 1 To 5
            texto = texto & Mid(BarCodes(f1), i, 1) + Mid(BarCodes(f2), i, 1)
          Next
          BarCodes(f) = texto
        Next
      Next
    End If
    
    'Desenho da barra
    
    retorno = retorno + "<img src=http://www.empresaltda.com.br/boletos/img/p.gif width=" & fino & " height=" & altura & "  border=0><img"
    retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/b.gif width=" & fino & "  height=" & altura & "  border=0><img"
    retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/p.gif width=" & fino & "  height=" & altura & "  border=0><img"
    retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/b.gif width=" & fino & "  height=" & altura & "  border=0><img"
    
    
    texto = VALOR
    If Len(texto) Mod 2 <> 0 Then
      texto = "0" & texto
    End If
    
    ' Draw dos dados
    Do While Len(texto) > 0
      i = CInt(Left(texto, 2))
      texto = Right(texto, Len(texto) - 2)
      f = BarCodes(i)
      For i = 1 To 10 Step 2
        If Mid(f, i, 1) = "0" Then
          f1 = fino
        Else
          f1 = largo
        End If
    
        retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/p.gif width=" & f1 & " height=" & altura & " border=0><img"
     
        If Mid(f, i + 1, 1) = "0" Then
          f2 = fino
        Else
          f2 = largo
        End If
    
        retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/b.gif width=" & f2 & " height=" & altura & " border=0><img"
    
      Next
    Loop
    
    ' Draw guarda final
    retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/p.gif width=" & largo & " height=" & altura & " border=0><img"
    retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/b.gif width=" & fino & " height=" & altura & " border=0><img"
    retorno = retorno + " src=http://www.empresaltda.com.br/boletos/img/p.gif width=" & f1 & " height=" & altura & " border=0>"
    
    WBarCode = retorno
    
End Function


Public Function Replace2(Txt, Busca, Troca)

Dim a As Integer
Dim aux As String

For a = 1 To Len(Txt)
   If Mid(Txt, a, Len(Busca)) = Busca Then
      aux = aux & Troca
      a = a + Len(Busca) - 1
   Else
      aux = aux & Mid(Txt, a, 1)
   End If
Next a

Replace2 = aux

End Function


