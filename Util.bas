Attribute VB_Name = "Util"
Option Explicit

'Função que retorna os centavos de um valor financeiro
Public Function retornaCentavos(pdblValor As Double) As Integer
   If InStr(CStr(pdblValor), ",") > 0 Then
      retornaCentavos = Right(CStr(pdblValor), Len(CStr(pdblValor)) - InStr(CStr(pdblValor), ","))
   Else
      retornaCentavos = 0
   End If
End Function

'Função que retorna o valor inteiro de um valor financeiro, desconsiderando os centavos
Public Function retornaValorInteiro(pdblValor As Double) As Long
   If InStr(CStr(pdblValor), ",") > 0 Then
      retornaValorInteiro = Left(CStr(pdblValor), InStr(CStr(pdblValor), ",") - 1)
   Else
      retornaValorInteiro = pdblValor
   End If
End Function

Public Function converteValorExtenso(plngValor As Long) As String
      Select Case plngValor
         Case 0: converteValorExtenso = "zero"
         Case 1: converteValorExtenso = "um"
         Case 2: converteValorExtenso = "dois"
         Case 3: converteValorExtenso = "três"
         Case 4: converteValorExtenso = "quatro"
         Case 5: converteValorExtenso = "cinco"
         Case 6: converteValorExtenso = "seis"
         Case 7: converteValorExtenso = "sete"
         Case 8: converteValorExtenso = "oito"
         Case 9: converteValorExtenso = "nove"
         Case 10: converteValorExtenso = "dez"
         Case 11: converteValorExtenso = "onze"
         Case 12: converteValorExtenso = "doze"
         Case 13: converteValorExtenso = "treze"
         Case 14: converteValorExtenso = "quatorze"
         Case 15: converteValorExtenso = "quinze"
         Case 16: converteValorExtenso = "dezesseis"
         Case 17: converteValorExtenso = "dezessete"
         Case 18: converteValorExtenso = "dezoito"
         Case 19: converteValorExtenso = "dezenove"
         Case 20: converteValorExtenso = "vinte"
         Case 30: converteValorExtenso = "trinta"
         Case 40: converteValorExtenso = "quarenta"
         Case 50: converteValorExtenso = "cinquenta"
         Case 60: converteValorExtenso = "sessenta"
         Case 70: converteValorExtenso = "setenta"
         Case 80: converteValorExtenso = "oitenta"
         Case 90: converteValorExtenso = "noventa"
         
         
      End Select
  
   
End Function


