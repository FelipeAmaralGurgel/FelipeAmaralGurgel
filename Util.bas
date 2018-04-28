Attribute VB_Name = "Util"
Option Explicit

'Função que retorna os centavos de um valor financeiro
Public Function retornaCentavos(pstrValor As String) As String
   If InStr(CStr(pstrValor), ",") > 0 Then
      retornaCentavos = Right(CStr(pstrValor), Len(CStr(pstrValor)) - InStr(CStr(pstrValor), ","))
   Else
      retornaCentavos = 0
   End If
End Function

'Função que retorna o valor inteiro de um valor financeiro, desconsiderando os centavos
Public Function retornaValorInteiro(pstrValor As String) As String
   If InStr(CStr(pstrValor), ",") > 0 Then
      retornaValorInteiro = Left(CStr(pstrValor), InStr(CStr(pstrValor), ",") - 1)
   Else
      retornaValorInteiro = pstrValor
   End If
End Function

Public Function converterValorExtensoUnidade(pstrValor As Integer) As String
   Dim intUnidade As Integer
   
      Select Case plngValor
         Case 1: converteValorExtenso = "um"
         Case 2: converteValorExtenso = "dois"
         Case 3: converteValorExtenso = "três"
         Case 4: converteValorExtenso = "quatro"
         Case 5: converteValorExtenso = "cinco"
         Case 6: converteValorExtenso = "seis"
         Case 7: converteValorExtenso = "sete"
         Case 8: converteValorExtenso = "oito"
         Case 9: converteValorExtenso = "nove"
'         Case Else
'            intUnidade = plngValor - CLng(Left(CStr(plngValor), 1) & "0")
      End Select
  
   
End Function

Public Function converterValorExtensoDez(pstrValor As String) As String
   Select Case pstrValor
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
   End Select
End Function

Public Function converterValorExtensoDezenas(pstrValor As String) As String
   Select Case pstrValor
      Case 2: converteValorExtenso = "vinte"
      Case 3: converteValorExtenso = "trinta"
      Case 4: converteValorExtenso = "quarenta"
      Case 5: converteValorExtenso = "cinquenta"
      Case 6: converteValorExtenso = "sessenta"
      Case 7: converteValorExtenso = "setenta"
      Case 8: converteValorExtenso = "oitenta"
      Case 9: converteValorExtenso = "noventa"
   End Select
End Function

