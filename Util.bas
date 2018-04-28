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
   
      Select Case pstrValor
         Case 1: converterValorExtensoUnidade = "um"
         Case 2: converterValorExtensoUnidade = "dois"
         Case 3: converterValorExtensoUnidade = "três"
         Case 4: converterValorExtensoUnidade = "quatro"
         Case 5: converterValorExtensoUnidade = "cinco"
         Case 6: converterValorExtensoUnidade = "seis"
         Case 7: converterValorExtensoUnidade = "sete"
         Case 8: converterValorExtensoUnidade = "oito"
         Case 9: converterValorExtensoUnidade = "nove"
'         Case Else
'            intUnidade = plngValor - CLng(Left(CStr(plngValor), 1) & "0")
      End Select
  
   
End Function

Public Function converterValorExtensoDez(pstrValor As String) As String
   Select Case pstrValor
      Case 10: converterValorExtensoDez = "dez"
      Case 11: converterValorExtensoDez = "onze"
      Case 12: converterValorExtensoDez = "doze"
      Case 13: converterValorExtensoDez = "treze"
      Case 14: converterValorExtensoDez = "quatorze"
      Case 15: converterValorExtensoDez = "quinze"
      Case 16: converterValorExtensoDez = "dezesseis"
      Case 17: converterValorExtensoDez = "dezessete"
      Case 18: converterValorExtensoDez = "dezoito"
      Case 19: converterValorExtensoDez = "dezenove"
   End Select
End Function

Public Function converterValorExtensoDezenas(pstrValor As String) As String
   Select Case pstrValor
      Case 2: converterValorExtensoDezenas = "vinte"
      Case 3: converterValorExtensoDezenas = "trinta"
      Case 4: converterValorExtensoDezenas = "quarenta"
      Case 5: converterValorExtensoDezenas = "cinquenta"
      Case 6: converterValorExtensoDezenas = "sessenta"
      Case 7: converterValorExtensoDezenas = "setenta"
      Case 8: converterValorExtensoDezenas = "oitenta"
      Case 9: converterValorExtensoDezenas = "noventa"
   End Select
End Function

