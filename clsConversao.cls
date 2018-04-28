VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConversao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Converter()
On Error Resume Next
   Dim strValor As String
   Dim strCentavos As String
   Dim strValorInteiro As String
   Dim strValorExtenso As String
   Dim strReais As String
   
   'Define valor padrão para o usuário não ficar sem retorno
   strValor = Format(InputBox("Valor: ", "Extenso"), "#,##0.00")
   
   strCentavos = Right(strValor, 2)
   
   'Pega somente o valor dos centavos
   strCentavos = retornaCentavos(strValor)
   strValorInteiro = retornaValorInteiro(strValor)
   
   If Val(strCentavos) > 0 Then
      If Mid(strCentavos, 1, 1) = 1 Then
        strValorExtenso = converterValorExtensoDez(strCentavos)
      Else
        strValorExtenso = converterValorExtensoDezenas(Left(strCentavos, 1)) & " e "
        strValorExtenso = strValorExtenso & converterValorExtensoUnidade(Right(strCentavos, 1))
      End If
      strValorExtenso = strValorExtenso & " centavos"
   End If
   
   If Val(strValorInteiro) > 0 Then
      If Mid(strValorInteiro, 1, 1) = 1 Then
        strValorExtenso = converterValorExtensoDez(strValorInteiro)
      Else
        strValorExtenso = converterValorExtensoDezenas(Left(strCentavos, 1)) & " e "
        strValorExtenso = strValorExtenso & converterValorExtensoUnidade(Right(strCentavos, 1))
      End If
      strValorExtenso = strValorExtenso & " centavos"
   End If
   
   
   If Len(strValorExtenso) > 0 Then
      MsgBox "Valor por extenso: " & vbCrLf & strValorExtenso, vbInformation + vbOKOnly, "Valor por Extenso"
   Else
      MsgBox "Não foi informado nenhum valor", vbInformation + vbOKOnly, "Não há informação"
   End If

End Function

