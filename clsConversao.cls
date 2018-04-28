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

Public Function Converter(pdblValor As Double, ByRef pstrValorExtenso As String)
   Dim strReais As String
   Dim strCentavos As String
   Dim intCentavos As Integer
   
   'Define valor padrão para o usuário não ficar sem retorno
   strReais = "Alguns reais"
   strCentavos = " e alguns centavos"
   'Pega somente o valor dos centavos
   intCentavos = retornaCentavos(pdblValor)
   If intCentavos > 0 Then
      
   Else
      'se não tem centavos, define zero
      strCentavos = " e zero centavos"
   End If
   
   'Retorna o valor dos reais concatenados com centavos
   pstrValorExtenso = strReais & strCentavos
   

End Function

