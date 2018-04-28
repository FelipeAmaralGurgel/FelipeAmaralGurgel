Attribute VB_Name = "Init"
Option Explicit

Public Sub main()
   Dim objConversao As New clsConversao
   Dim strValorExtenso As String
   
   'A classe clsConversao é responsável por todos os tratamentos para exibir valores
   Call objConversao.Converter(50500.71, strValorExtenso)
   
   MsgBox ("Você converteu " & strValorExtenso)
End Sub
