Attribute VB_Name = "Init"
Option Explicit

Public Sub main()
   Dim objConversao As New clsConversao
   Dim strValorExtenso As String
   
   Call objConversao.Converter(947.71, strValorExtenso)
   
   MsgBox ("Você converteu " & strValorExtenso)
End Sub
