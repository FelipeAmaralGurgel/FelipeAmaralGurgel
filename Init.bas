Attribute VB_Name = "Init"
Option Explicit

Public Sub main()
   Dim objConversao As New clsConversao
   Dim strValorExtenso As String
   
   Call objConversao.Converter(1.5, strValorExtenso)
End Sub
