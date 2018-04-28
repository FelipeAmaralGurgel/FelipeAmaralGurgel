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
