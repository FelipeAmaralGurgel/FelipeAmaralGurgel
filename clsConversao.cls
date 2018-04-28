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
   
   strReais = "Alguns reais"
   strCentavos = " e alguns centavos"
   If InStr(CStr(pdblValor), ",") > 0 Then
      
   Else
      strCentavos = " e zero centavos"
   End If
   
  pstrValorExtenso = strReais & strCentavos
   

End Function

