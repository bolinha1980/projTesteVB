VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function FuncaoRetornaSoma(ByVal Vr1 As Currency, ByVal Vr2 As Currency) As Currency
    FuncaoRetornaSoma = Vr1 + Vr2
End Function

Public Function FuncaoRetornaMult(ByVal Vr1 As Currency, ByVal Vr2 As Currency) As Currency
    FuncaoRetornaMult = Vr1 * Vr2
End Function

