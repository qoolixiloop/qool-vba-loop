VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExcelFunctions
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This class is exposed as an Automation Add-In. When this Automation Add-In is loaded in Excel
' (versions 2002 and later), functions in this module may be called directly from Excel worksheet
' cells.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function DoubleIt(D As Double) As Double
'''''''''''''''''''''''''''''''''''''''''''''
' DoubleIt
' This is an example function that simply
' doubles the input and returns the result.
'''''''''''''''''''''''''''''''''''''''''''''
DoubleIt = D * 2
End Function
