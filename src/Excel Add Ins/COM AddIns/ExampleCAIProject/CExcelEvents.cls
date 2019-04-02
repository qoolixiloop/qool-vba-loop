VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExcelEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CExcelEvents
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This module handles events from the Excel application.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private WithEvents pExcelApp As Excel.Application
Attribute pExcelApp.VB_VarHelpID = -1

Friend Property Get ExcelApp() As Excel.Application
    Set ExcelApp = pExcelApp
End Property

Friend Property Set ExcelApp(XLApp As Excel.Application)
    Set pExcelApp = XLApp
End Property

Private Sub Class_Terminate()
    Set pExcelApp = Nothing
End Sub

Private Sub pExcelApp_NewWorkbook(ByVal Wb As Excel.Workbook)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' pExcelApp_NewWorkbook
' Example Excel application event. Display a MsgBox
' when a new workbook is opened.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MsgBox ThisCAI.ProgId & " NewWorkbook: " & Wb.Name
End Sub


