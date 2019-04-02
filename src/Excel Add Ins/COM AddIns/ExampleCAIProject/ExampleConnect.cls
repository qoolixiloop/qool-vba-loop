VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExampleConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExampleConnect
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This is the connection class for the COM Add-In (CAI). This will be registered with Windows
' and configured in the registry. See www.cpearson.com/excel/ComAddinInVB6.htm for details
' about this project.
'
' Required References:
' --------------------
' Microsoft Add-In Designer
' Microsoft Office 2003
' Microsoft Excel 2003
' Microsoft PowerPoint 2003
'
' This example CAI supports only Excel and PowerPoint. You can, of course, support other and/or
' additional applications by modifying the code in this class and add-in additional classes.
' This example creates only a single menu item on the Tools menu in Excel or PowerPoint. This is
' done in the CExcelControls and CWordControls classes. You can provide additional code in those
' classes to support many menu items and/or commandbar controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''
' Implement the COM Add-In interface.
'''''''''''''''''''''''''''''''''''''''''''''''''
Implements AddInDesignerObjects.IDTExtensibility2

'''''''''''''''''''''''''''''''''''''''''''''''''
' Declare procedures for all the methods defined
' in the IDTExtensibility2 interface.
'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant)
' not used
End Sub

Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant)
' not used
End Sub

Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant)
' not used
End Sub

Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, _
        ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
        ByVal AddInInst As Object, custom() As Variant)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IDTExtensibility2_OnConnection
' This is called when the add-in is connected to the host application (e.g., Excel).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''
' Set ThisCAI to the instance of this COM Add-In.
'''''''''''''''''''''''''''''''''''''''''''''''''
Set ThisCAI = AddInInst
Set ConnectInst = Me

'''''''''''''''''''''''''''''''''''''''''''''''''
' See if we're connecting to Excel. If so, set up
' the appropriate variables.
'''''''''''''''''''''''''''''''''''''''''''''''''
If TypeOf Application Is Excel.Application Then
    Set ExcelApp = Application
    Set ExcelEvents = New CExcelEvents
    Set ExcelEvents.ExcelApp = Application
    Set ExcelControls = New CExcelControls
    Set ExcelControls.ExcelApp = Application
    ExcelControls.CreateControls
End If

'''''''''''''''''''''''''''''''''''''''''''''''''
' See if we're connecting to PowerPoint. If so, set up
' the appropriate variables.
'''''''''''''''''''''''''''''''''''''''''''''''''
If TypeOf Application Is PowerPoint.Application Then
    Set PowerPointApp = Application
    Set PowerPointEvents = New CPowerPointEvents
    Set PowerPointEvents.PowerPointApp = Application
    Set PowerPointControls = New CPowerPointControls
    Set PowerPointControls.PowerPointApp = Application
    PowerPointControls.CreateControls
End If

End Sub

Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
    custom() As Variant)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IDTExtensibility2_OnDisconnection
' This executes when the COM Add-In is unloaded by the host application.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''
' Clean up Excel
''''''''''''''''''''''''''''''''
If Not ExcelApp Is Nothing Then
    '''''''''''''''''''''''''''''''''
    ' Set ALL Excel objects to Nothing.
    ' Otherwise they can cause the
    ' Excel process to continue to
    ' run in the background even
    ' after it appears to be closed.
    '''''''''''''''''''''''''''''''''
    Set ExcelEvents.ExcelApp = Nothing
    Set ExcelEvents = Nothing
    ExcelControls.DeleteControls
    Set ExcelControls = Nothing
    Set ExcelApp = Nothing
End If

''''''''''''''''''''''''''''''''
' Clean up PowerPoint
''''''''''''''''''''''''''''''''
If Not PowerPointApp Is Nothing Then
    '''''''''''''''''''''''''''''''''
    ' Set ALL PowerPoint objects to Nothing.
    ' Otherwise they can cause the
    ' Excel process to continue to
    ' run in the background even
    ' after it appears to be closed.
    '''''''''''''''''''''''''''''''''
    Set PowerPointEvents.PowerPointApp = Nothing
    Set PowerPointEvents = Nothing
    PowerPointControls.DeleteControls
    Set PowerPointControls = Nothing
    Set PowerPointApp = Nothing
End If

Set ThisCAI = Nothing

End Sub

