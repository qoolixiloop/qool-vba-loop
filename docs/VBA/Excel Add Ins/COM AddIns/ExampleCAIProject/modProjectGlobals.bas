Attribute VB_Name = "modProjectGlobals"
Option Explicit
Option Compare Text
Option Private Module
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExampleConnect
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This module contains project-wide globals and constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''
' CONSTANTS
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const C_EXCEL_TOOLS_MENU_ID As Long = 30007
Public Const C_POWERPOINT_TOOLS_MENU_ID As Long = 30007

'''''''''''''''''''''''''''''''''''''''''''''''''''
' VARIABLES
'''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''
' This is an instance of the COM Add-In inself.
''''''''''''''''''''''''''''''''''''''''''''''''
Public ThisCAI As Office.COMAddIn
''''''''''''''''''''''''''''''''''''''''''''''''
' These are references to the host applications.
''''''''''''''''''''''''''''''''''''''''''''''''
Public ExcelApp As Excel.Application
Public PowerPointApp  As PowerPoint.Application
''''''''''''''''''''''''''''''''''''''''''''''''
' These are the application event handlers.
''''''''''''''''''''''''''''''''''''''''''''''''
Public ExcelEvents As CExcelEvents
Public PowerPointEvents As CPowerPointEvents
''''''''''''''''''''''''''''''''''''''''''''''''
' This references the Connect object.
''''''''''''''''''''''''''''''''''''''''''''''''
Public ConnectInst As ExampleConnect

''''''''''''''''''''''''''''''''''''''''''''''''
' Instance of CExcelMenus
''''''''''''''''''''''''''''''''''''''''''''''''
Public ExcelControls As CExcelControls
Public PowerPointControls As CPowerPointControls


