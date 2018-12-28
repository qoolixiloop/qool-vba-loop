VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPowerPointEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents pPowerPointApp As PowerPoint.Application
Attribute pPowerPointApp.VB_VarHelpID = -1

Friend Property Get PowerPointApp() As PowerPoint.Application
    Set PowerPointApp = pPowerPointApp
End Property


Friend Property Set PowerPointApp(PPTApp As PowerPoint.Application)
    Set pPowerPointApp = PPTApp
End Property

Private Sub pPowerPointApp_NewPresentation(ByVal Pres As PowerPoint.Presentation)
    MsgBox "Event from " & ThisCAI.ProgId & "NewPresentation: " & Pres.Name
End Sub

